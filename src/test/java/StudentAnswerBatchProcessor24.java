import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.io.*;
import java.util.*;
import java.util.regex.*;

public class StudentAnswerBatchProcessor24 {
    static File modelFile = null;

    public static void main(String[] args) throws Exception {
        File folder = new File("F:\\Desktop backup 10-10-2025\\New folder\\New folder\\New folder");
        File compareDir = new File(folder.getParentFile(), folder.getName() + "_Compared");
        if (!compareDir.exists()) compareDir.mkdirs();

        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Results");
        String[] columns = {"File Name", "Total Mistakes", "Final Marks",
                "Extra Space", "Extra Tab", "Extra Enter",
                "Extra Word", "Missing Word", "Wrong Word",
                "Mistakes Summary"};
        Row header = sheet.createRow(0);
        for (int i = 0; i < columns.length; i++) {
            header.createCell(i).setCellValue(columns[i]);
        }

        int rowNum = 1;

        File[] docxFiles = folder.listFiles((dir, name) ->
                name.toLowerCase().endsWith(".docx") && !name.startsWith("~$"));

        if (docxFiles == null) {
            System.out.println("\u274C No .docx files found.");
            return;
        }

        for (File studentFile : docxFiles) {
            String stdfile = studentFile.getName();
            String[] parts = stdfile.split("_");
            String seatno = parts[1];
            String batchname = parts[2];
            if (seatno.length() >= 10) {
                String course = seatno.substring(4, 6);
                int course1 = course.equals("15") ? 1 : (course.equals("16") ? 2 : 0);
                FileInputStream fis = new FileInputStream("F:\\GCC-TBC October  repeaters\\eng allocation\\Batch Wise Subjective (1).xlsx");
                Workbook workbook1 = new XSSFWorkbook(fis);
                Sheet sheet1 = workbook1.getSheetAt(0);

                for (Row row : sheet1) {
                    if (row.getRowNum() == 0) continue;
                    Cell batchCell = row.getCell(6);
                    Cell courseCell = row.getCell(5);
                    Cell subjectiveCell = row.getCell(2);
                    if (batchCell == null || courseCell == null || subjectiveCell == null) continue;
                    String excelBatch = getCellValueAsString(batchCell);
                    int excelCourse1 = Integer.parseInt(getCellValueAsString(courseCell));
                    if (excelBatch.equals(batchname) && excelCourse1 == course1) {
                        String fileCandidate = getCellValueAsString(subjectiveCell);
                        if ((fileCandidate.startsWith("Eng30 Speed") && course1 == 1) ||
                            (fileCandidate.startsWith("Eng 40 Speed") && course1 == 2)) {
                            modelFile = new File("F:\\GCC-TBC October  repeaters\\Question files\\SUBJECTIE-\\ENG-40\\ENG 40-SPEED\\" + fileCandidate);
                            break;
                        }
                    }
                }
            }

            List<String> modelParas = extractFormattedParagraphs(modelFile);
            List<String> studentParas = extractFormattedParagraphs(studentFile);
            Map<String, Integer> mistakeCounts = new HashMap<>();
            List<String[]> wordMistakes = new ArrayList<>();
            XWPFDocument outputDoc = new XWPFDocument();

            int refIndex = 0, stuIndex = 0;
            while (stuIndex < studentParas.size()) {
                String studentLine = studentParas.get(stuIndex);
                String modelLine = refIndex < modelParas.size() ? modelParas.get(refIndex) : "";

                int blankLineCount = 0;
                int tempIndex = stuIndex;
                while (tempIndex < studentParas.size()) {
                    String tempLine = studentParas.get(tempIndex).trim();
                    String modelTempLine = refIndex < modelParas.size() ? modelParas.get(refIndex).trim() : "";
                    if (tempLine.isEmpty() && !modelTempLine.isEmpty()) {
                        blankLineCount++;
                        tempIndex++;
                    } else break;
                }
                if (blankLineCount > 0) {
                    for (int b = 0; b < blankLineCount; b++) {
                        XWPFParagraph para = outputDoc.createParagraph();
                        XWPFRun run = para.createRun();
                        run.setText("[Extra Enter]");
                        run.setColor("FF0000");
                        incrementMistake("Extra Enter", mistakeCounts);
                    }
                    stuIndex += blankLineCount;
                    continue;
                }

                XWPFParagraph para = outputDoc.createParagraph();
                Pattern wsPattern = Pattern.compile("( {2,})|(\\t+)");
                Matcher wsMatcher = wsPattern.matcher(studentLine);
                int lastIndex = 0;
                StringBuilder cleanStuLine = new StringBuilder();

                while (wsMatcher.find()) {
                    int start = wsMatcher.start();
                    int end = wsMatcher.end();
                    if (start > lastIndex) {
                        String normal = studentLine.substring(lastIndex, start);
                        addTextRun(para, normal, false);
                        cleanStuLine.append(normal);
                    }
                    if (wsMatcher.group(1) != null) {
                        String spaceGroup = wsMatcher.group(1);
                        int count = spaceGroup.length();
                        int fullGroups = count / 2;
                        int leftover = count % 2;
                        for (int i = 0; i < fullGroups; i++) {
                            addTextRun(para, "  ", true);
                            cleanStuLine.append(" ");
                            incrementMistake("Extra Space", mistakeCounts);
                        }
                        if (leftover > 0) {
                            addTextRun(para, " ", false);
                            cleanStuLine.append(" ");
                        }
                    } else if (wsMatcher.group(2) != null) {
                        String tabGroup = wsMatcher.group(2);
                        int tabCount = tabGroup.length();
                        XWPFRun run = para.createRun();
                        run.setText("[Extra Tab" + (tabCount > 1 ? "s" : "") + "]");
                        run.setColor("FF0000");
                        cleanStuLine.append(" ");
                        for (int i = 0; i < tabCount; i++) incrementMistake("Extra Tab", mistakeCounts);
                    }
                    lastIndex = end;
                }
                if (lastIndex < studentLine.length()) {
                    String remaining = studentLine.substring(lastIndex);
                    addTextRun(para, remaining, false);
                    cleanStuLine.append(remaining);
                }

                compareWords(modelLine, cleanStuLine.toString().trim(), para, mistakeCounts, wordMistakes, studentFile.getName());
                refIndex++;
                stuIndex++;
            }

            // Add summary section
            XWPFParagraph summary = outputDoc.createParagraph();
            XWPFRun run = summary.createRun();
            run.setBold(true);
            run.setText("Summary of Mistakes:");
            run.addBreak();
            int extraWord = 0, missingWord = 0, wrongWord = 0;
            for (String[] m : wordMistakes) {
                switch (m[1]) {
                    case "Extra Word": extraWord++; break;
                    case "Missing Word": missingWord++; break;
                    case "Wrong Word": wrongWord++; break;
                }
            }
            for (String key : Arrays.asList("Extra Space", "Extra Tab", "Extra Enter")) {
                int c = mistakeCounts.getOrDefault(key, 0);
                if (c > 0) {
                    run = summary.createRun();
                    run.setColor("FF0000");
                    run.setText(key + ": " + c);
                    run.addBreak();
                }
            }
            run = summary.createRun();
            run.setText("Extra Word: " + extraWord);
            run.addBreak();
            run.setText("Missing Word: " + missingWord);
            run.addBreak();
            run.setText("Wrong Word: " + wrongWord);
            run.addBreak();

            int totalMistakes = wordMistakes.size() +
                mistakeCounts.getOrDefault("Extra Space", 0) +
                mistakeCounts.getOrDefault("Extra Tab", 0) +
                mistakeCounts.getOrDefault("Extra Enter", 0);
            int obtainedMarks = Math.max(0, 40 - totalMistakes);
            run.setText("Total Mistakes: " + totalMistakes);
            run.addBreak();
            run.setText("Final Marks: " + obtainedMarks + " / 40");

            File comparedFile = new File(compareDir, studentFile.getName().replace(".docx", "_Compared.docx"));
            try (FileOutputStream out = new FileOutputStream(comparedFile)) {
                outputDoc.write(out);
            }

            Row row = sheet.createRow(rowNum++);
            row.createCell(0).setCellValue(studentFile.getName());
            row.createCell(1).setCellValue(totalMistakes);
            row.createCell(2).setCellValue(obtainedMarks);
            row.createCell(3).setCellValue(mistakeCounts.getOrDefault("Extra Space", 0));
            row.createCell(4).setCellValue(mistakeCounts.getOrDefault("Extra Tab", 0));
            row.createCell(5).setCellValue(mistakeCounts.getOrDefault("Extra Enter", 0));
            row.createCell(6).setCellValue(extraWord);
            row.createCell(7).setCellValue(missingWord);
            row.createCell(8).setCellValue(wrongWord);

            StringBuilder summaryStr = new StringBuilder();
            for (String[] m : wordMistakes) summaryStr.append(m[1]).append(" - ").append(m[2]).append(" | ");
            row.createCell(9).setCellValue(summaryStr.toString());
        }

        for (int i = 0; i < columns.length; i++) sheet.autoSizeColumn(i);
        try (FileOutputStream out = new FileOutputStream(new File(folder.getParent(), "Comparison_Result.xlsx"))) {
            workbook.write(out);
        }
        workbook.close();
        System.out.println("\u2705 All students processed.");
    }

    private static void compareWords(String model, String student, XWPFParagraph para, Map<String, Integer> mistakeCounts, List<String[]> wordMistakes, String fileName) {
        String[] modelWords = model.trim().split("\\s+");
        String[] studentWords = student.trim().split("\\s+");
        int i = 0, j = 0;
        while (i < modelWords.length || j < studentWords.length) {
            String refWordRaw = i < modelWords.length ? modelWords[i] : null;
            String stuWordRaw = j < studentWords.length ? studentWords[j] : null;
            String refWord = normalizeWord(refWordRaw);
            String stuWord = normalizeWord(stuWordRaw);
            if (refWord != null && stuWord != null && refWord.equals(stuWord)) {
                para.createRun().setText(stuWord + " "); i++; j++;
            } else if (refWord != null && stuWord != null) {
                boolean matched = false;
                for (int k = 1; k <= 3 && (i + k) < modelWords.length; k++) {
                    if (normalizeWord(modelWords[i + k]).equals(stuWord)) {
                        for (int x = i; x < i + k; x++) {
                            highlight(para.createRun(), "[" + modelWords[x] + "]");
                            wordMistakes.add(new String[]{fileName, "Missing Word", modelWords[x]});
                            incrementMistake("Missing Word", mistakeCounts);
                        }
                        i += k; matched = true; break;
                    }
                }
                if (matched) continue;
                for (int k = 1; k <= 3 && (j + k) < studentWords.length; k++) {
                    if (normalizeWord(studentWords[j + k]).equals(refWord)) {
                        for (int x = j; x < j + k; x++) {
                            highlight(para.createRun(), "[" + studentWords[x] + "]");
                            wordMistakes.add(new String[]{fileName, "Extra Word", studentWords[x]});
                            incrementMistake("Extra Word", mistakeCounts);
                        }
                        j += k; matched = true; break;
                    }
                }
                if (matched) continue;
                highlight(para.createRun(), "[" + refWord + "/" + stuWord + "]");
                wordMistakes.add(new String[]{fileName, "Wrong Word", refWord + "/" + stuWord});
                incrementMistake("Wrong Word", mistakeCounts); i++; j++;
            } else if (stuWord != null) {
                highlight(para.createRun(), "[" + stuWord + "]");
                wordMistakes.add(new String[]{fileName, "Extra Word", stuWord});
                incrementMistake("Extra Word", mistakeCounts); j++;
            } else if (refWord != null) {
                highlight(para.createRun(), "[" + refWord + "]");
                wordMistakes.add(new String[]{fileName, "Missing Word", refWord});
                incrementMistake("Missing Word", mistakeCounts); i++;
            }
        }
    }

    private static void addTextRun(XWPFParagraph para, String text, boolean highlight) {
        XWPFRun run = para.createRun();
        run.setText(text);
        if (highlight) {
            CTRPr rpr = run.getCTR().isSetRPr() ? run.getCTR().getRPr() : run.getCTR().addNewRPr();
            CTShd shd = rpr.addNewShd();
            shd.setVal(STShd.CLEAR);
            shd.setColor("auto");
            shd.setFill("FF0000");
        }
    }

    private static void highlight(XWPFRun run, String text) {
        run.setText(text + " ");
        run.setColor("FF0000");
    }

    private static List<String> extractFormattedParagraphs(File file) throws IOException {
        List<String> paras = new ArrayList<>();
        try (FileInputStream fis = new FileInputStream(file);
             XWPFDocument doc = new XWPFDocument(fis)) {
            for (XWPFParagraph p : doc.getParagraphs()) {
                StringBuilder fullText = new StringBuilder();
                for (XWPFRun run : p.getRuns()) fullText.append(run.toString());
                paras.add(fullText.toString());
            }
        }
        return paras;
    }

    private static void incrementMistake(String type, Map<String, Integer> counter) {
        counter.put(type, counter.getOrDefault(type, 0) + 1);
    }

    private static String normalizeWord(String word) {
        return word == null ? null : word.replace("’", "'").replace("‘", "'");
    }

    private static String getCellValueAsString(Cell cell) {
        switch (cell.getCellType()) {
            case STRING: return cell.getStringCellValue().trim();
            case NUMERIC: return String.valueOf((int) cell.getNumericCellValue());
            default: return "";
        }
    }
}
