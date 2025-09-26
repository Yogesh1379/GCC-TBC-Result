// Full working merged code with "Extra Space", "Extra Tab", "Extra Enter" and original wrong-word logic from Processor22
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.io.*;
import java.util.*;
import java.util.regex.*;

public class StudentAnswerBatchProcessor28 {
    static File modelFile = null;
    static List<String> modelParas;

    public static void main(String[] args) throws Exception {
        File folder = new File("F:\\desktop backup 5 july 25\\GCC TBC MAY 2025 AUTO\\marking\\New folder");
        File compareDir = new File(folder.getParentFile(), folder.getName() + "_Compared");
        if (!compareDir.exists()) compareDir.mkdirs();

        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Results");
        String[] columns = {"File Name", "Extra Word", "Missing Word", "Wrong Word",
                "Extra Space", "Extra Tab", "Extra Enter", "Total Mistakes", "Final Marks", "Mistakes Summary"};
        Row header = sheet.createRow(0);
        for (int i = 0; i < columns.length; i++) header.createCell(i).setCellValue(columns[i]);

        int rowNum = 1;
        File[] docxFiles = folder.listFiles((dir, name) -> name.toLowerCase().endsWith(".docx") && !name.startsWith("~$"));
        if (docxFiles == null) return;

        for (File studentFile : docxFiles) {
            // Map to model
            String[] parts = studentFile.getName().split("_");
            String seatno = parts[1], batchname = parts[2];
            int course1 = seatno.substring(4, 6).equals("15") ? 1 : 2;
            try (FileInputStream fis = new FileInputStream("F:\\desktop backup 5 july 25\\GCC TBC MAY 2025 AUTO\\allocation\\BathwiseSubjective.xlsx");
                 Workbook wb1 = new XSSFWorkbook(fis)) {
                for (Row row : wb1.getSheetAt(0)) {
                    if (row.getRowNum() == 0) continue;
                    if (getCellValueAsString(row.getCell(6)).equals(batchname)
                            && Integer.parseInt(getCellValueAsString(row.getCell(5))) == course1) {
                        String fileName = getCellValueAsString(row.getCell(2));
                        if ((course1 == 1 && fileName.startsWith("Eng30 Speed")) ||
                            (course1 == 2 && fileName.startsWith("Eng 40 Speed"))) {
                            modelFile = new File("F:\\desktop backup 5 july 25\\GCC TBC MAY 2025 AUTO\\allocation\\New folder\\All_Subjective_Eng30\\" + fileName);
                            break;
                        }
                    }
                }
            }

            modelParas = extractFormattedParagraphs(modelFile);
            List<String> studentParas = extractFormattedParagraphs(studentFile);
            List<String[]> wordMistakes = new ArrayList<>();
            Map<String, Integer> formattingMistakes = new HashMap<>();

            XWPFDocument outputDoc = new XWPFDocument();
            int refIndex = 0, stuIndex = 0;
            while (stuIndex < studentParas.size()) {
                String studentLine = studentParas.get(stuIndex);
                String modelLine = refIndex < modelParas.size() ? modelParas.get(refIndex) : "";

                // Extra Enters
                int blankCount = 0, tempIdx = stuIndex;
                while (tempIdx < studentParas.size() && studentParas.get(tempIdx).trim().isEmpty()) {
                    blankCount++; tempIdx++;
                }
                if (blankCount >= 3) {
                    XWPFParagraph para = outputDoc.createParagraph();
                    XWPFRun run = para.createRun();
                    run.setText("[Extra Enters Group]"); run.setColor("FF0000");
                    incrementMistake("Extra Enter", formattingMistakes);
                    stuIndex += blankCount;
                    continue;
                }

                XWPFParagraph para = outputDoc.createParagraph();
                studentLine = handleExtraSpacesAndTabs(para, studentLine, formattingMistakes);
                compareWordsUsing22Logic(modelLine, studentLine, para, formattingMistakes, wordMistakes, studentFile.getName());
                refIndex++; stuIndex++;
            }

            // Tally counts
            int extraWord = 0, missingWord = 0, wrongWord = 0;
            for (String[] m : wordMistakes) {
                switch (m[1]) {
                    case "Extra Word": extraWord++; break;
                    case "Missing Word": missingWord++; break;
                    case "Wrong Word": wrongWord++; break;
                }
            }
            int totalFormatting = formattingMistakes.values().stream().mapToInt(i -> i).sum();
            int totalMistakes = wordMistakes.size() + totalFormatting;
            int obtainedMarks = Math.max(0, 40 - totalMistakes);

            // Summary in DOCX
            XWPFParagraph summary = outputDoc.createParagraph();
            XWPFRun run = summary.createRun();
            run.setText("Extra Word: " + extraWord); run.addBreak();
            run.setText("Missing Word: " + missingWord); run.addBreak();
//            run.setText("Wrong Word: " + wrongWord); run.addBreak();
            for (Map.Entry<String, Integer> e : formattingMistakes.entrySet()) {
                run.setText(e.getKey() + ": " + e.getValue()); run.addBreak();
            }
            run.setText("Total Mistakes: " + totalMistakes); run.addBreak();
            run.setText("Final Marks: " + obtainedMarks + " / 40");

            // Save DOCX
            File compared = new File(compareDir, studentFile.getName().replace(".docx", "_Compared.docx"));
            try (FileOutputStream out = new FileOutputStream(compared)) {
                outputDoc.write(out);
            }

            // Write Excel row
            Row row = sheet.createRow(rowNum++);
            row.createCell(0).setCellValue(studentFile.getName());
            row.createCell(1).setCellValue(extraWord);
            row.createCell(2).setCellValue(missingWord);
            row.createCell(3).setCellValue(wrongWord);
            row.createCell(4).setCellValue(formattingMistakes.getOrDefault("Extra Space", 0));
            row.createCell(5).setCellValue(formattingMistakes.getOrDefault("Extra Tab", 0));
            row.createCell(6).setCellValue(formattingMistakes.getOrDefault("Extra Enter", 0));
            row.createCell(7).setCellValue(totalMistakes);
            row.createCell(8).setCellValue(obtainedMarks);
            StringBuilder sb = new StringBuilder();
            for (String[] m : wordMistakes) sb.append("[").append(m[1]).append(" - ").append(m[2]).append("] ");
            row.createCell(9).setCellValue(sb.toString().trim());
        }

        // Autosize & save Excel
        for (int i = 0; i < sheet.getRow(0).getLastCellNum(); i++) sheet.autoSizeColumn(i);
        try (FileOutputStream out = new FileOutputStream(new File(folder.getParent(), "Comparison_Result.xlsx"))) {
            workbook.write(out);
        }
        workbook.close();
        System.out.println("✅ Processing complete.");
    }

    private static String handleExtraSpacesAndTabs(XWPFParagraph para, String text, Map<String,Integer> mistakes) {
        Matcher m = Pattern.compile("( {3,})").matcher(text);
        int last = 0;
        while (m.find()) {
            addTextRun(para, text.substring(last, m.start()), false);
            addTextRun(para, m.group(), true);
            incrementMistake("Extra Space", mistakes);
            last = m.end();
        }
        if (last < text.length()) addTextRun(para, text.substring(last), false);
        if (text.contains("\t\t\t")) {
            addTextRun(para, "[Extra Tabs]", true);
            incrementMistake("Extra Tab", mistakes);
        }
        return text;
    }

    private static void compareWordsUsing22Logic(String model, String student,
                                                 XWPFParagraph para,
                                                 Map<String,Integer> mistakeCounts,
                                                 List<String[]> wordMistakes,
                                                 String fileName) {
        model = model.trim().replaceAll("\\s+([,;:!?])","$1");
        student = student.trim().replaceAll("\\s+([,;:!?])","$1");
        model = model.replaceAll("\\.(?!\\s|$)",". ");
        student = student.replaceAll("\\.(?!\\s|$)",". ");

        String[] modelWords = model.split("\\s+");
        String[] studentWords = student.split("\\s+");
        int i=0, j=0;
        while (i<modelWords.length || j<studentWords.length) {
            String refRaw = i<modelWords.length? modelWords[i]:null;
            String stuRaw = j<studentWords.length? studentWords[j]:null;
            String ref = normalize(refRaw), stu = normalize(stuRaw);
            XWPFRun run = para.createRun();
            if (ref!=null && stu!=null && ref.equals(stu)) {
                run.setText(stu+" "); i++; j++;
            } else if (ref!=null && stu!=null) {
                if (j+1<studentWords.length && ref.equals(normalize(studentWords[j+1]))) {
                    highlight(run, "["+stu+"]");
                    wordMistakes.add(new String[]{fileName,"Extra Word",stu});
                    incrementMistake("Extra Word",mistakeCounts);
                    j++;
                } else if (i+1<modelWords.length && normalize(modelWords[i+1]).equals(stu)) {
                    highlight(run, "["+ref+"]");
                    wordMistakes.add(new String[]{fileName,"Missing Word",ref});
                    incrementMistake("Missing Word",mistakeCounts);
                    i++;
                } else {
                    highlight(run, "["+stu+"/"+ref+"]");
                    wordMistakes.add(new String[]{fileName,"Wrong Word",stu+"/"+ref});
                    incrementMistake("Wrong Word",mistakeCounts);
                    i++; j++;
                }
            } else if (stu!=null) {
                highlight(run, "["+stu+"]");
                wordMistakes.add(new String[]{fileName,"Extra Word",stu});
                incrementMistake("Extra Word",mistakeCounts);
                j++;
            } else if (ref!=null) {
                highlight(run, "["+ref+"]");
                wordMistakes.add(new String[]{fileName,"Missing Word",ref});
                incrementMistake("Missing Word",mistakeCounts);
                i++;
            }
        }
    }

    private static void addTextRun(XWPFParagraph para, String text, boolean highlight) {
        XWPFRun run = para.createRun(); run.setText(text);
        if (highlight) {
            CTRPr rpr = run.getCTR().isSetRPr() ? run.getCTR().getRPr() : run.getCTR().addNewRPr();
            CTShd shd = rpr.addNewShd(); shd.setVal(STShd.CLEAR); shd.setFill("FF0000");
        }
    }

    private static void highlight(XWPFRun run, String text) {
        run.setText(text + " "); run.setColor("FF0000");
    }

    private static void incrementMistake(String type, Map<String,Integer> counter) {
        counter.put(type, counter.getOrDefault(type, 0) + 1);
    }

    private static List<String> extractFormattedParagraphs(File file) throws IOException {
        List<String> paras = new ArrayList<>();
        try (FileInputStream fis = new FileInputStream(file);
             XWPFDocument doc = new XWPFDocument(fis)) {
            for (XWPFParagraph p : doc.getParagraphs()) {
                StringBuilder sb = new StringBuilder();
                for (XWPFRun r : p.getRuns()) sb.append(r.toString());
                paras.add(sb.toString());
            }
        }
        return paras;
    }

    private static String getCellValueAsString(Cell cell) {
        switch (cell.getCellType()) {
            case STRING: return cell.getStringCellValue().trim();
            case NUMERIC: return String.valueOf((int)cell.getNumericCellValue());
            default: return "";
        }
    }

    private static String normalize(String w) {
        return w == null ? null : w.trim().replace("’","'").replace("‘","'");
    }

    private static int countMatches(String text, String regex) {
        return (int)Pattern.compile(regex).matcher(text).results().count();
    }
}
