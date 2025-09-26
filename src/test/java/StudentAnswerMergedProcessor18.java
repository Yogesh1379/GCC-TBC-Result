
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.io.*;
import java.util.*;
import java.util.regex.*;

public class StudentAnswerMergedProcessor18 {

    public static void main(String[] args) throws Exception {
        File folder = new File("C:\\Users\\User\\Desktop\\New folder (4)\\SpeedFiles");
        File compareDir = new File(folder.getParentFile(), folder.getName() + "_Compared");
        if (!compareDir.exists()) compareDir.mkdirs();
        File modelFile = new File("C:\\Users\\User\\Desktop\\Eng30 Speed 4.docx");
        List<String> modelParas = extractFormattedParagraphs(modelFile);

        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Results");
        String[] columns = {
                "File Name", "Total Mistakes", "Final Marks",
                "Extra Space", "Extra Tab", "Extra Enter",
                "Extra Word", "Missing Word", "Wrong Word",
                "Mistakes Summary"
        };
        Row header = sheet.createRow(0);
        for (int i = 0; i < columns.length; i++) header.createCell(i).setCellValue(columns[i]);

        int rowNum = 1;
        File[] docxFiles = folder.listFiles(new FilenameFilter() {
            public boolean accept(File dir, String name) {
                return name.toLowerCase().endsWith(".docx") && !name.startsWith("~$");
            }
        });
        if (docxFiles == null) return;

        for (File studentFile : docxFiles) {
            List<String> studentParas = extractFormattedParagraphs(studentFile);
            Map<String, Integer> mistakeCounts = new HashMap<String, Integer>();
            List<String[]> wordMistakes = new ArrayList<String[]>();

            XWPFDocument outputDoc = new XWPFDocument();
            int refIndex = 0, stuIndex = 0;

            while (stuIndex < studentParas.size()) {
                String studentLine = studentParas.get(stuIndex);
                String modelLine = refIndex < modelParas.size() ? modelParas.get(refIndex) : "";

                int blankLineCount = 0, tempIndex = stuIndex;
                while (tempIndex < studentParas.size() && studentParas.get(tempIndex).trim().isEmpty()) {
                    if (refIndex < modelParas.size() && !modelParas.get(refIndex).trim().isEmpty())
                        blankLineCount++;
                    tempIndex++;
                }
                if (blankLineCount >= 3) {
                    XWPFParagraph para = outputDoc.createParagraph();
                    XWPFRun run = para.createRun();
                    run.setText("[Extra Enters Group]");
                    run.setColor("FF0000");
                    incrementMistake("Extra Enter", mistakeCounts);
                    stuIndex += blankLineCount;
                    continue;
                }

                XWPFParagraph para = outputDoc.createParagraph();
                Pattern pattern = Pattern.compile("( {3,})");
                Matcher matcher = pattern.matcher(studentLine);
                int lastIndex = 0;
                while (matcher.find()) {
                    int start = matcher.start(1);
                    int end = matcher.end(1);
                    if (start > lastIndex)
                        addTextRun(para, studentLine.substring(lastIndex, start), false);

                    int spaceCount = matcher.group(1).length();
                    for (int i = 0; i < spaceCount / 3; i++) {
                        addTextRun(para, "   ", true);
                        incrementMistake("Extra Space", mistakeCounts);
                    }
                    if (spaceCount % 3 != 0)
                        addTextRun(para, matcher.group(1).substring(spaceCount - (spaceCount % 3)), false);
                    lastIndex = end;
                }
                if (lastIndex < studentLine.length())
                    addTextRun(para, studentLine.substring(lastIndex), false);

                int tabCount = countMatches(studentLine, "\\t{3,}");
                for (int i = 0; i < tabCount; i++)
                    incrementMistake("Extra Tab", mistakeCounts);
                if (tabCount > 0) {
                    XWPFRun run = para.createRun();
                    run.setText("[Extra Tabs]");
                    run.setColor("FF0000");
                }

                if (studentLine.startsWith("\t") && !modelLine.startsWith("\t")) {
                    incrementMistake("Extra Tab", mistakeCounts);
                    XWPFRun run = para.createRun();
                    run.setText("[Extra Tab]");
                    run.setColor("FF0000");
                    studentLine = studentLine.replaceFirst("^\\t+", "");
                }

                compareWords(modelLine, studentLine, para, mistakeCounts, wordMistakes, studentFile.getName());
                refIndex++;
                stuIndex++;
            }

            int extraWord = countByType(wordMistakes, "Extra Word");
            int missingWord = countByType(wordMistakes, "Missing Word");
            int wrongWord = countByType(wordMistakes, "Wrong Word");

            int groupedMistakes = mistakeCounts.getOrDefault("Extra Enter", 0) +
                    mistakeCounts.getOrDefault("Extra Space", 0) +
                    (mistakeCounts.getOrDefault("Extra Tab", 0) / 3);

            int totalMistakes = extraWord + missingWord + wrongWord + groupedMistakes;
            int marks = Math.max(0, 40 - totalMistakes);

            XWPFParagraph summary = outputDoc.createParagraph();
            addTextRun(summary, "\nSummary of Mistakes:", false);
            if (mistakeCounts.get("Extra Space") != null)
                addTextRun(summary, "Extra Space: " + mistakeCounts.get("Extra Space"), true);
            if (mistakeCounts.get("Extra Tab") != null)
                addTextRun(summary, "Extra Tab: " + mistakeCounts.get("Extra Tab"), true);
            if (mistakeCounts.get("Extra Enter") != null)
                addTextRun(summary, "Extra Enter: " + mistakeCounts.get("Extra Enter"), true);
            if (extraWord > 0) addTextRun(summary, "Extra Word: " + extraWord, false);
            if (missingWord > 0) addTextRun(summary, "Missing Word: " + missingWord, false);
            if (wrongWord > 0) addTextRun(summary, "Wrong Word: " + wrongWord, false);
            addTextRun(summary, "Total Mistakes: " + totalMistakes, false);
            addTextRun(summary, "Final Marks: " + marks + " / 40", false);

            File outFile = new File(compareDir, studentFile.getName().replace(".docx", "_Compared.docx"));
            FileOutputStream out = null;
            try {
                out = new FileOutputStream(outFile);
                outputDoc.write(out);
            } finally {
                if (out != null) out.close();
            }

            Row row = sheet.createRow(rowNum++);
            row.createCell(0).setCellValue(studentFile.getName());
            row.createCell(1).setCellValue(totalMistakes);
            row.createCell(2).setCellValue(marks);
            row.createCell(3).setCellValue(mistakeCounts.getOrDefault("Extra Space", 0));
            row.createCell(4).setCellValue(mistakeCounts.getOrDefault("Extra Tab", 0));
            row.createCell(5).setCellValue(mistakeCounts.getOrDefault("Extra Enter", 0));
            row.createCell(6).setCellValue(extraWord);
            row.createCell(7).setCellValue(missingWord);
            row.createCell(8).setCellValue(wrongWord);

            StringBuilder sb = new StringBuilder();
            for (String[] m : wordMistakes) sb.append(m[1]).append(" - ").append(m[2]).append(" | ");
            row.createCell(9).setCellValue(sb.toString());
        }

        for (int i = 0; i < columns.length; i++) sheet.autoSizeColumn(i);
        FileOutputStream excelOut = null;
        try {
            excelOut = new FileOutputStream(new File(folder.getParent(), "Comparison_Result.xlsx"));
            workbook.write(excelOut);
        } finally {
            if (excelOut != null) excelOut.close();
        }
        workbook.close();
        System.out.println("\u2705 All students processed.");
    }

    private static void addTextRun(XWPFParagraph para, String text, boolean highlight) {
        XWPFRun run = para.createRun();
        run.setText(text + " ");
        if (highlight) run.setColor("FF0000");
    }

    private static void compareWords(String model, String student, XWPFParagraph para,
                                     Map<String, Integer> mistakeCounts,
                                     List<String[]> wordMistakes, String fileName) {
        String[] a = model.trim().split("\\s+");
        String[] b = student.trim().split("\\s+");
        int i = 0, j = 0;
        while (i < a.length || j < b.length) {
            String aw = i < a.length ? normalizeWord(a[i]) : null;
            String bw = j < b.length ? normalizeWord(b[j]) : null;
            if (aw != null && bw != null && aw.equals(bw)) {
                para.createRun().setText(b[j] + " ");
                i++; j++;
            } else if (bw != null && (j + 1 < b.length && normalizeWord(b[j + 1]).equals(aw))) {
                addTextRun(para, "[" + b[j] + "]", true);
                wordMistakes.add(new String[]{fileName, "Extra Word", b[j]});
                j++;
            } else if (aw != null && (i + 1 < a.length && normalizeWord(a[i + 1]).equals(bw))) {
                addTextRun(para, "[" + a[i] + "]", true);
                wordMistakes.add(new String[]{fileName, "Missing Word", a[i]});
                i++;
            } else if (aw != null && bw != null) {
                addTextRun(para, "[" + b[j] + "/" + a[i] + "]", true);
                wordMistakes.add(new String[]{fileName, "Wrong Word", b[j] + "/" + a[i]});
                i++; j++;
            } else if (bw != null) {
                addTextRun(para, "[" + b[j] + "]", true);
                wordMistakes.add(new String[]{fileName, "Extra Word", b[j]});
                j++;
            } else if (aw != null) {
                addTextRun(para, "[" + a[i] + "]", true);
                wordMistakes.add(new String[]{fileName, "Missing Word", a[i]});
                i++;
            }
        }
    }

    private static String normalizeWord(String word) {
        if (word == null) return null;
        return word.replace("’", "'");
    }

    private static void incrementMistake(String key, Map<String, Integer> map) {
        map.put(key, map.containsKey(key) ? map.get(key) + 1 : 1);
    }

    private static int countByType(List<String[]> list, String type) {
        int count = 0;
        for (String[] arr : list) {
            if (arr[1].equals(type)) count++;
        }
        return count;
    }

    private static int countMatches(String text, String regex) {
        return Pattern.compile(regex).matcher(text).find() ? 1 : 0;
    }

    private static List<String> extractFormattedParagraphs(File file) throws IOException {
        List<String> paras = new ArrayList<String>();
        FileInputStream fis = null;
        XWPFDocument doc = null;
        try {
            fis = new FileInputStream(file);
            doc = new XWPFDocument(fis);
            for (XWPFParagraph p : doc.getParagraphs()) {
                StringBuilder fullText = new StringBuilder();
                for (XWPFRun run : p.getRuns()) fullText.append(run.toString());
                paras.add(fullText.toString());
            }
        } finally {
            if (doc != null) doc.close();
            if (fis != null) fis.close();
        }
        return paras;
    }
}
