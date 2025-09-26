import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import java.util.regex.Pattern;

import java.io.*;
import java.util.*;

public class letter7 {
    static File modelFile = null;
    static List<String> modelParas;

    @SuppressWarnings("resource")
    public static void main(String[] args) throws Exception {
        File folder = new File("F:\\desktop backup 5 july 25\\GCC TBC MAY 2025 AUTO\\New folder\\New folder");
        File compareDir = new File(folder.getParentFile(), folder.getName() + "_Compared");
        if (!compareDir.exists()) compareDir.mkdirs();

        Workbook resultWb   = new XSSFWorkbook();
        Sheet    resultSheet = resultWb.createSheet("Results");
        String[] columns     = {
            "File Name", "Extra Word", "Missing Word", "Wrong Word",
            "Total Mistakes", "Final Marks", "Mistakes Summary"
        };
        Row header = resultSheet.createRow(0);
        for (int i = 0; i < columns.length; i++) {
            header.createCell(i).setCellValue(columns[i]);
        }

        File[] docxFiles = folder.listFiles((dir, name) ->
            name.toLowerCase().endsWith(".docx") && !name.startsWith("~$")
        );
        if (docxFiles == null) {
            System.out.println("❌ No .docx files found.");
            return;
        }

        int    rowNum     = 1;
        double MAX_MARKS  = 7.5;

        for (File studentFile : docxFiles) {
            // determine modelFile
            String[] parts = studentFile.getName().split("_");
            if (parts.length < 3) continue;
            String seatno    = parts[1],
                   batchname = parts[2];
            System.out.println(seatno);
            if (seatno.length() >= 10) {
                String course = seatno.substring(4, 6);
                int course1 = course.equals("15") ? 1 : course.equals("16") ? 2 : 0;
                try (FileInputStream fis = new FileInputStream(
                        "F:\\desktop backup 5 july 25\\GCC TBC MAY 2025 AUTO\\allocation\\BathwiseSubjective.xlsx"
                )) {
                    Workbook allocWb = new XSSFWorkbook(fis);
                    Sheet    allocSheet = allocWb.getSheetAt(0);
                    for (Row r : allocSheet) {
                        Cell bc = r.getCell(6), cc = r.getCell(5), sc = r.getCell(2);
                        if (bc == null || cc == null || sc == null) continue;
                        if (!getCellValueAsString(bc).equals(batchname)) continue;
                        if (Integer.parseInt(getCellValueAsString(cc)) != course1) continue;
                        String cand = getCellValueAsString(sc);
                        if ((cand.startsWith("Eng30 Ltr")  && course1 == 1) ||
                            ((cand.startsWith("Eng40 Resume") || cand.startsWith("Eng40 B Ltr")) && course1 == 2)) {
                            System.out.println(cand);
                            modelFile = new File(
                              "F:\\desktop backup 5 july 25\\GCC TBC MAY 2025 AUTO\\allocation\\New folder\\Model Anwers\\" + cand
                            );
                            break;
                        }
                    }
                    allocWb.close();
                }
            }

            // extract paragraphs & table lines
            modelParas     = extractParasAndTables(modelFile);
            List<String> studentParas = extractParasAndTables(studentFile);

            List<String[]> wordMistakes = new ArrayList<>();
            XWPFDocument outputDoc = new XWPFDocument();

            // --- compare 2nd-letter addresses (between "To," and "Subject:") ---
            String modelAddress2 = extractNthAddress(modelParas, 2);
            String studentAddress2 = extractNthAddress(studentParas, 2);
            if (!modelAddress2.isEmpty() || !studentAddress2.isEmpty()) {
                XWPFParagraph addrHdr = outputDoc.createParagraph();
                addrHdr.createRun().setBold(true);
                addrHdr.createRun().setText("Address (2nd letter) comparison:");
                XWPFParagraph addrPara = outputDoc.createParagraph();
                compareAddresses(modelAddress2, studentAddress2, addrPara, wordMistakes, studentFile.getName());
            }

            // compare up through and including the first "Encl" line
            int maxLines = Math.max(modelParas.size(), studentParas.size());
            for (int i = 0; i < maxLines; i++) {
                String mLine = i < modelParas.size()   ? modelParas.get(i)   : "";
                String sLine = i < studentParas.size() ? studentParas.get(i) : "";
                XWPFParagraph p = outputDoc.createParagraph();
                compareWords(mLine, sLine, p, wordMistakes, studentFile.getName());
                if (mLine.contains("Encl") || sLine.contains("Encl")) {
                    break;
                }
            }

            // tally
            int extra = 0, missing = 0, wrong = 0;
            for (String[] m : wordMistakes) {
                switch (m[1]) {
                    case "Extra Word":   extra++;   break;
                    case "Missing Word": missing++; break;
                    case "Wrong Word":   wrong++;   break;
                }
            }
            int   totalMistakes = wordMistakes.size();
            double obtainedMarks = Math.max(0.0, MAX_MARKS - totalMistakes * 0.5);

            // summary in DOCX
            XWPFParagraph sumPara = outputDoc.createParagraph();
            if (extra   > 0) sumPara.createRun().setText("Extra Word: " + extra);
            if (missing > 0) sumPara.createRun().setText("Missing Word: " + missing);
            if (wrong   > 0) sumPara.createRun().setText("Wrong Word: " + wrong);
            sumPara.createRun().setText("Total Mistakes: " + totalMistakes);
            sumPara.createRun().setText(
                String.format("Final Marks: %.1f / %.1f", obtainedMarks, MAX_MARKS)
            );

            File comparedFile = new File(compareDir,
                studentFile.getName().replace(".docx", "_Compared.docx")
            );
            try (FileOutputStream out = new FileOutputStream(comparedFile)) {
                outputDoc.write(out);
            }

            // write Excel row
            Row row = resultSheet.createRow(rowNum++);
            row.createCell(0).setCellValue(studentFile.getName());
            row.createCell(1).setCellValue(extra);
            row.createCell(2).setCellValue(missing);
            row.createCell(3).setCellValue(wrong);
            row.createCell(4).setCellValue(totalMistakes);
            row.createCell(5).setCellValue(obtainedMarks);
            StringBuilder sb = new StringBuilder();
            for (String[] m : wordMistakes) {
                sb.append("[").append(m[1]).append(" - ").append(m[2]).append("] ");
            }
            row.createCell(6).setCellValue(sb.toString().trim());
        }

        // autosize & save
        for (int i = 0; i < columns.length; i++) resultSheet.autoSizeColumn(i);
        try (FileOutputStream out = new FileOutputStream(
                new File(folder.getParent(), "Comparison_Result.xlsx")
        )) {
            resultWb.write(out);
        }
        resultWb.close();
        System.out.println("✅ All students processed.");
    }

    private static List<String> extractParasAndTables(File file) throws IOException {
        List<String> paras = new ArrayList<>();
        try (FileInputStream fis = new FileInputStream(file);
             XWPFDocument doc = new XWPFDocument(fis)) {

            for (IBodyElement e : doc.getBodyElements()) {
                if (e instanceof XWPFParagraph) {
                    paras.addAll(splitRunsIntoLines((XWPFParagraph)e));
                } else if (e instanceof XWPFTable) {
                    addTableLines((XWPFTable)e, paras);
                }
            }
            for (XWPFTable t : doc.getTables()) {
                addTableLines(t, paras);
            }
        }
        return paras;
    }

    private static void addTableLines(XWPFTable table, List<String> paras) {
        for (XWPFTableRow row : table.getRows()) {
            for (XWPFTableCell cell : row.getTableCells()) {
                for (XWPFParagraph p : cell.getParagraphs()) {
                    paras.addAll(splitRunsIntoLines(p));
                }
            }
        }
    }

    private static List<String> splitRunsIntoLines(XWPFParagraph p) {
        StringBuilder sb = new StringBuilder();
        for (XWPFRun run : p.getRuns()) {
            sb.append(run.toString().replace("\t","    "));
            int brCount = run.getCTR().getBrList().size();
            for (int i=0; i<brCount; i++) sb.append("\n");
        }
        List<String> out = new ArrayList<>();
        for (String line : sb.toString().split("\\r?\\n")) {
            if (!line.trim().isEmpty()) out.add(line.trim());
        }
        return out;
    }

    private static void compareWords(String model, String student,
                                     XWPFParagraph para,
                                     List<String[]> mistakes,
                                     String fileName) {
        String modelOrig = model == null ? "" : model.trim();
        String studentOrig = student == null ? "" : student.trim();

        // Ensure student spacing mirrors model before tokenization/comparison
        studentOrig = ensureStudentSpacingFromModel(modelOrig, studentOrig);

        // Now apply your existing normalization rules (space before punctuation removal, single space after period etc.)
        modelOrig   = modelOrig.replaceAll("\\s+([,;:!?])","$1");
        studentOrig = studentOrig.replaceAll("\\s+([,;:!?])","$1");

        // ensure exactly one space after '.' and ':' (and collapse multiples)
        modelOrig = modelOrig.replaceAll("\\.\\s*", ". ").replaceAll(":\\s*", ": ");
        studentOrig = studentOrig.replaceAll("\\.\\s*", ". ").replaceAll(":\\s*", ": ");

        String[] mW = modelOrig.isEmpty() ? new String[0] : modelOrig.split("\\s+");
        String[] sW = studentOrig.isEmpty() ? new String[0] : studentOrig.split("\\s+");
        List<int[]> lcs = computeLCS(mW, sW);

        int i=0, j=0, k=0;
        while (i<mW.length || j<sW.length) {
            if (k<lcs.size() && i==lcs.get(k)[0] && j==lcs.get(k)[1]) {
                addTextRun(para, sW[j]+" ", false);
                i++; j++; k++;
            }
            else if (i<mW.length && j<sW.length &&
                     (k>=lcs.size() || (i<lcs.get(k)[0] && j<lcs.get(k)[1])) &&
                     !normalize(mW[i]).equals(normalize(sW[j]))) {
                highlight(para.createRun(),"["+mW[i]+"/"+sW[j]+"]");
                mistakes.add(new String[]{fileName,"Wrong Word",mW[i]+"/"+sW[j]});
                i++; j++;
            }
            else if (k<lcs.size() && i<lcs.get(k)[0]) {
                highlight(para.createRun(),"["+mW[i]+"]");
                mistakes.add(new String[]{fileName,"Missing Word",mW[i]});
                i++;
            }
            else if (k<lcs.size() && j<lcs.get(k)[1]) {
                highlight(para.createRun(),"["+sW[j]+"]");
                mistakes.add(new String[]{fileName,"Extra Word",sW[j]});
                j++;
            }
            else if (i<mW.length) {
                highlight(para.createRun(),"["+mW[i]+"]");
                mistakes.add(new String[]{fileName,"Missing Word",mW[i]});
                i++;
            }
            else if (j<sW.length) {
                highlight(para.createRun(),"["+sW[j]+"]");
                mistakes.add(new String[]{fileName,"Extra Word",sW[j]});
                j++;
            }
            else {
                i++; j++;
            }
        }
    }

    private static List<int[]> computeLCS(String[] a, String[] b) {
        int m=a.length, n=b.length;
        int[][] dp = new int[m+1][n+1];
        for (int i=1; i<=m; i++) {
            for (int j=1; j<=n; j++) {
                if (normalize(a[i-1]).equals(normalize(b[j-1])))
                    dp[i][j] = dp[i-1][j-1] + 1;
                else
                    dp[i][j] = Math.max(dp[i-1][j], dp[i][j-1]);
            }
        }
        List<int[]> lcs = new ArrayList<>();
        int i=m, j=n;
        while (i>0 && j>0) {
            if (normalize(a[i-1]).equals(normalize(b[j-1]))) {
                lcs.add(0, new int[]{i-1,j-1}); i--; j--;
            } else if (dp[i-1][j]>=dp[i][j-1]) i--; else j--;
        }
        return lcs;
    }

    private static String getCellValueAsString(Cell cell) {
        switch (cell.getCellType()) {
            case STRING:  return cell.getStringCellValue().trim();
            case NUMERIC: return String.valueOf((int)cell.getNumericCellValue());
            default:      return "";
        }
    }

    private static String normalize(String w) {
        return w==null? "": w.trim().replace("’","'").replace("‘","'");
    }

    /**
     * Ensure student spacing around punctuation matches model's pattern.
     * For each punctuation in {.,;:!?,-}:
         */
    private static String ensureStudentSpacingFromModel(String model, String student) {
        if (student == null) student = "";
        if (model == null) model = "";

        // punctuation set
        char[] puncts = {'.', ',', ';', ':', '!', '?', '-'};

        for (char p : puncts) {
            String pQ = Pattern.quote(String.valueOf(p)); // safe regex-escaped punctuation

            // does model show a space before this punctuation anywhere? (e.g. " word ,")
            boolean modelHasSpaceBefore = Pattern.compile("\\s" + pQ).matcher(model).find();

            // does model show a space after this punctuation anywhere? (e.g. ", word")
            boolean modelHasSpaceAfter  = Pattern.compile(pQ + "\\s").matcher(model).find();

            // If model requires space BEFORE punctuation, add space before any punctuation in student that lacks it
            if (modelHasSpaceBefore) {
                // replace punctuation not preceded by whitespace with " <punct>"
                student = student.replaceAll("(?<!\\s)" + pQ, " " + p);
            }

            // If model requires space AFTER punctuation, add a single space after any punctuation in student that lacks it
            if (modelHasSpaceAfter) {
                // replace punctuation not followed by whitespace with "<punct> "
                student = student.replaceAll(pQ + "(?!\\s)", p + " ");
            }
        }

        // Ensure exactly one space after '.' and ':' (user requested only one space there)
//        student = student.replaceAll("\\.\\s*", ". ");
//        student = student.replaceAll(":\\s*", ": ");

        // Collapse multiple spaces to single and trim
        student = student.replaceAll("\\s+", " ").trim();

        return student;
    }

    private static void addTextRun(XWPFParagraph para, String text, boolean highlight) {
        XWPFRun run = para.createRun();
        run.setText(text);
        if (highlight) {
            CTRPr rpr = run.getCTR().isSetRPr()? run.getCTR().getRPr(): run.getCTR().addNewRPr();
            CTShd shd = rpr.addNewShd();
            shd.setVal(STShd.CLEAR);
            shd.setFill("FF0000");
        }
    }

    private static void highlight(XWPFRun run, String text) {
        run.setText(text + " ");
        run.setColor("FF0000");
    }

   
     
    private static String extractNthAddress(List<String> paras, int n) {
        if (paras == null || paras.isEmpty() || n <= 0) return "";
        int count = 0;
        boolean collecting = false;
        StringBuilder sb = new StringBuilder();

        for (String raw : paras) {
            if (raw == null) continue;
            String line = raw.trim();
            if (!collecting) {
                // Detect a standalone "To" line (e.g. "To," or "To") OR inline "To, Name..."
                if (line.matches("(?i)^to\\s*,?$")) {
                    count++;
                    if (count == n) {
                        collecting = true;
                        continue; // next lines are part of address
                    }
                } else if (line.toLowerCase().startsWith("to,") || line.toLowerCase().startsWith("to ")) {
                    // Inline To followed by address on same line
                    count++;
                    if (count == n) {
                        String after = line.replaceFirst("(?i)^to\\s*,?\\s*", "");
                        if (!after.isEmpty()) sb.append(after).append(" ");
                        collecting = true;
                        continue;
                    }
                }
            } else {
                // We are collecting address lines until we hit a "Subject" marker
                if (line.toLowerCase().contains("subject")) {
                    break;
                }
                if (!line.isEmpty()) {
                    sb.append(line).append(" ");
                }
            }
        }
        return sb.toString().trim();
    }

    /**
     * Compare two addresses word-by-word and write highlighted differences into para.
     * Differences are added to mistakes list using same format as compareWords.
     */
    private static void compareAddresses(String modelAddr, String studentAddr,
                                         XWPFParagraph para,
                                         List<String[]> mistakes,
                                         String fileName) {

        String modelOrig = modelAddr == null ? "" : modelAddr.trim();
        String studentOrig = studentAddr == null ? "" : studentAddr.trim();

        // adjust student spacing to match model punctuation spacing rules
        studentOrig = ensureStudentSpacingFromModel(modelOrig, studentOrig);

        // Normalize punctuation spacing similar to compareWords
        modelOrig = modelOrig.replaceAll("\\s+([,;:!?])","$1").replaceAll("\\.(?!\\s|$)",". ");
        studentOrig = studentOrig.replaceAll("\\s+([,;:!?])","$1").replaceAll("\\.(?!\\s|$)",". ");

        para.createRun().setText("  (Model) " + modelOrig);
        para.createRun().addBreak();
        para.createRun().setText("  (Student) ");

        String[] mW = modelOrig.isEmpty() ? new String[0] : modelOrig.split("\\s+");
        String[] sW = studentOrig.isEmpty() ? new String[0] : studentOrig.split("\\s+");

        List<int[]> lcs = computeLCS(mW, sW);

        int i = 0, j = 0, k = 0;
        while (i < mW.length || j < sW.length) {
            if (k < lcs.size() && i == lcs.get(k)[0] && j == lcs.get(k)[1]) {
                addTextRun(para, sW[j] + " ", false);
                i++; j++; k++;
            }
            else if (i < mW.length && j < sW.length &&
                    (k >= lcs.size() || (i < lcs.get(k)[0] && j < lcs.get(k)[1])) &&
                     !normalize(mW[i]).equals(normalize(sW[j]))) {
                // same as 'Wrong Word' in lines compare
                highlight(para.createRun(), "[" + mW[i] + "/" + sW[j] + "]");
                mistakes.add(new String[]{fileName, "Wrong Word", mW[i] + "/" + sW[j]});
                i++; j++;
            }
            else if (k < lcs.size() && i < lcs.get(k)[0]) {
                highlight(para.createRun(), "[" + mW[i] + "]");
                mistakes.add(new String[]{fileName, "Missing Word", mW[i]});
                i++;
            }
            else if (k < lcs.size() && j < lcs.get(k)[1]) {
                highlight(para.createRun(), "[" + sW[j] + "]");
                mistakes.add(new String[]{fileName, "Extra Word", sW[j]});
                j++;
            }
            else if (i < mW.length) {
                highlight(para.createRun(), "[" + mW[i] + "]");
                mistakes.add(new String[]{fileName, "Missing Word", mW[i]});
                i++;
            }
            else if (j < sW.length) {
                highlight(para.createRun(), "[" + sW[j] + "]");
                mistakes.add(new String[]{fileName, "Extra Word", sW[j]});
                j++;
            }
            else {
                i++; j++;
            }
        }
        para.createRun().addBreak();
    }
    
}
