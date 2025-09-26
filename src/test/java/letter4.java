import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.io.*;
import java.util.*;

public class letter4 {
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
        model   = model.trim().replaceAll("\\s+([,;:!?])","$1");
        student = student.trim().replaceAll("\\s+([,;:!?])","$1");
        model   = model.replaceAll("\\.(?!\\s|$)",". ");
        student = student.replaceAll("\\.(?!\\s|$)",". ");

        String[] mW = model.split("\\s+");
        String[] sW = student.split("\\s+");
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
}
