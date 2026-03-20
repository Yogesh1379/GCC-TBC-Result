package StudentData_Verification;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;

public class ExcelMismatchBothValues1 {

    static final String[] COLS = {
            "INSTID", "STUDENT_FIRST_NAME", "STUDENT_MIDDLE_NAME",
            "STUDENT_LAST_NAME", "MOTHERS_NAME", "GENDER",
            "HANDICAP", "DATE_OF_BIRTH", "COURSE_SUBJECT"
    };

    static final int COL_COUNT = COLS.length;

    public static void main(String[] args) throws Exception {

        String file1 = "F:\\GCC  TBC December 2025\\marathi\\Student data\\Marathi data from support team.xlsx";
        String file2 =  "F:\\GCC  TBC December 2025\\marathi\\Student data\\Marathi data from msce.xlsx";
        String output = "F:\\GCC  TBC December 2025\\marathi\\Student data\\Comparison_Report1.xlsx";

        Map<String, Row> file2Map = loadFile2Rows(file2);

        Workbook wb1 = new XSSFWorkbook(new FileInputStream(file1));
        Sheet s1 = wb1.getSheetAt(0);

        Workbook outWb = new XSSFWorkbook();
        Sheet matched = outWb.createSheet("MATCHED");
        Sheet mismatched = outWb.createSheet("MISMATCHED");

        createMatchedHeader(matched);
        createMismatchHeader(mismatched);

        int mRow = 1, mmRow = 1;

        for (int i = 1; i <= s1.getLastRowNum(); i++) {
            Row r1 = s1.getRow(i);
            if (r1 == null) continue;

            String key = buildBaseKey(r1);

            if (!file2Map.containsKey(key)) {
                writeMismatchRow(r1, null, mismatched.createRow(mmRow++),
                        "MISMATCHED", "RECORD_NOT_FOUND_IN_FILE2");
                continue;
            }

            Row r2 = file2Map.get(key);
            List<String> diffs = new ArrayList<>();

            for (int c = 0; c < COL_COUNT; c++) {
                if (!get(r1, c).equalsIgnoreCase(get(r2, c))) {
                    diffs.add(COLS[c]);
                }
            }

            if (diffs.isEmpty()) {
                writeMatchedRow(r1, matched.createRow(mRow++));
            } else {
                writeMismatchRow(r1, r2, mismatched.createRow(mmRow++),
                        "MISMATCHED", String.join(", ", diffs));
            }
        }

        wb1.close();

        FileOutputStream fos = new FileOutputStream(output);
        outWb.write(fos);
        outWb.close();
        fos.close();

        System.out.println("Excel comparison completed.");
    }

    // ---------------- HELPERS ----------------

    static Map<String, Row> loadFile2Rows(String path) throws Exception {
        Map<String, Row> map = new HashMap<>();

        Workbook wb = new XSSFWorkbook(new FileInputStream(path));
        Sheet s = wb.getSheetAt(0);

        for (int i = 1; i <= s.getLastRowNum(); i++) {
            Row r = s.getRow(i);
            if (r != null) {
                map.put(buildBaseKey(r), r);
            }
        }
        return map;
    }

    // Base identity (without COURSE_SUBJECT to detect mismatches)
    static String buildBaseKey(Row r) {
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < COL_COUNT - 1; i++) {
            sb.append(get(r, i).toLowerCase()).append("|");
        }
        return sb.toString();
    }

    static void createMatchedHeader(Sheet s) {
        Row h = s.createRow(0);
        int c = 0;
        for (String col : COLS) {
            h.createCell(c++).setCellValue(col);
        }
        h.createCell(c).setCellValue("STATUS");
    }

    static void createMismatchHeader(Sheet s) {
        Row h = s.createRow(0);
        int c = 0;

        for (String col : COLS) {
            h.createCell(c++).setCellValue(col + "_F1");
        }
        for (String col : COLS) {
            h.createCell(c++).setCellValue(col + "_F2");
        }
        h.createCell(c++).setCellValue("STATUS");
        h.createCell(c).setCellValue("MISMATCH_REASON");
    }

    static void writeMatchedRow(Row src, Row tgt) {
        int c = 0;
        for (int i = 0; i < COL_COUNT; i++) {
            tgt.createCell(c++).setCellValue(get(src, i));
        }
        tgt.createCell(c).setCellValue("MATCHED");
    }

    static void writeMismatchRow(Row r1, Row r2, Row tgt,
                                 String status, String reason) {
        int c = 0;

        for (int i = 0; i < COL_COUNT; i++) {
            tgt.createCell(c++).setCellValue(get(r1, i));
        }
        for (int i = 0; i < COL_COUNT; i++) {
            tgt.createCell(c++).setCellValue(r2 == null ? "" : get(r2, i));
        }
        tgt.createCell(c++).setCellValue(status);
        tgt.createCell(c).setCellValue(reason);
    }

    static String get(Row r, int i) {
        if (r == null || r.getCell(i) == null) return "";
        r.getCell(i).setCellType(CellType.STRING);
        return r.getCell(i).getStringCellValue().trim();
    }
}
