package StudentData_Verification;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;

public class ExcelCompareBothValues_FIXED {

    // ===== PRIMARY KEY COLUMNS (MATCHING) =====
    static final int[] KEY_COLS = {0, 1, 2, 3}; // INSTID, FIRST, MIDDLE, LAST

    // ===== COLUMNS TO COMPARE =====
    static final int[] COMPARE_COLS = {
            0, 1, 2, 3, 4, 7, 10, 13
    };

    static final String[] COL_NAMES = {
            "INSTID", "STUDENT_FIRST_NAME", "STUDENT_MIDDLE_NAME",
            "STUDENT_LAST_NAME", "MOTHERS_NAME",
            "GENDER", "DATE_OF_BIRTH", "COURSE_SUBJECT"
    };

    static final int COL_COUNT = COMPARE_COLS.length;

    public static void main(String[] args) throws Exception {

        String file1Path = "F:\\GCC  TBC December 2025\\marathi\\Student data\\Marathi data from support team.xlsx";
        String file2Path = "F:\\GCC  TBC December 2025\\marathi\\Student data\\Marathi data from msce.xlsx";
        String outputPath = "F:\\GCC  TBC December 2025\\marathi\\Student data\\Comparison_Report2_FIXED.xlsx";

        // Load File2 into Map
        Map<String, List<String>> file2Map = loadFileData(file2Path);

        Workbook wb1 = new XSSFWorkbook(new FileInputStream(file1Path));
        Sheet sheet1 = wb1.getSheetAt(0);

        Workbook outWb = new XSSFWorkbook();
        Sheet matchedSheet = outWb.createSheet("MATCHED");
        Sheet mismatchedSheet = outWb.createSheet("MISMATCHED");

        createMatchedHeader(matchedSheet);
        createMismatchHeader(mismatchedSheet);

        int matchRow = 1;
        int mismatchRow = 1;

        for (int i = 1; i <= sheet1.getLastRowNum(); i++) {
            Row row1 = sheet1.getRow(i);
            if (row1 == null) continue;

            String key = buildPrimaryKey(row1);
            List<String> row2 = file2Map.get(key);

            if (row2 == null) {
                // Not found
                writeMismatchRow(row1, null,
                        mismatchedSheet.createRow(mismatchRow++),
                        "RECORD_NOT_FOUND_IN_FILE2");
            } else {
                // Compare column-wise
                List<String> mismatches = new ArrayList<>();

                for (int c = 0; c < COL_COUNT; c++) {
                    String v1 = getCellValue(row1.getCell(COMPARE_COLS[c])).toLowerCase();
                    String v2 = row2.get(c).toLowerCase();
                    if (!v1.equals(v2)) {
                        mismatches.add(COL_NAMES[c]);
                    }
                }

                if (mismatches.isEmpty()) {
                    writeMatchedRow(row1, matchedSheet.createRow(matchRow++));
                } else {
                    writeMismatchRow(row1, row2,
                            mismatchedSheet.createRow(mismatchRow++),
                            String.join(", ", mismatches));
                }
            }
        }

        wb1.close();

        try (FileOutputStream fos = new FileOutputStream(outputPath)) {
            outWb.write(fos);
        }
        outWb.close();

        System.out.println("✅ Comparison report generated successfully.");
    }

    // ================= HELPERS =================

    static Map<String, List<String>> loadFileData(String path) throws Exception {
        Map<String, List<String>> map = new HashMap<>();
        Workbook wb = new XSSFWorkbook(new FileInputStream(path));
        Sheet sheet = wb.getSheetAt(0);

        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row == null) continue;

            map.put(buildPrimaryKey(row), extractRow(row));
        }
        wb.close();
        return map;
    }

    // PRIMARY KEY (ONLY FEW COLUMNS)
    static String buildPrimaryKey(Row row) {
        StringBuilder sb = new StringBuilder();
        for (int col : KEY_COLS) {
            sb.append(getCellValue(row.getCell(col)).toLowerCase().trim()).append("|");
        }
        return sb.toString();
    }

    static List<String> extractRow(Row row) {
        List<String> list = new ArrayList<>();
        for (int c : COMPARE_COLS) {
            list.add(getCellValue(row.getCell(c)));
        }
        return list;
    }

    // ================= HEADERS =================

    static void createMatchedHeader(Sheet sheet) {
        Row h = sheet.createRow(0);
        int c = 0;
        for (String col : COL_NAMES) {
            h.createCell(c++).setCellValue(col);
        }
        h.createCell(c).setCellValue("STATUS");
    }

    static void createMismatchHeader(Sheet sheet) {
        Row h = sheet.createRow(0);
        int c = 0;

        for (String col : COL_NAMES) h.createCell(c++).setCellValue(col + "_FILE1");
        for (String col : COL_NAMES) h.createCell(c++).setCellValue(col + "_FILE2");

        h.createCell(c++).setCellValue("STATUS");
        h.createCell(c).setCellValue("MISMATCH_REASON");
    }

    // ================= WRITE =================

    static void writeMatchedRow(Row src, Row tgt) {
        int c = 0;
        for (int col : COMPARE_COLS) {
            tgt.createCell(c++).setCellValue(getCellValue(src.getCell(col)));
        }
        tgt.createCell(c).setCellValue("MATCHED");
    }

    static void writeMismatchRow(Row row1, List<String> row2,
                                 Row tgt, String reason) {
        int c = 0;

        for (int col : COMPARE_COLS) {
            tgt.createCell(c++).setCellValue(getCellValue(row1.getCell(col)));
        }

        if (row2 != null) {
            for (String v : row2) tgt.createCell(c++).setCellValue(v);
        } else {
            for (int i = 0; i < COL_COUNT; i++) tgt.createCell(c++).setCellValue("");
        }

        tgt.createCell(c++).setCellValue("MISMATCHED");
        tgt.createCell(c).setCellValue(reason);
    }

    static String getCellValue(Cell cell) {
        if (cell == null) return "";
        return new DataFormatter().formatCellValue(cell).trim();
    }
}
