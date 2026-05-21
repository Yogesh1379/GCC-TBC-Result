package BatchWise_Compare_With_Result;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;

public class ExcelSeatComparison {

    public static void main(String[] args) throws Exception {

        String file1Path = "F:\\GCCTBC-APR 2026\\Batchwise\\final\\HIN-40\\Batchwise_HIN-40-APRIL-2026.xlsx";// batchwise
        String file2Path = "F:\\GCCTBC-APR 2026\\Result\\marathi\\TentativeResult3.xlsx";  /// result 
        String outputPath = "F:\\GCCTBC-APR 2026\\Batchwise\\final\\HIN-40\\Mismatch_Report hindi40.xlsx";

        FileInputStream fis1 = new FileInputStream(file1Path);
        FileInputStream fis2 = new FileInputStream(file2Path);

        Workbook wb1 = new XSSFWorkbook(fis1);
        Workbook wb2 = new XSSFWorkbook(fis2);

        Sheet sheet1 = wb1.getSheetAt(0);
        Sheet sheet2 = wb2.getSheetAt(1);

        // Columns to compare
        List<String> columnsToCompare = Arrays.asList(
                
                "LetterTotalMarks",
                "SpeedMistakeMarks",
                 "ExcelTotalMarks",
                "Email Marks",
                "Objective Marks"
        );

        // Get header mapping for both files
        Map<String, Integer> headerMap1 = getHeaderMap(sheet1);
        Map<String, Integer> headerMap2 = getHeaderMap(sheet2);

        int seatCol1 = headerMap1.get("Seat No.");
        int seatCol2 = headerMap2.get("Seat No.");

        // Store File2 data in Map for fast lookup
        Map<String, Row> file2Data = new HashMap<>();
        for (int i = 1; i <= sheet2.getLastRowNum(); i++) {
            Row row = sheet2.getRow(i);
            if (row != null) {
                String seat = getCellValue(row.getCell(seatCol2));
                file2Data.put(seat, row);
            }
        }

        // Create output workbook
        Workbook outputWb = new XSSFWorkbook();
        Sheet outputSheet = outputWb.createSheet("Mismatches");

        int outputRowNum = 0;

        // Create header row
        Row header = outputSheet.createRow(outputRowNum++);
        header.createCell(0).setCellValue("Seat No");
        header.createCell(1).setCellValue("Column Name");
        header.createCell(2).setCellValue("Batchwise");
        header.createCell(3).setCellValue("Result");

        // Compare rows
        for (int i = 1; i <= sheet1.getLastRowNum(); i++) {
            Row row1 = sheet1.getRow(i);
            if (row1 == null) continue;

            String seat = getCellValue(row1.getCell(seatCol1));

            if (!file2Data.containsKey(seat)) {
                Row outRow = outputSheet.createRow(outputRowNum++);
                outRow.createCell(0).setCellValue(seat);
                outRow.createCell(1).setCellValue("Seat Missing in File2");
                outRow.createCell(2).setCellValue("Present");
                outRow.createCell(3).setCellValue("Not Found");
                continue;
            }

            Row row2 = file2Data.get(seat);

            for (String colName : columnsToCompare) {

                Integer colIndex1 = headerMap1.get(colName);
                Integer colIndex2 = headerMap2.get(colName);

                if (colIndex1 == null || colIndex2 == null) continue;

                String value1 = getCellValue(row1.getCell(colIndex1));
                String value2 = getCellValue(row2.getCell(colIndex2));

                if (!Objects.equals(value1, value2)) {
                    Row outRow = outputSheet.createRow(outputRowNum++);
                    outRow.createCell(0).setCellValue(seat);
                    outRow.createCell(1).setCellValue(colName);
                    outRow.createCell(2).setCellValue(value1);
                    outRow.createCell(3).setCellValue(value2);
                }
            }
        }

        // Write output file
        FileOutputStream fos = new FileOutputStream(outputPath);
        outputWb.write(fos);

        fos.close();
        wb1.close();
        wb2.close();

        System.out.println("Mismatch report generated successfully!");
    }

    // Create header map
    private static Map<String, Integer> getHeaderMap(Sheet sheet) {
        Map<String, Integer> map = new HashMap<>();
        Row headerRow = sheet.getRow(0);

        for (Cell cell : headerRow) {
            map.put(cell.getStringCellValue().trim(), cell.getColumnIndex());
        }
        return map;
    }
    private static String getCellValue(Cell cell) {
        if (cell == null) return "";

        switch (cell.getCellType()) {

            case STRING:
                return cell.getStringCellValue().trim();

            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                } else {
                    // Correct way to read numeric without casting to int
                    return new java.math.BigDecimal(cell.getNumericCellValue())
                            .stripTrailingZeros()
                            .toPlainString();
                }

            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());

            case FORMULA:
                return cell.getCellFormula();

            default:
                return "";
        }
    }}
    // Get cell value as String
//    private static String getCellValue(Cell cell) {
//        if (cell == null) return "";
//
//        switch (cell.getCellType()) {
//            case STRING:
//                return cell.getStringCellValue().trim();
//            case NUMERIC:
//                return String.valueOf((int) cell.getNumericCellValue());
//            case BOOLEAN:
//                return String.valueOf(cell.getBooleanCellValue());
//            default:
//                return "";
//        }
//    }
//}