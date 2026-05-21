package Email_Marking;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

public class Markcompare {

    public static void main(String[] args) throws IOException {

        String file1Path = "F:\\GCCTBC-APR 2026\\Email marking\\840 stud 501 batch marathi\\Email marathi30 Marks.xlsx";
        String file2Path = "F:\\GCCTBC-APR 2026\\Email marking\\840 stud 501 batch marathi\\New Microsoft Excel Worksheet.xlsx";//vj
        String outputPath = "F:\\GCCTBC-APR 2026\\Email marking\\840 stud 501 batch marathi\\comparisonOutput_marathi30.xlsx";

        DataFormatter formatter = new DataFormatter();

        // =======================
        // Read First Excel into Map<String, List<Double>>
        // =======================
        Map<String, List<Double>> firstExcelData = new HashMap<>();

        try (Workbook wb1 = new XSSFWorkbook(new FileInputStream(file1Path))) {

            Sheet sheet1 = wb1.getSheetAt(0);

            for (int i = 1; i <= sheet1.getLastRowNum(); i++) {
                Row row = sheet1.getRow(i);
                if (row == null) continue;

                Cell seatCell = row.getCell(0);  // Column A
                Cell marksCell = row.getCell(6); // Column G

                if (seatCell == null || marksCell == null) continue;

                String seatNumber = formatter.formatCellValue(seatCell).trim();
                double marks = marksCell.getNumericCellValue();

                firstExcelData.computeIfAbsent(seatNumber, k -> new ArrayList<>()).add(marks);
            }
        }

        // =======================
        // Read Second Excel into Map<String, List<Double>>
        // =======================
        Map<String, List<Double>> secondExcelData = new HashMap<>();

        try (Workbook wb2 = new XSSFWorkbook(new FileInputStream(file2Path))) {
            Sheet sheet2 = wb2.getSheetAt(0);

            for (int i = 1; i <= sheet2.getLastRowNum(); i++) {
                Row row = sheet2.getRow(i);
                if (row == null) continue;

                Cell seatCell = row.getCell(0);  // Column A
                Cell marksCell = row.getCell(1); // Column B

                if (seatCell == null || marksCell == null) continue;

                String seatNumber = formatter.formatCellValue(seatCell).trim();
                double marks = marksCell.getNumericCellValue();

                secondExcelData.computeIfAbsent(seatNumber, k -> new ArrayList<>()).add(marks);
            }
        }

        // =======================
        // Prepare Output Workbook with two sheets
        // =======================
        try (Workbook outputWorkbook = new XSSFWorkbook()) {

            Sheet multipleSheet = outputWorkbook.createSheet("Multiple SeatNumbers");
            Sheet uniqueSheet = outputWorkbook.createSheet("Unique SeatNumbers");

            createHeaderRow(multipleSheet);
            createHeaderRow(uniqueSheet);

            int multipleRowNum = 1;
            int uniqueRowNum = 1;

            // =======================
            // Compare Seat Numbers
            // =======================
            Set<String> allSeats = new HashSet<>();
            allSeats.addAll(firstExcelData.keySet());
            allSeats.addAll(secondExcelData.keySet());

            for (String seat : allSeats) {

                List<Double> marksList1 = firstExcelData.getOrDefault(seat, new ArrayList<>());
                List<Double> marksList2 = secondExcelData.getOrDefault(seat, new ArrayList<>());

                // Sort descending for top marks first
                marksList1.sort(Collections.reverseOrder());
                marksList2.sort(Collections.reverseOrder());

                int size = Math.max(marksList1.size(), marksList2.size());

                Sheet targetSheet = marksList1.size() > 1 ? multipleSheet : uniqueSheet;
                int rowNum = marksList1.size() > 1 ? multipleRowNum : uniqueRowNum;

                for (int i = 0; i < size; i++) {

                    Double marks1 = i < marksList1.size() ? marksList1.get(i) : null;
                    Double marks2 = i < marksList2.size() ? marksList2.get(i) : null;

                    Row outRow = targetSheet.createRow(rowNum++);
                    outRow.createCell(0).setCellValue(seat);

                    if (marks1 != null) outRow.createCell(1).setCellValue(marks1);
                    else outRow.createCell(1).setCellValue("Not Found");

                    if (marks2 != null) outRow.createCell(2).setCellValue(marks2);
                    else outRow.createCell(2).setCellValue("Not Found");

                    if (marks1 != null && marks2 != null) {
                        if (Double.compare(marks1, marks2) == 0) {
                            outRow.createCell(3).setCellValue("MATCH");
                        } else {
                            outRow.createCell(3).setCellValue("NOT MATCH");
                        }
                        outRow.createCell(4).setCellValue(Math.abs(marks1 - marks2));
                    } else {
                        outRow.createCell(3).setCellValue("NOT MATCH");
                        outRow.createCell(4).setCellValue("-");
                    }
                }

                if (marksList1.size() > 1) multipleRowNum = rowNum;
                else uniqueRowNum = rowNum;
            }

            // Auto size columns
            for (Sheet sheet : Arrays.asList(multipleSheet, uniqueSheet)) {
                for (int i = 0; i <= 4; i++) {
                    sheet.autoSizeColumn(i);
                }
            }

            // Write Output Excel
            try (FileOutputStream fos = new FileOutputStream(outputPath)) {
                outputWorkbook.write(fos);
            }
        }

        System.out.println("Comparison Completed Successfully. Output Excel generated at: " + outputPath);
    }

    private static void createHeaderRow(Sheet sheet) {
        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue("Seat Number");
        header.createCell(1).setCellValue("Marks (File1)");
        header.createCell(2).setCellValue("Marks (File2)");
        header.createCell(3).setCellValue("Result");
        header.createCell(4).setCellValue("Difference");
    }
}
