package Present_Absent;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;

public class PresentAbsentFromFilename {

    static class RowData {
        String seatNo;
        String batch;
        String speedFile = "";
        String excelFile = "";
        String letterFile = "";

        RowData(String seatNo, String batch) {
            this.seatNo = seatNo;
            this.batch = batch;
        }

        boolean hasEmptySection(String section) {
            if ("speed".equals(section)) return speedFile.isEmpty();
            if ("excel".equals(section)) return excelFile.isEmpty();
            if ("letter".equals(section)) return letterFile.isEmpty();
            return false;
        }

        void setSection(String section, String fileName) {
            if ("speed".equals(section)) speedFile = fileName;
            else if ("excel".equals(section)) excelFile = fileName;
            else if ("letter".equals(section)) letterFile = fileName;
        }
    }

    public static void main(String[] args) {

        String inputPath = "C:\\Users\\User\\Desktop\\marathi file names.xlsx";
        String outputPath = "C:\\Users\\User\\Desktop\\present Absent1.xlsx";

        long start = System.currentTimeMillis();

        try (FileInputStream fis = new FileInputStream(inputPath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0);
            DataFormatter df = new DataFormatter();

            // Key = seatNo_batch
            Map<String, List<RowData>> dataMap = new LinkedHashMap<>();

            for (int i = 1; i <= sheet.getLastRowNum(); i++) {

                Row row = sheet.getRow(i);
                if (row == null) continue;

                String fileName = df.formatCellValue(
                        row.getCell(0, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK)
                ).trim();

                if (fileName.isEmpty()) continue;

                // Split filename
                String[] parts = fileName.split("_");
                if (parts.length < 3) continue;

                String rawSection = parts[0].toLowerCase();
                String seatNo = parts[1].trim();
                String batch = parts[2].trim();

                String section;

                if (rawSection.contains("speed"))
                    section = "speed";
                else if (rawSection.contains("excel"))
                    section = "excel";
                else if (rawSection.contains("letter"))
                    section = "letter";
                else
                    continue;

                String key = seatNo + "_" + batch;

                dataMap.putIfAbsent(key, new ArrayList<>());
                List<RowData> rowList = dataMap.get(key);

                boolean placed = false;

                for (RowData rd : rowList) {
                    if (rd.hasEmptySection(section)) {
                        rd.setSection(section, fileName);
                        placed = true;
                        break;
                    }
                }

                if (!placed) {
                    RowData newRow = new RowData(seatNo, batch);
                    newRow.setSection(section, fileName);
                    rowList.add(newRow);
                }
            }

            // Create Output File
            try (Workbook outWorkbook = new XSSFWorkbook();
                 FileOutputStream fos = new FileOutputStream(outputPath)) {

                Sheet outSheet = outWorkbook.createSheet("Output");

                // Header
                Row header = outSheet.createRow(0);
                header.createCell(0).setCellValue("Seat No");
                header.createCell(1).setCellValue("Batch");
                header.createCell(2).setCellValue("Speed File");
                header.createCell(3).setCellValue("Excel File");
                header.createCell(4).setCellValue("Letter File");
                header.createCell(5).setCellValue("Speed count");
                header.createCell(6).setCellValue("Stmt count");
                header.createCell(7).setCellValue("Letter count");

                int rowIndex = 1;

                for (List<RowData> rowList : dataMap.values()) {
                    for (RowData rd : rowList) {

                        Row outRow = outSheet.createRow(rowIndex++);
                        outRow.createCell(0).setCellValue(rd.seatNo);
                        outRow.createCell(1).setCellValue(rd.batch);
                        outRow.createCell(2).setCellValue(rd.speedFile);
                        if(!rd.speedFile.isBlank()) {
                            outRow.createCell(5).setCellValue(1);}
                            else { outRow.createCell(5).setCellValue(0);
                            	
                            }
                        outRow.createCell(3).setCellValue(rd.excelFile);
                        if(!rd.excelFile.isBlank()) {
                            outRow.createCell(6).setCellValue(1);}
                            else { outRow.createCell(6).setCellValue(0);
                            	
                            }
                        outRow.createCell(4).setCellValue(rd.letterFile);
                        if(!rd.letterFile.isBlank()) {
                        outRow.createCell(7).setCellValue(1);}
                        else { outRow.createCell(7).setCellValue(0);
                        	
                        }
                    }
                }

                // Auto size columns
                for (int i = 0; i < 5; i++) {
                    outSheet.autoSizeColumn(i);
                }

                outWorkbook.write(fos);
            }

            long end = System.currentTimeMillis();
            System.out.println("✅ Completed Successfully!");
            System.out.println("Time Taken: " + (end - start) + " ms");

        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}