package Student_Data_Verification_Pre_Process;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.HashSet;
import java.util.Set;

public class MissingPhotosCheck {

    public static void main(String[] args) throws Exception {

        String excelPath = "F:\\GCC  TBC December 2025\\student data\\StudentData.xlsx";
        String photoFolderPath = "F:\\GCC  TBC December 2025\\Photo\\CTIMAGES\\";

        File folder = new File(photoFolderPath);
        File[] files = folder.listFiles();

        // Store all photo names in a SET for fast lookup
        Set<String> photoSet = new HashSet<>();

        if (files != null) {
            for (File f : files) {
                photoSet.add(f.getName().trim().toLowerCase());
            }
        }


        FileInputStream fis = new FileInputStream(excelPath);
        Workbook wb = new XSSFWorkbook(fis);
        Sheet sheet = wb.getSheetAt(0);

        DataFormatter df = new DataFormatter();


        Workbook wbOut = new XSSFWorkbook();
        Sheet outSheet = wbOut.createSheet("MissingPhotos");

        Row header = outSheet.createRow(0);
        header.createCell(0).setCellValue("Missing Photo Name");
        header.createCell(1).setCellValue("Row Number in Original File");

        int outputRow = 1;


        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
System.out.println(i);
            Row row = sheet.getRow(i);
            if (row == null) continue;

            String photoName = df.formatCellValue(row.getCell(16)).trim(); // CHANGE column index
            System.out.println( photoName);
            if (photoName.isEmpty()) continue;

            // Convert to lowercase for safe match
            String photoLower = photoName.toLowerCase();

                        if (!photoSet.contains(photoLower)) {

                Row outRow = outSheet.createRow(outputRow++);
                outRow.createCell(0).setCellValue(photoName);
                outRow.createCell(1).setCellValue(i);

                System.out.println("Missing: " + photoName);
            }
        }


        FileOutputStream fos = new FileOutputStream(
                "F:\\GCC  TBC December 2025\\Photo\\Missing_Photos_Report.xlsx"
        );
        wbOut.write(fos);
        fos.close();

        wb.close();
        wbOut.close();

        System.out.println("Missing_Photos_Report.xlsx created successfully!");
    }
}
