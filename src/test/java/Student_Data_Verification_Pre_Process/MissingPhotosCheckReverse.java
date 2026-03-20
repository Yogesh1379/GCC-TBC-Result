package Student_Data_Verification_Pre_Process;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.HashSet;
import java.util.Set;

public class MissingPhotosCheckReverse {

    public static void main(String[] args) throws Exception {

        String excelPath = "F:\\GCC  TBC December 2025\\student data\\StudentData.xlsx";
        String photoFolderPath = "F:\\GCC  TBC December 2025\\Photo\\CTIMAGES\\";

        // ---------------------------------------------------------
        //  STEP 1: Load Excel Photo Names into SET
        // ---------------------------------------------------------
        FileInputStream fis = new FileInputStream(excelPath);
        Workbook wb = new XSSFWorkbook(fis);
        Sheet sheet = wb.getSheetAt(0);

        DataFormatter df = new DataFormatter();

        Set<String> excelPhotoNames = new HashSet<>();

        for (int i = 1; i <= sheet.getLastRowNum(); i++) {

            Row row = sheet.getRow(i);
            if (row == null) continue;

            String photoName = df.formatCellValue(row.getCell(16)).trim(); // column 16
            if (!photoName.isEmpty()) {
                excelPhotoNames.add(photoName.toLowerCase());
            }
        }

        wb.close();
        fis.close();

        // ---------------------------------------------------------
        //  STEP 2: Read Photos From Folder
        // ---------------------------------------------------------
        File folder = new File(photoFolderPath);
        File[] files = folder.listFiles();

        // ---------------------------------------------------------
        //  STEP 3: Output Excel for Extra Photos (not in Excel)
        // ---------------------------------------------------------
        Workbook wbOut = new XSSFWorkbook();
        Sheet outSheet = wbOut.createSheet("PhotosNotInExcel");

        Row header = outSheet.createRow(0);
        header.createCell(0).setCellValue("Photo Name (Not in Excel)");

        int rowIndex = 1;

        if (files != null) {
            for (File f : files) {
                String fileName = f.getName().trim().toLowerCase();

                // -------------------------------------------------
                //  CHECK: Photo exists in Excel?
                // -------------------------------------------------
                if (!excelPhotoNames.contains(fileName)) {

                    Row r = outSheet.createRow(rowIndex++);
                    r.createCell(0).setCellValue(f.getName());

                    System.out.println("Not in Excel: " + f.getName());
                }
            }
        }

        // ---------------------------------------------------------
        //  STEP 4: Save Output File
        // ---------------------------------------------------------
        FileOutputStream fos = new FileOutputStream(
                "F:\\GCC  TBC December 2025\\Photo\\Extra_Photos_Not_In_Excel.xlsx"
        );
        wbOut.write(fos);

        fos.close();
        wbOut.close();

        System.out.println("Extra_Photos_Not_In_Excel.xlsx created successfully!");
    }
}
