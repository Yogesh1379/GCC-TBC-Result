package Question_Data_name_verify;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashSet;
import java.util.Iterator;
import java.util.Set;

public class Question_file_name {

    public static void main(String[] args) {
        String excelPath = "F:\\GCC  TBC December 2025\\Question\\allocation\\marathi\\SUbjectiveMarathiHin.xlsx"; // Input Excel
        String folderPath = "F:\\GCC  TBC December 2025\\Question\\allocation\\marathi\\All"; // Local folder path
        String outputPath = "F:\\GCC  TBC December 2025\\Question\\allocation\\marathi\\Wrong Files que Name Mar1.xlsx"; // Output Excel change name accordingly

        try {
            // Step 1: Read Excel
            FileInputStream fis = new FileInputStream(excelPath);
            Workbook workbook = new XSSFWorkbook(fis);
            Sheet sheet = workbook.getSheetAt(0);		//change index of sheet to 1 for Marathi verification

            // Step 2: Read all filenames from 2 column (index 1)
            Set<String> excelFileNames = new HashSet<>();
            DataFormatter formatter = new DataFormatter();

            Iterator<Row> rowIterator = sheet.iterator();
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                Cell cell = row.getCell(2); // 3rd column, index = 2
                if (cell != null) {
                    String fileName = formatter.formatCellValue(cell).trim();
                    if (!fileName.isEmpty()) {
                        excelFileNames.add(fileName);
                    }
                }
            }
            workbook.close();

            // Step 3: Read local folder file names
            File folder = new File(folderPath);
            File[] files = folder.listFiles();
            Set<String> localFileNames = new HashSet<>();
            if(files==null) {
            	System.out.println("Folder path is wrong");
            }

            if (files != null) {
                for (File f : files) {
                    if (f.isFile()) {
                        localFileNames.add(f.getName());
                    }
                }
            }

            // Step 4: Create new workbook for results
            Workbook outputWorkbook = new XSSFWorkbook();

            // Missing files sheet
            Sheet missingSheet = outputWorkbook.createSheet("Missing Files");
            int missingIndex = 0;
            Row headerRow = missingSheet.createRow(missingIndex++);
            headerRow.createCell(0).setCellValue("Missing File Names");

            // Match files sheet
            Sheet matchSheet = outputWorkbook.createSheet("Match Files");
            int matchIndex = 0;
            Row headerRow1 = matchSheet.createRow(matchIndex++);
            headerRow1.createCell(0).setCellValue("Match File Names");
            
            //Folder File Names which not match with Name format
            Sheet folderfile = outputWorkbook.createSheet("Folder Files");
            int folderIndex = 0;
            Row headerRow2 = folderfile.createRow(folderIndex++);
            headerRow2.createCell(0).setCellValue("Folder File Names which not match with Name format");
            

            // Step 5: Compare and record results
            for (String excelFile : excelFileNames) {
                if (!localFileNames.contains(excelFile)) {
                    Row row = missingSheet.createRow(missingIndex++);
                    row.createCell(0).setCellValue(excelFile);
                } else {
                    Row row = matchSheet.createRow(matchIndex++);
                    row.createCell(0).setCellValue(excelFile);
                }
            }
            
            // for file name from folder not match with excel file
            for(String lcname:localFileNames)
            {
            	if(!excelFileNames.contains(lcname))
            	{
            		Row row = folderfile.createRow(folderIndex);
            		row.createCell(0).setCellValue(lcname);
            	}
            }

            // Step 6: Save output Excel
            FileOutputStream fos = new FileOutputStream(outputPath);
            outputWorkbook.write(fos);
            fos.close();
            outputWorkbook.close();

            System.out.println("✅ Results exported to: " + outputPath);

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
