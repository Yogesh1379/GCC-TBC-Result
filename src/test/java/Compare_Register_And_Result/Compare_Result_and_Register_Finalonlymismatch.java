package Compare_Register_And_Result;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Compare_Result_and_Register_Finalonlymismatch {

    public static void main(String[] args) throws EncryptedDocumentException, IOException {

        IOUtils.setByteArrayMaxOverride(1024 * 1024 * 1024);

        String outputPath =
                "F:\\GCCTBC-APR 2026\\Register data\\New folder\\50-60compare\\50Comparison_Output FINAL1.xlsx";

        File oldFile =
                new File("F:\\GCCTBC-APR 2026\\Register data\\New folder\\GCC-TBC-JANUARY-2026-ENGLISH_50WPM RESULT\\GCC-TBC-JANUARY-2026-ENGLISH_50WPM RESULT.xlsx"); //result

        File newFile =
                new File("F:\\GCCTBC-APR 2026\\Register data\\New folder\\ENGLISH50REGISTERDATA.xlsx"); // register

        Workbook oldWb = WorkbookFactory.create(new FileInputStream(oldFile));
        Workbook newWb = WorkbookFactory.create(new FileInputStream(newFile));

        Sheet oldSheet = oldWb.getSheetAt(0);   // RESULT sheet
        Sheet newSheet = newWb.getSheetAt(0);        // REGISTER sheet

        DataFormatter df = new DataFormatter();

        /*
          STEP 1 : Load OLD RESULT data
         */
        Map<String, String[]> oldData = new HashMap<>();

        for (int i = 1; i <= oldSheet.getLastRowNum(); i++) {

            Row r = oldSheet.getRow(i);

            if (r == null)
                continue;

            String seat = df.formatCellValue(r.getCell(3)); // SEAT NO

            if (seat.isEmpty())
                continue;

            oldData.put(seat, new String[] {

                    df.formatCellValue(r.getCell(12)), // Obj Grace
                    df.formatCellValue(r.getCell(19)), // Email+Stmt+Letter Grace
                    df.formatCellValue(r.getCell(17)), // Speed Grace
                    df.formatCellValue(r.getCell(21)), // Percentage
                    df.formatCellValue(r.getCell(22)), // Result
                    df.formatCellValue(r.getCell(23)), // Grade
                    df.formatCellValue(r.getCell(6)),  // Center Code
                    df.formatCellValue(r.getCell(7)),  // Center Address
                    df.formatCellValue(r.getCell(8)),  // Institute Code
                    df.formatCellValue(r.getCell(1))   // Student Name
            });
        }

        /*
          STEP 2 : Prepare OUTPUT sheet
         */
        Workbook outWb = new XSSFWorkbook();
        Sheet outSheet = outWb.createSheet("Comparison");

        Row header = outSheet.createRow(0);

        String[] headers = {
                "Seat No",
                "Result's Obj Grace", "Register's Obj Grace",
                "Result's Email Grace", "Register's Email Grace",
                "Result's Speed Grace", "Register's Speed Grace",
                "Result's Percentage", "Register's Percentage",
                "Result's Result", "Register's Result",
                "Result's Grade", "Register's Grade",
                "Result's Center Code", "Register's Center Code",
                "Result's Center Address", "Register's Center Address",
                "Result's Inst Code", "Register's Inst Code",
                "Result's Std Name", "Register's Std Name",
                "Status"
        };

        for (int i = 0; i < headers.length; i++) {
            header.createCell(i).setCellValue(headers[i]);
        }

        int rowIndex = 1;

        /*
          STEP 3 : Compare with REGISTER
         */
        for (int i = 1; i <= newSheet.getLastRowNum(); i++) {

            Row r = newSheet.getRow(i);

            if (r == null)
                continue;

            String seat = df.formatCellValue(r.getCell(6)); // Seat Number

            if (seat.isEmpty())
                continue;

            System.out.println("Checking Seat No : " + seat);

            String newObjGrace   = df.formatCellValue(r.getCell(10));
            String newEmailGrace = df.formatCellValue(r.getCell(11));
            String newSpeedGrace = df.formatCellValue(r.getCell(12));
            String newPercent    = df.formatCellValue(r.getCell(13));
            String newResult     = df.formatCellValue(r.getCell(14));
            String newGrade      = df.formatCellValue(r.getCell(15));
            String newCenterCode = df.formatCellValue(r.getCell(4));
            String newCenterAddr = df.formatCellValue(r.getCell(5));
            String newInstCode   = df.formatCellValue(r.getCell(7));
            String newStdName    = df.formatCellValue(r.getCell(8));

            String[] old = oldData.get(seat);

            /*
              SEAT NOT FOUND
             */
            if (old == null) {

                Row outRow = outSheet.createRow(rowIndex++);

                outRow.createCell(0).setCellValue(seat);
                outRow.createCell(21).setCellValue("SEAT NOT FOUND IN RESULT");

                continue;
            }

            /*
              Compare data
             */
            boolean match =

                    old[0].equals(newObjGrace) &&
                    old[1].equals(newEmailGrace) &&
                    old[2].equals(newSpeedGrace) &&
                    old[3].equals(newPercent) &&
                    old[4].equals(newResult) &&
                    old[5].equals(newGrade) &&
                    old[6].equals(newCenterCode) &&
                    // old[7].equals(newCenterAddr) &&
                    old[8].equals(newInstCode) &&
                    old[9].equals(newStdName);

            /*
              WRITE ONLY MISMATCH RECORDS
             */
            if (!match) {

                Row outRow = outSheet.createRow(rowIndex++);

                outRow.createCell(0).setCellValue(seat);

                int c = 1;

                outRow.createCell(c++).setCellValue(old[0]);
                outRow.createCell(c++).setCellValue(newObjGrace);

                outRow.createCell(c++).setCellValue(old[1]);
                outRow.createCell(c++).setCellValue(newEmailGrace);

                outRow.createCell(c++).setCellValue(old[2]);
                outRow.createCell(c++).setCellValue(newSpeedGrace);

                outRow.createCell(c++).setCellValue(old[3]);
                outRow.createCell(c++).setCellValue(newPercent);

                outRow.createCell(c++).setCellValue(old[4]);
                outRow.createCell(c++).setCellValue(newResult);

                outRow.createCell(c++).setCellValue(old[5]);
                outRow.createCell(c++).setCellValue(newGrade);

                outRow.createCell(c++).setCellValue(old[6]);
                outRow.createCell(c++).setCellValue(newCenterCode);

                outRow.createCell(c++).setCellValue(old[7]);
                outRow.createCell(c++).setCellValue(newCenterAddr);

                outRow.createCell(c++).setCellValue(old[8]);
                outRow.createCell(c++).setCellValue(newInstCode);

                outRow.createCell(c++).setCellValue(old[9]);
                outRow.createCell(c++).setCellValue(newStdName);

                outRow.createCell(c).setCellValue("MISMATCH");
            }
        }

        /*
          STEP 4 : Auto-size columns
         */
        for (int i = 0; i < headers.length; i++) {
            outSheet.autoSizeColumn(i);
        }

        /*
          STEP 5 : Save output
         */
        FileOutputStream fos = new FileOutputStream(outputPath);

        outWb.write(fos);

        fos.close();

        outWb.close();
        oldWb.close();
        newWb.close();

        System.out.println("Comparison completed successfully.");
        System.out.println("Output file : " + outputPath);
    }
}