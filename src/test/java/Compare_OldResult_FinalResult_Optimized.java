import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;
import org.apache.poi.util.IOUtils;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Compare_OldResult_FinalResult_Optimized {

    public static void main(String[] args) throws EncryptedDocumentException, IOException {
        IOUtils.setByteArrayMaxOverride(200000000);

        File file1 = new File("C:\\Users\\User\\Desktop\\New folder (2)\\Book1.xlsx");// old

        File file2 = new File("C:\\Users\\User\\Desktop\\New folder (2)\\Result.xlsx");// new

        FileInputStream fis1 = new FileInputStream(file1);
        FileInputStream fis2 = new FileInputStream(file2);

        Workbook wb1 = WorkbookFactory.create(fis1);
        Workbook wb2 = WorkbookFactory.create(fis2);

        Sheet sheet1 = wb1.getSheet("Sheet1");
        Sheet sheet2 = wb2.getSheet("Sheet2");

        DataFormatter df = new DataFormatter();

        // 🔹 Step 1: Load OLD data into HashMap
        Map<String, String[]> oldData = new HashMap<>();
        for (int i = 1; i <= sheet1.getLastRowNum(); i++) {
            Row r = sheet1.getRow(i);
            if (r == null) continue;
            String seat = df.formatCellValue(r.getCell(3));
            if (seat.isEmpty()) continue;
            oldData.put(seat, new String[]{
                    df.formatCellValue(r.getCell(21)), // Marks
                    df.formatCellValue(r.getCell(22)), // Result
                    df.formatCellValue(r.getCell(23))  // Grade
            });
        }

        // 🔹 Step 2: Prepare output workbook
        Workbook outWb = new XSSFWorkbook();
        Sheet outSheet = outWb.createSheet("Comparison");

        // Header Row
        Row header = outSheet.createRow(0);
        header.createCell(0).setCellValue("Seat Number");
        header.createCell(1).setCellValue("Old Marks");
        header.createCell(2).setCellValue("New Marks");
        header.createCell(3).setCellValue("Old Result");
        header.createCell(4).setCellValue("New Result");
        header.createCell(5).setCellValue("Old Grade");
        header.createCell(6).setCellValue("New Grade");
        header.createCell(7).setCellValue("Status");

        int rowIndex = 1;

        // 🔹 Step 3: Compare NEW data with OLD data
        for (int i = 1; i <= sheet2.getLastRowNum(); i++) {
            Row r = sheet2.getRow(i);
            if (r == null) continue;

            String seat = df.formatCellValue(r.getCell(3));
            System.out.println(seat);
            if (seat.isEmpty()) continue;

            String newMarks = df.formatCellValue(r.getCell(21));
            String newResult = df.formatCellValue(r.getCell(22));
            String newGrade = df.formatCellValue(r.getCell(23));

            String[] oldValues = oldData.get(seat);

            if (oldValues == null) {
                // Seat not found in old file → write difference
                Row outRow = outSheet.createRow(rowIndex++);
                outRow.createCell(0).setCellValue(seat);
                outRow.createCell(1).setCellValue("N/A");
                outRow.createCell(2).setCellValue(newMarks);
                outRow.createCell(3).setCellValue("N/A");
                outRow.createCell(4).setCellValue(newResult);
                outRow.createCell(5).setCellValue("N/A");
                outRow.createCell(6).setCellValue(newGrade);
                outRow.createCell(7).setCellValue("Seat not in OLD");
                continue;
            }

            String oldMarks = oldValues[0];
            String oldResult = oldValues[1];
            String oldGrade = oldValues[2];

            // ✅ Only write if values differ
            if (!(oldMarks.equals(newMarks) &&
                  oldResult.equals(newResult) &&
                  oldGrade.equals(newGrade))) {

                Row outRow = outSheet.createRow(rowIndex++);
                outRow.createCell(0).setCellValue(seat);
                outRow.createCell(1).setCellValue(oldMarks);
                outRow.createCell(2).setCellValue(newMarks);
                outRow.createCell(3).setCellValue(oldResult);
                outRow.createCell(4).setCellValue(newResult);
                outRow.createCell(5).setCellValue(oldGrade);
                outRow.createCell(6).setCellValue(newGrade);
                outRow.createCell(7).setCellValue("DOES NOT MATCH");
            }
        }

        // 🔹 Step 4: Save output
        FileOutputStream fos = new FileOutputStream("C:\\Users\\User\\Desktop\\Compare ResultSSD.xlsx");
        outWb.write(fos);
        fos.close();
        outWb.close();
        wb1.close();
        wb2.close();
        fis1.close();
        fis2.close();

        System.out.println("✅ Comparison completed. Only mismatches saved: Compare Result.xlsx");
    }
}
