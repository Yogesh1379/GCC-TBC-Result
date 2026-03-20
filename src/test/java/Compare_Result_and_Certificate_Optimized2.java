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

public class Compare_Result_and_Certificate_Optimized2 {

    public static void main(String[] args) throws EncryptedDocumentException, IOException {
        IOUtils.setByteArrayMaxOverride(1024 * 1024 * 1024);
String outputPath = "F:\\GCC-TBC October  repeaters\\Register Data\\Compare_Result and register.xlsx";
        File file1 = new File("F:\\GCC-TBC October  repeaters\\Register Data\\30RESULT.xlsx");// old--register OR result
        File file2 = new File("F:\\GCC-TBC October  repeaters\\Register Data\\OctRegisterData.xlsx");// new--result

        FileInputStream fis1 = new FileInputStream(file1);
        FileInputStream fis2 = new FileInputStream(file2);

        Workbook wb1 = WorkbookFactory.create(fis1);
        Workbook wb2 = WorkbookFactory.create(fis2);

        Sheet sheet1 = wb1.getSheet("Sheet2");
        Sheet sheet2 = wb2.getSheet("English (2)");

        DataFormatter df = new DataFormatter();

        // 🔹 Step 1: Load OLD data into HashMap (skip grace marks from comparison)
        Map<String, String[]> oldData = new HashMap<>();
        for (int i = 1; i <= sheet1.getLastRowNum(); i++) {
            Row r = sheet1.getRow(i);
            if (r == null) continue;
            String seat = df.formatCellValue(r.getCell(3));
            if (seat.isEmpty()) continue;
            oldData.put(seat, new String[]{
//                df.formatCellValue(r.getCell(32)), // grace Marks (IGNORED in comparison)
                df.formatCellValue(r.getCell(21)),  // percent
                df.formatCellValue(r.getCell(22)),  // result
                df.formatCellValue(r.getCell(6)),   // center code
                df.formatCellValue(r.getCell(7)),   // center address
                df.formatCellValue(r.getCell(8)),   // inst code
                df.formatCellValue(r.getCell(12)),  // obj grace
                df.formatCellValue(r.getCell(17)),  // speed grace
                df.formatCellValue(r.getCell(19)),  // email stmt ltr garce
                df.formatCellValue(r.getCell(23))   // grade
            });
        }

        // 🔹 Step 2: Prepare output workbook
        Workbook outWb = new XSSFWorkbook();
        Sheet outSheet = outWb.createSheet("Comparison");

        // Header Row
        Row header = outSheet.createRow(0);
        header.createCell(0).setCellValue("Seat Number");
        header.createCell(1).setCellValue("Old Grace Marks");
        header.createCell(2).setCellValue("New Grace Marks");
        header.createCell(3).setCellValue("Old Percent");
        header.createCell(4).setCellValue("New Percent");
        header.createCell(5).setCellValue("Old Result");
        header.createCell(6).setCellValue("New Result");
        header.createCell(7).setCellValue("Old Center Code");
        header.createCell(8).setCellValue("New Center Code");
        header.createCell(9).setCellValue("Old Center Address");
        header.createCell(10).setCellValue("New Center Address");
        header.createCell(11).setCellValue("Old Inst Code");
        header.createCell(12).setCellValue("New Inst Code");
        header.createCell(13).setCellValue("Old Obj Grace");
        header.createCell(14).setCellValue("New Obj Grace");
        header.createCell(15).setCellValue("Old Speed Grace");
        header.createCell(16).setCellValue("New Speed Grace");
        header.createCell(17).setCellValue("Old Email Stmt Ltr Grace");
        header.createCell(18).setCellValue("New Email Stmt Ltr Grace");
        header.createCell(19).setCellValue("Old Grade");
        header.createCell(20).setCellValue("New Grade");
        header.createCell(21).setCellValue("Status");

        int rowIndex = 1;

        // 🔹 Step 3: Compare NEW data with OLD data
        for (int i = 1; i <= sheet2.getLastRowNum(); i++) {
            Row r = sheet2.getRow(i);
            if (r == null) continue;

            String seat = df.formatCellValue(r.getCell(6));
            System.out.println(seat);
            if (seat.isEmpty()) continue;

            // Read new values
            String newGraceMarks   = df.formatCellValue(r.getCell(32)); // will display but not compare
            String newPercent      = df.formatCellValue(r.getCell(13));
            String newResult       = df.formatCellValue(r.getCell(14));
            String newCenterCode   = df.formatCellValue(r.getCell(4));
            String newCenterAddr   = df.formatCellValue(r.getCell(5));
            String newInstCode     = df.formatCellValue(r.getCell(7));
            String newObjGrace     = df.formatCellValue(r.getCell(10));
            String newSpeedGrace   = df.formatCellValue(r.getCell(12));
            String newEmailGrace   = df.formatCellValue(r.getCell(11));
            String newGrade        = df.formatCellValue(r.getCell(15));

            String[] oldValues = oldData.get(seat);
            if (oldValues == null) {
                Row outRow = outSheet.createRow(rowIndex++);
                outRow.createCell(0).setCellValue(seat);
                outRow.createCell(1).setCellValue("N/A"); // old grace marks not found
                outRow.createCell(2).setCellValue(newGraceMarks);
                outRow.createCell(3).setCellValue("N/A");
                outRow.createCell(4).setCellValue(newPercent);
                outRow.createCell(5).setCellValue("N/A");
                outRow.createCell(6).setCellValue(newResult);
                outRow.createCell(7).setCellValue("N/A");
                outRow.createCell(8).setCellValue(newCenterCode);
                outRow.createCell(9).setCellValue("N/A");
                outRow.createCell(10).setCellValue(newCenterAddr);
                outRow.createCell(11).setCellValue("N/A");
                outRow.createCell(12).setCellValue(newInstCode);
                outRow.createCell(13).setCellValue("N/A");
                outRow.createCell(14).setCellValue(newObjGrace);
                outRow.createCell(15).setCellValue("N/A");
                outRow.createCell(16).setCellValue(newSpeedGrace);
                outRow.createCell(17).setCellValue("N/A");
                outRow.createCell(18).setCellValue(newEmailGrace);
                outRow.createCell(19).setCellValue("N/A");
                outRow.createCell(20).setCellValue(newGrade);
                outRow.createCell(21).setCellValue("Seat not in OLD");
                continue;
            }

            // mapping old values (shifted since grace is skipped)
            String oldPercent    = oldValues[0];
            String oldResult     = oldValues[1];
            String oldCenterCode = oldValues[2];
            String oldCenterAddr = oldValues[3];
            String oldInstCode   = oldValues[4];
            String oldObjGrace   = oldValues[5];
            String oldSpeedGrace = oldValues[6];
            String oldEmailGrace = oldValues[7];
            String oldGrade      = oldValues[8];

            // ✅ Compare everything except grace marks
            if (!( oldPercent.equals(newPercent) &&
                   oldResult.equals(newResult) &&
                   oldCenterCode.equals(newCenterCode) &&
                   oldCenterAddr.equals(newCenterAddr) &&
                   oldInstCode.equals(newInstCode) &&
                   oldObjGrace.equals(newObjGrace) &&
                   oldSpeedGrace.equals(newSpeedGrace) &&
                   oldEmailGrace.equals(newEmailGrace) &&
                   oldGrade.equals(newGrade))) {

                Row outRow = outSheet.createRow(rowIndex++);
                outRow.createCell(0).setCellValue(seat);
                outRow.createCell(1).setCellValue("IGNORED"); // grace not compared
                outRow.createCell(2).setCellValue(newGraceMarks);
                outRow.createCell(3).setCellValue(oldPercent);
                outRow.createCell(4).setCellValue(newPercent);
                outRow.createCell(5).setCellValue(oldResult);
                outRow.createCell(6).setCellValue(newResult);
                outRow.createCell(7).setCellValue(oldCenterCode);
                outRow.createCell(8).setCellValue(newCenterCode);
                outRow.createCell(9).setCellValue(oldCenterAddr);
                outRow.createCell(10).setCellValue(newCenterAddr);
                outRow.createCell(11).setCellValue(oldInstCode);
                outRow.createCell(12).setCellValue(newInstCode);
                outRow.createCell(13).setCellValue(oldObjGrace);
                outRow.createCell(14).setCellValue(newObjGrace);
                outRow.createCell(15).setCellValue(oldSpeedGrace);
                outRow.createCell(16).setCellValue(newSpeedGrace);
                outRow.createCell(17).setCellValue(oldEmailGrace);
                outRow.createCell(18).setCellValue(newEmailGrace);
                outRow.createCell(19).setCellValue(oldGrade);
                outRow.createCell(20).setCellValue(newGrade);
                outRow.createCell(21).setCellValue("DOES NOT MATCH (Grace Ignored)");
            }
        }

        // 🔹 Step 4: Save output
        FileOutputStream fos = new FileOutputStream(outputPath);
        outWb.write(fos);
        fos.close();
        outWb.close();
        wb1.close();
        wb2.close();
        fis1.close();
        fis2.close();

        System.out.println("✅ Comparison completed. Grace marks ignored. Output saved.");
    }
}
