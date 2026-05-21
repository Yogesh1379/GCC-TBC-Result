package GRACE_MARKS;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.Reporter;
import org.testng.annotations.Test;

public class newResultWithGraceFinal implements IAutoConstant {
    @Test
    public void resultTest() throws EncryptedDocumentException, IOException {
        File file = new File("F:\\GCCTBC-APR 2026\\Result\\Final\\FinalResult - Copy.xlsx");
        FileInputStream fis = new FileInputStream(file);
        IOUtils.setByteArrayMaxOverride(200000000);
        Workbook wb = WorkbookFactory.create(fis);
        Sheet sheet = wb.getSheetAt(0);
        int rc = sheet.getLastRowNum();
        DataFormatter df = new DataFormatter();
        @SuppressWarnings("resource")
		Workbook workbook1 = new XSSFWorkbook();
        IOUtils.setByteArrayMaxOverride(200000000);
        Sheet sheet1 = workbook1.createSheet("TestSheet");
        Row header = sheet1.createRow(0);
        header.createCell(0).setCellValue("seatNumber");
        header.createCell(1).setCellValue("percent");
        header.createCell(2).setCellValue("Expected percent");
        header.createCell(3).setCellValue("result");
        header.createCell(4).setCellValue("Expected result");
        header.createCell(5).setCellValue("grade");
        header.createCell(6).setCellValue("Expected grade");
        header.createCell(7).setCellValue("Grace");
        header.createCell(8).setCellValue("Expected Grace");
        header.createCell(9).setCellValue("PerAfter Grace");
        header.createCell(10).setCellValue("Status");

        for (int i = 1; i <= rc; i++) {

            Row row = sheet.getRow(i);
            String seatNumber = df.formatCellValue(row.getCell(3));// old 3
            String obj = df.formatCellValue(row.getCell(11));//11
            Double objmarks = Double.valueOf(obj);
            String speed = df.formatCellValue(row.getCell(16));//16
            Double speedMarks = Double.valueOf(speed);
            String eel = df.formatCellValue(row.getCell(18));//18
            Double EmailExcelLetter = Double.valueOf(eel);
            String per = df.formatCellValue(row.getCell(21));//21
            Double percent = Double.valueOf(per);
            String result = df.formatCellValue(row.getCell(22));//22
            String grade = df.formatCellValue(row.getCell(23));//23
            String presentAbsent = df.formatCellValue(row.getCell(30));//30
            String graceex = df.formatCellValue(row.getCell(32));//32
            Double graceExcel = Double.valueOf(graceex);
            String email = df.formatCellValue(row.getCell(13));//13
            Double email1 = Double.valueOf(email);
            String stmt = df.formatCellValue(row.getCell(15));//15
            Double stmt1 = Double.valueOf(stmt);
            String ltr = df.formatCellValue(row.getCell(14));//14
            Double ltr1 = Double.valueOf(ltr);
            System.out.println(i + ":- " + seatNumber);
            String Result = "";
            double total = objmarks + speedMarks + EmailExcelLetter;
//			 assertEquals(percent, total);
            Row row1 = sheet1.createRow(i);
            row1.createCell(0).setCellValue(seatNumber);
            row1.createCell(1).setCellValue(percent);
            row1.createCell(3).setCellValue(result);
            row1.createCell(5).setCellValue(grade);
            row1.createCell(2).setCellValue(total);
            row1.createCell(7).setCellValue(graceExcel);
            row1.createCell(8).setCellValue(0);
            row1.createCell(9).setCellValue(total);
            String Grade="";
            //grace mark calculation
            double objGrace1 = 0;
            double EmailExcelLetterGrace1=0;
            double speedGrace1=0;
            double totalGrace = 0;
            double percentAfterGarce=total;

//            String graceex = df.formatCellValue(row.getCell(32));
//            Double graceExcel = Double.valueOf(graceex);
            double EmailExcelLetter1 = email1+stmt1+ltr1;
            double objGrace = 10-objmarks;
            double EmailExcelLetterGrace = 14-EmailExcelLetter1;
            double speedGrace = 16-speedMarks;
            if(presentAbsent.equals("Present")) {
                if(objGrace >= 0 && objGrace <= 2)
                {
                    objGrace1 = objGrace;
                }
                if(EmailExcelLetterGrace >= 0 && EmailExcelLetterGrace <= 2 ) {
                    EmailExcelLetterGrace1=EmailExcelLetterGrace;
                }
                if(speedGrace >= 0 && speedGrace <= 2)
                {
                    speedGrace1=speedGrace;
                }
                totalGrace = objGrace1+speedGrace1+EmailExcelLetterGrace1;
                double speedExcelgrace = speedGrace1+EmailExcelLetterGrace1;
                if(totalGrace==4 || totalGrace<=4 && totalGrace!=0 ) {
                    if (EmailExcelLetter1 >= 12 && objmarks >= 8 && speedMarks >= 14 && speedExcelgrace <= 2) {
                        if (EmailExcelLetterGrace1 != 0 || EmailExcelLetterGrace1 <= 1 ||
                                speedGrace1 != 0 || speedGrace1 <= 1 || objGrace1 >= 0 || objGrace1 <= 2) {
                            Reporter.log(seatNumber + " " + totalGrace, true);
//                            Row row1 = sheet1.createRow(i);
//                            row1.createCell(0).setCellValue(seatNumber);
//                            row1.createCell(1).setCellValue(graceExcel);
                            row1.createCell(8).setCellValue(totalGrace);
                            percentAfterGarce = (double) total + totalGrace;
                            row1.createCell(9).setCellValue(percentAfterGarce);
                            Reporter.log(seatNumber + " tatal grace :- " + totalGrace);
                        } else {
                            totalGrace = 0;
                        }
                    } else {
                        totalGrace = 0;
                        Reporter.log(seatNumber + " student is fail");
//                        row1.createCell(8).setCellValue(0);
                    }
                }
            }

            else
            {
                totalGrace = 0;
                Reporter.log(seatNumber+"  Student is Absent");

            }

//            testing
            if (presentAbsent.equals("Present") && result.equalsIgnoreCase("ABSENT")) {

                Reporter.log(seatNumber + "  Support Remark");
                row1.createCell(4).setCellValue("Support Remark");
            }
            else if (presentAbsent.equals("Present")) {
                if (objmarks >= 10 && EmailExcelLetter >= 14 && speedMarks >= 16) {

                    Result = "PASS";
                    row1.createCell(4).setCellValue(Result);
                    Reporter.log(seatNumber + " is " + Result, true);
//					assertEquals(Result, result);

                } else if (percentAfterGarce>=40&&totalGrace!=0&&(objmarks >= 8 && EmailExcelLetter >= 12
                        && speedMarks >= 14)) {
                    Result = "PASS";
                    row1.createCell(4).setCellValue(Result);
                    Reporter.log(seatNumber + " is " + Result, true);
//					assertEquals(Result, result);

                } else if (presentAbsent.equalsIgnoreCase("Present") && result.equalsIgnoreCase("RESERVED")) {
                    String reslt = "RESERVED";
                    row1.createCell(4).setCellValue(reslt);
                    Reporter.log(seatNumber + " is " + reslt, true);
//					assertEquals(reslt, result);
                } else if (presentAbsent.equalsIgnoreCase("Present") && result.equalsIgnoreCase("CANCELLED")) {
                    String reslt = "CANCELLED";
                    row1.createCell(4).setCellValue(reslt);
                    Reporter.log(seatNumber + " is " + reslt, true);
//					assertEquals(reslt, result);
                } else {
                     Result = "FAIL";
                    Grade="-";
                    row1.createCell(4).setCellValue(Result);
                    Reporter.log(seatNumber + " is " + Result, true);
//					assertEquals(result1, result);
                }

                if (percentAfterGarce >= 75) {
                    if (Result.equals("PASS")
//							&&grade.equals("A+")
                    ) {
                        Grade="A+";
                        Reporter.log(seatNumber + " is PASS with A+ grade", true);
                        row1.createCell(6).setCellValue("A+");
                    } else {
                        Grade="-";
                        Reporter.log(seatNumber + " is FAIL with - grade", true);
                        row1.createCell(6).setCellValue("-");
                        //					assertEquals( "A+",grade);
                    }
                } else if (percentAfterGarce >= 60 && percentAfterGarce <= 74.99) {
                    if (Result.equals("PASS")
//							&&grade.equals("A")
                    ) {
                        Grade="A";
                        Reporter.log(seatNumber + " is PASS with A grade", true);
                        row1.createCell(6).setCellValue("A");
                    } else {
                        Grade="-";
                        Reporter.log(seatNumber + " is FAIL with - grade", true);
                        row1.createCell(6).setCellValue("-");
                        //					assertEquals( "A",grade);
                    }
                } else if (percentAfterGarce >= 50 && percentAfterGarce <= 59.99) {
                    if (Result.equals("PASS")
//							&&grade.equals("B")
                    ) {
                        Grade="B";
                        Reporter.log(seatNumber + " is PASS with B grade", true);
                        row1.createCell(6).setCellValue("B");
                    } else {
                        Grade="-";
                        Reporter.log(seatNumber + " is FAIL with - grade", true);
                        row1.createCell(6).setCellValue("-");
                        //					assertEquals( "B",grade);
                    }
                } else if (percentAfterGarce >= 40 && percentAfterGarce <= 49.99) {
                    if (Result.equals("PASS")
//							&&grade.equals("C")
                    ) {
                        Grade="C";
                        Reporter.log(seatNumber + " is PASS with C grade", true);
                        row1.createCell(6).setCellValue("C");
                    } else {
                        Grade="-";
                        Reporter.log(seatNumber + " is FAIL with - grade", true);
                        row1.createCell(6).setCellValue("-");
                        //					assertEquals( "C",grade);
                    }
                } else {
                    Reporter.log(seatNumber + " Grade is - ", true);
                    row1.createCell(6).setCellValue("-");
                }

            }
            else if (presentAbsent.equals("Absent")) {
                Result="ABSENT";
                Reporter.log(seatNumber + "  Student is Absent");
                row1.createCell(4).setCellValue("ABSENT");
                row1.createCell(6).setCellValue("-");
                Grade="-";
            }




            if( result.equalsIgnoreCase(Result) && (percentAfterGarce==percent)
            &&(grade.equalsIgnoreCase(Grade)) && graceExcel==totalGrace ){
                row1.createCell(10).setCellValue("Matched");
            }

            else {
                row1.createCell(10).setCellValue("Mis-Matched");
            }

        }
        FileOutputStream fileOut = new FileOutputStream(new File("F:\\GCCTBC-APR 2026\\Result\\Final\\30RESULT WITH GRACE.xlsx"));
        // Write the workbook data to the file
        workbook1.write(fileOut);
    }
}


