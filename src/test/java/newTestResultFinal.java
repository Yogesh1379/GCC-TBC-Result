import static org.testng.Assert.assertEquals;

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

public class newTestResultFinal implements IAutoConstant{
	@Test
	public void resultTest() throws EncryptedDocumentException, IOException
	{
		File file = new File(Excel_Path);
		FileInputStream fis = new FileInputStream(file);
		IOUtils.setByteArrayMaxOverride(200000000);
		Workbook wb = WorkbookFactory.create(fis);
		Sheet sheet = wb.getSheet("30");
		int rc = sheet.getLastRowNum();
		DataFormatter df = new DataFormatter();
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
		for(int i=1;i<=rc;i++)
		{
			Row row = sheet.getRow(i);
			String seatNumber = df.formatCellValue(row.getCell(3));
			String obj = df.formatCellValue(row.getCell(12));
			Double objmarks = Double.valueOf(obj);
			String speed = df.formatCellValue(row.getCell(17));
			Double speedMarks = Double.valueOf(speed);
			String eel = df.formatCellValue(row.getCell(19));
			Double EmailExcelLetter = Double.valueOf(eel);
			String per = df.formatCellValue(row.getCell(21));
			Double percent = Double.valueOf(per);
			String result=df.formatCellValue(row.getCell(22));
			String grade = df.formatCellValue(row.getCell(23));
			String presentAbsent = df.formatCellValue(row.getCell(30));
			System.out.println(i+":- "+seatNumber);
			String Result="PASS";
double total = objmarks+speedMarks+EmailExcelLetter;
//			assertEquals(percent, total);
			 Row row1 = sheet1.createRow(i);
		        row1.createCell(0).setCellValue(seatNumber);
		        row1.createCell(1).setCellValue(percent);
		        row1.createCell(3).setCellValue(result);
		        row1.createCell(5).setCellValue(grade);
		        row1.createCell(2).setCellValue(total);
		         if(presentAbsent.equals("Present") && result.equalsIgnoreCase("ABSENT")) {
					
					Reporter.log(seatNumber+"  Support Remark");
					row1.createCell(4).setCellValue("Support Remark");
				}
		         else	if(presentAbsent.equals("Present")) {
				if(objmarks>=10 && EmailExcelLetter>=14 && speedMarks>=16)
				{
					
					Result="PASS";
					row1.createCell(4).setCellValue(Result);
					Reporter.log(seatNumber+" is "+Result,true);
					assertEquals(Result, result);

				}
				else
				{
					String result1="FAIL";
					row1.createCell(4).setCellValue(result1);
					Reporter.log(seatNumber+" is "+result1,true);
					assertEquals(result1, result);
				}

				if(percent>=75)
				{
					if(result.equals("PASS")&&grade.equals("A+"))
					{
						Reporter.log(seatNumber+ " is PASS with A+ grade",true);
						row1.createCell(6).setCellValue("A+");
					}
					else
					{
						Reporter.log(seatNumber+" is FAIL with - grade",true);
						row1.createCell(6).setCellValue("-");
						//					assertEquals( "A+",grade);
					}
				}
				else if(percent >= 60 && percent<=74.99)
				{
					if(result.equals("PASS")&&grade.equals("A"))
					{
						Reporter.log(seatNumber+" is PASS with A grade",true);
						row1.createCell(6).setCellValue("A");
					}
					else
					{
						Reporter.log(seatNumber+ " is FAIL with - grade",true);
						row1.createCell(6).setCellValue("-");
						//					assertEquals( "A",grade);
					}
				}
				else if(percent >= 50 && percent <= 59.99)
				{
					if(result.equals("PASS")&&grade.equals("B"))
					{
						Reporter.log(seatNumber+" is PASS with B grade",true);
						row1.createCell(6).setCellValue("B");
					}
					else
					{
						Reporter.log(seatNumber+ " is FAIL with - grade",true);
						row1.createCell(6).setCellValue("-");
						//					assertEquals( "B",grade);
					}
				}
				else if(percent >= 40 && percent <= 49.99)
				{
					if(result.equals("PASS")&&grade.equals("C"))
					{
						Reporter.log(seatNumber+" is PASS with C grade",true);
						row1.createCell(6).setCellValue("C");
					}
					else
					{
						Reporter.log(seatNumber+" is FAIL with - grade",true);
						row1.createCell(6).setCellValue("-");
						//					assertEquals( "C",grade);
					}
				}

				else
				{
					Reporter.log(seatNumber+ " Grade is - ",true);
					row1.createCell(6).setCellValue("-");
				}
				

			}
			
			else
			{
				Reporter.log(seatNumber+"  Student is Absent");
			}
		}FileOutputStream fileOut = new FileOutputStream(new File("C:\\Users\\User\\Desktop\\TestExcelFileSSD.xlsx"));
	     // Write the workbook data to the file
        workbook1.write(fileOut);
		}
	}


