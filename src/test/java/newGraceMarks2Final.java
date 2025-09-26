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
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.testng.Reporter;
import org.testng.annotations.Test;
import org.apache.poi.util.IOUtils;

public class newGraceMarks2Final implements IAutoConstant {
	@Test
	public void GraceMarksCal() throws EncryptedDocumentException, IOException
	{
		File file = new File("C:\\Users\\User\\Desktop\\Grace marks\\ssd\\25-08-2025 After second change\\RESULT.xlsx");
		FileInputStream fis = new FileInputStream(file);
		IOUtils.setByteArrayMaxOverride(200000000);
		Workbook wb = WorkbookFactory.create(fis);
		Sheet sheet = wb.getSheet("30");
		int rc = sheet.getLastRowNum();
		Reporter.log(Integer.toString(rc),true);
		DataFormatter df = new DataFormatter();
		@SuppressWarnings("resource")
		Workbook workbook1 = new SXSSFWorkbook();
		IOUtils.setByteArrayMaxOverride(200000000);
		Sheet sheet1 = workbook1.createSheet("TestSheet");
		Row header = sheet1.createRow(0);
		header.createCell(0).setCellValue("seatNumber");
		header.createCell(1).setCellValue("ExcelGraceMark");
		header.createCell(2).setCellValue("GraceMark");

		for(int i=1;i<=rc;i++)
		{
			double objGrace1 = 0;
			double EmailExcelLetterGrace1=0;
			double speedGrace1=0;
			Row row = sheet.getRow(i);
			String seatNumber = df.formatCellValue(row.getCell(3));
			String obj = df.formatCellValue(row.getCell(11));
			Double objMarks = Double.valueOf(obj);
			String speed = df.formatCellValue(row.getCell(16));
			Double SpeedMarks = Double.valueOf(speed);
			String email = df.formatCellValue(row.getCell(13));
			Double email1 = Double.valueOf(email);
			String stmt = df.formatCellValue(row.getCell(15));
			Double stmt1 = Double.valueOf(stmt);
			String ltr = df.formatCellValue(row.getCell(14));
			Double ltr1 = Double.valueOf(ltr);
			String presentAbsent = df.formatCellValue(row.getCell(30));
			String graceex = df.formatCellValue(row.getCell(34));
			Double graceExcel = Double.valueOf(graceex);
			double EmailExcelLetter = email1+stmt1+ltr1;
			double objGrace = 10-objMarks;
			double EmailExcelLetterGrace = 14-EmailExcelLetter;
			double speedGrace = 16-SpeedMarks;
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
			double totalGrace = objGrace1+speedGrace1+EmailExcelLetterGrace1;
			double speedExcelgrace = speedGrace1+EmailExcelLetterGrace1;
			if(totalGrace==4 || totalGrace<=4 && totalGrace!=0 )
				if(EmailExcelLetter>=12 && objMarks>=8 && SpeedMarks>=14 && speedExcelgrace<=2) 
				{ 
					if(EmailExcelLetterGrace1!=0||EmailExcelLetterGrace1<=1 ||
							speedGrace1!=0||speedGrace1<=1 || objGrace1 >= 0 || objGrace1 <= 2  )
					{
						Reporter.log(seatNumber+ " "+totalGrace, true);						
						Row row1 = sheet1.createRow(i);
						row1.createCell(0).setCellValue(seatNumber);
						row1.createCell(1).setCellValue(graceExcel);
						row1.createCell(2).setCellValue(totalGrace);
						Reporter.log(seatNumber+ " " +totalGrace);
					}
				}
				else {
					Reporter.log(seatNumber+" student is fail");			
				}			
		}
			
		else
		{
			Reporter.log(seatNumber+"  Student is Absent");
		}
			}
		FileOutputStream fileOut = new FileOutputStream(new File("C:\\Users\\User\\Desktop\\graceTestSSD.xlsx"));
		// Write the workbook data to the file
		workbook1.write(fileOut);
		
		 ((SXSSFWorkbook) workbook1).dispose();
	}

}
