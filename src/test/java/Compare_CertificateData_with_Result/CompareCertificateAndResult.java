package Compare_CertificateData_with_Result;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

public class CompareCertificateAndResult {
	
	@Test 
	public void compare() throws EncryptedDocumentException, IOException {
		//final result  msce result 
	File resultFile=new File("F:\\GCCTBC-APR 2026\\Result\\MSCE FINAL\\FinalResult.xlsx");
	File CertificateFile=new File("F:\\GCCTBC-APR 2026\\certificate data\\CertificateData.xlsx");
	String outputPath="F:\\GCCTBC-APR 2026\\Result\\MSCE FINAL\\compare Certificate result.xlsx";
	
	FileInputStream resultF = new FileInputStream(resultFile);
	Workbook Rwb = WorkbookFactory.create(resultF);
	Workbook Cwb = WorkbookFactory.create(new FileInputStream(CertificateFile));
	Sheet Rsheet = Rwb.getSheetAt(0);
	Sheet CSheet = Cwb.getSheetAt(0);
	
	DataFormatter df = new DataFormatter();
	
	HashMap<String, Integer> RHeader=new HashMap<>();  
	Row RheaderRow = Rsheet.getRow(0);
	int idex=0;
	for(int i=0;i<RheaderRow.getLastCellNum();i++) {
		String header = df.formatCellValue(RheaderRow.getCell(i)).trim();
		RHeader.put(header, idex);
		
		idex++;
	}
//	System.out.println(RHeader);
	
	
	Map<String, String[]>resultData=new HashMap<>();
	
	for(int i=1;i<=Rsheet.getLastRowNum();i++) {
		Row Rrow = Rsheet.getRow(i);
		if(Rrow==null)
			continue;
		
		String SEAT_NO = df.formatCellValue(Rrow.getCell(RHeader.get("SEAT NO")));
//		System.out.println(SEAT_NO);
		if(SEAT_NO.isEmpty())
			continue;
		
		resultData.put(SEAT_NO, new String[]{
			df.formatCellValue(Rrow.getCell(RHeader.get("NAME"))),
			df.formatCellValue(Rrow.getCell(RHeader.get("MOTHER'S NAME"))),
			df.formatCellValue(Rrow.getCell(RHeader.get("EXAM CENTER CODE"))),
			df.formatCellValue(Rrow.getCell(RHeader.get("INSTITUTE CODE"))),
			df.formatCellValue(Rrow.getCell(RHeader.get("INSTITUTE NAME"))),
			df.formatCellValue(Rrow.getCell(RHeader.get("OBJECTIVE MARKS WITH GRACE"))),
			df.formatCellValue(Rrow.getCell(RHeader.get("EMAIL MARKS"))),
			df.formatCellValue(Rrow.getCell(RHeader.get("LETTER MARKS"))),
			df.formatCellValue(Rrow.getCell(RHeader.get("STATEMENT MARKS"))),
			df.formatCellValue(Rrow.getCell(RHeader.get("SPEED MARKS WITH GRACE"))),
			df.formatCellValue(Rrow.getCell(RHeader.get("SECTION 1 (LETTER+STATEMENT+EMAIL) WITH GRACE"))),
			df.formatCellValue(Rrow.getCell(RHeader.get("TOTAL MARKS"))),
			df.formatCellValue(Rrow.getCell(RHeader.get("RESULT"))),
			df.formatCellValue(Rrow.getCell(RHeader.get("GRADE")))				
		});
	}
//	System.out.println(resultData);
	 Workbook outWb = new XSSFWorkbook();
     Sheet outSheet = outWb.createSheet("Comparison");

     Row header = outSheet.createRow(0);
 
     String[] headers = {
             "Seat No",
             "Result's Std Name", "Certificate's Std Name",
             "Result's Mother Name", "Certificate's Mother Name",
             "Result's Obj Grace", "Certificate's Obj Grace",
             "Result's SECTION1LETTE Grace", "Certificate's SECTION1LETTE Grace",
             "Result's Speed Grace", "Certificate's Speed Grace",
             "Result's Percentage", "Certificate's Percentage",
             "Result's Result", "Certificate's Result",
             "Result's Grade", "Certificate's Grade",
             "Result's Center Code", "Certificate's Center Code",
             "Result's Inst Code", "Certificate's Inst Code",  
             "Result's Inst name", "Certificate's Inst name",  
             "Status","Wrong field"
     };

     for (int i = 0; i < headers.length; i++) {
         header.createCell(i).setCellValue(headers[i]);
     }

     int rowIndex = 1;
     
     HashMap<String, Integer> CHeader=new HashMap<>();  
 	Row CheaderRow = CSheet.getRow(0);
 	int index=0;
 	
 
 	
 	for(int i=0;i<CheaderRow.getLastCellNum();i++) {
 		String Cheader = df.formatCellValue(CheaderRow.getCell(i)).trim();
 		CHeader.put(Cheader, index);
 		
 		index++;
 	}
//	System.out.println(CHeader);
 	HashSet<String> set = new HashSet<>();
 	
	for(int i=1;i<=CSheet.getLastRowNum();i++)
	{
		Row Crow = CSheet.getRow(i);
		
		if(Crow==null)
			continue;
		
		String SeatNo = df.formatCellValue(Crow.getCell(CHeader.get("SeatNo")));
		System.out.println("Checking for seat no :- "+ SeatNo);
		if(SeatNo.isEmpty())
			continue;
		String InstCode = df.formatCellValue(Crow.getCell(CHeader.get("InstCode")));
		String NAME = df.formatCellValue(Crow.getCell(CHeader.get("NAME")));
		String MotherName = df.formatCellValue(Crow.getCell(CHeader.get("MotherName")));
		String Grade = df.formatCellValue(Crow.getCell(CHeader.get("Grade")));
		String OBJECTIVE = df.formatCellValue(Crow.getCell(CHeader.get("OBJECTIVE")));
		String Email = df.formatCellValue(Crow.getCell(CHeader.get("Email")));
		String Letter = df.formatCellValue(Crow.getCell(CHeader.get("Letter")));
		String Statement =	df.formatCellValue(Crow.getCell(CHeader.get("Statement")));
		String SPEED = df.formatCellValue(Crow.getCell(CHeader.get("SPEED")));
		String TOTAL = df.formatCellValue(Crow.getCell(CHeader.get("TOTAL")));
		String EXAMCENTERCODE = df.formatCellValue(Crow.getCell(CHeader.get("EXAMCENTERCODE")));
		String SECTION1LETTE = df.formatCellValue(Crow.getCell(CHeader.get("SECTION1LETTE")));
		String INSTITUTENAME = df.formatCellValue(Crow.getCell(CHeader.get("INSTITUTENAME")));
		String certificateNo = df.formatCellValue(Crow.getCell(CHeader.get("SRNO")));
		int rowi=1;
		if(!set.add(certificateNo)) {
			 Sheet outSheet1 = outWb.createSheet("Duplicate cert no.");

		     Row header1 = outSheet1.createRow(0);
		 header1.createCell(0).setCellValue("Duplicate Certificate no.");
		Row cerrow = outSheet1.createRow(rowi++);
		int c=0;
		cerrow.createCell(c++).setCellValue(certificateNo);
		}
		
		String[] result=resultData.get(SeatNo);

		if(result==null) {
			Row outRow = outSheet.createRow(rowIndex++);
			outRow.createCell(0).setCellValue(SeatNo);
			outRow.createCell(23).setCellValue("Seat number not found in Result file");
		continue;		
		}
		
		boolean match=
				result[0].equals(NAME)&&
				result[1].equals(MotherName)&&
				result[2].equals(EXAMCENTERCODE)&&
				result[3].equals(InstCode)&&
				result[4].equals(INSTITUTENAME)&&
				result[5].equals(OBJECTIVE)&&
				result[6].equals(Email)&&
				result[7].equals(Letter)&&
				result[8].equals(Statement)&&
				result[9].equals(SPEED)&&
				result[10].equals(SECTION1LETTE)&&
				result[11].equals(TOTAL)&&
				result[13].equals(Grade);
		
		if(!match) {
			Row outRow = outSheet.createRow(rowIndex++);
			outRow.createCell(0).setCellValue(SeatNo);
			int c=1;
			outRow.createCell(c++).setCellValue(result[0]); //name
			outRow.createCell(c++).setCellValue(NAME);
			outRow.createCell(c++).setCellValue(result[1]); //mother
			outRow.createCell(c++).setCellValue(MotherName);
			outRow.createCell(c++).setCellValue(result[5]); //obj
			outRow.createCell(c++).setCellValue(OBJECTIVE);
			outRow.createCell(c++).setCellValue(result[10]); //section1
			outRow.createCell(c++).setCellValue(SECTION1LETTE);
			outRow.createCell(c++).setCellValue(result[9]); // speed
			outRow.createCell(c++).setCellValue(SPEED);
			outRow.createCell(c++).setCellValue(result[11]); // percent
			outRow.createCell(c++).setCellValue(TOTAL);
			outRow.createCell(c++).setCellValue(result[12]); //result
			outRow.createCell(c++).setCellValue("PASS");
			outRow.createCell(c++).setCellValue(result[13]);  //grade 
			outRow.createCell(c++).setCellValue(Grade); 
			outRow.createCell(c++).setCellValue(result[2]); //center code
			outRow.createCell(c++).setCellValue(EXAMCENTERCODE); 
			outRow.createCell(c++).setCellValue(result[3]);  //inst code
			outRow.createCell(c++).setCellValue(InstCode); 
			outRow.createCell(c++).setCellValue(result[4]);
			outRow.createCell(c++).setCellValue(INSTITUTENAME);
			 outRow.createCell(c).setCellValue("MISMATCH");
			    StringBuilder mismatchColumns = new StringBuilder();

			    if (!result[0].equals(NAME))
			        mismatchColumns.append("NAME, ");

			    if (!result[1].equals(MotherName))
			        mismatchColumns.append("MOTHER NAME, ");

			    if (!result[2].equals(EXAMCENTERCODE))
			        mismatchColumns.append("EXAM CENTER CODE, ");

			    if (!result[3].equals(InstCode))
			        mismatchColumns.append("INSTITUTE CODE, ");

			    if (!result[4].equals(INSTITUTENAME))
			        mismatchColumns.append("INSTITUTE NAME, ");

			    if (!result[5].equals(OBJECTIVE))
			        mismatchColumns.append("OBJECTIVE, ");

			    if (!result[6].equals(Email))
			        mismatchColumns.append("EMAIL, ");

			    if (!result[7].equals(Letter))
			        mismatchColumns.append("LETTER, ");

			    if (!result[8].equals(Statement))
			        mismatchColumns.append("STATEMENT, ");

			    if (!result[9].equals(SPEED))
			        mismatchColumns.append("SPEED, ");

			    if (!result[10].equals(SECTION1LETTE))
			        mismatchColumns.append("SECTION1LETTE, ");

			    if (!result[11].equals(TOTAL))
			        mismatchColumns.append("TOTAL, ");

			    if (!result[13].equals(Grade))
			        mismatchColumns.append("GRADE, ");
			    outRow.createCell(24).setCellValue(
			            mismatchColumns.toString()
			    );    
			    
			
		}
				
		
	}
	
	for (int i = 0; i < headers.length; i++) {
        outSheet.autoSizeColumn(i);
    }
	 FileOutputStream fos = new FileOutputStream(outputPath);

     outWb.write(fos);

     fos.close();

     outWb.close();
     Rwb.close();
     Cwb.close();

     System.out.println("Comparison completed successfully.");
     System.out.println("Output file : " + outputPath);
 	
	}

}
