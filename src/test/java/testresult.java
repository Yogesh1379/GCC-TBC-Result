import static org.testng.Assert.assertEquals;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbookFactory;
import org.testng.annotations.Test;

public class testresult {
	
		@Test(priority = 1)
		public  void readexcel() throws InterruptedException, IOException
		{
			
			
			File file =new File("C:\\Users\\User\\Desktop\\Book2.xlsx");
		      //Create object of FileInputStream class to read excel file
		      FileInputStream inputstream=new FileInputStream(file);
		      Workbook wb = new XSSFWorkbookFactory().create(inputstream);
		      //Creating Sheet object using the sheet Name
		      XSSFSheet sheet=(XSSFSheet) wb.getSheet("Sheet1");
		      int rowxount=sheet.getLastRowNum();
		      int columncount=sheet.getRow(1).getLastCellNum();
		      
		      DataFormatter df=new DataFormatter();
		      
		      
		      for (int i = 1; i < rowxount; i++) {
		          XSSFRow r=sheet.getRow(i);
		          String seatnumber=df.formatCellValue(r.getCell(0));
		          String obj=df.formatCellValue(r.getCell(13));
		          Double obj1 = Double.valueOf(obj);
		          String speed=df.formatCellValue(r.getCell(15));
		          Double speed1 = Double.valueOf(speed);
	              String stmtletter=df.formatCellValue(r.getCell(14));
	              Double stmtletter1 = Double.valueOf(stmtletter);
		          String percentage=df.formatCellValue(r.getCell(7));
		          Double percentage1 = Double.valueOf(percentage);
		          String result =df.formatCellValue(r.getCell(8));
		          String grade=df.formatCellValue(r.getCell(9));
		         
		       System.out.println("seatnumber is: "+seatnumber);
		       String PASS="PASS";
		       String Gradeactual="A+";
		       String GradeactualA="A";
		       String GradeactualB="B";
		       String GradeactualC="C";
		      
		       
		       if (obj1>=10 && stmtletter1>=14 && speed1>=16) {
		    	   
		    	   System.out.println("Result is: "+PASS);
		    	 
		    	   assertEquals(PASS, result);
		    	   
				
			} else {
				String FAIL="FAIL";
		    	   System.out.println("Result is: "+FAIL);
		    	   

			}
		       if(percentage1>=75)
		       {
		    	   
		           if(result.equals(PASS) && grade.equals(Gradeactual))
		           {
		               System.out.println("Student Grade:  "+Gradeactual);
		           }
		           else
		        	   System.out.println("Student fail with -  ");
		       }
		       
		       
		       else if (percentage1 >= 60 && percentage1<=74.99)
		       {
		    	   if(result.equals(PASS) &&  grade.equals(GradeactualA))
		           {
		               System.out.println("Student Grade:  "+GradeactualA);
		           }
		           else
		        	   System.out.println("Student fail with -");
		    	       
		       }
		       
		       
		       else if (percentage1 >= 50  && percentage1<=59.99)
		       {
		    	   if(result.equals(PASS) &&  grade.equals(GradeactualB))
		           {
		               System.out.println("Student Grade:  "+GradeactualB);
		           }
		           else
		        	   System.out.println("Student fail with - ");
		    	      
		       }
		       
		       
		       else if (percentage1 >= 40   && percentage1<=49.99)
		       {
		    	   if(result.equals(PASS)&&  grade.equals(GradeactualC))
		           {
		               System.out.println("Student Grade:  "+GradeactualC);
		           }
		           else
		        	   System.out.println("Student fail with -");
		    	      
		       }
		       
		       else
		       {
		           System.out.println("grade is -");
		       }
	      
		      }
	    
		}
	}



