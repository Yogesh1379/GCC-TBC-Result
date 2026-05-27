package Speed_Question_Paper_Verification;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.testng.Reporter;
import org.testng.annotations.Test;


import SendEmail_Package.SendMailofSpeedQuestionVerify;

public class Speed_question_Paper_verify {
	@Test
	public void questionFormatting() throws IOException {

		String folderpath="F:\\GCCTBC-APR 2026\\wrong support question\\New folder (2)";
		String Outpath="F:\\GCCTBC-APR 2026\\qestion allocation\\FINAL Repeater Question files April 2026\\New folder\\bookmark verification1.xlsx";
		Workbook outworkbook = new XSSFWorkbook();
		Sheet outsheet = outworkbook.createSheet("Bookmark Validation");
		Row outheader = outsheet.createRow(0);
		outheader.createCell(0).setCellValue("File Name");
		outheader.createCell(1).setCellValue("Errors");
		outheader.createCell(2).setCellValue("Words");
		File folder = new File(folderpath);

		File[] files = folder.listFiles((dir, name) ->name.endsWith(".docx") );

		if(files==null) {
//			System.out.println("Folder is empty");
			Reporter.log("Folder is empty",true);
		}
		int rowcount=1;
		for(File file:files) {
			
			
			String filename = file.getName();
			FileInputStream fis = new FileInputStream(file);
			@SuppressWarnings("resource")
			XWPFDocument doc = new XWPFDocument(fis);
			List<XWPFParagraph> para = doc.getParagraphs();

			int numberofpara = para.size();

			int TabCount = 0;
			
			for(XWPFParagraph paragraph:para) {

				String text = paragraph.getText();
				if(text==null) {
					Row row = outsheet.createRow(rowcount++);
					
					Reporter.log(filename+" :-paragraph is null",true);
					row.createCell(0).setCellValue(filename);
					row.createCell(1).setCellValue("paragraph is null");
				}
				if(text!=null) {
					if(text.contains("\t")) {
						int count = text.length() - text.replace("\t", "").length();
						TabCount+=count;
					}
					String[] words = text.split("\\s+");
					for(String word:words) {

						if(word.matches("[a-zA-Z]+[.,][a-zA-Z]+")) {
							Row row = outsheet.createRow(rowcount++);
						
							Reporter.log(word,true);
							row.createCell(0).setCellValue(filename);
							row.createCell(1).setCellValue("Space is missing after word");
							row.createCell(2).setCellValue(word);
						}
					}

					Pattern pattern = Pattern.compile("(\\S+)\\s{2,}");
					Matcher matcher = pattern.matcher(text);
					while(matcher.find()) {
						Row row = outsheet.createRow(rowcount++);
						String previousWord = matcher.group(1);
						row.createCell(0).setCellValue(filename);
						row.createCell(1).setCellValue("Two / More than 2 spaces after word");
						row.createCell(2).setCellValue(previousWord);
//						System.out.println(previousWord+" **********");
					}
					Pattern patter1 = Pattern.compile("(\\S+)\\t");
					Matcher matcher1 = patter1.matcher(text);
					while(matcher1.find()) {
						Row row = outsheet.createRow(rowcount++);
//						String previousWord = matcher.group(1);
						row.createCell(0).setCellValue(filename);
						row.createCell(1).setCellValue("Before tab character is present");
//						row.createCell(2).setCellValue(previousWord);
					}
					
					Pattern patter2 = Pattern.compile("\\s+\\t");
					Matcher matcher2 = patter2.matcher(text);
					while(matcher2.find()) {
						Row row = outsheet.createRow(rowcount++);
//						String previousWord = matcher.group(1);
						row.createCell(0).setCellValue(filename);
						row.createCell(1).setCellValue("Before tab Space is present");
					}

				}

			}

			if(TabCount!=numberofpara) {
				Row row = outsheet.createRow(rowcount++);
		
				Reporter.log(filename+" :- Tab is missing",true);
				row.createCell(0).setCellValue(filename);
				row.createCell(1).setCellValue("Tab is missing");
				row.createCell(2).setCellValue("Required Tab count :-"+numberofpara+", Tabs in the file is:-"+TabCount);
			
				Reporter.log(TabCount +" :-TabCount",true);
		
				Reporter.log(numberofpara+" :- numberofpara",true);
			}
			
		}
		for(int i=0; i<3;i++) {
			outsheet.autoSizeColumn(i);
		}
		FileOutputStream fos = new FileOutputStream(Outpath);
		outworkbook.write(fos);
		outworkbook.close();

		Reporter.log("Question verification done and excel is written in output path",true);
		try {

			SendMailofSpeedQuestionVerify.sendReport(Outpath);

		} catch (Exception e) {

		    e.printStackTrace();
		}
	}
}


