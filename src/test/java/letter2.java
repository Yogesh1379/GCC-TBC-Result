import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.io.*;
import java.util.*;
import java.util.regex.*;

public class letter2 {
	static File modelFile = null;
	static List<String> modelParas;

	@SuppressWarnings("resource")
	public static void main(String[] args) throws Exception {
		File folder = new File("F:\\desktop backup 5 july 25\\GCC TBC MAY 2025 AUTO\\New folder\\New folder");
		File compareDir = new File(folder.getParentFile(), folder.getName() + "_Compared");
		if (!compareDir.exists()) compareDir.mkdirs();

		Workbook workbook = new XSSFWorkbook();
		Sheet sheet = workbook.createSheet("Results");
		String[] columns = {
				"File Name", "Extra Word", "Missing Word", "Wrong Word", "Total Mistakes", "Final Marks", "Mistakes Summary"
		};
		Row header = sheet.createRow(0);
		for (int i = 0; i < columns.length; i++) {
			header.createCell(i).setCellValue(columns[i]);
		}

		int rowNum = 1;
		File[] docxFiles = folder.listFiles((dir, name) -> name.toLowerCase().endsWith(".docx") && !name.startsWith("~$"));
		if (docxFiles == null) {
			System.out.println("❌ No .docx files found.");
			return;
		}

		for (File studentFile : docxFiles) {

			String stdfile = studentFile.getName();
			//        	System.out.println(stdfile);

			String[] parts = stdfile.split("_");
			String seatno = parts[1];
			String batchname = parts[2];
			String section = parts[0];
			System.out.println(seatno);
			if (seatno.length()>=10)
			{
				String course = seatno.substring(4, 6);
				int course1=0;
				if(course.equals("15")  )
				{
					//        			nteger.parseInt(course);
					course1=1;
				}
				else if(course.equals("16")) {
					//        			System.out.println("16");
					course1=2;
				}
				//        		System.out.println(course);
				FileInputStream fis = new FileInputStream("F:\\desktop backup 5 july 25\\GCC TBC MAY 2025 AUTO\\allocation\\BathwiseSubjective.xlsx");
				Workbook workbook1 = new XSSFWorkbook(fis);
				Sheet sheet1 = workbook1.getSheetAt(0); // assuming data is in first sheet

				//        		     modelFile = null;
				int rowCount = 0;
				for (Row row : sheet1) {
					if (rowCount++ == 0) continue;
					Cell batchCell = row.getCell(6); // assuming Batch ID in col 0
					Cell courseCell = row.getCell(5); // assuming Course in col 1
					Cell subjectiveCell = row.getCell(2); // assuming SubjectiveFile in col 2

					if (batchCell == null || courseCell == null || subjectiveCell == null) continue;

					String excelBatch = getCellValueAsString(batchCell);
					String excelCourse = getCellValueAsString(courseCell);

					int excelCourse1 = Integer.parseInt(excelCourse);
					if (excelBatch.equals(batchname)&& excelCourse1==course1) {

						String fileCandidate = getCellValueAsString(subjectiveCell);
						//        		        	System.out.println(course1);
						if (fileCandidate.startsWith("Eng30 Ltr") && course1==1) {
//							System.out.println(fileCandidate);
							modelFile=new File("F:\\desktop backup 5 july 25\\GCC TBC MAY 2025 AUTO\\allocation\\New folder\\Model Anwers\\"+fileCandidate);

							break;   
						}

						else if(fileCandidate.startsWith("Eng40 Resume")||fileCandidate.startsWith("Eng40 B Ltr") && course1==2) {
//							System.out.println(fileCandidate);
							modelFile=new File("F:\\desktop backup 5 july 25\\GCC TBC MAY 2025 AUTO\\allocation\\New folder\\Model Anwers\\"+fileCandidate);

							break;
						}
					}

					//        		    workbook.close();

				}
			}
		
			modelParas = extractFormattedParagraphs(modelFile);
			List<String> studentParas = extractFormattedParagraphs(studentFile);
			List<String[]> wordMistakes = new ArrayList<>();

			XWPFDocument outputDoc = new XWPFDocument();
			int refIndex = 0, stuIndex = 0;
			while (stuIndex < studentParas.size()) {
				String studentLine = studentParas.get(stuIndex);
				String modelLine = refIndex < modelParas.size() ? modelParas.get(refIndex) : "";

//				*********
				
//				*******
				XWPFParagraph para = outputDoc.createParagraph();
				compareWords(modelLine, studentLine, para, wordMistakes, studentFile.getName());
				refIndex++;
				stuIndex++;
			}

			// Count word-level mistakes
			int extraWord = 0, missingWord = 0, wrongWord = 0;
			for (String[] m : wordMistakes) {
				switch (m[1]) {
				case "Extra Word": extraWord++; break;
				case "Missing Word": missingWord++; break;
				case "Wrong Word": wrongWord++; break;
				}
			}
			int totalMistakes = wordMistakes.size();
			int obtainedMarks = Math.max(0, 40 - totalMistakes);

			// Write summary in DOCX
			XWPFParagraph summaryPara = outputDoc.createParagraph();
			XWPFRun run;
			if (extraWord > 0) { run = summaryPara.createRun(); run.setText("Extra Word: " + extraWord); run.addBreak(); }
			if (missingWord > 0) { run = summaryPara.createRun(); run.setText("Missing Word: " + missingWord); run.addBreak(); }
			if (wrongWord > 0) { run = summaryPara.createRun(); run.setText("Wrong Word: " + wrongWord); run.addBreak(); }
			run = summaryPara.createRun(); run.setText("Total Mistakes: " + totalMistakes); run.addBreak();
			run = summaryPara.createRun(); run.setText("Final Marks: " + obtainedMarks + " / 40");

			// Save compared DOCX
			File comparedFile = new File(compareDir, studentFile.getName().replace(".docx", "_Compared.docx"));
			try (FileOutputStream out = new FileOutputStream(comparedFile)) {
				outputDoc.write(out);
			}

			// Write to Excel
			Row row = sheet.createRow(rowNum++);
			row.createCell(0).setCellValue(studentFile.getName());
			row.createCell(1).setCellValue(extraWord);
			row.createCell(2).setCellValue(missingWord);
			row.createCell(3).setCellValue(wrongWord);
			row.createCell(4).setCellValue(totalMistakes);
			row.createCell(5).setCellValue(obtainedMarks);
			StringBuilder summary = new StringBuilder();
			for (String[] m : wordMistakes) {
				summary.append("[").append(m[1]).append(" - ").append(m[2]).append("] ");
			}
			row.createCell(6).setCellValue(summary.toString().trim());
		}

		for (int i = 0; i < columns.length; i++) sheet.autoSizeColumn(i);
		try (FileOutputStream out = new FileOutputStream(new File(folder.getParent(), "Comparison_Result.xlsx"))) {
			workbook.write(out);
		}
		workbook.close();
		System.out.println("✅ All students processed.");
	}

	private static void compareWords(String model, String student, XWPFParagraph para,
			List<String[]> wordMistakes, String fileName) {
		
		model   = model.trim().replaceAll("\\s+([,;:!?])", "$1");
        student = student.trim().replaceAll("\\s+([,;:!?])", "$1");
        model = model.replaceAll("\\.(?!\\s|$)", ". ");   // add space after . if not followed by space or end
        student = student.replaceAll("\\.(?!\\s|$)", ". ");
    	
		String[] modelWords = model.trim().split("\\s+");
		String[] studentWords = student.trim().split("\\s+");
		List<int[]> lcs = computeLCS(modelWords, studentWords);

		int i = 0, j = 0, k = 0;
		while (i < modelWords.length || j < studentWords.length) {

			// Matching word via LCS
			if (k < lcs.size() && i == lcs.get(k)[0] && j == lcs.get(k)[1]) {
				addTextRun(para, studentWords[j] + " ", false);
				i++; j++; k++;
			}
			// Wrong word: both model and student have a word, but it's not in LCS and mismatched
			else if (i < modelWords.length && j < studentWords.length &&
					(k >= lcs.size() || i < lcs.get(k)[0] && j < lcs.get(k)[1]) &&
					!normalize(modelWords[i]).equals(normalize(studentWords[j]))) {
				highlight(para.createRun(), "[" + modelWords[i] + "/" + studentWords[j] + "]");
				wordMistakes.add(new String[]{fileName, "Wrong Word", modelWords[i] + "/" + studentWords[j]});
				i++; j++;
			}
			// Missing word from student (model has, student doesn't)
			else if (k < lcs.size() && i < lcs.get(k)[0]) {
				highlight(para.createRun(), "[" + modelWords[i] + "]");
				wordMistakes.add(new String[]{fileName, "Missing Word", modelWords[i]});
				i++;
			}
			// Extra word in student (student has, model doesn't)
			else if (k < lcs.size() && j < lcs.get(k)[1]) {
				highlight(para.createRun(), "[" + studentWords[j] + "]");
				wordMistakes.add(new String[]{fileName, "Extra Word", studentWords[j]});
				j++;
			}
			// Remaining model words (after LCS)
			else if (i < modelWords.length && j >= studentWords.length) {
				highlight(para.createRun(), "[" + modelWords[i] + "]");
				wordMistakes.add(new String[]{fileName, "Missing Word", modelWords[i]});
				i++;
			}
			// Remaining student words (after LCS)
			else if (j < studentWords.length && i >= modelWords.length) {
				highlight(para.createRun(), "[" + studentWords[j] + "]");
				wordMistakes.add(new String[]{fileName, "Extra Word", studentWords[j]});
				j++;
			}
			else {
				// Safety fallback
				i++;
				j++;
			}
		}
	}


	private static void addTextRun(XWPFParagraph para, String text, boolean highlight) {
		XWPFRun run = para.createRun();
		run.setText(text);
		if (highlight) {
			CTRPr rpr = run.getCTR().isSetRPr() ? run.getCTR().getRPr() : run.getCTR().addNewRPr();
			CTShd shd = rpr.addNewShd();
			shd.setVal(STShd.CLEAR);
			shd.setFill("FF0000");
		}
	}

	private static void highlight(XWPFRun run, String text) {
		run.setText(text + " ");
		run.setColor("FF0000");
	}

	private static List<String> extractFormattedParagraphs(File file) throws IOException {
		List<String> paras = new ArrayList<>();
		try (FileInputStream fis = new FileInputStream(file);
				XWPFDocument doc = new XWPFDocument(fis)) {
			for (XWPFParagraph p : doc.getParagraphs()) {
				StringBuilder fullText = new StringBuilder();
				for (XWPFRun run : p.getRuns()) fullText.append(run.toString());
				paras.add(fullText.toString());
			}
		}
		return paras;
	}

	private static List<int[]> computeLCS(String[] a, String[] b) {
		int m = a.length, n = b.length;
		int[][] dp = new int[m+1][n+1];
		for (int i=1; i<=m; i++)
			for (int j=1; j<=n; j++) {
			if (normalize(a[i-1]).equals(normalize(b[j-1]))) dp[i][j] = dp[i-1][j-1] + 1;
			else dp[i][j] = Math.max(dp[i-1][j], dp[i][j-1]);
		}
		List<int[]> lcs = new ArrayList<>();
		int i=m, j=n;
		while (i>0 && j>0) {
			if (normalize(a[i-1]).equals(normalize(b[j-1]))) {
				lcs.add(0, new int[]{i-1,j-1}); i--; j--;
			} else if (dp[i-1][j]>=dp[i][j-1]) i--; else j--;
		}
		return lcs;
	}
	private static String getCellValueAsString(Cell cell) {
		switch (cell.getCellType()) {
		case STRING:
			return cell.getStringCellValue().trim();
		case NUMERIC:
			return String.valueOf((int) cell.getNumericCellValue());
		default:
			return "";
		}
	}
	private static String normalize(String w) {
		if (w==null) return null;
		return w.trim()
				.replace("’","'")
				.replace("‘","'")
				 .replace("‘", "'")
			        .replace("‘", "'");
	}
}
