import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.io.*;
import java.util.*;
import java.util.regex.*;

public class StudentAnswerBatchProcessor22 {
	static File modelFile=null;
	static List<String> modelParas;
	static int blankLineCount = 0;
    @SuppressWarnings("resource")
	public static void main(String[] args) throws Exception {
        File folder = new File("F:\\GCC  TBC December 2025\\Marking\\Eng 40\\Diff_SpeedAnswer\\zero cases");
        File compareDir = new File(folder.getParentFile(), folder.getName() + "_Compared");
        if (!compareDir.exists()) compareDir.mkdirs();
//        File modelFile = new File("C:\\Users\\User\\Desktop\\Eng30 Speed 4.docx");
//        List<String> modelParas = extractFormattedParagraphs(modelFile);

        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Results");
        String[] columns = {
                "File Name", "Total Mistakes", "Final Marks",
                "Extra Space", "Extra Tab", "Extra Enter",
                "Extra Word", "Missing Word", "Wrong Word",
                "Mistakes Summary"
        };
        Row header = sheet.createRow(0);
        for (int i = 0; i < columns.length; i++) {
            header.createCell(i).setCellValue(columns[i]);
        }

        int rowNum = 1;

        File[] docxFiles = folder.listFiles((dir, name) ->
                name.toLowerCase().endsWith(".docx") && !name.startsWith("~$"));

        if (docxFiles == null) {
            System.out.println("❌ No .docx files found.");
            return;
        }

        for (File studentFile : docxFiles) {
        	// test
        	
      String stdfile = studentFile.getName();
//        	System.out.println(stdfile);
        	
        	String[] parts = stdfile.split("_");
        	String seatno = parts[1];
        	String batchname = parts[2];
//        	String section = parts[0];
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
        		 FileInputStream fis = new FileInputStream("F:\\GCC  TBC December 2025\\Question\\allocation\\Batch Wise Subjective.xlsx");
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
        		            if (fileCandidate.startsWith("Eng30 Speed") && course1==1) {
        		            	System.out.println(fileCandidate);
        		            	modelFile=new File("F:\\GCC  TBC December 2025\\Question\\ENGLISH\\All 30&40\\"+fileCandidate);
        		            	
        		            	 break;   
        		            }
        		          
        		            else if(fileCandidate.startsWith("Eng 40 Speed") && course1==2) {
        		            	System.out.println(fileCandidate);
        		               	modelFile=new File("F:\\GCC  TBC December 2025\\Question\\ENGLISH\\All 30&40\\"+fileCandidate);
        		               
        		                break;
        		        }
        		    }

//        		    workbook.close();
        		       		
        		    }
        	}
        	
        	 List<String> modelParas = extractFormattedParagraphs(modelFile);
        	 
        	//test
            List<String> studentParas = extractFormattedParagraphs(studentFile);
            Map<String, Integer> mistakeCounts = new HashMap<>();
            List<String[]> wordMistakes = new ArrayList<>();

            XWPFDocument outputDoc = new XWPFDocument();
            int refIndex = 0, stuIndex = 0;

            while (stuIndex < studentParas.size()) {
                String studentLine = studentParas.get(stuIndex);
                String modelLine = refIndex < modelParas.size() ? modelParas.get(refIndex) : "";
// testttttt
//             // Trim off any trailing spaces first
//                studentLine = studentLine.stripTrailing();
//                modelLine   = modelLine.stripTrailing();
//
//                // Now check for a final dot:
//                boolean studentEndsWithDot = studentLine.endsWith(".");
//                boolean modelEndsWithDot   = modelLine.endsWith(".");
//
//                // Example: if you want to ignore any difference in that dot, you could:
//                //   if only one ends with a dot, strip it so further comparison treats them equally:
//                if (studentEndsWithDot && !modelEndsWithDot) {
//                    studentLine = studentLine.substring(0, studentLine.length() - 1);
//                }
//                else if (modelEndsWithDot && !studentEndsWithDot) {
//                    modelLine = modelLine.substring(0, modelLine.length() - 1);
//                }    
//                studentLine.lines().toString().matches("' \\$'");
            
                
                
                // Check for group of 3 or more consecutive extra blank lines
//                int blankLineCount = 0;
                int tempIndex = stuIndex;
                while (tempIndex < studentParas.size()) {
                    String tempLine = studentParas.get(tempIndex).trim();
                    String modelTempLine = (refIndex < modelParas.size()) ? modelParas.get(refIndex).trim() : "";
                    if (tempLine.isEmpty() && !modelTempLine.isEmpty()) {
                        blankLineCount++;
                        tempIndex++;
                    } else {
                        break;
                    }
                }
                
                if (blankLineCount >= 3) {
                    XWPFParagraph para = outputDoc.createParagraph();
                    XWPFRun run = para.createRun();
                    run.setText("[Extra Enters Group]");
                    run.setColor("FF0000");
                    incrementMistake("Extra Enter", mistakeCounts);
                    stuIndex += blankLineCount;
                    continue;
                }

                XWPFParagraph para = outputDoc.createParagraph();

                // Highlight 3-space groups
                Pattern pattern = Pattern.compile("( {3,})");
                Matcher matcher = pattern.matcher(studentLine);
                int lastIndex = 0;
                while (matcher.find()) {
                    int start = matcher.start(1);
                    int end = matcher.end(1);
                    String matchedSpaces = matcher.group(1);
                    if (start > lastIndex) {
                        addTextRun(para, studentLine.substring(lastIndex, start), false);
                    }
                    int spaceCount = matchedSpaces.length();
                    int fullGroups = spaceCount / 3;
                    int leftover = spaceCount % 3;
                    for (int i = 0; i < fullGroups; i++) {
                        addTextRun(para, "   ", true);
                        incrementMistake("Extra Space", mistakeCounts);
                    }
                    if (leftover > 0) {
                        addTextRun(para, matchedSpaces.substring(spaceCount - leftover), false);
                    }
                    lastIndex = end;
                }
                if (lastIndex < studentLine.length()) {
                    addTextRun(para, studentLine.substring(lastIndex), false);
                }

                // Tabs
                int tabCount = countMatches(studentLine, "\\t{3,}");
                if (tabCount > 0) {
                    XWPFRun run = para.createRun();
                    run.setText("[Extra Tabs]");
                    run.setColor("FF0000");
                    for (int i = 0; i < tabCount; i++) {
                        incrementMistake("Extra Tab", mistakeCounts);
                    }
                }

                // Extra tab at start
                if (studentLine.startsWith("\t") && !modelLine.startsWith("\t")) {
                    incrementMistake("Extra Tab", mistakeCounts);
                    XWPFRun run = para.createRun();
                    run.setText("[Extra Tab]");
                    run.setColor("FF0000");
                    studentLine = studentLine.replaceFirst("^\\t+", "");
                }

                
                compareWords(modelLine, studentLine, para, mistakeCounts, wordMistakes, studentFile.getName());
                refIndex++;
                stuIndex++;
            }

            // Add summary to output DOCX
            XWPFParagraph summaryTitle = outputDoc.createParagraph();
            XWPFRun titleRun = summaryTitle.createRun();
            titleRun.setText("Summary of Mistakes:");
            titleRun.setBold(true);
            titleRun.setFontSize(12);
            titleRun.setColor("000000");
            titleRun.addBreak();

            XWPFParagraph summaryPara = outputDoc.createParagraph();
            String[] types = {"Extra Space", "Extra Tab", "Extra Enter","Missing Tab"};
            for (String type : types) {
                int count = mistakeCounts.getOrDefault(type, 0);
                if (count > 0) {
                    XWPFRun run = summaryPara.createRun();
                    run.setText(type + ": " + count);
                    run.setColor("FF0000");
                    run.addBreak();
                }
            }

            int extraWord = 0, missingWord = 0, wrongWord = 0, wrongTab=0;
            for (String[] mistake : wordMistakes) {
                switch (mistake[1]) {
                    case "Extra Word":
                        extraWord++;
                        break;
                    case "Missing Word":
                        missingWord++;
                        break;
                    case "Wrong Word":
                        wrongWord++;
                        break;
                    case "Missing Tab":
                    	wrongTab++;
                    	break;
                }
            }

            if (extraWord > 0) summaryPara.createRun().setText("Extra Word: " + extraWord);
            if (extraWord > 0) summaryPara.createRun().addBreak();
            if (missingWord > 0) summaryPara.createRun().setText("Missing Word: " + missingWord);
            if (missingWord > 0) summaryPara.createRun().addBreak();
            if (wrongWord > 0) summaryPara.createRun().setText("Wrong Word: " + wrongWord);
            if (wrongWord > 0) summaryPara.createRun().addBreak();
            if (wrongTab > 0) summaryPara.createRun().setText("Missing Tab" + wrongTab);
            if (wrongTab > 0) summaryPara.createRun().addBreak();


            int groupedMistakes = 0;
            int groupedMistakes1=0;
            for (String key : Arrays.asList("Extra Tab", "Extra Space")) {
                int count = mistakeCounts.getOrDefault(key, 0);
                if (key.equals("Extra Space")) {
                    groupedMistakes += count;
                } else if (count >= 3) {
                    groupedMistakes++;
                }
            }
            groupedMistakes += mistakeCounts.getOrDefault("Extra Enter", 0);
//             groupedMistakes1 += mistakeCounts.getOrDefault("Missing Tab", 0);
            int totalMistakes = wordMistakes.size() + groupedMistakes+groupedMistakes1;
            int obtainedMarks = Math.max(0, 40 - totalMistakes);

            summaryPara.createRun().addBreak();
            summaryPara.createRun().setText("Total Mistakes: " + totalMistakes);
            summaryPara.createRun().addBreak();
            summaryPara.createRun().setText("Final Marks: " + obtainedMarks + " / 40");

            // Save compared DOCX
            File comparedFile = new File(compareDir, studentFile.getName().replace(".docx", "_Compared.docx"));
            try (FileOutputStream out = new FileOutputStream(comparedFile)) {
                outputDoc.write(out);
            }

            // Excel output
            Row row = sheet.createRow(rowNum++);
            row.createCell(0).setCellValue(studentFile.getName());
            row.createCell(1).setCellValue(totalMistakes);
            row.createCell(2).setCellValue(obtainedMarks);
            row.createCell(3).setCellValue(mistakeCounts.getOrDefault("Extra Space", 0));
            row.createCell(4).setCellValue(mistakeCounts.getOrDefault("Extra Tab", 0));
            row.createCell(5).setCellValue(mistakeCounts.getOrDefault("Extra Enter", 0));
            row.createCell(6).setCellValue(extraWord);
            row.createCell(7).setCellValue(missingWord);
            row.createCell(8).setCellValue(wrongWord);

            StringBuilder summary = new StringBuilder();
            for (String[] m : wordMistakes) {
                summary.append(m[1]).append(" - ").append(m[2]).append(" | ");
            }
            row.createCell(9).setCellValue(summary.toString().trim());
        }

        for (int i = 0; i < columns.length; i++) sheet.autoSizeColumn(i);
        try (FileOutputStream out = new FileOutputStream(new File(folder.getParent(), "Comparison_Result.xlsx"))) {
            workbook.write(out);
        }
        workbook.close();
        System.out.println("✅ All students processed.");
    }
    

    private static void addTextRun(XWPFParagraph para, String text, boolean highlight) {
        XWPFRun run = para.createRun();
        run.setText(text);
        if (highlight) {
            CTRPr rpr = run.getCTR().isSetRPr() ? run.getCTR().getRPr() : run.getCTR().addNewRPr();
            CTShd shd = rpr.addNewShd();
            shd.setVal(STShd.CLEAR);
            shd.setColor("auto");
            shd.setFill("FF0000");
        }
    }

    private static void incrementMistake(String type, Map<String, Integer> counter) {
        counter.put(type, counter.getOrDefault(type, 0) + 1);
    }

    private static void compareWords(String model, String student, XWPFParagraph para,
                                     Map<String, Integer> mistakeCounts,
                                     List<String[]> wordMistakes, String fileName) {
    	
//   testtttt*********************************
    	int modelTabs   = countMatches(model,   "\t");
        int studentTabs = countMatches(student, "\t");
        if (studentTabs < modelTabs) {
            int missing = modelTabs - studentTabs;
            for (int k = 0; k < missing; k++) {
                incrementMistake("Missing Tab", mistakeCounts);
//         /       wordMistakes.add(new String[]{fileName, "Missing Tab", });
                XWPFRun run = para.createRun();
                run.setText("[Missing Tab]");
                run.setColor("FF0000");
            }
        }    	
  //    	 model   = model.replaceAll(",\\s+", ",");
//    	    student = student.replaceAll(",\\s+", ",");
    	    	model   = model.trim().replaceAll("\\s+([,;:!?])", "$1");
        student = student.trim().replaceAll("\\s+([,;:!?])", "$1");        
 //        model   = model.replaceAll("(['‘’])\\s+", "$1");
//        student = student.replaceAll("(['‘’])\\s+", "$1");     // to remove space after punctuation       	//    	
//        student = student.replaceAll("([.,;:!?])(\\S)", "$1 $2"); // to add space after , .        
        model = model.replaceAll("\\.(?!\\s|$)", ". ");   // add space after . if not followed by space or end
        student = student.replaceAll("\\.(?!\\s|$)", ". ");    	
        String[] modelWords = model.trim().split("\\s+");
        String[] studentWords = student.trim().split("\\s+");
        
        int i = 0, j = 0;

        while (i < modelWords.length || j < studentWords.length) {
//            String refWord = i < modelWords.length ? modelWords[i] : null;
//            String stuWord = j < studentWords.length ? studentWords[j] : null;
        	String refWordRaw = i < modelWords.length ? modelWords[i] : null;
        	String stuWordRaw = j < studentWords.length ? studentWords[j] : null;

        	String refWord = normalizeWord(refWordRaw);
        	String stuWord = normalizeWord(stuWordRaw);

            XWPFRun run = para.createRun();
            
//            if (stuWordRaw != null && refWordRaw != null) {
//                boolean studentHasEnter = stuWordRaw.contains("\\n") || stuWordRaw.contains("\\r");
//                boolean modelHasEnter   = refWordRaw.contains("\\n") || refWordRaw.contains("\\r");
//
//                if (studentHasEnter && !modelHasEnter) {
//                    stuWordRaw = stuWordRaw.replaceAll("[\\r\\n]+", " ");
//                    stuWord = normalizeWord(stuWordRaw); // re-normalize after replacing
//                }
//            }

            if (refWord != null && stuWord != null && refWord.equals(stuWord)) {
                run.setText(stuWord + " ");
                i++; j++;
            }
                       
            // testing from   *********************************************       
       
//            else if (refWord != null && stuWord != null && !refWord.equals(stuWord))
// {
//            	highlight(run, "[" + refWord + "]");
//                wordMistakes.add(new String[]{fileName, "Missing Word", refWord});
//            }
             
            // testing to 
            else if (refWord != null && stuWord != null) {
                if (j + 1 < studentWords.length && refWord.equals(studentWords[j + 1])) {
                    highlight(run, "[" + stuWord + "]");
                    wordMistakes.add(new String[]{fileName, "Extra Word", stuWord});
                    j++;
                }        
                else if (i + 1 < modelWords.length&& modelWords[i + 1].equals(stuWord)) {
                    highlight(run, "[" + refWord + "]");
                    wordMistakes.add(new String[]{fileName, "Missing Word", refWord});
                    i++;
                } 
                
                
//               testing  
                
                
//                
                
                else {
                    highlight(run, "[" + stuWord + "/" + refWord + "]");
                    wordMistakes.add(new String[]{fileName, "Wrong Word", stuWord + "/" + refWord});
                    i++; j++;
                }
            } else if (stuWord != null) {
                highlight(run, "[" + stuWord + "]");
                wordMistakes.add(new String[]{fileName, "Extra Word", stuWord});
                j++;
            } else if (refWord != null) {
                highlight(run, "[" + refWord + "]");
                wordMistakes.add(new String[]{fileName, "Missing Word", refWord});
                i++;  
            }
            
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
                for (XWPFRun run : p.getRuns()) {
                    fullText.append(run.toString());
                }
                paras.add(fullText.toString());
            }
        }
        return paras;
    }
    private static String normalizeWord(String word) {
        if (word == null) return null;
        return word
            .replace("’", "'")  // curly apostrophe to straight
            .replace("‘", "'")  // opening curly single quote
        .replace("‘", "'")
        .replace("‘", "'");
        
        
//            .replace("“", "\"") // curly double quote
//            .replace("”", "\"") // closing curly double quote
//            .replace("´", "'")  // acute accent often used as apostrophe
//            .replace("`", "'"); // grave accent
        
//        String w = word.replace("’", "'")
//                .replace("‘", "'");
// // strip a trailing full‑stop if present:
// w = w.replaceAll("\\.$", "");
// return w;
    }


    private static int countMatches(String text, String regex) {
        return (int) Pattern.compile(regex).matcher(text).results().count();
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
    }
