package BookmarkTesting;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;

public class EmailMarksCalculator30 {

    public static void main(String[] args) throws IOException {
        FileInputStream fis = new FileInputStream("C:\\Users\\User\\Desktop\\EmailMarking Oct\\eng3040emailoctexam.xlsx");
        Workbook workbook = new XSSFWorkbook(fis);
        Sheet sheet = workbook.getSheetAt(0);

        // Create output workbook
        Workbook outputWorkbook = new XSSFWorkbook();
        Sheet outputSheet = outputWorkbook.createSheet("Marks");

        // Header row
        Row headerRow = outputSheet.createRow(0);
        String[] headers = {"SeatNumber", "AttachmentMarks", "SubjectMarks", "ToSendMailMarks", "ContentMarks", "IsSendMarks", "TotalMarks","Batch", "WordMistakes"};
        for (int i = 0; i < headers.length; i++) {
            headerRow.createCell(i).setCellValue(headers[i]);
        }

        int outputRowNum = 1;
        for (Row row : sheet) {
            if (row.getRowNum() == 0) continue; // Skip header

            String Questionattachment1 = getCellValue(row.getCell(1));
            String questioncontent = getCellValue(row.getCell(2));
            String questionsubject = getCellValue(row.getCell(3));
            String questiontosend = getCellValue(row.getCell(4));

            String Attachment1 = getCellValue(row.getCell(5));
            String Contents = getCellValue(row.getCell(6));
            String Subject = getCellValue(row.getCell(7));
            String ToSendMail = getCellValue(row.getCell(8));
            String BatchName = getCellValue(row.getCell(21));
            String IsSendStr = getCellValue(row.getCell(9));
            int IsSend = IsSendStr.equals("1") ? 1 : 0;

            String SeatNumber = getCellValue(row.getCell(10));

            // Calculate content marks + mistakes
            ContentResult result = calculateContentMarks(questioncontent, Contents);

            // Marks calculation
            int attachmentMarks = Questionattachment1.equals(Attachment1) ? 1 : 0;
//            int subjectMarks = questionsubject.equals(Subject) ? 1 : 0;
            // Normalize spaces before comparing Subject
            String normalizedQuestionSubject = questionsubject.replaceAll("\\s+", " ").trim();
            String normalizedSubject = Subject.replaceAll("\\s+", " ").trim();
            int subjectMarks = normalizedQuestionSubject.equals(normalizedSubject) ? 1 : 0;

            int toSendMarks = questiontosend.equals(ToSendMail) ? 1 : 0;
            int isSendMarks = IsSend;
            double totalMarks = attachmentMarks + subjectMarks + toSendMarks + result.contentMarks + isSendMarks;

            // Write to output
            Row outRow = outputSheet.createRow(outputRowNum++);
            outRow.createCell(0).setCellValue(SeatNumber);
            outRow.createCell(1).setCellValue(attachmentMarks);
            outRow.createCell(2).setCellValue(subjectMarks);
            outRow.createCell(3).setCellValue(toSendMarks);
            outRow.createCell(4).setCellValue(result.contentMarks);
            outRow.createCell(5).setCellValue(isSendMarks);
            outRow.createCell(6).setCellValue(totalMarks);
            outRow.createCell(7).setCellValue(BatchName);
            outRow.createCell(8).setCellValue(result.wordMistakes); // ✅ Added mistakes text
        }

        // Write Excel
        FileOutputStream fos = new FileOutputStream("C:\\Users\\User\\Desktop\\EmailMarking Oct\\Marks\\EmailEnglish30OCTMarks.xlsx");
        outputWorkbook.write(fos);
        fos.close();
        workbook.close();
        outputWorkbook.close();

        System.out.println("Marks calculation completed successfully!");
    }

    // Helper for cell value
    private static String getCellValue(Cell cell) {
        if (cell == null) return "";
        if (cell.getCellType() == CellType.STRING) return cell.getStringCellValue().trim();
        else if (cell.getCellType() == CellType.NUMERIC) {
            if (DateUtil.isCellDateFormatted(cell)) {
                return cell.getDateCellValue().toString();
            } else {
                long value = (long) cell.getNumericCellValue();
                return String.valueOf(value);
            }
        } else return cell.toString().trim();
    }

    // ✅ New class to hold marks + word mistakes
    static class ContentResult {
        double contentMarks;
        String wordMistakes;
    }

    // ✅ Improved method that ignores tabs/spaces/enters and tracks word mistakes
//    private static ContentResult calculateContentMarks(String questionContent, String answerContent) {
//        ContentResult result = new ContentResult();
//
//        if (questionContent == null) questionContent = "";
//        if (answerContent == null) answerContent = "";
//
//        // Normalize whitespace
//        questionContent = questionContent.replaceAll("[\\t\\n\\r]+", " ").trim();
//        answerContent = answerContent.replaceAll("[\\t\\n\\r]+", " ").trim();
//
//        if (questionContent.isEmpty() && answerContent.isEmpty()) {
//            result.contentMarks = 0.0;
//            result.wordMistakes = "";
//            return result;
//        }
//
//        if (questionContent.equals(answerContent)) {
//            result.contentMarks = 1.0;
//            result.wordMistakes = "";
//            return result;
//        }
//
//        String[] questionWords = questionContent.split("\\s+");
//        String[] answerWords = answerContent.split("\\s+");
//
//        int mistakes = 0;
//        int len = Math.min(questionWords.length, answerWords.length);
//        StringBuilder wordMistakes = new StringBuilder();
//
//        for (int i = 0; i < len; i++) {
//            if (!questionWords[i].trim().equals(answerWords[i].trim())) {
//                mistakes++;
//                wordMistakes.append(questionWords[i]).append(" | ").append(answerWords[i]).append("\n");
//            }
//        }
//
//        // Extra/missing words
//        if (questionWords.length > answerWords.length) {
//            for (int i = len; i < questionWords.length; i++) {
//                mistakes++;
//                wordMistakes.append(questionWords[i]).append(" | (missing)\n");
//            }
//        } else if (answerWords.length > questionWords.length) {
//            for (int i = len; i < answerWords.length; i++) {
//                mistakes++;
//                wordMistakes.append("(extra) | ").append(answerWords[i]).append("\n");
//            }
//        }
//
//        // Marks rule
//        if (mistakes == 0) result.contentMarks = 1.0;
//        else if (mistakes == 1) result.contentMarks = 0.5;
//        else result.contentMarks = 0.0;
//
//        result.wordMistakes = wordMistakes.toString().trim();
//        return result;
//    }
    private static ContentResult calculateContentMarks(String questionContent, String answerContent) {
        ContentResult result = new ContentResult();

        if (questionContent == null) questionContent = "";
        if (answerContent == null) answerContent = "";

        // Normalize whitespace (tabs/newlines -> single space) and trim
        questionContent = questionContent.replaceAll("[\\t\\n\\r]+", " ").trim();
        answerContent = answerContent.replaceAll("[\\t\\n\\r]+", " ").trim();

        if (questionContent.isEmpty() && answerContent.isEmpty()) {
            result.contentMarks = 0.0;
            result.wordMistakes = "";
            return result;
        }

        if (questionContent.equalsIgnoreCase(answerContent)) {
            result.contentMarks = 1.0;
            result.wordMistakes = "";
            return result;
        }

        String[] questionWords = questionContent.split("\\s+");
        String[] answerWords = answerContent.split("\\s+");

        int mistakes = 0;
        StringBuilder wordMistakes = new StringBuilder();

        int i = 0; // index for questionWords
        int j = 0; // index for answerWords
        while (i < questionWords.length && j < answerWords.length) {
            String qw = questionWords[i].trim();
            String aw = answerWords[j].trim();

            if (qw.equalsIgnoreCase(aw)) {
                // exact match, advance both
                i++;
                j++;
            } else {
                // Try to detect an extra word in answer (answer has an inserted word)
                if (j + 1 < answerWords.length && qw.equalsIgnoreCase(answerWords[j + 1].trim())) {
                    mistakes++;
                    wordMistakes.append("(extra) | ").append(aw).append("\n");
                    j++; // consume extra word from answer
                }
                // Try to detect a missing word in answer (answer missing qw)
                else if (i + 1 < questionWords.length && questionWords[i + 1].trim().equalsIgnoreCase(aw)) {
                    mistakes++;
                    wordMistakes.append(qw).append(" | (missing)\n");
                    i++; // consume missing word from question
                }
                // General mismatch (substitution)
                else {
                    mistakes++;
                    wordMistakes.append(qw).append(" | ").append(aw).append("\n");
                    i++;
                    j++;
                }
            }
        }

        // Remaining question words => missing in answer
        while (i < questionWords.length) {
            mistakes++;
            wordMistakes.append(questionWords[i].trim()).append(" | (missing)\n");
            i++;
        }

        // Remaining answer words => extra in answer
        while (j < answerWords.length) {
            mistakes++;
            wordMistakes.append("(extra) | ").append(answerWords[j].trim()).append("\n");
            j++;
        }

        // Marks rule
        if (mistakes == 0) result.contentMarks = 1.0;
        else if (mistakes == 1) result.contentMarks = 0.5;
        else result.contentMarks = 0.0;

        result.wordMistakes = wordMistakes.toString().trim();
        return result;
    }

}
