package BookmarkTesting_letter;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.io.*;
import java.util.*;

public class BookmarkParagraphExcelWriter {

    public static void main(String[] args) {
        String folderPath = "C:\\Users\\User\\Desktop\\bookmaek";  // <-- change this path // <-- change this path
        String outputPath = "C:\\Users\\User\\Desktop\\bookmaek\\Bookmark_Report.xlsx"; // <-- change this

        File folder = new File(folderPath);
        File[] files = folder.listFiles((dir, name) -> name.toLowerCase().endsWith(".docx"));

        if (files == null || files.length == 0) {
            System.out.println("❌ No .docx files found in folder.");
            return;
        }

        // Create Excel workbook & sheet
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Bookmark Report");

        // Create header row
        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue("File Name");
        header.createCell(1).setCellValue("Bookmark Name");
        header.createCell(2).setCellValue("Paragraphs Between");
        header.createCell(3).setCellValue("Body Count");
        header.createCell(4).setCellValue("Status");

        int rowIndex = 1; // Start writing data from row 2

        for (File file : files) {
            try (FileInputStream fis = new FileInputStream(file);
                 XWPFDocument document = new XWPFDocument(fis)) {

                int startIndex = -1;
                int endIndex = -1;
                int bodyBookmarkCount = 0;
                List<String> allBookmarks = new ArrayList<>();

                List<XWPFParagraph> paragraphs = document.getParagraphs();

                for (int i = 0; i < paragraphs.size(); i++) {
                    XWPFParagraph para = paragraphs.get(i);
                    CTP ctp = para.getCTP();

                    for (CTBookmark bookmark : ctp.getBookmarkStartList()) {
                        String name = bookmark.getName();
                        allBookmarks.add(name);

                        String lower = name.toLowerCase();

                        if (lower.contains("salut")) {
                            startIndex = i;
                        } else if (lower.contains("sign")) {
                            endIndex = i;
                        } else if (lower.contains("body")) {
                            bodyBookmarkCount++;
                        }
                    }
                }

                int paragraphBetweenCount = 0;
                String status = "N/A";

                if (startIndex != -1 && endIndex != -1) {
                    for (int i = startIndex + 1; i < endIndex; i++) {
                        XWPFParagraph para = paragraphs.get(i);
                        if (!para.getText().trim().isEmpty()) {
                            paragraphBetweenCount++;
                        }
                    }

                    status = (paragraphBetweenCount == bodyBookmarkCount) ? "Match" : "Not Match";
                } else {
                    status = "Missing salution/sign";
                }

                // Write each bookmark name to Excel
                for (String b : allBookmarks) {
                    Row row = sheet.createRow(rowIndex++);
                    row.createCell(0).setCellValue(file.getName());
                    row.createCell(1).setCellValue(b);
                    row.createCell(2).setCellValue(paragraphBetweenCount);
                    row.createCell(3).setCellValue(bodyBookmarkCount);
                    row.createCell(4).setCellValue(status);
                }

                System.out.println("✅ Processed: " + file.getName());

            } catch (Exception e) {
                System.out.println("❌ Error processing " + file.getName() + ": " + e.getMessage());
            }
        }

        // Autosize columns
        for (int i = 0; i < 5; i++) {
            sheet.autoSizeColumn(i);
        }

        // Save Excel file
        try (FileOutputStream fos = new FileOutputStream(outputPath)) {
            workbook.write(fos);
            workbook.close();
            System.out.println("\n📘 Excel report created at: " + outputPath);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
