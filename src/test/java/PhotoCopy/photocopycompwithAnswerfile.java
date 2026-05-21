package PhotoCopy;

import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

import java.io.*;
import java.util.regex.*;

public class photocopycompwithAnswerfile {

    // ✅ Read Word file text
    public static String readWordFile(File file) throws IOException {
        FileInputStream fis = new FileInputStream(file);
        XWPFDocument document = new XWPFDocument(fis);
        StringBuilder sb = new StringBuilder();
        for (XWPFParagraph para : document.getParagraphs()) {
            sb.append(para.getText()).append("\n");
        }
        document.close();
        fis.close();
        return sb.toString().trim();
    }

    // ✅ Read PDF file text
    public static String readPdfFile(File file) throws IOException {
        PDDocument document = PDDocument.load(file);
        PDFTextStripper stripper = new PDFTextStripper();
        String text = stripper.getText(document);
        document.close();
        return text.trim();
    }

    // ✅ Extract valid 10-digit seat number from filename (after SpeedAnswer_)
    public static String extractSeatNumber(String fileName) {
        String nameWithoutExt = fileName.replaceAll("\\.(docx|pdf)$", "");
        Matcher matcher = Pattern.compile("SpeedAnswer_(\\d{10})").matcher(nameWithoutExt);
        if (matcher.find()) {
            String seatNumber = matcher.group(1);
            String fifthSixth = seatNumber.substring(4, 6);
            if (fifthSixth.equals("15") || fifthSixth.equals("16")) {
                return seatNumber;
            } else {
                System.out.println("⚠ Skipped invalid seat number (5th–6th not 15/16): " + seatNumber + " in " + fileName);
            }
        }
        return null;
    }

    // ✅ Compare text ignoring case and extra spaces
    public static boolean compareText(String wordText, String pdfText) {
        wordText = wordText.replaceAll("\\s+", " ").trim().toLowerCase();
        pdfText = pdfText.replaceAll("\\s+", " ").trim().toLowerCase();
        return wordText.equals(pdfText);
    }

    // ✅ Find matching PDF file with same filename before extension
    private static File findMatchingPdf(File pdfFolder, File wordFile) {
        String baseName = wordFile.getName().replaceAll("(?i)\\.docx$", ""); // remove .docx or .DOCX
        for (File pdf : pdfFolder.listFiles()) {
            if (!pdf.isFile() || !pdf.getName().toLowerCase().endsWith(".pdf")) continue;
            String pdfBase = pdf.getName().replaceAll("(?i)\\.pdf$", "");
            if (pdfBase.equals(baseName)) {
                return pdf; // ✅ exact filename match
            }
        }
        return null; // ❌ not found
    }

    public static void main(String[] args) throws Exception {
        WebDriver driver = new ChromeDriver();
        driver.manage().window().maximize();

        File wordFolder = new File("F:\\GCC  TBC December 2025\\PHOTO COPY\\SpeedPhotoCopy");
        File pdfFolder = new File("F:\\GCC  TBC December 2025\\PHOTO COPY\\Photocopy");

        if (!wordFolder.exists() || !pdfFolder.exists()) {
            System.out.println("❌ One of the folders does not exist!");
            return;
        }

        // --- Create Excel workbook ---
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Comparison Results");

        // Header row
        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue("Seat Number");
        header.createCell(1).setCellValue("Word File");
        header.createCell(2).setCellValue("PDF File");
        header.createCell(3).setCellValue("Status");

        int rowNum = 1;

        // --- Process all Word files ---
        for (File wordFile : wordFolder.listFiles()) {
            if (wordFile.isFile() && wordFile.getName().toLowerCase().endsWith(".docx")) {

                String seatNumber = extractSeatNumber(wordFile.getName());
                if (seatNumber == null) continue;

                File pdfFile = findMatchingPdf(pdfFolder, wordFile);
                String status;
                String pdfFileName = (pdfFile != null) ? pdfFile.getName() : "PDF Not Found";

                if (pdfFile != null) {
                    String wordText = readWordFile(wordFile);
                    String pdfText = readPdfFile(pdfFile);
                    status = compareText(wordText, pdfText) ? "MATCH" : "MISMATCH";
                } else {
                    status = "PDF Not Found";
                }

                // --- Write result to Excel ---
                Row row = sheet.createRow(rowNum++);
                row.createCell(0).setCellValue(seatNumber);
                row.createCell(1).setCellValue(wordFile.getName());
                row.createCell(2).setCellValue(pdfFileName);
                row.createCell(3).setCellValue(status);

                // --- Console output ---
                System.out.println(status + " → " + wordFile.getName());
            }
        }

        // --- Autosize columns ---
        for (int i = 0; i < 4; i++) sheet.autoSizeColumn(i);

        // --- Save Excel file ---
        FileOutputStream fileOut = new FileOutputStream("F:\\GCC  TBC December 2025\\PHOTO COPY\\ComparisonResults.xlsx");
        workbook.write(fileOut);
        fileOut.close();
        workbook.close();

        driver.quit();
        System.out.println("✅ Excel report generated successfully!");
    }
}
