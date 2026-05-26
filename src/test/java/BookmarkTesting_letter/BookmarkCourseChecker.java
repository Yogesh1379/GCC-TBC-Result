package BookmarkTesting_letter;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.io.*;
import java.util.*;

public class BookmarkCourseChecker {

    public static void main(String[] args) {
        String folderPath = "C:\\Users\\User\\Desktop\\bookmaek";
        String outputExcel = "C:\\Users\\User\\Desktop\\bookmaek\\Bookmark_Report.xlsx";

        // --- Step 1: Expected bookmark sets ---
        Map<String, List<String>> expectedBookmarks = new HashMap<>();
        expectedBookmarks.put("Eng30_Letter4Para", Arrays.asList(
                "heading_10","ref_05","address_10","subject_10","salutation_05",
                "body1_05","body2_05","body3_05","body4_05","sign_10","enclosure_05"
        ));
        expectedBookmarks.put("Eng30_Letter3Para", Arrays.asList(
                "heading_10","ref_05","address_10","subject_10","salutation_05",
                "body1_65","body2_65","body3_65","sign_10","enclosure_05"
        ));
        expectedBookmarks.put("Eng30_Letter2Para", Arrays.asList(
                "heading_10","ref_05","address_10","subject_10","salutation_05",
                "body1_10","body2_10","sign_10","enclosure_05"
        ));
        expectedBookmarks.put("Eng40_BLetter3Para", Arrays.asList(
                "Bheading_10","Bref_05","Baddress_20","Bsubject_05","Bsalution_05",
                "Bbody1_65","Bbody2_65","Bbody3_65","Bsign_05","Bend_05"
        ));
        expectedBookmarks.put("Eng40_BLetter4Para", Arrays.asList(
                "Bheading_10","Bref_05","Baddress_20","Bsubject_05","Bsalution_05",
                "Bbody1_05","Bbody2_05","Bbody3_05","Bbody4_05","Bsign_05","Bend_05"
        ));
        expectedBookmarks.put("Eng40_RLetter", Arrays.asList(
                "heading_05","address_20","subject_05","salutation_05",
                "body1_05","body2_05","body3_10","body4_10","sign_05","enclosure_05"
        ));

        // --- Step 2: Prepare Excel workbook ---
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Bookmark Validation");
        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue("File Name");
        header.createCell(1).setCellValue("Course");
        header.createCell(2).setCellValue("Expected Bookmark");
        header.createCell(3).setCellValue("Found");
        header.createCell(4).setCellValue("Status");

        int rowIndex = 1;

        File folder = new File(folderPath);
        File[] files = folder.listFiles((dir, name) -> name.toLowerCase().endsWith(".docx"));
        if (files == null) {
            System.out.println("No .docx files found.");
            return;
        }

        // --- Step 3: Process each file ---
        for (File file : files) {
            String name = file.getName().toLowerCase();
            String courseKey = "";

            if (name.contains("30") && name.contains("letter4")) courseKey = "Eng30_Letter4Para";
            else if (name.contains("30") && name.contains("letter3")) courseKey = "Eng30_Letter3Para";
            else if (name.contains("30") && name.contains("letter2")) courseKey = "Eng30_Letter2Para";
            else if (name.contains("40") && name.contains("b") && name.contains("letter4")) courseKey = "Eng40_BLetter4Para";
            else if (name.contains("40") && name.contains("b") && name.contains("letter3")) courseKey = "Eng40_BLetter3Para";
            else if (name.contains("40") && name.contains("rletter")) courseKey = "Eng40_RLetter";

            if (courseKey.isEmpty()) {
                System.out.println("⚠️ Unknown course type for file: " + file.getName());
                continue;
            }

            List<String> expected = expectedBookmarks.get(courseKey);
            Set<String> found = new HashSet<>();

            try (FileInputStream fis = new FileInputStream(file);
                 XWPFDocument doc = new XWPFDocument(fis)) {

                for (XWPFParagraph p : doc.getParagraphs()) {
                    for (CTBookmark bm : p.getCTP().getBookmarkStartList()) {
                        String bmName = bm.getName();
                        if (!bmName.equalsIgnoreCase("_GoBack")) {
                            found.add(bmName);
                        }
                    }
                }

            } catch (Exception e) {
                System.out.println("❌ Error reading " + file.getName() + ": " + e.getMessage());
                continue;
            }

            // --- Step 4: Compare and write results ---
            for (String expectedName : expected) {
                Row row = sheet.createRow(rowIndex++);
                row.createCell(0).setCellValue(file.getName());
                row.createCell(1).setCellValue(courseKey);
                row.createCell(2).setCellValue(expectedName);
                row.createCell(3).setCellValue(found.contains(expectedName) ? "Yes" : "No");
                row.createCell(4).setCellValue(found.contains(expectedName) ? "Match" : "Missing");
            }
        }

        // --- Step 5: Save Excel ---
        for (int i = 0; i < 5; i++) sheet.autoSizeColumn(i);

        try (FileOutputStream fos = new FileOutputStream(outputExcel)) {
            workbook.write(fos);
            workbook.close();
            System.out.println("\n📘 Excel report saved at: " + outputExcel);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
