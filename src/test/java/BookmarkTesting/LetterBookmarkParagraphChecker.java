package BookmarkTesting;

import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.List;

public class LetterBookmarkParagraphChecker {

    public static void main(String[] args) {
        String filePath = "C:\\Users\\User\\Desktop\\bookmaek\\Eng30 Ltr 1.docx"; // <-- Change this path

        try (FileInputStream fis = new FileInputStream(filePath);
             XWPFDocument document = new XWPFDocument(fis)) {

            int startIndex = -1;
            int endIndex = -1;
            int bodyBookmarkCount = 0;

            List<XWPFParagraph> paragraphs = document.getParagraphs();

            for (int i = 0; i < paragraphs.size(); i++) {
                XWPFParagraph para = paragraphs.get(i);
                CTP ctp = para.getCTP();

                for (CTBookmark bookmark : ctp.getBookmarkStartList()) {
                    String name = bookmark.getName().toLowerCase(); // make it case-insensitive

                    // Find start and end bookmarks dynamically
                    if (name.contains("salu")) {
                        startIndex = i;
                    } else if (name.contains("sign")) {
                        endIndex = i;
                    } 
                    // Count Body/body bookmarks
                    else if (name.contains("body")) {
                        bodyBookmarkCount++;
                    }
                }
            }

            // Check if we found the bookmarks
            if (startIndex == -1 || endIndex == -1) {
                System.out.println("❌ Could not find bookmarks containing 'salution' or 'sign'.");
                return;
            }

            // Count paragraphs between salution and sign
            int paragraphBetweenCount = 0;
            for (int i = startIndex + 1; i < endIndex; i++) {
                XWPFParagraph para = paragraphs.get(i);
                if (!para.getText().trim().isEmpty()) {
                    paragraphBetweenCount++;
                }
            }

            // Display results
            System.out.println("📘 Bookmark Range: salution ➜ sign");
            System.out.println("🧾 Paragraphs Between: " + paragraphBetweenCount);
            System.out.println("🏷️ Body/Body_ Bookmark Count: " + bodyBookmarkCount);

            // Compare counts
            if (paragraphBetweenCount == bodyBookmarkCount) {
                System.out.println("✅ Paragraph count matches number of Body bookmarks!");
            } else {
                System.out.println("⚠️ Mismatch: Paragraph count and Body bookmarks differ.");
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
