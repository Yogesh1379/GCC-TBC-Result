package BookmarkTesting_letter;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.io.*;
import java.util.*;

public class LetterBookmarkParagraphFolderChecker {

    public static void main(String[] args) {
        // 📂 Folder containing all Word files
        String folderPath = "C:\\Users\\User\\Desktop\\bookmaek";
        File folder = new File(folderPath);
        File[] files = folder.listFiles((dir, name) -> name.toLowerCase().endsWith(".docx"));

        if (files == null || files.length == 0) {
            System.out.println("❌ No .docx files found in folder.");
            return;
        }

        List<String> mismatchFiles = new ArrayList<>();

        for (File file : files) {
            try (FileInputStream fis = new FileInputStream(file);
                 XWPFDocument document = new XWPFDocument(fis)) {

                int startIndex = -1;
                int endIndex = -1;
                int bodyBookmarkCount = 0;

                List<XWPFParagraph> paragraphs = document.getParagraphs();

                for (int i = 0; i < paragraphs.size(); i++) {
                    XWPFParagraph para = paragraphs.get(i);
                    CTP ctp = para.getCTP();

                    for (CTBookmark bookmark : ctp.getBookmarkStartList()) {
                        String name = bookmark.getName().toLowerCase();

                        if (name.contains("salution")) {
                            startIndex = i;
                        } else if (name.contains("sign")) {
                            endIndex = i;
                        } else if (name.contains("body")) {
                            bodyBookmarkCount++;
                        }
                    }
                }

                if (startIndex == -1 || endIndex == -1) {
                    System.out.println("⚠️ Skipping " + file.getName() +
                            " (missing 'salution' or 'sign' bookmarks)");
                    continue;
                }

                // Count paragraphs between start and end
                int paragraphBetweenCount = 0;
                for (int i = startIndex + 1; i < endIndex; i++) {
                    XWPFParagraph para = paragraphs.get(i);
                    if (!para.getText().trim().isEmpty()) {
                        paragraphBetweenCount++;
                    }
                }

                // Compare and show result
                System.out.println("🧾 File: " + file.getName());
                System.out.println("   Paragraphs Between: " + paragraphBetweenCount);
                System.out.println("   Body Bookmark Count: " + bodyBookmarkCount);

                if (paragraphBetweenCount == bodyBookmarkCount) {
                    System.out.println("   ✅ Match\n");
                } else {
                    System.out.println("   ❌ Mismatch\n");
                    mismatchFiles.add(file.getName());
                }

            } catch (Exception e) {
                System.out.println("❌ Error reading " + file.getName() + ": " + e.getMessage());
            }
        }

        // Summary
        if (mismatchFiles.isEmpty()) {
            System.out.println("🎉 All files matched successfully!");
        } else {
            System.out.println("\n⚠️ Files with mismatched paragraph count and Body bookmarks:");
            for (String f : mismatchFiles) {
                System.out.println(" - " + f);
            }
        }
    }
}
