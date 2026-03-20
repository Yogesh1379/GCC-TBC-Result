import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.List;

public class ReadBookmarksFromWord {
    public static void main(String[] args) {
        String filePath = "F:\\GCC-TBC October  repeaters\\Question files\\SUBJECTIE-\\ENG-40\\ENG 40-LETTER\\Business_Letter\\Eng40 B Ltr Answer\\Eng40 B Ltr Question1.docx"; // Change this path

        try (FileInputStream fis = new FileInputStream(filePath);
             XWPFDocument document = new XWPFDocument(fis)) {

            // Loop through all paragraphs in the document
            for (XWPFParagraph para : document.getParagraphs()) {

                // Access the underlying XML (CTP)
                CTP ctp = para.getCTP();

                // Get bookmark start elements
                List<CTBookmark> bookmarks = ctp.getBookmarkStartList();

                for (CTBookmark bookmark : bookmarks) {
                    System.out.println("📘 Bookmark Name: " + bookmark.getName());
                    System.out.println("📍 Bookmark ID: " + bookmark.getId());
                    System.out.println("📄 Paragraph Text: " + para.getText());
                    System.out.println("------------------------------------");
                }
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
