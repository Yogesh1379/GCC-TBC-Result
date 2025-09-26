import org.apache.poi.xwpf.usermodel.*;

import java.io.FileInputStream;
import java.util.List;

public class DetectExtraEnter {
    public static void main(String[] args) throws Exception {
        XWPFDocument modelDoc = new XWPFDocument(new FileInputStream("F:\\desktop backup 5 july 25\\GCC TBC MAY 2025 AUTO\\allocation\\New folder\\Model Anwers\\Eng40 B Ltr Question5.docx"));
        XWPFDocument studentDoc = new XWPFDocument(new FileInputStream("F:\\desktop backup 5 july 25\\GCC TBC MAY 2025 AUTO\\New folder\\New folder\\New folder\\letterAnswer_3102160006_204_19-06-2025_16-21-40.docx"));

        List<XWPFParagraph> modelParas = modelDoc.getParagraphs();
        List<XWPFParagraph> studentParas = studentDoc.getParagraphs();

        int modelIndex = 0;
        int studentIndex = 0;

        while (modelIndex < modelParas.size() && studentIndex < studentParas.size()) {
            String modelText = normalize(modelParas.get(modelIndex).getText());
            String studentText = normalize(studentParas.get(studentIndex).getText());

            if (modelText.startsWith(studentText) && !modelText.equals(studentText)) {
                // Student pressed Enter early → extra paragraph
                System.out.println("Extra Enter detected at student paragraph: " + (studentIndex + 1));
                studentIndex++; // move to next student paragraph
                continue;
            }

            modelIndex++;
            studentIndex++;
        }
    }

    private static String normalize(String text) {
        return text == null ? "" : text.replaceAll("\\s+", " ").trim();
    }
}
