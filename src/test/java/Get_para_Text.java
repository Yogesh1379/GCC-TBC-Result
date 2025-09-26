import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class Get_para_Text {
         public static void main(String[] args) throws IOException {

            File modelfile = new File("F:\\desktop backup 5 july 25\\GCC TBC MAY 2025 AUTO\\allocation\\English_30_40_Quetion_paper\\All_Subjective_Eng30\\Eng30 Speed 1.docx");
           File studentfile= new File("F:\\desktop backup 5 july 25\\GCC TBC MAY 2025 AUTO\\marking\\New folder\\SpeedAnswer_1101150001_101_6-18-2025_9-17-10.docx");
            List<String> modelPara = extractFormattedParagraphs(modelfile);
            List<String> studentpara = extractFormattedParagraphs(studentfile);
             int minSize = Math.max(modelPara.size(), studentpara.size());
             for (int i = 0; i < minSize; i++) {
            	 String modelText = (i < modelPara.size()) ? modelPara.get(i).trim() : "";
            	    String studentText = (i < studentpara.size()) ? studentpara.get(i).trim() : "";
                 if (modelText.equals(studentText)) {
                     System.out.println("Paragraph " + (i + 1) + ": Match ✅");
                 } else if(!modelText.isBlank() && studentText.isBlank())
                 {
                     System.out.println("Paragraph "+(i+1)+ ": Missing ✅❌");
                     System.out.println("Model   : " + modelText);
                 }
                 else {
                     System.out.println("Paragraph " + (i + 1) + ": Mismatch ❌");
                     System.out.println("Model   : " + modelText);
                     System.out.println("Student : " + studentText);
                     List<String> missingWords = getStrings(modelText, studentText);
                     if (!missingWords.isEmpty()) {
                         System.out.println("Missing words from student❌✅: " + missingWords);
                     } else {
                         System.out.println("Words are similar but order or format may differ.");
                     }
                 }

             }
        }

    private static List<String> getStrings(String modelText, String studentText) {
        String[] modelWords = modelText.split("\\s+");
        String[] studentWords = studentText.split("\\s+");

        List<String> missingWords = new ArrayList<>();
        for (String word : modelWords) {
            boolean found = false;
            for (String sWord : studentWords) {
            	
                if (word.equals(sWord)) {
                    found = true;
                    break;
                }
            
            for(int i=0;i<modelWords.length;i++)
            {
            if (!word.isEmpty()&&sWord.isEmpty() ) {
                missingWords.add(word);
            }
        }}
        }
        return missingWords;
    }

    private static List<String> extractFormattedParagraphs(File file) throws IOException {
            @SuppressWarnings({ "rawtypes", "unchecked" })
			List<String> paras = new ArrayList();
            FileInputStream fis = new FileInputStream(file);
            @SuppressWarnings("resource")
			XWPFDocument doc = new XWPFDocument(fis);

            for(XWPFParagraph p : doc.getParagraphs()) {
                StringBuilder fullText = new StringBuilder();

                for(XWPFRun run : p.getRuns()) {
                    fullText.append(run.toString());
                }

                paras.add(fullText.toString());
            }

            return paras;
        }
    }


