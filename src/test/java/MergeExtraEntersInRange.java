import org.apache.poi.xwpf.usermodel.*;
import java.io.*;
import java.util.*;

public class MergeExtraEntersInRange {

    public static void main(String[] args) throws Exception {

        // === HARD-CODED PATHS ===
        File modelFile = new File("F:\\desktop backup 5 july 25\\GCC TBC MAY 2025 AUTO\\allocation\\New folder\\Model Anwers\\Eng40 B Ltr Question5.docx");
        File studentFile = new File("F:\\desktop backup 5 july 25\\GCC TBC MAY 2025 AUTO\\New folder\\New folder\\New folder\\letterAnswer_3102160016_204_6-19-2025_4-22-29.docx");
        // ========================

        if (!modelFile.exists() || !studentFile.exists()) {
            System.err.println("Model or Student file not found.");
            return;
        }

        List<String> modelParas = readParagraphTexts(modelFile);
        int modelStart = findParagraphIndexContainingTextList(modelParas, "reference");
        int modelEnd = findParagraphIndexStartingWithTextList(modelParas, "encl");

        if (modelStart < 0 || modelEnd < 0 || modelEnd <= modelStart) {
            System.err.println("Could not find 'reference' to 'encl' range in model.");
            return;
        }

        // Get model no-blank pairs in range
        Set<Integer> noBlankPairs = new HashSet<>();
        for (int i = modelStart; i < modelEnd; i++) {
            String cur = modelParas.get(i).trim();
            String next = modelParas.get(i + 1).trim();
            if (!cur.isEmpty() && !next.isEmpty()) {
                noBlankPairs.add(i);
            }
        }

        // Open student file for editing
        try (FileInputStream fis = new FileInputStream(studentFile);
             XWPFDocument studDoc = new XWPFDocument(fis)) {

            List<XWPFParagraph> paras = studDoc.getParagraphs();

            // Find matching start and end in student using keywords only
            int studStart = findParagraphIndexContainingParaList(paras, "reference");
            int studEnd = findParagraphIndexStartingWithParaList(paras, "encl");

            if (studStart < 0 || studEnd < 0 || studEnd <= studStart) {
                System.err.println("Could not find matching range in student.");
                return;
            }

            boolean changed = false;
            int i = studStart;
            while (i < studEnd) {
                if (noBlankPairs.contains(modelStart + (i - studStart))) {
                    // This pair should have no blank in model
                    int j = i + 1;
                    int emptyCount = 0;
                    while (j < studEnd && paras.get(j).getText().trim().isEmpty()) {
                        emptyCount++;
                        j++;
                    }
                    if (emptyCount >= 1 && j < studEnd) {
                        String merged = paras.get(i).getText().trim() + " " + paras.get(j).getText().trim();
                        clearParagraph(paras.get(i));
                        paras.get(i).createRun().setText(merged);
                        // Remove blanks + merged para
                        for (int r = 0; r < (j - i); r++) {
                            studDoc.removeBodyElement(studDoc.getPosOfParagraph(paras.get(i + 1)));
                        }
                        changed = true;
                        continue; // recheck same i
                    }
                }
                i++;
            }

            // Save as new file in same folder
            String studentName = studentFile.getName();
            String newName = studentName.replaceAll("\\.docx$", "") + "_fixed.docx";
            File newFile = new File(studentFile.getParentFile(), newName);

            try (FileOutputStream out = new FileOutputStream(newFile)) {
                studDoc.write(out);
            }

            if (changed) {
                System.out.println("Modified file created: " + newFile.getAbsolutePath());
            } else {
                System.out.println("No changes needed. Copy created anyway: " + newFile.getAbsolutePath());
            }
        }
    }

    // ---------- Utility Methods ----------

    private static List<String> readParagraphTexts(File file) throws IOException {
        try (FileInputStream fis = new FileInputStream(file);
             XWPFDocument doc = new XWPFDocument(fis)) {
            List<String> list = new ArrayList<>();
            for (XWPFParagraph p : doc.getParagraphs()) {
                list.add(p.getText() == null ? "" : p.getText());
            }
            return list;
        }
    }

    // For List<String> paragraphs
    private static int findParagraphIndexContainingTextList(List<String> paras, String keyword) {
        String kw = keyword.toLowerCase();
        for (int i = 0; i < paras.size(); i++) {
            if (paras.get(i).toLowerCase().contains(kw)) return i;
        }
        return -1;
    }

    private static int findParagraphIndexStartingWithTextList(List<String> paras, String keyword) {
        String kw = keyword.toLowerCase();
        for (int i = 0; i < paras.size(); i++) {
            if (paras.get(i).trim().toLowerCase().startsWith(kw)) return i;
        }
        return -1;
    }

    // For List<XWPFParagraph> paragraphs
    private static int findParagraphIndexContainingParaList(List<XWPFParagraph> paras, String keyword) {
        String kw = keyword.toLowerCase();
        for (int i = 0; i < paras.size(); i++) {
            if (paras.get(i).getText().toLowerCase().contains(kw)) return i;
        }
        return -1;
    }

    private static int findParagraphIndexStartingWithParaList(List<XWPFParagraph> paras, String keyword) {
        String kw = keyword.toLowerCase();
        for (int i = 0; i < paras.size(); i++) {
            if (paras.get(i).getText().trim().toLowerCase().startsWith(kw)) return i;
        }
        return -1;
    }

    private static void clearParagraph(XWPFParagraph para) {
        int runs = para.getRuns() == null ? 0 : para.getRuns().size();
        for (int i = runs - 1; i >= 0; i--) {
            para.removeRun(i);
        }
    }
}
