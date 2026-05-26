package BookmarkTesting_letter;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.*;
import java.util.*;

public class MarathiHiddenCharFinder {

    public static void main(String[] args) throws IOException {

        String inputFile = "C:\\Users\\User\\Desktop\\EmailMarking Oct\\Marks\\EmailMarathi30OCTMarks.xlsx";
        String outputFile = "C:\\Users\\User\\Desktop\\EmailMarking Oct\\Marks\\HiddenUnicodeReport1.xlsx";

        Workbook wb = new XSSFWorkbook(new FileInputStream(inputFile));
        Sheet sh = wb.getSheetAt(0);

        Workbook outWb = new XSSFWorkbook();
        Sheet out = outWb.createSheet("HiddenUnicode");

        // Add a new header column "Actual Word"
        Row head = out.createRow(0);
        head.createCell(0).setCellValue("Seat No");
        head.createCell(1).setCellValue("Word 1");
        head.createCell(2).setCellValue("Word 2");
        head.createCell(3).setCellValue("Extra Codepoints");
        head.createCell(4).setCellValue("Actual Word"); // New column

        int rowNum = 1;

        for (int i = 1; i <= sh.getLastRowNum(); i++) {
            Row r = sh.getRow(i);
            if (r == null) continue;

            Cell seatCell = r.getCell(0);   // Seat No (A)
            Cell wordCell = r.getCell(8);   // Marathi text (I)
            if (seatCell == null || wordCell == null) continue;

            String seat = seatCell.toString().trim();
            String text = wordCell.toString().trim();
            if (!text.contains("|")) continue;

            String[] parts = text.split("\\|");
            if (parts.length < 2) continue;

            String w1 = parts[0].trim();
            String w2 = parts[1].trim();

            // find invisible codepoints present only in w2 (or w1)
            Set<String> diffs = hiddenCodes(w1, w2);
            if (!diffs.isEmpty()) {
                Row or = out.createRow(rowNum++);
                or.createCell(0).setCellValue(seat);
                or.createCell(1).setCellValue(w1);
                or.createCell(2).setCellValue(w2);
                or.createCell(3).setCellValue(String.join(", ", diffs));
                or.createCell(4).setCellValue(text); // Write actual original word
            }
        }

        try (FileOutputStream fos = new FileOutputStream(outputFile)) {
            outWb.write(fos);
        }
        wb.close();
        outWb.close();

        System.out.println("✅ Report written to " + outputFile);
    }

    // return any hidden/invisible Unicode codes present in one string but not the other
    private static Set<String> hiddenCodes(String w1, String w2) {
        Set<Integer> set1 = w1.codePoints().collect(HashSet::new, HashSet::add, HashSet::addAll);
        Set<Integer> set2 = w2.codePoints().collect(HashSet::new, HashSet::add, HashSet::addAll);
        Set<Integer> diff = new HashSet<>(set1);
        diff.addAll(set2);
        set1.retainAll(set2);
        diff.removeAll(set1);

        Set<String> hidden = new LinkedHashSet<>();
        for (int cp : diff) {
            if (cp == 0x200C || cp == 0x200D || cp == 0xFEFF || cp == 0x2060)
                hidden.add(String.format("U+%04X", cp));
        }
        return hidden;
    }
}
