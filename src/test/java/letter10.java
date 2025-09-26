import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import java.util.regex.Pattern;
import java.util.regex.Matcher;

import java.io.*;
import java.util.*;

public class letter10 {
    static File modelFile = null;
    static List<String> modelParas;

    @SuppressWarnings("resource")
    public static void main(String[] args) throws Exception {
        File folder = new File("F:\\desktop backup 5 july 25\\GCC TBC MAY 2025 AUTO\\New folder\\New folder\\New folder");
        File compareDir = new File(folder.getParentFile(), folder.getName() + "_Compared");
        if (!compareDir.exists()) compareDir.mkdirs();

        Workbook resultWb   = new XSSFWorkbook();
        Sheet    resultSheet = resultWb.createSheet("Results");
        String[] columns     = {
            "File Name", "Extra Word", "Missing Word", "Wrong Word",
            "Total Mistakes", "Final Marks", "Mistakes Summary"
        };
        Row header = resultSheet.createRow(0);
        for (int i = 0; i < columns.length; i++) {
            header.createCell(i).setCellValue(columns[i]);
        }

        File[] docxFiles = folder.listFiles((dir, name) ->
            name.toLowerCase().endsWith(".docx") && !name.startsWith("~$")
        );
        if (docxFiles == null) {
            System.out.println("❌ No .docx files found.");
            return;
        }

        int    rowNum     = 1;
        double MAX_MARKS  = 7.5;

        for (File studentFile : docxFiles) {
            // determine modelFile
            String[] parts = studentFile.getName().split("_");
            if (parts.length < 3) continue;
            String seatno    = parts[1],
                   batchname = parts[2];
            System.out.println(seatno);
            if (seatno.length() >= 10) {
                String course = seatno.substring(4, 6);
                int course1 = course.equals("15") ? 1 : course.equals("16") ? 2 : 0;
                try (FileInputStream fis = new FileInputStream(
                        "F:\\desktop backup 5 july 25\\GCC TBC MAY 2025 AUTO\\allocation\\BathwiseSubjective.xlsx"
                )) {
                    Workbook allocWb = new XSSFWorkbook(fis);
                    Sheet    allocSheet = allocWb.getSheetAt(0);
                    for (Row r : allocSheet) {
                        Cell bc = r.getCell(6), cc = r.getCell(5), sc = r.getCell(2);
                        if (bc == null || cc == null || sc == null) continue;
                        if (!getCellValueAsString(bc).equals(batchname)) continue;
                        if (Integer.parseInt(getCellValueAsString(cc)) != course1) continue;
                        String cand = getCellValueAsString(sc);
                        if ((cand.startsWith("Eng30 Ltr")  && course1 == 1) ||
                            ((cand.startsWith("Eng40 Resume") || cand.startsWith("Eng40 B Ltr")) && course1 == 2)) {
                            System.out.println(cand);
                            modelFile = new File(
                              "F:\\desktop backup 5 july 25\\GCC TBC MAY 2025 AUTO\\allocation\\New folder\\Model Anwers\\" + cand
                            );
                            break;
                        }
                    }
                    allocWb.close();
                }
            }

            // extract paragraphs & table lines
            modelParas     = extractParasAndTables(modelFile);
            List<String> studentParas = extractParasAndTables(studentFile);

            // --- NEW: adjust student enters after the first "Reference" if model has none ---
            adjustStudentEntersAfterKeyword(modelParas, studentParas); // uses multiple regex variants

            List<String[]> wordMistakes = new ArrayList<>();
            XWPFDocument outputDoc = new XWPFDocument();

            // --- compare 2nd-letter addresses (between "To," and "Subject:") ---
            String modelAddress2 = extractNthAddress(modelParas, 2);
            String studentAddress2 = extractNthAddress(studentParas, 2);

            if (!modelAddress2.isEmpty() && !studentAddress2.isEmpty()) {
                XWPFParagraph addrHdr = outputDoc.createParagraph();
                addrHdr.createRun().setBold(true);
                addrHdr.createRun().setText("Address (2nd letter) comparison:");
                XWPFParagraph addrPara = outputDoc.createParagraph();
                compareAddresses(modelAddress2, studentAddress2, addrPara, wordMistakes, studentFile.getName());
            } else if (!modelAddress2.isEmpty() && studentAddress2.isEmpty()) {
                // Student missing the 2nd address entirely -> count as 4 mistakes
                XWPFParagraph addrHdr = outputDoc.createParagraph();
                addrHdr.createRun().setBold(true);
                addrHdr.createRun().setText("Address (2nd letter) comparison:");
                XWPFParagraph addrPara = outputDoc.createParagraph();
                addrPara.createRun().setText("  (Model) " + modelAddress2);
                addrPara.createRun().addBreak();
                addrPara.createRun().setText("  (Student) " + "(missing)");
                // Add four missing-word mistakes so totalMistakes increases by 4
                for (int t = 0; t < 4; t++) {
                    wordMistakes.add(new String[]{studentFile.getName(), "Missing Word", "Address (2nd letter) missing"});
                }
                // Prevent double-counting: blank the model's address lines so main loop won't re-detect them.
                removeNthAddressLines(modelParas, 2);
            }


            // compare up through and including the first "Encl" line
            int mi = 0, si = 0;
            while (mi < modelParas.size() || si < studentParas.size()) {
                String mLine = mi < modelParas.size()   ? modelParas.get(mi)   : "";
                String sLine = si < studentParas.size() ? studentParas.get(si) : "";

                // If either side has empty lines (enters) -> compute consecutive counts and handle
                if (mLine.isEmpty() || sLine.isEmpty()) {
                    int mCount = 0, sCount = 0;
                    while (mi + mCount < modelParas.size() && modelParas.get(mi + mCount).isEmpty()) mCount++;
                    while (si + sCount < studentParas.size() && studentParas.get(si + sCount).isEmpty()) sCount++;

                    // NEW RULE: tolerate differences up to 2 (inclusive). Only mark when abs diff > 2.
                    if (Math.abs(mCount - sCount) <= 2) {
                        // skip those empty-lines on each side
                        mi += mCount;
                        si += sCount;
                        if (mCount == 0 && sCount == 0) { mi++; si++; }
                        continue;
                    } else {
                        // mark as Enter / blank-line mismatch
                        XWPFParagraph p = outputDoc.createParagraph();
                        highlight(p.createRun(), "[Enters:" + mCount + "/" + sCount + "]");
                        wordMistakes.add(new String[]{studentFile.getName(), "Enter Mismatch", mCount + "/" + sCount});
                        mi += mCount;
                        si += sCount;
                        continue;
                    }
                }

                // normal non-empty line compare
                XWPFParagraph p = outputDoc.createParagraph();
                compareWords(mLine, sLine, p, wordMistakes, studentFile.getName());
                // stop if "Encl" found
                if (mLine.toLowerCase().contains("encl") || sLine.toLowerCase().contains("encl")) {
                    break;
                }
                mi++; si++;
            }

            // tally
            int extra = 0, missing = 0, wrong = 0;
            for (String[] m : wordMistakes) {
                switch (m[1]) {
                    case "Extra Word":   extra++;   break;
                    case "Missing Word": missing++; break;
                    case "Wrong Word":   wrong++;   break;
                }
            }
            int   totalMistakes = wordMistakes.size();
            double obtainedMarks = Math.max(0.0, MAX_MARKS - totalMistakes * 0.5);

            // summary in DOCX
            XWPFParagraph sumPara = outputDoc.createParagraph();
            if (extra   > 0) sumPara.createRun().setText("Extra Word: " + extra);
            if (missing > 0) sumPara.createRun().setText("Missing Word: " + missing);
            if (wrong   > 0) sumPara.createRun().setText("Wrong Word: " + wrong);
            sumPara.createRun().setText("Total Mistakes: " + totalMistakes);
            sumPara.createRun().setText(
                String.format("Final Marks: %.1f / %.1f", obtainedMarks, MAX_MARKS)
            );

            File comparedFile = new File(compareDir,
                studentFile.getName().replace(".docx", "_Compared.docx")
            );
            try (FileOutputStream out = new FileOutputStream(comparedFile)) {
                outputDoc.write(out);
            }

            // write Excel row
            Row row = resultSheet.createRow(rowNum++);
            row.createCell(0).setCellValue(studentFile.getName());
            row.createCell(1).setCellValue(extra);
            row.createCell(2).setCellValue(missing);
            row.createCell(3).setCellValue(wrong);
            row.createCell(4).setCellValue(totalMistakes);
            row.createCell(5).setCellValue(obtainedMarks);
            StringBuilder sb = new StringBuilder();
            for (String[] m : wordMistakes) {
                sb.append("[").append(m[1]).append(" - ").append(m[2]).append("] ");
            }
            row.createCell(6).setCellValue(sb.toString().trim());
        }

        // autosize & save
        for (int i = 0; i < columns.length; i++) resultSheet.autoSizeColumn(i);
        try (FileOutputStream out = new FileOutputStream(
                new File(folder.getParent(), "Comparison_Result.xlsx")
        )) {
            resultWb.write(out);
        }
        resultWb.close();
        System.out.println("✅ All students processed.");
    }

    /**
     * Robust version: try multiple keyword regex variants (Ref / Ref. / Reference / Ref No / Reference No).
     * If model has ZERO empty lines after its FIRST match, but student has 1 or 2 empties there,
     * merge them: keyword-line + " " + next non-empty line (remove the empties and the merged line).
     *
     * Debug prints to System.out when a merge happens.
     */
    /**
     * Robust adjuster tuned to the actual files you uploaded.
     *
     * Behavior:
     *  - Find first occurrence of the word "reference" (case-insensitive) in modelParas.
     *  - Count how many empty lines follow that occurrence in model (mEmpty).
     *  - Find the corresponding "reference" in studentParas and count student's empty lines (sEmpty).
     *  - If student has MORE empty lines than model, and (sEmpty - mEmpty) <= 2 and sEmpty > 0,
     *    then merge the student's keyword line with the next non-empty line (or append space if none).
     */
    private static void adjustStudentEntersAfterKeyword(List<String> modelParas, List<String> studentParas) {
        if (modelParas == null || studentParas == null) return;

        // pattern: only 'reference' (case-insensitive). We'll locate occurrences anywhere in a paragraph.
        Pattern referencePat = Pattern.compile("(?i)\\breference\\b");

        // 1) find first 'reference' in model
        int mIdx = -1;
        for (int i = 0; i < modelParas.size(); i++) {
            String r = normalizeForPattern(modelParas.get(i));
            if (r == null) continue;
            Matcher mm = referencePat.matcher(r);
            if (mm.find()) { mIdx = i; break; }
        }
        if (mIdx < 0) return; // no reference in model

        // 2) count empty lines after model's occurrence
        int mEmpty = 0;
        for (int i = mIdx + 1; i < modelParas.size(); i++) {
            String rr = modelParas.get(i);
            if (rr == null || rr.trim().isEmpty()) mEmpty++;
            else break;
        }

        // 3) find first 'reference' in student
        int sIdx = -1;
        for (int i = 0; i < studentParas.size(); i++) {
            String r = normalizeForPattern(studentParas.get(i));
            if (r == null) continue;
            Matcher ms = referencePat.matcher(r);
            if (ms.find()) { sIdx = i; break; }
        }
        if (sIdx < 0) return; // student doesn't have reference -> nothing to do

        // 4) count student's empty lines after its occurrence
        int sEmpty = 0;
        int p = sIdx + 1;
        while (p < studentParas.size() && (studentParas.get(p) == null || studentParas.get(p).trim().isEmpty())) {
            sEmpty++; p++;
        }

        // 5) decide whether to merge:
        // proceed when student has more empties than model AND difference <= 2 AND student has at least 1 empty
        if (sEmpty == 0) return;
        int diff = sEmpty - mEmpty;
        if (diff <= 0 || diff > 2) return;

        // 6) perform merge: base line + " " + next non-empty (if present), then remove the empty lines and the merged next
        String base = studentParas.get(sIdx) == null ? "" : studentParas.get(sIdx);
        if (p < studentParas.size()) {
            String next = studentParas.get(p) == null ? "" : studentParas.get(p);
            String merged = base + " " + next;
            studentParas.set(sIdx, merged);
            // remove items sIdx+1 .. p (that includes sEmpty empties + the next non-empty which we've merged)
            int removes = p - sIdx;
            for (int r = 0; r < removes; r++) {
                if (sIdx + 1 < studentParas.size()) studentParas.remove(sIdx + 1);
            }
            System.out.println("Merged after 'reference' at student index " + sIdx + " (sEmpty=" + sEmpty + ", mEmpty=" + mEmpty + ") -> merged line created.");
        } else {
            // no following non-empty line — just append single space and remove blank lines
            studentParas.set(sIdx, base + " ");
            for (int r = 0; r < sEmpty; r++) {
                if (sIdx + 1 < studentParas.size()) studentParas.remove(sIdx + 1);
            }
            System.out.println("Appended space after 'reference' at student index " + sIdx + " (no following non-empty line).");
        }
    }

    /** Replace NBSP with normal space, but keep other whitespace as-is (no trimming) — used for matching. */
    private static String normalizeForPattern(String text) {
        if (text == null) return null;
        return text.replace('\u00A0', ' ');
    }


    private static List<String> extractParasAndTables(File file) throws IOException {
        List<String> paras = new ArrayList<>();
        try (FileInputStream fis = new FileInputStream(file);
             XWPFDocument doc = new XWPFDocument(fis)) {

            for (IBodyElement e : doc.getBodyElements()) {
                if (e instanceof XWPFParagraph) {
                    paras.addAll(splitRunsIntoLines((XWPFParagraph)e));
                } else if (e instanceof XWPFTable) {
                    addTableLines((XWPFTable)e, paras);
                }
            }
            for (XWPFTable t : doc.getTables()) {
                addTableLines(t, paras);
            }
        }
        return paras;
    }

    private static void addTableLines(XWPFTable table, List<String> paras) {
        for (XWPFTableRow row : table.getRows()) {
            for (XWPFTableCell cell : row.getTableCells()) {
                for (XWPFParagraph p : cell.getParagraphs()) {
                    paras.addAll(splitRunsIntoLines(p));
                }
            }
        }
    }

    private static List<String> splitRunsIntoLines(XWPFParagraph p) {
        StringBuilder sb = new StringBuilder();
        for (XWPFRun run : p.getRuns()) {
            // PRESERVE tab characters (\t) so they are counted as tabs (not converted to spaces)
            sb.append(run.toString());
            int brCount = run.getCTR().getBrList().size();
            for (int i=0; i<brCount; i++) sb.append("\n");
        }
        List<String> out = new ArrayList<>();
        // preserve empty lines so we can detect multiple enters
        String[] lines = sb.toString().split("\\r?\\n", -1); // include trailing empties
        for (String line : lines) {
            if (line == null) continue;
            if (line.trim().isEmpty()) out.add(""); // preserve blank line
            else out.add(line); // IMPORTANT: keep exact internal/leading/trailing whitespace
        }
        return out;
    }

    /**
     * Compare two lines word-by-word, and check per-type whitespace (spaces/tabs/enters) between tokens.
     * Spacing rule changed: count as mismatch only when absolute difference > 2.
     */
    private static void compareWords(String model, String student,
                                     XWPFParagraph para,
                                     List<String[]> mistakes,
                                     String fileName) {
        String modelOrig = model == null ? "" : model;
        String studentOrig = student == null ? "" : student;

        // Ensure student spacing mirrors model punctuation rules before token parsing (but don't collapse multiples)
        studentOrig = ensureStudentSpacingFromModel(modelOrig, studentOrig);

        // Normalize newlines to '\n' for counting consistency
        modelOrig = modelOrig.replace("\r\n", "\n").replace("\r", "\n");
        studentOrig = studentOrig.replace("\r\n", "\n").replace("\r", "\n");

        // Parse tokens and whitespace counts
        TokenCounts modelTC = parseTokensAndCounts(modelOrig);
        TokenCounts studentTC = parseTokensAndCounts(studentOrig);

        String[] mW = modelTC.words.toArray(new String[0]);
        String[] sW = studentTC.words.toArray(new String[0]);

        // Compare leading whitespace (before first token) if present
        if (modelTC.hasLeading || studentTC.hasLeading) {
            int mSpaces = modelTC.leadingSpaces;
            int mTabs   = modelTC.leadingTabs;
            int mEnters = modelTC.leadingEnters;
            int sSpaces = studentTC.leadingSpaces;
            int sTabs   = studentTC.leadingTabs;
            int sEnters = studentTC.leadingEnters;

            boolean leadMismatch = false;
            StringBuilder detail = new StringBuilder();
            if (isSpacingMismatch(mSpaces, sSpaces)) {
                leadMismatch = true;
                detail.append("lead-space:").append(mSpaces).append("/").append(sSpaces).append(";");
            }
            if (isSpacingMismatch(mTabs, sTabs)) {
                leadMismatch = true;
                detail.append("lead-tab:").append(mTabs).append("/").append(sTabs).append(";");
            }
            if (isSpacingMismatch(mEnters, sEnters)) {
                leadMismatch = true;
                detail.append("lead-enter:").append(mEnters).append("/").append(sEnters).append(";");
            }
            if (leadMismatch) {
                highlight(para.createRun(), "[spc:" + detail.toString() + "]");
                mistakes.add(new String[]{fileName, "Spacing", "Leading->" + detail.toString()});
            }
        }

        // Normalize punctuation spacing for display & LCS (but counts are already captured)
        String modelForCompare = modelOrig.replaceAll("\\s+([,;:!?])","$1").replaceAll("\\.\\s*", ". ").replaceAll(":\\s*", ": ");
        String studentForCompare = studentOrig.replaceAll("\\s+([,;:!?])","$1").replaceAll("\\.\\s*", ". ").replaceAll(":\\s*", ": ");

        // Use LCS on tokens
        List<int[]> lcs = computeLCS(mW, sW);

        int i = 0, j = 0, k = 0;
        while (i < mW.length || j < sW.length) {
            if (k < lcs.size() && i == lcs.get(k)[0] && j == lcs.get(k)[1]) {
                // matched token -> write token
                addTextRun(para, sW[j] + " ", false);

                // check whitespace after this token (space, tab, enter)
                int mSpaces = (i < modelTC.spacesAfter.size()) ? modelTC.spacesAfter.get(i)[0] : 0;
                int mTabs   = (i < modelTC.spacesAfter.size()) ? modelTC.spacesAfter.get(i)[1] : 0;
                int mEnters = (i < modelTC.spacesAfter.size()) ? modelTC.spacesAfter.get(i)[2] : 0;

                int sSpaces = (j < studentTC.spacesAfter.size()) ? studentTC.spacesAfter.get(j)[0] : 0;
                int sTabs   = (j < studentTC.spacesAfter.size()) ? studentTC.spacesAfter.get(j)[1] : 0;
                int sEnters = (j < studentTC.spacesAfter.size()) ? studentTC.spacesAfter.get(j)[2] : 0;

                StringBuilder mismatchDetail = new StringBuilder();
                boolean anyMismatch = false;
                if (isSpacingMismatch(mSpaces, sSpaces)) {
                    anyMismatch = true;
                    mismatchDetail.append("space:").append(mSpaces).append("/").append(sSpaces).append(";");
                }
                if (isSpacingMismatch(mTabs, sTabs)) {
                    anyMismatch = true;
                    mismatchDetail.append("tab:").append(mTabs).append("/").append(sTabs).append(";");
                }
                if (isSpacingMismatch(mEnters, sEnters)) {
                    anyMismatch = true;
                    mismatchDetail.append("enter:").append(mEnters).append("/").append(sEnters).append(";");
                }
                if (anyMismatch) {
                    highlight(para.createRun(), "[spc:" + mismatchDetail.toString() + "]");
                    mistakes.add(new String[]{fileName, "Spacing", mW[i] + " -> " + mismatchDetail.toString()});
                }

                i++; j++; k++;
            }
            else if (i < mW.length && j < sW.length &&
                     (k >= lcs.size() || (i < lcs.get(k)[0] && j < lcs.get(k)[1])) &&
                     !normalize(mW[i]).equals(normalize(sW[j]))) {
                highlight(para.createRun(),"["+mW[i]+"/"+sW[j]+"]");
                mistakes.add(new String[]{fileName,"Wrong Word",mW[i]+"/"+sW[j]});
                i++; j++;
            }
            else if (k < lcs.size() && i < lcs.get(k)[0]) {
                highlight(para.createRun(),"["+mW[i]+"]");
                mistakes.add(new String[]{fileName,"Missing Word",mW[i]});
                i++;
            }
            else if (k < lcs.size() && j < lcs.get(k)[1]) {
                highlight(para.createRun(),"["+sW[j]+"]");
                mistakes.add(new String[]{fileName,"Extra Word",sW[j]});
                j++;
            }
            else if (i < mW.length) {
                highlight(para.createRun(),"["+mW[i]+"]");
                mistakes.add(new String[]{fileName,"Missing Word",mW[i]});
                i++;
            }
            else if (j < sW.length) {
                highlight(para.createRun(),"["+sW[j]+"]");
                mistakes.add(new String[]{fileName,"Extra Word",sW[j]});
                j++;
            }
            else {
                i++; j++;
            }
        }
    }

    /**
     * Return true when whitespace counts should be considered a mismatch according to rule:
     *  - count as mismatch only when absolute difference > 2
     */
    private static boolean isSpacingMismatch(int modelCount, int studentCount) {
        return Math.abs(modelCount - studentCount) > 2;
    }

    private static List<int[]> computeLCS(String[] a, String[] b) {
        int m=a.length, n=b.length;
        int[][] dp = new int[m+1][n+1];
        for (int i=1; i<=m; i++) {
            for (int j=1; j<=n; j++) {
                if (normalize(a[i-1]).equals(normalize(b[j-1])))
                    dp[i][j] = dp[i-1][j-1] + 1;
                else
                    dp[i][j] = Math.max(dp[i-1][j], dp[i][j-1]);
            }
        }
        List<int[]> lcs = new ArrayList<>();
        int i=m, j=n;
        while (i>0 && j>0) {
            if (normalize(a[i-1]).equals(normalize(b[j-1]))) {
                lcs.add(0, new int[]{i-1,j-1}); i--; j--;
            } else if (dp[i-1][j]>=dp[i][j-1]) i--; else j--;
        }
        return lcs;
    }

    private static String getCellValueAsString(Cell cell) {
        switch (cell.getCellType()) {
            case STRING:  return cell.getStringCellValue().trim();
            case NUMERIC: return String.valueOf((int)cell.getNumericCellValue());
            default:      return "";
        }
    }

    private static String normalize(String w) {
        return w==null? "": w.trim().replace("’","'").replace("‘","'");
    }

    /**
     * Ensure student spacing around punctuation matches model's pattern.
     * For each punctuation in {.,;:!?,-}:
     *
     * NOTE: This adjusts spacing for punctuation but does NOT collapse multiple whitespace,
     * so counts of spaces/tabs/enters are preserved for later comparison.
     */
    private static String ensureStudentSpacingFromModel(String model, String student) {
        if (student == null) student = "";
        if (model == null) model = "";

        // --- Protect specific sequences in student's text so they remain unchanged ---
        Pattern[] protectPatterns = new Pattern[] {
            // Ref. No.  or Ref. No. :  (case-insensitive, allow variable spaces)
            Pattern.compile("(?i)Ref\\.\\s*No\\.\\s*:?"),
            // Encl.  or Encl. :  (case-insensitive, allow variable spaces)
            Pattern.compile("(?i)Encl\\.\\s*:?")
        };
        // keep order and original text for restoration
        Map<String,String> placeholders = new LinkedHashMap<>();
        int protIdx = 0;
        for (Pattern p : protectPatterns) {
            Matcher m = p.matcher(student);
            StringBuffer sb = new StringBuffer();
            while (m.find()) {
                String found = m.group();
                String key = "__PROT" + (protIdx++) + "__";
                placeholders.put(key, found);
                m.appendReplacement(sb, Matcher.quoteReplacement(key));
            }
            m.appendTail(sb);
            student = sb.toString();
        }

        // --- punctuation spacing logic (applies to the student string with protected parts replaced) ---
        char[] puncts = {'.', ',', ';', ':', '!', '?', '-'};

        for (char p : puncts) {
            String pQ = Pattern.quote(String.valueOf(p)); // safe regex-escaped punctuation

            // does model show a space before this punctuation anywhere? (e.g. " word ,")
            boolean modelHasSpaceBefore = Pattern.compile("\\s" + pQ).matcher(model).find();

            // does model show a space after this punctuation anywhere? (e.g. ", word")
            boolean modelHasSpaceAfter  = Pattern.compile(pQ + "\\s").matcher(model).find();

            // If model requires space BEFORE punctuation, add space before any punctuation in student that lacks it
            if (modelHasSpaceBefore) {
                // replace punctuation not preceded by whitespace with " <punct>"
                student = student.replaceAll("(?<!\\s)" + pQ, " " + p);
            }

            // If model requires space AFTER punctuation, add a single space after any punctuation in student that lacks it
            if (modelHasSpaceAfter) {
                // replace punctuation not followed by whitespace with "<punct> "
                student = student.replaceAll(pQ + "(?!\\s)", p + " ");
            }
        }

        // IMPORTANT: do NOT collapse multiple spaces here and do NOT trim — we must preserve exact counts

        // --- Restore protected substrings back to their original text ---
        for (Map.Entry<String,String> e : placeholders.entrySet()) {
            student = student.replace(e.getKey(), e.getValue());
        }

        return student;
    }

    private static void addTextRun(XWPFParagraph para, String text, boolean highlight) {
        XWPFRun run = para.createRun();
        run.setText(text);
        if (highlight) {
            CTRPr rpr = run.getCTR().isSetRPr()? run.getCTR().getRPr(): run.getCTR().addNewRPr();
            CTShd shd = rpr.addNewShd();
            shd.setVal(STShd.CLEAR);
            shd.setFill("FF0000");
        }
    }

    private static void highlight(XWPFRun run, String text) {
        run.setText(text + " ");
        run.setColor("FF0000");
    }

    // Helper structure to store tokens with counts of (spaces, tabs, enters) AFTER each token.
    private static class TokenCounts {
        List<String> words = new ArrayList<>();
        // each int[] is {spaces, tabs, enters}
        List<int[]> spacesAfter = new ArrayList<>();
        // leading whitespace counts before first token
        boolean hasLeading = false;
        int leadingSpaces = 0;
        int leadingTabs = 0;
        int leadingEnters = 0;
    }

    /**
     * Parse tokens (non-whitespace sequences) and count number of spaces/tabs/enters between tokens.
     * Leading whitespace (before first token) is stored separately.
     */
    private static TokenCounts parseTokensAndCounts(String s) {
        TokenCounts tc = new TokenCounts();
        if (s == null || s.isEmpty()) return tc;

        // normalize newlines
        s = s.replace("\r\n", "\n").replace("\r", "\n");

        Pattern p = Pattern.compile("\\S+");
        Matcher m = p.matcher(s);
        List<int[]> positions = new ArrayList<>();
        while (m.find()) {
            positions.add(new int[]{m.start(), m.end()});
        }
        if (positions.isEmpty()) {
            // no tokens, but we may still have leading whitespace (whole string)
            int spaces = 0, tabs = 0, enters = 0;
            for (int k = 0; k < s.length(); k++) {
                char c = s.charAt(k);
                if (c == ' ') spaces++;
                else if (c == '\t') tabs++;
                else if (c == '\n') enters++;
            }
            if (spaces + tabs + enters > 0) {
                tc.hasLeading = true;
                tc.leadingSpaces = spaces;
                tc.leadingTabs = tabs;
                tc.leadingEnters = enters;
            }
            return tc;
        }

        // leading whitespace before first token
        int firstStart = positions.get(0)[0];
        if (firstStart > 0) {
            tc.hasLeading = true;
            for (int k = 0; k < firstStart; k++) {
                char c = s.charAt(k);
                if (c == ' ') tc.leadingSpaces++;
                else if (c == '\t') tc.leadingTabs++;
                else if (c == '\n') tc.leadingEnters++;
            }
        }

        for (int idx = 0; idx < positions.size(); idx++) {
            int start = positions.get(idx)[0];
            int end = positions.get(idx)[1];
            tc.words.add(s.substring(start, end));
            int wsStart = end;
            int wsEnd;
            if (idx + 1 < positions.size()) wsEnd = positions.get(idx + 1)[0];
            else wsEnd = s.length();
            int spaces = 0, tabs = 0, enters = 0;
            for (int k = wsStart; k < wsEnd; k++) {
                char c = s.charAt(k);
                if (c == ' ') spaces++;
                else if (c == '\t') tabs++;
                else if (c == '\n') enters++;
            }
            tc.spacesAfter.add(new int[]{spaces, tabs, enters});
        }
        return tc;
    }

    /**
     * Compare two addresses word-by-word and check per-type whitespace counts similarly to compareWords.
     */
    private static void compareAddresses(String modelAddr, String studentAddr,
                                         XWPFParagraph para,
                                         List<String[]> mistakes,
                                         String fileName) {

        String modelOrig = modelAddr == null ? "" : modelAddr;
        String studentOrig = studentAddr == null ? "" : studentAddr;

        studentOrig = ensureStudentSpacingFromModel(modelOrig, studentOrig);

        modelOrig = modelOrig.replace("\r\n", "\n").replace("\r", "\n");
        studentOrig = studentOrig.replace("\r\n", "\n").replace("\r", "\n");

        TokenCounts modelTC = parseTokensAndCounts(modelOrig);
        TokenCounts studentTC = parseTokensAndCounts(studentOrig);

        para.createRun().setText("  (Model) " + modelOrig);
        para.createRun().addBreak();
        para.createRun().setText("  (Student) ");

        String[] mW = modelTC.words.toArray(new String[0]);
        String[] sW = studentTC.words.toArray(new String[0]);

        // leading whitespace check
        if (modelTC.hasLeading || studentTC.hasLeading) {
            int mSpaces = modelTC.leadingSpaces, mTabs = modelTC.leadingTabs, mEnters = modelTC.leadingEnters;
            int sSpaces = studentTC.leadingSpaces, sTabs = studentTC.leadingTabs, sEnters = studentTC.leadingEnters;
            boolean leadMismatch = false;
            StringBuilder detail = new StringBuilder();
            if (isSpacingMismatch(mSpaces, sSpaces)) { leadMismatch = true; detail.append("lead-space:").append(mSpaces).append("/").append(sSpaces).append(";"); }
            if (isSpacingMismatch(mTabs, sTabs))   { leadMismatch = true; detail.append("lead-tab:").append(mTabs).append("/").append(sTabs).append(";"); }
            if (isSpacingMismatch(mEnters, sEnters)) { leadMismatch = true; detail.append("lead-enter:").append(mEnters).append("/").append(sEnters).append(";"); }
            if (leadMismatch) {
                highlight(para.createRun(), "[spc:" + detail.toString() + "]");
                mistakes.add(new String[]{fileName, "Spacing", "Leading->" + detail.toString()});
            }
        }

        List<int[]> lcs = computeLCS(mW, sW);

        int i = 0, j = 0, k = 0;
        while (i < mW.length || j < sW.length) {
            if (k < lcs.size() && i == lcs.get(k)[0] && j == lcs.get(k)[1]) {
                addTextRun(para, sW[j] + " ", false);

                int[] mAfter = (i < modelTC.spacesAfter.size()) ? modelTC.spacesAfter.get(i) : new int[]{0,0,0};
                int[] sAfter = (j < studentTC.spacesAfter.size()) ? studentTC.spacesAfter.get(j) : new int[]{0,0,0};

                StringBuilder mismatchDetail = new StringBuilder();
                boolean anyMismatch = false;
                if (isSpacingMismatch(mAfter[0], sAfter[0])) { anyMismatch = true; mismatchDetail.append("space:").append(mAfter[0]).append("/").append(sAfter[0]).append(";"); }
                if (isSpacingMismatch(mAfter[1], sAfter[1])) { anyMismatch = true; mismatchDetail.append("tab:").append(mAfter[1]).append("/").append(sAfter[1]).append(";"); }
                if (isSpacingMismatch(mAfter[2], sAfter[2])) { anyMismatch = true; mismatchDetail.append("enter:").append(mAfter[2]).append("/").append(sAfter[2]).append(";"); }

                if (anyMismatch) {
                    highlight(para.createRun(), "[spc:" + mismatchDetail.toString() + "]");
                    mistakes.add(new String[]{fileName, "Spacing", mW[i] + " -> " + mismatchDetail.toString()});
                }

                i++; j++; k++;
            }
            else if (i < mW.length && j < sW.length &&
                    (k >= lcs.size() || (i < lcs.get(k)[0] && j < lcs.get(k)[1])) &&
                     !normalize(mW[i]).equals(normalize(sW[j]))) {
                highlight(para.createRun(), "[" + mW[i] + "/" + sW[j] + "]");
                mistakes.add(new String[]{fileName, "Wrong Word", mW[i] + "/" + sW[j]});
                i++; j++;
            }
            else if (k < lcs.size() && i < lcs.get(k)[0]) {
                highlight(para.createRun(), "[" + mW[i] + "]");
                mistakes.add(new String[]{fileName, "Missing Word", mW[i]});
                i++;
            }
            else if (k < lcs.size() && j < lcs.get(k)[1]) {
                highlight(para.createRun(), "[" + sW[j] + "]");
                mistakes.add(new String[]{fileName, "Extra Word", sW[j]});
                j++;
            }
            else if (i < mW.length) {
                highlight(para.createRun(), "[" + mW[i] + "]");
                mistakes.add(new String[]{fileName, "Missing Word", mW[i]});
                i++;
            }
            else if (j < sW.length) {
                highlight(para.createRun(), "[" + sW[j] + "]");
                mistakes.add(new String[]{fileName, "Extra Word", sW[j]});
                j++;
            }
            else {
                i++; j++;
            }
        }
        para.createRun().addBreak();
    }

    private static String extractNthAddress(List<String> paras, int n) {
        if (paras == null || paras.isEmpty() || n <= 0) return "";
        int count = 0;
        boolean collecting = false;
        StringBuilder sb = new StringBuilder();

        for (String raw : paras) {
            if (raw == null) continue;
            String line = raw.trim();
            if (!collecting) {
                // Detect a standalone "To" line (e.g. "To," or "To") OR inline "To, Name..."
                if (line.matches("(?i)^to\\s*,?$")) {
                    count++;
                    if (count == n) {
                        collecting = true;
                        continue; // next lines are part of address
                    }
                } else if (line.toLowerCase().startsWith("to,") || line.toLowerCase().startsWith("to ")) {
                    // Inline To followed by address on same line
                    count++;
                    if (count == n) {
                        String after = line.replaceFirst("(?i)^to\\s*,?\\s*", "");
                        if (!after.isEmpty()) sb.append(after).append(" ");
                        collecting = true;
                        continue;
                    }
                }
            } else {
                // We are collecting address lines until we hit a "Subject" marker
                if (line.toLowerCase().contains("subject")) {
                    break;
                }
                if (!line.isEmpty()) {
                    sb.append(line).append(" ");
                }
            }
        }
        return sb.toString().trim();
    }

    /**
     * Remove (blank out) the lines which formed the nth "To" address (used by extractNthAddress).
     * This prevents the main compare loop from counting the same address as missing again.
     */
    private static void removeNthAddressLines(List<String> paras, int n) {
        if (paras == null || paras.isEmpty() || n <= 0) return;
        int count = 0;
        boolean collecting = false;

        for (int idx = 0; idx < paras.size(); idx++) {
            String raw = paras.get(idx);
            String line = raw == null ? "" : raw.trim();

            if (!collecting) {
                // match a standalone "To" line like "To" or "To,"
                if (line.matches("(?i)^to\\s*,?$")) {
                    count++;
                    if (count == n) {
                        // blank this "To" line and start collecting subsequent address lines
                        paras.set(idx, "");
                        collecting = true;
                    }
                }
                // inline "To, Name..." on the same line
                else if (line.toLowerCase().startsWith("to,") || line.toLowerCase().startsWith("to ")) {
                    count++;
                    if (count == n) {
                        // blank this line (it contained the address) and start collecting
                        paras.set(idx, "");
                        collecting = true;
                    }
                }
            } else {
                // we're collecting address lines until we hit "Subject"
                if (line.toLowerCase().contains("subject")) {
                    break;
                }
                // blank this address line
                paras.set(idx, "");
            }
        }
    }

}
