package GRACE_MARKS;
import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.xml.parsers.SAXParserFactory;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler.SheetContentsHandler;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import org.apache.poi.xssf.usermodel.XSSFComment;
import org.xml.sax.SAXException;

public class Compare_Result_Streaming {

    public static void main(String[] args) throws Exception {

        String oldFile = "C:\\Users\\User\\Desktop\\Register june\\final\\Book2.xlsx";
        String newFile = "C:\\Users\\User\\Desktop\\Register june\\Data\\Eng30 Result data.xlsx";

        // Map to store old data
        Map<String, String[]> oldData = new HashMap<>();

        // Read OLD Excel using streaming
        readExcelStreaming(oldFile, "Register", oldData);

        // Prepare SXSSFWorkbook for writing output (memory-efficient)
        SXSSFWorkbook outWb = new SXSSFWorkbook(100); // keep 100 rows in memory
        Sheet outSheet = outWb.createSheet("Comparison");

        // Header row
        Row header = outSheet.createRow(0);
        String[] headers = {
                "Seat Number",
                "Old Grace Marks","New Grace Marks",
                "Old Percent","New Percent",
                "Old Result","New Result",
                "Old Center Code","New Center Code",
                "Old Center Address","New Center Address",
                "Old Inst Code","New Inst Code",
                "Old Obj Grace","New Obj Grace",
                "Old Speed Grace","New Speed Grace",
                "Old Email Stmt Ltr Grace","New Email Stmt Ltr Grace",
                "Old Grade","New Grade",
                "Status"
        };
        for(int i=0;i<headers.length;i++){
            header.createCell(i).setCellValue(headers[i]);
        }

        // Read NEW Excel using streaming and compare
        compareExcelStreaming(newFile, "Sheet1", oldData, outSheet);

        // Write output
        try (FileOutputStream fos = new FileOutputStream("C:\\Users\\User\\Desktop\\Register june\\final\\Compare_Result_Eng30.xlsx")) {
            outWb.write(fos);
        }
        outWb.dispose(); // important for SXSSFWorkbook
        outWb.close();

        System.out.println("✅ Comparison completed successfully (streaming).");
    }

    private static void readExcelStreaming(String filePath, String sheetName, Map<String, String[]> oldData) throws Exception {

        OPCPackage pkg = OPCPackage.open(new File(filePath));
        ReadOnlySharedStringsTable strings = new ReadOnlySharedStringsTable(pkg);
        XSSFReader xssfReader = new XSSFReader(pkg);
        XSSFReader.SheetIterator iter = (XSSFReader.SheetIterator) xssfReader.getSheetsData();

        DataFormatter df = new DataFormatter();

        while(iter.hasNext()) {
            try (InputStream sheetInputStream = iter.next()) {
                String name = iter.getSheetName();
                if(!name.equalsIgnoreCase(sheetName)) continue;

                XSSFSheetXMLHandler.SheetContentsHandler handler = new SheetContentsHandler() {
                    int currentRow = -1;
                    List<String> rowData = new ArrayList<>();

                    @Override
                    public void startRow(int rowNum) {
                        currentRow = rowNum;
                        rowData.clear();
                    }

                    @Override
                    public void endRow(int rowNum) {
                        if(currentRow == 0) return; // skip header
                        if(rowData.size() < 25) return; // skip incomplete rows

                        String seat = rowData.get(3);
                        if(seat.isEmpty()) return;
                        System.out.println(seat);
                        oldData.put(seat, new String[]{
                                rowData.get(21), // grace Marks
                                rowData.get(22), // percent
                                rowData.get(23), // result
                                rowData.get(6),  // center code
                                rowData.get(7),  // center address
                                rowData.get(9),  // inst code
                                rowData.get(13), // obj grace
                                rowData.get(18), // speed grace
                                rowData.get(20), // email stmt ltr grace
                                rowData.get(24)  // grade
                        });
                    }

                    @Override
                    public void cell(String cellReference, String formattedValue, XSSFComment comment) {
                        rowData.add(formattedValue);
                    }

                    @Override
                    public void headerFooter(String text, boolean isHeader, String tagName) {}
                };

                XSSFSheetXMLHandler xmlHandler = new XSSFSheetXMLHandler(
                        xssfReader.getStylesTable(), null, strings, handler, df, false);

                SAXParserFactory.newInstance()
                        .newSAXParser()
                        .parse(sheetInputStream, new org.xml.sax.helpers.DefaultHandler() {
                            @Override
                            public void startElement(String uri, String localName, String qName, org.xml.sax.Attributes attributes) throws SAXException {
                                xmlHandler.startElement(uri, localName, qName, attributes);
                            }

                            @Override
                            public void endElement(String uri, String localName, String qName) throws SAXException {
                                xmlHandler.endElement(uri, localName, qName);
                            }

                            @Override
                            public void characters(char[] ch, int start, int length) throws SAXException {
                                xmlHandler.characters(ch, start, length);
                            }
                        });
            }
        }
        pkg.close();
    }

    private static void compareExcelStreaming(String filePath, String sheetName, Map<String,String[]> oldData, Sheet outSheet) throws Exception {

        OPCPackage pkg = OPCPackage.open(new File(filePath));
        ReadOnlySharedStringsTable strings = new ReadOnlySharedStringsTable(pkg);
        XSSFReader xssfReader = new XSSFReader(pkg);
        XSSFReader.SheetIterator iter = (XSSFReader.SheetIterator) xssfReader.getSheetsData();
        DataFormatter df = new DataFormatter();

        final int[] rowIndex = {1};

        while(iter.hasNext()) {
            try (InputStream sheetInputStream = iter.next()) {
                String name = iter.getSheetName();
                if(!name.equalsIgnoreCase(sheetName)) continue;

                XSSFSheetXMLHandler.SheetContentsHandler handler = new SheetContentsHandler() {
                    int currentRow = -1;
                    List<String> rowData = new ArrayList<>();

                    @Override
                    public void startRow(int rowNum) {
                        currentRow = rowNum;
                        rowData.clear();
                    }

                    @Override
                    public void endRow(int rowNum) {
                        if(currentRow==0) return; // skip header
                        if(rowData.size()<25) return; // skip incomplete

                        String seat = rowData.get(3);
                        if(seat.isEmpty()) return;

                        String newGraceMarks   = rowData.get(21);
                        String newPercent      = rowData.get(22);
                        String newResult       = rowData.get(23);
                        String newCenterCode   = rowData.get(6);
                        String newCenterAddr   = rowData.get(7);
                        String newInstCode     = rowData.get(9);
                        String newObjGrace     = rowData.get(13);
                        String newSpeedGrace   = rowData.get(18);
                        String newEmailGrace   = rowData.get(20);
                        String newGrade        = rowData.get(24);

                        String[] oldValues = oldData.get(seat);

                        Row outRow = outSheet.createRow(rowIndex[0]++);
                        if(oldValues==null){
                            outRow.createCell(0).setCellValue(seat);
                            outRow.createCell(1).setCellValue("N/A");
                            outRow.createCell(2).setCellValue(newGraceMarks);
                            outRow.createCell(3).setCellValue("N/A");
                            outRow.createCell(4).setCellValue(newPercent);
                            outRow.createCell(5).setCellValue("N/A");
                            outRow.createCell(6).setCellValue(newResult);
                            outRow.createCell(7).setCellValue("N/A");
                            outRow.createCell(8).setCellValue(newCenterCode);
                            outRow.createCell(9).setCellValue("N/A");
                            outRow.createCell(10).setCellValue(newCenterAddr);
                            outRow.createCell(11).setCellValue("N/A");
                            outRow.createCell(12).setCellValue(newInstCode);
                            outRow.createCell(13).setCellValue(newObjGrace);
                            outRow.createCell(14).setCellValue(newSpeedGrace);
                            outRow.createCell(15).setCellValue(newEmailGrace);
                            outRow.createCell(16).setCellValue(newGrade);
                            outRow.createCell(17).setCellValue("Seat not in OLD");
                            return;
                        }

                        String oldGraceMarks = oldValues[0];
                        String oldPercent    = oldValues[1];
                        String oldResult     = oldValues[2];
                        String oldCenterCode = oldValues[3];
                        String oldCenterAddr = oldValues[4];
                        String oldInstCode   = oldValues[5];
                        String oldObjGrace   = oldValues[6];
                        String oldSpeedGrace = oldValues[7];
                        String oldEmailGrace = oldValues[8];
                        String oldGrade      = oldValues[9];

                        boolean matches = oldGraceMarks.equals(newGraceMarks)
                                && oldPercent.equals(newPercent)
                                && oldResult.equals(newResult)
                                && oldCenterCode.equals(newCenterCode)
                                && oldCenterAddr.equals(newCenterAddr)
                                && oldInstCode.equals(newInstCode)
                                && oldObjGrace.equals(newObjGrace)
                                && oldSpeedGrace.equals(newSpeedGrace)
                                && oldEmailGrace.equals(newEmailGrace)
                                && oldGrade.equals(newGrade);

                        if(!matches){
                            outRow.createCell(0).setCellValue(seat);
                            outRow.createCell(1).setCellValue(oldGraceMarks);
                            outRow.createCell(2).setCellValue(newGraceMarks);
                            outRow.createCell(3).setCellValue(oldPercent);
                            outRow.createCell(4).setCellValue(newPercent);
                            outRow.createCell(5).setCellValue(oldResult);
                            outRow.createCell(6).setCellValue(newResult);
                            outRow.createCell(7).setCellValue(oldCenterCode);
                            outRow.createCell(8).setCellValue(newCenterCode);
                            outRow.createCell(9).setCellValue(oldCenterAddr);
                            outRow.createCell(10).setCellValue(newCenterAddr);
                            outRow.createCell(11).setCellValue(oldInstCode);
                            outRow.createCell(12).setCellValue(newInstCode);
                            outRow.createCell(13).setCellValue(oldObjGrace);
                            outRow.createCell(14).setCellValue(newObjGrace);
                            outRow.createCell(15).setCellValue(oldSpeedGrace);
                            outRow.createCell(16).setCellValue(newSpeedGrace);
                            outRow.createCell(17).setCellValue(oldEmailGrace);
                            outRow.createCell(18).setCellValue(newEmailGrace);
                            outRow.createCell(19).setCellValue(oldGrade);
                            outRow.createCell(20).setCellValue(newGrade);
                            outRow.createCell(21).setCellValue("DOES NOT MATCH");
                        }
                    }

                    @Override
                    public void cell(String cellReference, String formattedValue, XSSFComment comment) { rowData.add(formattedValue); }
                    @Override public void headerFooter(String text, boolean isHeader, String tagName) {}
                };

                XSSFSheetXMLHandler xmlHandler = new XSSFSheetXMLHandler(
                        xssfReader.getStylesTable(), null, strings, handler, df, false);

                SAXParserFactory.newInstance()
                        .newSAXParser()
                        .parse(sheetInputStream, new org.xml.sax.helpers.DefaultHandler() {
                            @Override
                            public void startElement(String uri, String localName, String qName, org.xml.sax.Attributes attributes) {
                                try { xmlHandler.startElement(uri, localName, qName, attributes); } catch (Exception e) {}
                            }
                            @Override
                            public void endElement(String uri, String localName, String qName) {
                                try { xmlHandler.endElement(uri, localName, qName); } catch (Exception e) {}
                            }
                            @Override
                            public void characters(char[] ch, int start, int length) {
                                try { xmlHandler.characters(ch, start, length); } catch (Exception e) {}
                            }
                        });
            }
        }
        pkg.close();
    }
}
