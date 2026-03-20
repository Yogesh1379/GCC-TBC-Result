package marking_verification;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.*;
import java.util.HashSet;

public class Check_Marking_run_allSeatnumber {

    static public void main(String[] args) throws IOException {
        File originalfile= new File("F:\\GCC  TBC December 2025\\ssd50\\WrongcomparefilechekcEng50.xlsx");// file which is given to marking with que name and std ans file
        File markingFile=new File("F:\\GCC  TBC December 2025\\Marking\\eng50\\ENG 50_MERGE DATA.xlsx"); //merge marking data
        DataFormatter df = new DataFormatter();
        HashSet<Object> filenameset = new HashSet<>();
        FileInputStream fisoriginal=new FileInputStream(originalfile);
        XSSFWorkbook workbook = new XSSFWorkbook(fisoriginal);
        Sheet  originalsheet=workbook.getSheetAt(0);
        FileInputStream fismark = new FileInputStream(markingFile);
        XSSFWorkbook Markworkbook = new XSSFWorkbook(fismark);
        XSSFSheet markSheet = Markworkbook.getSheetAt(0);
        // CREATE RESULT WORKBOOK
        Workbook resultWorkbook = new XSSFWorkbook();
        
        Sheet matchSheet = resultWorkbook.createSheet("Match Result");
        Sheet mismatchSheet = resultWorkbook.createSheet("Mismatch Result");

        Row matchHeader = matchSheet.createRow(0);
        matchHeader.createCell(0).setCellValue("Matched Files");

        Row mismatchHeader = mismatchSheet.createRow(0);
        mismatchHeader.createCell(0).setCellValue("Mis Matched Files");

        int matchRow = 1;
        int notFoundRow = 1;


        for(int i=1;i<=markSheet.getLastRowNum();i++)
        {
            Row row=markSheet.getRow(i);
           Cell cell= row.getCell(15);  //have to change for lrt spd stmt
            String cellValue = df.formatCellValue(cell).replaceAll("DiffDoc_","");
System.out.println(cellValue);
            filenameset.add(cellValue);
        }

        for(int i=1;i<=originalsheet.getLastRowNum();i++)
        {
            Row row=originalsheet.getRow(i);
            Cell cell= row.getCell(3);  //have to change for lrt spd stmt
            String cellValue = df.formatCellValue(cell);
            if(filenameset.contains(cellValue)){
                Row r = matchSheet.getRow(matchRow);
                if (r == null) r = matchSheet.createRow(matchRow);
                r.createCell(0).setCellValue(cellValue);
                matchRow++;
                System.out.println("match :- "+cellValue);
            }
            else{
                Row r = mismatchSheet.getRow(notFoundRow);
                if (r == null) r = mismatchSheet.createRow(notFoundRow);
                r.createCell(0).setCellValue(cellValue);
                notFoundRow++;
                System.out.println("notfount :- "+cellValue);
            }
        }

        // SAVE FILE
        FileOutputStream fos = new FileOutputStream(
                "F:\\GCC  TBC December 2025\\Marking\\eng50\\exl Mismatch File.xlsx");
        resultWorkbook.write(fos);

        fos.close();
        resultWorkbook.close();
        workbook.close();
        Markworkbook.close();

        System.out.println(" Result Excel created successfully");



    }
}
