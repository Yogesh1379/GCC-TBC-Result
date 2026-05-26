package Bookmark_Verification_Statement;

import org.testng.annotations.Test;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Bookmark_Testing {

	@Test
	public void excelBookmark() throws EncryptedDocumentException, IOException {
		String folderpath="F:\\GCCTBC-APR 2026\\ENGLISH-BOOK_MARK 28-04-2026\\SRP_E-40-Statement\\New folder";
		String Outpath="F:\\GCCTBC-APR 2026\\ENGLISH-BOOK_MARK 28-04-2026\\SRP_E-40-Statement\\New folder\\bookmark verification.xlsx";

		File folder = new File(folderpath);

		File[] files = folder.listFiles((dir, name) ->name.endsWith(".xlsx") );

		if(files==null) {
			System.out.println("Folder is empty");
		}
		Workbook outworkbook = new XSSFWorkbook();
		Sheet outsheet = outworkbook.createSheet("Bookmark Validation");
		Row outheader = outsheet.createRow(0);
		outheader.createCell(0).setCellValue("File Name");
		outheader.createCell(1).setCellValue("Cell range and Error");
		try {
			int rowidx=1;
			for(File file:files) {
				List<String> errors = new ArrayList<>();

				String filename = file.getName().toLowerCase();
				System.out.println(filename);

				Workbook wb = WorkbookFactory.create(file);
				Sheet sh = wb.getSheetAt(0);

				boolean headingFound = false;
				boolean columnHeaderFound = false;
				boolean tableBodyFound = false;
				boolean formulaFound = false;

				for( Name name:wb.getAllNames())
				{

					String bookmarkname = name.getNameName();
					String formula = name.getRefersToFormula();
					//					List<String> errors = new ArrayList<>();

					if(bookmarkname.equals("Heading")){
						headingFound=true;
					}
					if(bookmarkname.equals("ColumnHeader")){
						columnHeaderFound=true;
					}
					if(bookmarkname.equals("TableBody")) {
						tableBodyFound=true;
					}
					if(bookmarkname.equals("Formula")) {
						formulaFound=true;
					}

					if((bookmarkname.equals("Heading")) && (formula.equals("Sheet1!$B$3:$H$4"))){

						//						List<String> errors = new ArrayList<>();

						// Check B3:H3
						checkFormattingMainHeader(sh, 2, 2, 1, 7, "B3:H3", errors);

						// Check B4:H4
						checkformatingSubHeading(sh, 3, 3, 1, 7, "B4:H4", errors);

						if (errors.isEmpty()) {
							System.out.println("Heading Formatting is correct");
						} else {
							System.out.println("Formatting issues found:");

							//							for (String err : errors) {
							//								System.out.println(err);
							//								Row row = outsheet.createRow(rowidx++);
							//								row.createCell(0).setCellValue(filename);
							//								row.createCell(1).setCellValue(err);
							//							}
						}

					}
					if((bookmarkname.equals("Heading")) && (!formula.equals("Sheet1!$B$3:$H$4"))){
						errors.add(" Heading Bookmark is at wrong cells, ");
					}


					if((bookmarkname.equals("ColumnHeader"))&&formula.equals("Sheet1!$B$5:$H$6")) {
						//						List<String> errors = new ArrayList<>();

						checkformatingColumnHeader(sh, 4, 5, 1, 1, "B5:B6", errors);

						checkformatingColumnHeader(sh, 4, 5, 2, 3, "C5:D6", errors);

						checkformatingColumnHeader(sh, 4, 4, 4, 5, "E5:F5", errors);

						checkformatingColumnHeader(sh, 4, 4, 6, 7, "G5:H5", errors);

						checkformatingSubColumnHeader(sh, 5, 5, 4, 4, "E6:E6", errors);

						checkformatingSubColumnHeader(sh, 5, 5, 5, 5, "F5:F5", errors);

						checkformatingSubColumnHeader(sh, 5, 5, 6, 6, "G6:G6", errors);

						checkformatingSubColumnHeader(sh, 5, 5, 7, 7, "H7:H7", errors);

						if (errors.isEmpty()) {
							System.out.println("Column Heading Formatting is correct");
						} else {
							System.out.println("Formatting issues found:");

							//							for (String err : errors) {
							//								System.out.println(err);
							//								Row row = outsheet.createRow(rowidx++);
							//								row.createCell(0).setCellValue(filename);
							//								row.createCell(1).setCellValue(err);
							//							}
						}
					}
					if((bookmarkname.equals("ColumnHeader"))&&!formula.equals("Sheet1!$B$5:$H$6")) {
						errors.add("column Heading Bookmark is at wrong cells, ");
					}

					if((bookmarkname.equals("TableBody"))&&formula.equals("Sheet1!$B$7:$H$14")) {

						//						List<String> errors = new ArrayList<>();
						checkformattingSerialNoColumn(sh, 6, 13, 1, 1, "B7:B14",errors );
						checkformattingItemsColumn(sh, 6, 13, 2, 2, "C7:C14", errors);
						checkFormattingForDcolumn(sh, 6, 13, 3, 3, "D7:D14", errors);
						checkformattingForNumberColumns(sh, 6, 13, 4, 4, "E7:E14", errors);
						checkformattingForNumberColumns(sh, 6, 13, 5, 5, "F7:F14", errors);
						checkformattingForNumberColumns(sh, 6, 13, 6, 6, "G7:G14", errors);
						checkformattingForNumberColumns(sh, 6, 13, 7, 7, "H7:H14", errors);

						if (errors.isEmpty()) {
							System.out.println("Column Heading Formatting is correct");
						} else {
							System.out.println("Formatting issues found:");

							//							for (String err : errors) {
							//								System.out.println(err);
							//								Row row = outsheet.createRow(rowidx++);
							//								row.createCell(0).setCellValue(filename);
							//								row.createCell(1).setCellValue(err);
							//							}
						}
					}
					if((bookmarkname.equals("TableBody"))&&!formula.equals("Sheet1!$B$7:$H$14")) {
						errors.add(" Table body bookmarks at wrong cell");
					}

					if((bookmarkname.equals("Formula"))&&formula.equals("Sheet1!$E$15:$H$15")) {
						Row row = sh.getRow(14);
						Cell E15cell = row.getCell(4);
						Cell F15cell = row.getCell(5);
						Cell G15cell = row.getCell(6);
						Cell H15cell = row.getCell(7);

						if (E15cell != null && E15cell.getCellType() != CellType.FORMULA) {
							errors.add("E15 cell has no formula, ");
							//						    String formula1 = E15cell.getCellFormula();
							//						    Row row1 = outsheet.createRow(rowidx++);
							//							row1.createCell(0).setCellValue(filename);
							//							row1.createCell(1).setCellValue("E15 cell has no formula");
							//						    System.out.println("Formula : " + formula1);
						}
						if (F15cell != null && F15cell.getCellType() != CellType.FORMULA) {
							errors.add("F15 cell has no formula, ");
							//						    String formula2 = F15cell.getCellFormula();
							//						    Row row1 = outsheet.createRow(rowidx++);
							//							row1.createCell(0).setCellValue(filename);
							//							row1.createCell(1).setCellValue("F15 cell has no formula");
							//						    System.out.println("Formula : " + formula2);
						}
						if (G15cell != null && G15cell.getCellType() != CellType.FORMULA) {
							//						    String formula3 = G15cell.getCellFormula();
							//						    Row row1 = outsheet.createRow(rowidx++);
							//							row1.createCell(0).setCellValue(filename);
							//							row1.createCell(1).setCellValue("G15 cell has no formula");
							errors.add("G15 cell has no formula, ");
							//						    System.out.println("Formula : " + formula3);
						}
						if (H15cell != null && H15cell.getCellType() != CellType.FORMULA) {

							//						    String formula4 = H15cell.getCellFormula();
							//						    Row row1 = outsheet.createRow(rowidx++);
							//							row1.createCell(0).setCellValue(filename);
							//							row1.createCell(1).setCellValue("H15 cell has no formula");
							errors.add("H15 cell has no formula, ");
							//						    System.out.println("Formula : " + formula4);
						}

					}
					if((bookmarkname.equals("Formula"))&&!formula.equals("Sheet1!$E$15:$H$15")) {
						errors.add("Formula bookmarks at wrong place");
					}


					if((!bookmarkname.equals("Heading")) &&
							(!bookmarkname.equals("ColumnHeader")) &&
							(!bookmarkname.equals("TableBody")) &&
							(!bookmarkname.equals("Formula")))
					{
						errors.add(bookmarkname + " is an invalid bookmark");
					}
				}


				if(!headingFound) {
					errors.add("Heading bookmark is not present, ");
				}
				if(!columnHeaderFound) {
					errors.add("Column Heading bookmark is not present, ");
				}
				if(!tableBodyFound) {
					errors.add("Table body bookmark is not present,  ");
				}
				if(!formulaFound) {
					errors.add("formula bookmark is not present,");
				}
				if(!errors.isEmpty()) {

					String finalErrors = String.join(" | ", errors);

					Row row = outsheet.createRow(rowidx++);

					row.createCell(0).setCellValue(filename);
					row.createCell(1).setCellValue(finalErrors);
				}

				wb.close();	} 
		}



		catch (Exception e) {
			System.out.println(e.getMessage());
		}
		FileOutputStream fos = new FileOutputStream(Outpath);
		outworkbook.write(fos);
		outworkbook.close();
		System.out.println("Bookmark verification done and excel is written in output path");


	}
	public static void checkFormattingMainHeader(
			Sheet sheet,
			int firstRow,
			int lastRow,
			int firstCol,
			int lastCol,
			String rangeName,
			List<String> errors) {

		boolean merged = false;

		// Check merged region
		for (int i = 0; i < sheet.getNumMergedRegions(); i++) {

			CellRangeAddress region = sheet.getMergedRegion(i);

			if (region.getFirstRow() == firstRow &&
					region.getLastRow() == lastRow &&
					region.getFirstColumn() == firstCol &&
					region.getLastColumn() == lastCol) {

				merged = true;
				break;
			}
		}

		if (!merged) {
			errors.add(rangeName + " is not merged correctly, ");
		}

		// Check first cell style
		Row row = sheet.getRow(firstRow);

		if (row == null) {
			errors.add(rangeName + " row not found, ");
			return;
		}

		Cell cell = row.getCell(firstCol);

		if (cell == null) {
			errors.add(rangeName + " cell not found, ");
			return;
		}

		CellStyle style = cell.getCellStyle();
		Workbook workbook = sheet.getWorkbook();
		Font font = workbook.getFontAt(style.getFontIndex());

		// Bold check
		if (!font.getBold()) {
			errors.add(rangeName + " text is not bold, ");
		}

		// Underline check
		if (font.getUnderline() == Font.U_NONE) {
			errors.add(rangeName + " text is not underlined, ");
		}

		// Horizontal center alignment
		if (style.getAlignment() != HorizontalAlignment.CENTER) {
			errors.add(rangeName + " is not center aligned, ");
		}

		// Vertical middle alignment
		if (style.getVerticalAlignment() != VerticalAlignment.CENTER) {
			errors.add(rangeName + " is not middle aligned, ");
		}
		//for border check
		if (style.getBorderTop() != BorderStyle.NONE ||
				style.getBorderBottom() != BorderStyle.NONE ||
				style.getBorderLeft() != BorderStyle.NONE ||
				style.getBorderRight() != BorderStyle.NONE) {

			errors.add(rangeName + " has unexpected borders, ");
		}
		// Font size validation (example: expected 11)
		if (font.getFontHeightInPoints() != 16) {
			errors.add(rangeName + " font size is not 16, ");
		}
	}

	public static void checkformatingSubHeading(
			Sheet sheet,
			int firstRow,
			int lastRow,
			int firstCol,
			int lastCol,
			String rangeName,
			List<String> errors) {

		boolean merged = false;

		// Check merged region
		for (int i = 0; i < sheet.getNumMergedRegions(); i++) {

			CellRangeAddress region = sheet.getMergedRegion(i);

			if (region.getFirstRow() == firstRow &&
					region.getLastRow() == lastRow &&
					region.getFirstColumn() == firstCol &&
					region.getLastColumn() == lastCol) {

				merged = true;
				break;
			}
		}


		if (!merged) {
			errors.add(rangeName + " is not merged correctly, ");
		}

		// Check first cell style
		Row row = sheet.getRow(firstRow);

		if (row == null) {
			errors.add(rangeName + " row not found, ");
			return;
		}

		Cell cell = row.getCell(firstCol);

		if (cell == null) {
			errors.add(rangeName + " cell not found, ");
			return;
		}

		CellStyle style = cell.getCellStyle();

		Workbook workbook = sheet.getWorkbook();
		Font font = workbook.getFontAt(style.getFontIndex());
		// Horizontal center alignment
		if (style.getAlignment() != HorizontalAlignment.RIGHT) {
			errors.add(rangeName + " is not center aligned, ");
		}

		// Vertical middle alignment
		if (style.getVerticalAlignment() != VerticalAlignment.CENTER) {
			errors.add(rangeName + " is not middle aligned, ");
		}
		// Underline should NOT exist
		if (font.getUnderline() != Font.U_NONE) {
			errors.add(rangeName + " has unexpected underline, ");
		}

		// Italic should NOT exist
		if (font.getItalic()) {
			errors.add(rangeName + " has unexpected italic formatting, ");
		}

		// Wrap text should NOT exist
		if (style.getWrapText()) {
			errors.add(rangeName + " has unexpected wrap text, ");
		}

		// Rotation should NOT exist
		if (style.getRotation() != 0) {
			errors.add(rangeName + " has unexpected text rotation, ");
		}

		// Background color should NOT exist
		if (style.getFillPattern() != FillPatternType.NO_FILL) {
			errors.add(rangeName + " has unexpected background color, ");
		}

		// Borders should NOT exist
		if (style.getBorderTop() != BorderStyle.NONE ||
				style.getBorderBottom() != BorderStyle.NONE ||
				style.getBorderLeft() != BorderStyle.NONE ||
				style.getBorderRight() != BorderStyle.NONE) {

			errors.add(rangeName + " has unexpected borders, ");
		}



		// Font name validation
		//  if (!font.getFontName().equalsIgnoreCase("Calibri")) {
		//      errors.add(rangeName + " font is not Calibri");
		//  }


	}

	public static void checkformatingColumnHeader(
			Sheet sheet,
			int firstRow,
			int lastRow,
			int firstCol,
			int lastCol,
			String rangeName,
			List<String> errors) {

		boolean merged = false;

		// Check merged region
		for (int i = 0; i < sheet.getNumMergedRegions(); i++) {

			CellRangeAddress region = sheet.getMergedRegion(i);

			if (region.getFirstRow() == firstRow &&
					region.getLastRow() == lastRow &&
					region.getFirstColumn() == firstCol &&
					region.getLastColumn() == lastCol) {

				merged = true;
				break;
			}
		}

		if (!merged) {
			errors.add(rangeName + " is not merged correctly, ");
		}

		// Check first cell style
		Row row = sheet.getRow(firstRow);

		if (row == null) {
			errors.add(rangeName + " row not found, ");
			return;
		}

		Cell cell = row.getCell(firstCol);

		if (cell == null) {
			errors.add(rangeName + " cell not found, ");
			return;
		}

		CellStyle style = cell.getCellStyle();
		Workbook workbook = sheet.getWorkbook();
		Font font = workbook.getFontAt(style.getFontIndex());

		// Bold check
		if (!font.getBold()) {
			errors.add(rangeName + " text is not bold, ");
		}

		// Horizontal center alignment
		if (style.getAlignment() != HorizontalAlignment.CENTER) {
			errors.add(rangeName + " is not center aligned, ");
		}

		// Vertical middle alignment
		if (style.getVerticalAlignment() != VerticalAlignment.CENTER) {
			errors.add(rangeName + " is not middle aligned, ");
		}

		// Underline should NOT exist
		if (font.getUnderline() != Font.U_NONE) {
			errors.add(rangeName + " has unexpected underline, ");
		}

		// Italic should NOT exist
		if (font.getItalic()) {
			errors.add(rangeName + " has unexpected italic formatting, ");
		}

		// Wrap text should NOT exist
		if (style.getWrapText()) {
			errors.add(rangeName + " has unexpected wrap text, ");
		}

		// Rotation should NOT exist
		if (style.getRotation() != 0) {
			errors.add(rangeName + " has unexpected text rotation, ");
		}

		// Background color should NOT exist
		if (style.getFillPattern() != FillPatternType.NO_FILL) {
			errors.add(rangeName + " has unexpected background color, ");
		}

		// Borders should be exist
		if (style.getBorderTop()== BorderStyle.HAIR ||
				style.getBorderBottom() == BorderStyle.HAIR ||
				style.getBorderLeft() == BorderStyle.HAIR ||
				style.getBorderRight() == BorderStyle.HAIR) {

			errors.add(rangeName + " has no borders, ");
		}

	}

	public static void checkformatingSubColumnHeader(
			Sheet sheet,
			int firstRow,
			int lastRow,
			int firstCol,
			int lastCol,
			String rangeName,
			List<String> errors) {

		boolean merged = false;

		// Check merged region
		for (int i = 0; i < sheet.getNumMergedRegions(); i++) {

			CellRangeAddress region = sheet.getMergedRegion(i);

			if (region.getFirstRow() == firstRow &&
					region.getLastRow() == lastRow &&
					region.getFirstColumn() == firstCol &&
					region.getLastColumn() == lastCol) {

				merged = true;
				break;
			}
		}

		if (merged) {
			errors.add(rangeName + " is not merged correctly, ");
		}

		// Check first cell style
		Row row = sheet.getRow(firstRow);

		if (row == null) {
			errors.add(rangeName + " row not found, ");
			return;
		}

		Cell cell = row.getCell(firstCol);

		if (cell == null) {
			errors.add(rangeName + " cell not found, ");
			return;
		}

		CellStyle style = cell.getCellStyle();
		Workbook workbook = sheet.getWorkbook();
		Font font = workbook.getFontAt(style.getFontIndex());

		// Bold check
		if (!font.getBold()) {
			errors.add(rangeName + " text is not bold, ");
		}

		// Horizontal center alignment
		if (style.getAlignment() != HorizontalAlignment.CENTER) {
			errors.add(rangeName + " is not center aligned, ");
		}

		// Vertical middle alignment
		if (style.getVerticalAlignment() != VerticalAlignment.CENTER) {
			errors.add(rangeName + " is not middle aligned, ");
		}

		// Underline should NOT exist
		if (font.getUnderline() != Font.U_NONE) {
			errors.add(rangeName + " has unexpected underline, ");
		}

		// Italic should NOT exist
		if (font.getItalic()) {
			errors.add(rangeName + " has unexpected italic formatting, ");
		}

		// Wrap text should NOT exist
		if (style.getWrapText()) {
			errors.add(rangeName + " has unexpected wrap text, ");
		}

		// Rotation should NOT exist
		if (style.getRotation() != 0) {
			errors.add(rangeName + " has unexpected text rotation, ");
		}

		// Background color should NOT exist
		if (style.getFillPattern() != FillPatternType.NO_FILL) {
			errors.add(rangeName + " has unexpected background color, ");
		}

		// Borders should be exist
		if (style.getBorderTop()== BorderStyle.HAIR ||
				style.getBorderBottom() == BorderStyle.HAIR ||
				style.getBorderLeft() == BorderStyle.HAIR ||
				style.getBorderRight() == BorderStyle.HAIR) {

			errors.add(rangeName + " has no borders, ");
		}

	}

	public static void checkformattingSerialNoColumn(
			Sheet sheet,
			int firstRow,
			int lastRow,
			int firstCol,
			int lastCol,
			String rangeName,
			List<String> errors) {

		boolean merged = false;

		// Check merged region
		for (int i = 0; i < sheet.getNumMergedRegions(); i++) {

			CellRangeAddress region = sheet.getMergedRegion(i);

			if (region.getFirstRow() == firstRow &&
					region.getLastRow() == lastRow &&
					region.getFirstColumn() == firstCol &&
					region.getLastColumn() == lastCol) {

				merged = true;
				break;
			}
		}

		if (merged) {
			errors.add(rangeName + " is  merged correctly, ");
		}

		// Check first cell style
		Row row = sheet.getRow(firstRow);

		if (row == null) {
			errors.add(rangeName + " row not found, ");
			return;
		}

		Cell cell = row.getCell(firstCol);

		if (cell == null) {
			errors.add(rangeName + " cell not found, ");
			return;
		}

		CellStyle style = cell.getCellStyle();
		Workbook workbook = sheet.getWorkbook();
		Font font = workbook.getFontAt(style.getFontIndex());

		// Bold check
		if (font.getBold()) {
			errors.add(rangeName + " text should not be bold, ");
		}

		// Horizontal center alignment
		if (style.getAlignment() != HorizontalAlignment.CENTER) {
			errors.add(rangeName + " is not center aligned, ");
		}

		// Vertical middle alignment
		if (style.getVerticalAlignment() != VerticalAlignment.CENTER) {
			errors.add(rangeName + " is not middle aligned, ");
		}

		// Underline should NOT exist
		if (font.getUnderline() != Font.U_NONE) {
			errors.add(rangeName + " has unexpected underline, ");
		}

		// Italic should NOT exist
		if (font.getItalic()) {
			errors.add(rangeName + " has unexpected italic formatting, ");
		}

		// Wrap text should NOT exist
		if (style.getWrapText()) {
			errors.add(rangeName + " has unexpected wrap text, ");
		}

		// Rotation should NOT exist
		if (style.getRotation() != 0) {
			errors.add(rangeName + " has unexpected text rotation, ");
		}

		// Background color should NOT exist
		if (style.getFillPattern() != FillPatternType.NO_FILL) {
			errors.add(rangeName + " has unexpected background color, ");
		}

		// Borders should be exist
		if (style.getBorderTop()== BorderStyle.HAIR ||
				style.getBorderBottom() == BorderStyle.HAIR ||
				style.getBorderLeft() == BorderStyle.HAIR ||
				style.getBorderRight() == BorderStyle.HAIR) {

			errors.add(rangeName + " has no borders, ");
		}
	}

	public static void checkformattingItemsColumn(
			Sheet sheet,
			int firstRow,
			int lastRow,
			int firstCol,
			int lastCol,
			String rangeName,
			List<String> errors) {

		boolean merged = false;

		// Check merged region
		for (int i = 0; i < sheet.getNumMergedRegions(); i++) {

			CellRangeAddress region = sheet.getMergedRegion(i);

			if (region.getFirstRow() == firstRow &&
					region.getLastRow() == lastRow &&
					region.getFirstColumn() == firstCol &&
					region.getLastColumn() == lastCol) {

				merged = true;
				break;
			}
		}

		if (merged) {
			errors.add(rangeName + " is  merged correctly, ");
		}

		// Check first cell style
		Row row = sheet.getRow(firstRow);

		if (row == null) {
			errors.add(rangeName + " row not found, ");
			return;
		}

		Cell cell = row.getCell(firstCol);

		if (cell == null) {
			errors.add(rangeName + " cell not found, ");
			return;
		}

		CellStyle style = cell.getCellStyle();
		Workbook workbook = sheet.getWorkbook();
		Font font = workbook.getFontAt(style.getFontIndex());

		// Bold check
		if (font.getBold()) {
			errors.add(rangeName + " text should not be bold, ");
		}

		// Horizontal center alignment
		if (style.getAlignment() != HorizontalAlignment.LEFT) {
			errors.add(rangeName + " is not center aligned, ");
		}

		// Vertical middle alignment
		if (style.getVerticalAlignment() != VerticalAlignment.CENTER) {
			errors.add(rangeName + " is not middle aligned, ");
		}

		// Underline should NOT exist
		if (font.getUnderline() != Font.U_NONE) {
			errors.add(rangeName + " has unexpected underline, ");
		}

		// Italic should NOT exist
		if (font.getItalic()) {
			errors.add(rangeName + " has unexpected italic formatting, ");
		}

		// Wrap text should NOT exist
		if (style.getWrapText()) {
			errors.add(rangeName + " has unexpected wrap text, ");
		}

		// Rotation should NOT exist
		if (style.getRotation() != 0) {
			errors.add(rangeName + " has unexpected text rotation, ");
		}

		// Background color should NOT exist
		if (style.getFillPattern() != FillPatternType.NO_FILL) {
			errors.add(rangeName + " has unexpected background color, ");
		}

		// Borders should be exist
		//		if (style.getBorderTop()== BorderStyle.HAIR ||
		//				style.getBorderBottom() == BorderStyle.HAIR ||
		//				style.getBorderLeft() == BorderStyle.HAIR ||
		//				style.getBorderRight() == BorderStyle.HAIR) {
		//
		//			errors.add(rangeName + " has no borders");
		//		}
	}

	public static void checkFormattingForDcolumn(
			Sheet sheet,
			int firstRow,
			int lastRow,
			int firstCol,
			int lastCol,
			String rangeName,
			List<String> errors) {

		boolean merged = false;

		// Check merged region
		for (int i = 0; i < sheet.getNumMergedRegions(); i++) {

			CellRangeAddress region = sheet.getMergedRegion(i);

			if (region.getFirstRow() == firstRow &&
					region.getLastRow() == lastRow &&
					region.getFirstColumn() == firstCol &&
					region.getLastColumn() == lastCol) {

				merged = true;
				break;
			}
		}

		if (merged) {
			errors.add(rangeName + " is  merged correctly, ");
		}

		// Check first cell style
		Row row = sheet.getRow(firstRow);

		if (row == null) {
			errors.add(rangeName + " row not found, ");
			return;
		}

		Cell cell = row.getCell(firstCol);

		if (cell == null) {
			errors.add(rangeName + " cell not found, ");
			return;
		}

		CellStyle style = cell.getCellStyle();
		Workbook workbook = sheet.getWorkbook();
		Font font = workbook.getFontAt(style.getFontIndex());

		// Bold check
		if (font.getBold()) {
			errors.add(rangeName + " text should not be bold, ");
		}

		// Horizontal center alignment
		if (style.getAlignment() != HorizontalAlignment.CENTER) {
			errors.add(rangeName + " is not center aligned, ");
		}

		// Vertical middle alignment
		if (style.getVerticalAlignment() != VerticalAlignment.CENTER) {
			errors.add(rangeName + " is not middle aligned, ");
		}

		// Underline should NOT exist
		if (font.getUnderline() != Font.U_NONE) {
			errors.add(rangeName + " has unexpected underline, ");
		}

		// Italic should NOT exist
		if (font.getItalic()) {
			errors.add(rangeName + " has unexpected italic formatting, ");
		}

		// Wrap text should NOT exist
		if (style.getWrapText()) {
			errors.add(rangeName + " has unexpected wrap text, ");
		}

		// Rotation should NOT exist
		if (style.getRotation() != 0) {
			errors.add(rangeName + " has unexpected text rotation, ");
		}

		// Background color should NOT exist
		if (style.getFillPattern() != FillPatternType.NO_FILL) {
			errors.add(rangeName + " has unexpected background color, ");
		}

		// Borders should be exist
		//		if (style.getBorderTop()== BorderStyle.HAIR ||
		//				style.getBorderBottom() == BorderStyle.HAIR ||
		//				style.getBorderLeft() != BorderStyle.HAIR ||
		//				style.getBorderRight() == BorderStyle.HAIR) {
		//
		//			errors.add(rangeName + " has no borders");
		//		}
	}

	public static void checkformattingForNumberColumns(
			Sheet sheet,
			int firstRow,
			int lastRow,
			int firstCol,
			int lastCol,
			String rangeName,
			List<String> errors) {

		boolean merged = false;

		// Check merged region
		for (int i = 0; i < sheet.getNumMergedRegions(); i++) {

			CellRangeAddress region = sheet.getMergedRegion(i);

			if (region.getFirstRow() == firstRow &&
					region.getLastRow() == lastRow &&
					region.getFirstColumn() == firstCol &&
					region.getLastColumn() == lastCol) {

				merged = true;
				break;
			}
		}

		if (merged) {
			errors.add(rangeName + " is  merged correctly, ");
		}

		// Check first cell style
		Row row = sheet.getRow(firstRow);

		if (row == null) {
			errors.add(rangeName + " row not found, ");
			return;
		}

		Cell cell = row.getCell(firstCol);

		if (cell == null) {
			errors.add(rangeName + " cell not found, ");
			return;
		}

		CellStyle style = cell.getCellStyle();
		Workbook workbook = sheet.getWorkbook();
		Font font = workbook.getFontAt(style.getFontIndex());

		// Bold check
		if (font.getBold()) {
			errors.add(rangeName + " text should not be bold, ");
		}

		// Horizontal center alignment
		if (style.getAlignment() == HorizontalAlignment.CENTER) {
			errors.add(rangeName + " is center aligned, ");
		}

		// Vertical middle alignment
		if (style.getVerticalAlignment() != VerticalAlignment.CENTER) {
			errors.add(rangeName + " is not middle aligned, ");
		}

		// Underline should NOT exist
		if (font.getUnderline() != Font.U_NONE) {
			errors.add(rangeName + " has unexpected underline, ");
		}

		// Italic should NOT exist
		if (font.getItalic()) {
			errors.add(rangeName + " has unexpected italic formatting, ");
		}

		// Wrap text should NOT exist
		if (style.getWrapText()) {
			errors.add(rangeName + " has unexpected wrap text, ");
		}

		// Rotation should NOT exist
		if (style.getRotation() != 0) {
			errors.add(rangeName + " has unexpected text rotation, ");
		}

		// Background color should NOT exist
		if (style.getFillPattern() != FillPatternType.NO_FILL) {
			errors.add(rangeName + " has unexpected background color, ");
		}

		// Borders should be exist
		if (style.getBorderTop()== BorderStyle.HAIR ||
				style.getBorderBottom() == BorderStyle.HAIR ||
				style.getBorderLeft() == BorderStyle.HAIR ||
				style.getBorderRight() == BorderStyle.HAIR) {

			errors.add(rangeName + " has no borders, ");
		}
	}

}
