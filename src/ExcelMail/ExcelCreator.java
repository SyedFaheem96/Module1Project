package ExcelMail;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

public class ExcelCreator {

	private String filename;
	private String highlightedFile;

	public void start(String filename, String highlighteFile) throws Exception {
		this.filename = filename;
		this.highlightedFile = highlighteFile;
		create();
		highlightRows();
	}

	private void create() throws Exception {

		System.out.println("Creating Excel file...");

		Workbook workbook = new HSSFWorkbook();

		Sheet sheet = workbook.createSheet("HITBATCH2");

		Row row = sheet.createRow(0);
		Cell cell = row.createCell(0);
		sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 3)); // merge
		cell.setCellValue("Covid-19 RECORD ");
		// TODO : center it in the cell

		row = sheet.createRow(1);
		row.createCell(0).setCellValue("Patient name");
		row.createCell(1).setCellValue("Vaccinated Record 1st dose");
		row.createCell(2).setCellValue("Address");
		row.createCell(3).setCellValue("Phone ");

		row = sheet.createRow(4);
		row.createCell(0).setCellValue("Suhail");
		row.createCell(1).setCellValue("No");
		row.createCell(2).setCellValue("No 2 sdasdsadsaasf a, fafasfda");
		row.createCell(3).setCellValue("9940201740");
		// TODO : use the row object and highlight it
		// TODO : if the value is 'No' then highlight the row

		row = sheet.createRow(2);
		row.createCell(0).setCellValue("Syed Faheem");
		row.createCell(1).setCellValue("yes");
		row.createCell(2).setCellValue("No 1/2 kodungaiyur chennai");
		row.createCell(3).setCellValue("9940201750");

		row = sheet.createRow(3);
		row.createCell(0).setCellValue("Aslam");
		row.createCell(1).setCellValue("No");
		row.createCell(2).setCellValue("2 L.F Road. Kanyakumari");
		row.createCell(3).setCellValue("9940201740");
		// TODO : use the row object and highlight it
		// TODO : if the value is 'No' then highlight the row

		// set auto column width
		sheet.autoSizeColumn(0);
		sheet.autoSizeColumn(1);
		sheet.autoSizeColumn(2);
		sheet.autoSizeColumn(3);

		File f = new File(filename);

		FileOutputStream fos = new FileOutputStream(f);
		workbook.write(fos);

		fos.close();
		workbook.close();
		System.out.print("Excel Created and Data added");
	}

	private void highlightRows() throws IOException {

		Workbook workbook = new HSSFWorkbook(new FileInputStream(filename));
		Sheet sheet = workbook.getSheetAt(0);

		int rows = sheet.getPhysicalNumberOfRows();

		// highlight cell
		CellStyle style = workbook.createCellStyle();
		style.setFillForegroundColor(IndexedColors.RED.getIndex());
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

		Row row;
		String vaccinated;
		for (int i = 2; i < rows; i++) {
			row = sheet.getRow(i);
			vaccinated = row.getCell(1).getStringCellValue(); // the vaccine cell

			if (vaccinated.equalsIgnoreCase("NO"))
				row.getCell(0).setCellStyle(style);
		}

		workbook.write(new FileOutputStream(highlightedFile));
		workbook.close();

		System.out.println("\t\tHighlighted specific data");
	}

}