package com.galaxy.migration;

import java.io.ByteArrayOutputStream;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.PrintWriter;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.HashSet;
import java.util.Iterator;
import java.util.Set;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ParseExcelDate {

	public static void readXLSXFile() throws IOException {

//		final int WORKSPACE_MODULE_ID = workspaceModuleId;

		InputStream excelFileToRead = new FileInputStream("src/main/resources/test.xlsx");

		FileWriter fileWriter = new FileWriter("src/main/resources/generatedProductSQL.txt");
		PrintWriter printWriter = new PrintWriter(fileWriter);

		XSSFWorkbook wb = new XSSFWorkbook(excelFileToRead);

		XSSFSheet sheet = wb.getSheetAt(0);
		XSSFRow row;
		XSSFCell cell;

		Iterator<Row> rows = sheet.rowIterator();
//		rows.next();
		int rowCount = 0;
		int succeededInsertCount = 0;
		int faildInsertCount = 0;
		int postMarktingProductsCount = 0;
		StringBuilder lsqQPPVSelect = new StringBuilder();
		Set<String> lsrQPPVEmailsSet = new HashSet<>();

		XSSFWorkbook wbCorrect = new XSSFWorkbook();
		Sheet sheetCorrect = wbCorrect.createSheet("dates");
		SimpleDateFormat sdf = new SimpleDateFormat("M/dd/yyyy");
		new_row: while (rows.hasNext()) {

			row = (XSSFRow) rows.next();
			System.out.println("bbbbbbbbbbbbbbbbbbbbbbbbbbbb" + row.getRowNum());
			Iterator<Cell> cells = row.cellIterator();
			cell = row.getCell(0);
			String v = "";
			try {

				v = sdf.format(cell.getDateCellValue());
			} catch (Exception e) {
				v = cell.getStringCellValue();
				v = v.replaceAll("\\\\", "/");
				v = v.replaceAll("_", "/");
				String[] x = v.split("/");
				if (x.length == 3) {

					int y = Integer.parseInt(x[0]);

					if (y > 12 && y < 32) {
						x[0] = x[1];
						x[1] = String.valueOf(y);
					}

					v = String.join("/", x);

				}
			}

			Row rowCorrect = sheetCorrect.createRow(rowCount);
			Cell cellCorrect = rowCorrect.createCell(0);

//			SimpleDateFormat sdf = new SimpleDateFormat("dd-M-yyyy hh:mm:ss");
//			String dateInString = "15-10-2015 10:20:56";
//			Date date = sdf.parse(dateInString);
//			System.out.println(date);

			try {
				Date d = sdf.parse(v);

				Calendar cal = Calendar.getInstance();
				cal.setTime(d);
				cellCorrect.setCellValue(cal);
			} catch (Exception e) {
				cellCorrect.setCellValue(v);
			}

//			printWriter.printf( DateUtil.isCellDateFormatted(cellCorrect)+"");
			printWriter.printf(v + "\n");

//			while (cells.hasNext()) {
//				cell = (XSSFCell) cells.next();
//				try {
//					
//					System.out.println(cell.getDateCellValue());
//				}catch (Exception e) {
//					System.out.println("errorrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrr");
//				}
//				
//			}
			rowCount++;
		}

		// Write the output to a file
		try (ByteArrayOutputStream fileOut = new ByteArrayOutputStream()) {
			wbCorrect.write(fileOut);
			try (OutputStream outputStream = new FileOutputStream("thefilename.xlsx")) {
				fileOut.writeTo(outputStream);
			}
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();

		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
//		printWriter.
		printWriter.close();
		wb.close();
		System.out.println(">>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>" + rowCount);
	}

}
