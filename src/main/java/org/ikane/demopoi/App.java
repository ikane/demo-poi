package org.ikane.demopoi;

import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Hello world!
 *
 */
public class App {
	public static void main(String[] args) {
		try {
			FileInputStream file = new FileInputStream(new File("LOADING 2015W43_complement.xlsx"));

			// Create Workbook instance holding reference to .xlsx file
			XSSFWorkbook workbook = new XSSFWorkbook(file);

			// Get first/desired sheet from the workbook
			XSSFSheet sheet = workbook.getSheetAt(0);

			// Iterate through each rows one by one
			Iterator<Row> rowIterator = sheet.iterator();

			int i = 0;
			while (rowIterator.hasNext()) {
				
				XSSFRow row = (XSSFRow) rowIterator.next();//increment the row iterator
				
				int fcell = row.getFirstCellNum();// first cell number of excel
				int lcell = row.getLastCellNum(); //last cell number of excel
				
				if(containsValue(row, fcell, lcell) == true){
					System.out.println("index:" + i++);
				}
			}

			/*
			 * while (rowIterator.hasNext()) { Row row = rowIterator.next();
			 * //For each row, iterate through all the columns Iterator<Cell>
			 * cellIterator = row.cellIterator();
			 * 
			 * while (cellIterator.hasNext()) { Cell cell = cellIterator.next();
			 * //Check the cell type and format accordingly switch
			 * (cell.getCellType()) { case Cell.CELL_TYPE_NUMERIC:
			 * System.out.print(cell.getNumericCellValue() + "t"); break; case
			 * Cell.CELL_TYPE_STRING: System.out.print(cell.getStringCellValue()
			 * + "t"); break; } } System.out.println(""); }
			 */

			file.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public static boolean containsValue(XSSFRow row, int fcell, int lcell) {
		boolean flag = false;
		for (int i = fcell; i < lcell; i++) {
			if (StringUtils.isEmpty(String.valueOf(row.getCell(i))) == true
					|| StringUtils.isWhitespace(String.valueOf(row.getCell(i))) == true
					|| StringUtils.isBlank(String.valueOf(row.getCell(i))) == true
					|| String.valueOf(row.getCell(i)).length() == 0 || row.getCell(i) == null) {
			} else {
				flag = true;
			}
		}
		return flag;
	}
}
