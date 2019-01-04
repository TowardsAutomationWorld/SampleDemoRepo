package com.mani.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelHandlingUsingApachePoi {

	public static void main(String[] args) throws IOException {

		File file = new File("G:\\Git_Repository\\ExcelSamples\\data\\AllContacts.xlsx");
		FileInputStream fis = new FileInputStream(file);

		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		XSSFSheet spreadSheet = workbook.getSheet("Contacts");
		int activeCellRowNumber = spreadSheet.getActiveCell().getRow();

		int rowCount = spreadSheet.getLastRowNum() - spreadSheet.getFirstRowNum();

		for (int i = 0; i < rowCount + 1; i++) {
			XSSFRow row = spreadSheet.getRow(i);
			// Create a loop to print cell values in a row
			for (int j = 0; j < row.getLastCellNum(); j++) {
				DataFormatter formatter = new DataFormatter();
				// Print Excel data in console
				formatter.formatCellValue(spreadSheet.getRow(i).getCell(j));
				String cell = formatter.formatCellValue(spreadSheet.getRow(i).getCell(j));
				System.out.println(cell);
			}
			System.out.println();
		}
	}
}
