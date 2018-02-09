package ru.unlimit.javapro.excel.lab09_xlsx;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class XLSXApp {

	public static void main(String[] args) throws IOException {
	    Workbook wb = new XSSFWorkbook();
	    FileOutputStream fileOut = new FileOutputStream("reports/workbook.xlsx");

		Sheet sheet0=wb.createSheet("Publishers");
		Row row = sheet0.createRow(3);
		Cell cell = row.createCell(4);
		cell.setCellValue("O'Reilly");
	    
	    
	    wb.write(fileOut);
	    fileOut.close();


	}

}
