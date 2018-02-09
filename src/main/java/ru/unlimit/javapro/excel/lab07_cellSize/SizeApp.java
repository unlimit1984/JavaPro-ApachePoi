package ru.unlimit.javapro.excel.lab07_cellSize;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

public class SizeApp {

	public static void main(String[] args) throws IOException {
		
		Workbook wb = new HSSFWorkbook();
		Sheet sheet=wb.createSheet("Лист_01");
		Row row = sheet.createRow(0);
		Cell cell = row.createCell(0);
		cell.setCellValue("Новая ячейка");
		//sheet.setColumnWidth(3, 5000);
		//sheet.autoSizeColumn(0);
		row.setHeightInPoints(15);
		
		sheet.addMergedRegion(new CellRangeAddress(0, 5, 0, 2));
		
		//FileOutputStream fos = new FileOutputStream("C:/ALL/tmp/Размер_Ячейки.xls");
		FileOutputStream fos = new FileOutputStream("Размер_Ячейки.xls");
		wb.write(fos);
		fos.close();	

	}

}
