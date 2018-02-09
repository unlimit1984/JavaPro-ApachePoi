package ru.unlimit.javapro.excel.lab11_pivotTables;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataConsolidateFunction;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFPivotTable;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class PivotTableApp {

	public static void main(String[] args) throws IOException {
		XSSFWorkbook wb = new XSSFWorkbook();

		// Генерация данных
		XSSFSheet sheet = wb.createSheet("Данные");
		for (int row_i = 0; row_i < 41; row_i++) {
			Row row = sheet.createRow(row_i);
			if (row_i == 0) {

				Cell head0 = row.createCell(0);
				head0.setCellValue("Отдел");

				Cell head1 = row.createCell(1);
				head1.setCellValue("Сотрудник");
				
				Cell head2 = row.createCell(2);
				head2.setCellValue("Таб.номер");
				
				Cell head3 = row.createCell(3);
				head3.setCellValue("Должность");

				Cell head4 = row.createCell(4);
				head4.setCellValue("Город");

				Cell head5 = row.createCell(5);
				head5.setCellValue("ЗП");

			} else {
				for (int column = 0; column < 6; column++) {
					Cell cell = row.createCell(column);
					switch (column) {
					case 0:
						cell.setCellValue("Отдел" + (int) (1 + Math.random() * 5));
						break;
					case 1:
						cell.setCellValue("ФИО" + row_i);
						break;
					case 2:
						cell.setCellValue("№" + 1000 + row_i);
						break;
					case 3:
						cell.setCellValue(randomStr("Менеджер","Инженер","Тестер"));
						break;
					case 4:
						cell.setCellValue(randomStr("Москва","Питер","Саратов","Ижевск"));
						break;
					case 5:
						cell.setCellValue((int) (70000 + Math.random() * 50000));
						break;
					}
				}
			}
		}
		
		// Генерация сводной таблицы
		/******************************************/
		XSSFSheet sheetReport = wb.createSheet("Report");

//		AreaReference area = new AreaReference("A1:F41");
		CellReference ref = new CellReference("A5");

//		XSSFPivotTable pivotTable = sheetReport.createPivotTable(area, ref,sheet);
		//XSSFPivotTable pivotTable = sheet.createPivotTable(area, ref);

//		pivotTable.addReportFilter(3);		//Должность
//		pivotTable.addReportFilter(4);		//Город
//		pivotTable.addRowLabel(0);			//Отдел
//		pivotTable.addRowLabel(1);			//ФИО сотрудника
		
		//pivotTable.addDataColumn(4, false);	//Город - не работает
		
//		pivotTable.addColumnLabel(DataConsolidateFunction.SUM, 5);
//		pivotTable.addColumnLabel(DataConsolidateFunction.MIN, 5);
		//pivotTable.addColumnLabel(DataConsolidateFunction.MAX, 4);
		
		/******************************************/		
		
		OutputStream fos = new FileOutputStream("C:/ALL/tmp/Компания.xlsx");
		wb.write(fos);
		fos.close();
		wb.close();
	}
	static String randomStr(String ... arr ){
		if(arr.length==0){
			throw new RuntimeException("Нет аргументов!");
		}
		int index = (int) (Math.random()*(arr.length));
		return arr[index];
	}

}
