package ru.unlimit.javapro.excel.labXX_pivotTable;

import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.Workbook;

public class SuperPivotTablesApp {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		
		Workbook wb = null;
		
		Name namedRange = wb.getName("");
		namedRange.setRefersToFormula("'Данные для отчета'!$A1:$J$5");

	}

}
