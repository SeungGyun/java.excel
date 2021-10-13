package com.ilhsk.excel;

import java.io.FileInputStream;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.junit.Test;

public class ExcelDefaultTest {

	
	public static void main(String[] args)throws Throwable {
		FileInputStream file = new FileInputStream("D:/ttt/Item_tt.xls");
		try {
			HSSFWorkbook workbook = new HSSFWorkbook(file);
			
			
			workbook.getCreationHelper().createFormulaEvaluator().evaluateAll();
			int sheetNum = workbook.getNumberOfSheets();
			
			FormulaEvaluator formulaEval = workbook.getCreationHelper().createFormulaEvaluator();
			workbook.getCreationHelper().createFormulaEvaluator().evaluateAllFormulaCells(workbook);
			for (int k = 0; k < sheetNum; k++) {
				System.out.println("Sheet Number : " + k);
				HSSFSheet sheet = workbook.getSheetAt(k);
				int rows = sheet.getPhysicalNumberOfRows();

				for (int r = 0; r < rows; r++) {
					HSSFRow row = sheet.getRow(r);
					
					int cells = row.getPhysicalNumberOfCells();

					for (short c = 0; c < cells; c++) {
						HSSFCell cell = row.getCell(c);
						CellType celltype = cell.getCellType();
						System.out.println(getValue(cell, formulaEval));
					}

				}

			}
		} catch (Exception e) {
			System.err.println(e.getMessage());
		}

	}
	
	static public String getValue(Cell cell,FormulaEvaluator formulaEval) {
		DataFormatter dataFormatter = new DataFormatter();
		switch (cell.getCellType()) {
			case FORMULA:
				if(formulaEval != null) {
					CellValue evaluate = formulaEval.evaluate(cell);
					//System.err.print(evaluate.getCellType().toString()+">> ");
					//System.out.print(evaluate.formatAsString() +" >>");
					if( evaluate != null ) {
						switch (evaluate.getCellType()) {
							case STRING : 
								return evaluate.formatAsString().replaceAll("\"", "");
							case NUMERIC:	
								
								return Long.toString((long) evaluate.getNumberValue()); 
							case BOOLEAN:
								return Boolean.toString(evaluate.getBooleanValue());
							case _NONE :
								return "";
							case BLANK : 
								return "";	
							default:
								return  evaluate.formatAsString().replaceAll("\"", "");
						}
					}else {
						return cell.getStringCellValue();
					}
				}
			case _NONE :
				return "";
			case BLANK : 
				return "";			
			default:
				return dataFormatter.formatCellValue( cell).trim();
		}
		
	}
	

}
