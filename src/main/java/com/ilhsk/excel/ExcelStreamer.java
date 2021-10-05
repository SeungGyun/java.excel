
package com.ilhsk.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import com.monitorjbl.xlsx.StreamingReader;


public class ExcelStreamer {

	public static void main(String[] args) throws Exception {
		try (InputStream is = new FileInputStream(new File("d:/tt.xlsx")); Workbook workbook = StreamingReader.builder().rowCacheSize(100).bufferSize(4096).open(is)) {
			for (Sheet sheet : workbook) {
				System.out.println(sheet.getSheetName());
				for (Row r : sheet) {
					for (Cell c : r) {
						System.out.println(c.getStringCellValue());
					}
				}
			}
		}
	}

}
