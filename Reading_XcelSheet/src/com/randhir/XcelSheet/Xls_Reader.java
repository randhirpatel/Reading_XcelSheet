package com.randhir.XcelSheet;

import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Xls_Reader {
	private static XSSFWorkbook wb=null;
	@SuppressWarnings("deprecation")
	public static void main(String[] args) {
		File myFile= new File(System.getProperty("user.dir")+"\\Data\\Data.xlsx");
		try {
			FileInputStream fs=new FileInputStream(myFile);
			wb=new XSSFWorkbook(fs);
		} catch (Exception e) {
			e.printStackTrace();
		}
		XSSFSheet mySheet= wb.getSheetAt(0);
		Iterator<Row> rowIterator=mySheet.iterator();
		while (rowIterator.hasNext()) {
			Row row =rowIterator.next();
			Iterator<Cell> cellIterator=row.cellIterator();
			while (cellIterator.hasNext()) {
				Cell cell = cellIterator.next();
				switch (cell.getCellType()) { 
				case Cell.CELL_TYPE_STRING:
				System.out.print(cell.getStringCellValue() + "\t"); 
				break; 
				case Cell.CELL_TYPE_NUMERIC:
				System.out.print(cell.getNumericCellValue() + "\t"); 
				break; 
				case Cell.CELL_TYPE_BOOLEAN: 
				System.out.print(cell.getBooleanCellValue() + "\t");
				break; 
				default :
			    }
			  }
			System.out.println("");
		}
		System.out.println("Adding a simple output change");
	}
}
