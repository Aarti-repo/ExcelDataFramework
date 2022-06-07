package Utility;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel_utils {

	static String path;
	static XSSFWorkbook workbook;
	static XSSFSheet sheet;
	
	
	public Excel_utils(String excelPath, String sheetName) {
		try {
			
			 workbook = new XSSFWorkbook(excelPath);
			 sheet = workbook.getSheet(sheetName);
		}	
		
		catch(Exception e) {
			e.printStackTrace();
		}
		
	}
	
	public static void main(String args[]) {
		getRowCount();
		getColCount();
	//	getStringCellData(1,0);
		getNumCellData(1,1);
	}
	
	
	public static int getRowCount() {
		
		int rowCount =0;
		try {
			
	//	 path = System.getProperty("user.dir");
	//	System.out.println("Excel Path :"+ path);
	//	 workbook = new XSSFWorkbook(path +"/Excel/Data.XLSX");
		// sheet = workbook.getSheet("sheet1");
			
			
		 rowCount = sheet.getPhysicalNumberOfRows();
		int allrows = sheet.getLastRowNum();
		System.out.println("All rows : "+ allrows);
		System.out.println("No. of filled rows in sheet :" + rowCount);
		
		}
		
		catch(Exception e) {
			
			e.getCause();
			System.out.println(e.getMessage());
			e.printStackTrace();
		}
		return rowCount;
	}
		
  public static int getColCount() {
	  
	  int colCount = 0;
	  
		try {
			
		colCount = sheet.getRow(0).getPhysicalNumberOfCells();
		System.out.println("No. of filled columns in sheet :" + colCount);
		
		}
		
		catch(Exception e) {
			
			e.getCause();
			System.out.println(e.getMessage());
			e.printStackTrace();
		}
		return colCount;
	}
	
	
	public  String  getStringCellData(int rowNum, int colNum) {             //use parameters to avoid hard code
		
		 String cellData = "";
		try {
			
	//	 workbook = new XSSFWorkbook(path +"/Excel/Data.XLSX");	  
	//	 sheet = workbook.getSheet("sheet1");
		 
	//	 String cellData = sheet.getRow(1).getCell(0).getStringCellValue();         // this is hard coded value like row no and column no.
		 
		 cellData = sheet.getRow(rowNum).getCell(colNum).getStringCellValue(); 
	//	 System.out.println("String Data in Cell : " + cellData);
		 
		}
		
		catch(Exception ex) {
			ex.getCause();
			System.out.println(ex.getMessage());
			ex.printStackTrace();

		}
		return cellData;
	}
	
	public static void getNumCellData(int rowNum, int colNum) {                               //use parameters to avoid hard code
		try {
			
		 double cellDataNum = sheet.getRow(rowNum).getCell(colNum).getNumericCellValue();
	//	 System.out.println(" Numeric data in cell : " + cellDataNum);
		}
		
		catch(Exception ex) {
			ex.getCause();
			System.out.println(ex.getMessage());
			ex.printStackTrace();

		}
	}
}
