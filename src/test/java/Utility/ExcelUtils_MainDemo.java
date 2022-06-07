package Utility;

import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

public class ExcelUtils_MainDemo {

/*	public static void main(String[] args) {

	String	path = System.getProperty("user.dir");
		System.out.println("Excel Path :"+ path);
		Excel_utils excel = new Excel_utils(path + "/Excel/Data.XLSX", "sheet1");
		
		
	//	excel.getRowCount();
	//	excel.getColCount();
	//	excel.getNumCellData(3,1);
	//	excel.getStringCellData(1, 0);
		
		testData(path + "/Excel/Data.XLSX", "sheet1");
	}
	
	*/
	
	@Test(dataProvider = "Test1")
	public void test1(String username, String password) {
		
		System.out.println(username+ " | "+password );
	}

 @DataProvider (name = "Test1")
 
 public Object[][] getData(){
	 
	 String	excelPath = "C://Users/Aarti/eclipse-workspace/ExcelData_Framework/Excel/Data.XLSX";
		
		Object data[][] = testData(excelPath, "sheet1");
		
		return data;
 }
 
	public static Object[][] testData(String excelPath, String sheetName) {
		
		Excel_utils excel = new Excel_utils(excelPath, sheetName);
		
		
		
		int rowCount = excel.getRowCount();
		int colCount = excel.getColCount();
		
		Object data[][] = new Object[rowCount-1][colCount]; 
		
		for(int i = 1; i<rowCount; i++) {
			for(int j = 0; j<colCount; j++) {
				
				String cellData = excel.getStringCellData(i, j);
				System.out.print(cellData + " | ");
				data[i-1][j] = cellData;
			}
			System.out.println();
		}
		return data;
	}
}
