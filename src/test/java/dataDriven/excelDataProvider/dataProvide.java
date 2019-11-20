package dataDriven.excelDataProvider;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

public class dataProvide {
	
	@DataProvider(name="driveTest")
	public Object[][] getData() throws IOException
	{
		// Object[][] data = { {"hello", "text", "1"}, {"bye", "message", "143"}, {"solo", "call", "453"} };
	
		// every row of excel = 1 array in multi-dimensional array
		FileInputStream fis = new FileInputStream("C:\\Users\\gregs\\eclipse-workspace\\excelDataProvider\\excelDriven.xlsx"); 
		// excel instance, takes the excel file as an arg
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		// return type of below method is XSSF sheet
		XSSFSheet sheet = wb.getSheetAt(0);
		// how to get the # of rows
		int rowCount = sheet.getPhysicalNumberOfRows();
		XSSFRow row = sheet.getRow(0);	
		// how to get columnCount (the last cell in a row!)
		int colCount = row.getLastCellNum();
		// declare multi-dimensional array with above variables 
		Object Data [][] = new Object[rowCount - 1][colCount];
		
		return Data;
	}
	
	@Test(dataProvider="driveTest")
	public void testCaseData(String greeting, String communication, String id)
	{
		System.out.println(greeting + communication + id);
	}
	
	
	
	
	
}
