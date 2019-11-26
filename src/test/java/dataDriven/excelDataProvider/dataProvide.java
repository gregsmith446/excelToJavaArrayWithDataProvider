package dataDriven.excelDataProvider;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

public class dataProvide {
	
	DataFormatter formatter = new DataFormatter();
	
	@Test(dataProvider="driveTest")
	public void testCaseData(String greeting, String communication, String id)
	{
		System.out.println(greeting + communication + id);
	}
	
	@DataProvider(name="driveTest")
	public Object[][] getData() throws IOException
	{	
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
		Object data [][] = new Object[rowCount - 1][colCount];
			
		// Object[][] data = { {"hello", "text", "1"}, {"bye", "message", "143"}, {"solo", "call", "453"} };
		for (int i = 0; i < rowCount - 1; i++)
		{
			row = sheet.getRow(i + 1);			
			for (int j = 0; j < colCount; j++)
			{
				XSSFCell cell = row.getCell(j);
				data[i][j] = formatter.formatCellValue(cell);
			}
		}
		return data;
	}
}

















