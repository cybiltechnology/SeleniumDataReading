package dataread;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.xslf.usermodel.XSLFSheet;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

public class ExcelDataReading {
	@Test(dataProvider="CreditCardData")
	public void ReadTestData(HashMap<Object,Object> incomingData) {
		System.out.println(incomingData.get("NameOnCard")+"; "+incomingData.get("DOB")+": "+incomingData.get("CVV")+"; "+incomingData.get("CreditCardNumber"));
		
	}
	@DataProvider(name = "CreditCardData")
	public Object[][] readData() throws Exception{
		System.out.println("test");
		FileInputStream file = new FileInputStream("C:\\Users\\Cybil\\OneDrive\\Desktop\\CreditCardTestData.xlsx");
		XSSFWorkbook workBook = new XSSFWorkbook(file);
		XSSFSheet worksheet = workBook.getSheet("TestData");
		int startRow = 1;
		int endRow = worksheet.getLastRowNum();
		int colCount = worksheet.getRow(0).getPhysicalNumberOfCells();
		System.out.println(endRow+"; "+colCount);
		Object[][] data = new Object[endRow][1];
		for(int i = startRow;i<=endRow;i++) {
			Map<Object,Object> dataMap = new HashMap<>();
			for(int j= 0;j<colCount;j++) {
				Object celldata = null;
				XSSFCell cell = worksheet.getRow(i).getCell(j);	
				System.out.println(cell);
				switch(cell.getCellType()) {
				case STRING:
					celldata = cell.getStringCellValue();
					break;
				case NUMERIC:
					celldata = cell.getNumericCellValue();
					break;
				case BOOLEAN:
					celldata = cell.getBooleanCellValue();
					break;
				case BLANK:
					celldata = null;
					break;
				default:
					break;
					
				}
				dataMap.put(worksheet.getRow(0).getCell(j).toString(), celldata);
			}
			data[i-1][0]= dataMap;
			
		}
		for(Object[] input : data) {
			for (Object map: input) {
				System.out.println(map);
				
			}
			
			
		}
		return data;
		
	}

}
