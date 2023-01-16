package exeleFileReadWriteDemo;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReadDemo {

	public static void main(String[] args) throws IOException {
		String excelfilepath = "C:\\eclipse-workspace\\exeleFileReadWrite\\Datafile\\Untitled spreadsheet (1).xlsx";
		                                                         //String excelfilepath = ".\\Datafile\\Untitled spreadsheet.xlsx";
		FileInputStream fis = new FileInputStream(excelfilepath); 

		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		// XSSFSheet sheet =workbook.getSheet("sheet1");  // this is another method to get the sheet
		XSSFSheet sheet = workbook.getSheetAt(0);
		
		//to get all data we can use for loop
	int rows =sheet.getLastRowNum();
	int cols =sheet.getRow(1).getLastCellNum();
	
	
	
	 for (int r=0;r<=rows;r++) {
		XSSFRow row=sheet.getRow(r);
		
		for (int c=0;c<cols;c++)
		{
			XSSFCell cell=row.getCell(c);
			
			switch (cell.getCellType())
			{
			case STRING: System.out.print(cell.getStringCellValue()) ; break;
			case NUMERIC:System.out.print(cell.getNumericCellValue()); break;
			case BOOLEAN: System.out.print(cell.getBooleanCellValue());break;
			
			}
			System.out.print("    |    ");
			
		}
		System.out.println();  
	} 
	
	//another method for same reading
	
	/* Iterator iterator=sheet.iterator();
	while (iterator.hasNext()) {
		XSSFRow row=(XSSFRow) iterator.next();
		Iterator celliterator=row.cellIterator();
		while (celliterator.hasNext()) {}
		XSSFCell cell=(XSSFCell) celliterator.next();
		
		switch (cell.getCellType())
		{
		case STRING: System.out.print(cell.getStringCellValue()) ; break;
		case NUMERIC:System.out.print(cell.getNumericCellValue()); break;
		case BOOLEAN: System.out.print(cell.getBooleanCellValue());break;
		default:
			break;
		
		}
		System.out.print("    |    ");
		
	}
System.out.println();  */
	
	
	}

}
