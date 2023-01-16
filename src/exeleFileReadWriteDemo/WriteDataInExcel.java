package exeleFileReadWriteDemo;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;  
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteDataInExcel {

	public static void main(String[] args) throws IOException {
		XSSFWorkbook workbook= new XSSFWorkbook ();
		XSSFSheet sheet=workbook.createSheet("emp info");
		Object empdata [] [] = { {"emp id", "emp name", "emp position"},
		                          {01,       "Sabbir",    "Manager"},
		                          {02,        "Ahmed",     "Owner"},
		                          {03,         "Ripon",     "Director"},
		                          {04,         "Sar",     "CEO"},
		                          {05,         "Forida",     "Director"}
		                        
		                          } ;

		int rows=empdata.length;  //we use rows<r coz here length counting start from 1 but in excel it start from 0
		int cols=empdata[0].length; //same for column as well
		System.out.println(rows);  //rows goes sideways 
		System.out.println(cols);  // column goes top to bottom
		
		for(int r=0; r<rows ;r++) { 
			XSSFRow row=sheet.createRow(r);
			for ( int c=0; c<cols;c++) {
				XSSFCell cell=row.createCell(c);
				Object value =empdata[r] [c];
				
				if(value instanceof String)
				cell.setCellValue((String)value);
				
				if (value instanceof Integer)
					cell.setCellValue((Integer)value);
				if (value instanceof Boolean)
					cell.setCellValue((Boolean)value);
				
			}
		}
	String filepath = "C:\\eclipse-workspace\\exeleFileReadWrite\\Datafile\\FirstexcellWritingDemo.xlsx" ;
	FileOutputStream fos = new FileOutputStream (filepath);
	workbook.write(fos);
	fos.close();
	System.out.println("excel written successfully");
	}

}
