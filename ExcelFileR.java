package ExcelFileRead;

import java.io.File;
import java.io.FileInputStream;


import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelFileR {

	public static void main(String[] args) throws Exception {
		File filename= new File("C:\\Users\\cheem\\eclipse-workspace\\Excel File\\ExcelPOI.xlsx");
		//Load the file
		FileInputStream LoadFile= new FileInputStream(filename);
		// Load workbook
		XSSFWorkbook wb=new XSSFWorkbook(LoadFile);
		   
		   // Load sheet- Here we are loading first sheetonly
		XSSFSheet sh1= wb.getSheetAt(0);//This means first sheet in excel fle
		String FirstRow= sh1.getRow(0).getCell(0).getStringCellValue();
		System.out.println("My first row value="+FirstRow);
		System.out.println("Second row="+sh1.getRow(0).getCell(1).getStringCellValue());
		wb.close();
		
		
	}

}
