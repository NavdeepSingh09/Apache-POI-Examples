package ExcelFileRead;

import java.io.File;
import java.io.FileInputStream;


import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelFileR2 {

	public static void main(String[] args) throws Exception {
		File filename= new File("C:\\Users\\cheem\\eclipse-workspace\\Excel File\\ExcelPOI.xlsx");
//Load the file
		FileInputStream LoadFile= new FileInputStream(filename);
// Load workbook
		XSSFWorkbook wb=new XSSFWorkbook(LoadFile);
// Load sheet- Here we are loading first sheet only
		XSSFSheet sh1= wb.getSheetAt(0);//This means first sheet in excel file
		int num= sh1.getPhysicalNumberOfRows();
		int num1=sh1.getRow(0).getPhysicalNumberOfCells();
		System.out.println(num);
		System.out.println(num1);
		
		
		for(int i=0;i<num;i++)
		{
			for(int j=0;j<num1;j++)
			{
				if(j==0)
				{
				String FirstRow= sh1.getRow(i).getCell(j).getStringCellValue();
				System.out.println("First Row="+FirstRow);
				}
				else if(j==1)
				{
				String FirstRow= sh1.getRow(i).getCell(j).getStringCellValue();
				System.out.println("Second Row="+FirstRow);	
				}
				else
				{
				String FirstRow= sh1.getRow(i).getCell(j).getStringCellValue();
				System.out.println("Third Row="+FirstRow);	
				}
			
			}
			System.out.println(" ");
		}
		wb.close();
		
		
	}

}
