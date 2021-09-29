package ExcelProgram;

import java.io.File;
import java.io.FileInputStream;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReadData 
{

	public static void main(String[] args) 
	{
		try
		{
	
		//create the file object to open a file
		File file = new File("D:\\S T U D Y\\ExcelReadWrite.xlsx");
		
		//Create FileInputStream - to read the file
		FileInputStream fis = new FileInputStream(file);
		
		//XSSFWorkbook - will open the Excel workbook of type xlsx 
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		
		//XSSFSheet object will open a sheet at 0th index
		XSSFSheet sh = wb.getSheetAt(0);
		
		//create a for loop to iterate through rows
		for (int i=0; i<sh.getLastRowNum()+1; i++)
		{
			
			//XSSFRow - it will access the row
			XSSFRow row = sh.getRow(i);
			
			//create for loop to access the columns or cells
			for (int j=0; j<2; j++)
			{
				//XSSFCell - it is used to access the cell 
				XSSFCell cellValue = row.getCell(j);
				
				//Print the data on console
				System.out.print(cellValue +" | ");
				
			}
			System.out.println();
		}
		
		//close the fileinputstream and workbook
		fis.close();
		wb.close();
		
		
		
		
		}
		catch(Exception e)
		{
			System.out.println(e);
		}
		

	}

}
