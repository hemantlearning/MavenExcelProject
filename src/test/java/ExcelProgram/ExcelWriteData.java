package ExcelProgram;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
public class ExcelWriteData 
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
		
		//count will give us the last line that  is displayed
		int rowcount = sh.getLastRowNum()+1;
		
		//data that to be written on the Excel sheet
		String[] value = {"4","Four"};
		
		//XSSFRow - will access that particular row
		XSSFRow row = sh.createRow(rowcount);
		
		for(int i=0; i<2; i++)
		{
			//XSSFCell - will access the exact cell location
			XSSFCell cell = row.createCell(i);
			
			//This will take the data from the variable to write to the cell location
			cell.setCellValue(value[i]);
			
			
		}
		
		//It will collect all the information to write data into the cell
		FileOutputStream fos = new FileOutputStream(file);
		
		//This will actually write the data into the cell which is collected in fos
		wb.write(fos);
		
		
		fis.close();
		fos.close();
		wb.close();
		
		
		
		}
		catch(Exception e) {
			System.out.println(e.getMessage());
		}
		

	}

}
