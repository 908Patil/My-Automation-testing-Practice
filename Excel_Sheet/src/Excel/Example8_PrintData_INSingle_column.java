package Excel;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Example8_PrintData_INSingle_column
{

	public static void main(String[] args) throws EncryptedDocumentException, IOException
	{
		
		FileInputStream file = new FileInputStream("C:\\Users\\dhira\\Desktop\\Manual Excel Sheet\\Demo Excel Sheet.xlsx");
		
		Sheet sh = WorkbookFactory.create(file).getSheet("Sheet1");
		
		int lastRowIndex = sh.getLastRowNum();
		
		for (int i=0; i<=lastRowIndex;i++) 
		{
			String value = sh.getRow(i).getCell(0).getStringCellValue();
			
			System.out.println(value+ " " );
		}
	}
}
