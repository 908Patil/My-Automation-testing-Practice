package Excel;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Example1_getStringdata 
{

	public static void main(String[] args) throws EncryptedDocumentException, IOException 
	{
		
		FileInputStream file = new FileInputStream("C:\\Users\\dhira\\Desktop\\Manual Excel Sheet\\Tset Case sample sheet.xlsx");
		
		String value = WorkbookFactory.create(file).getSheet("Test Case Templates").getRow(1).getCell(2).getStringCellValue();
		
		System.out.println(value);
	}
}
