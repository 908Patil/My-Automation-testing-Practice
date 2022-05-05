package Excel;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Example4_getRowSize 
{

	public static void main(String[] args) throws EncryptedDocumentException, IOException 
	{
		
		FileInputStream file = new FileInputStream("C:\\Users\\dhira\\Desktop\\Manual Excel Sheet\\Tset Case sample sheet.xlsx");
		
		int Value = WorkbookFactory.create(file).getSheet("Test Case Templates").getLastRowNum()+1;
		
		
		System.out.println(Value);
	}
}
