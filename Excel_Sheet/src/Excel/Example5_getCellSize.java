package Excel;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Example5_getCellSize 
{

	public static void main(String[] args) throws EncryptedDocumentException, IOException 
	{
		
		FileInputStream file = new FileInputStream("C:\\Users\\dhira\\Desktop\\Manual Excel Sheet\\Tset Case sample sheet.xlsx");
		
		int Value = WorkbookFactory.create(file).getSheet("Test Case Templates").getRow(0).getLastCellNum();
		
		
		System.out.println(Value);
	}
}
