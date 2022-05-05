package Excel;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Example3_getBooleanValue 
{

	public static void main(String[]args) throws EncryptedDocumentException, IOException 
	{
		FileInputStream file = new FileInputStream("C:\\Users\\dhira\\Desktop\\Manual Excel Sheet\\Tset Case sample sheet.xlsx");
		
		
		boolean Value = WorkbookFactory.create(file).getSheet("Test Case Templates").getRow(2).getCell(4).getBooleanCellValue();
		
		System.out.println(Value);
	}
}
