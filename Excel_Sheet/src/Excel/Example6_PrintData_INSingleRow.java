package Excel;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Example6_PrintData_INSingleRow 
{

	public static void main(String[] args) throws EncryptedDocumentException, IOException 
	{
		
		
		FileInputStream file = new FileInputStream("C:\\Users\\dhira\\Desktop\\Manual Excel Sheet\\Demo Excel Sheet.xlsx");
		
		Sheet sh = WorkbookFactory.create(file).getSheet("Sheet1");
		
	   int lastCellIndex = sh.getRow(0).getLastCellNum()-1;
	   
	   for(int i=0; i<= lastCellIndex; i++) 
	   {
		   String Value = sh.getRow(0).getCell(i).getStringCellValue();
		   
		   System.out.print(Value+  " ");
	   }
	
	}
	
}
