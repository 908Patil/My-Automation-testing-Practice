
package Excel;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Example11_Verify_CellType 
{

	public static void main(String[] args) throws EncryptedDocumentException, IOException 
	{
		
		FileInputStream file = new FileInputStream("C:\\Users\\dhira\\Desktop\\Manual Excel Sheet\\Demo Excel Sheet.xlsx");
		
		Cell cellinfo = WorkbookFactory.create(file).getSheet("Sheet2").getRow(0).getCell(0);
		
		CellType CellType = cellinfo.getCellType();
		System.out.println(CellType);
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
	}
}
