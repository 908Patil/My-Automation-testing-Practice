package Excel;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Example10_getCelltype2 
{
public static void main(String[] args) throws EncryptedDocumentException, IOException 
{
	
	
	FileInputStream file = new FileInputStream("C:\\Users\\dhira\\Desktop\\Manual Excel Sheet\\Demo Excel Sheet.xlsx");
	
	Sheet sh = WorkbookFactory.create(file).getSheet("Sheet2");
	Cell cellinfo = sh.getRow(0).getCell(2);
	
	CellType TypeOfcell = cellinfo.getCellType();
	
	System.out.println(TypeOfcell);
	
	
	if(TypeOfcell == CellType.STRING) 
	{
		
		String value  = cellinfo.getStringCellValue();
		
		System.out.println(value);
	}
	
	else if(TypeOfcell == CellType.NUMERIC) 
	{
		
		double value = cellinfo.getNumericCellValue();
		
		System.out.println(value);
	}
	
	else if(TypeOfcell == CellType.BOOLEAN) 
	{
		boolean value = cellinfo.getBooleanCellValue();
		
		System.out.println(value);
	}
	
}
}
