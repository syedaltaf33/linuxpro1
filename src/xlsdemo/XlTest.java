package xlsdemo;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class XlTest {

	public static void main(String[] args) throws IOException 
	{
		FileInputStream fi = new FileInputStream("H:\\althaf.xlsx");
		Workbook wb = new XSSFWorkbook(fi);
		
		Sheet ws = wb.getSheet("Sheet2");
		Row r = ws.getRow(1);
		Cell c = r.createCell(3);
		c.setCellValue("ALTHAF   ");
		
		FileOutputStream fo = new FileOutputStream("H:\\\\althafg.xlsx");
		wb.write(fo);
		wb.close();
		fi.close();
		fo.close();	
		
		
	
		

	
	
	
	
	
	}

}
