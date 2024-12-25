package tests;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.commons.collections4.bag.SynchronizedSortedBag;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadData {

	public static void main(String[] args) throws Exception {
		
		FileInputStream file=new FileInputStream("src\\test\\resources\\exceldata\\Registration.xlsx");
		XSSFWorkbook wb=new XSSFWorkbook(file);
		XSSFSheet sheet=wb.getSheet("Sheet1");  
		int totalrows=sheet.getLastRowNum();
		System.out.println(totalrows+" rows");
				
		int totalcells=sheet.getRow(0).getLastCellNum();
		System.out.println(totalcells+" cells");//4 --1,2,3,4
				
		for(int row=0;row<=totalrows;row++)
		{
			XSSFRow Currentrow=sheet.getRow(row);
			for(int cell=0;cell<totalcells;cell++)
			{
				XSSFCell currentcell=Currentrow.getCell(cell);
				System.out.print(currentcell.toString()+"  ");
			}
			System.out.println();
		}
		wb.close();
		file.close();
		
	}
}
