package tests;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.util.Scanner;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteDynamicdata2 {

	public static void main(String[] args) throws Exception {
		//new file
		FileOutputStream file=new FileOutputStream("src\\test\\resources\\exceldata\\EmptyData.xlsx");
		//new workbook
		XSSFWorkbook wb=new XSSFWorkbook();
		//new sheet
		XSSFSheet sheet=wb.createSheet("DynamicData");
		Scanner sc=new Scanner(System.in);
		System.out.println("enter how many rows you need");
		int rowsno=sc.nextInt();
		System.out.println("enter how many cells you need in every row");
		int cellsno=sc.nextInt();
		for(int r=0;r<=rowsno;r++)
		{
			XSSFRow currentrow=sheet.createRow(r);
			for(int c=0;c<cellsno;c++)
			{
				XSSFCell currentcell=currentrow.createCell(c);
				currentcell.setCellValue(sc.next());//here the user knows which data type he is gonna give
			}
		}
		 wb.write(file);//connection to file
		 wb.close();
		 file.close();
		 System.out.println("-----file is created---");
		
		

	}
}
