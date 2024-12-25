package tests;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WritePractice {
    public static void main(String[] args) throws Exception {
    	
    FileOutputStream file=new FileOutputStream("src\\test\\resources\\exceldata\\Empty.xlsx");
    XSSFWorkbook wb=new XSSFWorkbook();
    XSSFSheet sh=wb.createSheet("Register");
    XSSFRow row1=sh.createRow(0);
     row1.createCell(0).setCellValue("Ravi");
     row1.createCell(1).setCellValue("1234");
     row1.createCell(2).setCellValue("banglr");
    XSSFRow row2=sh.createRow(1);
    row2.createCell(0).setCellValue("hari");
    row2.createCell(1).setCellValue("12131");
    row2.createCell(2).setCellValue("chennai");
    XSSFRow row3=sh.createRow(2);
    row3.createCell(0).setCellValue("wewe");
    row3.createCell(1).setCellValue("14");
    row3.createCell(2).setCellValue("blr");
    wb.write(file);
    wb.close();
    file.close();
    FileInputStream file1=new FileInputStream("src\\test\\resources\\exceldata\\Empty.xlsx");
    XSSFWorkbook wb1=new XSSFWorkbook(file1);
    XSSFSheet sh1=wb1.getSheet("Register");
    int rows=sh1.getLastRowNum();
    int cells=sh1.getRow(0).getLastCellNum();
    for(int i=0;i<=rows;i++)
    {
    	XSSFRow currentrow=sh1.getRow(i);
    	for(int j=0;j<=cells;j++)
    	{
    		XSSFCell currencell=currentrow.getCell(j);
    		System.out.print(currencell+" ");
    	}
    	System.out.println();
    }
    
    }
}
