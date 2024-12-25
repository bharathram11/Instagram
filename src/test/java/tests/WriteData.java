package tests;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteData {

	public static void main(String[] args) throws Exception {
		//new file
		FileOutputStream file=new FileOutputStream("src\\test\\resources\\exceldata\\Empty.xlsx");
		//new workbook
		XSSFWorkbook wb=new XSSFWorkbook();
		//new sheet
		XSSFSheet sheet=wb.createSheet("School");
		//new row
		XSSFRow row1=sheet.createRow(0);
		 row1.createCell(0).setCellValue("Narayana");
		 row1.createCell(1).setCellValue(12345);
		 row1.createCell(2).setCellValue("Banglore");
		XSSFRow row2=sheet.createRow(1);
		 row2.createCell(0).setCellValue("Chaithanya");
		 row2.createCell(1).setCellValue(123);
		 row2.createCell(2).setCellValue("Banglore");
		XSSFRow row3=sheet.createRow(2);
		 row3.createCell(0).setCellValue("oxford");
		 row3.createCell(1).setCellValue(145);
		 row3.createCell(2).setCellValue("Chennai");
		 wb.write(file);//connection to file
		 wb.close();
		 file.close();
		 System.out.println("-----file is created---");
		
/*
		 		//new file
		 		FileOutputStream file=new FileOutputStream("src\\test\\resources\\exceldata\\Ram.xlsx");
		 		XSSFWorkbook wb=new XSSFWorkbook();
		 		XSSFSheet sh=wb.createSheet("Students");
		 		XSSFRow row1=sh.createRow(0);
		 		  row1.createCell(0).setCellValue("bharath");
		 		  row1.createCell(1).setCellValue("ram");
		 		  row1.createCell(2).setCellValue("17");
		 		  row1.createCell(3).setCellValue("2000");*/
		 		  


	}
}
