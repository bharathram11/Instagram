package opencart;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class StoreData {

	public static void main(String[] args) throws Exception {
		// TODO Auto-generated method stub
		FileOutputStream file=new FileOutputStream("C:\\Users\\bhara\\Downloads\\Opencart.xlsx");
		XSSFWorkbook wb=new XSSFWorkbook();
		XSSFSheet sh=wb.createSheet("Sheet2");
		XSSFRow row1=sh.createRow(0);
		  row1.createCell(0).setCellValue("hello");
		  row1.createCell(1).setCellValue("namste");
		  row1.createCell(2).setCellValue("f8888");
		wb.write(file);
		wb.close();
		file.close();

	}

}
