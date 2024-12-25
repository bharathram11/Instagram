package tests;
import java.io.FileOutputStream;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
public class WriteRandom {

	public static void main(String[] args) throws Exception {
		//new file
		FileOutputStream file=new FileOutputStream("src\\test\\resources\\exceldata\\Random.xlsx");
		//new workbook
		XSSFWorkbook wb=new XSSFWorkbook();
		//new sheet
		XSSFSheet sheet=wb.createSheet("School");
		//new row
		XSSFRow row=sheet.createRow(4);
		XSSFCell cell=row.createCell(7);
		cell.setCellValue("bharath");
		wb.write(file);//connection to file
		wb.close();
		file.close();
		System.out.println("-----file is created---");
	}
}
