package opencart;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
//TIHSFS
public class GetData {
    public static void main(String[] args) throws Exception {
        WebDriver driver = new ChromeDriver();
        driver.manage().window().maximize();
        driver.get("https://demo.opencart.com/index.php?route=account/login");

        // Load the Excel file
        FileInputStream file = new FileInputStream("C:\\Users\\bhara\\Downloads\\Opencart.xlsx");
        XSSFWorkbook wb = new XSSFWorkbook(file);
        XSSFSheet sh = wb.getSheet("Sheet1");

        int count = 1; // Starting count value

        // Get the number of rows
        int rows = sh.getLastRowNum();

        for (int row = 1; row <= rows; row++) { // Start from 1 if the first row contains headers
            XSSFRow currentRow = sh.getRow(row);
            // Get email and password from cells
            String email = currentRow.getCell(0).getStringCellValue(); // Email column
            String password = currentRow.getCell(1).getStringCellValue(); // Password column
            // Perform actions in the browser
            driver.findElement(By.xpath("//input[@id='input-email']")).clear();
            driver.findElement(By.xpath("//input[@id='input-email']")).sendKeys(email);
            driver.findElement(By.xpath("//input[@id='input-password']")).clear();
            driver.findElement(By.xpath("//input[@id='input-password']")).sendKeys(password);
            driver.findElement(By.xpath("//*[@type='submit']")).click();
            // Write the count value to the Excel sheet
            currentRow.createCell(2).setCellValue(count);
            count++;
        }
        // Save the updated Excel file
        FileOutputStream outFile = new FileOutputStream("C:\\Users\\bhara\\Downloads\\Opencart.xlsx");
        wb.write(outFile);

        // Close resources
        outFile.close();
        wb.close();
        file.close();
        driver.quit();
    }
}
