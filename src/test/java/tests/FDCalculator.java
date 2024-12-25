package tests;

import java.io.FileInputStream;
import java.io.IOException;
import java.time.Duration;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.Select;

import utilities.ExcelUtils;

public class FDCalculator {

	public static void main(String[] args) throws Exception {
		ChromeOptions options = new ChromeOptions();
        options.addArguments("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36");
		WebDriver driver=new ChromeDriver();
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(30));
		driver.get("https://www.moneycontrol.com/fixed-income/calculator/state-bank-of-india-sbi/fixed-deposit-calculator-SBI-BSB001.html");
		driver.manage().window().maximize();
        String filepath="src\\test\\resources\\exceldata\\Data.xlsx";
        int rows = ExcelUtils.getRowCount(filepath, "Sheet1");
        for(int i=1;i<=rows;i++)
        {//read the data from the sheet
        	String principleamount=ExcelUtils.getCellData(filepath,"Sheet1",i,0);
        	String roi=ExcelUtils.getCellData(filepath,"Sheet1",i,1);
        	String period1=ExcelUtils.getCellData(filepath,"Sheet1",i,2);
        	String period2=ExcelUtils.getCellData(filepath,"Sheet1",i,3);
        	String frequency=ExcelUtils.getCellData(filepath,"Sheet1",i,4);
        	String expmaturedvalue=ExcelUtils.getCellData(filepath,"Sheet1",i,5);
        	//pass above data into application
        	driver.findElement(By.xpath("//input[@id='principal']")).sendKeys(principleamount);
        	driver.findElement(By.xpath("//input[@id='interest']")).sendKeys(roi);
        	driver.findElement(By.xpath("//input[@id='tenure']")).sendKeys(period1);
        	Select dropdown=new Select(driver.findElement(By.xpath("//select[@id='tenurePeriod']")));
        	dropdown.selectByVisibleText(period2);
        	Select freq=new Select(driver.findElement(By.xpath("//select[@id='frequency']")));
        	freq.selectByVisibleText(frequency);
        	
        	driver.findElement(By.xpath("//img[@src='https://images.moneycontrol.com/images/mf_revamp/btn_calcutate.gif']")).click();
        	String actmaturedvalue=driver.findElement(By.xpath("//span[@name='resp_matval']//strong")).getText();
        	if(Double.parseDouble(expmaturedvalue)==Double.parseDouble(actmaturedvalue))
        	{
        		System.out.println("test passed");
        		ExcelUtils.setCellData(filepath,"Sheet1", i,7,"Passed");
        		ExcelUtils.fillGreenColor(filepath,"Sheet1",i,7);
        	}
        	else
        	{
        		System.out.println("test failed");
        		ExcelUtils.setCellData(filepath,"Sheet1", i,7,"Failed");
        		ExcelUtils.fillRedColor(filepath,"Sheet1",i,7);
        	}
        	driver.findElement(By.xpath("//img[@class='PL5']")).click();
        }
        driver.close();
	}
}
