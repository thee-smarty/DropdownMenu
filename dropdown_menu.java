package miniproject;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.time.Duration;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.support.ui.Select;

public class dropdown_menu 
{
	public static void main(String[] args) throws InterruptedException, IOException 
	{
		System.setProperty("webdriver.edge.driver", "C:\\Users\\2389221\\eclipse-workspace\\Selenium\\Driver\\msedgedriver.exe");
		
		InputStream fileIn = new FileInputStream("C:\\Users\\2389221\\eclipse-workspace\\Selenium\\src\\miniproject\\DataInputs.xls");
		Workbook wbr = WorkbookFactory.create(fileIn);
		Sheet sheet1 = wbr.getSheetAt(0);
		Row row1 = sheet1.getRow(0);
		// Launch the browser.
		WebDriver driver=new EdgeDriver();
		driver.manage().window().maximize();
		// Open the URL: https://mail.rediff.com/cgi-bin/login.cgi
		driver.get("https://mail.rediff.com/cgi-bin/login.cgi");
		// Click on "Create a new account link"
		driver.findElement(By.xpath("//a[@title='Create new Rediffmail account']")).click();
		// Enter the name as “Kamal” and Rediff mail as “kamal1234” and click on “Check Availability” button.
		driver.findElement(By.xpath("//input[1]")).sendKeys(row1.getCell(0).toString());
		driver.findElement(By.xpath("//tr[7]//input[1]")).sendKeys(row1.getCell(1).toString());
		driver.findElement(By.xpath("//input[@type='button']")).click();
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(10));
		// Select one of the auto suggested mail options
		driver.findElement(By.xpath("//div[@id='recommend_text']//input")).click();
		// Enter password as “Kamal@1234”.
		driver.findElement(By.xpath("//tr[9]//input")).sendKeys(row1.getCell(2).toString());
		driver.findElement(By.xpath("//tr[11]//input")).sendKeys(row1.getCell(2).toString());		
		// Select the Checkbox “Click if you don't have an alternate ID”
		driver.findElement(By.xpath("//tr[15]//input")).click();		
		// Select the Date of Birth “20-Jun-2000”
		driver.findElement(By.xpath("//tr[22]//select[1]")).sendKeys(row1.getCell(3).toString());
		driver.findElement(By.xpath("//tr[22]//select[2]")).sendKeys(row1.getCell(4).toString());
		driver.findElement(By.xpath("//tr[22]//select[3]")).sendKeys(row1.getCell(5).toString());
		fileIn.close();
		// Click on Country dropdown list box
		WebElement country= driver.findElement(By.xpath("//select[@id='country']"));
		country.click();
		// Fetch all the available country names and display on console
		Select select= new Select(country);
		List<WebElement> list = select.getOptions();
		int i=1;
		try {
			OutputStream fileOut = new FileOutputStream("C:\\Users\\2389221\\eclipse-workspace\\Selenium\\src\\miniproject\\DataOutputs.xls");
			Workbook wbw = new HSSFWorkbook();
			Sheet sheet2 = wbw.createSheet("outputs");
			sheet2.createRow(0).createCell(0).setCellValue("Country");
			
			for(WebElement x : list){
				System.out.println(x.getText());
				sheet2.createRow(i).createCell(0).setCellValue(x.getText());
				i++;
			}
			wbw.write(fileOut);
			wbw.close();
		}catch(FileNotFoundException e) {
			System.out.println("File not found exception "+e);
		}
		// Print the total count of countries.
		System.out.println("Total count of countries : "+ select.getOptions().size());
		// Select Country with Visible Text=”India” 
		select.selectByVisibleText("India");
		// Print the name of country selected on console
		String text = select.getFirstSelectedOption().getText();
		System.out.println("Country selected : "+text);
		// Validate the selected country against the expected i.e, India
		System.out.println(text.equals("India")?"Valid Country":"Not a valid Country");
		// Close the browser
		driver.quit();
	}	
}
