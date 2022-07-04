package StepDefinations;

import java.io.File;
import java.io.FileOutputStream;
import java.util.Iterator;
import java.util.List;
import java.util.Set;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.AfterMethod;

import io.cucumber.java.en.And;
import io.cucumber.java.en.Given;
import io.cucumber.java.en.Then;
import io.cucumber.java.en.When;

public class CCS {

public WebDriver driver;

@Given("^GreenKart page details are display")
public void greenkart_page_details_are_display() throws Throwable{
	System.setProperty("webdriver.chrome.driver", "C:\\Users\\0024C1744\\Downloads\\Training Material\\API_Udemy\\chromedriver_win32\\chromedriver.exe");
	driver=new ChromeDriver();
	driver.navigate().to("https://rahulshettyacademy.com/seleniumPractise/#/");
	Thread.sleep(2000);
	driver.manage().window().maximize();
	Thread.sleep(3000);	
}
@When("User click on Top deals link")
public void user_click_on_top_deals_link() throws Throwable{
	driver.findElement(By.linkText("Top Deals")).click();
	Thread.sleep(2000);
	Set<String> s1=driver.getWindowHandles();System.setProperty("webdriver.chrome.driver", "C:\\Users\\0024C1744\\Downloads\\Training Material\\API_Udemy\\chromedriver_win32\\chromedriver.exe");
	Iterator<String> i1=s1.iterator();
	String ParentWindow=i1.next();
	String ChildWindow=i1.next();
	driver.switchTo().window(ChildWindow);
	Thread.sleep(2000);	
}


@Then("New window will open")
public void new_window_will_open()  throws Throwable{
	String NW=driver.getCurrentUrl();
	String sub="offers";
	System.out.println(NW.contains(NW));
}

@And("Top deals details are display")
public void top_deals_details_are_display() throws Throwable{
	//driver.findElement(By.xpath(null))
	//driver.findElement(By.xpath("//span[@=\"root\"]/div/div/div/div/div/div/table/thead/tr/th[1]/span[2]")).click();
	Thread.sleep(2000);
	WebElement wb=driver.findElement(By.xpath("/html/body/div/div/div/div/div/div/div/table"));
	
	List<WebElement> rows=wb.findElements(By.tagName("tr"));
	XSSFWorkbook workbook=new XSSFWorkbook();
	XSSFSheet sheet=workbook.createSheet();
	
	for(int i=0;i<rows.size();i++)
	{
		
		List<WebElement> cols=rows.get(i).findElements(By.tagName("td"));
		Row ro=sheet.createRow(i);
		for(int j=0;j<cols.size();j++)
		{
			String cellText=cols.get(j).getText();
			ro.createCell(j).setCellValue(cellText);
		}
		
	}
	String path="C:\\Users\\0024C1744\\Downloads\\Training Material\\Applications\\demo.xlsx";
	File f=new File(path);
		FileOutputStream FO=new FileOutputStream(f);
		workbook.write(FO);
		FO.close();
		
	System.out.println("Data added successfully");
	
	

    // Write code here that turns the phrase above into concrete actions
	driver.quit();
	Runtime.getRuntime().exec("taskkill /F /IM chromedriver.exe /T");
	
}


}


