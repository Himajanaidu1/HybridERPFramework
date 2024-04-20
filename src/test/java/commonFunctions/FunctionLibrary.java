package commonFunctions;

import static org.testng.Assert.assertEquals;

import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.util.Date;
import java.util.Properties;

import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.Assert;
import org.testng.Reporter;
public class FunctionLibrary {
	public static Properties conpro;
	public static WebDriver driver;
	//method for launching browser
	public static WebDriver startBrowser()throws Throwable
	{
		conpro=new Properties();
		//load file
		conpro.load(new FileInputStream("./propertyFiles/Environment.properties"));
		if (conpro.getProperty("Browser").equalsIgnoreCase("chrome"))
		{
			driver=new ChromeDriver();
			driver.manage().window().maximize();
		}
		else if(conpro.getProperty("Browser").equalsIgnoreCase("Firefox"))
		{
			driver=new FirefoxDriver();
		}
		else
		{
			Reporter.log("Browser key value is not matching",true);
		}
		return driver;
	}
	//method for launching url
	public static void openUrl()
	{
		driver.get(conpro.getProperty("Url"));

	}
	//method for webelement to wait
	public static void waitForElement(String Locatorname,String Locatorvalue,String TestData)
	{
		WebDriverWait mywait=new WebDriverWait(driver,Duration.ofSeconds(Integer.parseInt(TestData)));
		if(Locatorname.equalsIgnoreCase("xpath"))
		{
			mywait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(Locatorvalue)));
		}

		if(Locatorname.equalsIgnoreCase("id"))
		{
			mywait.until(ExpectedConditions.visibilityOfElementLocated(By.id(Locatorvalue)));
		}
		if(Locatorname.equalsIgnoreCase("name"))
		{
			mywait.until(ExpectedConditions.visibilityOfElementLocated(By.name(Locatorvalue)));
		}
	}
	//method for type action
	public static void typeAction(String Locatorname,String Locatorvalue,String TestData)
	{
		if(Locatorname.equalsIgnoreCase("xpath"))
		{
			driver.findElement(By.xpath(Locatorvalue)).clear();
			driver.findElement(By.xpath(Locatorvalue)).sendKeys(TestData);
		}
		if(Locatorname.equalsIgnoreCase("name"))
		{
			driver.findElement(By.name(Locatorvalue)).clear();
			driver.findElement(By.name(Locatorvalue)).sendKeys(TestData);
		}
		if(Locatorname.equalsIgnoreCase("id"))
		{
			driver.findElement(By.id(Locatorvalue)).clear();
			driver.findElement(By.id(Locatorvalue)).sendKeys(TestData);
		}
	}
	//method for click action
	public static void clickAction(String Locatorname,String Locatorvalue)
	{
		if(Locatorname.equalsIgnoreCase("xpath"))
		{
			driver.findElement(By.xpath(Locatorvalue)).click();
		}
		if(Locatorname.equalsIgnoreCase("id"))
		{
			driver.findElement(By.id(Locatorvalue)).sendKeys(Keys.ENTER);
		}
		if(Locatorname.equalsIgnoreCase("name"))
		{
			driver.findElement(By.name(Locatorvalue)).click();
		}
	}
	//method for page title validation
	public static void validateTitle(String Expected_Title)
	{
		String Actual_Title=driver.getTitle();
		try {
			Assert.assertEquals(Actual_Title, Expected_Title, "Title is not matching");
		}catch(AssertionError a)
		{
			System.out.println(a.getMessage());
		}
	}
	//method for closing browser
	public static void closeBrowser()
	{
		driver.quit();	
	}
	//method for date generation
	public static String generateDate()
	{
		Date date=new Date();
		DateFormat df= new SimpleDateFormat("YYYY_MM_DD_mm_ss");
		return df.format(date);

	}
	//method for mouse click
	public static void mouseClick() throws Throwable
	{
		Actions ac=new Actions(driver);
		ac.moveToElement(driver.findElement(By.xpath("//a[starts-with(text(),'Stock Items ')]"))).perform();
		Thread.sleep(3000);
		ac.moveToElement(driver.findElement(By.xpath("(//a[contains(text(),'Stock Categories')])[2]"))).click().perform();
		Thread.sleep(3000);
	}
	//method for category table
	public static void categoryTable(String Exp_Data) throws Throwable
	{
		//if search textbox not displayed click search panel
		if(!driver.findElement(By.xpath(conpro.getProperty("search-textbox"))).isDisplayed())
			//click search panel button
			driver.findElement(By.xpath(conpro.getProperty("search-panel"))).click();
		driver.findElement(By.xpath(conpro.getProperty("search-textbox"))).clear();
		driver.findElement(By.xpath(conpro.getProperty("search-textbox"))).sendKeys(Exp_Data);
		driver.findElement(By.xpath(conpro.getProperty("search-button"))).click();
		Thread.sleep(3000);
		driver.findElement(By.xpath(conpro.getProperty("search-button"))).click();
		String Act_Data=driver.findElement(By.xpath("//table[@class='table ewTable']/tbody/tr[1]/td[4]/div/span/span")).getText();
		Reporter.log(Exp_Data+"   "+Act_Data,true);
		try {
			Assert.assertEquals(Act_Data, Exp_Data, "Category Name is not matching");
		}catch(AssertionError a)
		{
			Reporter.log(a.getMessage(),true);

		}
	}
	//method for listboxes
	public static void dropDownAction(String LocatorName,String LocatorValue,String TestData)
	{
		if(LocatorName.equalsIgnoreCase("xpath"))
		{
			int value = Integer.parseInt(TestData);
			Select element=new Select(driver.findElement(By.xpath(LocatorValue)));
			element.selectByIndex(value);
		}
		if(LocatorName.equalsIgnoreCase("name"))
		{
			int value = Integer.parseInt(TestData);
			Select element=new Select(driver.findElement(By.name(LocatorValue)));
			element.selectByIndex(value);
		}
		if(LocatorName.equalsIgnoreCase("id"))
		{
			int value = Integer.parseInt(TestData);
			Select element=new Select(driver.findElement(By.id(LocatorValue)));
			element.selectByIndex(value);
		}
	}
	//method for capture stock number into notepad
	public static void captureStock(String LocatorName,String LocatorValue) throws Throwable
	{
		String stockNumber="";
		if(LocatorName.equalsIgnoreCase("xpath"))
		{
			stockNumber=driver.findElement(By.xpath(LocatorValue)).getAttribute("value");
		}
		if(LocatorName.equalsIgnoreCase("id"))
		{
			stockNumber=driver.findElement(By.id(LocatorValue)).getAttribute("value");
		}
		if(LocatorName.equalsIgnoreCase("name"))
		{
			stockNumber=driver.findElement(By.name(LocatorValue)).getAttribute("value");
		}
		//Create notepad and write stockNumber
		FileWriter fw=new FileWriter("./CaptureData/stocknumber.txt");
		BufferedWriter bw=new BufferedWriter(fw);
		bw.write(stockNumber);
		bw.flush();
		bw.close();
	}
	//method for stock table
	public static void stockTable() throws Throwable
	{
		//read stock number from notepad
		FileReader fr= new FileReader("./CaptureData/stocknumber.txt");
		BufferedReader br=new BufferedReader(fr);
		String Exp_Data=br.readLine();
		//if search textbox not displayed click search panel
		if(!driver.findElement(By.xpath(conpro.getProperty("search-textbox"))).isDisplayed())
			//click search panel button
			driver.findElement(By.xpath(conpro.getProperty("search-panel"))).click();
		driver.findElement(By.xpath(conpro.getProperty("search-textbox"))).clear();
		driver.findElement(By.xpath(conpro.getProperty("search-textbox"))).sendKeys(Exp_Data);
		driver.findElement(By.xpath(conpro.getProperty("search-button"))).click();
		Thread.sleep(3000);
		String Act_Data= driver.findElement(By.xpath("//table[@class='table ewTable']/tbody/tr[1]/td[8]/div/span/span")).getText();
		Reporter.log(Act_Data+"  "+Exp_Data,true);
		try {
			Assert.assertEquals(Act_Data, Exp_Data, "Stock number is not matching");
		}catch(AssertionError a)
		{
			Reporter.log(a.getMessage(),true);

		}

	}
}


