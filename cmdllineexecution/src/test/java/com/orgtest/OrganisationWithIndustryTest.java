package com.orgtest;

import java.io.FileInputStream;
import java.time.Duration;
import java.util.Properties;
import java.util.Random;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.support.ui.Select;
import org.testng.annotations.Test;

public class OrganisationWithIndustryTest {
@Test
public void organisationWithIndustry() throws Exception{
        Random ra=new Random();
        int rint=ra.nextInt(1000);
		FileInputStream fis = new FileInputStream("C:\\Users\\viran\\OneDrive\\Documents\\Selenium\\vtige.properties");
		Properties pob = new Properties();
		pob.load(fis);
		String BROWSER = pob.getProperty("browser");
		String URL = pob.getProperty("url");
		String USN = pob.getProperty("username");
		String PWD = pob.getProperty("password");
		FileInputStream fos=new FileInputStream("C:\\Users\\viran\\OneDrive\\Documents\\Selenium\\Excel with condition.xlsx");
		Workbook wb=WorkbookFactory.create(fos);
		Sheet s=wb.getSheet("contact");
		        
		        String orgname=s.getRow(1).getCell(1).toString()+rint;
		        String industry=s.getRow(1).getCell(5).toString();
		        String type=s.getRow(1).getCell(6).toString();
		        
		        wb.close();
		  WebDriver driver=null;
		  if(BROWSER.equals("chrome"))
		  {
			  driver =new ChromeDriver();
		  }
		  else if(BROWSER.equals("edge"))
		  {
			  driver=new EdgeDriver();
		  }
		  else
		  {
			  driver=new ChromeDriver();
		  }
	driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(15));
	driver.manage().window().maximize();
		  driver.get(URL); 
		  driver.findElement(By.name("user_name")).sendKeys(USN);
		  driver.findElement(By.name("user_password")).sendKeys(PWD);
		  driver.findElement(By.id("submitButton")).submit();
		  driver.findElement(By.linkText("Organizations")).click();
		  driver.findElement(By.xpath("//img[@title='Create Organization...']")).click();
		  driver.findElement(By.name("accountname")).sendKeys(orgname);
		  WebElement industrywe=driver.findElement(By.name("industry"));
		  Select se=new Select(industrywe);
		  se.selectByVisibleText(industry);
		  Thread.sleep(1000);
		 WebElement typewe=driver.findElement(By.name("accounttype"));
		 Select sel=new Select(typewe);
		 sel.selectByVisibleText(type);
		  
		  driver.findElement(By.xpath("//input[@title='Save [Alt+S]']")).click();
		  
	     String actindustries=driver.findElement(By.id("dtlview_Industry")).getText();

	if(actindustries.contains(industry))
	{
	System.out.println(industry +" "+ "is verified");	
	}
	else
	{
		System.out.println(industry +" " + "is not verified");
	}
	String acttype=driver.findElement(By.id("dtlview_Type")).getText();

	if(acttype.contains(type))
	{
	System.out.println(acttype +" "+ "is verified");	
	}
	else
	{
		System.out.println(acttype +" " + "is not verified");
	}
	driver.close();
	}
}


