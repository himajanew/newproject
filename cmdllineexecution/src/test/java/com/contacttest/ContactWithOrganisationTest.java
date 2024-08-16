package com.contacttest;

import java.io.FileInputStream;
import java.time.Duration;
import java.util.Iterator;
import java.util.Properties;
import java.util.Random;
import java.util.Set;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.edge.EdgeDriver;
import org.testng.annotations.Test;

public class ContactWithOrganisationTest {
@Test
public void contactWithOrg() throws Exception
{
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
	        
	        String orgname=s.getRow(1).getCell(8).toString()+rint;
	        String lastname=s.getRow(1).getCell(9).toString()+rint;
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
	  driver.findElement(By.xpath("//input[@title='Save [Alt+S]']")).click();
	  String info=driver.findElement(By.xpath("//span[@class='dvHeaderText']")).getText();
	 if(info.contains(orgname))
	 {
		 System.out.println(orgname + "is created");
	 }
	 else
	 {
		 System.out.println(orgname + " is not created");
	 }
	  driver.findElement(By.linkText("Contacts")).click();
	  driver.findElement(By.xpath("//img[@title='Create Contact...']")).click();
	  driver.findElement(By.name("lastname")).sendKeys(lastname);
 	  driver.findElement(By.xpath("//input[@name='account_name']/following-sibling::img")).click();
 	  Set<String> set=driver.getWindowHandles();
      Iterator<String> it=set.iterator();
 	  while(it.hasNext())
 	  {
 		  String windowid=it.next();
 		  driver.switchTo().window(windowid);
 	  
 	  String actualurl=driver.getCurrentUrl();
 	  if(actualurl.contains("module=Accounts")) {
 		  break;	 }	  }
 	  driver.findElement(By.name("search_text")).sendKeys(orgname);
 	  driver.findElement(By.name("search")).click();
 	 driver.findElement(By.xpath("//a[text()='"+orgname+"']")).click();
 	 Set<String> sett=driver.getWindowHandles();
      Iterator<String> its=sett.iterator();
 	  while(its.hasNext())
 	  {
 		  String windowid=its.next();
 		  driver.switchTo().window(windowid);
 	  
 	  String acturl=driver.getCurrentUrl();
 	  if(acturl.contains("Contacts&action")) {
 		  break;
 	  }
 	  }
 	 driver.findElement(By.xpath("//input[@title='Save [Alt+S]']")).click();
 	  info=driver.findElement(By.xpath("//span[@class='dvHeaderText']")).getText();
	 if(info.contains(orgname))
	 {
		 System.out.println(orgname + "is created");
	 }
	 else
	 {
		 System.out.println(orgname + "is not created");
	 }
	  
 String actualorgname=driver.findElement(By.id("mouseArea_Organization Name")).getText();

if(actualorgname.trim().contains(orgname))
{
System.out.println(orgname +" "+ "is created");	
}
else
{
	System.out.println(orgname +" " + "is not created");
}
driver.close();
}

}

