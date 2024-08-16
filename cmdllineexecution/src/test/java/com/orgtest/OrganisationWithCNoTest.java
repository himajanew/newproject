
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
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.edge.EdgeDriver;
import org.testng.annotations.Test;

public class OrganisationWithCNoTest {

@Test
public void createOrganisationWithPhNum() throws Exception
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
        
        String orgname=s.getRow(1).getCell(2).toString()+rint;
        String phno=s.getRow(1).getCell(7).toString();
        
        
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
  driver.findElement(By.id("phone")).sendKeys(phno);
  
 	  
  driver.findElement(By.xpath("//input[@title='Save [Alt+S]']")).click();
  
 String actctno=driver.findElement(By.id("dtlview_Phone")).getText();

if(actctno.contains(phno))
{
System.out.println(phno +" "+ "is verified");	
}
else
{
System.out.println(phno +" " + "is not verified");
}


driver.close();
}
}