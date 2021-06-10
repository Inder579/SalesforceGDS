package resources;

import java.io.IOException;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;

public class Popup extends BrowserDriver{
	
	WebDriver driver = null;
	public void initialize1() throws IOException
	{
	
		 driver =browser();

	}
	
	public String popupfeature( WebDriver driver) throws InterruptedException
	
	{
	// path ="C:\\Users\\ibhangu\\Desktop\\confirmationAlert.html";
	   driver.get("confirmationAlert.html");
	   
	   String result = null;
	   WebElement clickalert = driver.findElement(By.xpath("//*[@name='alert']"));
	   
	   clickalert.click();
	   
	   Thread.sleep(10000);
	   
	   try {
	   if(driver.findElement(By.xpath("//*[@id='msg']")).isDisplayed() == true) {
		   
		    result=driver.findElement(By.xpath("//*[@id='msg']")).getText();
		   System.out.println(result);
	   }
	   
	   }
	  

	    catch(Exception e) {
	    	
	    	System.out.println(e.getMessage());
	    } 
	
	return result;
	
	}
	

}
