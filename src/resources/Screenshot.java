package resources;

import java.io.File;
import java.io.IOException;

import org.apache.commons.io.FileUtils;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;

import resources.dateandTimeUtility;

public class Screenshot {
	
	 WebDriver driver = null;
	
	public static void  getScreenShot(WebDriver driver, String ScreenshotName) {
		
		TakesScreenshot ts = (TakesScreenshot) driver;
		
		File src = ts.getScreenshotAs(OutputType.FILE);
		
		String Timestamp = dateandTimeUtility.getCurrentTimeStampF1();
		
		String Path = System.getProperty("user.dir")+"//Screenshot//" + ScreenshotName+"_"+Timestamp+".png";
		
		//System.out.println(Path);
		
		File destinationpath = new File(Path);
		
		try {
			FileUtils.copyFile(src, destinationpath);
			
		} catch (IOException e) {
			// TODO Auto-generated catch block
			System.out.println("Capture Failed "+e.getMessage());
		}
		
		//return Path;
		
		
	} 
	
	public static String  getScreenShottwo(WebDriver driver, String ScreenshotName) {
		
		TakesScreenshot ts = (TakesScreenshot) driver;
		
		File src = ts.getScreenshotAs(OutputType.FILE);
		
		String Timestamp = dateandTimeUtility.getCurrentTimeStampF1();
		
		String Path = System.getProperty("user.dir")+"\\Screenshot" +ScreenshotName+"_"+Timestamp+".png";
		
		System.out.println(Path);
		
		File destinationpath = new File(Path);
		
		try {
			FileUtils.copyFile(src, destinationpath);
			
		} catch (IOException e) {
			// TODO Auto-generated catch block
			System.out.println("Capture Failed "+e.getMessage());
		}
		
		return Path;
		
		
	} 
	
        public static String  getScreenShotPath(WebDriver driver, String MethodName) {
		
		TakesScreenshot ts = (TakesScreenshot) driver;
		
		File src = ts.getScreenshotAs(OutputType.FILE);
		
		String Date = dateandTimeUtility.getCurrentDate();
		
		String Time = dateandTimeUtility.getCurrentDate();
		
		String Path = System.getProperty("user.dir")+"//FailedScreenshot//"+Date+Time+ MethodName+".png";
		
		File destinationpath = new File(Path);
		
		try {
			FileUtils.copyFile(src, destinationpath);
			
		} catch (IOException e) {
			// TODO Auto-generated catch block
			System.out.println("Capture Failed "+e.getMessage());
		}
		
		return Path;
		
		
		
		
		
	} 

        public static String capture(WebDriver driver,String screenShotName) throws IOException
	    {
        	String TimeStamp = dateandTimeUtility.getCurrentTimeStampF1();
        	
	        TakesScreenshot ts = (TakesScreenshot)driver;
	        File source = ts.getScreenshotAs(OutputType.FILE);
	        String dest = System.getProperty("user.dir") +"\\screenshot\\"+screenShotName+TimeStamp+".png";
	        File destination = new File(dest);
	        FileUtils.copyFile(source, destination);        
	                     
	        return dest;
	    } 
        
  
}
