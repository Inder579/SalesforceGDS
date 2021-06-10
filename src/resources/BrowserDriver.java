package resources;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Properties;

import org.openqa.selenium.PageLoadStrategy;
import org.openqa.selenium.UnexpectedAlertBehaviour;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.remote.CapabilityType;
import org.openqa.selenium.remote.DesiredCapabilities;





public class BrowserDriver extends ExtendReport {
	
	public static WebDriver driver;
	public Properties prop;
	public String downloadPath;
	
	public WebDriver browser() throws IOException
	{
		prop = new Properties();
		FileInputStream fy =new FileInputStream(System.getProperty("user.dir")+"\\src\\main\\resources\\objects.properties");
		prop.load(fy);
		
		downloadPath=System.getProperty("user.dir")+"\\src\\main\\resources\\logs";
		
		
		//Set path for driver exe 
		System.setProperty("webdriver.chrome.driver",System.getProperty("user.dir")+"\\src\\main\\resources\\Browser Drivers\\chromedriver_win32\\chromedriver.exe");
		HashMap<String, Object> chromePrefs = new HashMap<String, Object>();
		chromePrefs.put("profile.default_content_settings.popups", 0);
		chromePrefs.put("safebrowsing.enabled", "true");
		chromePrefs.put("download.default_directory", downloadPath);
		ChromeOptions options=new ChromeOptions();
		options.setExperimentalOption("prefs", chromePrefs);
		

			//Add chrome switch to disable notification - "**--disable-notifications**"
			options.addArguments("--disable-notifications");
			options.setPageLoadStrategy(PageLoadStrategy.NONE);
			options.addArguments("--disable-gpu");
			
			//Pass ChromeOptions instance to ChromeDriver Constructor
			DesiredCapabilities capabilities = new DesiredCapabilities();
			capabilities.setCapability(ChromeOptions.CAPABILITY, options);
			capabilities.setCapability(CapabilityType.UNEXPECTED_ALERT_BEHAVIOUR, UnexpectedAlertBehaviour.IGNORE);
			
					
					capabilities.setCapability("autoAcceptAlerts", true);
			ChromeDriver driver = new ChromeDriver(capabilities);
			 driver.manage().window().maximize();
		
	
		return driver;
	
}
	public String getdwnldpath()
	{
		return downloadPath;
	}
	

}

