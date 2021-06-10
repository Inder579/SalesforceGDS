package resources;

import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;

import org.apache.commons.mail.EmailException;
import org.testng.ITestContext;
import org.testng.annotations.AfterSuite;
import org.testng.annotations.BeforeSuite;

import com.aventstack.extentreports.ExtentReports;
import com.aventstack.extentreports.ExtentTest;
import com.aventstack.extentreports.reporter.ExtentHtmlReporter;

public class ExtendReport {
	
       public static ExtentTest test;
       public static ExtentHtmlReporter htmlReporter;
       public static ExtentReports Extent;
       //public static WebDriver driver;
	
	@BeforeSuite
	public void Setup(ITestContext Result) {
		
		String xmlSuiteName = Result.getCurrentXmlTest().getSuite().getName();
		System.out.println(xmlSuiteName);
		//Set time stamp for test suite run
		Calendar cal = Calendar.getInstance();
        Date date=cal.getTime();
        DateFormat dateFormat = new SimpleDateFormat("HH:mm:ss");
        String formattedDate=dateFormat.format(date);
				
		
		if (xmlSuiteName.equalsIgnoreCase("Suite")){
			
			htmlReporter = new ExtentHtmlReporter(".//Reports//"+xmlSuiteName+".html");
			
			Extent = new ExtentReports();			
			Extent.attachReporter(htmlReporter);
		}
	}
	
	/*

	@AfterMethod()
	public void getResult(ITestResult Result)  {
		
		if (Result.getStatus() == ITestResult.SUCCESS) {
			
			test.log(Status.PASS, MarkupHelper.createLabel(Result.getName() + " is Passed ", ExtentColor.GREEN));
			
		}
		
	    else if(Result.getStatus() == ITestResult.FAILURE) {
		 test.log(Status.FAIL, MarkupHelper.createLabel(Result.getName() + "is Failed" , ExtentColor.RED));

		}
		
         else if(Result.getStatus() == ITestResult.SKIP) {
			
			test.log(Status.SKIP, MarkupHelper.createLabel(Result.getName() + " is Skipped ", ExtentColor.ORANGE));
					
		}
	}
	
	*/
	
	@AfterSuite()
	public void tearDown(ITestContext Result) throws EmailException {
		

		
		Extent.flush();
		
		//sendEmail ReportEmail = new sendEmail();
	   // ReportEmail.sendEmailResultwithAttachment();
	}
}
