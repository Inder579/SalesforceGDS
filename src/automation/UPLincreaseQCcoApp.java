package automation;
import java.awt.Dimension;
import java.awt.Font;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileReader;
import java.io.IOException;
import java.util.Iterator;
import java.util.List;
import java.util.concurrent.TimeUnit;

import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JTextPane;
import javax.swing.UIManager;
import javax.swing.plaf.FontUIResource;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.dom4j.DocumentException;
import org.joda.time.DateTime;
import org.joda.time.format.DateTimeFormat;
import org.json.simple.JSONArray;
import org.json.simple.JSONObject;
import org.json.simple.parser.JSONParser;
import org.json.simple.parser.ParseException;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import com.aventstack.extentreports.Status;
import com.aventstack.extentreports.markuputils.ExtentColor;
import com.aventstack.extentreports.markuputils.MarkupHelper;

import resources.BrowserDriver;
import resources.Screenshot;
public class UPLincreaseQCcoApp extends BrowserDriver {
	
	public static int attemptNo = 0;
	public String screenShotPathforInterestRate;
	public WebDriver driver;
	int cvScore, cvScoreCoapp, BehaviourScore;
	public String ActualIncome, appType, Province, ApplicationID, MorgagePayment, CoAppProvince, behscore, cobehscore,
			appbehscore;
	public String TotalIncomeAmount, IntRate, cabKey, cabKeyCoApp, LoanType, contributer;
	public double TotalIncome, RemainingIncome, TotalDebt, ActualQLA, ExpectedQLA,ExpectedMaxHA,ActualMaxHA, ExpInt, ActualQLAapp, ActualQLAcoapp,
			ActualFinalCoApp, SPLltv, Maxltv, HomeEquity, PropertyVal, LtvMax;

	String lowefs, loanType, highefs, Prov, provinceGroup, bkStrategy, ps = "", code = null, propertyType = "",
			propertyLocation = "";
	double lef, hef, calRemIn, QLA, InterestRate, remIn, remInNaPrev, remInNaAfter, interestcontrolgroup1,
			interestcontrolgroup2;
	int fcol, lcol, col, coldiff, rowNum, RiskGroup, SPLTotalDebt, lastNumRow, bkDecreaseAmount;
	String stringSplit[], Strategy;
	
	@BeforeTest
	public void initialize1() throws IOException {

		driver = browser();

	}
	@Test()
	public void CoApp1() throws Exception
	{
		// Login as Admin
				loginAsAdmin();
				// Login as FSR-Application
				loginAsFSR();
				Thread.sleep(3000);
	//			 driver.get("https://c.cs41.visual.force.com/apex/LAMSApplicationView1?id=a0855000006IFGa&sfdc.override=1");
				// WAIT for User to Submit
				waitForFirstSubmission();
				Thread.sleep(3000);
				firstPopup();
				Thread.sleep(4000);
				WebDriverWait wait = new WebDriverWait(driver, 360, 0000);
    			wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("getapptype"))));
				appType = driver.findElement(By.xpath(prop.getProperty("getapptype"))).getText();
		
		System.out.println("UPL");
		getAppDetails();
		calculateIncome();
		Thread.sleep(2000);
		calculateLiability();
		// SWITCH IFRAME DEFAULT
		driver.switchTo().defaultContent();
		getAppTimestampLogs();
		// Interest rate Calculations
		interestRateCalculation();
		checkContributer();

	 if (LoanType.contains("Increase")) {
			Thread.sleep(3000);
			
			calculateQLA();
		}
		maxHA();
		ReasonCode();
		SecondPopup();
	}
	public void ReasonCode() throws InterruptedException, IOException {
		// TODO Auto-generated method stub
		if (attemptNo == 0) {
			test = Extent.createTest("Reason Codes");
		}

		else {

			test = Extent.createTest("Resubmission Result: Attempt#  " + attemptNo + " - Reason Codes");
		}

		Thread.sleep(3000);
		driver.switchTo().defaultContent();
		JavascriptExecutor js = (JavascriptExecutor) driver;
		js.executeScript("window.scrollTo(0, document.body.scrollHeight)");

		Thread.sleep(3000);

		String screenShotPathforInterestRate = Screenshot.capture(driver, "ReasonCode");

		driver.switchTo().frame(driver.findElement(By.xpath(prop.getProperty("switchIframe"))));

		test.log(Status.PASS, MarkupHelper.createLabel(" Reason Codes with GDS Decision", ExtentColor.GREEN));
		test.log(Status.PASS, "Snapshot below: " + test.addScreenCaptureFromPath(screenShotPathforInterestRate));

		Thread.sleep(3000);
	}

	private void getAppTimestampLogs() throws InterruptedException {
		File newFile = new File(System.getProperty("user.dir")+"\\src\\main\\resources\\logs\\Response.json");
		 if (newFile.exists()) {
				Thread.sleep(3000);
				newFile.delete();
				Thread.sleep(3000);
				System.out.println("Response File deleted");
			}
		
		
		driver.findElement(By.xpath(prop.getProperty("gotopg2"))).click();
		Thread.sleep(5000);
		WebDriverWait waitl = new WebDriverWait(driver, 360, 0000);
		waitl.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("getTsApp"))));
		String ts = driver.findElement(By.xpath(prop.getProperty("getTsApp"))).getText();
		System.out.println("Decision Received TimeStamp :" + ts);
		// Split Time-stamp, 1/31/2020 12:16 PM to 1/31/2020 & 12:16 PM
		String[] parts = ts.split(" ");
		String tsdate = parts[0];
		String tstime = parts[1];

		String logs = tsdate + " " + tstime + ":00";

		// parse the string
		org.joda.time.format.DateTimeFormatter dtf = DateTimeFormat.forPattern("MM/dd/yyyy HH:mm:ss");
		// Parsing the date
		DateTime jodatime = dtf.parseDateTime(logs);

		// add two hours
		DateTime date = jodatime.plusMinutes(1);
		DateTime dateTime = jodatime.plusMinutes(2); // easier than mucking about with Calendar and constants

		String ant1 = String.valueOf(date);
		String ts1 = ":" + Character.toString(ant1.charAt(14)) + Character.toString(ant1.charAt(15));

		String ant = String.valueOf(dateTime);
		String ts2 = ":" + Character.toString(ant.charAt(14)) + Character.toString(ant.charAt(15));

		Thread.sleep(2000);

		// Logout as FSR
		WebDriverWait wait1 = new WebDriverWait(driver, 360, 0000);
		wait1.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("logoutmenu"))));
		driver.findElement(By.xpath(prop.getProperty("logoutmenu"))).click();
		Thread.sleep(3000);
		WebDriverWait wait2 = new WebDriverWait(driver, 360, 0000);
		wait2.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("logoutbutton"))));
		driver.findElement(By.xpath(prop.getProperty("logoutbutton"))).click();
		System.out.println("Log out As FSR");

		// Download logs as Admin User

		Thread.sleep(4000);
		WebDriverWait wait3 = new WebDriverWait(driver, 360, 0000);
		wait3.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("clicklogs"))));
		driver.findElement(By.xpath(prop.getProperty("clicklogs"))).click();
		Thread.sleep(4000);
		WebDriverWait waitview = new WebDriverWait(driver, 360, 0000);
		waitview.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("clickview"))));
		WebElement selectView = driver.findElement(By.xpath(prop.getProperty("clickview")));
		Select view = new Select(selectView);
		view.selectByVisibleText("CMO");
		Thread.sleep(3000);
		WebDriverWait wait4 = new WebDriverWait(driver, 360, 0000);
		wait4.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("clickedit"))));
		driver.findElement(By.xpath(prop.getProperty("clickedit"))).click();
		Thread.sleep(3000);
		WebDriverWait wait5 = new WebDriverWait(driver, 360, 0000);
		wait5.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("addtimestampdate"))));
		driver.findElement(By.xpath(prop.getProperty("addtimestampdate"))).clear();
		driver.findElement(By.xpath(prop.getProperty("addtimestampdate"))).sendKeys(tsdate);
		Thread.sleep(3000);
		driver.findElement(By.xpath(prop.getProperty("addfsrname"))).clear();
		driver.findElement(By.xpath(prop.getProperty("addfsrname"))).sendKeys(prop.getProperty("fsrName"));
		Thread.sleep(3000);
		driver.findElement(By.xpath(prop.getProperty("clicksave"))).click();
		Thread.sleep(3000);

		// Re-arrange List and Match Timestamp(ts) from page2 with Logs Timestamp
		System.out.println("Download Logs");
		WebDriverWait wait = new WebDriverWait(driver, 360, 0000);
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("getTsLog"))));
		String Timestamp = driver.findElement(By.xpath(prop.getProperty("getTsLog"))).getText();
		System.out.println("Logs TimeStamp :" + Timestamp);

		String stringSplitter[];

		// Splitting Name & TimeStamp
		stringSplitter = Timestamp.split(",");

		String fsrName = stringSplitter[0];
		String logTimeStamp = stringSplitter[1];

		// Create method to access download path

		if (Timestamp.contains(ts)) {
			WebDriverWait wait6 = new WebDriverWait(driver, 360, 0000);
			wait6.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("clicklog"))));
			driver.findElement(By.xpath(prop.getProperty("clicklog"))).click();
			Thread.sleep(3000);
			WebDriverWait wait7 = new WebDriverWait(driver, 360, 0000);
			wait7.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("downloadlog"))));
			driver.findElement(By.xpath(prop.getProperty("downloadlog"))).click();
			Thread.sleep(3000);
			WebDriverWait wait8 = new WebDriverWait(driver, 360, 0000);
			wait8.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("clickviewfile"))));
			driver.findElement(By.xpath(prop.getProperty("clickviewfile"))).click();
			Thread.sleep(5000);

		}

		else if (Timestamp.contains(ts1)) {
			WebDriverWait wait6 = new WebDriverWait(driver, 360, 0000);
			wait6.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("clicklog"))));
			driver.findElement(By.xpath(prop.getProperty("clicklog"))).click();
			Thread.sleep(3000);
			WebDriverWait wait7 = new WebDriverWait(driver, 360, 0000);
			wait7.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("downloadlog"))));
			driver.findElement(By.xpath(prop.getProperty("downloadlog"))).click();
			Thread.sleep(3000);
			WebDriverWait wait8 = new WebDriverWait(driver, 360, 0000);
			wait8.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("clickviewfile"))));
			driver.findElement(By.xpath(prop.getProperty("clickviewfile"))).click();
			Thread.sleep(5000);

		} else if (Timestamp.contains(ts2)) {
			WebDriverWait wait6 = new WebDriverWait(driver, 360, 0000);
			wait6.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("clicklog"))));
			driver.findElement(By.xpath(prop.getProperty("clicklog"))).click();
			Thread.sleep(3000);
			WebDriverWait wait7 = new WebDriverWait(driver, 360, 0000);
			wait7.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("downloadlog"))));
			driver.findElement(By.xpath(prop.getProperty("downloadlog"))).click();
			Thread.sleep(3000);
			WebDriverWait wait8 = new WebDriverWait(driver, 360, 0000);
			wait8.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("clickviewfile"))));
			driver.findElement(By.xpath(prop.getProperty("clickviewfile"))).click();
			Thread.sleep(5000);

		} else if (!logTimeStamp.contains(ts)) {
			driver.findElement(By.xpath(prop.getProperty("clickname"))).click();
			Thread.sleep(5000);
			String Timestamp1 = driver.findElement(By.xpath(prop.getProperty("getTsLog"))).getText();
			System.out.println("Logs TimeStamp :" + Timestamp1);
			if (Timestamp1.contains(ts)) {
				WebDriverWait wait6 = new WebDriverWait(driver, 360, 0000);
				wait6.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("clicklog"))));
				driver.findElement(By.xpath(prop.getProperty("clicklog"))).click();
				Thread.sleep(3000);
				WebDriverWait wait7 = new WebDriverWait(driver, 360, 0000);
				wait7.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("downloadlog"))));
				driver.findElement(By.xpath(prop.getProperty("downloadlog"))).click();
				Thread.sleep(3000);
				WebDriverWait wait8 = new WebDriverWait(driver, 360, 0000);
				wait8.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("clickviewfile"))));
				driver.findElement(By.xpath(prop.getProperty("clickviewfile"))).click();
				Thread.sleep(5000);

			}

			else if (Timestamp1.contains(ts1)) {
				WebDriverWait wait6 = new WebDriverWait(driver, 360, 0000);
				wait6.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("clicklog"))));
				driver.findElement(By.xpath(prop.getProperty("clicklog"))).click();
				Thread.sleep(3000);
				WebDriverWait wait7 = new WebDriverWait(driver, 360, 0000);
				wait7.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("downloadlog"))));
				driver.findElement(By.xpath(prop.getProperty("downloadlog"))).click();
				Thread.sleep(3000);
				WebDriverWait wait8 = new WebDriverWait(driver, 360, 0000);
				wait8.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("clickviewfile"))));
				driver.findElement(By.xpath(prop.getProperty("clickviewfile"))).click();
				Thread.sleep(5000);

			} else if (Timestamp1.contains(ts2)) {
				WebDriverWait wait6 = new WebDriverWait(driver, 360, 0000);
				wait6.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("clicklog"))));
				driver.findElement(By.xpath(prop.getProperty("clicklog"))).click();
				Thread.sleep(3000);
				WebDriverWait wait7 = new WebDriverWait(driver, 360, 0000);
				wait7.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("downloadlog"))));
				driver.findElement(By.xpath(prop.getProperty("downloadlog"))).click();
				Thread.sleep(3000);
				WebDriverWait wait8 = new WebDriverWait(driver, 360, 0000);
				wait8.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("clickviewfile"))));
				driver.findElement(By.xpath(prop.getProperty("clickviewfile"))).click();
				Thread.sleep(5000);

			}

		}
	}
	
	private void calculateLiability() throws IOException, InterruptedException {
		// TODO Auto-generated method stub
		Thread.sleep(4000);
		if (attemptNo == 0) {
			test = Extent.createTest("Total Liability Calcuation");
		}

		else {

			test = Extent.createTest("Resubmission Result: Attempt#  " + attemptNo + " - Total Liability Calcuation");
		}

		// Liabilities Table
		System.out.println("Liabilities Table");
		// SWITCH IFRAME DEFAULT
		driver.switchTo().defaultContent();

		driver.manage().timeouts().implicitlyWait(0, TimeUnit.SECONDS);
		WebElement Liabilitiestable = driver.findElement(By.xpath(prop.getProperty("liabilitiesTable")));
		new Actions(driver).moveToElement(Liabilitiestable).perform();
		Thread.sleep(2000);
		List<WebElement> rowValsLiabilities = Liabilitiestable.findElements(By.tagName("tr"));
		int rowNumLiabilities = Liabilitiestable.findElements(By.tagName("tr")).size();

		int colNumLiabilities = driver.findElements(By.xpath(prop.getProperty("colliabilities"))).size();
		System.out.println("Total number of rows = " + rowNumLiabilities);
		System.out.println("Total number of columns = " + colNumLiabilities);

		double LiabilitiesValue = 0;
		String str = null;
		double subTotal;
		for (int i = 0; i < rowNumLiabilities; i++) {
			// Get each row's column values by tag name
			List<WebElement> colValsLiabilities = rowValsLiabilities.get(i).findElements(By.tagName("td"));
			String LiabilitiesAmount = colValsLiabilities.get(4).getText();
			System.out.println(LiabilitiesAmount);
			if (LiabilitiesAmount.contains(",")) {
				str = LiabilitiesAmount.replace(",", "");
				subTotal = Double.parseDouble(str.replace("$", ""));
			} else {
				subTotal = Double.parseDouble(LiabilitiesAmount.replace("$", ""));
			}
			System.out.println(subTotal);

			LiabilitiesValue += subTotal;

			System.out.println("---------------");
		}

		System.out.println(LiabilitiesValue);

		// Rent Expense
		System.out.println("Rent Expense Table");
		driver.manage().timeouts().implicitlyWait(0, TimeUnit.SECONDS);
		WebElement Renttable = driver.findElement(By.xpath(prop.getProperty("rent")));
		new Actions(driver).moveToElement(Renttable).perform();
		Thread.sleep(2000);
		List<WebElement> rowRent = Renttable.findElements(By.xpath(prop.getProperty("rowrentco")));

		int rowNumRentco = Renttable.findElements(By.xpath(prop.getProperty("rowrentco"))).size();
		int rowNumRent = rowNumRentco / 2;
		int colNumRent = driver.findElements(By.xpath(prop.getProperty("colrentco"))).size();
		System.out.println("Total number of Rows = " + rowNumRent);
		System.out.println("Total number of columns = " + colNumRent);
		double RentValue = 0;
		String str1 = null;
		double subTotal1;
		for (int i = 0; i < rowNumRent; i++) {
			// Get each row's column values by tag name
			List<WebElement> colValsRent = rowRent.get(i).findElements(By.tagName("td"));
			String RentAmount = colValsRent.get(2).getText();
			System.out.println(RentAmount);

			if (RentAmount.contains(",")) {
				str1 = RentAmount.replace(",", "");
				subTotal1 = Double.parseDouble(str1.replace("$", ""));
			} else {
				subTotal1 = Double.parseDouble(RentAmount.replace("$", ""));
			}

			System.out.println(subTotal1);

			RentValue += subTotal1;

			System.out.println("---------------");
		}

		System.out.println(RentValue);

		// Mortgages Table
		System.out.println("Mortgages Table");

		double MortgagesValue = 0;
		String str2 = null;
		double subTotal2;
		driver.manage().timeouts().implicitlyWait(0, TimeUnit.SECONDS);
		int colNumMortgages = driver.findElements(By.xpath(prop.getProperty("colmortgagesco"))).size();

		WebElement Mortgagestable = driver.findElement(By.xpath(prop.getProperty("mortgage")));

		List<WebElement> rowMortgages = Mortgagestable.findElements(By.xpath(prop.getProperty("rowmortgagesco")));

		int rowNumMortgagesco = Mortgagestable.findElements(By.xpath(prop.getProperty("rowmortgagesco"))).size();
		int rowNumMortgages = rowNumMortgagesco / 2;
		System.out.println("Total number of columns = " + rowNumMortgages);

		System.out.println("Total number of columns = " + colNumMortgages);

		for (int i = 0; i < rowNumMortgages; i++) {
			// Get each row's column values by tag name
			List<WebElement> colValsMortgages = rowMortgages.get(i).findElements(By.tagName("td"));
			String RentAmount = colValsMortgages.get(4).getText();
			System.out.println(RentAmount);
			if (RentAmount.contains(",")) {
				str2 = RentAmount.replace(",", "");
				subTotal2 = Double.parseDouble(str2.replace("$", ""));
			} else {
				subTotal2 = Double.parseDouble(RentAmount.replace("$", ""));
			}

			System.out.println(subTotal2);

			MortgagesValue += subTotal2;

			System.out.println("---------------");
		}

		System.out.println(MortgagesValue);

		// Other Table
		System.out.println("Other Table");
		driver.manage().timeouts().implicitlyWait(0, TimeUnit.SECONDS);
		WebElement Othertable = driver.findElement(By.xpath(prop.getProperty("other")));
		new Actions(driver).moveToElement(Othertable).perform();

		List<WebElement> rowOther = Othertable.findElements(By.xpath(prop.getProperty("rowotherco")));

		int rowNumOtherco = Othertable.findElements(By.xpath(prop.getProperty("rowotherco"))).size();
		int rowNumOther = rowNumOtherco / 2;
		int colNumOther = driver.findElements(By.xpath(prop.getProperty("colotherco"))).size();
		System.out.println("Total number of rows = " + rowNumOther);
		System.out.println("Total number of columns = " + colNumOther);
		double OtherValue = 0;
		String str3 = null;
		double subTotal3;
		for (int i = 0; i < rowNumOther; i++) {
			// Get each row's column values by tag name
			List<WebElement> colValsOther = rowOther.get(i).findElements(By.tagName("td"));
			String RentAmount = colValsOther.get(4).getText();
			System.out.println(RentAmount);
			if (RentAmount.contains(",")) {
				str3 = RentAmount.replace(",", "");
				subTotal3 = Double.parseDouble(str3.replace("$", ""));
			} else {
				subTotal3 = Double.parseDouble(RentAmount.replace("$", ""));
			}

			System.out.println(subTotal3);

			OtherValue += subTotal3;

			System.out.println("---------------");
		}

		System.out.println(OtherValue);

		double TotalLaibility = LiabilitiesValue + RentValue + OtherValue;
		System.out.println("Actual Liabilities: $" + TotalLaibility);
		// test = Extent.createTest(" Calculate Liability");
		// String ActualLiability = "$"+TotalLaibility;
		// test.info("Actual Liability ="+ActualLiability);

		// Get Total Debt value - Liabilities Comparison
		Thread.sleep(2000);

		JavascriptExecutor js = (JavascriptExecutor) driver;
		js.executeScript("document.body.style.zoom='90%'");
		Thread.sleep(3000);

		String screenShotPath = Screenshot.capture(driver, "CaculateLiability");
		js.executeScript("document.body.style.zoom='100%'");
		Thread.sleep(3000);

		driver.switchTo().frame(driver.findElement(By.xpath(prop.getProperty("switchIframe"))));
		WebElement Totalpay = driver.findElement(By.xpath(prop.getProperty("totalpayment")));

		String TotalDebtAmount = Totalpay.getText();
		// test.info("Expected Liability ="+TotalDebtAmount);
		String str4 = TotalDebtAmount.replace(",", "");
		TotalDebt = Double.parseDouble(str4.replace("$", ""));

		System.out.println("Expected Liabilities: $" + TotalDebt);
		if (TotalLaibility == TotalDebt) {
			System.out.println("Laibilities:Passed");
			test.log(Status.PASS, MarkupHelper.createLabel("Total Liability - Actual Value   =  $" + TotalLaibility,
					ExtentColor.GREEN));
			test.log(Status.PASS,
					MarkupHelper.createLabel("Total Liability - Expected Value =  $" + TotalDebt, ExtentColor.GREEN));

			test.log(Status.PASS,
					MarkupHelper.createLabel("Liability is Matching with GDS Decision ", ExtentColor.GREEN));
			test.log(Status.PASS, "Snapshot below: " + test.addScreenCaptureFromPath(screenShotPath));

			// Assert.assertTrue(true);
		} else {
			System.out.println("Laibilities:Failed");

			test.log(Status.FAIL, MarkupHelper.createLabel("Total Liability - Actual Value   =  $" + TotalLaibility,
					ExtentColor.RED));
			test.log(Status.FAIL,
					MarkupHelper.createLabel("Total Liability - Expected Value =  $" + TotalDebt, ExtentColor.RED));

			test.log(Status.FAIL,
					MarkupHelper.createLabel("Liability is not Matching with GDS Decision ", ExtentColor.RED));
			test.log(Status.FAIL, "Snapshot below: " + test.addScreenCaptureFromPath(screenShotPath));

			// Assert.assertTrue(false);
		}
	}

	private void calculateIncome() throws IOException, InterruptedException {
		// TODO Auto-generated method stub
		System.out.println("Resubmission attempt #" + attemptNo);
		if (attemptNo == 0) {
			test = Extent.createTest("Total Income Calcuation");
		}

		else {

			test = Extent.createTest("Resubmission Result: Attempt#  " + attemptNo + " - Total Income Calcuation");
		}

		// SWITCH IFRAME DEFAULT
		driver.switchTo().defaultContent();

		WebDriverWait waitLoad = new WebDriverWait(driver, 360, 0000);
		waitLoad.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("summary"))));
		driver.findElement(By.xpath(prop.getProperty("summary"))).click();

		// Income Table
		Thread.sleep(2000);

		WebElement Incometable = driver.findElement(By.xpath(prop.getProperty("incomeTable")));
		new Actions(driver).moveToElement(Incometable).perform();

		Thread.sleep(2000);
		List<WebElement> rowValsIncome = Incometable.findElements(By.tagName("tr"));
		int rowNumIncome = Incometable.findElements(By.tagName("tr")).size();

		int colNumIncome = driver.findElements(By.xpath(prop.getProperty("colincome"))).size();
		System.out.println("Income Table");

		double IncomeValue = 0;
		String str = null;

		for (int i = 0; i < rowNumIncome; i++) {
			double subTotal = 0;
			double subValue = 0;
			// Get each row's column values by tag name
			List<WebElement> colValsIncome = rowValsIncome.get(i).findElements(By.tagName("td"));
			String IncomeFreq = colValsIncome.get(3).getText();
			System.out.println(IncomeFreq);
			String IncomeAmount = colValsIncome.get(4).getText();
			System.out.println(IncomeAmount);

			if (IncomeAmount.contains(","))

			{
				str = IncomeAmount.replace(",", "");
				subValue = Double.parseDouble(str.replace("$", ""));
			} else {
				subValue = Double.parseDouble(IncomeAmount.replace("$", ""));
			}
			System.out.println(subValue);

			if (IncomeFreq.equals("Weekly")) {
				subTotal = subValue * 4.0;

			}

			else if (IncomeFreq.equals("Bi-Weekly")) {
				subTotal = subValue * 2.0;

			} else if (IncomeFreq.equals("Semi-Monthly")) {
				subTotal = subValue * 2.0;

			} else if (IncomeFreq.equals("Monthly")) {
				subTotal = subValue * 1.0;

			}

			IncomeValue += subTotal;
			System.out.println("---------------");

		}

		System.out.println("Actual Income: $" + IncomeValue);

		// Income field Comparison
		Thread.sleep(2000);
		JavascriptExecutor js = (JavascriptExecutor) driver;
		js.executeScript("document.body.style.zoom='80%'");
		Thread.sleep(3000);
		String screenShotPath = Screenshot.capture(driver, "CalculateIncome");
		js.executeScript("document.body.style.zoom='100%'");
		Thread.sleep(5000);
		driver.switchTo().frame(driver.findElement(By.xpath(prop.getProperty("switchIframe"))));

		WebElement Totalinc = driver.findElement(By.xpath(prop.getProperty("totalincome")));

		TotalIncomeAmount = Totalinc.getText();
		// test.info("Expected Income ="+TotalIncomeAmount);
		String st = TotalIncomeAmount.replace(",", "");
		TotalIncome = Double.parseDouble(st.replace("$", ""));

		System.out.println("Expected Income: $" + TotalIncome);
		if (IncomeValue == TotalIncome) {

			test.log(Status.PASS,
					MarkupHelper.createLabel(" Total Income  :Actual Value =  $" + IncomeValue, ExtentColor.GREEN));
			test.log(Status.PASS,
					MarkupHelper.createLabel("Total Income  :Expected Value =  $" + TotalIncome, ExtentColor.GREEN));
			test.log(Status.PASS, MarkupHelper.createLabel("Income is Matching with GDS Decision", ExtentColor.GREEN));
			test.log(Status.PASS, "Snapshot below: " + test.addScreenCaptureFromPath(screenShotPath));
			System.out.println("Income:Passed");

		} else {

			test.log(Status.FAIL,
					MarkupHelper.createLabel("Total Income : Actual Value =  $" + IncomeValue, ExtentColor.RED));
			test.log(Status.FAIL,
					MarkupHelper.createLabel("Total Income : Expected Value =  $" + TotalIncome, ExtentColor.RED));
			test.log(Status.FAIL, MarkupHelper.createLabel("Income Not Matching with GDS Decision", ExtentColor.RED));
			test.log(Status.FAIL, "Snapshot below: " + test.addScreenCaptureFromPath(screenShotPath));
			System.out.println("Income Not Match");
		}
	}

	private void getAppDetails() throws InterruptedException {
		// TODO Auto-generated method stub
		// TODO Auto-generated method stub
				WebDriverWait waitLoad1 = new WebDriverWait(driver, 360, 0000);
				waitLoad1.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("getprovince"))));
				Province = driver.findElement(By.xpath(prop.getProperty("getprovince"))).getText();
				CoAppProvince = driver.findElement(By.xpath(prop.getProperty("getCoAppprovince"))).getText();
				LoanType = driver.findElement(By.xpath(prop.getProperty("loantype"))).getText();
				ApplicationID = driver.findElement(By.xpath(prop.getProperty("getappid"))).getText();
				appType = driver.findElement(By.xpath(prop.getProperty("getapptype"))).getText();
				
				System.out.println("App Type :" + appType);
				System.out.println("Loan Type :" + LoanType);
				driver.switchTo().frame(driver.findElement(By.xpath(prop.getProperty("switchIframe"))));
				Thread.sleep(4000);
				IntRate = driver.findElement(By.xpath(prop.getProperty("get%"))).getText();
				ExpInt = Double.parseDouble(IntRate.replace("%", ""));

				String intqla = driver.findElement(By.xpath(prop.getProperty("getQla"))).getText();
				String qla = intqla.replace(",", "");
				
				ExpectedQLA = Double.parseDouble(qla.replace("$", ""));
				String intHA=driver.findElement(By.xpath(prop.getProperty("getmaxHA"))).getText();
				String HA = intHA.replace(",", "");

				ExpectedMaxHA = Double.parseDouble(HA.replace("$", ""));
				String cvscore = driver.findElement(By.xpath(prop.getProperty("getcvscore"))).getText();
				cvScore = Integer.parseInt(cvscore);
				String cvscoreco = driver.findElement(By.xpath(prop.getProperty("getcvcoapp"))).getText();
				cvScoreCoapp = Integer.parseInt(cvscoreco);

				Strategy = driver.findElement(By.xpath(prop.getProperty("adCoAppStrategy"))).getText();
				if (LoanType.contains("New")) {
					cabKey = driver.findElement(By.xpath(prop.getProperty("cabkeyApplicant"))).getText();
					cabKeyCoApp = driver.findElement(By.xpath(prop.getProperty("cabkeyCoApp"))).getText();
					bkStrategy = driver.findElement(By.xpath(prop.getProperty("bkstrategy"))).getText();
				}
				riskFactor();
				decisionCode();
				System.out.println("ApplicationID: " + ApplicationID);
				System.out.println("Province: " + Province);

				System.out.println("Interest Rate :" + ExpInt);
				System.out.println("CV Score: " + cvScore);
				System.out.println("QLA: " + ExpectedQLA);
				System.out.println("Strategy :" + Strategy);
	}

	public void loginAsAdmin() throws InterruptedException {
		driver.get(prop.getProperty("sfUrl"));
		Thread.sleep(2000);
		WebDriverWait waitLoad = new WebDriverWait(driver, 360, 0000);
		waitLoad.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("username"))));
//		driver.findElement(By.xpath(prop.getProperty("username"))).sendKeys(decodeString(prop.getProperty("AdminEmail")));
		driver.findElement(By.xpath(prop.getProperty("username"))).sendKeys(prop.getProperty("AdminEmail"));
		Thread.sleep(2000);
//		driver.findElement(By.cssSelector(prop.getProperty("password"))).sendKeys(decodeString(prop.getProperty("AdminPassword")));
		driver.findElement(By.cssSelector(prop.getProperty("password"))).sendKeys(prop.getProperty("AdminPassword"));
		driver.findElement(By.xpath(prop.getProperty("clicklogin"))).click();
		
		
		System.out.println("Logged in As Admin");
		Thread.sleep(10000);
		String baseUrl = driver.getCurrentUrl();
		System.out.println(baseUrl);
		if (baseUrl.contains(prop.getProperty("lightningUrl1"))
				|| baseUrl.contains(prop.getProperty("lightningUrl2"))) {
			// SWITCH TO CLASSIC VIEW
			Thread.sleep(3000);
			WebDriverWait wait1 = new WebDriverWait(driver, 360, 0000);
			wait1.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("userIcon"))));
			driver.findElement(By.xpath(prop.getProperty("userIcon"))).click();
			Thread.sleep(3000);
			WebDriverWait wait2 = new WebDriverWait(driver, 360, 0000);
			wait2.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("switchToSFClassic"))));
			driver.findElement(By.xpath(prop.getProperty("switchToSFClassic"))).click();
		}
	}
	public void loginAsFSR() throws InterruptedException {
		Thread.sleep(5000);
		WebDriverWait waithome = new WebDriverWait(driver, 360, 0000);
		waithome.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("hometab"))));
		driver.findElement(By.xpath(prop.getProperty("hometab"))).click();
		Thread.sleep(4000);
		WebDriverWait waitLoad = new WebDriverWait(driver, 360, 0000);
		waitLoad.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("searchmenu"))));
		Thread.sleep(4000);
		driver.findElement(By.xpath(prop.getProperty("searchmenu"))).sendKeys(prop.getProperty("fsrName"));
		Thread.sleep(3000);
		driver.findElement(By.xpath(prop.getProperty("searchbutton"))).click();
		Thread.sleep(3000);
		WebDriverWait wait = new WebDriverWait(driver, 360, 0000);
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("clickfsr"))));

		driver.findElement(By.xpath(prop.getProperty("clickfsr"))).click();
		Thread.sleep(3000);
		WebDriverWait wait1 = new WebDriverWait(driver, 360, 0000);
		wait1.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("fsrmenubutton"))));
		driver.findElement(By.xpath(prop.getProperty("fsrmenubutton"))).click();
		Thread.sleep(3000);
		WebDriverWait wait2 = new WebDriverWait(driver, 360, 0000);
		wait2.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("userdetail"))));
		driver.findElement(By.xpath(prop.getProperty("userdetail"))).click();
		Thread.sleep(3000);
		WebDriverWait wait3 = new WebDriverWait(driver, 360, 0000);
		wait3.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("fsrloginbutton"))));
		Thread.sleep(3000);
		driver.findElement(By.xpath(prop.getProperty("fsrloginbutton"))).click();
		Thread.sleep(2000);
		System.out.println("Logged in As FSR");
	}
/*		public void waitForFirstSubmission() throws Exception {

	 

    // Accept Add Co-Applicant Alert



    WebDriverWait waitCheckBox = new WebDriverWait(driver, 360, 000000000);
    waitCheckBox.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("AddCoApp"))));
    System.out.println("Add CoApplicant");
    WebDriverWait waitAlertPresent = new WebDriverWait(driver, 360, 0000000000);
    waitAlertPresent.until(ExpectedConditions.alertIsPresent());
    driver.switchTo().alert().accept();
    System.out.println("accepted");



    Thread.sleep(10000);
    WebDriverWait waitLoading = new WebDriverWait(driver, 360, 0000000);
    waitLoading.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("alertmessage"))));
    System.out.println("Alert is displayed");
    System.out.println("Clicked Save button");
    Thread.sleep(10000);
    WebDriverWait waitLoad12 = new WebDriverWait(driver, 360, 0000000);
    waitLoad12.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath(prop.getProperty("alertmessage"))));
    System.out.println("Alert not displayed");



    Thread.sleep(10000);
    System.out.println("Clicked Re-submit");
    Thread.sleep(7000);
}
*/

	public void waitForFirstSubmission() throws Exception {

		// Accept Applicant Section Complete Alert

		// We are declaring the frame
		JFrame frmOpt = new JFrame(); // We are declaring the frame
		frmOpt.setAlwaysOnTop(true);// This is the line for displaying it above all windows

		Thread.sleep(1000);
		String s = "<html>Press 1 For Calculations<br>Press 2 For Results<br>";

		JLabel label = new JLabel(s);
		JTextPane jtp = new JTextPane();
		jtp.setSize(new Dimension(480, 10));
		jtp.setPreferredSize(new Dimension(480, jtp.getPreferredSize().height));
		label.setFont(new Font("Arial", Font.BOLD, 26));
		UIManager.put("OptionPane.minimumSize", new Dimension(700, 300));
		UIManager.put("TextField.font", new FontUIResource(new Font("Verdana", Font.BOLD, 20)));
		// Getting Input from user

		String option = JOptionPane.showInputDialog(frmOpt, label);

		int useroption = Integer.parseInt(option);

		switch (useroption) {

		case 1:

			// Function for Re-Submission

			break;

		case 2:

			System.out.println("Results");
			if (attemptNo == 0) {
				test = Extent.createTest("Resubmission - Confirmation ");
				test.info(" The test run complete. Please review test result(s)");
			}

			else {

				test = Extent.createTest("Resubmission - Confirmation ");
				test.info(" The test run complete. Please review test result(s)");
			}

			Thread.sleep(3000);

			driver.close();
			driver.quit();
			break;

		}

}

public void firstPopup() throws InterruptedException {
	Thread.sleep(10000);
	// First Pop-up
	driver.get(System.getProperty("user.dir") + "\\src\\main\\resources\\confirmationAlert1.html");
	WebDriverWait waitLoad = new WebDriverWait(driver, 360, 00000000);
	waitLoad.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@name='alert']")));
	WebElement clickalert = driver.findElement(By.xpath("//*[@name='alert']"));
	clickalert.click();

	Thread.sleep(12000);
	String response = null;

	try {
		if (driver.findElement(By.xpath("//*[@id='msg']")).isDisplayed() == true) {
			response = driver.findElement(By.xpath("//*[@id='msg']")).getText();

			System.out.println(response);
		}

	}

	catch (Exception e) {

		System.out.println(e.getMessage());
	}

	if (response.contains("OK")) {

		driver.navigate().back();
		Thread.sleep(5000);
	} else if (response.contains("CANCEL")) {

		test = Extent.createTest("Get Application Details ");

		Thread.sleep(3000);

		driver.close();
		driver.quit();

		test.info("You opted to Close the Automation Test Run");

		test.log(Status.PASS, MarkupHelper.createLabel("Automation Exited", ExtentColor.GREEN));
	}
	

}
	
public void interestRateCalculation() throws DocumentException, InterruptedException, IOException, ParseException {

	if (attemptNo == 0) {
		test = Extent.createTest("Interest Rate Calculation");
	}

	else {

		test = Extent.createTest("Resubmission Result: Attempt#  " + attemptNo + " - InterestRate Calculation");
	}
	// Read Random number from XML
	Thread.sleep(8000);
	
	// TODO Auto-generated method stub
	File file1 = new File(System.getProperty("user.dir")+"\\src\\main\\resources\\logs\\Response.xml");
    File newFile = new File(System.getProperty("user.dir")+"\\src\\main\\resources\\logs\\Response.json");
    if(file1.renameTo(newFile)){
        System.out.println("File rename success");
    }else{
        System.out.println("File rename failed");
    }
    JSONParser parser = new JSONParser();
    Object obj = parser.parse(new FileReader(System.getProperty("user.dir")+"\\src\\main\\resources\\logs\\Response.json"));
    JSONObject jsonObject =  (JSONObject) obj;
    JSONArray cars = (JSONArray) jsonObject.get("Response");
    String txt =cars.toString();
    
    int index1 = txt.indexOf("<DE_UPL_Credit_Contributor>");
    contributer  = txt.substring(index1+27, index1+38);
	
	File inputFile = new File(System.getProperty("user.dir") + "\\src\\main\\resources\\logs\\Response.xml");


	if (LoanType.contains("New")) {
		if (bkStrategy.contains("Bankruptcy QLA Decrease")) {
			
			if (contributer.contains("Applicant 2"))
			{
				int index4 = txt.indexOf("<DE_UPL_App2_BKQualifiedLoanAmount_Decrement>");
		        String roar2 = txt.substring(index4+45, index4+50);
		        if(roar2.contains("-"))
		        {
		        double bkdec = Double.valueOf(roar2);
				bkDecreaseAmount = (int) bkdec;
		        }
			}
			else
			{
			int index4 = txt.indexOf("<DE_UPL_App1_BKQualifiedLoanAmount_Decrement>");
	        String roar2 = txt.substring(index4+45, index4+50);
	        if(roar2.contains("-"))
	        {
	        double bkdec = Double.valueOf(roar2);
			bkDecreaseAmount = (int) bkdec;
	        }
			}
         	System.out.println("BK Decrease Amount : " + bkDecreaseAmount);
		}
	}
	double intRate = 35.00;

	// Delete Response File
	 if (file1.exists()) {
			Thread.sleep(3000);
			file1.delete();
			Thread.sleep(3000);
			System.out.println("Response File deleted");
		}
		
		if (newFile.exists()) {
			Thread.sleep(3000);
			newFile.delete();
			Thread.sleep(3000);
			System.out.println("Response File deleted");
		}
	if (inputFile.exists()) {
		Thread.sleep(3000);
		inputFile.delete();
		Thread.sleep(3000);
		System.out.println("Response File deleted");
	}
	
	loginAsFSR();
	Thread.sleep(3000);
	landOnAppPage();
	Thread.sleep(5000);

	driver.switchTo().defaultContent();
	WebDriverWait wait = new WebDriverWait(driver, 360, 0000);
	wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("re-submit"))));
	WebElement Int = driver.findElement(By.xpath(prop.getProperty("re-submit")));
	new Actions(driver).moveToElement(Int).perform();

	Thread.sleep(3000);
	JavascriptExecutor js = (JavascriptExecutor) driver;
	js.executeScript("arguments[0].focus();", Int);
	js.executeScript("document.body.style.zoom='50%'");
	Thread.sleep(3000);

	String screenShotPathforInterestRate = Screenshot.capture(driver, "CaculateInterestRate");
	js.executeScript("document.body.style.zoom='100%'");
	driver.switchTo().frame(driver.findElement(By.xpath(prop.getProperty("switchIframe"))));

	// Displaying Interest Rate result
	double Actint = intRate;
	System.out.println("Actual Interest rate: " + Actint);
	System.out.println("Expected Interest rate: " + ExpInt);

	if (ExpInt == Actint) {

		test.log(Status.PASS, MarkupHelper.createLabel("InterestRate Percentage Actual value : " + Actint + "%",
				ExtentColor.GREEN));
		test.log(Status.PASS, MarkupHelper.createLabel("InterestRate Percentage Expected value : " + ExpInt + "%",
				ExtentColor.GREEN));

		test.log(Status.PASS, MarkupHelper.createLabel(" Interest Rate Calculation is Matching with GDS Decision",
				ExtentColor.GREEN));
		test.log(Status.PASS, "Snapshot below: " + test.addScreenCaptureFromPath(screenShotPathforInterestRate));

		System.out.println("PASSED in Interest Verification");
	} else {

		test.log(Status.FAIL, MarkupHelper.createLabel("InterestRate Percentage Actual value : " + Actint + "%",
				ExtentColor.RED));
		test.log(Status.FAIL, MarkupHelper.createLabel("InterestRate Percentage Expected value : " + ExpInt + "%",
				ExtentColor.RED));

		test.log(Status.FAIL, MarkupHelper.createLabel(" Interest Rate Calculation not Matching with GDS Decision",
				ExtentColor.RED));
		test.log(Status.FAIL, "Snapshot below: " + test.addScreenCaptureFromPath(screenShotPathforInterestRate));

		System.out.println("FAILED in Interest Rate Verification");
	}
	Thread.sleep(3000);
}

public void landOnAppPage() throws InterruptedException {
	System.out.println("Go To Applicant's Page");
	Thread.sleep(3000);
	WebDriverWait wait = new WebDriverWait(driver, 360, 0000);
	wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("applicationsbutton"))));
	driver.findElement(By.xpath(prop.getProperty("applicationsbutton"))).click();
	Thread.sleep(6000);
	WebDriverWait wait1 = new WebDriverWait(driver, 360, 0000);
	wait1.until(ExpectedConditions
			.visibilityOfElementLocated(By.xpath("//a[normalize-space(text())=\"" + ApplicationID + "\"]")));
	WebElement link = driver.findElement(By.xpath("//a[normalize-space(text())=\"" + ApplicationID + "\"]"));

	link.click();
}
public void checkContributer() {
	// TODO Auto-generated method stub
	if (contributer.contains("Applicant 1")) {
		System.out.println("Credit_Contributer = Applicant 1");

	} else if (contributer.contains("Applicant 2")) {
		System.out.println("Credit_Contributer = Applicant 2");
		cvScore = cvScoreCoapp;
		Province = CoAppProvince;

	} else {
		System.out.println("Credit_Contributer = Shared");

	}

}
public void calculateQLA() throws InterruptedException, IOException {

	Thread.sleep(5000);
	if (attemptNo == 0) {
		test = Extent.createTest("QLA Calculation");
	}

	else {

		test = Extent.createTest("Resubmission Result: Attempt#  " + attemptNo + " - QLA Calculation");
	}

	File file = new File(System.getProperty("user.dir") + "\\src\\main\\resources\\Excel\\EFS-CV-Grids-FINAL.xlsx");

	FileInputStream inputStream = new FileInputStream(file);

	Workbook workbook = new XSSFWorkbook(inputStream);
	String IntRate = String.valueOf(ExpInt);

	org.apache.poi.ss.usermodel.Sheet sheet = workbook.getSheet("35.99");

	Iterator<Row> rows = sheet.iterator();

	Row row = rows.next();

	Iterator<Cell> cell = row.cellIterator();

	Cell value;

	RemainingIncome = TotalIncome * 0.55 - TotalDebt;
	System.out.println("RemainingIncome ="+RemainingIncome);
	// Setting Province label

	System.out.println("Province is " + Province);

	stringSplit = Province.split(" - ");
	String Prov = stringSplit[0];
	// System.out.println(Prov);

	if (Prov.contains("ON") || Prov.contains("MB")) {
		provinceGroup = "ON,MB";
	} else if (Prov.contains("NL")) {
		provinceGroup = "NL";
	} else if (Prov.contains("SK")) {
		provinceGroup = "SK";
	} else {
		provinceGroup = "OTHER";
	}

	// Identifying the province
	while (cell.hasNext()) {

		value = cell.next();
		if (value.getStringCellValue().contains(provinceGroup)) {
			fcol = value.getColumnIndex();
			break;
		}

	}

	// Setting first column and last column for the table

	lcol = fcol + 7;

	// Reading through CV score row

	for (int c = fcol; c < (lcol + 1); c++) {

		String efs = sheet.getRow(2).getCell(c).getStringCellValue();
		value = sheet.getRow(2).getCell(c);

		if (efs.contains(">")) {
			lowefs = efs.replace(">=", "").trim();
			// System.out.println(lowefs);
			lef = Double.parseDouble(lowefs);
			// System.out.println(lef);
			if (cvScore >= lef) {
				col = value.getColumnIndex();
				break;

			}

		} else if (efs.contains("-")) {
			stringSplit = efs.split("-");
			lowefs = stringSplit[0];
			highefs = stringSplit[1];
			lef = Double.parseDouble(lowefs);
			hef = Double.parseDouble(highefs);

			if (cvScore >= lef) {
				col = value.getColumnIndex();
				// System.out.println(col);
				break;
			} else if (((563) <= cvScore) && (cvScore <= (575))) {
				col = fcol + 6;
				// System.out.println(col);

				break;
			}

		}

		// System.out.println(col+" is the column to be verified");

	}

	// Reading through Remaining Income given in the excel sheet

	int strcounter = 1; // Counter is intialized for NA counts

	for (int r = 4; r <= sheet.getLastRowNum(); r++) {

		try {
			if (sheet.getRow(r).getCell(col).getCellTypeEnum() == CellType.NUMERIC) {

				double remIn = sheet.getRow(r).getCell(col).getNumericCellValue();
				lastNumRow = r; // Row above NAs are stored as separate variable for calculation

				if (RemainingIncome < remIn) {
					rowNum = sheet.getRow(r - 1).getCell(col).getRowIndex();
					break;

				}

				if (r == sheet.getLastRowNum()) {
					rowNum = r;

				}

			}

			else if (sheet.getRow(r).getCell(col).getCellTypeEnum() == CellType.STRING) {
				int rowAboveNA = r - 1;
				double remInNaPrev = sheet.getRow(r - 1).getCell(col).getNumericCellValue();

				while (sheet.getRow(r + 1).getCell(col).getCellTypeEnum() == CellType.STRING) {
					strcounter++;
					;
					r++;
					if (r == sheet.getLastRowNum()) {
						rowNum = lastNumRow;
					}
				}

				double remInNaAfter = sheet.getRow(rowAboveNA + strcounter + 1).getCell(col).getNumericCellValue();

				if ((RemainingIncome > remInNaPrev) && (RemainingIncome < remInNaAfter)) {
					rowNum = rowAboveNA;
					break;
				}
			}

			strcounter = 1;
		}

		catch (IllegalStateException | NumberFormatException | NullPointerException e) {

		}

	}

	// Calculation QLA

	double ActualQLA;
	QLA = sheet.getRow(rowNum).getCell(lcol).getNumericCellValue();
	Thread.sleep(3000);

	if (QLA == 0.0) {
		ActualQLA = QLA;
	} else {
		ActualQLA = QLA + 100;
	}
	if (ActualQLA != ExpectedQLA) {
		if (loanType.contains("New")) {
			if (bkStrategy.contains("Bankruptcy QLA Decrease")) {
				ActualQLA = ActualQLA + bkDecreaseAmount;

			}
		}
	}
	
	
	System.out.println("Actual QLA :$" + ActualQLA);
	System.out.println("Expected QLA :$" + ExpectedQLA);

	Thread.sleep(3000);
	driver.switchTo().defaultContent();
	WebElement Int = driver.findElement(By.xpath(prop.getProperty("re-submit")));
	new Actions(driver).moveToElement(Int).perform();
	Thread.sleep(3000);
	JavascriptExecutor js = (JavascriptExecutor) driver;
	js.executeScript("document.body.style.zoom='70%'");
	Thread.sleep(3000);
	String screenShotPathforQLA = Screenshot.capture(driver, "CaculateQLA");
	js.executeScript("document.body.style.zoom='100%'");

	// Displaying QLA result

	if (ActualQLA == ExpectedQLA) {

		test.log(Status.PASS, MarkupHelper.createLabel("QLA Actual value :  $" + ActualQLA, ExtentColor.GREEN));
		test.log(Status.PASS,
				MarkupHelper.createLabel("QLA Expected value :  $ " + ExpectedQLA, ExtentColor.GREEN));

		test.log(Status.PASS,
				MarkupHelper.createLabel(" QLA Calculation is Matching with GDS Decision", ExtentColor.GREEN));
		test.log(Status.PASS, "Snapshot below: " + test.addScreenCaptureFromPath(screenShotPathforQLA));
		System.out.println("PASSED in QLA Verification");
	} else {
		System.out.println(ExpectedQLA + " is the expected value.");

		test.log(Status.FAIL, MarkupHelper.createLabel("QLA Actual value :  $" + ActualQLA, ExtentColor.RED));
		test.log(Status.FAIL, MarkupHelper.createLabel("QLA Expected value :  $ " + ExpectedQLA, ExtentColor.RED));

		test.log(Status.FAIL,
				MarkupHelper.createLabel(" QLA Calculation not Matching with GDS Decision", ExtentColor.RED));
		test.log(Status.FAIL, "Snapshot below: " + test.addScreenCaptureFromPath(screenShotPathforQLA));
		System.out.println("FAILED in QLA Verification");
	}

}
public void maxHA() throws IOException, InterruptedException {
	if (attemptNo == 0) {
		test = Extent.createTest("MAX H&A Calculation");
	}

	else {

		test = Extent.createTest("Resubmission Result: Attempt#  " + attemptNo + " - MAX H&A Calculation");
	}

	org.apache.poi.ss.usermodel.Sheet sheet;

	File file = new File(
			System.getProperty("user.dir") + "\\src\\main\\resources\\Excel\\UAT-SF-GDScalculation.xlsx");

	FileInputStream inputStream = new FileInputStream(file);

	Workbook workbook = new XSSFWorkbook(inputStream);

	sheet = workbook.getSheet("Max H&A");

	// For UPL
	double qla1 = sheet.getRow(3).getCell(0).getNumericCellValue();
	double qla2 = sheet.getRow(3).getCell(2).getNumericCellValue();
	double max1 = sheet.getRow(3).getCell(3).getNumericCellValue();
	double qla3 = sheet.getRow(4).getCell(0).getNumericCellValue();
	double qla4 = sheet.getRow(4).getCell(2).getNumericCellValue();
	double max2 = sheet.getRow(4).getCell(3).getNumericCellValue();
	double qla5 = sheet.getRow(5).getCell(0).getNumericCellValue();
	double qla6 = sheet.getRow(5).getCell(2).getNumericCellValue();
	double max3 = sheet.getRow(5).getCell(3).getNumericCellValue();
	double qla7 = sheet.getRow(6).getCell(0).getNumericCellValue();
	double qla8 = sheet.getRow(6).getCell(2).getNumericCellValue();
	double max4 = sheet.getRow(6).getCell(3).getNumericCellValue();
	double qla9 = sheet.getRow(7).getCell(0).getNumericCellValue();
	double qla10 = sheet.getRow(7).getCell(2).getNumericCellValue();
	double max5 = sheet.getRow(7).getCell(3).getNumericCellValue();
	// For SPL
	double qla11 = sheet.getRow(11).getCell(0).getNumericCellValue();
	double qla12 = sheet.getRow(11).getCell(2).getNumericCellValue();
	double max6 = sheet.getRow(11).getCell(3).getNumericCellValue();
	
	if(ExpectedQLA<qla1)
	{
		ActualMaxHA=0.0;
	}
	else if(ExpectedQLA>=qla1 && ExpectedQLA<=qla2)
	{
		ActualMaxHA=max1;
	}
	else if(ExpectedQLA>=qla3 && ExpectedQLA<=qla4)
	{
		ActualMaxHA=max2;
	}
	else if(ExpectedQLA>=qla5 && ExpectedQLA<=qla6)
	{
		ActualMaxHA=max3;
	}
	else if(ExpectedQLA>=qla7 && ExpectedQLA<=qla8)
	{
		ActualMaxHA=max4;
	}
	else if(ExpectedQLA>=qla9 && ExpectedQLA<=qla10)
	{
		ActualMaxHA=max5;
	}
	else if(ExpectedQLA>=qla11 && ExpectedQLA<=qla12)
	{
		ActualMaxHA=max6;
	}
	driver.switchTo().defaultContent();
	WebElement Int = driver.findElement(By.xpath(prop.getProperty("re-submit")));
	new Actions(driver).moveToElement(Int).perform();
	Thread.sleep(3000);
	JavascriptExecutor js = (JavascriptExecutor) driver;
	js.executeScript("document.body.style.zoom='70%'");
	Thread.sleep(3000);

	String screenShotPathforInterestRate = Screenshot.capture(driver, "CaculateInterestRate");
	js.executeScript("document.body.style.zoom='100%'");
	driver.switchTo().frame(driver.findElement(By.xpath(prop.getProperty("switchIframe"))));

	// Displaying Interest Rate result
	System.out.println("Actual MaxH&A: " + ActualMaxHA);
	System.out.println("Expected MaxH&A: " + ExpectedMaxHA);

	if (ExpectedMaxHA == ActualMaxHA) {

		test.log(Status.PASS, MarkupHelper.createLabel("MaxH&A Actual value : " + ActualMaxHA + "%",
				ExtentColor.GREEN));
		test.log(Status.PASS, MarkupHelper.createLabel("MaxH&A Expected value : " + ExpectedMaxHA + "%",
				ExtentColor.GREEN));

		test.log(Status.PASS, MarkupHelper.createLabel(" MaxH&A Calculation is Matching with GDS Decision",
				ExtentColor.GREEN));
		test.log(Status.PASS, "Snapshot below: " + test.addScreenCaptureFromPath(screenShotPathforInterestRate));
		System.out.println("PASSED in MaxH&A Verification");
	} else {
		

		test.log(Status.FAIL, MarkupHelper.createLabel("MaxH&A Actual value : " + ActualMaxHA + "%",
				ExtentColor.RED));
		test.log(Status.FAIL, MarkupHelper.createLabel("MaxH&A Expected value : " + ExpectedMaxHA + "%",
				ExtentColor.RED));

		test.log(Status.FAIL, MarkupHelper.createLabel(" MaxH&A Calculation not Matching with GDS Decision",
				ExtentColor.RED));
		test.log(Status.FAIL, "Snapshot below: " + test.addScreenCaptureFromPath(screenShotPathforInterestRate));
		System.out.println("FAILED in MaxH&A Verification");
	}
	Thread.sleep(3000);
}
public void riskFactor() throws InterruptedException {
	// Risk Table
	System.out.println("---------------");
	System.out.println("Risk Factors");
	driver.manage().timeouts().implicitlyWait(0, TimeUnit.SECONDS);

	WebDriverWait waitLoad = new WebDriverWait(driver, 360, 0000);
	waitLoad.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("getRiskfactor"))));

	WebElement Decisiontable = driver.findElement(By.xpath(prop.getProperty("RiskTable")));
	Thread.sleep(2000);
	List<WebElement> rowValsDecision = Decisiontable.findElements(By.tagName("tr"));
	int rowNumDecision = Decisiontable.findElements(By.tagName("tr")).size();

	int colNumDecision = driver.findElements(By.xpath(prop.getProperty("colRisk"))).size();

	for (int i = 0; i < rowNumDecision; i++) {
		// Get each row's column values by tag name
		List<WebElement> colValsDecision = rowValsDecision.get(i).findElements(By.tagName("td"));
		System.out.println("---------------");
		String reasoncode = colValsDecision.get(0).getText();
		System.out.println("Code: " + reasoncode);
		String status = colValsDecision.get(1).getText();
		System.out.println("Description: " + status);
		String ReqOverride = colValsDecision.get(2).getText();
		System.out.println("Weight: " + ReqOverride);
		String group = colValsDecision.get(3).getText();
		System.out.println("Applies To: " + group);

		System.out.println("---------------");
	}
}

public void decisionCode() throws InterruptedException {
	// Decision Table
	System.out.println("Decision Table");
	driver.manage().timeouts().implicitlyWait(0, TimeUnit.SECONDS);

	WebDriverWait waitLoad = new WebDriverWait(driver, 360, 0000);
	waitLoad.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("getDesisioncode"))));

	WebElement Decisiontable = driver.findElement(By.xpath(prop.getProperty("DecisionTable")));
	Thread.sleep(2000);
	List<WebElement> rowValsDecision = Decisiontable.findElements(By.tagName("tr"));
	int rowNumDecision = Decisiontable.findElements(By.tagName("tr")).size();

	int colNumDecision = driver.findElements(By.xpath(prop.getProperty("colDecision"))).size();

	for (int i = 0; i < rowNumDecision; i++) {
		// Get each row's column values by tag name
		List<WebElement> colValsDecision = rowValsDecision.get(i).findElements(By.tagName("td"));
		System.out.println("---------------");
		String reasoncode = colValsDecision.get(0).getText();
		System.out.println("Reason Code: " + reasoncode);
		String status = colValsDecision.get(1).getText();
		System.out.println("Status: " + status);
		String ReqOverride = colValsDecision.get(2).getText();
		System.out.println("Required Override Level: " + ReqOverride);
		String group = colValsDecision.get(3).getText();
		System.out.println("Group: " + group);
		String description = colValsDecision.get(4).getText();
		System.out.println("Description: " + description);

		System.out.println("---------------");
	}

}
public void SecondPopup() throws Exception {
	attemptNo++;
	driver.switchTo().defaultContent();
	WebElement page1 = driver.findElement(By.xpath(prop.getProperty("uplPage1")));
	new Actions(driver).moveToElement(page1).perform();

	JFrame frmOpt = new JFrame(); // We are declaring the frame
	frmOpt.setAlwaysOnTop(true);// This is the line for displaying it above all windows

	Thread.sleep(1000);
	String s = "<html>Press 1 for Re-Submission with Co Applicant<br>Press 2 for Re-Submission with Removal of Co Applicant<br>";
	s += "Press 3 for Results</html>";
	JLabel label = new JLabel(s);
	JTextPane jtp = new JTextPane();
	jtp.setSize(new Dimension(480, 10));
	jtp.setPreferredSize(new Dimension(480, jtp.getPreferredSize().height));
	label.setFont(new Font("Arial", Font.BOLD, 26));
	UIManager.put("OptionPane.minimumSize", new Dimension(700, 300));
	UIManager.put("TextField.font", new FontUIResource(new Font("Verdana", Font.BOLD, 20)));
	// Getting Input from user

	String option = JOptionPane.showInputDialog(frmOpt, label);

	int useroption = Integer.parseInt(option);

	switch (useroption) {

	case 1:

		// Function for Add Co App
		System.out.println("Re-Submission with Co Applicant");
		resubmitForDecision();

		break;

	case 2:

		// Function for Removal Co App
		System.out.println("Re-Submission with Removal of Co Applicant");
		removeCoApp();
		break;
	case 3:

		// Function for Removal Co App
		System.out.println("Results");
		if (attemptNo == 0) {
			test = Extent.createTest("Resubmission - Confirmation ");
			test.info(" The test run complete. Please review test result(s)");
		}

		else {

			test = Extent.createTest("Resubmission - Confirmation ");
			test.info(" The test run complete. Please review test result(s)");
		}

		Thread.sleep(3000);

		driver.close();
		driver.quit();
		break;

	}

}
public void resubmitForDecision() throws Exception {

	System.out.println(attemptNo);

	
	firstPopup();
	// SWITCH IFRAME DEFAULT
	driver.switchTo().defaultContent();
	getAppDetails();
	calculateIncome();
	Thread.sleep(2000);
	calculateLiability();
	driver.switchTo().defaultContent();
	getAppTimestampLogs();

	// Interest rate Calculations

	interestRateCalculation();
	checkContributer();
	Thread.sleep(3000);
	// Logging in as FSR
	 if (LoanType.contains("Increase")) {
		
		calculateQLA();
	}
	maxHA();
	ReasonCode();
	SecondPopup();
}

public void removeCoApp() throws Exception {
	
	firstPopup();
	// SWITCH IFRAME DEFAULT
	driver.switchTo().defaultContent();
	Thread.sleep(3000);
	getApplicationDetails();
	calculateAppIncome();
	Thread.sleep(2000);
	calculateAppLiability();

	Thread.sleep(3000);
	// Logging in as FSR
	 if (LoanType.contains("Increase")) {
		
		calculateQLA();
	}
	maxHA();
	ReasonCode();
	SecondPopup();

}
public void getApplicationDetails() throws InterruptedException, IOException {
	
	WebDriverWait waitLoad = new WebDriverWait(driver, 360, 0000);
	waitLoad.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("getprovince"))));
	Province = driver.findElement(By.xpath(prop.getProperty("getprovince"))).getText();
	ApplicationID = driver.findElement(By.xpath(prop.getProperty("getappid"))).getText();
	appType = driver.findElement(By.xpath(prop.getProperty("getapptype"))).getText();
	LoanType = driver.findElement(By.xpath(prop.getProperty("loantype"))).getText();
	System.out.println("App Type :" + appType);

	driver.switchTo().frame(driver.findElement(By.xpath(prop.getProperty("switchIframe"))));
	Thread.sleep(4000);
	IntRate = driver.findElement(By.xpath(prop.getProperty("get%"))).getText();
	ExpInt = Double.parseDouble(IntRate.replace("%", ""));

	String intqla = driver.findElement(By.xpath(prop.getProperty("getQla"))).getText();
	String qla = intqla.replace(",", "");
	
	ExpectedQLA = Double.parseDouble(qla.replace("$", ""));
	String intHA=driver.findElement(By.xpath(prop.getProperty("getmaxHA"))).getText();
	String HA = intHA.replace(",", "");

	ExpectedMaxHA = Double.parseDouble(HA.replace("$", ""));
	String cvscore = driver.findElement(By.xpath(prop.getProperty("getcvscore"))).getText();
	cvScore = Integer.parseInt(cvscore);
	Strategy = driver.findElement(By.xpath(prop.getProperty("adstrategy"))).getText();
	if (LoanType.contains("New")) {
		cabKey = driver.findElement(By.xpath("//th[contains(text(),'CAB Key')]/following-sibling::td[1]/span"))
				.getText();
		bkStrategy = driver.findElement(By.xpath(prop.getProperty("bkstrategy"))).getText();
	}
	riskFactor();
	decisionCode();
	System.out.println("ApplicationID: " + ApplicationID);
	System.out.println("Province: " + Province);

	System.out.println("Interest Rate :" + ExpInt);
	System.out.println("CV Score: " + cvScore);
	System.out.println("QLA: " + ExpectedQLA);
	System.out.println("Strategy :" + Strategy);
}
public void calculateAppIncome() throws InterruptedException, IOException {

	System.out.println("Resubmission attempt #" + attemptNo);
	if (attemptNo == 0) {
		test = Extent.createTest("Total Income Calcuation");
	}

	else {

		test = Extent.createTest("Resubmission Result: Attempt#  " + attemptNo + " - Total Income Calcuation");
	}

	// SWITCH IFRAME DEFAULT
	driver.switchTo().defaultContent();

	// Income Table
	Thread.sleep(2000);

	WebElement Incometable = driver.findElement(By.xpath(prop.getProperty("incomeTable")));
	new Actions(driver).moveToElement(Incometable).perform();

	Thread.sleep(2000);
	List<WebElement> rowValsIncome = Incometable.findElements(By.tagName("tr"));
	int rowNumIncome = Incometable.findElements(By.tagName("tr")).size();

	int colNumIncome = driver.findElements(By.xpath(prop.getProperty("colincome"))).size();
	System.out.println("Income Table");

	double IncomeValue = 0;
	String str = null;
	for (int i = 0; i < rowNumIncome; i++) {
		double subTotal = 0;
		double subValue = 0; // Get each row's column values by tag name
		List<WebElement> colValsIncome = rowValsIncome.get(i).findElements(By.tagName("td"));
		String IncomeFreq = colValsIncome.get(2).getText();
		System.out.println(IncomeFreq);
		String IncomeAmount = colValsIncome.get(3).getText();
		System.out.println(IncomeAmount);

		if (IncomeAmount.contains(","))

		{
			str = IncomeAmount.replace(",", "");
			subValue = Double.parseDouble(str.replace("$", ""));
		} else {
			subValue = Double.parseDouble(IncomeAmount.replace("$", ""));
		}
		System.out.println(subValue);

		if (IncomeFreq.equals("Weekly")) {
			subTotal = subValue * 4.0;

		}

		else if (IncomeFreq.equals("Bi-Weekly")) {
			subTotal = subValue * 2.0;

		} else if (IncomeFreq.equals("Semi-Monthly")) {
			subTotal = subValue * 2.0;

		} else if (IncomeFreq.equals("Monthly")) {
			subTotal = subValue * 1.0;

		}

		IncomeValue += subTotal;
		System.out.println("---------------");

	}

	System.out.println("Actual Income: $" + IncomeValue);
	// test = Extent.createTest(" Calculate Income");

	// ActualIncome = "$"+IncomeValue;
	// test.info("Actual Income ="+ActualIncome);

	// Income field Comparison
	Thread.sleep(2000);
	JavascriptExecutor js = (JavascriptExecutor) driver;
	js.executeScript("document.body.style.zoom='80%'");
	Thread.sleep(3000);
	String screenShotPath = Screenshot.capture(driver, "CalculateIncome");
	js.executeScript("document.body.style.zoom='100%'");
	Thread.sleep(5000);
	driver.switchTo().frame(driver.findElement(By.xpath(prop.getProperty("switchIframe"))));

	WebElement Totalinc = driver.findElement(By.xpath(prop.getProperty("totalincome")));

	TotalIncomeAmount = Totalinc.getText();
	// test.info("Expected Income ="+TotalIncomeAmount);
	String st = TotalIncomeAmount.replace(",", "");
	TotalIncome = Double.parseDouble(st.replace("$", ""));

	System.out.println("Expected Income: $" + TotalIncome);
	if (IncomeValue == TotalIncome) {

		test.log(Status.PASS,
				MarkupHelper.createLabel(" Total Income  :Actual Value =  $" + IncomeValue, ExtentColor.GREEN));
		test.log(Status.PASS,
				MarkupHelper.createLabel("Total Income  :Expected Value =  $" + TotalIncome, ExtentColor.GREEN));
		test.log(Status.PASS, MarkupHelper.createLabel("Income is Matching with GDS Decision", ExtentColor.GREEN));
		test.log(Status.PASS, "Snapshot below: " + test.addScreenCaptureFromPath(screenShotPath));
		System.out.println("Income:Passed");

	} else {

		test.log(Status.FAIL,
				MarkupHelper.createLabel("Total Income : Actual Value =  $" + IncomeValue, ExtentColor.RED));
		test.log(Status.FAIL,
				MarkupHelper.createLabel("Total Income : Expected Value =  $" + TotalIncome, ExtentColor.RED));
		test.log(Status.FAIL, MarkupHelper.createLabel("Income Not Matching with GDS Decision", ExtentColor.RED));
		test.log(Status.FAIL, "Snapshot below: " + test.addScreenCaptureFromPath(screenShotPath));

		System.out.println("Income Not Match");
	}

}

public void calculateAppLiability() throws InterruptedException, IOException {
	Thread.sleep(4000);
	if (attemptNo == 0) {
		test = Extent.createTest("Total Liability Calcuation");
	}

	else {

		test = Extent.createTest("Resubmission Result: Attempt#  " + attemptNo + " - Total Liability Calcuation");
	}

	// Liabilities Table
	System.out.println("Liabilities Table");
	// SWITCH IFRAME DEFAULT
	driver.switchTo().defaultContent();

	driver.manage().timeouts().implicitlyWait(0, TimeUnit.SECONDS);
	WebElement Liabilitiestable = driver.findElement(By.xpath(prop.getProperty("liabilitiesTable")));
	new Actions(driver).moveToElement(Liabilitiestable).perform();
	Thread.sleep(2000);
	List<WebElement> rowValsLiabilities = Liabilitiestable.findElements(By.tagName("tr"));
	int rowNumLiabilities = Liabilitiestable.findElements(By.tagName("tr")).size();

	int colNumLiabilities = driver.findElements(By.xpath(prop.getProperty("colliabilities"))).size();
	System.out.println("Total number of rows = " + rowNumLiabilities);
	System.out.println("Total number of columns = " + colNumLiabilities);

	double LiabilitiesValue = 0;
	for (int i = 0; i < rowNumLiabilities; i++) {
		// Get each row's column values by tag name
		List<WebElement> colValsLiabilities = rowValsLiabilities.get(i).findElements(By.tagName("td"));
		String LiabilitiesAmount = colValsLiabilities.get(3).getText();
		System.out.println(LiabilitiesAmount);
		String str = LiabilitiesAmount.replace(",", "");
		double subTotal = Double.parseDouble(str.replace("$", ""));
		System.out.println(subTotal);

		LiabilitiesValue += subTotal;

		System.out.println("---------------");
	}

	System.out.println(LiabilitiesValue);

	// Rent Expense
	System.out.println("Rent Expense Table");
	driver.manage().timeouts().implicitlyWait(0, TimeUnit.SECONDS);
	WebElement Renttable = driver.findElement(By.xpath(prop.getProperty("rentTable")));
	new Actions(driver).moveToElement(Renttable).perform();
	Thread.sleep(2000);
	List<WebElement> rowRent = Renttable.findElements(By.xpath(prop.getProperty("rowrent")));

	int rowNumRent = Renttable.findElements(By.xpath(prop.getProperty("rowrent"))).size();

	int colNumRent = driver.findElements(By.xpath(prop.getProperty("colrent"))).size();
	System.out.println("Total number of columns = " + colNumRent);
	double RentValue = 0;
	for (int i = 0; i < rowNumRent; i++) {
		// Get each row's column values by tag name
		List<WebElement> colValsRent = rowRent.get(i).findElements(By.tagName("td"));
		String RentAmount = colValsRent.get(1).getText();
		System.out.println(RentAmount);
		String str = RentAmount.replace(",", "");
		double subTotal = Double.parseDouble(str.replace("$", ""));

		System.out.println(subTotal);

		RentValue += subTotal;

		System.out.println("---------------");
	}

	System.out.println(RentValue);

	// Mortgages Table
	System.out.println("Mortgages Table");

	double MortgagesValue = 0;
	driver.manage().timeouts().implicitlyWait(0, TimeUnit.SECONDS);
	int colNumMortgages = driver.findElements(By.xpath(prop.getProperty("colmortgages"))).size();

	WebElement Mortgagestable = driver.findElement(By.xpath(prop.getProperty("mortgagesTable")));

	List<WebElement> rowMortgages = Mortgagestable.findElements(By.xpath(prop.getProperty("rowmortgages")));

	int rowNumMortgages = Mortgagestable.findElements(By.xpath(prop.getProperty("rowmortgages"))).size();

	System.out.println("Total number of columns = " + colNumMortgages);

	for (int i = 0; i < rowNumMortgages; i++) {
		// Get each row's column values by tag name
		List<WebElement> colValsMortgages = rowMortgages.get(i).findElements(By.tagName("td"));
		String RentAmount = colValsMortgages.get(3).getText();
		System.out.println(RentAmount);
		String str = RentAmount.replace(",", "");
		double subTotal = Double.parseDouble(str.replace("$", ""));

		System.out.println(subTotal);

		MortgagesValue += subTotal;

		System.out.println("---------------");
	}

	System.out.println(MortgagesValue);

	// Other Table
	System.out.println("Other Table");
	driver.manage().timeouts().implicitlyWait(0, TimeUnit.SECONDS);
	WebElement Othertable = driver.findElement(By.xpath(prop.getProperty("otherTable")));
	new Actions(driver).moveToElement(Othertable).perform();

	List<WebElement> rowOther = Othertable.findElements(By.xpath(prop.getProperty("rowother")));

	int rowNumOther = Othertable.findElements(By.xpath(prop.getProperty("rowother"))).size();

	int colNumOther = driver.findElements(By.xpath(prop.getProperty("colother"))).size();
	System.out.println("Total number of columns = " + colNumOther);
	double OtherValue = 0;
	for (int i = 0; i < rowNumOther; i++) {
		// Get each row's column values by tag name
		List<WebElement> colValsOther = rowOther.get(i).findElements(By.tagName("td"));
		String RentAmount = colValsOther.get(3).getText();
		System.out.println(RentAmount);
		String str = RentAmount.replace(",", "");
		double subTotal = Double.parseDouble(str.replace("$", ""));

		System.out.println(subTotal);

		OtherValue += subTotal;

		System.out.println("---------------");
	}

	System.out.println(OtherValue);

	double TotalLaibility = LiabilitiesValue + RentValue + OtherValue + MortgagesValue;
	System.out.println("Actual Liabilities: $" + TotalLaibility);
	// test = Extent.createTest(" Calculate Liability");
	// String ActualLiability = "$"+TotalLaibility;
	// test.info("Actual Liability ="+ActualLiability);

	// Get Total Debt value - Liabilities Comparison
	Thread.sleep(2000);

	JavascriptExecutor js = (JavascriptExecutor) driver;
	js.executeScript("document.body.style.zoom='90%'");
	Thread.sleep(3000);

	String screenShotPath = Screenshot.capture(driver, "CaculateLiability");
	js.executeScript("document.body.style.zoom='100%'");
	Thread.sleep(3000);

	driver.switchTo().frame(driver.findElement(By.xpath(prop.getProperty("switchIframe"))));
	WebElement Totalpay = driver.findElement(By.xpath(prop.getProperty("totalpayment")));

	String TotalDebtAmount = Totalpay.getText();
	// test.info("Expected Liability ="+TotalDebtAmount);
	String str = TotalDebtAmount.replace(",", "");
	TotalDebt = Double.parseDouble(str.replace("$", ""));

	System.out.println("Expected Liabilities: $" + TotalDebt);
	if (TotalLaibility == TotalDebt) {
		System.out.println("Laibilities:Passed");
		test.log(Status.PASS, MarkupHelper.createLabel("Total Liability - Actual Value   =  $" + TotalLaibility,
				ExtentColor.GREEN));
		test.log(Status.PASS,
				MarkupHelper.createLabel("Total Liability - Expected Value =  $" + TotalDebt, ExtentColor.GREEN));

		test.log(Status.PASS,
				MarkupHelper.createLabel("Liability is Matching with GDS Decision ", ExtentColor.GREEN));
		test.log(Status.PASS, "Snapshot below: " + test.addScreenCaptureFromPath(screenShotPath));

	} else {
		System.out.println("Laibilities:Failed");

		test.log(Status.FAIL, MarkupHelper.createLabel("Total Liability - Actual Value   =  $" + TotalLaibility,
				ExtentColor.RED));
		test.log(Status.FAIL,
				MarkupHelper.createLabel("Total Liability - Expected Value =  $" + TotalDebt, ExtentColor.RED));

		test.log(Status.FAIL,
				MarkupHelper.createLabel("Liability is not Matching with GDS Decision ", ExtentColor.RED));
		test.log(Status.FAIL, "Snapshot below: " + test.addScreenCaptureFromPath(screenShotPath));

		// Assert.assertTrue(false);
	}
}
}
