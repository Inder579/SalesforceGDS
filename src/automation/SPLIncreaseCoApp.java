package automation;

import org.openqa.selenium.WebDriver;
import java.awt.Dimension;
import java.awt.Font;
import java.awt.HeadlessException;
import java.awt.RenderingHints.Key;
import java.awt.Toolkit;
import java.awt.datatransfer.DataFlavor;
import java.awt.datatransfer.UnsupportedFlavorException;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.IOException;
import java.io.InputStream;
import java.sql.Date;
import java.sql.Timestamp;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Locale;
import java.util.NoSuchElementException;
import java.util.concurrent.TimeUnit;

import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JTextPane;
import javax.swing.UIManager;
import javax.swing.plaf.FontUIResource;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.dom4j.DocumentException;
import org.joda.time.DateTime;
import org.joda.time.LocalDateTime;
import org.joda.time.format.DateTimeFormat;
import org.json.simple.JSONArray;
import org.json.simple.JSONObject;
import org.json.simple.parser.JSONParser;
import org.json.simple.parser.ParseException;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.StaleElementReferenceException;
import org.openqa.selenium.TimeoutException;
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
import com.monitorjbl.xlsx.StreamingReader;

import resources.BrowserDriver;
import resources.Screenshot;

public class SPLIncreaseCoApp extends BrowserDriver {

	public static int attemptNo = 0;
	public String screenShotPathforInterestRate;
	public WebDriver driver;
	int cvScore,cvScoreCoapp, BehaviourScore;
	public String ActualIncome, appType, loanType, splloanType, Province,CoAppProvince, ApplicationID, MorgagePayment,behscore, cobehscore,
	appbehscore;
	public String TotalIncomeAmount, IntRate, cabKey, qlaStrategy, applicationType, cabKeyCoApp, contributer;
	public double TotalIncome, RemainingIncome, TotalDebt, ExpectedQLA, ExpInt, SPLltv, Maxltv, HomeEquity, PropertyVal, ActualQLAcoapp,
	ActualFinalCoApp;

	String lowefs, highefs, Prov, provinceGroup, bkStrategy, ps = "", code = null, propertyType = "",
			propertyLocation = "";
	double lef, hef, calRemIn, QLA, remIn, remInNaPrev, remInNaAfter, LtvMax, ActualMaxHA, ExpectedMaxHA, IncomeValue, TotalLaibility;

	int fcol, lcol, col, coldiff, rowNum, RiskGroup, SPLTotalDebt, lastNumRow, bkDecreaseAmount, RandomNumberResponse, length;
	String stringSplit[], Strategy;

	String lastname, firstname, address, city, dob, clprod, loanpurpose, hearabout, Referral, livingsituation, email,
			lengthofstay;

	String phone, loanamount, landlordname, landlordnumber, Employername, Employerposition, Incomeamt, Incomefreq,
			Employmentstatus, Supervisorname, Supervisornumber, lengthofemployment, previousemployer,
			lengthpreviousemployer, preferedLang;
	String QualifiedLoanAmount, totalincomeFusion, totaldebtFusion, MaximumLTV, ApplicantEFSCVScore, UplStrategyFusion,
			CurrentAddress, postalcode, MortgageBalances, PropertyValue, url, IncomeLiabilityScreen, QLAInterestScreen,
			tsdate, SPLBuydown, masterID,ts0, ts1, ts2, ReasonCodeSPLFullScreen1, ReasonCodeSPLFullScreen2,
			ReasonCodeSPLFullScreen3, MaxHASPLFullScreen;
	
	
	@BeforeTest
	public void initialize1() throws IOException {

		driver = browser();

	}
	@Test()
	public void m1() throws Exception {

		// Login as Admin
		loginAsAdmin();
	    waitForFirstSubmission();
	    firstPopup();
	    
		getAddress();
		getPartyDetailsCoApp();
		getCollateral();

		getAppDetailsCoApp();
		
		calculateIncome();
		calculateSPLLiability();
		premulesoft();
		mulesoft();
		checkContributerincrease();
        splIncreaseInterestRate();
		splLTV();
		splIncreaseRemInc();
		MaxAfforableIncreaseQLA();
		urbancode();
		splFinalIncreaseQLA();
		maxHA();
		ReasonCode();
		Thread.sleep(4000);
		SecondPopupSpl();
	}

	public void loginAsAdmin() throws InterruptedException, IOException, UnsupportedFlavorException {
		 driver.get(prop.getProperty("sfUrl"));

	//	driver.get("https://goeasy--uatpreview.lightning.force.com/lightning/r/genesis__Applications__c/a5yf0000000HsxbAAC/view");
		// driver.get("https://goeasy--goeasyqasb.my.salesforce.com/");
		Thread.sleep(2000);
		WebDriverWait waitLoad = new WebDriverWait(driver, 360, 0000);
		waitLoad.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("username"))));
		// driver.findElement(By.xpath(prop.getProperty("username"))).sendKeys(decodeString(prop.getProperty("AdminEmail")));
		driver.findElement(By.xpath(prop.getProperty("username"))).sendKeys(prop.getProperty("AdminEmailFusion"));
		Thread.sleep(2000);
		// driver.findElement(By.cssSelector(prop.getProperty("password"))).sendKeys(decodeString(prop.getProperty("AdminPassword")));
		driver.findElement(By.cssSelector(prop.getProperty("password")))
				.sendKeys(prop.getProperty("AdminPasswordFusion"));
		driver.findElement(By.xpath(prop.getProperty("clicklogin"))).click();

		// https://goeasy--goeasyqasb.lightning.force.com/lightning/r/genesis__Applications__c/a5y1h0000002C2NAAU/view
		System.out.println("Logged in As Admin");
	}
	public void premulesoft() throws InterruptedException
	{
		// mulesoft
		driver.get(prop.getProperty("mulesoft"));
		// Usecustomdomain
		WebDriverWait waitcustom = new WebDriverWait(driver, 360, 0000);
		waitcustom.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("Usecustomdomain"))));
		// driver.findElement(By.xpath(prop.getProperty("username"))).sendKeys(decodeString(prop.getProperty("AdminEmail")));
		Thread.sleep(2000);
		driver.findElement(By.xpath(prop.getProperty("Usecustomdomain"))).click();
		// customDomain
		WebDriverWait waitcustomDomain = new WebDriverWait(driver, 360, 0000);
		waitcustomDomain
				.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("customDomain"))));
		Thread.sleep(2000);
		// driver.findElement(By.xpath(prop.getProperty("username"))).sendKeys(decodeString(prop.getProperty("AdminEmail")));
		driver.findElement(By.xpath(prop.getProperty("customDomain"))).sendKeys(prop.getProperty("Organizationdomain"));
		// ContinueOrg
		driver.findElement(By.xpath(prop.getProperty("ContinueOrg"))).click();
	}

	public void mulesoft() throws Exception {

		/*
		 * Timestamp timeStamp = new Timestamp(System.currentTimeMillis()); String
		 * Time=timeStamp.toString(); System.out.println(timeStamp); String[] arrSplit =
		 * Time.split(" "); String date = arrSplit[0]; String time = arrSplit[1];
		 * System.out.println(date+" "+time);
		 */
		
		
		// RuntimeManager
		WebDriverWait waitRuntimeManager = new WebDriverWait(driver, 360, 0000);
		waitRuntimeManager
				.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("RuntimeManager"))));
		driver.findElement(By.xpath(prop.getProperty("RuntimeManager"))).click();

		// SearchApplications
		WebDriverWait waitSearchApplications = new WebDriverWait(driver, 360, 0000);
		waitSearchApplications
				.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("SearchApplications"))));
		// driver.findElement(By.xpath(prop.getProperty("username"))).sendKeys(decodeString(prop.getProperty("AdminEmail")));
		driver.findElement(By.xpath(prop.getProperty("SearchApplications"))).sendKeys(prop.getProperty("clsdev"));

		// clsdevclick

		WebDriverWait waitclsdevclick = new WebDriverWait(driver, 360, 0000);
		waitclsdevclick.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("clsdevclick"))));
		driver.findElement(By.xpath(prop.getProperty("clsdevclick"))).click();

		// Logs
		WebDriverWait waitLogs = new WebDriverWait(driver, 360, 0000);
		
		
		
		waitLogs.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("Logs"))));
    	waitLogs.until(ExpectedConditions.elementToBeClickable(By.xpath(prop.getProperty("Logs"))));
    	WebElement log = driver.findElement(By.xpath(prop.getProperty("Logs")));
		Actions act = new Actions(driver);
		
        int attempts = 0;
        while (attempts < 3)
        {
            try
            {
            	Thread.sleep(5000);
            	WebDriverWait waitLog = new WebDriverWait(driver, 360, 0000);
        		waitLog.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("Logs"))));
        		JavascriptExecutor executor = (JavascriptExecutor)driver;
        		executor.executeScript("arguments[0].click();", log);
      
                break;
            }
            catch (StaleElementReferenceException e)
            {
                System.out.println("StaleElementReference");
                driver.get(prop.getProperty("mulesoft"));
                mulesoft();
                checkContributerincrease();
                splIncreaseInterestRate();
    			splLTV();
    			splIncreaseRemInc();
    			MaxAfforableIncreaseQLA();
    			urbancode();
    			splFinalIncreaseQLA();
    			maxHA();
    			ReasonCode();
    			Thread.sleep(4000);
    			SecondPopupSpl();
                
            }
            attempts++;
        }
       
		
//		act.clickAndHold();
//		act.release().perform();

		/*
		 * // getlogs // WebDriverWait waitgetlogs = new WebDriverWait(driver, 360,
		 * 0000); //
		 * waitgetlogs.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop
		 * .getProperty("getlogs")))); String logs =
		 * driver.findElement(By.xpath(prop.getProperty("getlogs"))).getText();
		 * System.out.println(logs);
		 */
		// closedeploy
		WebDriverWait waitclosedeploy = new WebDriverWait(driver, 360, 0000);
		waitclosedeploy.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("closedeploy"))));
		driver.findElement(By.xpath(prop.getProperty("closedeploy"))).click();

		try
		{
		
		Thread.sleep(6000);

		// searchlogs
		WebDriverWait waitsearchlogs = new WebDriverWait(driver, 360, 0000);
		waitsearchlogs.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("searchlogs"))));
		driver.findElement(By.xpath(prop.getProperty("Advanced"))).click();
		// Enter Date and time
		WebDriverWait waitstartDateInput = new WebDriverWait(driver, 360, 0000);
		waitstartDateInput
				.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("startDateInput"))));
		driver.findElement(By.xpath(prop.getProperty("startDateInput"))).sendKeys(tsdate);
		driver.findElement(By.xpath(prop.getProperty("endDateInput"))).sendKeys(tsdate);
		driver.findElement(By.xpath(prop.getProperty("startTime"))).sendKeys(ts0);
		driver.findElement(By.xpath(prop.getProperty("endTime"))).sendKeys(ts0);
		
		driver.findElement(By.xpath(prop.getProperty("searchlogs"))).sendKeys(masterID); // clickarrow
		driver.findElement(By.xpath(prop.getProperty("clickarrow"))).click();
		WebDriverWait waitdebugpriority = new WebDriverWait(driver, 360, 0000);
		waitdebugpriority
				.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("debugpriority"))));
		driver.findElement(By.xpath(prop.getProperty("debugpriority"))).click();
		// Apply
		driver.findElement(By.xpath(prop.getProperty("Apply"))).click();
		Thread.sleep(2000);
		driver.findElement(By.xpath(prop.getProperty("searchlogs"))).sendKeys(Keys.ENTER);

		Thread.sleep(8000);
		List<WebElement> responsefiles = driver.findElements(By.xpath(prop.getProperty("listdebug")));
		WebElement target = responsefiles.get(0);
		Thread.sleep(2000);
		act.moveToElement(target);
		Thread.sleep(2000);
		act.clickAndHold();
		act.release().perform();
		
		}

		catch(IndexOutOfBoundsException e)
		{
			Thread.sleep(6000);

			// searchlogs
			WebDriverWait waitsearchlogs = new WebDriverWait(driver, 360, 0000);
			waitsearchlogs.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("searchlogs"))));
			driver.findElement(By.xpath(prop.getProperty("Advanced"))).click();
			// Enter Date and time
			WebDriverWait waitstartDateInput = new WebDriverWait(driver, 360, 0000);
			waitstartDateInput
					.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("startDateInput"))));
			driver.findElement(By.xpath(prop.getProperty("startDateInput"))).clear();
			driver.findElement(By.xpath(prop.getProperty("startDateInput"))).sendKeys(tsdate);
			driver.findElement(By.xpath(prop.getProperty("endDateInput"))).clear();
			driver.findElement(By.xpath(prop.getProperty("endDateInput"))).sendKeys(tsdate);
			driver.findElement(By.xpath(prop.getProperty("startTime"))).clear();
			driver.findElement(By.xpath(prop.getProperty("startTime"))).sendKeys(ts1);
			driver.findElement(By.xpath(prop.getProperty("endTime"))).clear();
			driver.findElement(By.xpath(prop.getProperty("endTime"))).sendKeys(ts2);
			driver.findElement(By.xpath(prop.getProperty("clickarrow"))).click();
			WebDriverWait waitdebugpriority = new WebDriverWait(driver, 360, 0000);
			waitdebugpriority
					.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("debugpriority"))));
			driver.findElement(By.xpath(prop.getProperty("debugpriority"))).click();
			// Apply
			driver.findElement(By.xpath(prop.getProperty("Apply"))).click();
			
			Thread.sleep(7000);
			List<WebElement> responsefiles = driver.findElements(By.xpath(prop.getProperty("listdebug")));
			WebElement target = responsefiles.get(0);
			Thread.sleep(2000);
			act.moveToElement(target);
			Thread.sleep(2000);
			act.clickAndHold();
			act.release().perform();
		}
		// **********************************************************
		/*
		 * //Advanced
		 * driver.findElement(By.xpath(prop.getProperty("Advanced"))).click();
		 * //Lasthour WebDriverWait waitLasthour = new WebDriverWait(driver, 360, 0000);
		 * waitLasthour.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(
		 * prop.getProperty("Lasthour"))));
		 * driver.findElement(By.xpath(prop.getProperty("Lasthour"))).click();
		 * 
		 * driver.findElement(By.xpath(prop.getProperty("searchlogs"))).sendKeys(
		 * masterID); //clickarrow
		 * driver.findElement(By.xpath(prop.getProperty("clickarrow"))).click();
		 * WebDriverWait waitdebugpriority = new WebDriverWait(driver, 360, 0000);
		 * waitdebugpriority.until(ExpectedConditions.visibilityOfElementLocated(By.
		 * xpath(prop.getProperty("debugpriority"))));
		 * driver.findElement(By.xpath(prop.getProperty("debugpriority"))).click();
		 * //Apply driver.findElement(By.xpath(prop.getProperty("Apply"))).click();
		 * Thread.sleep(4000); List<WebElement> responsefiles
		 * =driver.findElements(By.xpath(prop.getProperty("listdebug"))); WebElement
		 * target = responsefiles.get(0); Thread.sleep(2000); act.moveToElement(target);
		 * Thread.sleep(2000); WebElement clickresponselink =
		 * responsefiles.get(responsefiles.size()-1);
		 * 
		 * 
		 * Thread.sleep(4000); act.moveToElement(clickresponselink); act.clickAndHold();
		 * act.release().perform(); Thread.sleep(4000);
		 */
		// *****************************************************************
		// Keys.RETURN
		// driver.findElement(By.xpath(prop.getProperty("searchlogs"))).sendKeys(masterID);
		// driver.findElement(By.xpath(prop.getProperty("searchlogs"))).sendKeys(Keys.RETURN);

		Thread.sleep(10000);
		// WebElement field=driver.findElement(By.xpath(prop.getProperty("getlogs")));
		// act.moveToElement(field).doubleClick().build().perform();

		act.keyDown(Keys.CONTROL).sendKeys("a").keyUp(Keys.CONTROL).build().perform();
		act.keyDown(Keys.CONTROL).sendKeys("c").keyUp(Keys.CONTROL).build().perform();
		String logs = (String) Toolkit.getDefaultToolkit().getSystemClipboard().getData(DataFlavor.stringFlavor);
		int index1 = logs.indexOf("RandomNumber_Internal_SPLInterestRate");
		String roar1 = logs.substring(index1 + 39, index1 + 42);
		String randomnumber = null;
		if (roar1.contains(",")) {
			randomnumber = roar1.replace(",", "");
		} else {
			randomnumber = roar1;
		}

		double RandomNum = Double.valueOf(randomnumber);
		RandomNumberResponse = (int) RandomNum;
		 int Cc = logs.indexOf("DE_SPL_Credit_Contributor");
	        contributer  = logs.substring(Cc+29, Cc+40);
	       
		int index2 = logs.indexOf("DE_SPL_Buydown");
		SPLBuydown = logs.substring(index2 + 18, index2 + 19);
		
		driver.get(url);
		// driver.close();
	}

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
		label.setFont(new Font("Arial", Font.BOLD, 20));
		UIManager.put("OptionPane.minimumSize", new Dimension(500, 200));
		UIManager.put("TextField.font", new FontUIResource(new Font("Verdana", Font.BOLD, 18)));
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
		Thread.sleep(9000);
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

	public void getAddress() throws InterruptedException, java.text.ParseException {

		switchtoIframe1();
		Thread.sleep(5000);

		switchtoIframe2();
		Thread.sleep(5000);
		// threedots
		WebDriverWait waitSetup2 = new WebDriverWait(driver, 360, 0000);
		waitSetup2.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("OfferSelection"))));
		driver.findElement(By.xpath(prop.getProperty("threedots"))).click();

		// EventHistory
		WebDriverWait waitEventHistory = new WebDriverWait(driver, 360, 0000);
		waitEventHistory
				.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("EventHistory"))));
		driver.findElement(By.xpath(prop.getProperty("EventHistory"))).click();
		// EventHistoryTable
		driver.switchTo().defaultContent();
		switchtoIframe1();
		Thread.sleep(5000);
		switchtoIframe3();
		Thread.sleep(8000);
		String Event = null;
		try
		{
		WebDriverWait waitEventHistoryTable = new WebDriverWait(driver, 360, 0000);
		waitEventHistoryTable
				.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("EventHistoryTable"))));

		WebElement EventHistoryTable = driver.findElement(By.xpath(prop.getProperty("EventHistoryTable")));
		List<WebElement> rowValsEventHistoryTable = EventHistoryTable.findElements(By.tagName("tr"));
		int rowNumEventHistoryTable = EventHistoryTable.findElements(By.tagName("tr")).size();
		String EventHistoryDate = null;
		
		for (int i = 0; i < rowNumEventHistoryTable; i++) {

			// Get each row's column values by tag name
			List<WebElement> colValsEventHistory = rowValsEventHistoryTable.get(i).findElements(By.tagName("td"));
			WebElement EventHistory = colValsEventHistory.get(0);
			EventHistoryDate = EventHistory.getText();
			if (EventHistoryDate.contains("GDS returns offer details")) {
				WebElement getEvent = colValsEventHistory.get(4);
				Event = getEvent.getText();
			}

		}
		}
		catch (IndexOutOfBoundsException e)
		{
			System.out.println("Event History table not correctly displayed");
		}
		
		System.out.println("EventHistory: " + Event);

		// Format of the date defined in the input String
		DateFormat df = new SimpleDateFormat("dd/MM/yyyy hh:mm aa");
		// Desired format: 24 hour format: Change the pattern as per the need
		DateFormat outputformat = new SimpleDateFormat("MM/dd/yyyy HH:mm");
		java.util.Date date2 = null;
		String output = null;
		// Converting the input String to Date
		date2 = df.parse(Event);
		// Changing the format of date and storing it in String
		output = outputformat.format(date2);
		// Displaying the date

		String[] parts = output.split(" ");
		tsdate = parts[0];
		String tstime = parts[1];

		String logs = tsdate + " " + tstime + ":00";

		// parse the string
		org.joda.time.format.DateTimeFormatter dtf = DateTimeFormat.forPattern("MM/dd/yyyy HH:mm:ss");
		// Parsing the date
		DateTime jodatime = dtf.parseDateTime(logs);
		String ant0=String.valueOf(jodatime);
		ts0 = Character.toString(ant0.charAt(11)) + Character.toString(ant0.charAt(12)) + ":"
			+ Character.toString(ant0.charAt(14)) + Character.toString(ant0.charAt(15));
		// add two hours
		DateTime date = jodatime.minusMinutes(1);
		DateTime dateTime = jodatime.plusMinutes(1); // easier than mucking about with Calendar and constants

		String ant1 = String.valueOf(date);
		ts1 = Character.toString(ant1.charAt(11)) + Character.toString(ant1.charAt(12)) + ":"
				+ Character.toString(ant1.charAt(14)) + Character.toString(ant1.charAt(15));

		String ant = String.valueOf(dateTime);
		ts2 = Character.toString(ant.charAt(11)) + Character.toString(ant.charAt(12)) + ":"
				+ Character.toString(ant.charAt(14)) + Character.toString(ant.charAt(15));

		driver.switchTo().defaultContent();
		switchtoIframe1();
		driver.findElement(By.xpath(prop.getProperty("closeIncome"))).click();

		driver.switchTo().defaultContent();
		switchtoIframe1();
		Thread.sleep(3000);
		switchtoIframe2();
		Thread.sleep(3000);
		WebDriverWait waitSetup = new WebDriverWait(driver, 360, 0000);
		waitSetup.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("Parties"))));
		driver.findElement(By.xpath(prop.getProperty("Parties"))).click();
		Thread.sleep(5000);
		switchtoIframe4();
		Thread.sleep(5000);
		// CurrentAddress
		WebDriverWait CurrentAddresswait = new WebDriverWait(driver, 360, 0000);
		CurrentAddresswait
				.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("CurrentAddress"))));
		CurrentAddress = driver.findElement(By.xpath(prop.getProperty("CurrentAddress"))).getText();
		String reverse = new StringBuffer(CurrentAddress).reverse().toString();
		reverse = reverse.replaceAll(" ", "");
		String part1 = reverse.substring(0, 6);
		postalcode = new StringBuffer(part1).reverse().toString();
		if (postalcode.contains(" ")) {
			postalcode = postalcode.replaceAll(" ", "");
		}

		if (CurrentAddress.contains("Ontario")) {
			Prov = "ON";
		} else if (CurrentAddress.contains("Alberta")) {
			Prov = "AB";
		} else if (CurrentAddress.contains("British Columbia")) {
			Prov = "BC";
		} else if (CurrentAddress.contains("Manitoba")) {
			Prov = "MB";
		} else if (CurrentAddress.contains("New Brunswick")) {
			Prov = "NB";
		} else if (CurrentAddress.contains("Newfoundland and Labrador")) {
			Prov = "NL";
		} else if (CurrentAddress.contains("Nova Scotia")) {
			Prov = "NS";
		} else if (CurrentAddress.contains("Northwest Territories")) {
			Prov = "NT";
		} else if (CurrentAddress.contains("Nunavut")) {
			Prov = "NU";
		} else if (CurrentAddress.contains("Prince Edward")) {
			Prov = "PE";
		} else if (CurrentAddress.contains("Quebec")) {
			Prov = "QC";
		} else if (CurrentAddress.contains("Saskatchewan")) {
			Prov = "SK";
		} else if (CurrentAddress.contains("Yukon")) {
			Prov = "YT";
		}

	}

	public void getPartyDetailsCoApp()
			throws InterruptedException, HeadlessException, UnsupportedFlavorException, IOException {
		// PartyDetails
		double AppIncome, CoAppIncome;
		WebDriverWait waitPartyDetails = new WebDriverWait(driver, 360, 0000);
		waitPartyDetails
				.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("PartyDetails"))));
		driver.findElement(By.xpath(prop.getProperty("PartyDetails"))).click();

		// EmploymentandIncome

		WebDriverWait waitEmploymentandIncome = new WebDriverWait(driver, 360, 0000);
		waitEmploymentandIncome.until(
				ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("EmploymentandIncome"))));
		driver.findElement(By.xpath(prop.getProperty("EmploymentandIncome"))).click();
		driver.switchTo().defaultContent();
		switchtoIframe1();
		switchtoIframe3();
		// Income table
		Thread.sleep(5000);
		WebDriverWait waitMonthlyIncometable = new WebDriverWait(driver, 360, 0000);
		waitMonthlyIncometable
				.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("MonthlyIncometable"))));

		WebElement MonthlyIncometable = driver.findElement(By.xpath(prop.getProperty("MonthlyIncometable")));
		List<WebElement> rowValsIncome = MonthlyIncometable.findElements(By.tagName("tr"));
		int rowNumIncome = MonthlyIncometable.findElements(By.tagName("tr")).size();

		String str = null;
		double ApplicantIncome = 0, OtherIncomeValue=0.0;
		for (int i = 0; i < rowNumIncome; i++) {

			double subValue = 0;
			// Get each row's column values by tag name
			List<WebElement> colValsIncome = rowValsIncome.get(i).findElements(By.tagName("td"));
			WebElement Income = colValsIncome.get(4);
			Actions act = new Actions(driver);
			Income.click();
			Thread.sleep(1000);
			Income.click();
			
			act.keyDown(Keys.CONTROL).sendKeys("a").keyUp(Keys.CONTROL).build().perform();
			act.keyDown(Keys.CONTROL).sendKeys("c").keyUp(Keys.CONTROL).build().perform();
			String Incomeamount = (String) Toolkit.getDefaultToolkit().getSystemClipboard()
					.getData(DataFlavor.stringFlavor);

			if (Incomeamount.contains(","))

			{
				str = Incomeamount.replace(",", "");
				subValue = Double.parseDouble(str.replace("$", ""));
			} else {
				subValue = Double.parseDouble(Incomeamount.replace("$", ""));
			}
			ApplicantIncome += subValue;
		}
		
		// otherIncometable
		WebElement otherIncometable = driver.findElement(By.xpath(prop.getProperty("otherIncometable")));
		List<WebElement> rowValsOther = otherIncometable.findElements(By.tagName("tr"));
		int rowNumOther = otherIncometable.findElements(By.tagName("tr")).size();
		String strOther = null;
		for (int i = 0; i < rowNumOther; i++) {

			double subValue = 0;
			// Get each row's column values by tag name
			List<WebElement> colValOther = rowValsOther.get(i).findElements(By.tagName("td"));
			WebElement Income = colValOther.get(4);
			Actions act = new Actions(driver);
			Income.click();
			Thread.sleep(1000);
			Income.click();
			
			act.keyDown(Keys.CONTROL).sendKeys("a").keyUp(Keys.CONTROL).build().perform();
			act.keyDown(Keys.CONTROL).sendKeys("c").keyUp(Keys.CONTROL).build().perform();
			String Incomeamount = (String) Toolkit.getDefaultToolkit().getSystemClipboard()
					.getData(DataFlavor.stringFlavor);

			if (Incomeamount.contains(","))

			{
				strOther = Incomeamount.replace(",", "");
				subValue = Double.parseDouble(strOther.replace("$", ""));
			} else {
				subValue = Double.parseDouble(Incomeamount.replace("$", ""));
			}
			OtherIncomeValue += subValue;

		}

		AppIncome = ApplicantIncome + OtherIncomeValue;
		
		// close
		driver.switchTo().defaultContent();
		switchtoIframe1();
		driver.findElement(By.xpath(prop.getProperty("closeIncome"))).click();
		switchtoIframe2();
		switchtoIframe4();
		Thread.sleep(2000);
		driver.findElement(By.xpath(prop.getProperty("PartyDetails"))).click();
		// Liabilities
		WebDriverWait waitLiabilities = new WebDriverWait(driver, 360, 0000);
		waitLiabilities.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("Liabilities"))));
		driver.findElement(By.xpath(prop.getProperty("Liabilities"))).click();
		driver.switchTo().defaultContent();
		switchtoIframe1();
		switchtoIframe3();
		Thread.sleep(8000);
		try
		{
		
		WebDriverWait waitLiabilitiesTable = new WebDriverWait(driver, 6);
		waitLiabilitiesTable
				.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("LiabilitiesTable"))));
		WebElement LiabilitiesTable = driver.findElement(By.xpath(prop.getProperty("LiabilitiesTable")));
		if (LiabilitiesTable.isDisplayed()) {
			// showrows
			String rows = driver.findElement(By.xpath(prop.getProperty("showrows"))).getText();
			int rowcount=0;

			if (rows.contains("+")) {
				do {
					Loadmore();
					Thread.sleep(5000);
					rows = driver.findElement(By.xpath(prop.getProperty("showrows"))).getText();
					rowcount++;
				} while (rows.contains("+"));
			}

			List<WebElement> rowValsLiability = LiabilitiesTable.findElements(By.tagName("tr"));
			int rowNumLiability = LiabilitiesTable.findElements(By.tagName("tr")).size();

			String strLiability = null;
			TotalLaibility=0;
			for (int i = 0; i < rowNumLiability; i++) {

				double subValue = 0;
				// Get each row's column values by tag name
				List<WebElement> colValsliability = rowValsLiability.get(i).findElements(By.tagName("td"));
				WebElement Liability = colValsliability.get(9);
				String Liabilityamount = Liability.getText();

				if (Liabilityamount.contains(","))

				{
					strLiability = Liabilityamount.replace(",", "");
					subValue = Double.parseDouble(strLiability.replace("$", ""));
				} else {
					subValue = Double.parseDouble(Liabilityamount.replace("$", ""));
				}
				TotalLaibility += subValue;
			}
		
			if(rowcount>4)
			{
				Thread.sleep(4000);
				//click2
				driver.findElement(By.xpath(prop.getProperty("next"))).click();
				Thread.sleep(3000);
				List<WebElement> rowValsLiability1 = LiabilitiesTable.findElements(By.tagName("tr"));
				int rowNumLiability1 = LiabilitiesTable.findElements(By.tagName("tr")).size();

				String strLiability1 = null;
				
				for (int i = 0; i < rowNumLiability1; i++) {

					double subValue = 0;
					// Get each row's column values by tag name
					List<WebElement> colValsliability = rowValsLiability1.get(i).findElements(By.tagName("td"));
					WebElement Liability = colValsliability.get(9);
					String Liabilityamount = Liability.getText();

					if (Liabilityamount.contains(","))

					{
						strLiability1 = Liabilityamount.replace(",", "");
						subValue = Double.parseDouble(strLiability1.replace("$", ""));
					} else {
						subValue = Double.parseDouble(Liabilityamount.replace("$", ""));
					}
					TotalLaibility += subValue;
				}
				if(rowcount>9 )
				{
					Thread.sleep(4000);
					//click3
					driver.findElement(By.xpath(prop.getProperty("next"))).click();
					Thread.sleep(3000);
					List<WebElement> rowValsLiability11 = LiabilitiesTable.findElements(By.tagName("tr"));
					int rowNumLiability11 = LiabilitiesTable.findElements(By.tagName("tr")).size();

					String strLiability11 = null;
					
					for (int i = 0; i < rowNumLiability11; i++) {

						double subValue = 0;
						// Get each row's column values by tag name
						List<WebElement> colValsliability = rowValsLiability11.get(i).findElements(By.tagName("td"));
						WebElement Liability = colValsliability.get(9);
						String Liabilityamount = Liability.getText();

						if (Liabilityamount.contains(","))

						{
							strLiability11 = Liabilityamount.replace(",", "");
							subValue = Double.parseDouble(strLiability11.replace("$", ""));
						} else {
							subValue = Double.parseDouble(Liabilityamount.replace("$", ""));
						}
						TotalLaibility += subValue;
					}
					
					if(rowcount>14 )
					{
						Thread.sleep(4000);
						//click4
						driver.findElement(By.xpath(prop.getProperty("next"))).click();
						Thread.sleep(3000);
						List<WebElement> rowValsLiability111 = LiabilitiesTable.findElements(By.tagName("tr"));
						int rowNumLiability111 = LiabilitiesTable.findElements(By.tagName("tr")).size();

						String strLiability111 = null;
						
						for (int i = 0; i < rowNumLiability111; i++) {

							double subValue = 0;
							// Get each row's column values by tag name
							List<WebElement> colValsliability = rowValsLiability111.get(i).findElements(By.tagName("td"));
							WebElement Liability = colValsliability.get(9);
							String Liabilityamount = Liability.getText();

							if (Liabilityamount.contains(","))

							{
								strLiability111 = Liabilityamount.replace(",", "");
								subValue = Double.parseDouble(strLiability111.replace("$", ""));
							} else {
								subValue = Double.parseDouble(Liabilityamount.replace("$", ""));
							}
							TotalLaibility += subValue;
						}
					}
					
				}
				
			}
			
			
		}
		
		}
		
		  catch (TimeoutException e)
        {
            System.out.println("Liability Not Displayed");
           

            
        }
		// close
					driver.switchTo().defaultContent();
					switchtoIframe1();
					driver.findElement(By.xpath(prop.getProperty("closeIncome"))).click();
	
		//Co-Borrower
driver.switchTo().defaultContent();
switchtoIframe1();
Thread.sleep(3000);
switchtoIframe2();
Thread.sleep(3000);
WebDriverWait ClickCoAppwait = new WebDriverWait(driver, 360, 0000);
ClickCoAppwait
		.until(ExpectedConditions.elementToBeClickable(By.xpath(prop.getProperty("ClickCoApp"))));
driver.findElement(By.xpath(prop.getProperty("ClickCoApp"))).click();
Thread.sleep(5000);
switchtoIframe4();
Thread.sleep(5000);
					
					// CurrentAddress
					WebDriverWait CurrentAddresswait = new WebDriverWait(driver, 360, 0000);
					CurrentAddresswait
							.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("CurrentAddress"))));
					CurrentAddress = driver.findElement(By.xpath(prop.getProperty("CurrentAddress"))).getText();
					String reverse = new StringBuffer(CurrentAddress).reverse().toString();
					reverse = reverse.replaceAll(" ", "");
					String part1 = reverse.substring(0, 6);
					postalcode = new StringBuffer(part1).reverse().toString();
					if (postalcode.contains(" ")) {
						postalcode = postalcode.replaceAll(" ", "");
					}

					if (CurrentAddress.contains("Ontario")) {
						CoAppProvince = "ON";
					} else if (CurrentAddress.contains("Alberta")) {
						CoAppProvince = "AB";
					} else if (CurrentAddress.contains("British Columbia")) {
						CoAppProvince = "BC";
					} else if (CurrentAddress.contains("Manitoba")) {
						CoAppProvince = "MB";
					} else if (CurrentAddress.contains("New Brunswick")) {
						CoAppProvince = "NB";
					} else if (CurrentAddress.contains("Newfoundland and Labrador")) {
						CoAppProvince = "NL";
					} else if (CurrentAddress.contains("Nova Scotia")) {
						CoAppProvince = "NS";
					} else if (CurrentAddress.contains("Northwest Territories")) {
						CoAppProvince = "NT";
					} else if (CurrentAddress.contains("Nunavut")) {
						CoAppProvince = "NU";
					} else if (CurrentAddress.contains("Prince Edward")) {
						CoAppProvince = "PE";
					} else if (CurrentAddress.contains("Quebec")) {
						CoAppProvince = "QC";
					} else if (CurrentAddress.contains("Saskatchewan")) {
						CoAppProvince = "SK";
					} else if (CurrentAddress.contains("Yukon")) {
						CoAppProvince = "YT";
					}
					System.out.println("CoAppProvince ="+CoAppProvince);
					
					
					// PartyDetails
					waitPartyDetails
							.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("PartyDetails"))));
					driver.findElement(By.xpath(prop.getProperty("PartyDetails"))).click();

					// EmploymentandIncome
					waitEmploymentandIncome.until(
							ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("EmploymentandIncome"))));
					driver.findElement(By.xpath(prop.getProperty("EmploymentandIncome"))).click();
					driver.switchTo().defaultContent();
					switchtoIframe1();
					switchtoIframe3();
					// Income table
					Thread.sleep(5000);
					WebDriverWait waitMonthly = new WebDriverWait(driver, 360, 0000);
					waitMonthly
							.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("MonthlyIncometable"))));

					WebElement Incometable = driver.findElement(By.xpath(prop.getProperty("MonthlyIncometable")));
					waitMonthlyIncometable
							.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("MonthlyIncometable"))));

					List<WebElement> rowValsIncomeCoApp = Incometable.findElements(By.tagName("tr"));
					int rowNumIncomeCoApp = Incometable.findElements(By.tagName("tr")).size();

					String strCoApp = null;
					double CoApplicantIncome = 0, CoOtherIncomeValue=0.0;
					for (int i = 0; i < rowNumIncomeCoApp; i++) {

						double subValue = 0;
						// Get each row's column values by tag name
						List<WebElement> colValsIncomeCoApp = rowValsIncomeCoApp.get(i).findElements(By.tagName("td"));
						WebElement Income = colValsIncomeCoApp.get(4);
						Actions act = new Actions(driver);
						Income.click();
						Thread.sleep(1000);
						Income.click();
						
						act.keyDown(Keys.CONTROL).sendKeys("a").keyUp(Keys.CONTROL).build().perform();
						act.keyDown(Keys.CONTROL).sendKeys("c").keyUp(Keys.CONTROL).build().perform();
						String IncomeamountCoApp = (String) Toolkit.getDefaultToolkit().getSystemClipboard()
								.getData(DataFlavor.stringFlavor);

						if (IncomeamountCoApp.contains(","))

						{
							strCoApp = IncomeamountCoApp.replace(",", "");
							subValue = Double.parseDouble(strCoApp.replace("$", ""));
						} else {
							subValue = Double.parseDouble(IncomeamountCoApp.replace("$", ""));
						}
						CoApplicantIncome += subValue;
					}
					
					// otherIncometable
					WebElement othertable = driver.findElement(By.xpath(prop.getProperty("otherIncometable")));
					List<WebElement> rowValsOtherCoApp = othertable.findElements(By.tagName("tr"));
					int rowNumOtherCoApp = othertable.findElements(By.tagName("tr")).size();
					String strOtherCoApp = null;
					for (int i = 0; i < rowNumOtherCoApp; i++) {

						double subValue = 0;
						// Get each row's column values by tag name
						List<WebElement> colValOtherCoApp = rowValsOtherCoApp.get(i).findElements(By.tagName("td"));
						WebElement Income = colValOtherCoApp.get(4);
						Actions act = new Actions(driver);
						Income.click();
						Thread.sleep(1000);
						Income.click();
						
						act.keyDown(Keys.CONTROL).sendKeys("a").keyUp(Keys.CONTROL).build().perform();
						act.keyDown(Keys.CONTROL).sendKeys("c").keyUp(Keys.CONTROL).build().perform();
						String Incomeamount = (String) Toolkit.getDefaultToolkit().getSystemClipboard()
								.getData(DataFlavor.stringFlavor);

						if (Incomeamount.contains(","))

						{
							strOtherCoApp = Incomeamount.replace(",", "");
							subValue = Double.parseDouble(strOtherCoApp.replace("$", ""));
						} else {
							subValue = Double.parseDouble(Incomeamount.replace("$", ""));
						}
						CoOtherIncomeValue += subValue;

					}

					CoAppIncome = CoApplicantIncome + CoOtherIncomeValue;
					IncomeValue=AppIncome+CoAppIncome;
					System.out.println("TotalIncome =$" + IncomeValue);			
					
					// close
					driver.switchTo().defaultContent();
					switchtoIframe1();
					driver.findElement(By.xpath(prop.getProperty("closeIncome"))).click();
					switchtoIframe2();
					switchtoIframe4();
					Thread.sleep(2000);
					driver.findElement(By.xpath(prop.getProperty("PartyDetails"))).click();
					// Liabilities
					WebDriverWait waitCoAppLiabilities = new WebDriverWait(driver, 360, 0000);
					waitCoAppLiabilities.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("Liabilities"))));
					driver.findElement(By.xpath(prop.getProperty("Liabilities"))).click();
					driver.switchTo().defaultContent();
					switchtoIframe1();
					switchtoIframe3();
					Thread.sleep(8000);
					
					try {
					CoAppLiabilities();
					}
					catch (TimeoutException e)
		            {
		                System.out.println("Liability Not Displayed");
		               
		                
		            }
					
		System.out.println("Total Liability =$" + TotalLaibility);
		// close
		driver.switchTo().defaultContent();
		switchtoIframe1();
		driver.findElement(By.xpath(prop.getProperty("closeIncome"))).click();
	}
	
	public void CoAppLiabilities() throws InterruptedException
	{
		WebDriverWait waitLiabilitiesTable = new WebDriverWait(driver, 6);
		waitLiabilitiesTable
				.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("LiabilitiesTable"))));
		WebElement LiabilitiesTable = driver.findElement(By.xpath(prop.getProperty("LiabilitiesTable")));
		if (LiabilitiesTable.isDisplayed()) {
			// showrows
			String CoApprows = driver.findElement(By.xpath(prop.getProperty("showrows"))).getText();
			int CoApprowcount=0;

			if (CoApprows.contains("+")) {
				do {
					Loadmore();
					Thread.sleep(5000);
					CoApprows = driver.findElement(By.xpath(prop.getProperty("showrows"))).getText();
					CoApprowcount++;
				} while (CoApprows.contains("+"));
			}

			List<WebElement> rowValsLiability = LiabilitiesTable.findElements(By.tagName("tr"));
			int rowNumLiability = LiabilitiesTable.findElements(By.tagName("tr")).size();

			String strLiability = null;
			for (int i = 0; i < rowNumLiability; i++) {

				double subValue = 0;
				// Get each row's column values by tag name
				List<WebElement> colValsliability = rowValsLiability.get(i).findElements(By.tagName("td"));
				WebElement Liability = colValsliability.get(9);
				String Liabilityamount = Liability.getText();

				if (Liabilityamount.contains(","))

				{
					strLiability = Liabilityamount.replace(",", "");
					subValue = Double.parseDouble(strLiability.replace("$", ""));
				} else {
					subValue = Double.parseDouble(Liabilityamount.replace("$", ""));
				}
				TotalLaibility += subValue;
			}
			if(CoApprowcount>4)
			{
				Thread.sleep(5000);
				//click2
				driver.findElement(By.xpath(prop.getProperty("next"))).click();
				Thread.sleep(3000);
				List<WebElement> rowValsLiability1 = LiabilitiesTable.findElements(By.tagName("tr"));
				int rowNumLiability1 = LiabilitiesTable.findElements(By.tagName("tr")).size();

				String strLiability1 = null;
				
				for (int i = 0; i < rowNumLiability1; i++) {

					double subValue = 0;
					// Get each row's column values by tag name
					List<WebElement> colValsliability = rowValsLiability1.get(i).findElements(By.tagName("td"));
					WebElement Liability = colValsliability.get(9);
					String Liabilityamount = Liability.getText();

					if (Liabilityamount.contains(","))

					{
						strLiability1 = Liabilityamount.replace(",", "");
						subValue = Double.parseDouble(strLiability1.replace("$", ""));
					} else {
						subValue = Double.parseDouble(Liabilityamount.replace("$", ""));
					}
					TotalLaibility += subValue;
				}
				if(CoApprowcount>9 )
				{
					Thread.sleep(4000);
					//click3
					driver.findElement(By.xpath(prop.getProperty("next"))).click();
					Thread.sleep(3000);
					List<WebElement> rowValsLiability11 = LiabilitiesTable.findElements(By.tagName("tr"));
					int rowNumLiability11 = LiabilitiesTable.findElements(By.tagName("tr")).size();

					String strLiability11 = null;
					
					for (int i = 0; i < rowNumLiability11; i++) {

						double subValue = 0;
						// Get each row's column values by tag name
						List<WebElement> colValsliability = rowValsLiability11.get(i).findElements(By.tagName("td"));
						WebElement Liability = colValsliability.get(9);
						String Liabilityamount = Liability.getText();

						if (Liabilityamount.contains(","))

						{
							strLiability11 = Liabilityamount.replace(",", "");
							subValue = Double.parseDouble(strLiability11.replace("$", ""));
						} else {
							subValue = Double.parseDouble(Liabilityamount.replace("$", ""));
						}
						TotalLaibility += subValue;
					}
					
					if(CoApprowcount>14 )
					{
						Thread.sleep(4000);
						//click4
						driver.findElement(By.xpath(prop.getProperty("next"))).click();
						Thread.sleep(3000);
						List<WebElement> rowValsLiability111 = LiabilitiesTable.findElements(By.tagName("tr"));
						int rowNumLiability111 = LiabilitiesTable.findElements(By.tagName("tr")).size();

						String strLiability111 = null;
						
						for (int i = 0; i < rowNumLiability111; i++) {

							double subValue = 0;
							// Get each row's column values by tag name
							List<WebElement> colValsliability = rowValsLiability111.get(i).findElements(By.tagName("td"));
							WebElement Liability = colValsliability.get(9);
							String Liabilityamount = Liability.getText();

							if (Liabilityamount.contains(","))

							{
								strLiability111 = Liabilityamount.replace(",", "");
								subValue = Double.parseDouble(strLiability111.replace("$", ""));
							} else {
								subValue = Double.parseDouble(Liabilityamount.replace("$", ""));
							}
							TotalLaibility += subValue;
						}
					}
					
				}
				
			}
			
			
		}
	}
	
	

	public void Loadmore() {

		// loadmore
		Actions action = new Actions(driver);

		if (driver.findElement(By.xpath(prop.getProperty("loadmore"))).isDisplayed()) {
			WebElement loadmore = driver.findElement(By.xpath(prop.getProperty("loadmore")));
			((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", loadmore);
			loadmore.click();
		}
	}

	public void getCollateral() {
		// Collateral
		switchtoIframe2();
		WebDriverWait waitCollateral = new WebDriverWait(driver, 360, 0000);
		waitCollateral.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("Collateral"))));
		driver.findElement(By.xpath(prop.getProperty("Collateral"))).click();
		WebDriverWait waitpropertyType = new WebDriverWait(driver, 360, 0000);
		waitpropertyType
				.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("PropertyType"))));
		propertyType = driver.findElement(By.xpath(prop.getProperty("PropertyType"))).getText();
		System.out.println(propertyType);
		MortgageBalances = driver.findElement(By.xpath(prop.getProperty("TotalMortgageBalanceOutstanding"))).getText();
		PropertyValue = driver.findElement(By.xpath(prop.getProperty("EstimatedPropertyValue"))).getText();
	}

	public void getAppDetailsCoApp() throws InterruptedException, IOException {
		// OfferSelection
		WebDriverWait waitAgentDashboard = new WebDriverWait(driver, 360, 0000);
		waitAgentDashboard
				.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("AgentDashboard"))));
		driver.findElement(By.xpath(prop.getProperty("AgentDashboard"))).click();
		WebDriverWait waitSetup2 = new WebDriverWait(driver, 360, 0000);
		waitSetup2.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("OfferSelection"))));
		url = driver.getCurrentUrl();
		// Get MasterID
		String reverseurl = new StringBuffer(url).reverse().toString();
		String[] Spliturl = reverseurl.split("/");
		String view = Spliturl[1];
		masterID = new StringBuffer(view).reverse().toString();
		System.out.println("MasterID =" + masterID);
		driver.findElement(By.xpath(prop.getProperty("OfferSelection"))).click();

		// Switch Iframes
		driver.switchTo().defaultContent();
		switchtoIframe1();
		switchtoIframe3();

		// Offer details

		Thread.sleep(12000);
		WebElement getQla = driver.findElement(By.xpath(prop.getProperty("QualifiedLoanAmount")));
		Actions action = new Actions(driver);

		QualifiedLoanAmount = getQla.getText();
		String qla;
		if (QualifiedLoanAmount.contains(",")) {
			qla = QualifiedLoanAmount.replace(",", "");
		} else {
			qla = QualifiedLoanAmount;
		}

		ExpectedQLA = Double.parseDouble(qla.replace("$", ""));
		System.out.println("ExpectedQLA =$" + ExpectedQLA);
		WebElement getInt = driver.findElement(By.xpath(prop.getProperty("InterestRate")));
		IntRate = getInt.getText();
		((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", getQla);
		Thread.sleep(2000);
		QLAInterestScreen = Screenshot.capture(driver, "CaculateQLA");
		System.out.println("InterestRate =" + IntRate);
		ExpInt = Double.parseDouble(IntRate.replace("%", ""));
		WebDriverWait waitOfferdetails = new WebDriverWait(driver, 360, 0000);
		waitOfferdetails
				.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("Offerdetails"))));

		// Thread.sleep(8000);
		driver.findElement(By.xpath(prop.getProperty("Offerdetails"))).click();

		// Total Income
		WebDriverWait totalincome = new WebDriverWait(driver, 360, 0000);
		totalincome
				.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("TotalIncomeFusion"))));
		totalincomeFusion = driver.findElement(By.xpath(prop.getProperty("TotalIncomeFusion"))).getText();
		String st = totalincomeFusion.replace(",", "");
		TotalIncome = Double.parseDouble(st.replace("$", ""));
		System.out.println("Total Income =$" + TotalIncome);
		IncomeLiabilityScreen = Screenshot.capture(driver, "CalculateIncome");
		// Total Debt
		totaldebtFusion = driver.findElement(By.xpath(prop.getProperty("TotalDebtFusion"))).getText();
		String str1 = totaldebtFusion.replace(",", "");
		TotalDebt = Double.parseDouble(str1.replace("$", ""));
		MaximumLTV = driver.findElement(By.xpath(prop.getProperty("MaximumLTV"))).getText();

		LtvMax = Double.parseDouble(MaximumLTV);

		System.out.println("Total Debt =$" + TotalDebt);
		System.out.println("MaximumLTV =" + LtvMax);

		// ExpectedMaxHA
		List<WebElement> HA = driver.findElements(By.xpath(prop.getProperty("HAFullSpl")));
		WebElement targetHA = HA.get(1);
		((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", targetHA);
		Thread.sleep(2000);
		MaxHASPLFullScreen = Screenshot.capture(driver, "MAxHA");
		String MaxHA = targetHA.getText();
		ExpectedMaxHA = Double.parseDouble(MaxHA.replace(",", ""));
		System.out.println("ExpectedMaxHA =" + ExpectedMaxHA);

		// ReasonCodeSPLFullScreen
		List<WebElement> reasoncode = driver.findElements(By.xpath(prop.getProperty("SPLRiskFactors")));
		length = reasoncode.size();
		if (length == 1) {
			WebElement desicioncode1 = reasoncode.get(0);
			((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", desicioncode1);
			Thread.sleep(2000);
			ReasonCodeSPLFullScreen1 = Screenshot.capture(driver, "ReasonCode1");
		} else if (length == 2) {
			WebElement desicioncode1 = reasoncode.get(0);
			((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", desicioncode1);
			Thread.sleep(2000);
			ReasonCodeSPLFullScreen1 = Screenshot.capture(driver, "ReasonCode1");
			WebElement desicioncode2 = reasoncode.get(1);
			((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", desicioncode2);
			Thread.sleep(2000);
			ReasonCodeSPLFullScreen2 = Screenshot.capture(driver, "ReasonCode2");
		}
		else if (length == 3) {
			WebElement desicioncode1 = reasoncode.get(0);
			((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", desicioncode1);
			Thread.sleep(2000);
			ReasonCodeSPLFullScreen1 = Screenshot.capture(driver, "ReasonCode1");
			WebElement desicioncode2 = reasoncode.get(1);
			((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", desicioncode2);
			Thread.sleep(2000);
			ReasonCodeSPLFullScreen2 = Screenshot.capture(driver, "ReasonCode2");
			WebElement desicioncode3 = reasoncode.get(2);
			((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", desicioncode3);
			Thread.sleep(2000);
			ReasonCodeSPLFullScreen3 = Screenshot.capture(driver, "ReasonCode3");
		}
		
		else {
			WebElement desicioncode1 = reasoncode.get(0);
			((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", desicioncode1);
			Thread.sleep(2000);
			ReasonCodeSPLFullScreen1 = Screenshot.capture(driver, "ReasonCode1");
			WebElement desicioncode2 = reasoncode.get(1);
			((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", desicioncode2);
			Thread.sleep(2000);
			ReasonCodeSPLFullScreen2 = Screenshot.capture(driver, "ReasonCode2");
		}
		// ApplicantEFSCVScore

		ApplicantEFSCVScore = driver.findElement(By.xpath(prop.getProperty("ApplicantEFSCVScore"))).getText();
		String parts = ApplicantEFSCVScore.substring(0, 3);
		cvScore = Integer.parseInt(parts);
		System.out.println("ApplicantEFSCVScore =" + cvScore);
		//cvScoreCoapp
		String cvscoreco = driver.findElement(By.xpath(prop.getProperty("Co-ApplicantEFSCVScore"))).getText();
		String partsCoapp = cvscoreco.substring(0, 3);
		cvScoreCoapp = Integer.parseInt(partsCoapp);
		System.out.println("CoApplicantEFSCVScore =" + cvScoreCoapp);
		appbehscore = driver.findElement(By.xpath(prop.getProperty("behscore"))).getText();
		cobehscore = driver.findElement(By.xpath(prop.getProperty("cobehscore"))).getText();

		// CloseOffer
		driver.findElement(By.xpath(prop.getProperty("CloseOffer"))).click();

	}

	public void calculateIncome() throws InterruptedException, IOException {

		System.out.println("Resubmission attempt #" + attemptNo);
		if (attemptNo == 0) {
			test = Extent.createTest("Total Income Calcuation");
		}

		else {

			test = Extent.createTest("Resubmission Result: Attempt#  " + attemptNo + " - Total Income Calcuation");
		}

		System.out.println("Actual Income: $" + IncomeValue);
		System.out.println("Expected Income: $" + TotalIncome);
		if (IncomeValue == TotalIncome) {

			test.log(Status.PASS,
					MarkupHelper.createLabel(" Total Income  :Actual Value =  $" + IncomeValue, ExtentColor.GREEN));
			test.log(Status.PASS,
					MarkupHelper.createLabel("Total Income  :Expected Value =  $" + TotalIncome, ExtentColor.GREEN));
			test.log(Status.PASS, MarkupHelper.createLabel("Income is Matching with GDS Decision", ExtentColor.GREEN));
			test.log(Status.PASS, "Snapshot below: " + test.addScreenCaptureFromPath(IncomeLiabilityScreen));
			System.out.println("Income:Passed");

		} else {

			test.log(Status.FAIL,
					MarkupHelper.createLabel("Total Income : Actual Value =  $" + IncomeValue, ExtentColor.RED));
			test.log(Status.FAIL,
					MarkupHelper.createLabel("Total Income : Expected Value =  $" + TotalIncome, ExtentColor.RED));
			test.log(Status.FAIL, MarkupHelper.createLabel("Income Not Matching with GDS Decision", ExtentColor.RED));
			test.log(Status.FAIL, "Snapshot below: " + test.addScreenCaptureFromPath(IncomeLiabilityScreen));
			System.out.println("Income Not Match");
		}

	}

	public void calculateSPLLiability() throws InterruptedException, IOException {

		if (attemptNo == 0) {
			test = Extent.createTest("Total Liability Calcuation");
		}

		else {

			test = Extent.createTest("Resubmission Result: Attempt#  " + attemptNo + " - Total Liability Calcuation");
		}
		System.out.println("Actual Liabilities: $" + TotalLaibility);
		System.out.println("Expected Liabilities: $" + TotalDebt);
		if (TotalLaibility == TotalDebt) {
			System.out.println("Laibilities:Passed");

			test.log(Status.PASS, MarkupHelper.createLabel("Total Liability - Actual Value   =  $" + TotalLaibility,
					ExtentColor.GREEN));
			test.log(Status.PASS,
					MarkupHelper.createLabel("Total Liability - Expected Value =  $" + TotalDebt, ExtentColor.GREEN));

			test.log(Status.PASS,
					MarkupHelper.createLabel("Liability is Matching with GDS Decision ", ExtentColor.GREEN));
			test.log(Status.PASS, "Snapshot below: " + test.addScreenCaptureFromPath(IncomeLiabilityScreen));
			// Assert.assertTrue(true);
		} else {
			System.out.println("Laibilities:Failed");

			test.log(Status.FAIL, MarkupHelper.createLabel("Total Liability - Actual Value   =  $" + TotalLaibility,
					ExtentColor.RED));
			test.log(Status.FAIL,
					MarkupHelper.createLabel("Total Liability - Expected Value =  $" + TotalDebt, ExtentColor.RED));

			test.log(Status.FAIL,
					MarkupHelper.createLabel("Liability is not Matching with GDS Decision ", ExtentColor.RED));
			test.log(Status.FAIL, "Snapshot below: " + test.addScreenCaptureFromPath(IncomeLiabilityScreen));

			// Assert.assertTrue(false);
		}
	}

	public void splLTV() {

		// SPL LTV Calculation - 1st Submission
		// SPL LTV= (Total Amount of Applicant Mortgage Balances Outstanding + Total
		// Credit Limits of Revolving Trades of Applicant)*100/Total Value of Property

		String str = MortgageBalances.replace(",", "");
		double MortgageBal = Double.parseDouble(str.replace("$", ""));
		System.out.println("Mortgage Balance = " + MortgageBal);
		String str1 = PropertyValue.replace(",", "");
		PropertyVal = Double.parseDouble(str1.replace("$", ""));
		System.out.println("Property Value = " + PropertyVal);
		SPLltv = MortgageBal * 100 / PropertyVal;
		System.out.println("SPL LTV =" + SPLltv);

	}

	public void splIncreaseRemInc() throws IOException {
		File file = new File(System.getProperty("user.dir")
				+ "\\src\\main\\resources\\Excel\\Credit Decisions Calculations - SPL Increase.xlsx");

		FileInputStream inputStream = new FileInputStream(file);

		Workbook workbook = new XSSFWorkbook(inputStream);

		org.apache.poi.ss.usermodel.Sheet sheet = workbook.getSheet("SPL Remaining Income");

		// CV Score
		ArrayList<Integer> efsCvScoreList1 = new ArrayList<Integer>();
		ArrayList<Integer> efsCvScoreList2 = new ArrayList<Integer>();
		// Behavior Score
		ArrayList<Integer> efsCvScoreList3 = new ArrayList<Integer>();
		ArrayList<Integer> efsCvScoreList4 = new ArrayList<Integer>();

		for (int x = 2; x < 24; x++) {
			int efsRange1 = (int) sheet.getRow(x).getCell(1).getNumericCellValue();

			efsCvScoreList1.add(efsRange1);
		}

		for (int x = 2; x < 24; x++) {

			int efsRange2 = (int) sheet.getRow(x).getCell(3).getNumericCellValue();
			efsCvScoreList2.add(efsRange2);
		}

		// Behavior Score list
		for (int x = 2; x < 24; x++) {

			int efsRange3 = (int) sheet.getRow(x).getCell(4).getNumericCellValue();
			efsCvScoreList3.add(efsRange3);
		}

		for (int x = 2; x < 24; x++) {

			int efsRange4 = (int) sheet.getRow(x).getCell(6).getNumericCellValue();
			efsCvScoreList4.add(efsRange4);
		}

		// Remaining Income
		ArrayList<Double> efsCvScoreList5 = new ArrayList<Double>();

		for (int x = 2; x < 24; x++) {

			double efsRange5 = sheet.getRow(x).getCell(8).getNumericCellValue();
			efsCvScoreList5.add(efsRange5);
		}

		double val = 0;
		// 3

		if (efsCvScoreList2.get(0) <= cvScore && efsCvScoreList4.get(0) < BehaviourScore) {
			val = efsCvScoreList5.get(0);
			RemainingIncome = TotalIncome * val - TotalDebt;
		}
		// 4
		else if (efsCvScoreList2.get(1) <= cvScore && (efsCvScoreList4.get(1) == BehaviourScore)) {
			val = efsCvScoreList5.get(1);
			RemainingIncome = TotalIncome * val - TotalDebt;
		}
		// 5
		else if (efsCvScoreList2.get(2) <= cvScore && (efsCvScoreList3.get(2) <= BehaviourScore)
				&& (BehaviourScore <= efsCvScoreList4.get(2))) {
			val = efsCvScoreList5.get(2);
			RemainingIncome = TotalIncome * val - TotalDebt;
		}
		// 6
		else if (efsCvScoreList2.get(3) <= cvScore
				&& efsCvScoreList4.get(3) >= BehaviourScore) {
			val = efsCvScoreList5.get(3);
			RemainingIncome = TotalIncome * val - TotalDebt;
		}
		// 7
		else if (efsCvScoreList1.get(4) <= cvScore && cvScore <= efsCvScoreList2.get(4)
				&& (efsCvScoreList3.get(4) <= BehaviourScore)
				&& (BehaviourScore <= efsCvScoreList4.get(4))) {
			val = efsCvScoreList5.get(4);
			RemainingIncome = TotalIncome * val - TotalDebt;
		}
		// 8
		else if (efsCvScoreList1.get(5) <= cvScore && cvScore <= efsCvScoreList2.get(5)
				&& efsCvScoreList4.get(5) >= BehaviourScore) {
			val = efsCvScoreList5.get(5);
			RemainingIncome = TotalIncome * val - TotalDebt;
		}
		// 9
		else if (efsCvScoreList1.get(6) <= cvScore && cvScore <= efsCvScoreList2.get(6)
				&& efsCvScoreList4.get(6) < BehaviourScore) {
			val = efsCvScoreList5.get(6);
			RemainingIncome = TotalIncome * val - TotalDebt;
		}
		// 10
		else if (efsCvScoreList1.get(7) <= cvScore && cvScore <= efsCvScoreList2.get(7)
				&& (efsCvScoreList3.get(7) <= BehaviourScore)
				&& (BehaviourScore <= efsCvScoreList4.get(7))) {
			val = efsCvScoreList5.get(7);
			RemainingIncome = TotalIncome * val - TotalDebt;
		}
		// 11
		else if (efsCvScoreList1.get(8) <= cvScore && cvScore <= efsCvScoreList2.get(8)
				&& efsCvScoreList3.get(8) <= BehaviourScore && BehaviourScore <= efsCvScoreList4.get(8)) {
			val = efsCvScoreList5.get(8);
			RemainingIncome = TotalIncome * val - TotalDebt;
		}
		// 12
		else if (efsCvScoreList1.get(9) <= cvScore && cvScore <= efsCvScoreList2.get(9)
				&& (efsCvScoreList3.get(9) <= BehaviourScore)
				&& (BehaviourScore <= efsCvScoreList4.get(9))) {
			val = efsCvScoreList5.get(9);
			RemainingIncome = TotalIncome * val - TotalDebt;
		}
		// 13
		else if (efsCvScoreList1.get(10) <= cvScore 
				&& (efsCvScoreList3.get(10) <= BehaviourScore)
				&& (BehaviourScore <= efsCvScoreList4.get(10))) {
			val = efsCvScoreList5.get(10);
			RemainingIncome = TotalIncome * val - TotalDebt;
		}
		// 14
		else if (efsCvScoreList2.get(11) <= cvScore 
				&& efsCvScoreList4.get(11) >= BehaviourScore) {
			val = efsCvScoreList5.get(11);
			RemainingIncome = TotalIncome * val - TotalDebt;
		}
		// 15
		else if (efsCvScoreList1.get(12) <= cvScore && cvScore <= efsCvScoreList2.get(12)
				&& (efsCvScoreList4.get(12) == BehaviourScore)) {
			val = efsCvScoreList5.get(12);
			RemainingIncome = TotalIncome * val - TotalDebt;
		}

		//16
		else if (efsCvScoreList1.get(13) <= cvScore && cvScore <= efsCvScoreList2.get(13)
				&& (efsCvScoreList4.get(13) == BehaviourScore)) {
			val = efsCvScoreList5.get(13);
			RemainingIncome = TotalIncome * val - TotalDebt;
		}
		//17
		else if (efsCvScoreList1.get(14) <= cvScore && cvScore <= efsCvScoreList2.get(14)
				&& (efsCvScoreList3.get(14) <= BehaviourScore)
				&& (BehaviourScore <= efsCvScoreList4.get(14))) {
			val = efsCvScoreList5.get(14);
			RemainingIncome = TotalIncome * val - TotalDebt;
		}
		//18
		else if (efsCvScoreList1.get(15) <= cvScore && cvScore <= efsCvScoreList2.get(15)
				&& efsCvScoreList4.get(15) < BehaviourScore) {
			val = efsCvScoreList5.get(15);
			RemainingIncome = TotalIncome * val - TotalDebt;
		}
		
		//19
		else if (efsCvScoreList1.get(16) <= cvScore && cvScore <= efsCvScoreList2.get(16)
				&& (efsCvScoreList3.get(16) <= BehaviourScore)
				&& (BehaviourScore <= efsCvScoreList4.get(16))) {
			val = efsCvScoreList5.get(16);
			RemainingIncome = TotalIncome * val - TotalDebt;
		}
		//20
		else if (efsCvScoreList1.get(17) <= cvScore && cvScore <= efsCvScoreList2.get(17)
				&& (efsCvScoreList3.get(17) <= BehaviourScore)
				&& (BehaviourScore <= efsCvScoreList4.get(17))) {
			val = efsCvScoreList5.get(17);
			RemainingIncome = TotalIncome * val - TotalDebt;
		}
		//21
		else if (efsCvScoreList1.get(18) <= cvScore && cvScore <= efsCvScoreList2.get(18)
				&& (efsCvScoreList3.get(18) <= BehaviourScore)
				&& (BehaviourScore <= efsCvScoreList4.get(18))) {
			val = efsCvScoreList5.get(18);
			RemainingIncome = TotalIncome * val - TotalDebt;
		}
		//22
		else if (efsCvScoreList1.get(19) <= cvScore && cvScore <= efsCvScoreList2.get(19)
				&& efsCvScoreList4.get(19) < BehaviourScore) {
			val = efsCvScoreList5.get(19);
			RemainingIncome = TotalIncome * val - TotalDebt;
		}
		//23
		else if (efsCvScoreList1.get(20) <= cvScore && cvScore <= efsCvScoreList2.get(20)
				&& efsCvScoreList4.get(20) < BehaviourScore) {
			val = efsCvScoreList5.get(20);
			RemainingIncome = TotalIncome * val - TotalDebt;
		}
		//24
		else if (efsCvScoreList1.get(21) <= cvScore && cvScore <= efsCvScoreList2.get(21)
				&& (efsCvScoreList3.get(21) <= BehaviourScore)
				&& (BehaviourScore <= efsCvScoreList4.get(21))) {
			val = efsCvScoreList5.get(21);
			RemainingIncome = TotalIncome * val - TotalDebt;
		}
		System.out.println("Val =" + val + " " + "RemainingIncome =" + RemainingIncome);

	}


	public void switchtoIframe1() {
		WebDriverWait wait = new WebDriverWait(driver, 360, 0000);
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("iframe1"))));
		WebElement iframe1 = driver.findElement(By.xpath(prop.getProperty("iframe1")));

		driver.switchTo().frame(iframe1);

	}

	public void switchtoIframe2() {
		WebDriverWait wait2 = new WebDriverWait(driver, 360, 0000);
		wait2.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("iframe2"))));
		WebElement iframe2 = driver.findElement(By.xpath(prop.getProperty("iframe2")));
		driver.switchTo().frame(iframe2);

	}

	public void switchtoIframe3() {
		WebDriverWait waitframe = new WebDriverWait(driver, 360, 0000);
		waitframe.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("iframe3"))));
		WebElement iframe3 = driver.findElement(By.xpath(prop.getProperty("iframe3")));
		driver.switchTo().frame(iframe3);

	}

	public void switchtoIframe4() {
		WebDriverWait waitparties = new WebDriverWait(driver, 360, 0000);
		waitparties.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("iframe4"))));
		WebElement iframe4 = driver.findElement(By.xpath(prop.getProperty("iframe4")));
		driver.switchTo().frame(iframe4);

	}

	public void splIncreaseInterestRate() throws IOException, InterruptedException, DocumentException, ParseException {

		if (attemptNo == 0) {
			test = Extent.createTest("Interest Rate Calculation");
		}

		else {

			test = Extent.createTest("Resubmission Result: Attempt#  " + attemptNo + " - InterestRate Calculation");
		}

		
		 int spl=0;
		// SPLBuydown
			if (SPLBuydown.contains("0") || SPLBuydown.contains("1")) {
				double splbuy = Double.valueOf(SPLBuydown);
				spl = (int) splbuy;
			}
	

		File file = new File(System.getProperty("user.dir")
				+ "\\src\\main\\resources\\Excel\\Credit Decisions Calculations - SPL Increase.xlsx");

		FileInputStream inputStream = new FileInputStream(file);

		Workbook workbook = new XSSFWorkbook(inputStream);

		org.apache.poi.ss.usermodel.Sheet sheet = workbook.getSheet("SPL Interest Rate");


		// CV Score
		ArrayList<Integer> efsCvScoreList1 = new ArrayList<Integer>();
		ArrayList<Integer> efsCvScoreList2 = new ArrayList<Integer>();
		// Behavior Score
		ArrayList<Integer> efsCvScoreList3 = new ArrayList<Integer>();
		ArrayList<Integer> efsCvScoreList4 = new ArrayList<Integer>();
		for (int x = 2; x <= sheet.getLastRowNum(); x++) {
			int efsRange1 = (int) sheet.getRow(x).getCell(1).getNumericCellValue();
			efsCvScoreList1.add(efsRange1);
		}

		for (int x = 2; x <= sheet.getLastRowNum(); x++) {

			int efsRange2 = (int) sheet.getRow(x).getCell(3).getNumericCellValue();
			efsCvScoreList2.add(efsRange2);
		}

		for (int x = 2; x <= sheet.getLastRowNum(); x++) {

			int efsRange3 = (int) sheet.getRow(x).getCell(4).getNumericCellValue();
			efsCvScoreList3.add(efsRange3);
		}

		for (int x = 2; x <= sheet.getLastRowNum(); x++) {

			int efsRange4 = (int) sheet.getRow(x).getCell(6).getNumericCellValue();
			efsCvScoreList4.add(efsRange4);
		}

		ArrayList<Double> efsCvScoreList5 = new ArrayList<Double>();

		for (int x = 2; x <= sheet.getLastRowNum(); x++) {

			double efsRange5 = sheet.getRow(x).getCell(7).getNumericCellValue();
			efsCvScoreList5.add(efsRange5);
		}

		double rate = 0;
		// Check Cvscore and Behavioral Score
		if (efsCvScoreList2.get(0) < cvScore && efsCvScoreList4.get(0) < BehaviourScore) {
			rate = efsCvScoreList5.get(0);
		}

		else if ((efsCvScoreList2.get(1) < cvScore) && (efsCvScoreList3.get(1) <= BehaviourScore)
				&& (BehaviourScore <= efsCvScoreList4.get(1))) {
			rate = efsCvScoreList5.get(1);
		}

		else if ((efsCvScoreList2.get(2) < cvScore) && (efsCvScoreList3.get(2) <= BehaviourScore)
				&& (BehaviourScore <= efsCvScoreList4.get(2))) {
			rate = efsCvScoreList5.get(2);
		}

		else if ((efsCvScoreList2.get(3) < cvScore) && (efsCvScoreList3.get(3) <= BehaviourScore)
				&& (BehaviourScore <= efsCvScoreList4.get(3))) {
			rate = efsCvScoreList5.get(3);
		} else if ((efsCvScoreList2.get(4) < cvScore) && (efsCvScoreList3.get(4) <= BehaviourScore)
				&& (BehaviourScore <= efsCvScoreList4.get(4))) {
			rate = efsCvScoreList5.get(4);
		}

		else if ((efsCvScoreList2.get(5) < cvScore) && (BehaviourScore == efsCvScoreList4.get(5))) {
			rate = efsCvScoreList5.get(5);
		}
		//
		else if (efsCvScoreList1.get(6) <= cvScore && cvScore <= efsCvScoreList2.get(6)
				&& efsCvScoreList4.get(6) < BehaviourScore) {
			rate = efsCvScoreList5.get(6);
		} else if (efsCvScoreList1.get(7) <= cvScore && cvScore <= efsCvScoreList2.get(7)
				&& (efsCvScoreList3.get(7) <= BehaviourScore) && (BehaviourScore <= efsCvScoreList4.get(7))) {
			rate = efsCvScoreList5.get(7);
		} else if (efsCvScoreList1.get(8) <= cvScore && cvScore <= efsCvScoreList2.get(8)
				&& (efsCvScoreList3.get(8) <= BehaviourScore) && (BehaviourScore <= efsCvScoreList4.get(8))) {
			rate = efsCvScoreList5.get(8);
		} else if (efsCvScoreList1.get(9) <= cvScore && cvScore <= efsCvScoreList2.get(9)
				&& (efsCvScoreList3.get(9) <= BehaviourScore) && (BehaviourScore <= efsCvScoreList4.get(9))) {
			rate = efsCvScoreList5.get(9);
		} else if (efsCvScoreList1.get(10) <= cvScore && cvScore <= efsCvScoreList2.get(10)
				&& (efsCvScoreList3.get(10) <= BehaviourScore) && (BehaviourScore <= efsCvScoreList4.get(10))) {
			rate = efsCvScoreList5.get(10);
		} else if (efsCvScoreList1.get(11) <= cvScore && cvScore <= efsCvScoreList2.get(11)
				&& (BehaviourScore == efsCvScoreList4.get(11))) {
			rate = efsCvScoreList5.get(11);
		}
		//
		else if (efsCvScoreList1.get(12) <= cvScore && cvScore <= efsCvScoreList2.get(12)
				&& efsCvScoreList4.get(12) < BehaviourScore) {
			rate = efsCvScoreList5.get(12);
		} else if (efsCvScoreList1.get(13) <= cvScore && cvScore <= efsCvScoreList2.get(13)
				&& (efsCvScoreList3.get(13) <= BehaviourScore) && (BehaviourScore <= efsCvScoreList4.get(13))) {
			rate = efsCvScoreList5.get(13);
		} else if (efsCvScoreList1.get(14) <= cvScore && cvScore <= efsCvScoreList2.get(14)
				&& (efsCvScoreList3.get(14) <= BehaviourScore) && (BehaviourScore <= efsCvScoreList4.get(14))) {
			rate = efsCvScoreList5.get(14);
		} else if (efsCvScoreList1.get(15) <= cvScore && cvScore <= efsCvScoreList2.get(15)
				&& (efsCvScoreList3.get(15) <= BehaviourScore) && (BehaviourScore <= efsCvScoreList4.get(15))) {
			rate = efsCvScoreList5.get(15);
		} else if (efsCvScoreList1.get(16) <= cvScore && cvScore <= efsCvScoreList2.get(16)
				&& (efsCvScoreList3.get(16) <= BehaviourScore) && (BehaviourScore <= efsCvScoreList4.get(16))) {
			rate = efsCvScoreList5.get(16);
		} else if (efsCvScoreList1.get(17) <= cvScore && cvScore <= efsCvScoreList2.get(17)
				&& (BehaviourScore == efsCvScoreList4.get(17))) {
			rate = efsCvScoreList5.get(17);
		}
		//
		else if (efsCvScoreList1.get(18) <= cvScore && cvScore <= efsCvScoreList2.get(18)
				&& efsCvScoreList4.get(18) < BehaviourScore) {
			rate = efsCvScoreList5.get(18);
		} else if (efsCvScoreList1.get(19) <= cvScore && cvScore <= efsCvScoreList2.get(19)
				&& (efsCvScoreList3.get(19) <= BehaviourScore) && (BehaviourScore <= efsCvScoreList4.get(19))) {
			rate = efsCvScoreList5.get(19);
		} else if (efsCvScoreList1.get(20) <= cvScore && cvScore <= efsCvScoreList2.get(20)
				&& (efsCvScoreList3.get(20) <= BehaviourScore) && (BehaviourScore <= efsCvScoreList4.get(20))) {
			rate = efsCvScoreList5.get(20);
		} else if (efsCvScoreList1.get(21) <= cvScore && cvScore <= efsCvScoreList2.get(21)
				&& (efsCvScoreList3.get(21) <= BehaviourScore) && (BehaviourScore <= efsCvScoreList4.get(21))) {
			rate = efsCvScoreList5.get(21);
		} else if (efsCvScoreList1.get(22) <= cvScore && cvScore <= efsCvScoreList2.get(22)
				&& (efsCvScoreList3.get(22) <= BehaviourScore) && (BehaviourScore <= efsCvScoreList4.get(22))) {
			rate = efsCvScoreList5.get(22);
		} else if (efsCvScoreList1.get(23) <= cvScore && cvScore <= efsCvScoreList2.get(23)
				&& (BehaviourScore == efsCvScoreList4.get(23))) {
			rate = efsCvScoreList5.get(23);
		}
		//
		else if (efsCvScoreList1.get(24) <= cvScore && cvScore <= efsCvScoreList2.get(24)
				&& efsCvScoreList4.get(24) < BehaviourScore) {
			rate = efsCvScoreList5.get(24);
		} else if (efsCvScoreList1.get(25) <= cvScore && cvScore <= efsCvScoreList2.get(25)
				&& (efsCvScoreList3.get(25) <= BehaviourScore) && (BehaviourScore <= efsCvScoreList4.get(25))) {
			rate = efsCvScoreList5.get(25);
		} else if (efsCvScoreList1.get(26) <= cvScore && cvScore <= efsCvScoreList2.get(26)
				&& (efsCvScoreList3.get(26) <= BehaviourScore) && (BehaviourScore <= efsCvScoreList4.get(26))) {
			rate = efsCvScoreList5.get(26);
		} else if (efsCvScoreList1.get(27) <= cvScore && cvScore <= efsCvScoreList2.get(27)
				&& (efsCvScoreList3.get(27) <= BehaviourScore) && (BehaviourScore <= efsCvScoreList4.get(27))) {
			rate = efsCvScoreList5.get(27);
		} else if (efsCvScoreList1.get(28) <= cvScore && cvScore <= efsCvScoreList2.get(28)
				&& (efsCvScoreList3.get(28) <= BehaviourScore) && (BehaviourScore <= efsCvScoreList4.get(28))) {
			rate = efsCvScoreList5.get(28);
		} else if (efsCvScoreList1.get(29) <= cvScore && cvScore <= efsCvScoreList2.get(29)
				&& (BehaviourScore == efsCvScoreList4.get(29))) {
			rate = efsCvScoreList5.get(29);
		}
		//
		else if (efsCvScoreList1.get(30) <= cvScore && cvScore <= efsCvScoreList2.get(30)
				&& efsCvScoreList4.get(30) < BehaviourScore) {
			rate = efsCvScoreList5.get(30);
		} else if (efsCvScoreList1.get(31) <= cvScore && cvScore <= efsCvScoreList2.get(31)
				&& (efsCvScoreList3.get(31) <= BehaviourScore) && (BehaviourScore <= efsCvScoreList4.get(31))) {
			rate = efsCvScoreList5.get(31);
		} else if (efsCvScoreList1.get(32) <= cvScore && cvScore <= efsCvScoreList2.get(32)
				&& (efsCvScoreList3.get(32) <= BehaviourScore) && (BehaviourScore <= efsCvScoreList4.get(32))) {
			rate = efsCvScoreList5.get(32);
		}

		else if (efsCvScoreList1.get(33) <= cvScore && cvScore <= efsCvScoreList2.get(33)
				&& (efsCvScoreList3.get(33) <= BehaviourScore) && (BehaviourScore <= efsCvScoreList4.get(33))) {
			rate = efsCvScoreList5.get(33);
		} else if (efsCvScoreList1.get(34) <= cvScore && cvScore <= efsCvScoreList2.get(34)
				&& (BehaviourScore == efsCvScoreList4.get(34))) {
			rate = efsCvScoreList5.get(34);
		} else if (efsCvScoreList1.get(35) <= cvScore && cvScore <= efsCvScoreList2.get(35)) {
			rate = efsCvScoreList5.get(35);
		} else {
			System.out.println("Behavior Score Missing");
		}
		

		if(spl==1)
		{
			rate=rate-5.0;
		}
		
		// Displaying Interest Rate result
		System.out.println("Actual Interest rate: " + rate);
		System.out.println("Expected Interest rate: " + ExpInt);

		if (ExpInt == rate) {

			test.log(Status.PASS, MarkupHelper.createLabel("InterestRate Percentage Actual value : " + rate + "%",
					ExtentColor.GREEN));
			test.log(Status.PASS, MarkupHelper.createLabel("InterestRate Percentage Expected value : " + ExpInt + "%",
					ExtentColor.GREEN));

			test.log(Status.PASS, MarkupHelper.createLabel(" Interest Rate Calculation is Matching with GDS Decision",
					ExtentColor.GREEN));
			test.log(Status.PASS, "Snapshot below: " + test.addScreenCaptureFromPath(QLAInterestScreen));
			System.out.println("PASSED in Interest Verification");
		} else {
			System.out.println(IntRate + " is the Actual value.");

			test.log(Status.FAIL,
					MarkupHelper.createLabel("InterestRate Percentage Actual value : " + rate + "%", ExtentColor.RED));
			test.log(Status.FAIL, MarkupHelper.createLabel("InterestRate Percentage Expected value : " + ExpInt + "%",
					ExtentColor.RED));

			test.log(Status.FAIL, MarkupHelper.createLabel(" Interest Rate Calculation not Matching with GDS Decision",
					ExtentColor.RED));
			test.log(Status.FAIL, "Snapshot below: " + test.addScreenCaptureFromPath(QLAInterestScreen));
			System.out.println("FAILED in Interest Rate Verification");
		}
		Thread.sleep(3000);

	}

	public void checkContributerincrease() {
		// TODO Auto-generated method stub
		if (contributer.contains("Applicant 1")) {
			System.out.println("Credit_Contributer = Applicant 1");
			behscore = appbehscore;
		} else if (contributer.contains("Applicant 2")) {
			System.out.println("Credit_Contributer = Applicant 2");
			cvScore = cvScoreCoapp;
			Province = CoAppProvince;
			behscore = cobehscore;

		} else {
			System.out.println("Credit_Contributer = Shared");
			behscore = appbehscore;
		}
		if (behscore.contains(" ")) {
			System.out.println("BehaviourScore:" + behscore);
		}
		else if (behscore.contains("-1.00")){
			BehaviourScore= -1;
			System.out.println("BehaviourScore:" + behscore);
		}
		else {
			BehaviourScore = Integer.parseInt(behscore);
			System.out.println("BehaviourScore:" + BehaviourScore);
		}
	}
	
	public void MaxAfforableIncreaseQLA() throws IOException {
		File file = new File(System.getProperty("user.dir")
				+ "\\src\\main\\resources\\Excel\\Credit Decisions Calculations - SPL Increase.xlsx");

		FileInputStream inputStream = new FileInputStream(file);

		Workbook workbook = new XSSFWorkbook(inputStream);
		// String IntRate = String.valueOf(ExpInt);

		org.apache.poi.ss.usermodel.Sheet sheet = workbook.getSheet("Maximum Affordable QLA");

		Row r = sheet.getRow(1);
		int lastCol = r.getLastCellNum(); // Gets last column index
		int lastrow = sheet.getLastRowNum(); // Gets last row num
	

		String stringSplit[];

		System.out.println("Int Rate = " + IntRate + " " + "Province = " + Prov);
		//stringSplit = Province.split(" - ");
		//String Prov = stringSplit[0];
		System.out.println("cvScore :" + cvScore + "\t" + "BehaviourScore :" + BehaviourScore);
		System.out.println("RemainingIncome:" + RemainingIncome);
		int fcol = 0;
		// Iterating the row which Interest value for identifying the right table
		for (int i = 10; i <= lastCol; i++) {
			try {
				if (sheet.getRow(1).getCell(i).getStringCellValue().contains(IntRate)) {
					fcol = i; // this would be the first column(Province) in the table

					break;
				}
			}

			catch (Exception e) {

			}
			i += 3;
		}

		// Since Interest Rate 28.99 table has less rows we are providing some specific
		// conditions
		if (IntRate.contains("28.99")) {
			lastrow = 114;

		}
		int z = 0;
		// Iterating Province column
		for (int j = 3; j <= lastrow; j++) {

			if (sheet.getRow(j).getCell(fcol).getStringCellValue().contains(Prov)) {

				// Once Province matched, it would iterate Remaining Income column
				double remInvalue = sheet.getRow(j).getCell(fcol + 1).getNumericCellValue();

				if (z == 0) {

					if (remInvalue > RemainingIncome) {
						QLA = 0.0;
						break;

					}
				}
				if (!(z == 0)) {
					if (remInvalue > RemainingIncome) {
						QLA = sheet.getRow(j - 1).getCell(fcol + 2).getNumericCellValue();
						break;

					}
				}

				// This block is for last line of the table alone
				if (j == sheet.getLastRowNum()) {
					QLA = sheet.getRow(sheet.getLastRowNum()).getCell(fcol + 2).getNumericCellValue();

				}

				// This block would be executed if the calculated remaining income is greater
				// than the maximum
				// remaining income in the table

				if (!sheet.getRow(j + 1).getCell(fcol).getStringCellValue().contains(Prov)) {
					if (remInvalue < RemainingIncome) {
						// Its picking the maximum value available
						QLA = sheet.getRow(j).getCell(fcol + 2).getNumericCellValue();
						break;
					}
				}
				z++;
			}

		}
		System.out.println("QLA from sheet: " + QLA);
		if(!(ExpectedQLA==QLA+100))
		{
			// CV Score
						ArrayList<Integer> efsCvScoreList1 = new ArrayList<Integer>();
						ArrayList<Integer> efsCvScoreList2 = new ArrayList<Integer>();
						// Behavior Score
						ArrayList<Integer> efsCvScoreList3 = new ArrayList<Integer>();
						ArrayList<Integer> efsCvScoreList4 = new ArrayList<Integer>();
						// QLA
						ArrayList<Double> efsCvScoreList5 = new ArrayList<Double>();
						ArrayList<Double> efsCvScoreList6 = new ArrayList<Double>();

						for (int x = 10; x < 33; x++) {
							int efsRange1 = (int) sheet.getRow(x).getCell(0).getNumericCellValue();

							efsCvScoreList1.add(efsRange1);
						}

						for (int x = 10; x < 33; x++) {

							int efsRange2 = (int) sheet.getRow(x).getCell(2).getNumericCellValue();
							efsCvScoreList2.add(efsRange2);
						}

						// Behavior Score list
						for (int x = 10; x < 33; x++) {

							int efsRange3 = (int) sheet.getRow(x).getCell(3).getNumericCellValue();
							efsCvScoreList3.add(efsRange3);
						}

						for (int x = 10; x < 33; x++) {

							int efsRange4 = (int) sheet.getRow(x).getCell(5).getNumericCellValue();
							efsCvScoreList4.add(efsRange4);
						}

						// QLA

						for (int x = 10; x < 33; x++) {

							double efsRange6 = sheet.getRow(x).getCell(7).getNumericCellValue();
							efsCvScoreList5.add(efsRange6);
						}
						// Reset Values

						for (int x = 10; x < 33; x++) {

							double efsRange6 = sheet.getRow(x).getCell(8).getNumericCellValue();
							efsCvScoreList6.add(efsRange6);
						}

						//11
						if (efsCvScoreList2.get(0) <= cvScore && efsCvScoreList4.get(0) <= BehaviourScore
								&& QLA > efsCvScoreList5.get(0)) {
							QLA = efsCvScoreList6.get(0);

						}
						// 12
						else if (efsCvScoreList2.get(1) <= cvScore 
								&& efsCvScoreList3.get(1) <= BehaviourScore && BehaviourScore <= efsCvScoreList4.get(1)
								&& QLA > efsCvScoreList5.get(1)) {
							QLA = efsCvScoreList6.get(1);

						}
						// 13
						else if (efsCvScoreList2.get(2) <= cvScore 
								&& efsCvScoreList3.get(2) <= BehaviourScore && BehaviourScore <= efsCvScoreList4.get(2)
								&& QLA > efsCvScoreList5.get(2)) {
							QLA = efsCvScoreList6.get(2);

						}
						// 14
						else if (efsCvScoreList2.get(3) <= cvScore 
								&& efsCvScoreList3.get(3) <= BehaviourScore && BehaviourScore <= efsCvScoreList4.get(3)
								&& QLA > efsCvScoreList5.get(3)) {
							QLA = efsCvScoreList6.get(3);

						}
						// 15
						else if (efsCvScoreList1.get(4) <= cvScore && cvScore <= efsCvScoreList2.get(4)
								&& efsCvScoreList4.get(4) <= BehaviourScore && QLA > efsCvScoreList5.get(4)) {
							QLA = efsCvScoreList6.get(4);

						}
						// 16
						else if (efsCvScoreList1.get(5) <= cvScore && cvScore <= efsCvScoreList2.get(5)
								&& efsCvScoreList3.get(5) <= BehaviourScore && BehaviourScore <= efsCvScoreList4.get(5)
								&& QLA > efsCvScoreList5.get(5)) {
							QLA = efsCvScoreList6.get(5);

						}

						// 17
						else if (efsCvScoreList1.get(6) <= cvScore && cvScore <= efsCvScoreList2.get(6)
								&& efsCvScoreList3.get(6) <= BehaviourScore && BehaviourScore <= efsCvScoreList4.get(6)
								&& QLA > efsCvScoreList5.get(6)) {
							QLA = efsCvScoreList6.get(6);

						}
						// 18
						else if (efsCvScoreList1.get(7) <= cvScore && cvScore <= efsCvScoreList2.get(7)
								&& efsCvScoreList3.get(7) <= BehaviourScore && BehaviourScore <= efsCvScoreList4.get(7)
								&& QLA > efsCvScoreList5.get(7)) {
							QLA = efsCvScoreList6.get(7);

						}
						// 19
						else if (efsCvScoreList1.get(8) <= cvScore && cvScore <= efsCvScoreList2.get(8)
								&& efsCvScoreList3.get(8) <= BehaviourScore && BehaviourScore <= efsCvScoreList4.get(8)
								&& QLA > efsCvScoreList5.get(8)) {
							QLA = efsCvScoreList6.get(8);

						}
						// 20
						else if (efsCvScoreList1.get(9) <= cvScore && cvScore <= efsCvScoreList2.get(9)
								&& efsCvScoreList3.get(9) <= BehaviourScore && BehaviourScore <= efsCvScoreList4.get(9)
								&& QLA > efsCvScoreList5.get(9)) {
							QLA = efsCvScoreList6.get(9);

						}
						// 21
						else if (efsCvScoreList1.get(10) <= cvScore && cvScore <= efsCvScoreList2.get(10)
								&& efsCvScoreList3.get(10) <= BehaviourScore && BehaviourScore <= efsCvScoreList4.get(10)
								&& QLA > efsCvScoreList5.get(10)) {
							QLA = efsCvScoreList6.get(10);

						}
						// 22
						else if (efsCvScoreList1.get(11) <= cvScore && cvScore <= efsCvScoreList2.get(11)
								&& efsCvScoreList3.get(11) <= BehaviourScore && BehaviourScore <= efsCvScoreList4.get(11) 
								&& QLA > efsCvScoreList5.get(11)) {
							QLA = efsCvScoreList6.get(11);

						}
						// 23
						else if (efsCvScoreList2.get(12) <= cvScore && efsCvScoreList3.get(12) <= BehaviourScore
								&& BehaviourScore <= efsCvScoreList4.get(12) && QLA > efsCvScoreList5.get(12)) {
							QLA = efsCvScoreList6.get(12);

						}
						// 24
						else if (efsCvScoreList2.get(13) <= cvScore && efsCvScoreList4.get(13) >= BehaviourScore
								&& QLA > efsCvScoreList5.get(13)) {
							QLA = efsCvScoreList6.get(13);

						}
						// 25
						else if (efsCvScoreList2.get(14) <= cvScore && efsCvScoreList4.get(14) == BehaviourScore
								&& QLA > efsCvScoreList5.get(14)) {
							QLA = efsCvScoreList6.get(14);

						}
						// 26
						else if (efsCvScoreList1.get(15) <= cvScore && cvScore <= efsCvScoreList2.get(15)
								&& efsCvScoreList4.get(15) == BehaviourScore
								 && QLA > efsCvScoreList5.get(15)) {
							QLA = efsCvScoreList6.get(15);

						}
						// 27
						else if (efsCvScoreList1.get(16) <= cvScore && cvScore <= efsCvScoreList2.get(16) 
								&& BehaviourScore <= efsCvScoreList4.get(16) 
								&& QLA > efsCvScoreList5.get(16)) {
							QLA = efsCvScoreList6.get(16);

						}
						// 28
						else if (efsCvScoreList1.get(17) <= cvScore && cvScore <= efsCvScoreList2.get(17)
								&& efsCvScoreList3.get(17) <= BehaviourScore && BehaviourScore <= efsCvScoreList4.get(17)
								&& QLA > efsCvScoreList5.get(17)) {
							QLA = efsCvScoreList6.get(17);

						}
						// 29
						else if (efsCvScoreList1.get(18) <= cvScore && cvScore <= efsCvScoreList2.get(18) 
								&& efsCvScoreList4.get(18) == BehaviourScore
								&& QLA > efsCvScoreList5.get(18)) {
							QLA = efsCvScoreList6.get(18);

						}
						// 30
						else if (efsCvScoreList1.get(19) <= cvScore && cvScore <= efsCvScoreList2.get(19)
								&& efsCvScoreList4.get(19) == BehaviourScore 
								&& QLA > efsCvScoreList5.get(19)) {
							QLA = efsCvScoreList6.get(19);

						}
						// 31
						else if (efsCvScoreList1.get(20) <= cvScore && cvScore <= efsCvScoreList2.get(20)
								&& efsCvScoreList4.get(20) >= BehaviourScore
								&& QLA > efsCvScoreList5.get(20)) {
							QLA = efsCvScoreList6.get(20);

						}

						//32
						else if (efsCvScoreList1.get(21) <= cvScore && cvScore <= efsCvScoreList2.get(21)
								&& efsCvScoreList4.get(21) > BehaviourScore
								&& QLA > efsCvScoreList5.get(21)) {
							QLA = efsCvScoreList6.get(21);

						}
						//33
						else if (efsCvScoreList1.get(22) <= cvScore && cvScore <= efsCvScoreList2.get(22)
								&& efsCvScoreList4.get(22) > BehaviourScore
								&& QLA > efsCvScoreList5.get(22)) {
							QLA = efsCvScoreList6.get(22);

						}
		System.out.println("SPL QLA Final is " + QLA);
		}

	}

	public void urbancode() throws FileNotFoundException {

		String ps1 = postalcode.substring(0, 3);
		String ps2 = " ";
		String ps3 = postalcode.substring(3, 6);

		ps = ps1 + ps2 + ps3;
		System.out.println("Checking Urban code: "+ps);
		InputStream is = new FileInputStream(
				new File(System.getProperty("user.dir") + "\\src\\main\\resources\\Excel\\FNF Urban Code.xlsx"));
		Workbook wb = StreamingReader.builder().sstCacheSize(100).open(is);
		org.apache.poi.ss.usermodel.Sheet sheet = wb.getSheet("FNF goeasy urbanization");

		Iterator<Row> rows = sheet.iterator();
		label: while (rows.hasNext()) {
			Row row = rows.next();
			Iterator<Cell> cell = row.cellIterator();
			Cell value = cell.next();
			// System.out.print(value.getStringCellValue());
			// System.out.println("");

			if (value.getStringCellValue().toLowerCase().contains(ps.toLowerCase())) {
				int i = 0;

				while (cell.hasNext()) {

					value = cell.next();
					if (i == 2) {
						code = value.getStringCellValue();
						System.out.println(code);
						break label;
					}

					i++;
				}

			}

		}
	}
	public void splFinalIncreaseQLA() throws IOException, InterruptedException {

		if (attemptNo == 0) {
			test = Extent.createTest("QLA Calculation");
		}

		else {

			test = Extent.createTest("Resubmission Result: Attempt#  " + attemptNo + " - QLA Calculation");
		}
		// Variable declaration. This needs to be done in Class level
		ArrayList rangeList;
		ArrayList<Double> efsCvScoreList = new ArrayList<Double>();
		ArrayList<String> RiskgroupList = new ArrayList<String>();

		String range, riskGroup = " ", propertyRange, riskGroupbeh = " ";
		Double drange, propRange, propRange1, propRange2;
		int thecol = 0, therow = 0, tabrow = 0;
		String stringSplitter[];

		// Sheet initalization.This needs to be done once in Class level, so that, we
		// dont have to initialize this in each function
		org.apache.poi.ss.usermodel.Sheet sheet;

		File file = new File(System.getProperty("user.dir")
				+ "\\src\\main\\resources\\Excel\\Credit Decisions Calculations - SPL Increase.xlsx");

		FileInputStream inputStream = new FileInputStream(file);

		Workbook workbook = new XSSFWorkbook(inputStream);

		sheet = workbook.getSheet("Max LTV");

		// Reading cv score from B column

		for (int i = 2; i < 9; i++) {
			range = sheet.getRow(i).getCell(1).getStringCellValue();

			stringSplitter = range.split("&");

			if (i == 2) // This is for first cell alone. To remove the text content in the first cell
			{
				stringSplitter = range.split("&");
				String range1 = stringSplitter[0];
				range1 = range1.replace("efs_cv_score >=", "");
				drange = Double.parseDouble(range1);
				efsCvScoreList.add(drange);
				continue;
			}

			String range1 = stringSplitter[0];
			range1 = range1.replace(">=", " ");
			drange = Double.parseDouble(range1);
			efsCvScoreList.add(drange);

			String range2 = stringSplitter[1];
			range2 = range2.replace("<=", "");
			drange = Double.parseDouble(range2);
			efsCvScoreList.add(drange);

			// Splitting the numbers and storing it in an arraylist

		}

		// [683.0, 646.0, 682.0, 627.0, 645.0, 610.0, 626.0, 593.0, 609.0, 577.0, 592.0,
		// 564.0, 576.0]

		String rangerisk1 = sheet.getRow(125).getCell(0).getStringCellValue();
		System.out.println(rangerisk1);
		// Assigning the risk group based on applicant's efs score

		// RiskgroupList
		for (int i = 2; i < 9; i++) {
			String rangerisk = sheet.getRow(i).getCell(2).getStringCellValue();
			RiskgroupList.add(rangerisk);
		}

		if (cvScore >= efsCvScoreList.get(0)) // 683
		{
			riskGroup = (String) RiskgroupList.get(0);
		} else if ((cvScore >= efsCvScoreList.get(1)) && (cvScore <= efsCvScoreList.get(2))) // 646<=cvScore<=682
		{
			riskGroup = (String) RiskgroupList.get(1);
		} else if ((cvScore >= efsCvScoreList.get(3)) && (cvScore <= efsCvScoreList.get(4))) // 627<=cvScore<=645
		{
			riskGroup = (String) RiskgroupList.get(2);
		} else if ((cvScore >= efsCvScoreList.get(5)) && (cvScore <= efsCvScoreList.get(6))) // 610<=cvScore<=626
		{
			riskGroup = (String) RiskgroupList.get(3);
		} else if ((cvScore >= efsCvScoreList.get(7)) && (cvScore <= efsCvScoreList.get(8))) // 6593<=cvScore<=609
		{
			riskGroup = (String) RiskgroupList.get(4);
		} else if ((cvScore >= efsCvScoreList.get(9)) && (cvScore <= efsCvScoreList.get(10))) // 577<=cvScore<=592
		{
			riskGroup = (String) RiskgroupList.get(5);
		} else if ((cvScore >= efsCvScoreList.get(11)) && (cvScore <= efsCvScoreList.get(12))) // 564<=cvScore<=576
		{
			riskGroup = (String) RiskgroupList.get(6);
		}

		System.out.println("EFS CV Risk group: " + riskGroup);

		// Identify Behavior Group BehaviourScore

		ArrayList<Double> behaviorScoreList = new ArrayList<Double>();
		ArrayList<String> Riskgroupbehavior = new ArrayList<String>();
		double rep, rep1;
		// behaviorScoreList
		for (int i = 3; i < 10; i++) {
			DataFormatter formatter = new DataFormatter();
			String rangerisk = formatter.formatCellValue(sheet.getRow(i).getCell(6));

			if (i == 3) {

				String repe = rangerisk.replace("<=", "");
				rep = Double.parseDouble(repe);
				behaviorScoreList.add(rep);
			}
			if (i > 3 && i < 8) {

				String stringSplit[] = rangerisk.split("-");
				String split1 = stringSplit[0];
				rep = Double.parseDouble(split1);
				behaviorScoreList.add(rep);
				String split2 = stringSplit[1];
				rep1 = Double.parseDouble(split2);
				behaviorScoreList.add(rep1);
			}
			if (i == 8) {

				String repe = rangerisk.replace(">", "");
				rep = Double.parseDouble(repe);
				behaviorScoreList.add(rep);
			}
			if (i == 9) {

				rep = Double.parseDouble(rangerisk);
				behaviorScoreList.add(rep);
			}

			// behaviorScoreList.add(rep);
		}

		// RiskgroupList
		for (int i = 3; i < 11; i++) {
			String rangerisk = sheet.getRow(i).getCell(7).getStringCellValue();
			Riskgroupbehavior.add(rangerisk);
		}

		if ((BehaviourScore >= 0) && BehaviourScore <= behaviorScoreList.get(0)) // 0 to 557
		{
			riskGroupbeh = (String) Riskgroupbehavior.get(0);
		} else if ((BehaviourScore >= behaviorScoreList.get(1)) && (BehaviourScore <= behaviorScoreList.get(2))) // 558<=cvScore<=595
		{
			riskGroupbeh = (String) Riskgroupbehavior.get(1);
		} else if ((BehaviourScore >= behaviorScoreList.get(3)) && (BehaviourScore <= behaviorScoreList.get(4))) // 596<=cvScore<=634
		{
			riskGroupbeh = (String) Riskgroupbehavior.get(2);
		} else if ((BehaviourScore >= behaviorScoreList.get(5)) && (BehaviourScore <= behaviorScoreList.get(6))) // 635<=cvScore<=658
		{
			riskGroupbeh = (String) Riskgroupbehavior.get(3);
		} else if ((BehaviourScore >= behaviorScoreList.get(7)) && (BehaviourScore <= behaviorScoreList.get(8))) // 659<=cvScore<=697
		{
			riskGroupbeh = (String) Riskgroupbehavior.get(4);
		} else if ((BehaviourScore > behaviorScoreList.get(9))) // >697
		{
			riskGroupbeh = (String) Riskgroupbehavior.get(5);
		} else if ((BehaviourScore == behaviorScoreList.get(10))) // -1
		{
			riskGroupbeh = (String) Riskgroupbehavior.get(6);
		}

		else // Blank
		{
			riskGroupbeh = (String) Riskgroupbehavior.get(7);
		}

		System.out.println("Behaviour Risk group: " + riskGroupbeh);

		// Iterating through the tables to identify the CV Risk Group
		for (int j = 16; j < 737; j++) // 9 is the first row where table starts & 86 is the last row in the table
		{

			if (sheet.getRow(j).getCell(0).getStringCellValue().contains(riskGroup)
					&& sheet.getRow(j + 1).getCell(0).getStringCellValue().contains(riskGroupbeh)) {
				tabrow = j;
				for (int k = 0; k < 20; k++) {
					propertyRange = sheet.getRow(j + 2).getCell(k).getStringCellValue();

					if (k == 0) {
						propertyRange = propertyRange.replace("Property Value ", "");
						propertyRange = propertyRange.replace("<$", "");
						propertyRange = propertyRange.replace("K", "000");
						propRange = Double.parseDouble(propertyRange);

						if (PropertyVal < propRange) {
							thecol = k;
							break;
						}
					}

					if (k == 5 && riskGroup.contains("Risk Group 1")) // Since, Risk Group 1 has only two property type
																		// tables
					{

						propertyRange = propertyRange.replace("Property Value ", "");
						propertyRange = propertyRange.replace(">= $", "");
						propertyRange = propertyRange.replace("K", "000");
						propRange = Double.parseDouble(propertyRange);

						if (PropertyVal >= propRange) {
							thecol = k;
							break;
						}

					}

					if (k == 5 && !riskGroup.contains("Risk Group 1")) {

						propertyRange = propertyRange.replace("Property Value ", "");
						propertyRange = propertyRange.replace("<$", "");
						propertyRange = propertyRange.replace(">= $", "");
						propertyRange = propertyRange.replace("K", "000");
						stringSplitter = propertyRange.split("and");

						String range1 = stringSplitter[0];
						String range2 = stringSplitter[1];

						propRange1 = Double.parseDouble(range1);
						propRange2 = Double.parseDouble(range2);

						if ((PropertyVal >= propRange1) && (PropertyVal < propRange2)) {
							thecol = k;
							break;
						}

					}

					if (k == 10) {

						propertyRange = propertyRange.replace("Property Value ", "");
						propertyRange = propertyRange.replace("<$", "");
						propertyRange = propertyRange.replace(">=$", "");
						propertyRange = propertyRange.replace("K", "000");
						stringSplitter = propertyRange.split("and");
						String range1 = stringSplitter[0];
						String range2 = stringSplitter[1];

						propRange1 = Double.parseDouble(range1);
						propRange2 = Double.parseDouble(range2);

						if ((PropertyVal >= propRange1) && (PropertyVal < propRange2)) {
							thecol = k;
							break;
						}
					}

					if (k == 15) {
						propertyRange = propertyRange.replace("Property Value ", "");
						propertyRange = propertyRange.replace(">=$", "");
						propertyRange = propertyRange.replace("K", "000");
						propRange = Double.parseDouble(propertyRange);

						if (PropertyVal >= propRange) {
							thecol = k;
							break;
						}
					}

					k += 4;

				}

			}
			j += 14;
		}

		int r = tabrow + 5;
		try {
			while (sheet.getRow(r).getCell(thecol).getCellTypeEnum() == CellType.STRING) {

				if (sheet.getRow(r).getCell(thecol).getStringCellValue().toLowerCase()
						.contains(propertyType.toLowerCase())) {
					therow = r;
					break;
				}
				r++;
			}
		} catch (Exception e) {

		}

		if (code.equalsIgnoreCase("Urban")) {
			Maxltv = sheet.getRow(therow).getCell(thecol + 1).getNumericCellValue();
		}

		else if (code.equalsIgnoreCase("Rural")) {
			Maxltv = sheet.getRow(therow).getCell(thecol + 2).getNumericCellValue();
		} else if (code.equalsIgnoreCase("Remote")) {
			Maxltv = sheet.getRow(therow).getCell(thecol + 3).getNumericCellValue();
		}

		System.out.println("Max LTV is " + Maxltv);

		// Home Equity Calculation
		HomeEquity = (LtvMax - SPLltv) * PropertyVal / 100;
		System.out.println("Home Equity =" + HomeEquity);
		double ActualQLA = 0;

		if (QLA == 0.0) {
			ActualQLA = QLA;
		} else if (HomeEquity > QLA) {
			ActualQLA = QLA + 100;
		} else if (HomeEquity < QLA) {
			ActualQLA = HomeEquity + 100;
		}

		System.out.println("Actual QLA :$" + ActualQLA);
		System.out.println("Expected QLA :$" + ExpectedQLA);


		// Displaying QLA result

		if (ActualQLA == ExpectedQLA) {

			test.log(Status.PASS, MarkupHelper.createLabel("QLA Actual value :  $" + ActualQLA, ExtentColor.GREEN));
			test.log(Status.PASS,
					MarkupHelper.createLabel("QLA Expected value :  $ " + ExpectedQLA, ExtentColor.GREEN));

			test.log(Status.PASS,
					MarkupHelper.createLabel(" QLA Calculation is Matching with GDS Decision", ExtentColor.GREEN));
			test.log(Status.PASS, "Snapshot below: " + test.addScreenCaptureFromPath(QLAInterestScreen));

			System.out.println("PASSED in QLA Verification");
		} else {
			System.out.println(ExpectedQLA + " is the expected value.");

			test.log(Status.FAIL, MarkupHelper.createLabel("QLA Actual value :  $" + ActualQLA, ExtentColor.RED));
			test.log(Status.FAIL, MarkupHelper.createLabel("QLA Expected value :  $ " + ExpectedQLA, ExtentColor.RED));

			test.log(Status.FAIL,
					MarkupHelper.createLabel(" QLA Calculation not Matching with GDS Decision", ExtentColor.RED));
			test.log(Status.FAIL, "Snapshot below: " + test.addScreenCaptureFromPath(QLAInterestScreen));

			System.out.println("FAILED in QLA Verification");
		}
	}



	// Check Max H&A
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

		if (ExpectedQLA < qla1) {
			ActualMaxHA = 0.0;
		} else if (ExpectedQLA >= qla1 && ExpectedQLA <= qla2) {
			ActualMaxHA = max1;
		} else if (ExpectedQLA >= qla3 && ExpectedQLA <= qla4) {
			ActualMaxHA = max2;
		} else if (ExpectedQLA >= qla5 && ExpectedQLA <= qla6) {
			ActualMaxHA = max3;
		} else if (ExpectedQLA >= qla7 && ExpectedQLA <= qla8) {
			ActualMaxHA = max4;
		} else if (ExpectedQLA >= qla9 && ExpectedQLA <= qla10) {
			ActualMaxHA = max5;
		} else if (ExpectedQLA >= qla11 && ExpectedQLA <= qla12) {
			ActualMaxHA = max6;
		}
		if(Prov.equalsIgnoreCase("MB"))
		{
			ActualMaxHA=0.0;
		}
		// Displaying Interest Rate result
		System.out.println("Actual MaxH&A: " + ActualMaxHA);
		System.out.println("Expected MaxH&A: " + ExpectedMaxHA);

		if (ExpectedMaxHA == ActualMaxHA) {

			test.log(Status.PASS,
					MarkupHelper.createLabel("MaxH&A Actual value : " + ActualMaxHA + "%", ExtentColor.GREEN));
			test.log(Status.PASS,
					MarkupHelper.createLabel("MaxH&A Expected value : " + ExpectedMaxHA + "%", ExtentColor.GREEN));

			test.log(Status.PASS,
					MarkupHelper.createLabel(" MaxH&A Calculation is Matching with GDS Decision", ExtentColor.GREEN));
			test.log(Status.PASS, "Snapshot below: " + test.addScreenCaptureFromPath(MaxHASPLFullScreen));
			System.out.println("PASSED in MaxH&A Verification");
		} else {

			test.log(Status.FAIL,
					MarkupHelper.createLabel("MaxH&A Actual value : " + ActualMaxHA + "%", ExtentColor.RED));
			test.log(Status.FAIL,
					MarkupHelper.createLabel("MaxH&A Expected value : " + ExpectedMaxHA + "%", ExtentColor.RED));

			test.log(Status.FAIL,
					MarkupHelper.createLabel(" MaxH&A Calculation not Matching with GDS Decision", ExtentColor.RED));
			test.log(Status.FAIL, "Snapshot below: " + test.addScreenCaptureFromPath(MaxHASPLFullScreen));
			System.out.println("FAILED in MaxH&A Verification");
		}
		Thread.sleep(3000);
	}

	public void ReasonCode() throws InterruptedException, IOException {
		// TODO Auto-generated method stub //MaxHASPLFullScreen
		if (attemptNo == 0) {
			test = Extent.createTest("Reason Codes/Risk Factors");
		}

		else {

			test = Extent.createTest("Resubmission Result: Attempt#  " + attemptNo + " - Reason Codes");
		}
		
if(length==1)
{
	test.log(Status.PASS,
			MarkupHelper.createLabel(" Reason Codes/Risk Factors with GDS Decision", ExtentColor.GREEN));
	test.log(Status.PASS,
			"Snapshot below: " + test.addScreenCaptureFromPath(MaxHASPLFullScreen) +  test.addScreenCaptureFromPath(ReasonCodeSPLFullScreen1));
}
else if (length==2)
{
	test.log(Status.PASS,
			MarkupHelper.createLabel(" Reason Codes/Risk Factors with GDS Decision", ExtentColor.GREEN));
	test.log(Status.PASS,
			"Snapshot below: " + test.addScreenCaptureFromPath(MaxHASPLFullScreen) + test.addScreenCaptureFromPath(ReasonCodeSPLFullScreen1)
					+ test.addScreenCaptureFromPath(ReasonCodeSPLFullScreen2));
}
else if (length==3)
{
	test.log(Status.PASS,
			MarkupHelper.createLabel(" Reason Codes/Risk Factors with GDS Decision", ExtentColor.GREEN));
	test.log(Status.PASS,
			"Snapshot below: " + test.addScreenCaptureFromPath(MaxHASPLFullScreen) + test.addScreenCaptureFromPath(ReasonCodeSPLFullScreen1)
					+ test.addScreenCaptureFromPath(ReasonCodeSPLFullScreen2)
					+ test.addScreenCaptureFromPath(ReasonCodeSPLFullScreen3));
}
else {
	test.log(Status.PASS,
			MarkupHelper.createLabel(" Reason Codes/Risk Factors with GDS Decision", ExtentColor.GREEN));
	test.log(Status.PASS,
			"Snapshot below: " + test.addScreenCaptureFromPath(MaxHASPLFullScreen) + test.addScreenCaptureFromPath(ReasonCodeSPLFullScreen1)
					+ test.addScreenCaptureFromPath(ReasonCodeSPLFullScreen2));
}
		Thread.sleep(3000);
	}

	public void SecondPopupSpl() throws Exception {

		attemptNo++;
		JFrame frmOpt = new JFrame(); // We are declaring the frame
		frmOpt.setAlwaysOnTop(true);// This is the line for displaying it above all windows

		Thread.sleep(1000);
		String s = "<html>Press 1 for Re-Submission with Co Applicant<br>Press 2 for Re-Submission with Removal of Co Applicant<br>";
		s += "Press 3 for Results</html>";

		JLabel label = new JLabel(s);
		JTextPane jtp = new JTextPane();
		jtp.setSize(new Dimension(480, 10));
		jtp.setPreferredSize(new Dimension(480, jtp.getPreferredSize().height));
		label.setFont(new Font("Arial", Font.BOLD, 20));
		UIManager.put("OptionPane.minimumSize", new Dimension(500, 200));
		UIManager.put("TextField.font", new FontUIResource(new Font("Verdana", Font.BOLD, 18)));
		// Getting Input from user

		String option = JOptionPane.showInputDialog(frmOpt, label);

		int useroption = Integer.parseInt(option);

		switch (useroption) {

		case 1:

			// Function for Re-Submission
			System.out.println("Re-Submission with Co-Applicant");
			
				resubmitForDecisionSpl();
			
			break;
		case 2:

			// Function for Removal Co App
			System.out.println("Re-Submission with Removal of Co Applicant");
				removeCoAppspl();
		
			break;
		case 3:

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
	
	public void removeCoAppspl() throws Exception
	{
		System.out.println("Re-Submission Attempt:"+attemptNo);

		firstPopup();

		Thread.sleep(3000);
		driver.get(url);
		getAddress();
		getPartyDetails();
		getCollateral();

		getAppDetails();
		
		calculateIncome();
		calculateSPLLiability();
		splLTV();
		splIncreaseRemInc();
		urbancode();
		MaxAfforableIncreaseQLA();
		splFinalIncreaseQLA();
		maxHA();
		ReasonCode();
		Thread.sleep(4000);
		SecondPopupSpl();
	}
	public void getPartyDetails()
			throws InterruptedException, HeadlessException, UnsupportedFlavorException, IOException {
		// PartyDetails
		WebDriverWait waitPartyDetails = new WebDriverWait(driver, 360, 0000);
		waitPartyDetails
				.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("PartyDetails"))));
		driver.findElement(By.xpath(prop.getProperty("PartyDetails"))).click();

		// EmploymentandIncome

		WebDriverWait waitEmploymentandIncome = new WebDriverWait(driver, 360, 0000);
		waitEmploymentandIncome.until(
				ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("EmploymentandIncome"))));
		driver.findElement(By.xpath(prop.getProperty("EmploymentandIncome"))).click();
		driver.switchTo().defaultContent();
		switchtoIframe1();
		switchtoIframe3();
		// Income table
		Thread.sleep(5000);
		WebDriverWait waitMonthlyIncometable = new WebDriverWait(driver, 360, 0000);
		waitMonthlyIncometable
				.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("MonthlyIncometable"))));

		WebElement MonthlyIncometable = driver.findElement(By.xpath(prop.getProperty("MonthlyIncometable")));
		List<WebElement> rowValsIncome = MonthlyIncometable.findElements(By.tagName("tr"));
		int rowNumIncome = MonthlyIncometable.findElements(By.tagName("tr")).size();

		String str = null;
		double ApplicantIncome = 0, OtherIncomeValue=0.0;
		for (int i = 0; i < rowNumIncome; i++) {

			double subValue = 0;
			// Get each row's column values by tag name
			List<WebElement> colValsIncome = rowValsIncome.get(i).findElements(By.tagName("td"));
			WebElement Income = colValsIncome.get(4);
			Actions act = new Actions(driver);
			Income.click();
			Thread.sleep(1000);
			Income.click();
			
			act.keyDown(Keys.CONTROL).sendKeys("a").keyUp(Keys.CONTROL).build().perform();
			act.keyDown(Keys.CONTROL).sendKeys("c").keyUp(Keys.CONTROL).build().perform();

			String Incomeamount = (String) Toolkit.getDefaultToolkit().getSystemClipboard()
					.getData(DataFlavor.stringFlavor);

			if (Incomeamount.contains(","))

			{
				str = Incomeamount.replace(",", "");
				subValue = Double.parseDouble(str.replace("$", ""));
			} else {
				subValue = Double.parseDouble(Incomeamount.replace("$", ""));
			}
			ApplicantIncome += subValue;
		}
		
		// otherIncometable
		WebElement otherIncometable = driver.findElement(By.xpath(prop.getProperty("otherIncometable")));
		List<WebElement> rowValsOther = otherIncometable.findElements(By.tagName("tr"));
		int rowNumOther = otherIncometable.findElements(By.tagName("tr")).size();
		String strOther = null;
		for (int i = 0; i < rowNumOther; i++) {

			double subValue = 0;
			// Get each row's column values by tag name
			List<WebElement> colValOther = rowValsOther.get(i).findElements(By.tagName("td"));
			WebElement Income = colValOther.get(4);
			Actions act = new Actions(driver);
			Income.click();
			Thread.sleep(1000);
			Income.click();
			
			act.keyDown(Keys.CONTROL).sendKeys("a").keyUp(Keys.CONTROL).build().perform();
			act.keyDown(Keys.CONTROL).sendKeys("c").keyUp(Keys.CONTROL).build().perform();

			String Incomeamount = (String) Toolkit.getDefaultToolkit().getSystemClipboard()
					.getData(DataFlavor.stringFlavor);

			if (Incomeamount.contains(","))

			{
				strOther = Incomeamount.replace(",", "");
				subValue = Double.parseDouble(strOther.replace("$", ""));
			} else {
				subValue = Double.parseDouble(Incomeamount.replace("$", ""));
			}
			OtherIncomeValue += subValue;

		}

		IncomeValue = ApplicantIncome + OtherIncomeValue;
		System.out.println("ApplicantIncome =$" + IncomeValue);
		// close
		driver.switchTo().defaultContent();
		switchtoIframe1();
		driver.findElement(By.xpath(prop.getProperty("closeIncome"))).click();
		switchtoIframe2();
		switchtoIframe4();
		Thread.sleep(2000);
		driver.findElement(By.xpath(prop.getProperty("PartyDetails"))).click();
		// Liabilities
		WebDriverWait waitLiabilities = new WebDriverWait(driver, 360, 0000);
		waitLiabilities.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("Liabilities"))));
		driver.findElement(By.xpath(prop.getProperty("Liabilities"))).click();
		driver.switchTo().defaultContent();
		switchtoIframe1();
		switchtoIframe3();
		Thread.sleep(8000);
		try
		{
		WebDriverWait waitLiabilitiesTable = new WebDriverWait(driver, 6);
		waitLiabilitiesTable
				.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("LiabilitiesTable"))));
		WebElement LiabilitiesTable = driver.findElement(By.xpath(prop.getProperty("LiabilitiesTable")));
		if (LiabilitiesTable.isDisplayed()) {
			// showrows
			String rows = driver.findElement(By.xpath(prop.getProperty("showrows"))).getText();
			int rowcount=0;

			if (rows.contains("+")) {
				do {
					Loadmore();
					Thread.sleep(5000);
					rows = driver.findElement(By.xpath(prop.getProperty("showrows"))).getText();
					rowcount++;
				} while (rows.contains("+"));
			}

			List<WebElement> rowValsLiability = LiabilitiesTable.findElements(By.tagName("tr"));
			int rowNumLiability = LiabilitiesTable.findElements(By.tagName("tr")).size();

			String strLiability = null;
			TotalLaibility=0;
			for (int i = 0; i < rowNumLiability; i++) {

				double subValue = 0;
				// Get each row's column values by tag name
				List<WebElement> colValsliability = rowValsLiability.get(i).findElements(By.tagName("td"));
				WebElement Liability = colValsliability.get(9);
				String Liabilityamount = Liability.getText();

				if (Liabilityamount.contains(","))

				{
					strLiability = Liabilityamount.replace(",", "");
					subValue = Double.parseDouble(strLiability.replace("$", ""));
				} else {
					subValue = Double.parseDouble(Liabilityamount.replace("$", ""));
				}
				TotalLaibility += subValue;
			}
		
			if(rowcount>4)
			{
				Thread.sleep(6000);
				//click2
				
				driver.findElement(By.xpath(prop.getProperty("next"))).click();
				Thread.sleep(3000);
				List<WebElement> rowValsLiability1 = LiabilitiesTable.findElements(By.tagName("tr"));
				int rowNumLiability1 = LiabilitiesTable.findElements(By.tagName("tr")).size();

				String strLiability1 = null;
				
				for (int i = 0; i < rowNumLiability1; i++) {

					double subValue = 0;
					// Get each row's column values by tag name
					List<WebElement> colValsliability = rowValsLiability1.get(i).findElements(By.tagName("td"));
					WebElement Liability = colValsliability.get(9);
					String Liabilityamount = Liability.getText();

					if (Liabilityamount.contains(","))

					{
						strLiability1 = Liabilityamount.replace(",", "");
						subValue = Double.parseDouble(strLiability1.replace("$", ""));
					} else {
						subValue = Double.parseDouble(Liabilityamount.replace("$", ""));
					}
					TotalLaibility += subValue;
				}
				if(rowcount>9 )
				{
					Thread.sleep(4000);
					//click3
					driver.findElement(By.xpath(prop.getProperty("next"))).click();
					Thread.sleep(3000);
					List<WebElement> rowValsLiability11 = LiabilitiesTable.findElements(By.tagName("tr"));
					int rowNumLiability11 = LiabilitiesTable.findElements(By.tagName("tr")).size();

					String strLiability11 = null;
					
					for (int i = 0; i < rowNumLiability11; i++) {

						double subValue = 0;
						// Get each row's column values by tag name
						List<WebElement> colValsliability = rowValsLiability11.get(i).findElements(By.tagName("td"));
						WebElement Liability = colValsliability.get(9);
						String Liabilityamount = Liability.getText();

						if (Liabilityamount.contains(","))

						{
							strLiability11 = Liabilityamount.replace(",", "");
							subValue = Double.parseDouble(strLiability11.replace("$", ""));
						} else {
							subValue = Double.parseDouble(Liabilityamount.replace("$", ""));
						}
						TotalLaibility += subValue;
					}
					
					if(rowcount>14 )
					{
						Thread.sleep(4000);
						//click4
						driver.findElement(By.xpath(prop.getProperty("next"))).click();
						Thread.sleep(3000);
						List<WebElement> rowValsLiability111 = LiabilitiesTable.findElements(By.tagName("tr"));
						int rowNumLiability111 = LiabilitiesTable.findElements(By.tagName("tr")).size();

						String strLiability111 = null;
						
						for (int i = 0; i < rowNumLiability111; i++) {

							double subValue = 0;
							// Get each row's column values by tag name
							List<WebElement> colValsliability = rowValsLiability111.get(i).findElements(By.tagName("td"));
							WebElement Liability = colValsliability.get(9);
							String Liabilityamount = Liability.getText();

							if (Liabilityamount.contains(","))

							{
								strLiability111 = Liabilityamount.replace(",", "");
								subValue = Double.parseDouble(strLiability111.replace("$", ""));
							} else {
								subValue = Double.parseDouble(Liabilityamount.replace("$", ""));
							}
							TotalLaibility += subValue;
						}
					}
					
				}
				
			}
		}
		}
		catch (TimeoutException e)
        {
            System.out.println("Liability Not Displayed");
           
            
        }

		System.out.println("LiabilityValue =$" + TotalLaibility);
		// close
		driver.switchTo().defaultContent();
		switchtoIframe1();
		driver.findElement(By.xpath(prop.getProperty("closeIncome"))).click();
	}


	public void getAppDetails() throws InterruptedException, IOException {
		// OfferSelection
		WebDriverWait waitAgentDashboard = new WebDriverWait(driver, 360, 0000);
		waitAgentDashboard
				.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("AgentDashboard"))));
		driver.findElement(By.xpath(prop.getProperty("AgentDashboard"))).click();
		WebDriverWait waitSetup2 = new WebDriverWait(driver, 360, 0000);
		waitSetup2.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("OfferSelection"))));
		url = driver.getCurrentUrl();
		// Get MasterID
		String reverseurl = new StringBuffer(url).reverse().toString();
		String[] Spliturl = reverseurl.split("/");
		String view = Spliturl[1];
		masterID = new StringBuffer(view).reverse().toString();
		System.out.println("MasterID =" + masterID);
		driver.findElement(By.xpath(prop.getProperty("OfferSelection"))).click();

		// Switch Iframes
		driver.switchTo().defaultContent();
		switchtoIframe1();
		switchtoIframe3();

		// Offer details

		Thread.sleep(12000);
		WebElement getQla = driver.findElement(By.xpath(prop.getProperty("QualifiedLoanAmount")));
		Actions action = new Actions(driver);

		QualifiedLoanAmount = getQla.getText();
		String qla;
		if (QualifiedLoanAmount.contains(",")) {
			qla = QualifiedLoanAmount.replace(",", "");
		} else {
			qla = QualifiedLoanAmount;
		}

		ExpectedQLA = Double.parseDouble(qla.replace("$", ""));
		System.out.println("ExpectedQLA =$" + ExpectedQLA);
		WebElement getInt = driver.findElement(By.xpath(prop.getProperty("InterestRate")));
		IntRate = getInt.getText();
		((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", getQla);
		Thread.sleep(2000);
		QLAInterestScreen = Screenshot.capture(driver, "CaculateQLA");
		System.out.println("InterestRate =" + IntRate);
		ExpInt = Double.parseDouble(IntRate.replace("%", ""));
		WebDriverWait waitOfferdetails = new WebDriverWait(driver, 360, 0000);
		waitOfferdetails
				.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("Offerdetails"))));

		// Thread.sleep(8000);
		driver.findElement(By.xpath(prop.getProperty("Offerdetails"))).click();

		// Total Income
		WebDriverWait totalincome = new WebDriverWait(driver, 360, 0000);
		totalincome
				.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("TotalIncomeFusion"))));
		totalincomeFusion = driver.findElement(By.xpath(prop.getProperty("TotalIncomeFusion"))).getText();
		String st = totalincomeFusion.replace(",", "");
		TotalIncome = Double.parseDouble(st.replace("$", ""));
		System.out.println("Total Income =$" + TotalIncome);
		IncomeLiabilityScreen = Screenshot.capture(driver, "CalculateIncome");
		// Total Debt
		totaldebtFusion = driver.findElement(By.xpath(prop.getProperty("TotalDebtFusion"))).getText();
		String str1 = totaldebtFusion.replace(",", "");
		TotalDebt = Double.parseDouble(str1.replace("$", ""));
		MaximumLTV = driver.findElement(By.xpath(prop.getProperty("MaximumLTV"))).getText();

		LtvMax = Double.parseDouble(MaximumLTV);

		System.out.println("Total Debt =$" + TotalDebt);
		System.out.println("MaximumLTV =" + LtvMax);

		// ExpectedMaxHA
		List<WebElement> HA = driver.findElements(By.xpath(prop.getProperty("HAFullSpl")));
		WebElement targetHA = HA.get(1);
		((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", targetHA);
		Thread.sleep(2000);
		MaxHASPLFullScreen = Screenshot.capture(driver, "MAxHA");
		String MaxHA = targetHA.getText();
		ExpectedMaxHA = Double.parseDouble(MaxHA.replace(",", ""));
		System.out.println("ExpectedMaxHA =" + ExpectedMaxHA);

		// ReasonCodeSPLFullScreen
		List<WebElement> reasoncode = driver.findElements(By.xpath(prop.getProperty("SPLRiskFactors")));
		length = reasoncode.size();
		if (length == 1) {
			WebElement desicioncode1 = reasoncode.get(0);
			((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", desicioncode1);
			Thread.sleep(2000);
			ReasonCodeSPLFullScreen1 = Screenshot.capture(driver, "ReasonCode1");
		} else if (length == 2) {
			WebElement desicioncode1 = reasoncode.get(0);
			((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", desicioncode1);
			Thread.sleep(2000);
			ReasonCodeSPLFullScreen1 = Screenshot.capture(driver, "ReasonCode1");
			WebElement desicioncode2 = reasoncode.get(1);
			((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", desicioncode2);
			Thread.sleep(2000);
			ReasonCodeSPLFullScreen2 = Screenshot.capture(driver, "ReasonCode2");
		}
		else if (length == 3) {
			WebElement desicioncode1 = reasoncode.get(0);
			((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", desicioncode1);
			Thread.sleep(2000);
			ReasonCodeSPLFullScreen1 = Screenshot.capture(driver, "ReasonCode1");
			WebElement desicioncode2 = reasoncode.get(1);
			((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", desicioncode2);
			Thread.sleep(2000);
			ReasonCodeSPLFullScreen2 = Screenshot.capture(driver, "ReasonCode2");
			WebElement desicioncode3 = reasoncode.get(2);
			((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", desicioncode3);
			Thread.sleep(2000);
			ReasonCodeSPLFullScreen3 = Screenshot.capture(driver, "ReasonCode3");
		}
		
		else {
			WebElement desicioncode1 = reasoncode.get(0);
			((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", desicioncode1);
			Thread.sleep(2000);
			ReasonCodeSPLFullScreen1 = Screenshot.capture(driver, "ReasonCode1");
			WebElement desicioncode2 = reasoncode.get(1);
			((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", desicioncode2);
			Thread.sleep(2000);
			ReasonCodeSPLFullScreen2 = Screenshot.capture(driver, "ReasonCode2");
		}
		// ApplicantEFSCVScore

		ApplicantEFSCVScore = driver.findElement(By.xpath(prop.getProperty("ApplicantEFSCVScore"))).getText();
		String parts = ApplicantEFSCVScore.substring(0, 3);
		cvScore = Integer.parseInt(parts);
		System.out.println("ApplicantEFSCVScore =" + cvScore);

		// CloseOffer
		driver.findElement(By.xpath(prop.getProperty("CloseOffer"))).click();

	}
	public void resubmitForDecisionSpl() throws Exception {

		System.out.println(attemptNo);

		firstPopup();

		Thread.sleep(3000);

		driver.get(url);
		getAddress();
		getPartyDetailsCoApp();
		getCollateral();

		getAppDetailsCoApp();
		
		calculateIncome();
		calculateSPLLiability();
		// mulesoft
		driver.get(prop.getProperty("mulesoft"));
		mulesoft();
		checkContributerincrease();
		splLTV();
		splIncreaseRemInc();
		MaxAfforableIncreaseQLA();
		urbancode();
		splFinalIncreaseQLA();
		maxHA();
		ReasonCode();
		Thread.sleep(4000);
		SecondPopupSpl();
	}

}
