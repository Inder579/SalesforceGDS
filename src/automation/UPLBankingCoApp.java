//ExpressFusionApp 
package automation;

import resources.BrowserDriver;
import resources.ReadExcel;

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

public class UPLBankingCoApp extends BrowserDriver {

	public static int attemptNo = 0;
	public String screenShotPathforInterestRate;
	public WebDriver driver;
	int cvScore, BehaviourScore, cvScoreCoapp;
	public String ActualIncome, appType, loanType, splloanType, Province, ApplicationID, MorgagePayment, NewStrategy,
			CoAppProvince;
	public String TotalIncomeAmount, IntRate, cabKey, qlaStrategy, applicationType, RiskGp, RiskGpCoApp, contributer;
	public double TotalIncome, RemainingIncome, TotalDebt, ExpectedQLA, ExpInt, SPLltv, Maxltv, HomeEquity, PropertyVal,
			ActualQLA1, ActualQLA;

	String lowefs, highefs, Prov, provinceGroup, bkStrategy, ps = "", code = null, propertyType = "",
			propertyLocation = "";
	double lef, hef, calRemIn, QLA, remIn, remInNaPrev, remInNaAfter, LtvMax, ActualMaxHA, ExpectedMaxHA, IncomeValue,
			TotalLaibility;

	int fcol, lcol, col, coldiff, rowNum, RiskGroup, SPLTotalDebt, lastNumRow, bkDecreaseAmount, RandomNumberResponse,
			length;
	String stringSplit[], Strategy, stringSplit2[];

	String lastname, firstname, address, city, dob, clprod, loanpurpose, hearabout, Referral, livingsituation, email,
			lengthofstay, interest, cabKeyApp, cabKeyCoApp;

	String phone, loanamount, landlordname, landlordnumber, Employername, Employerposition, Incomeamt, Incomefreq,
			Employmentstatus, Supervisorname, Supervisornumber, lengthofemployment, previousemployer,
			lengthpreviousemployer, preferedLang;
	String QualifiedLoanAmount, totalincomeFusion, totaldebtFusion, MaximumLTV, ApplicantEFSCVScore, UplStrategyFusion,
			CurrentAddress, postalcode, MortgageBalances, PropertyValue, url, IncomeLiabilityScreen, QLAInterestScreen,
			tsdate, SPLBuydown, masterID, ts0, ts1, ts2, ReasonCodeSPLFullScreen1, ReasonCodeSPLFullScreen2,
			ReasonCodeSPLFullScreen3, MaxHASPLFullScreen, appTypeCoApp;

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
		getAppDetails();

		calculateIncome();
		calculateLiability();
		premulesoft();
		mulesoft();
		Thread.sleep(4000);
		if (Strategy.contains("Banking Strategy")) {
			interestRateBanking();
		} else if (Strategy.contains("Banking Declined")) {
			interestRateBankingDecline();
		}
		Thread.sleep(3000);
		checkContributer();
		remInCalBanking();
		calculateQLABank();
		maxHA();
		ReasonCode();
		Thread.sleep(4000);
		SecondPopup();
	}

	public void loginAsAdmin() throws InterruptedException, IOException, UnsupportedFlavorException {
		driver.get(prop.getProperty("sfUrl"));

		// driver.get("https://goeasy--uatpreview.lightning.force.com/lightning/r/genesis__Applications__c/a5yf0000000HsxbAAC/view");
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

	public void premulesoft() throws InterruptedException {
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

	public void mulesoftremoval() throws Exception {

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
		while (attempts < 3) {
			try {
				Thread.sleep(5000);
				WebDriverWait waitLog = new WebDriverWait(driver, 360, 0000);
				waitLog.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("Logs"))));
				JavascriptExecutor executor = (JavascriptExecutor) driver;
				executor.executeScript("arguments[0].click();", log);

				break;
			} catch (StaleElementReferenceException e) {
				System.out.println("StaleElementReference");
				driver.get(prop.getProperty("mulesoft"));
				mulesoftremoval();

				Thread.sleep(3000);
				remInCalBanking();
				calculateQLABank();
				maxHA();
				ReasonCode();
				Thread.sleep(4000);
				SecondPopup();

			}
			attempts++;
		}

		// act.clickAndHold();
		// act.release().perform();

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

		// ********************

		// ********************
		try {

			Thread.sleep(6000);

			// searchlogs
			WebDriverWait waitsearchlogs = new WebDriverWait(driver, 360, 0000);
			waitsearchlogs
					.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("searchlogs"))));
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
			Thread.sleep(7000);
			List<WebElement> responsefiles = driver.findElements(By.xpath(prop.getProperty("listdebug")));
			WebElement target = responsefiles.get(0);
			Thread.sleep(2000);
			act.moveToElement(target);
			Thread.sleep(2000);
			act.clickAndHold();
			act.release().perform();

		}

		catch (IndexOutOfBoundsException e) {
			Thread.sleep(6000);

			// searchlogs
			WebDriverWait waitsearchlogs = new WebDriverWait(driver, 360, 0000);
			waitsearchlogs
					.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("searchlogs"))));
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

		Thread.sleep(10000);

		act.keyDown(Keys.CONTROL).sendKeys("a").keyUp(Keys.CONTROL).build().perform();
		act.keyDown(Keys.CONTROL).sendKeys("c").keyUp(Keys.CONTROL).build().perform();
		String logs = (String) Toolkit.getDefaultToolkit().getSystemClipboard().getData(DataFlavor.stringFlavor);
		int index1 = logs.indexOf("DE_NewInterestRate_RandomNumber");
		String roar1 = logs.substring(index1 + 35, index1 + 37);

		String randomnumber = null;
		if (roar1.contains(",")) {
			randomnumber = roar1.replace(",", "");
		} else if (roar1.contains(".")) {
			randomnumber = roar1.replace(".", "");

		} else {
			randomnumber = roar1;
		}
		double RandomNum = Double.valueOf(randomnumber);
		RandomNumberResponse = (int) RandomNum;
		System.out.println("Random Number =" + RandomNumberResponse);

		// Risk group
		int index3 = logs.indexOf("DE_App1_Banking_Adjusted_CV_RiskGroup");
		RiskGp = logs.substring(index3 + 41, index3 + 53);
		System.out.println("Risk Group =" + RiskGp);

		// Interest Rate
		int index = logs.indexOf("DE_UPL_NewInterestRate");
		interest = logs.substring(index + 26, index + 31);
		System.out.println("Interest Rate =" + interest);

		// BK Decrease Amount
		if (bkStrategy.contains("BKQLADecrease")) {

			int index2 = logs.indexOf("DE_UPL_App1_BKQualifiedLoanAmount_Decrement");
			String roar2 = logs.substring(index2 + 47, index2 + 52);

			if (roar2.contains("-")) {
				double bkdec = Double.valueOf(roar2);
				bkDecreaseAmount = (int) bkdec;
			}
			System.out.println("BK Decrease Amount : " + bkDecreaseAmount);
		}
		driver.get(url);
		Thread.sleep(5000);
		// driver.close();

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
		while (attempts < 3) {
			try {
				Thread.sleep(5000);
				WebDriverWait waitLog = new WebDriverWait(driver, 360, 0000);
				waitLog.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("Logs"))));
				JavascriptExecutor executor = (JavascriptExecutor) driver;
				executor.executeScript("arguments[0].click();", log);

				break;
			} catch (StaleElementReferenceException e) {
				System.out.println("StaleElementReference");
				driver.get(prop.getProperty("mulesoft"));
				mulesoft();
				Thread.sleep(4000);
				if (Strategy.contains("Banking Strategy")) {
					interestRateBanking();
				} else if (Strategy.contains("Banking Declined")) {
					interestRateBankingDecline();
				}
				Thread.sleep(3000);
				checkContributer();
				remInCalBanking();
				calculateQLABank();
				maxHA();
				ReasonCode();
				Thread.sleep(4000);
				SecondPopup();

			}
			attempts++;
		}

		// act.clickAndHold();
		// act.release().perform();

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

		// ********************

		// ********************
		try {

			Thread.sleep(6000);

			// searchlogs
			WebDriverWait waitsearchlogs = new WebDriverWait(driver, 360, 0000);
			waitsearchlogs
					.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("searchlogs"))));
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
			Thread.sleep(7000);
			List<WebElement> responsefiles = driver.findElements(By.xpath(prop.getProperty("listdebug")));
			WebElement target = responsefiles.get(0);
			Thread.sleep(2000);
			act.moveToElement(target);
			Thread.sleep(2000);
			act.clickAndHold();
			act.release().perform();

		}

		catch (IndexOutOfBoundsException e) {
			Thread.sleep(6000);

			// searchlogs
			WebDriverWait waitsearchlogs = new WebDriverWait(driver, 360, 0000);
			waitsearchlogs
					.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("searchlogs"))));
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

		Thread.sleep(10000);

		act.keyDown(Keys.CONTROL).sendKeys("a").keyUp(Keys.CONTROL).build().perform();
		act.keyDown(Keys.CONTROL).sendKeys("c").keyUp(Keys.CONTROL).build().perform();
		String logs = (String) Toolkit.getDefaultToolkit().getSystemClipboard().getData(DataFlavor.stringFlavor);
		// Random Number
		int index1 = logs.indexOf("DE_NewInterestRate_RandomNumber");
		String roar1 = logs.substring(index1 + 35, index1 + 37);

		String randomnumber = null;
		if (roar1.contains(",")) {
			randomnumber = roar1.replace(",", "");
		} else if (roar1.contains(".")) {
			randomnumber = roar1.replace(".", "");

		} else {
			randomnumber = roar1;
		}
		double RandomNum = Double.valueOf(randomnumber);
		RandomNumberResponse = (int) RandomNum;
		System.out.println("Random Number =" + RandomNumberResponse);

		// Risk group for Applicant and Co-Applicant
		int index3 = logs.indexOf("DE_App1_Banking_Adjusted_CV_RiskGroup");
		RiskGp = logs.substring(index3 + 41, index3 + 53);
		System.out.println("Risk Group Applicant=" + RiskGp);
		int index4 = logs.indexOf("DE_App2_Banking_Adjusted_CV_RiskGroup");
		RiskGpCoApp = logs.substring(index4 + 41, index4 + 53);
		System.out.println("Risk Group Co-Applicant=" + RiskGpCoApp);

		// Interest Rate
		int index = logs.indexOf("DE_UPL_NewInterestRate");
		interest = logs.substring(index + 26, index + 31);
		// System.out.println("Interest Rate =" + interest);

		// Credit contributor from mulesoft
		int Cc = logs.indexOf("DE_UPL_Credit_Contributor");
		contributer = logs.substring(Cc + 29, Cc + 40);

		// BK Decrease Amount
		if (bkStrategy.contains("BKQLADecrease")) {

			int index2 = logs.indexOf("DE_UPL_App1_BKQualifiedLoanAmount_Decrement");
			String roar2 = logs.substring(index2 + 47, index2 + 52);

			if (roar2.contains("-")) {
				double bkdec = Double.valueOf(roar2);
				bkDecreaseAmount = (int) bkdec;
			}
			System.out.println("BK Decrease Amount : " + bkDecreaseAmount);
		}
		driver.get(url);
		Thread.sleep(5000);
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
		Thread.sleep(6000);
		switchtoIframe3();
		Thread.sleep(8000);
		String Event = null;
		try {
			WebDriverWait waitEventHistoryTable = new WebDriverWait(driver, 360, 0000);
			waitEventHistoryTable.until(
					ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("EventHistoryTable"))));

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
		} catch (IndexOutOfBoundsException e) {
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

		// add two hours
		DateTime date = jodatime.minusMinutes(1);

		DateTime dateTime = jodatime.plusMinutes(1); // easier than mucking about with Calendar and constants

		String ant0 = String.valueOf(jodatime);
		ts0 = Character.toString(ant0.charAt(11)) + Character.toString(ant0.charAt(12)) + ":"
				+ Character.toString(ant0.charAt(14)) + Character.toString(ant0.charAt(15));

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
		// appTypeFusion
		appType = driver.findElement(By.xpath(prop.getProperty("appTypeFusion"))).getText();
		System.out.println("Living Status Applicant= " + appType);
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
		url = driver.getCurrentUrl();

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
		double ApplicantIncome = 0, OtherIncomeValue = 0.0;
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
		try {
			WebDriverWait waitLiabilitiesTable = new WebDriverWait(driver, 6);
			waitLiabilitiesTable.until(
					ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("LiabilitiesTable"))));
			WebElement LiabilitiesTable = driver.findElement(By.xpath(prop.getProperty("LiabilitiesTable")));
			if (LiabilitiesTable.isDisplayed()) {
				// showrows
				String rows = driver.findElement(By.xpath(prop.getProperty("showrows"))).getText();
				int rowcount = 0;

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
				TotalLaibility = 0;
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

				if (rowcount > 4) {
					Thread.sleep(6000);
					// click2

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
					if (rowcount > 9) {
						Thread.sleep(4000);
						// click3
						driver.findElement(By.xpath(prop.getProperty("next"))).click();
						Thread.sleep(3000);
						List<WebElement> rowValsLiability11 = LiabilitiesTable.findElements(By.tagName("tr"));
						int rowNumLiability11 = LiabilitiesTable.findElements(By.tagName("tr")).size();

						String strLiability11 = null;

						for (int i = 0; i < rowNumLiability11; i++) {

							double subValue = 0;
							// Get each row's column values by tag name
							List<WebElement> colValsliability = rowValsLiability11.get(i)
									.findElements(By.tagName("td"));
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

						if (rowcount > 14) {
							Thread.sleep(4000);
							// click4
							driver.findElement(By.xpath(prop.getProperty("next"))).click();
							Thread.sleep(3000);
							List<WebElement> rowValsLiability111 = LiabilitiesTable.findElements(By.tagName("tr"));
							int rowNumLiability111 = LiabilitiesTable.findElements(By.tagName("tr")).size();

							String strLiability111 = null;

							for (int i = 0; i < rowNumLiability111; i++) {

								double subValue = 0;
								// Get each row's column values by tag name
								List<WebElement> colValsliability = rowValsLiability111.get(i)
										.findElements(By.tagName("td"));
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
		} catch (TimeoutException e) {
			System.out.println("Liability Not Displayed");

		}

		System.out.println("LiabilityValue =$" + TotalLaibility);
		// close
		driver.switchTo().defaultContent();
		switchtoIframe1();
		driver.findElement(By.xpath(prop.getProperty("closeIncome"))).click();
	}

	public void getPartyDetailsCoApp()
			throws InterruptedException, HeadlessException, UnsupportedFlavorException, IOException {
		// PartyDetails for co-applicant
		double AppIncome, CoAppIncome;
		WebDriverWait waitPartyDetails = new WebDriverWait(driver, 360, 0000);
		waitPartyDetails
				.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("PartyDetails"))));
		driver.findElement(By.xpath(prop.getProperty("PartyDetails"))).click();

		// Employment and Income for co-applicant

		WebDriverWait waitEmploymentandIncome = new WebDriverWait(driver, 360, 0000);
		waitEmploymentandIncome.until(
				ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("EmploymentandIncome"))));
		driver.findElement(By.xpath(prop.getProperty("EmploymentandIncome"))).click();
		driver.switchTo().defaultContent();
		switchtoIframe1();
		switchtoIframe3();
		// Income table for co-applicant
		Thread.sleep(5000);
		WebDriverWait waitMonthlyIncometable = new WebDriverWait(driver, 360, 0000);
		waitMonthlyIncometable
				.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("MonthlyIncometable"))));

		WebElement MonthlyIncometable = driver.findElement(By.xpath(prop.getProperty("MonthlyIncometable")));
		List<WebElement> rowValsIncome = MonthlyIncometable.findElements(By.tagName("tr"));
		int rowNumIncome = MonthlyIncometable.findElements(By.tagName("tr")).size();

		String str = null;
		double ApplicantIncome = 0, OtherIncomeValue = 0.0;
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
		try {

			WebDriverWait waitLiabilitiesTable = new WebDriverWait(driver, 6);
			waitLiabilitiesTable.until(
					ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("LiabilitiesTable"))));
			WebElement LiabilitiesTable = driver.findElement(By.xpath(prop.getProperty("LiabilitiesTable")));
			if (LiabilitiesTable.isDisplayed()) {
				// showrows
				String rows = driver.findElement(By.xpath(prop.getProperty("showrows"))).getText();
				int rowcount = 0;

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
				TotalLaibility = 0;
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

				if (rowcount > 4) {
					Thread.sleep(4000);
					// click2
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
					if (rowcount > 9) {
						Thread.sleep(4000);
						// click3
						driver.findElement(By.xpath(prop.getProperty("next"))).click();
						Thread.sleep(3000);
						List<WebElement> rowValsLiability11 = LiabilitiesTable.findElements(By.tagName("tr"));
						int rowNumLiability11 = LiabilitiesTable.findElements(By.tagName("tr")).size();

						String strLiability11 = null;

						for (int i = 0; i < rowNumLiability11; i++) {

							double subValue = 0;
							// Get each row's column values by tag name
							List<WebElement> colValsliability = rowValsLiability11.get(i)
									.findElements(By.tagName("td"));
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

						if (rowcount > 14) {
							Thread.sleep(4000);
							// click4
							driver.findElement(By.xpath(prop.getProperty("next"))).click();
							Thread.sleep(3000);
							List<WebElement> rowValsLiability111 = LiabilitiesTable.findElements(By.tagName("tr"));
							int rowNumLiability111 = LiabilitiesTable.findElements(By.tagName("tr")).size();

							String strLiability111 = null;

							for (int i = 0; i < rowNumLiability111; i++) {

								double subValue = 0;
								// Get each row's column values by tag name
								List<WebElement> colValsliability = rowValsLiability111.get(i)
										.findElements(By.tagName("td"));
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

		catch (TimeoutException e) {
			System.out.println("Liability Not Displayed");

		}
		// close
		driver.switchTo().defaultContent();
		switchtoIframe1();
		driver.findElement(By.xpath(prop.getProperty("closeIncome"))).click();

		// Co-Borrower
		driver.switchTo().defaultContent();
		switchtoIframe1();
		Thread.sleep(3000);
		switchtoIframe2();
		Thread.sleep(3000);
		WebDriverWait ClickCoAppwait = new WebDriverWait(driver, 360, 0000);
		ClickCoAppwait.until(ExpectedConditions.elementToBeClickable(By.xpath(prop.getProperty("ClickCoApp"))));
		driver.findElement(By.xpath(prop.getProperty("ClickCoApp"))).click();
		Thread.sleep(5000);
		switchtoIframe4();
		Thread.sleep(5000);

		// CurrentAddress
		WebDriverWait CurrentAddresswait = new WebDriverWait(driver, 360, 0000);
		CurrentAddresswait
				.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("CurrentAddress"))));
		CurrentAddress = driver.findElement(By.xpath(prop.getProperty("CurrentAddress"))).getText();
		appTypeCoApp = driver.findElement(By.xpath(prop.getProperty("appTypeFusion"))).getText();
		System.out.println("Living Status Co-Applicant = " + appTypeCoApp);
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
		System.out.println("CoAppProvince =" + CoAppProvince);

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
		double CoApplicantIncome = 0, CoOtherIncomeValue = 0.0;
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
		IncomeValue = AppIncome + CoAppIncome;
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
		waitCoAppLiabilities
				.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("Liabilities"))));
		driver.findElement(By.xpath(prop.getProperty("Liabilities"))).click();
		driver.switchTo().defaultContent();
		switchtoIframe1();
		switchtoIframe3();
		Thread.sleep(8000);

		try {
			CoAppLiabilities();
		} catch (TimeoutException e) {
			System.out.println("Liability Not Displayed");

		}

		System.out.println("Total Liability =$" + TotalLaibility);
		// close
		driver.switchTo().defaultContent();
		switchtoIframe1();
		driver.findElement(By.xpath(prop.getProperty("closeIncome"))).click();
	}

	public void CoAppLiabilities() throws InterruptedException {
		WebDriverWait waitLiabilitiesTable = new WebDriverWait(driver, 6);
		waitLiabilitiesTable
				.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("LiabilitiesTable"))));
		WebElement LiabilitiesTable = driver.findElement(By.xpath(prop.getProperty("LiabilitiesTable")));
		if (LiabilitiesTable.isDisplayed()) {
			// showrows
			String CoApprows = driver.findElement(By.xpath(prop.getProperty("showrows"))).getText();
			int CoApprowcount = 0;

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
			if (CoApprowcount > 4) {
				Thread.sleep(5000);
				// click2
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
				if (CoApprowcount > 9) {
					Thread.sleep(4000);
					// click3
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

					if (CoApprowcount > 14) {
						Thread.sleep(4000);
						// click4
						driver.findElement(By.xpath(prop.getProperty("next"))).click();
						Thread.sleep(3000);
						List<WebElement> rowValsLiability111 = LiabilitiesTable.findElements(By.tagName("tr"));
						int rowNumLiability111 = LiabilitiesTable.findElements(By.tagName("tr")).size();

						String strLiability111 = null;

						for (int i = 0; i < rowNumLiability111; i++) {

							double subValue = 0;
							// Get each row's column values by tag name
							List<WebElement> colValsliability = rowValsLiability111.get(i)
									.findElements(By.tagName("td"));
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

	public void getAppDetails() throws InterruptedException, IOException {
		// OfferSelection
		switchtoIframe2();
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
		System.out.println("Total Debt =$" + TotalDebt);

		// ExpectedMaxHA
		List<WebElement> HA = driver.findElements(By.xpath(prop.getProperty("HAFullSpl")));

		WebElement targetHA = HA.get(3);
		((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", targetHA);
		Thread.sleep(2000);
		MaxHASPLFullScreen = Screenshot.capture(driver, "MAxHA");
		String MaxHA = targetHA.getText();
		ExpectedMaxHA = Double.parseDouble(MaxHA.replace(",", ""));
		System.out.println("ExpectedMaxHA =" + ExpectedMaxHA);

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
		} else if (length == 3) {
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
		// cvScoreCoapp
		String cvscoreco = driver.findElement(By.xpath(prop.getProperty("Co-ApplicantEFSCVScore"))).getText();
		String partsCoapp = cvscoreco.substring(0, 3);
		cvScoreCoapp = Integer.parseInt(partsCoapp);
		System.out.println("CoApplicantEFSCVScore =" + cvScoreCoapp);
		// CabKey for Applicant and Co-Applicant
		cabKeyApp = driver.findElement(By.xpath(prop.getProperty("ApplicantCABKey"))).getText();
		System.out.println("Risk Group Applicant :" + cabKeyApp);
		cabKeyCoApp = driver.findElement(By.xpath(prop.getProperty("Co-ApplicantCABKey"))).getText();
		System.out.println("Risk Group Co-Applicant :" + cabKeyCoApp);
		// cabKey =
		// driver.findElement(By.xpath(prop.getProperty("cabKeyFusionApp"))).getText();
		// System.out.println("Risk Group :" + cabKey);

		// UplStrategyFusion
		Strategy = driver.findElement(By.xpath(prop.getProperty("UplExpressStrategyFusion"))).getText();
		System.out.println("UplStrategyFusion = " + Strategy);
		// BkStrategyFusion
		bkStrategy = driver.findElement(By.xpath(prop.getProperty("BkStrategyFusion"))).getText();
		System.out.println("BkStrategyFusion = " + bkStrategy);

		// CloseOffer
		driver.findElement(By.xpath(prop.getProperty("CloseOffer"))).click();

	}

	public void getStrategy() {

		bkStrategy = driver.findElement(By.xpath(prop.getProperty("bkstrategy"))).getText();
		System.out.println(bkStrategy);

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

	public void calculateLiability() throws InterruptedException, IOException {

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

	public void interestRateBanking() throws DocumentException, InterruptedException, IOException, ParseException {

		if (attemptNo == 0) {
			test = Extent.createTest("Interest Rate Calculation");
		}

		else {

			test = Extent.createTest("Resubmission Result: Attempt#  " + attemptNo + " - InterestRate Calculation");
		}

		int RandomNumber = RandomNumberResponse;
		// Reading values for CV score, Random number and Interest Rate from Excel
		ReadExcel objExcelFile = new ReadExcel();

		// Prepare the path of excel file

		String filePath = System.getProperty("user.dir") + "\\src\\main\\resources\\Excel";

		// Call read file method of the class to read data

		objExcelFile.readExcel(filePath, "UAT-SF-GDScalculation.xlsx", "UPL Banking New Loan Primary");
		ArrayList List1 = objExcelFile.getlist1();
		ArrayList List2 = objExcelFile.getlist2();
		ArrayList List3 = objExcelFile.getlist3();

		int RiskGroup1 = (int) List1.get(0);
		int RiskGroup2 = (int) List1.get(8);
		int RiskGroup3 = (int) List1.get(16);

		// Get Random Numbers from Excel
		// Riskgrp1
		// List1
		int randomCndition1 = (int) List1.get(2);
		int randomCndition2 = (int) List1.get(3);
		int randomCndition3 = (int) List1.get(4);
		int randomCndition4 = (int) List1.get(5);
		// List3
		int ranCndition1 = (int) List3.get(2);
		int ranCndition2 = (int) List3.get(3);
		int ranCndition3 = (int) List3.get(4);
		int ranCndition4 = (int) List3.get(5);
		// Riskgrp2
		// List1
		int randomCndition5 = (int) List1.get(10);
		int randomCndition6 = (int) List1.get(11);
		int randomCndition7 = (int) List1.get(12);
		int randomCndition9 = (int) List1.get(13);
		// List3
		int ranCndition5 = (int) List3.get(10);
		int ranCndition6 = (int) List3.get(11);
		int ranCndition7 = (int) List3.get(12);
		int ranCndition9 = (int) List3.get(13);
		// Riskgrp3
		// List1
		int randomCndition10 = (int) List1.get(18);
		int randomCndition11 = (int) List1.get(19);
		int randomCndition12 = (int) List1.get(20);
		// List3
		int ranCndition10 = (int) List3.get(18);
		int ranCndition11 = (int) List3.get(19);
		int ranCndition12 = (int) List3.get(20);

		// Get Interest Rates from Excel
		// Riskgrp1
		double interestCondition1 = (double) List2.get(0);
		double interestCondition2 = (double) List2.get(1);
		double interestCondition3 = (double) List2.get(2);
		double interestCondition4 = (double) List2.get(3);
		// Riskgrp2
		double interestCondition5 = (double) List2.get(8);
		double interestCondition6 = (double) List2.get(9);
		double interestCondition7 = (double) List2.get(10);
		double interestCondition8 = (double) List2.get(11);
		// Riskgrp3
		double interestCondition9 = (double) List2.get(16);
		double interestCondition10 = (double) List2.get(17);
		double interestCondition11 = (double) List2.get(18);
		// Riskgrp=Any
		double interestCondition12 = (double) List2.get(23);
		double interestCondition13 = (double) List2.get(28);

		// Check Interest Rate
		double intRate = 0;

		if (Province != "Quebec") {

			if (RiskGroup == RiskGroup1)

			{
				if ((RandomNumber >= randomCndition1) && (RandomNumber <= ranCndition1)) {
					intRate = interestCondition1;
				} else if ((RandomNumber >= randomCndition2) && (RandomNumber <= ranCndition2)) {
					intRate = interestCondition2;

				} else if ((RandomNumber >= randomCndition3) && (RandomNumber <= ranCndition3)) {
					intRate = interestCondition3;
					System.out.println(intRate);
				} else if ((RandomNumber >= randomCndition4) && (RandomNumber <= ranCndition4)) {
					intRate = interestCondition4;
				}

			}

			else if (RiskGroup == RiskGroup2)

			{
				if ((RandomNumber >= randomCndition5) && (RandomNumber <= ranCndition5)) {
					intRate = interestCondition5;
				} else if ((RandomNumber >= randomCndition6) && (RandomNumber <= ranCndition6)) {
					intRate = interestCondition6;
				} else if ((RandomNumber >= randomCndition7) && (RandomNumber <= ranCndition7)) {
					intRate = interestCondition7;
				} else if ((RandomNumber >= randomCndition9) && (RandomNumber <= ranCndition9)) {
					intRate = interestCondition8;
				}
			}

			else if (RiskGroup == RiskGroup3)

			{
				if ((RandomNumber >= randomCndition10) && (RandomNumber <= ranCndition10)) {
					intRate = interestCondition9;
				} else if ((RandomNumber >= randomCndition11) && (RandomNumber <= ranCndition11)) {
					intRate = interestCondition10;
				} else if ((RandomNumber >= randomCndition12) && (RandomNumber <= ranCndition12)) {
					intRate = interestCondition11;
				}

			} else {

				intRate = interestCondition12;
			}
		}

		if (Province == "Quebec") {
			intRate = interestCondition13;
		}

		// Displaying Interest Rate result
		System.out.println("Actual Interest rate: " + intRate);
		System.out.println("Expected Interest rate: " + ExpInt);

		if (ExpInt == intRate) {

			test.log(Status.PASS, MarkupHelper.createLabel("InterestRate Percentage Actual value : " + intRate + "%",
					ExtentColor.GREEN));
			test.log(Status.PASS, MarkupHelper.createLabel("InterestRate Percentage Expected value : " + ExpInt + "%",
					ExtentColor.GREEN));

			test.log(Status.PASS, MarkupHelper.createLabel(" Interest Rate Calculation is Matching with GDS Decision",
					ExtentColor.GREEN));
			test.log(Status.PASS, "Snapshot below: " + test.addScreenCaptureFromPath(QLAInterestScreen));
			System.out.println("PASSED in Interest Verification");
		} else {
			System.out.println(intRate + " is the expected value.");

			test.log(Status.FAIL, MarkupHelper.createLabel("InterestRate Percentage Actual value : " + intRate + "%",
					ExtentColor.RED));
			test.log(Status.FAIL, MarkupHelper.createLabel("InterestRate Percentage Expected value : " + ExpInt + "%",
					ExtentColor.RED));

			test.log(Status.FAIL, MarkupHelper.createLabel(" Interest Rate Calculation not Matching with GDS Decision",
					ExtentColor.RED));
			test.log(Status.FAIL, "Snapshot below: " + test.addScreenCaptureFromPath(QLAInterestScreen));
			System.out.println("FAILED in Interest Rate Verification");
		}
		Thread.sleep(3000);
	}
	
	public void interestRateBankingDecline() throws DocumentException, InterruptedException, IOException, ParseException {

		if (attemptNo == 0) {
			test = Extent.createTest("Interest Rate Calculation");
		}

		else {

			test = Extent.createTest("Resubmission Result: Attempt#  " + attemptNo + " - InterestRate Calculation");
		}
		double intRate = 46.96;
		// Displaying Interest Rate result
		System.out.println("Actual Interest rate: " + intRate);
		System.out.println("Expected Interest rate: " + ExpInt);

		if (ExpInt == intRate) {

			test.log(Status.PASS, MarkupHelper.createLabel("InterestRate Percentage Actual value : " + intRate + "%",
					ExtentColor.GREEN));
			test.log(Status.PASS, MarkupHelper.createLabel("InterestRate Percentage Expected value : " + ExpInt + "%",
					ExtentColor.GREEN));

			test.log(Status.PASS, MarkupHelper.createLabel(" Interest Rate Calculation is Matching with GDS Decision",
					ExtentColor.GREEN));
			test.log(Status.PASS, "Snapshot below: " + test.addScreenCaptureFromPath(QLAInterestScreen));
			System.out.println("PASSED in Interest Verification");
		} else {
			System.out.println(intRate + " is the expected value.");

			test.log(Status.FAIL, MarkupHelper.createLabel("InterestRate Percentage Actual value : " + intRate + "%",
					ExtentColor.RED));
			test.log(Status.FAIL, MarkupHelper.createLabel("InterestRate Percentage Expected value : " + ExpInt + "%",
					ExtentColor.RED));

			test.log(Status.FAIL, MarkupHelper.createLabel(" Interest Rate Calculation not Matching with GDS Decision",
					ExtentColor.RED));
			test.log(Status.FAIL, "Snapshot below: " + test.addScreenCaptureFromPath(QLAInterestScreen));
			System.out.println("FAILED in Interest Rate Verification");
		}
		Thread.sleep(3000);
	}
	
	public void checkContributer() {
		// TODO Auto-generated method stub
		if (contributer.contains("Applicant 1")) {
			System.out.println("Credit_Contributer = Applicant 1");

		} else if (contributer.contains("Applicant 2")) {
			System.out.println("Credit_Contributer = Applicant 2");
			cvScore = cvScoreCoapp;
			Prov = CoAppProvince;
			appType = appTypeCoApp;
			RiskGp = RiskGpCoApp;

		} else {
			System.out.println("Credit_Contributer = Shared");
		}
	}

	public void remInCalBanking() throws IOException {
		// TODO Auto-generated method stub
		org.apache.poi.ss.usermodel.Sheet sheet;

		File file = new File(
				System.getProperty("user.dir") + "\\src\\main\\resources\\Excel\\UAT-SF-GDScalculation.xlsx");

		FileInputStream inputStream = new FileInputStream(file);

		Workbook workbook = new XSSFWorkbook(inputStream);

		sheet = workbook.getSheet("UPL Banking RemainingIncome");

		// Making cell values as variable

		String groupA = sheet.getRow(2).getCell(4).getStringCellValue();
		String[] Riskgrp = groupA.split(" or ");
		String Rg1 = Riskgrp[0];
		String Rg2 = Riskgrp[1];

		double value1 = sheet.getRow(2).getCell(6).getNumericCellValue();
		double value2 = sheet.getRow(3).getCell(6).getNumericCellValue();
		double value3 = sheet.getRow(4).getCell(6).getNumericCellValue();
		double value4 = sheet.getRow(5).getCell(6).getNumericCellValue();
		double value5 = sheet.getRow(6).getCell(6).getNumericCellValue();
		double value6 = sheet.getRow(7).getCell(6).getNumericCellValue();

		System.out.println("AppType :" + appType);
		System.out.println("Strategy :" + Strategy);
		System.out.println("Total Income :$" + TotalIncome);
		System.out.println("Total Debt :$" + TotalDebt);

		if (appType.equalsIgnoreCase("Own")) {
			if (RiskGp.contains(Rg1) || RiskGp.contains(Rg2)) {
				RemainingIncome = TotalIncome * value1 - TotalDebt;
			} else if (!(RiskGp.contains(Rg1) || RiskGp.contains(Rg2))) {
				RemainingIncome = TotalIncome * value3 - TotalDebt;
			}
		}

		else if (appType.equalsIgnoreCase("Rent")) {
			if (RiskGp.contains(Rg1) || RiskGp.contains(Rg2)) {
				RemainingIncome = TotalIncome * value2 - TotalDebt;
			} else if (!(RiskGp.contains(Rg1) || RiskGp.contains(Rg2))) {
				RemainingIncome = TotalIncome * value3 - TotalDebt;
			}
		}

		System.out.println("RemainingIncome :$" + RemainingIncome);

	}

	public void calculateQLABank() throws InterruptedException, IOException {
		// TODO Auto-generated method stub
		Thread.sleep(5000);
		if (attemptNo == 0) {
			test = Extent.createTest("QLA Calculation");
		}

		else {

			test = Extent.createTest("Resubmission Result: Attempt#  " + attemptNo + " - QLA Calculation");
		}

		File file = new File(
				System.getProperty("user.dir") + "\\src\\main\\resources\\Excel\\EFS-CV-Grids-FINAL_Risk.xlsx");

		FileInputStream inputStream = new FileInputStream(file);

		Workbook workbook = new XSSFWorkbook(inputStream);
		String IntRate = String.valueOf(ExpInt);

		org.apache.poi.ss.usermodel.Sheet sheet = workbook.getSheet(IntRate);

		Iterator<Row> rows = sheet.iterator();

		Row row = rows.next();

		Iterator<Cell> cell = row.cellIterator();

		Cell value;

		// Setting Province label

		System.out.println("Province is " + Prov);

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
		stringSplit2 = RiskGp.split("p ");
		String rk = stringSplit2[1];
		int rg = Integer.parseInt(rk);
		for (int c = fcol; c < (lcol + 1); c++) {

			int efs = (int) sheet.getRow(2).getCell(c).getNumericCellValue();

			value = sheet.getRow(2).getCell(c);

			if (efs == rg) {

				col = value.getColumnIndex();
				break;

			}

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

			if (bkStrategy.contains("BKQLADecrease")) {
				ActualQLA = ActualQLA + bkDecreaseAmount;

			}

		}

		if (Strategy.contains("Banking Declined")) {
			if (ActualQLA > 4100 && RiskGp.equalsIgnoreCase("Risk Group 3")) {
				ActualQLA = 4100;
			}
			if (ActualQLA > 3100 && RiskGp.equalsIgnoreCase("Risk Group 4")) {
				ActualQLA = 3100;
			}
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
		if (Prov.equalsIgnoreCase("MB")) {
			ActualMaxHA = 0.0;
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

		if (length == 1) {
			test.log(Status.PASS,
					MarkupHelper.createLabel(" Reason Codes/Risk Factors with GDS Decision", ExtentColor.GREEN));
			test.log(Status.PASS, "Snapshot below: " + test.addScreenCaptureFromPath(MaxHASPLFullScreen)
					+ test.addScreenCaptureFromPath(ReasonCodeSPLFullScreen1));
		} else if (length == 2) {
			test.log(Status.PASS,
					MarkupHelper.createLabel(" Reason Codes/Risk Factors with GDS Decision", ExtentColor.GREEN));
			test.log(Status.PASS,
					"Snapshot below: " + test.addScreenCaptureFromPath(MaxHASPLFullScreen)
							+ test.addScreenCaptureFromPath(ReasonCodeSPLFullScreen1)
							+ test.addScreenCaptureFromPath(ReasonCodeSPLFullScreen2));
		} else if (length == 3) {
			test.log(Status.PASS,
					MarkupHelper.createLabel(" Reason Codes/Risk Factors with GDS Decision", ExtentColor.GREEN));
			test.log(Status.PASS,
					"Snapshot below: " + test.addScreenCaptureFromPath(MaxHASPLFullScreen)
							+ test.addScreenCaptureFromPath(ReasonCodeSPLFullScreen1)
							+ test.addScreenCaptureFromPath(ReasonCodeSPLFullScreen2)
							+ test.addScreenCaptureFromPath(ReasonCodeSPLFullScreen3));
		} else {
			test.log(Status.PASS,
					MarkupHelper.createLabel(" Reason Codes/Risk Factors with GDS Decision", ExtentColor.GREEN));
			test.log(Status.PASS,
					"Snapshot below: " + test.addScreenCaptureFromPath(MaxHASPLFullScreen)
							+ test.addScreenCaptureFromPath(ReasonCodeSPLFullScreen1)
							+ test.addScreenCaptureFromPath(ReasonCodeSPLFullScreen2));
		}
		Thread.sleep(3000);
	}

	public void SecondPopup() throws Exception {

		attemptNo++;
		driver.switchTo().defaultContent();

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
			System.out.println("Re-Submission with  Applicant");
			resubmitForDecision();

			break;

		case 2:

			// Function for Removal Co App
			System.out.println("Re-Submission with Removal of Co Applicant");
			removeCoAppUpl();

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

	public void resubmitForDecision() throws Exception {
		System.out.println("Re-Submission Attempt:" + attemptNo);

		firstPopup();
		Thread.sleep(3000);
		driver.get(url);
		getAddress();
		getPartyDetailsCoApp();
		getAppDetails();

		calculateIncome();
		calculateLiability();
		if (bkStrategy.contains("BKQLADecrease")) {
			// mulesoft
			driver.get(prop.getProperty("mulesoft"));
			mulesoft();
		}
		Thread.sleep(3000);
		remInCalBanking();
		calculateQLABank();
		maxHA();
		ReasonCode();
		Thread.sleep(4000);
		SecondPopup();
	}

	public void removeCoAppUpl() throws Exception {

		System.out.println("Re-Submission Attempt:" + attemptNo);

		firstPopup();
		Thread.sleep(3000);
		driver.get(url);
		getAddress();
		getPartyDetails();
		getAppDetails();

		calculateIncome();
		calculateLiability();

		driver.get(prop.getProperty("mulesoft"));
		mulesoftremoval();

		Thread.sleep(3000);
		remInCalBanking();
		calculateQLABank();
		maxHA();
		ReasonCode();
		Thread.sleep(4000);
		SecondPopup();
	}

}