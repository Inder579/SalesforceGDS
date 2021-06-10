package automation;

import java.awt.Color;
import java.awt.Dimension;
import java.awt.Font;
import java.awt.Frame;
import java.awt.GridLayout;
import java.awt.Insets;
import java.awt.LayoutManager;
import java.awt.Toolkit;
import java.awt.Window;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Iterator;
import java.util.List;
import java.util.concurrent.TimeUnit;

import javax.swing.JButton;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JTextPane;
import javax.swing.SwingUtilities;
import javax.swing.UIManager;
import javax.swing.plaf.FontUIResource;

import org.apache.commons.codec.binary.Base64;
import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.dom4j.DocumentException;
import org.dom4j.io.SAXReader;
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
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;
import com.aventstack.extentreports.Status;
import com.aventstack.extentreports.markuputils.ExtentColor;
import com.aventstack.extentreports.markuputils.MarkupHelper;
import com.monitorjbl.xlsx.StreamingReader;

import resources.BrowserDriver;
import resources.ReadExcel;
import resources.Screenshot;

public class GdsTestPrequal extends BrowserDriver {

	public static int attemptNo = 0;
	public String screenShotPathforInterestRate;
	public WebDriver driver;
	int cvScore, BehaviourScore;
	public String ActualIncome, appType, loanType, splloanType, Province, ApplicationID, MorgagePayment;
	public String TotalIncomeAmount, IntRate, cabKey, qlaStrategy, applicationType;
	public double TotalIncome, RemainingIncome, TotalDebt, ExpectedQLA, ExpInt, SPLltv, Maxltv, HomeEquity, PropertyVal;

	String lowefs, highefs, Prov, provinceGroup, bkStrategy, ps = "", code = null, propertyType = "",
			propertyLocation = "";
	double lef, hef, calRemIn, QLA, InterestRate, remIn, remInNaPrev, remInNaAfter, LtvMax,ActualMaxHA,ExpectedMaxHA;
	int fcol, lcol, col, coldiff, rowNum, RiskGroup, SPLTotalDebt, lastNumRow, bkDecreaseAmount;
	String stringSplit[], Strategy;

	@BeforeTest
	public void initialize1() throws IOException {

		driver = browser();

	}

	/*
	 * @Test(priority=2) public void m2() { test =
	 * Extent.createTest(" Calculate Liability"); System.out.println("test");
	 * test.info("test"); }
	 */

	@Test()
	public void m1() throws Exception {

		// Login as Admin
		loginAsAdmin();

		// Login as FSR-Application
		loginAsFSR();
		Thread.sleep(4000);
//		 driver.get("https://c.cs41.visual.force.com/apex/LAMSApplicationView1?srPos=0&srKp=a08&id=a0855000006hhgp&sfdc.override=1");
		// WAIT for User to Submit
		waitForFirstSubmission();
		

		firstPopup();
		Thread.sleep(4000);
		WebDriverWait wait = new WebDriverWait(driver, 360, 0000);
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("getapptype"))));
		appType = driver.findElement(By.xpath(prop.getProperty("getapptype"))).getText();
		applicationType = driver.findElement(By.xpath(prop.getProperty("applicationType"))).getText();

		if (appType.contains("Own")) {
			System.out.println("SPL");
			// SWITCH IFRAME DEFAULT
			driver.switchTo().defaultContent();
			WebDriverWait waitln = new WebDriverWait(driver, 360, 0000);
			waitln.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("splloantype"))));
			splloanType = driver.findElement(By.xpath(prop.getProperty("splloantype"))).getText();

		
			System.out.println("SPL page 1");
			Thread.sleep(2000);
			getAppDetails();

			calculateIncome();
			calculateSPLLiability();

			// SWITCH IFRAME DEFAULT
			driver.switchTo().defaultContent();

			if (loanType.contains("New")) {
				// Go to page 2 (Get time stamp for Decision received)
				System.out.println("SPL New");
				getAppTimestampLogs();
				splinterestRateCalculation();
				splLTV();
				splRemInCal();
				splQLA();
				urbancode();
				splFinalQLA();
				maxHA();
				// Resubmission
				SecondPopupSpl();
			} else if (loanType.contains("Increase")) {
				System.out.println("SPL Increase");

				driver.switchTo().frame(driver.findElement(By.xpath(prop.getProperty("switchIframe"))));
				String behscore = driver.findElement(By.xpath(prop.getProperty("getbehscore"))).getText();

				if (behscore.contains(" ")) {
					System.out.println("BehaviourScore:" + behscore);
				} else {
					BehaviourScore = Integer.parseInt(behscore);
					System.out.println("BehaviourScore:" + BehaviourScore);
				}

				splIncreaseInterestRate();
				splLTV();
				splIncreaseRemInc();
				MaxAfforableIncreaseQLA();
				urbancode();
				splFinalIncreaseQLA();
				maxHA();
				SecondPopupIncreaseSpl();

			}

		} else {
			System.out.println("UPL");
			getUPLdetails();
			if (loanType.contains("New")) {
				// Get Strategy
				getStrategy();
			}

			// Calculate Income & Liabilities
			calculateIncome();
			calculateLiability();

			// SWITCH IFRAME DEFAULT
			driver.switchTo().defaultContent();

			// Go to page 2 (Get time stamp for Decision received)
			getAppTimestampLogs();

			// Interest rate Calculations
			interestRateCalculation();

			Thread.sleep(3000);
			// Logging in as FSR
			if (loanType.contains("New")) {
				// Check Strategy
				if (Strategy.contains("Credit Vision")) {
					remInCal();
					calculateQLA();
				} else if (Strategy.contains("CAB") && IntRate.contains("46.96")) {
					cabQla();
				} else {
					remInCal();
					calculateQLA();
				}
			} else if (loanType.contains("Increase")) {
				remInCal();
				calculateQLA();
			}
			maxHA();
			// Second Pop-up - Resubmission
			SecondPopup();
		}

	}
	static String decodeString(String value)
	 {
	  byte[] decodedString = Base64.decodeBase64(value);
	  return(new String(decodedString));
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

	// resubmitForDecisionIncreaseSpl

	public void resubmitForDecisionIncreaseSpl() throws Exception {
		System.out.println(attemptNo);

		

		firstPopup();
		// SWITCH IFRAME DEFAULT
		driver.switchTo().defaultContent();

		Thread.sleep(2000);

		getAppDetails();

		calculateIncome();
		SPLLiability();

		// SWITCH IFRAME DEFAULT
		driver.switchTo().defaultContent();

		Thread.sleep(3000);
		// QLA calculation
		splLTVResubmit();
		Thread.sleep(3000);
		splIncreaseRemInc();
		MaxAfforableIncreaseQLA();
		splFinalIncreaseQLA();
		maxHA();
		Thread.sleep(3000);
		SecondPopupIncreaseSpl();
	}

	public void SecondPopupIncreaseSpl() throws Exception {
		// TODO Auto-generated method stub
		attemptNo++;
		driver.switchTo().defaultContent();

		WebElement page1 = driver.findElement(By.xpath(prop.getProperty("page1")));
		new Actions(driver).moveToElement(page1).perform();

		JFrame frmOpt = new JFrame(); // We are declaring the frame
		frmOpt.setAlwaysOnTop(true);// This is the line for displaying it above all windows

		Thread.sleep(1000);
		String s = "<html>Press 1 For Re-Submission with Applicant<br>Press 2 For Results<br>";

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
			System.out.println("Re-Submission with  Applicant");
			resubmitForDecisionIncreaseSpl();

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

		Thread.sleep(3000);
		driver.switchTo().defaultContent();
		WebElement Int = driver.findElement(By.xpath(prop.getProperty("re-submit")));
		new Actions(driver).moveToElement(Int).perform();
		Thread.sleep(3000);
		JavascriptExecutor js = (JavascriptExecutor) driver;
		js.executeScript("document.body.style.zoom='60%'");
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

		System.out.println("Int Rate = " + IntRate + " " + "Province = " + Province);
		stringSplit = Province.split(" - ");
		String Prov = stringSplit[0];
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
		if (!(ExpectedQLA == QLA + 100)) {
			// CV Score
			ArrayList<Integer> efsCvScoreList1 = new ArrayList<Integer>();
			ArrayList<Integer> efsCvScoreList2 = new ArrayList<Integer>();
			// Behavior Score
			ArrayList<Integer> efsCvScoreList3 = new ArrayList<Integer>();
			ArrayList<Integer> efsCvScoreList4 = new ArrayList<Integer>();
			// QLA
			ArrayList<Double> efsCvScoreList5 = new ArrayList<Double>();
			ArrayList<Double> efsCvScoreList6 = new ArrayList<Double>();

			for (int x = 10; x < 31; x++) {
				int efsRange1 = (int) sheet.getRow(x).getCell(0).getNumericCellValue();

				efsCvScoreList1.add(efsRange1);
			}

			for (int x = 10; x < 31; x++) {

				int efsRange2 = (int) sheet.getRow(x).getCell(2).getNumericCellValue();
				efsCvScoreList2.add(efsRange2);
			}

			// Behavior Score list
			for (int x = 10; x < 31; x++) {

				int efsRange3 = (int) sheet.getRow(x).getCell(3).getNumericCellValue();
				efsCvScoreList3.add(efsRange3);
			}

			for (int x = 10; x < 31; x++) {

				int efsRange4 = (int) sheet.getRow(x).getCell(5).getNumericCellValue();
				efsCvScoreList4.add(efsRange4);
			}

			// QLA

			for (int x = 10; x < 31; x++) {

				double efsRange6 = sheet.getRow(x).getCell(7).getNumericCellValue();
				efsCvScoreList5.add(efsRange6);
			}
			// Reset Values

			for (int x = 10; x < 31; x++) {

				double efsRange6 = sheet.getRow(x).getCell(8).getNumericCellValue();
				efsCvScoreList6.add(efsRange6);
			}

			// 0
			if (efsCvScoreList2.get(0) <= cvScore && efsCvScoreList4.get(0) <= BehaviourScore
					&& QLA > efsCvScoreList5.get(0)) {
				QLA = efsCvScoreList6.get(0);

			}
			// 1
			else if (efsCvScoreList1.get(1) <= cvScore && cvScore <= efsCvScoreList2.get(1)
					&& efsCvScoreList4.get(1) <= BehaviourScore && QLA > efsCvScoreList5.get(1)) {
				QLA = efsCvScoreList6.get(1);

			}
			// 2
			else if (efsCvScoreList1.get(2) <= cvScore && cvScore <= efsCvScoreList2.get(2)
					&& efsCvScoreList3.get(2) <= BehaviourScore && BehaviourScore <= efsCvScoreList4.get(2)
					&& QLA > efsCvScoreList5.get(2)) {
				QLA = efsCvScoreList6.get(2);

			}
			// 3
			else if (efsCvScoreList1.get(3) <= cvScore && cvScore <= efsCvScoreList2.get(3)
					&& efsCvScoreList3.get(3) <= BehaviourScore && BehaviourScore <= efsCvScoreList4.get(3)
					&& QLA > efsCvScoreList5.get(3)) {
				QLA = efsCvScoreList6.get(3);

			}
			// 4
			else if (efsCvScoreList1.get(4) <= cvScore && cvScore <= efsCvScoreList2.get(4)
					&& efsCvScoreList4.get(4) <= BehaviourScore && QLA > efsCvScoreList5.get(4)) {
				QLA = efsCvScoreList6.get(4);

			}
			// 5
			else if (efsCvScoreList1.get(5) <= cvScore && cvScore <= efsCvScoreList2.get(5)
					&& efsCvScoreList3.get(5) <= BehaviourScore && BehaviourScore <= efsCvScoreList4.get(5)
					&& QLA > efsCvScoreList5.get(5)) {
				QLA = efsCvScoreList6.get(5);

			}

			// 6
			else if (efsCvScoreList1.get(6) <= cvScore && cvScore <= efsCvScoreList2.get(6)
					&& efsCvScoreList3.get(6) <= BehaviourScore && BehaviourScore <= efsCvScoreList4.get(6)
					&& QLA > efsCvScoreList5.get(6)) {
				QLA = efsCvScoreList6.get(6);

			}
			// 7
			else if (efsCvScoreList2.get(7) >= cvScore && efsCvScoreList4.get(7) <= BehaviourScore
					&& QLA > efsCvScoreList5.get(7)) {
				QLA = efsCvScoreList6.get(7);

			}
			// 8
			else if (efsCvScoreList2.get(8) >= cvScore && efsCvScoreList3.get(8) <= BehaviourScore
					&& BehaviourScore <= efsCvScoreList4.get(8) && QLA > efsCvScoreList5.get(8)) {
				QLA = efsCvScoreList6.get(8);

			}
			// 9
			else if (efsCvScoreList2.get(9) >= cvScore && efsCvScoreList3.get(9) <= BehaviourScore
					&& BehaviourScore <= efsCvScoreList4.get(9) && QLA > efsCvScoreList5.get(9)) {
				QLA = efsCvScoreList6.get(9);

			}
			// 10
			else if (efsCvScoreList2.get(10) <= cvScore && efsCvScoreList3.get(10) <= BehaviourScore
					&& BehaviourScore <= efsCvScoreList4.get(10) && QLA > efsCvScoreList5.get(10)) {
				QLA = efsCvScoreList6.get(10);

			}
			// 11
			else if (efsCvScoreList2.get(11) <= cvScore && efsCvScoreList3.get(11) <= BehaviourScore
					&& BehaviourScore <= efsCvScoreList4.get(11) && QLA > efsCvScoreList5.get(11)) {
				QLA = efsCvScoreList6.get(11);

			}
			// 12
			else if (efsCvScoreList2.get(12) < cvScore && efsCvScoreList3.get(12) <= BehaviourScore
					&& BehaviourScore <= efsCvScoreList4.get(12) && QLA > efsCvScoreList5.get(12)) {
				QLA = efsCvScoreList6.get(12);

			}
			// 13
			else if (efsCvScoreList2.get(13) <= cvScore && efsCvScoreList4.get(13) == BehaviourScore
					&& QLA > efsCvScoreList5.get(13)) {
				QLA = efsCvScoreList6.get(13);

			}
			// 14
			else if (efsCvScoreList2.get(14) >= cvScore && efsCvScoreList4.get(14) <= BehaviourScore
					&& QLA > efsCvScoreList5.get(14)) {
				QLA = efsCvScoreList6.get(14);

			}
			// 15
			else if (efsCvScoreList2.get(15) >= cvScore && efsCvScoreList3.get(15) <= BehaviourScore
					&& BehaviourScore <= efsCvScoreList4.get(15) && QLA > efsCvScoreList5.get(15)) {
				QLA = efsCvScoreList6.get(15);

			}
			// 16
			else if (efsCvScoreList2.get(16) >= cvScore && efsCvScoreList3.get(16) <= BehaviourScore
					&& BehaviourScore <= efsCvScoreList4.get(16) && QLA > efsCvScoreList5.get(16)) {
				QLA = efsCvScoreList6.get(16);

			}
			// 17
			else if (efsCvScoreList2.get(17) >= cvScore && efsCvScoreList4.get(17) == BehaviourScore
					&& QLA > efsCvScoreList5.get(17)) {
				QLA = efsCvScoreList6.get(17);

			}
			// 18
			else if (efsCvScoreList2.get(18) >= cvScore && efsCvScoreList4.get(18) >= BehaviourScore
					&& QLA > efsCvScoreList5.get(18)) {
				QLA = efsCvScoreList6.get(18);

			}
			// 19
			else if (efsCvScoreList1.get(19) <= cvScore && cvScore <= efsCvScoreList2.get(19)
					&& efsCvScoreList4.get(19) >= BehaviourScore && QLA > efsCvScoreList5.get(19)) {
				QLA = efsCvScoreList6.get(19);

			}
			// 20
			else if (efsCvScoreList1.get(20) <= cvScore && cvScore <= efsCvScoreList2.get(20)
					&& efsCvScoreList3.get(20) <= BehaviourScore && BehaviourScore <= efsCvScoreList4.get(20)
					&& QLA > efsCvScoreList5.get(20)) {
				QLA = efsCvScoreList6.get(20);

			}

			System.out.println("SPL QLA Final is " + QLA);
		}

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

		for (int x = 5; x < 15; x++) {
			int efsRange1 = (int) sheet.getRow(x).getCell(1).getNumericCellValue();

			efsCvScoreList1.add(efsRange1);
		}

		for (int x = 2; x < 15; x++) {

			int efsRange2 = (int) sheet.getRow(x).getCell(3).getNumericCellValue();
			efsCvScoreList2.add(efsRange2);
		}

		// Behavior Score list
		for (int x = 2; x < 15; x++) {

			int efsRange3 = (int) sheet.getRow(x).getCell(4).getNumericCellValue();
			efsCvScoreList3.add(efsRange3);
		}

		for (int x = 2; x < 15; x++) {

			int efsRange4 = (int) sheet.getRow(x).getCell(6).getNumericCellValue();
			efsCvScoreList4.add(efsRange4);
		}

		// Remaining Income
		ArrayList<Double> efsCvScoreList5 = new ArrayList<Double>();

		for (int x = 2; x < 15; x++) {

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
		else if (efsCvScoreList1.get(0) <= cvScore && cvScore <= efsCvScoreList2.get(3)
				&& efsCvScoreList4.get(3) >= BehaviourScore) {
			val = efsCvScoreList5.get(3);
			RemainingIncome = TotalIncome * val - TotalDebt;
		}
		// 7
		else if (efsCvScoreList1.get(1) <= cvScore && cvScore <= efsCvScoreList2.get(4)
				&& efsCvScoreList4.get(4) < BehaviourScore) {
			val = efsCvScoreList5.get(4);
			RemainingIncome = TotalIncome * val - TotalDebt;
		}
		// 8
		else if (efsCvScoreList1.get(2) <= cvScore && cvScore <= efsCvScoreList2.get(5)
				&& efsCvScoreList3.get(5) <= BehaviourScore && BehaviourScore < efsCvScoreList4.get(5)) {
			val = efsCvScoreList5.get(5);
			RemainingIncome = TotalIncome * val - TotalDebt;
		}
		// 9
		else if (efsCvScoreList1.get(3) <= cvScore && cvScore <= efsCvScoreList2.get(6)
				&& efsCvScoreList4.get(6) == BehaviourScore) {
			val = efsCvScoreList5.get(6);
			RemainingIncome = TotalIncome * val - TotalDebt;
		}
		// 10
		else if (efsCvScoreList1.get(4) <= cvScore && cvScore <= efsCvScoreList2.get(7)
				&& efsCvScoreList4.get(7) < BehaviourScore) {
			val = efsCvScoreList5.get(7);
			RemainingIncome = TotalIncome * val - TotalDebt;
		}
		// 11
		else if (efsCvScoreList1.get(5) <= cvScore && cvScore <= efsCvScoreList2.get(8)
				&& efsCvScoreList3.get(8) <= BehaviourScore && BehaviourScore < efsCvScoreList4.get(8)) {
			val = efsCvScoreList5.get(8);
			RemainingIncome = TotalIncome * val - TotalDebt;
		}
		// 12
		else if (efsCvScoreList1.get(6) <= cvScore && cvScore <= efsCvScoreList2.get(9)
				&& efsCvScoreList4.get(9) == BehaviourScore) {
			val = efsCvScoreList5.get(9);
			RemainingIncome = TotalIncome * val - TotalDebt;
		}
		// 13
		else if (efsCvScoreList1.get(7) <= cvScore && cvScore <= efsCvScoreList2.get(10)
				&& efsCvScoreList4.get(10) < BehaviourScore) {
			val = efsCvScoreList5.get(10);
			RemainingIncome = TotalIncome * val - TotalDebt;
		}
		// 14
		else if (efsCvScoreList1.get(8) <= cvScore && cvScore <= efsCvScoreList2.get(11)
				&& efsCvScoreList4.get(11) >= BehaviourScore) {
			val = efsCvScoreList5.get(11);
			RemainingIncome = TotalIncome * val - TotalDebt;
		}
		// 15
		else if (efsCvScoreList2.get(12) >= cvScore) {
			val = efsCvScoreList5.get(12);
			RemainingIncome = TotalIncome * val - TotalDebt;
		}

		System.out.println("Val =" + val + " " + "RemainingIncome =" + RemainingIncome);

	}

	public void splIncreaseInterestRate() throws IOException, InterruptedException {

		if (attemptNo == 0) {
			test = Extent.createTest("Interest Rate Calculation");
		}

		else {

			test = Extent.createTest("Resubmission Result: Attempt#  " + attemptNo + " - InterestRate Calculation");
		}

		File file = new File(System.getProperty("user.dir")
				+ "\\src\\main\\resources\\Excel\\Credit Decisions Calculations - SPL Increase.xlsx");

		FileInputStream inputStream = new FileInputStream(file);

		Workbook workbook = new XSSFWorkbook(inputStream);

		org.apache.poi.ss.usermodel.Sheet sheet = workbook.getSheet("SPL Interest Rate");

		// Storing random numbers as separate variable

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
		//
		System.out.println(rate);

		Thread.sleep(3000);
		driver.switchTo().defaultContent();

		WebElement Int = driver.findElement(By.xpath(prop.getProperty("otherTable")));
		new Actions(driver).moveToElement(Int).perform();
		Thread.sleep(3000);
		JavascriptExecutor js = (JavascriptExecutor) driver;

		js.executeScript("document.body.style.zoom='70%'");
		Thread.sleep(3000);

		String screenShotPathforInterestRate = Screenshot.capture(driver, "CaculateInterestRate");
		js.executeScript("document.body.style.zoom='100%'");
		// driver.switchTo().frame(driver.findElement(By.xpath(prop.getProperty("switchIframe"))));

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
			test.log(Status.PASS, "Snapshot below: " + test.addScreenCaptureFromPath(screenShotPathforInterestRate));
			System.out.println("PASSED in Interest Verification");
		} else {
			System.out.println(IntRate + " is the Actual value.");

			test.log(Status.FAIL,
					MarkupHelper.createLabel("InterestRate Percentage Actual value : " + rate + "%", ExtentColor.RED));
			test.log(Status.FAIL, MarkupHelper.createLabel("InterestRate Percentage Expected value : " + ExpInt + "%",
					ExtentColor.RED));

			test.log(Status.FAIL, MarkupHelper.createLabel(" Interest Rate Calculation not Matching with GDS Decision",
					ExtentColor.RED));
			test.log(Status.FAIL, "Snapshot below: " + test.addScreenCaptureFromPath(screenShotPathforInterestRate));
			System.out.println("FAILED in Interest Rate Verification");
		}
		Thread.sleep(3000);

	}

	public void urbancode() throws FileNotFoundException {

		String p = driver.findElement(By.xpath(prop.getProperty("getpostalcode"))).getText();

		ps = p.replace("_", " ");
		System.out.println(ps);
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

	public void splFinalQLA() throws InterruptedException, IOException {
		// Home Equity = (Max LTV - SPL LTV)*Property Value/100
		Thread.sleep(5000);

		if (attemptNo == 0) {
			test = Extent.createTest("QLA Calculation");
		}

		else {

			test = Extent.createTest("Resubmission Result: Attempt#  " + attemptNo + " - QLA Calculation");
		}
		// Maxltv Calculation

		ArrayList rangeList;
		ArrayList<Double> efsCvScoreList = new ArrayList<Double>();
		ArrayList<String> RiskgroupList = new ArrayList<String>();
		String range, riskGroup = " ", propertyRange;
		Double drange, propRange, propRange1, propRange2;
		int thecol = 0, therow = 0, tabrow = 0;
		String stringSplitter[];

		// Sheet initalization.This needs to be done once in Class level, so that, we
		// dont have to initialize this in each function
		org.apache.poi.ss.usermodel.Sheet sheet1;

		File file = new File(System.getProperty("user.dir")
				+ "\\src\\main\\resources\\Excel\\Credit Decision Calculations - SPL New.xlsx");

		FileInputStream inputStream = new FileInputStream(file);

		Workbook workbook = new XSSFWorkbook(inputStream);

		sheet1 = workbook.getSheet("Max LTV");

		// Reading cv score from B column

		for (int i = 0; i < 7; i++) {
			range = sheet1.getRow(i).getCell(1).getStringCellValue();

			stringSplitter = range.split("&");

			if (i == 0) // This is for first cell alone. To remove the text content in the first cell
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
		// RiskgroupList
		for (int i = 0; i < 7; i++) {
			String rangerisk = sheet1.getRow(i).getCell(2).getStringCellValue();
			RiskgroupList.add(rangerisk);
		}

		// Assigning the risk group based on applicant's efs score
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

		System.out.println(riskGroup);

		// Iterating through the tables to identify the right Risk Group
		for(int j=9;j<86;j++) //9 is the first row where table starts & 86 is the last row in the table
		{

		if(sheet1.getRow(j).getCell(1).getStringCellValue().contains(riskGroup))
		{
		        tabrow=j;
		        for(int k=1;k<19;k++)
		        {
		                      propertyRange=sheet1.getRow(j+1).getCell(k).getStringCellValue();
		                                                                 
		                      if(k==1)
		                      {
		                                     propertyRange = propertyRange.replace("Property Value ","");
		                                     propertyRange = propertyRange.replace("<$","");
		                                     propertyRange = propertyRange.replace("K","000");
		                                     propRange = Double.parseDouble(propertyRange);
		                                     
		                                     if(PropertyVal<propRange)
		                                     {
		                                                   thecol=k;
		                                                   break;
		                                     }
		                      }
		                      
		                      
		                      if(k==6 && riskGroup.equalsIgnoreCase("Risk Group 1")) //Since, Risk Group 1 has only two property type tables
		                      {

		                                     propertyRange = propertyRange.replace("Property Value ","");
		                                     propertyRange = propertyRange.replace(">= $","");
		                                     propertyRange = propertyRange.replace("K","000");
		                                     propRange = Double.parseDouble(propertyRange);
		                                     
		                                     if(PropertyVal>=propRange)
		                                     {
		                                                   thecol=k;
		                                                   break;
		                                     }
		                                     
		                      }
		                      
		                      if(k==6 && !riskGroup.equalsIgnoreCase("Risk Group 1"))   
		                      {             
		                                     propertyRange = propertyRange.replace("Property Value and ","");
		                                     propertyRange = propertyRange.replace("<$","");
		                                     propertyRange = propertyRange.replace(">= $","");
		                                     propertyRange = propertyRange.replace("K","000");
		                      stringSplitter = propertyRange.split("and");
		                      String range1 = stringSplitter[0];
		                      String range2 = stringSplitter[1];
		                      
		                                     propRange1 = Double.parseDouble(range1);
		                                     propRange2 = Double.parseDouble(range2);

		                                    if((PropertyVal>=propRange1)&&(PropertyVal<propRange2))                                                                     
		                                     {
		                                                   thecol=k;
		                                                   break;
		                                     }
		                                     
		                      }
		                      
		                      
		                      if(k==11)
		                      {
		                                     
		                                     propertyRange = propertyRange.replace("Property Value ","");
		                                     propertyRange = propertyRange.replace("<$","");
		                                     propertyRange = propertyRange.replace(">=$","");
		                                     propertyRange = propertyRange.replace("K","000");
		                      stringSplitter = propertyRange.split("and");
		                      String range1 = stringSplitter[0];
		                      String range2 = stringSplitter[1];
		                      
		                                     propRange1 = Double.parseDouble(range1);
		                                     propRange2 = Double.parseDouble(range2);
		                                     
		                                    if((PropertyVal>=propRange1)&&(PropertyVal<propRange2))                                                                     
		                                     {
		                                                   thecol=k;
		                                                   break;
		                                     }
		                      }
		                      
		                      if(k==16)
		                      {
		                                     propertyRange = propertyRange.replace("Property Value ","");
		                                     propertyRange = propertyRange.replace(">=$","");
		                                     propertyRange = propertyRange.replace("K","000");
		                                     propRange = Double.parseDouble(propertyRange);
		                                     
		                                     if(PropertyVal>=propRange)
		                                     {
		                                                   thecol=k;
		                                                   break;
		                                     }
		                      }
		                      
		                      k+=4;

		        }
		        
		}
		j+=12;
		}

		int r = tabrow + 4;
		try {
			while (sheet1.getRow(r).getCell(thecol).getCellTypeEnum() == CellType.STRING) {

				if (sheet1.getRow(r).getCell(thecol).getStringCellValue().toLowerCase()
						.contains(propertyType.toLowerCase())) {

					therow = r;
					break;
				}
				r++;
			}
		} catch (Exception e) {

		}

		if (code.equalsIgnoreCase("Urban")) {
			Maxltv = sheet1.getRow(therow).getCell(thecol + 1).getNumericCellValue();
		}

		if (code.equalsIgnoreCase("Rural")) {
			Maxltv = sheet1.getRow(therow).getCell(thecol + 2).getNumericCellValue();
		}
		if (code.equalsIgnoreCase("Remote")) {

			Maxltv = sheet1.getRow(therow).getCell(thecol + 3).getNumericCellValue();
		}

		System.out.println("Max LTV is " + Maxltv);

		// Home Equity Calculation
		HomeEquity = (LtvMax - SPLltv) * PropertyVal / 100;
		System.out.println("Home Equity =" + HomeEquity);
		double ActualQLA = 0;

		if (QLA==0.0) {
			ActualQLA = QLA;
		} 
		else if(HomeEquity==0.0)
		{
			ActualQLA=HomeEquity;
		}
		else if (HomeEquity > QLA) {
			ActualQLA = QLA + 100;
		} else if (HomeEquity < QLA) {
			ActualQLA = HomeEquity + 100;
		}

		System.out.println("Actual QLA :$" + ActualQLA);
		System.out.println("Expected QLA :$" + ExpectedQLA);

		Thread.sleep(3000);
		driver.switchTo().defaultContent();
		WebElement Int = driver.findElement(By.xpath(prop.getProperty("re-submit")));
		new Actions(driver).moveToElement(Int).perform();
		Thread.sleep(3000);
		JavascriptExecutor js = (JavascriptExecutor) driver;
		js.executeScript("document.body.style.zoom='60%'");
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

	public void splRemInCal() throws InterruptedException, IOException {

		org.apache.poi.ss.usermodel.Sheet sheet;

		File file = new File(System.getProperty("user.dir")
				+ "\\src\\main\\resources\\Excel\\Credit Decision Calculations - SPL New.xlsx");

		FileInputStream inputStream = new FileInputStream(file);

		Workbook workbook = new XSSFWorkbook(inputStream);

		sheet = workbook.getSheet("SPL Remaining Income");

		// Making cell values as variable

		int cvSco1 = (int) sheet.getRow(3).getCell(1).getNumericCellValue();
		int cvSco2 = (int) sheet.getRow(4).getCell(1).getNumericCellValue();
		int cvSco3 = (int) sheet.getRow(5).getCell(1).getNumericCellValue();

		double value1 = sheet.getRow(3).getCell(3).getNumericCellValue();
		double value2 = sheet.getRow(4).getCell(3).getNumericCellValue();
		double value3 = sheet.getRow(5).getCell(3).getNumericCellValue();

		if (cvScore <= cvSco1) {
			RemainingIncome = TotalIncome * value1 - TotalDebt;
		}

		else if ((cvScore > cvSco1) && (cvScore <= cvSco2)) {
			RemainingIncome = TotalIncome * value2 - TotalDebt;
		}

		else if (cvScore >= cvSco3) {
			RemainingIncome = TotalIncome * value3 - TotalDebt;
		}

		System.out.println("RemainingIncome = " + RemainingIncome);

	}

	public void splQLA() throws InterruptedException, IOException {

		File file = new File(System.getProperty("user.dir")
				+ "\\src\\main\\resources\\Excel\\Credit Decision Calculations - SPL New.xlsx");

		FileInputStream inputStream = new FileInputStream(file);

		Workbook workbook = new XSSFWorkbook(inputStream);
		// String IntRate = String.valueOf(ExpInt);

		org.apache.poi.ss.usermodel.Sheet sheet = workbook.getSheet("SPL QLA Interest");

		Row r = sheet.getRow(1);
		int lastCol = r.getLastCellNum(); // Gets last column index
		int lastrow = sheet.getLastRowNum(); // Gets last row num

		System.out.println("Int Rate = " + IntRate + " " + "Province = " + Province);
		stringSplit = Province.split(" - ");
		String Prov = stringSplit[0];
		System.out.println(cvScore);
		int fcol = 0;
		// Iterating the row which Interest value for identifying the right table
		for (int i = 7; i <= lastCol; i++) {
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

		// CV Score
		int CVval1 = (int) sheet.getRow(8).getCell(0).getNumericCellValue();
		int CVval2 = (int) sheet.getRow(8).getCell(2).getNumericCellValue();
		int CVval3 = (int) sheet.getRow(9).getCell(0).getNumericCellValue();
		int CVval4 = (int) sheet.getRow(9).getCell(2).getNumericCellValue();
		int CVval5 = (int) sheet.getRow(10).getCell(2).getNumericCellValue();
		int CVval6 = (int) sheet.getRow(11).getCell(2).getNumericCellValue();

		// Max QLA
		int QLAval1 = (int) sheet.getRow(8).getCell(4).getNumericCellValue();
		int QLAval2 = (int) sheet.getRow(9).getCell(4).getNumericCellValue();
		int QLAval3 = (int) sheet.getRow(10).getCell(4).getNumericCellValue();
		int QLAval4 = (int) sheet.getRow(11).getCell(4).getNumericCellValue();

		// Reset Values
		int Resetval1 = (int) sheet.getRow(8).getCell(5).getNumericCellValue();
		int Resetval2 = (int) sheet.getRow(9).getCell(5).getNumericCellValue();
		int Resetval3 = (int) sheet.getRow(10).getCell(5).getNumericCellValue();
		int Resetval4 = (int) sheet.getRow(11).getCell(5).getNumericCellValue();

		// Max QLA Conditions
		if ((QLA > QLAval1) && (cvScore > CVval1) && (cvScore < CVval2)) {
			QLA = Resetval1;
		}
		if ((QLA > QLAval2) && (cvScore > CVval3) && (cvScore < CVval4)) {
			QLA = Resetval2;
		}
		if ((QLA > QLAval3) && (cvScore < CVval5)) {
			QLA = Resetval3;
		}
		if ((QLA > QLAval4) && (cvScore < CVval6)) {
			QLA = Resetval4;

		}

		System.out.println("SPL QLA from Excel is " + QLA);
	}

	public void splLTV() {

		// SWITCH IFRAME DEFAULT
		driver.switchTo().defaultContent();

		// SPL LTV Calculation - 1st Submission
		// SPL LTV= (Total Amount of Applicant Mortgage Balances Outstanding + Total
		// Credit Limits of Revolving Trades of Applicant)*100/Total Value of Property
		String MortgageBalances = driver.findElement(By.xpath(prop.getProperty("getMorgageBalance"))).getText();
		String PropertyValue = driver.findElement(By.xpath(prop.getProperty("getPropertyValue"))).getText();
		String str = MortgageBalances.replace(",", "");
		double MortgageBal = Double.parseDouble(str.replace("$", ""));
		System.out.println("Mortgage Balance = " + MortgageBal);
		String str1 = PropertyValue.replace(",", "");
		PropertyVal = Double.parseDouble(str1.replace("$", ""));
		System.out.println("Property Value = " + PropertyVal);
		SPLltv = MortgageBal * 100 / PropertyVal;
		System.out.println("SPL LTV =" + SPLltv);

	}

	public void SecondPopupSpl() throws Exception {

		attemptNo++;
		driver.switchTo().defaultContent();

		JFrame frmOpt = new JFrame(); // We are declaring the frame
		frmOpt.setAlwaysOnTop(true);// This is the line for displaying it above all windows

		Thread.sleep(1000);
		String s = "<html>Press 1 For Re-Submission with Applicant<br>Press 2 For Results<br>";

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
			System.out.println("Re-Submission with  Applicant");
			resubmitForDecisionSpl();

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

	public void resubmitForDecisionSpl() throws Exception {

		System.out.println(attemptNo);

	
		firstPopup();
		// SWITCH IFRAME DEFAULT
		driver.switchTo().defaultContent();

		Thread.sleep(2000);

		getAppDetails();

		calculateIncome();
		SPLLiability();

		// SWITCH IFRAME DEFAULT
		driver.switchTo().defaultContent();

		Thread.sleep(3000);
		// QLA calculation
		splLTVResubmit();
		Thread.sleep(3000);
		splRemInCal();
		Thread.sleep(3000);
		splQLA();
		splFinalQLA();
		maxHA();
		Thread.sleep(3000);
		SecondPopupSpl();
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

	public void SPLLiability() throws InterruptedException, IOException {
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
			String str11 = RentAmount.replace(",", "");
			double subTotal = Double.parseDouble(str11.replace("$", ""));

			System.out.println(subTotal);

			RentValue += subTotal;

			System.out.println("---------------");
		}

		System.out.println(RentValue);
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
			String str2 = RentAmount.replace(",", "");
			double subTotal = Double.parseDouble(str2.replace("$", ""));

			System.out.println(subTotal);

			OtherValue += subTotal;

			System.out.println("---------------");
		}

		System.out.println(OtherValue);

		double TotalLaibility = LiabilitiesValue + MortgagesValue + OtherValue + RentValue;
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
		String str1 = TotalDebtAmount.replace(",", "");
		TotalDebt = Double.parseDouble(str1.replace("$", ""));

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

	public void splLTVResubmit() {
		String Bal;
		double Balances, CreditLimit;
		// SWITCH IFRAME DEFAULT
		driver.switchTo().defaultContent();
		// SPL LTV Calculation - Re Submission
		String PropertyValue = driver.findElement(By.xpath(prop.getProperty("getPropertyValue"))).getText();
		String str1 = PropertyValue.replace(",", "");
		PropertyVal = Double.parseDouble(str1.replace("$", ""));

		driver.manage().timeouts().implicitlyWait(0, TimeUnit.SECONDS);
		int colNumMortgages = driver.findElements(By.xpath(prop.getProperty("colmortgages"))).size();

		WebElement Mortgagestable = driver.findElement(By.xpath(prop.getProperty("mortgagesTable")));

		List<WebElement> rowMortgages = Mortgagestable.findElements(By.xpath(prop.getProperty("rowmortgages")));

		int rowNumMortgages = Mortgagestable.findElements(By.xpath(prop.getProperty("rowmortgages"))).size();

		for (int i = 0; i < rowNumMortgages; i++) {
			// Get each row's column values by tag name
			List<WebElement> colValsMortgages = rowMortgages.get(i).findElements(By.tagName("td"));
			String AccountType = colValsMortgages.get(5).getText();

			if (AccountType.contains("Installment Account")) {
				Bal = colValsMortgages.get(4).getText();

				System.out.println("Balance :" + Bal);
				String str2 = Bal.replace(",", "");
				Balances = Double.parseDouble(str2.replace("$", ""));

				SPLltv = Balances * 100 / PropertyVal;
				System.out.println("SPL LTV =" + SPLltv);
			} else {
				Bal = colValsMortgages.get(6).getText();

				System.out.println("Credit :" + Bal);
				String str2 = Bal.replace(",", "");
				CreditLimit = Double.parseDouble(str2.replace("$", ""));

				SPLltv = CreditLimit * 100 / PropertyVal;
				System.out.println("SPL LTV =" + SPLltv);
			}

			System.out.println("---------------");
		}
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

	public void getUPLdetails() throws InterruptedException {

		WebDriverWait waitLoad = new WebDriverWait(driver, 360, 0000);
		waitLoad.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("getprovince"))));
		Province = driver.findElement(By.xpath(prop.getProperty("getprovince"))).getText();
		ApplicationID = driver.findElement(By.xpath(prop.getProperty("getappid"))).getText();
		loanType = driver.findElement(By.xpath(prop.getProperty("getloantype"))).getText();
		applicationType = driver.findElement(By.xpath(prop.getProperty("applicationType"))).getText();
		appType = driver.findElement(By.xpath(prop.getProperty("getapptype"))).getText();
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
		if (loanType.contains("New")) {
			cabKey = driver.findElement(By.xpath("//th[contains(text(),'CAB Key')]/following-sibling::td[1]/span"))
					.getText();
			System.out.println("Risk Group :" + cabKey);
		}
		riskFactor();
		decisionCode();

		System.out.println("ApplicationID: " + ApplicationID);
		System.out.println("Province: " + Province);

		System.out.println("Interest Rate :" + ExpInt);
		System.out.println("CV Score: " + cvScore);
		System.out.println("QLA: " + ExpectedQLA);
		System.out.println("Max H&A: " + ExpectedMaxHA);
		System.out.println("Strategy :" + Strategy);
	}

	public void getAppDetails() throws InterruptedException, IOException {

		WebDriverWait waitLoad = new WebDriverWait(driver, 360, 0000);
		waitLoad.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("getprovince"))));
		Province = driver.findElement(By.xpath(prop.getProperty("getprovince"))).getText();
		ApplicationID = driver.findElement(By.xpath(prop.getProperty("getappid"))).getText();
		propertyType = driver.findElement(By.xpath(prop.getProperty("getPropertyType"))).getText();
		System.out.println(propertyType);
		propertyLocation = driver.findElement(By.xpath(prop.getProperty("getUrbanCode"))).getText();
		System.out.println("App Type :" + appType);
		loanType = driver.findElement(By.xpath(prop.getProperty("getloantype"))).getText();
		System.out.println("Loan Type :" + loanType);
		driver.switchTo().frame(driver.findElement(By.xpath(prop.getProperty("switchIframe"))));
		Thread.sleep(3000);
		IntRate = driver.findElement(By.xpath(prop.getProperty("get%"))).getText();
		ExpInt = Double.parseDouble(IntRate.replace("%", ""));
		String maxltv = driver.findElement(By.xpath(prop.getProperty("Maxltv"))).getText();
		String s = maxltv.replace("%", "");
		LtvMax = Double.parseDouble(s);
		String intqla = driver.findElement(By.xpath(prop.getProperty("getQla"))).getText();
		String qla = intqla.replace(",", "");
		ExpectedQLA = Double.parseDouble(qla.replace("$", ""));
		String intHA=driver.findElement(By.xpath(prop.getProperty("getmaxHA"))).getText();
		String HA = intHA.replace(",", "");

		ExpectedMaxHA = Double.parseDouble(HA.replace("$", ""));
		String cvscore = driver.findElement(By.xpath(prop.getProperty("getcvscore"))).getText();
		cvScore = Integer.parseInt(cvscore);
		Strategy = driver.findElement(By.xpath(prop.getProperty("adstrategy"))).getText();
		qlaStrategy = driver.findElement(By.xpath(prop.getProperty("qlastrategy"))).getText();
		riskFactor();
		decisionCode();

		System.out.println("ApplicationID: " + ApplicationID);
		System.out.println("Province: " + Province);

		System.out.println("Interest Rate :" + ExpInt);
		System.out.println("CV Score: " + cvScore);
		System.out.println("QLA: " + ExpectedQLA);
		System.out.println("Strategy :" + Strategy);
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
			double subValue = 0;
			// Get each row's column values by tag name
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

	public void calculateLiability() throws InterruptedException, IOException {
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

	public void calculateSPLLiability() throws InterruptedException, IOException {
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
		MorgagePayment = driver.findElement(By.xpath(prop.getProperty("getMorgagePayment"))).getText();
		System.out.println("Morgage Payment :" + MorgagePayment);
		driver.manage().timeouts().implicitlyWait(0, TimeUnit.SECONDS);
		WebElement Liabilitiestable = driver.findElement(By.xpath(prop.getProperty("liabilitiesTable")));
		new Actions(driver).moveToElement(Liabilitiestable).perform();
		Thread.sleep(2000);
		List<WebElement> rowValsLiabilities = Liabilitiestable.findElements(By.tagName("tr"));
		int rowNumLiabilities = Liabilitiestable.findElements(By.tagName("tr")).size();

		int colNumLiabilities = driver.findElements(By.xpath(prop.getProperty("colliabilities"))).size();

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

		String str = MorgagePayment.replace(",", "");
		double MorgateTotal = Double.parseDouble(str.replace("$", ""));
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
			String str11 = RentAmount.replace(",", "");
			double subTotal = Double.parseDouble(str11.replace("$", ""));

			System.out.println(subTotal);

			RentValue += subTotal;

			System.out.println("---------------");
		}

		System.out.println(RentValue);

		// Other Table
		System.out.println("Other Table");
		driver.manage().timeouts().implicitlyWait(0, TimeUnit.SECONDS);
		WebElement Othertable = driver.findElement(By.xpath(prop.getProperty("otherTable")));
		new Actions(driver).moveToElement(Othertable).perform();

		List<WebElement> rowOther = Othertable.findElements(By.xpath(prop.getProperty("rowother")));

		int rowNumOther = Othertable.findElements(By.xpath(prop.getProperty("rowother"))).size();

		int colNumOther = driver.findElements(By.xpath(prop.getProperty("colother"))).size();

		double OtherValue = 0;
		for (int i = 0; i < rowNumOther; i++) {
			// Get each row's column values by tag name
			List<WebElement> colValsOther = rowOther.get(i).findElements(By.tagName("td"));
			String RentAmount = colValsOther.get(3).getText();
			System.out.println(RentAmount);
			String str2 = RentAmount.replace(",", "");
			double subTotal = Double.parseDouble(str2.replace("$", ""));

			System.out.println(subTotal);

			OtherValue += subTotal;

			System.out.println("---------------");
		}

		System.out.println(OtherValue);

		double TotalLaibility = LiabilitiesValue + MorgateTotal + OtherValue + RentValue;
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
		String str1 = TotalDebtAmount.replace(",", "");
		TotalDebt = Double.parseDouble(str1.replace("$", ""));

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

	public void getAppTimestampLogs() throws InterruptedException, DocumentException {

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

	public void splinterestRateCalculation() throws IOException, DocumentException, InterruptedException, ParseException {
		if (attemptNo == 0) {
			test = Extent.createTest("Interest Rate Calculation");
		}

		else {

			test = Extent.createTest("Resubmission Result: Attempt#  " + attemptNo + " - InterestRate Calculation");
		}
		// Read Random number from XML
		Thread.sleep(7000);
		
		// TODO Auto-generated method stub
		File file1 = new File(System.getProperty("user.dir")+"\\src\\main\\resources\\logs\\Response.xml");
        File newFile = new File(System.getProperty("user.dir")+"\\src\\main\\resources\\logs\\Response.json");
        if(file1.renameTo(newFile)){
            System.out.println("File rename success");;
        }else{
            System.out.println("File rename failed");
        }
        JSONParser parser = new JSONParser();
        Object obj = parser.parse(new FileReader(System.getProperty("user.dir")+"\\src\\main\\resources\\logs\\Response.json"));
        JSONObject jsonObject =  (JSONObject) obj;
        JSONArray cars = (JSONArray) jsonObject.get("Response");
        String txt =cars.toString();
        int RandomNumber;
        double InterestRate = 0;
        if(cvScore<=682)
        {
        int index1 = txt.indexOf("<RandomNumber_Internal_SPLInterestRate>");
        String roar1 = txt.substring(index1+39, index1+41);
        double RandomNum = Double.valueOf(roar1);
		 RandomNumber = (int) RandomNum;
        }
        else
        {
        	RandomNumber = 0;
        	InterestRate=19.99;
        }
		
		
		File inputFile = new File(System.getProperty("user.dir") + "\\src\\main\\resources\\logs\\Response.json");
/*		SAXReader saxReader = new SAXReader();
		org.dom4j.Document document = saxReader.read(inputFile);

		String randomNumber = document.selectSingleNode("//DecisionEngine/RandomNumber_Internal_SPLInterestRate")
				.getText();
*/
		
		System.out.println("RandomNumber: " + RandomNumber);
		

		// Delete Response File
		if (inputFile.exists()) {
			Thread.sleep(3000);
			inputFile.delete();
			Thread.sleep(3000);
			System.out.println("Response File deleted");
		}

		double randomNumOne, randomNumTwo, randomNumThree, randomNumFour;
		
		String randomNumRange, efsCvScoreRange;

		File file = new File(System.getProperty("user.dir")
				+ "\\src\\main\\resources\\Excel\\Credit Decision Calculations - SPL New.xlsx");

		FileInputStream inputStream = new FileInputStream(file);

		Workbook workbook = new XSSFWorkbook(inputStream);

		org.apache.poi.ss.usermodel.Sheet sheet = workbook.getSheet("SPL Interest Rate");

		// Storing random numbers as separate variable

		randomNumRange = sheet.getRow(2).getCell(1).getStringCellValue();

		String[] stringSplit = randomNumRange.split("-");

		randomNumOne = Double.parseDouble(stringSplit[0]); // 1
		randomNumTwo = Double.parseDouble(stringSplit[1]); // 50

		// System.out.println(randomNumOne);

		randomNumRange = sheet.getRow(3).getCell(1).getStringCellValue();

		stringSplit = randomNumRange.split("-");

		randomNumThree = Double.parseDouble(stringSplit[0]); // 51
		randomNumFour = Double.parseDouble(stringSplit[1]); // 100

		// System.out.println(randomNumFour);

		ArrayList<Double> efsCvScoreList = new ArrayList<Double>();

		for (int x = 2; x <= 10; x++) {
			efsCvScoreRange = sheet.getRow(x).getCell(0).getStringCellValue();
			// System.out.println(efsCvScoreRange);
			if (efsCvScoreRange.contains("-")) {
				stringSplit = efsCvScoreRange.split("-");
				efsCvScoreList.add(Double.parseDouble(stringSplit[0]));
				efsCvScoreList.add(Double.parseDouble(stringSplit[1]));
			}
			if (efsCvScoreRange.contains("=")) {

				efsCvScoreRange = efsCvScoreRange.replace("<=", "");
				efsCvScoreList.add(Double.parseDouble(efsCvScoreRange));

			}
			x++;
		}

		if ((randomNumOne <= RandomNumber) && (RandomNumber <= randomNumTwo)) // 1 and 50
		{
			if ((efsCvScoreList.get(0) <= cvScore) && (cvScore <= efsCvScoreList.get(1))) // 646<=cvScore<=682
			{
				InterestRate = sheet.getRow(2).getCell(2).getNumericCellValue();
			}
			if ((efsCvScoreList.get(2) <= cvScore) && (cvScore <= efsCvScoreList.get(3))) // 627<=cvScore<=645
			{
				InterestRate = sheet.getRow(4).getCell(2).getNumericCellValue();
			}
			if ((efsCvScoreList.get(4) <= cvScore) && (cvScore <= efsCvScoreList.get(5))) // 610<=cvScore<=625
			{
				InterestRate = sheet.getRow(6).getCell(2).getNumericCellValue();
			}
			if ((efsCvScoreList.get(6) <= cvScore) && (cvScore <= efsCvScoreList.get(7))) // 593<=cvScore<=609
			{
				InterestRate = sheet.getRow(8).getCell(2).getNumericCellValue();
			}
			if ((cvScore <= efsCvScoreList.get(8))) {
				InterestRate = sheet.getRow(10).getCell(2).getNumericCellValue();
			}
		} else if ((randomNumThree <= RandomNumber) && (RandomNumber <= randomNumFour)) // 51 and 100
		{
			if ((efsCvScoreList.get(0) <= cvScore) && (cvScore <= efsCvScoreList.get(1))) // 646<=cvScore<=682
			{
				InterestRate = sheet.getRow(3).getCell(2).getNumericCellValue();
			}
			if ((efsCvScoreList.get(2) <= cvScore) && (cvScore <= efsCvScoreList.get(3))) // 627<=cvScore<=645
			{
				InterestRate = sheet.getRow(5).getCell(2).getNumericCellValue();
			}
			if ((efsCvScoreList.get(4) <= cvScore) && (cvScore <= efsCvScoreList.get(5))) // 610<=cvScore<=625
			{
				InterestRate = sheet.getRow(7).getCell(2).getNumericCellValue();
			}
			if ((efsCvScoreList.get(6) <= cvScore) && (cvScore <= efsCvScoreList.get(7))) // 593<=cvScore<=609
			{
				InterestRate = sheet.getRow(9).getCell(2).getNumericCellValue();
			}
			if ((cvScore <= efsCvScoreList.get(8))) {
				InterestRate = sheet.getRow(10).getCell(2).getNumericCellValue();
			}

		}

		else {
			if ((cvScore <= efsCvScoreList.get(8))) // cvScore 592
			{
				InterestRate = sheet.getRow(10).getCell(2).getNumericCellValue();
			}
		}
		double inrate = Double.valueOf(InterestRate);
		System.out.println(inrate);

		loginAsFSR();
		Thread.sleep(3000);
		landOnAppPage();
		Thread.sleep(5000);
		driver.switchTo().defaultContent();

		WebElement Int = driver.findElement(By.xpath(prop.getProperty("otherTable")));
		new Actions(driver).moveToElement(Int).perform();
		Thread.sleep(3000);
		JavascriptExecutor js = (JavascriptExecutor) driver;
		js.executeScript("document.body.style.zoom='70%'");
		Thread.sleep(3000);

		String screenShotPathforInterestRate = Screenshot.capture(driver, "CaculateInterestRate");
		js.executeScript("document.body.style.zoom='100%'");
		driver.switchTo().frame(driver.findElement(By.xpath(prop.getProperty("switchIframe"))));

		// Displaying Interest Rate result
		System.out.println("Actual Interest rate: " + inrate);
		System.out.println("Expected Interest rate: " + ExpInt);

		if (ExpInt == inrate) {

			test.log(Status.PASS, MarkupHelper.createLabel("InterestRate Percentage Actual value : " + inrate + "%",
					ExtentColor.GREEN));
			test.log(Status.PASS, MarkupHelper.createLabel("InterestRate Percentage Expected value : " + ExpInt + "%",
					ExtentColor.GREEN));

			test.log(Status.PASS, MarkupHelper.createLabel(" Interest Rate Calculation is Matching with GDS Decision",
					ExtentColor.GREEN));
			test.log(Status.PASS, "Snapshot below: " + test.addScreenCaptureFromPath(screenShotPathforInterestRate));
			System.out.println("PASSED in Interest Verification");
		} else {
			System.out.println(inrate + " is the Actual value.");

			test.log(Status.FAIL, MarkupHelper.createLabel("InterestRate Percentage Actual value : " + inrate + "%",
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
            System.out.println("File rename success");;
        }else{
            System.out.println("File rename failed");
        }
        JSONParser parser = new JSONParser();
        Object obj = parser.parse(new FileReader(System.getProperty("user.dir")+"\\src\\main\\resources\\logs\\Response.json"));
        JSONObject jsonObject =  (JSONObject) obj;
        JSONArray cars = (JSONArray) jsonObject.get("Response");
        String txt =cars.toString();
        int index1 = txt.indexOf("<DE_NewInterestRate_RandomNumber>");
        String roar1 = txt.substring(index1+33, index1+35);
        
		if (loanType.contains("New")) {
			if (bkStrategy.contains("Bankruptcy QLA Decrease")) {
				
				int index2 = txt.indexOf("<DE_UPL_App1_BKQualifiedLoanAmount_Decrement>");
		        String roar2 = txt.substring(index2+45, index2+50);
		        if(roar2.contains("-"))
		        {
		        double bkdec = Double.valueOf(roar2);
				bkDecreaseAmount = (int) bkdec;
		        }
             	System.out.println("BK Decrease Amount : " + bkDecreaseAmount);
			}
		}
		double RandomNum = Double.valueOf(roar1);
		int RandomNumber = (int) RandomNum;
        System.out.println("RandomNumber: " + RandomNumber);
        
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
		
		
		
		// Reading values for CV score, Random number and Interest Rate from Excel
		ReadExcel objExcelFile = new ReadExcel();

		// Prepare the path of excel file

		String filePath = System.getProperty("user.dir") + "\\src\\main\\resources\\Excel";

		// Call read file method of the class to read data

		objExcelFile.readExcel(filePath, "UAT-SF-GDScalculation.xlsx", "UPLInterestRate(Bankruptcy)");
		ArrayList List1 = objExcelFile.getlist1();
		ArrayList List2 = objExcelFile.getlist2();
		ArrayList List3 = objExcelFile.getlist3();
		// Get cv Scores from Excel

		int cvCondition1 = (int) List1.get(0);
		int cvCondition2 = (int) List1.get(8);
		int cvCondition3 = (int) List3.get(8);

		// Get Random Numbers from Excel

		int randomCndition1 = (int) List1.get(2);
		int randomCndition2 = (int) List1.get(3);
		int randomCndition9 = (int) List3.get(3);
		int randomCndition3 = (int) List1.get(4);
		int randomCndition10 = (int) List3.get(4);
		int randomCndition4 = (int) List1.get(5);
		int randomCndition11 = (int) List3.get(5);
		int randomCndition5 = (int) List1.get(10);
		int randomCndition6 = (int) List1.get(11);
		int randomCndition13 = (int) List3.get(11);
		int randomCndition7 = (int) List1.get(12);
		int randomCndition14 = (int) List3.get(12);

		// Get Interest Rates from Excel
		double interestCondition1 = (double) List2.get(0);
		double interestCondition2 = (double) List2.get(1);
		double interestCondition3 = (double) List2.get(2);
		double interestCondition4 = (double) List2.get(3);
		double interestCondition5 = (double) List2.get(8);
		double interestCondition6 = (double) List2.get(9);
		double interestCondition7 = (double) List2.get(10);
		double interestCondition8 = (double) List2.get(15);
		double interestCondition9 = (double) List2.get(20);

		// Check Interest Rate

		double intRate = 0;

		if (Province != "Quebec") {

			if (cvScore >= cvCondition1)

			{
				if (RandomNumber == randomCndition1) {
					intRate = interestCondition1;
				} else if ((RandomNumber >= randomCndition2) && (RandomNumber <= randomCndition9)) {
					intRate = interestCondition2;
					System.out.println(intRate);
				} else if ((RandomNumber >= randomCndition3) && (RandomNumber <= randomCndition10)) {
					intRate = interestCondition3;
					System.out.println(intRate);
				} else if ((RandomNumber >= randomCndition4) && (RandomNumber <= randomCndition11)) {
					intRate = interestCondition4;
				}

			}

			else if ((cvScore >= cvCondition2) && (cvScore < cvCondition3))

			{
				if (RandomNumber == randomCndition5) {
					intRate = interestCondition5;
				} else if ((RandomNumber >= randomCndition6) && (RandomNumber <= randomCndition13)) {
					intRate = interestCondition6;
				} else if ((RandomNumber >= randomCndition7) && (RandomNumber <= randomCndition14)) {
					intRate = interestCondition7;
				}

			} else {

				intRate = interestCondition8;
			}
		}

		if (Province == "Quebec") {
			intRate = interestCondition9;
		}
		loginAsFSR();
		Thread.sleep(3000);
		landOnAppPage();
		Thread.sleep(5000);
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
		System.out.println("Actual Interest rate: " + intRate);
		System.out.println("Expected Interest rate: " + ExpInt);

		if (ExpInt == intRate) {

			test.log(Status.PASS, MarkupHelper.createLabel("InterestRate Percentage Actual value : " + intRate + "%",
					ExtentColor.GREEN));
			test.log(Status.PASS, MarkupHelper.createLabel("InterestRate Percentage Expected value : " + ExpInt + "%",
					ExtentColor.GREEN));

			test.log(Status.PASS, MarkupHelper.createLabel(" Interest Rate Calculation is Matching with GDS Decision",
					ExtentColor.GREEN));
			test.log(Status.PASS, "Snapshot below: " + test.addScreenCaptureFromPath(screenShotPathforInterestRate));
			System.out.println("PASSED in Interest Verification");
		} else {
			System.out.println(intRate + " is the expected value.");

			test.log(Status.FAIL, MarkupHelper.createLabel("InterestRate Percentage Actual value : " + intRate + "%",
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

	public void SecondPopup() throws Exception {

		attemptNo++;
		driver.switchTo().defaultContent();
		WebElement page1 = driver.findElement(By.xpath(prop.getProperty("uplPage1")));
		new Actions(driver).moveToElement(page1).perform();

		// We are declaring the frame
		JFrame frmOpt = new JFrame(); // We are declaring the frame
		frmOpt.setAlwaysOnTop(true);// This is the line for displaying it above all windows

		Thread.sleep(1000);
		String s = "<html>Press 1 For Re-Submission with Applicant<br>Press 2 For Results<br>";

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
			System.out.println("Re-Submission with  Applicant");
			resubmitForDecision();

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

	public void resubmitForDecision() throws Exception {

		System.out.println(attemptNo);

		Thread.sleep(10000);

		// If Edit Menu entered
		try {
			driver.manage().timeouts().implicitlyWait(0, TimeUnit.SECONDS);

			if (driver.findElement(By.xpath(prop.getProperty("Appedit"))).isDisplayed()) {
				System.out.println("Edit Menu");
				JavascriptExecutor execute = (JavascriptExecutor) driver;
				execute.executeScript("window.alert = function () { return true}");

				Thread.sleep(5000);

			}
		} catch (Exception e) {
			System.out.println("Not Editing");
			Thread.sleep(5000);
			JavascriptExecutor execute = (JavascriptExecutor) driver;
			execute.executeScript("window.alert = function () { return true}");
		}

		WebDriverWait waitLoading = new WebDriverWait(driver, 360, 0000000);
		waitLoading.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("alertmessage"))));
		System.out.println("Alert is displayed");
		System.out.println("Clicked Save button");
		Thread.sleep(10000);
		WebDriverWait waitLoad12 = new WebDriverWait(driver, 360, 0000000);
		waitLoad12.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath(prop.getProperty("alertmessage"))));
		System.out.println("Alert not displayed");

		Thread.sleep(7000);
		System.out.println("Clicked Re-submit");
		Thread.sleep(7000);
		firstPopup();
		getUPLdetails();
		if (loanType.contains("New")) {
			getStrategy();
		}
		calculateIncome();
		calculateLiability();

		// SWITCH IFRAME

		Thread.sleep(3000);
		if (loanType.contains("New")) {
			// Check Strategy
			if (Strategy.contains("Credit Vision")) {
				remInCal();
				calculateQLA();
			} else if (Strategy.contains("CAB") && IntRate.contains("46.96")) {
				cabQla();
			} else {
				remInCal();
				calculateQLA();
			}
		} else if (loanType.contains("Increase")) {
			remInCal();
			calculateQLA();
		}
		maxHA();
		Thread.sleep(3000);
		SecondPopup();
	}

	public void waitForLoadingIconToFinish() throws Exception {

		// Wait for loading icon to disappear before recording objects (which is in main
		// frame)

		// Switch to parent frame as loading icons are located there
		driver.switchTo().defaultContent();

		// Wait for main loading icon to disappear if it is present
		// Set implicit webdriver wait to 0 to search faster
		driver.manage().timeouts().implicitlyWait(0, TimeUnit.SECONDS);

		try {
			if (driver.findElements(By.xpath("//*[@id='loadingPanel']")).size() > 0) {
				WebElement loadingIcon = driver.findElement(By.xpath("//*[@id='loadingPanel']"));
				WebDriverWait waitLoadingIcon = new WebDriverWait(driver, 360, 0000);
				waitLoadingIcon.until(ExpectedConditions.invisibilityOf(loadingIcon));
				System.out.println("waitLoadingIcon");
			}

			// Wait for transfer loading icon to disappear if it is present
			if (driver.findElements(By.xpath("//*[@id='loadingTransferPanel']")).size() > 0) {
				WebElement loadingTransferIcon = driver.findElement(By.xpath("//*[@id='loadingTransferPanel']"));
				WebDriverWait waitLoadingTransferIcon = new WebDriverWait(driver, 360, 0000);
				waitLoadingTransferIcon.until(ExpectedConditions.invisibilityOf(loadingTransferIcon));
				System.out.println("loadingTransferIcon");
			}
		} catch (org.openqa.selenium.NoSuchElementException | org.openqa.selenium.StaleElementReferenceException
				| org.openqa.selenium.TimeoutException e) {
		}

		// Set webdriver timeout back to 60 seconds
		driver.manage().timeouts().implicitlyWait(60, TimeUnit.SECONDS);

	}

	public void remInCal() throws InterruptedException, IOException {
		org.apache.poi.ss.usermodel.Sheet sheet;

		File file = new File(
				System.getProperty("user.dir") + "\\src\\main\\resources\\Excel\\UAT-SF-GDScalculation.xlsx");

		FileInputStream inputStream = new FileInputStream(file);

		Workbook workbook = new XSSFWorkbook(inputStream);

		sheet = workbook.getSheet("UPLRemainingIncome");

		// Making cell values as variable

		int cvSco = (int) sheet.getRow(2).getCell(1).getNumericCellValue();

		double intRate1 = sheet.getRow(5).getCell(5).getNumericCellValue();
		double intRate2 = sheet.getRow(5).getCell(7).getNumericCellValue();

		String groupA = sheet.getRow(7).getCell(8).getStringCellValue();
		String[] Riskgrp = groupA.split(" or ");
		String Rg1 = Riskgrp[0];
		String Rg2 = Riskgrp[1];
		String groupB = sheet.getRow(9).getCell(8).getStringCellValue();
		String[] Riskgroup = groupB.split(",");
		String Rg3 = Riskgroup[0];
		String Rg4 = Riskgroup[1];
		String Rg5 = Riskgroup[2];
		String Rg6 = Riskgroup[3];

		double value1 = sheet.getRow(2).getCell(10).getNumericCellValue();
		double value2 = sheet.getRow(3).getCell(10).getNumericCellValue();
		double value3 = sheet.getRow(4).getCell(10).getNumericCellValue();
		double value4 = sheet.getRow(5).getCell(10).getNumericCellValue();
		double value5 = sheet.getRow(6).getCell(10).getNumericCellValue();
		double value6 = sheet.getRow(7).getCell(10).getNumericCellValue();
		double value7 = sheet.getRow(8).getCell(10).getNumericCellValue();
		double value8 = sheet.getRow(9).getCell(10).getNumericCellValue();

		double InterestRate = Double.parseDouble(IntRate.replace("%", ""));

		System.out.println("CV Score :" + cvScore);
		System.out.println("AppType :" + appType);
		System.out.println("Strategy :" + Strategy);
		System.out.println("InterestRate :" + InterestRate + "%");
		System.out.println("Total Income :$" + TotalIncome);
		System.out.println("Total Debt :$" + TotalDebt);

		if (cvScore >= cvSco) {
			if (appType.equalsIgnoreCase("Own"))

			{
				if (Strategy.contains("Credit Vision") || Strategy.contains("Increase")) {
					RemainingIncome = TotalIncome * value1 - TotalDebt;
				} else if (Strategy.contains("CAB")) {
					if (!(InterestRate == intRate1) && !(InterestRate == intRate2)) {
						RemainingIncome = TotalIncome * value4 - TotalDebt;
					} else if (InterestRate == intRate1 || InterestRate == intRate2) {
						if (cabKey.contains(Rg1) || cabKey.contains(Rg2)) {
							RemainingIncome = TotalIncome * value6 - TotalDebt;
						} else if (cabKey.contains(Rg3) || cabKey.contains(Rg4) || cabKey.contains(Rg5)
								|| cabKey.contains(Rg6)) {
							RemainingIncome = TotalIncome * value8 - TotalDebt;
						}
					}
				}
			}

			else if (appType.equalsIgnoreCase("Rent")) {
				if (Strategy.contains("Credit Vision") || Strategy.contains("Increase")) {
					RemainingIncome = TotalIncome * value2 - TotalDebt;
				} else if (Strategy.contains("CAB")) {
					if (!(InterestRate == intRate1) && !(InterestRate == intRate2)) {
						RemainingIncome = TotalIncome * value5 - TotalDebt;
					} else if (InterestRate == intRate1 || InterestRate == intRate2) {
						if (cabKey.contains(Rg1) || cabKey.contains(Rg2)) {
							RemainingIncome = TotalIncome * value7 - TotalDebt;
						} else if (cabKey.contains(Rg3) || cabKey.contains(Rg4) || cabKey.contains(Rg5)
								|| cabKey.contains(Rg6)) {
							RemainingIncome = TotalIncome * value8 - TotalDebt;
						}
					}
				}
			}
		}

		else if (cvScore < cvSco) {
			if (appType.equalsIgnoreCase("Own")) {

				if (Strategy.contains("Credit Vision") || Strategy.contains("Increase")) {
					RemainingIncome = TotalIncome * value3 - TotalDebt;
				} else if (Strategy.equalsIgnoreCase("CAB")) {
					if (InterestRate == intRate1 || InterestRate == intRate2) {
						if (cabKey.contains(Rg1) || cabKey.contains(Rg2)) {
							RemainingIncome = TotalIncome * value6 - TotalDebt;
						} else if (cabKey.contains(Rg3) || cabKey.contains(Rg4) || cabKey.contains(Rg5)
								|| cabKey.contains(Rg6)) {
							RemainingIncome = TotalIncome * value8 - TotalDebt;
						}
					}
				}
			} else if (appType.equalsIgnoreCase("Rent")) {
				if (Strategy.contains("Credit Vision") || Strategy.contains("Increase")) {
					RemainingIncome = TotalIncome * value3 - TotalDebt;
				} else if (Strategy.contains("CAB")) {
					if ((InterestRate == intRate1) || (InterestRate == intRate2)) {
						if (cabKey.contains(Rg1) || cabKey.contains(Rg2)) {
							RemainingIncome = TotalIncome * value7 - TotalDebt;
						} else if (cabKey.contains(Rg3) || cabKey.contains(Rg4) || cabKey.contains(Rg5)
								|| cabKey.contains(Rg6)) {
							RemainingIncome = TotalIncome * value8 - TotalDebt;
						}
					}
				}
			}

		}

		System.out.println("RemainingIncome :$" + RemainingIncome);
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

		org.apache.poi.ss.usermodel.Sheet sheet = workbook.getSheet(IntRate);

		Iterator<Row> rows = sheet.iterator();

		Row row = rows.next();

		Iterator<Cell> cell = row.cellIterator();

		Cell value;

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
		
		if (Strategy.contains("Banking Declined"))
		{
			if(ActualQLA>4100)
			{
				ActualQLA=4100;
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

	public void cabQla() throws InterruptedException, IOException {
		Thread.sleep(5000);
		if (attemptNo == 0) {
			test = Extent.createTest("QLA Calculation");
		}

		else {

			test = Extent.createTest("Resubmission Result: Attempt#  " + attemptNo + " - QLA Calculation");
		}

		String riskGroupValue;
		double newpaymentQLA = 0;
		String cabKey = driver.findElement(By.xpath(prop.getProperty("cabKey"))).getText();
		System.out.println("Risk Group :" + cabKey);

		riskGroupValue = cabKey;

		File file = new File(System.getProperty("user.dir")
				+ "\\src\\main\\resources\\Excel\\EFS CAB Lending Limit Grid V4 Latest.xlsx");

		FileInputStream inputStream = new FileInputStream(file);

		Workbook workbook = new XSSFWorkbook(inputStream);

		org.apache.poi.ss.usermodel.Sheet sheet = workbook.getSheet("EFS CAB Lending Limit Grid V3 C");

		// Reading through CAB Lending Limit sheet for identifying RISK GROUP
		for (int r = 1; r <= sheet.getLastRowNum(); r++) {
			if (sheet.getRow(r).getCell(2).getStringCellValue().equalsIgnoreCase(riskGroupValue)) {
				newpaymentQLA = sheet.getRow(r).getCell(1).getNumericCellValue();
			}

		}
		System.out.println("New Payment QLA is " + newpaymentQLA);

		if (newpaymentQLA != 0) {

			sheet = workbook.getSheet("46.96");

			Iterator<Row> rows = sheet.iterator();

			Row row = rows.next();

			Iterator<Cell> cell = row.cellIterator();

			Cell value;

			// Setting Province label

			System.out.println("Province is " + Province);

			stringSplit = Province.split(" - ");
			String Prov = stringSplit[0];

			if (Prov.contains("ON") || Prov.contains("MB")) {
				provinceGroup = "ON,MB";
			} else if (Prov.contains("NL")) {
				provinceGroup = "NL";
			} else if (Prov.contains("SK")) {
				provinceGroup = "SK";
			} else {
				provinceGroup = "Other";
			}

			// Identifying the province
			while (cell.hasNext()) {

				value = cell.next();

				if (value.getStringCellValue().contains(provinceGroup)) {
					fcol = value.getColumnIndex();
					lcol = fcol + 1;
					break;
				}

			}
			// System.out.println(fcol);

			// Reading through 46.96 sheet for identifying RISK GROUP

			for (int r = 2; r <= sheet.getLastRowNum(); r++) {
				if (sheet.getRow(r).getCell(lcol).getNumericCellValue() > newpaymentQLA) {

					// System.out.println(sheet.getRow(r-1).getCell(lcol).getNumericCellValue());
					QLA = sheet.getRow(r - 1).getCell(fcol).getNumericCellValue();

					break;
				}
			}
		} else {
			QLA = 0.0;
			System.out.println("QLA is 0.0"); // This is for 0s in qla_selected column in CAB Lending Limit sheet
		}

		// Calculation QLA

		double ActualQLA;

		Thread.sleep(3000);

		if (QLA == 0.0) {
			ActualQLA = QLA;
		} else {
			ActualQLA = QLA + 100;
		}
		if (ActualQLA != ExpectedQLA) {

			if (bkStrategy.contains("Bankruptcy QLA Decrease")) {
				ActualQLA = ActualQLA + bkDecreaseAmount;

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

	@AfterTest()
	public void resetAttemptCount() {

		attemptNo = 0;

	}

}

	
	
	


