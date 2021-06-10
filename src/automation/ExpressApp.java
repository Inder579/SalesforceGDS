package automation;

import java.awt.Dimension;
import java.awt.Font;
import java.io.IOException;
import java.util.concurrent.TimeUnit;

import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JTextPane;
import javax.swing.UIManager;
import javax.swing.plaf.FontUIResource;

import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.Test;

import com.aventstack.extentreports.Status;
import com.aventstack.extentreports.markuputils.ExtentColor;
import com.aventstack.extentreports.markuputils.MarkupHelper;

import resources.Screenshot;

public class ExpressApp extends GdsTest {
	
	
	
	@Test
	public void m1() throws Exception
	{
		loginAsAdmin();
		loginAsFSR();
		Thread.sleep(4000);
	//	driver.get("https://c.cs29.visual.force.com/apex/LAMSApplicationView1?id=a080r00000149V4&sfdc.override=1");
		waitForExFirstSubmission();
		Thread.sleep(5000);
		
		firstPopup();
		Thread.sleep(5000);
		getUPLExdetails();
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
		Thread.sleep(3000);
		
		// Interest rate Calculations
		interestRateCalculation();
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
		ExReasonCode();
		// Second Pop-up - Resubmission
		SecoundPopEx();
		
	}
	public void ExReasonCode() throws InterruptedException, IOException {
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

	public void waitForExFirstSubmission() throws Exception {

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
	public void getUPLExdetails() throws InterruptedException {

		WebDriverWait waitLoad = new WebDriverWait(driver, 360, 0000);
		waitLoad.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("getprovince"))));
		Province = driver.findElement(By.xpath(prop.getProperty("getprovince"))).getText();
		ApplicationID = driver.findElement(By.xpath(prop.getProperty("getappid"))).getText();
		loanType = driver.findElement(By.xpath(prop.getProperty("getloantype"))).getText();
		applicationType=driver.findElement(By.xpath(prop.getProperty("applicationType"))).getText();
		appType = driver.findElement(By.xpath(prop.getProperty("getapptype"))).getText();
		driver.switchTo().frame(driver.findElement(By.xpath(prop.getProperty("switchIframe"))));
		Thread.sleep(4000);
		IntRate = driver.findElement(By.xpath(prop.getProperty("getExpress%"))).getText();
		ExpInt = Double.parseDouble(IntRate.replace("%", ""));

		String intqla = driver.findElement(By.xpath(prop.getProperty("getExpressQla"))).getText();
		String qla = intqla.replace(",", "");
		
		ExpectedQLA = Double.parseDouble(qla.replace("$", ""));
		String intHA=driver.findElement(By.xpath(prop.getProperty("getExmaxHA"))).getText();
		String HA = intHA.replace(",", "");

		ExpectedMaxHA = Double.parseDouble(HA.replace("$", ""));
		String cvscore = driver.findElement(By.xpath(prop.getProperty("getExpresscvscore"))).getText();
		cvScore = Integer.parseInt(cvscore);
		Strategy = driver.findElement(By.xpath(prop.getProperty("qlastrategy"))).getText();
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
		System.out.println("Strategy :" + Strategy);
	}
	
	public void SecoundPopEx() throws InterruptedException, IOException
	{
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
			ResubmitEx();

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

	public void ResubmitEx() throws InterruptedException, IOException
	{
		System.out.println(attemptNo);

	
		firstPopup();
		getUPLExdetails();
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
		ExReasonCode();
		Thread.sleep(3000);
		SecoundPopEx();
	}
	
}
