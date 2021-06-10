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

public class ExpressCoApp extends GDScoApp {
	
	@Test
	public void m1() throws Exception
	{
		loginAsAdmin();
		loginAsFSR();
		Thread.sleep(4000);
//	 driver.get("https://c.cs29.visual.force.com/apex/LAMSApplicationView1?id=a080r0000014IQ8&sfdc.override=1");
		
		waitForExFirstSubmission();
		Thread.sleep(3000);
		firstPopup();
		Thread.sleep(4000);
		
		getUPLExdetails();
		calculateIncome();
		Thread.sleep(2000);
		calculateLiability();
		// SWITCH IFRAME DEFAULT
		driver.switchTo().defaultContent();
		getAppTimestampLogs();
		// Interest rate Calculations
		interestRateCalculation();
		checkContributer();
		if (LoanType.contains("New")) {
			Thread.sleep(3000);
			// Check Strategy
			if (Strategy.contains("Credit Vision")) {
				remInCal();
				calculateQLA();
			} else if (Strategy.contains("CAB") && IntRate.contains("44.96")) {
				cabQla();
				cabqlaCoApp();
			} else {
				remInCal();
				calculateQLA();
			}
		} else if (LoanType.contains("Increase")) {
			Thread.sleep(3000);
			remInCal();
			calculateQLA();
		}
		maxHA();
		ExReasonCode();
		SecondPopupEx();
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
	private void waitForLoadingIconToFinish() {
		// TODO Auto-generated method stub
		
	}
	public void getUPLExdetails() throws InterruptedException {

		WebDriverWait waitLoad = new WebDriverWait(driver, 360, 0000);
		waitLoad.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("getprovince"))));
		Province = driver.findElement(By.xpath(prop.getProperty("getprovince"))).getText();
		CoAppProvince = driver.findElement(By.xpath(prop.getProperty("getCoAppprovince"))).getText();
		ApplicationID = driver.findElement(By.xpath(prop.getProperty("getappid"))).getText();
		loanType = driver.findElement(By.xpath(prop.getProperty("getloantype"))).getText();
		LoanType = driver.findElement(By.xpath(prop.getProperty("loantype"))).getText();
		appType = driver.findElement(By.xpath(prop.getProperty("getapptype"))).getText();
		Thread.sleep(3000);
     	driver.switchTo().frame(driver.findElement(By.xpath(prop.getProperty("switchIframe"))));
		Thread.sleep(3000);
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
		String cvscoreco = driver.findElement(By.xpath(prop.getProperty("getcvcoapp"))).getText();
		cvScoreCoapp = Integer.parseInt(cvscoreco);
		Strategy = driver.findElement(By.xpath(prop.getProperty("qlastrategy"))).getText();
		if (loanType.contains("New")) {
			cabKey = driver.findElement(By.xpath("//th[contains(text(),'CAB Key')]/following-sibling::td[1]/span"))
					.getText();
			cabKeyCoApp = driver.findElement(By.xpath(prop.getProperty("cabkeyCoApp"))).getText();
			bkStrategy = driver.findElement(By.xpath(prop.getProperty("bkstrategy"))).getText();
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
	public void SecondPopupEx() throws Exception {
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
			resubmitForDecisionEx();

			break;

		case 2:

			// Function for Removal Co App
			System.out.println("Re-Submission with Removal of Co Applicant");
			removeCoAppEx();
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
	public void resubmitForDecisionEx() throws Exception {

		System.out.println(attemptNo);

		
		firstPopup();
		// SWITCH IFRAME DEFAULT
		driver.switchTo().defaultContent();
		getUPLExdetails();
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
		if (LoanType.contains("New")) {
			// Check Strategy
			if (Strategy.contains("Credit Vision")) {
				remInCal();
				calculateQLA();
			} else if (Strategy.contains("CAB") && IntRate.contains("44.96")) {
				cabQla();
				cabqlaCoApp();
			} else {
				remInCal();
				calculateQLA();
			}
		} else if (LoanType.contains("Increase")) {
			remInCal();
			calculateQLA();
		}
		maxHA();
		ExReasonCode();
		SecondPopupEx();
	}
	public void removeCoAppEx() throws Exception {
		
		
		firstPopup();
		// SWITCH IFRAME DEFAULT
		driver.switchTo().defaultContent();

		getApplicationDetailsEx();
		calculateAppIncome();
		Thread.sleep(2000);
		calculateAppLiability();

		Thread.sleep(3000);
		// Logging in as FSR
		if (LoanType.contains("New")) {
			// Check Strategy
			if (Strategy.contains("Credit Vision")) {
				remInCal();
				calculateQLA();
			} else if (Strategy.contains("CAB") && IntRate.contains("46.96")) {
				qla();
			} else {
				remInCal();
				calculateQLA();
			}
		} else if (LoanType.contains("Increase")) {
			remInCal();
			calculateQLA();
		}
		maxHA();
		ExReasonCode();
		SecondPopupEx();

	}
	public void getApplicationDetailsEx() throws InterruptedException, IOException {
		
		WebDriverWait waitLoad = new WebDriverWait(driver, 360, 0000);
		waitLoad.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("getprovince"))));
		Province = driver.findElement(By.xpath(prop.getProperty("getprovince"))).getText();
		ApplicationID = driver.findElement(By.xpath(prop.getProperty("getappid"))).getText();
		appType = driver.findElement(By.xpath(prop.getProperty("getapptype"))).getText();
		LoanType = driver.findElement(By.xpath(prop.getProperty("loantype"))).getText();
		System.out.println("App Type :" + appType);

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
		String cvscore = driver.findElement(By.xpath(prop.getProperty("getcvscore"))).getText();
		cvScore = Integer.parseInt(cvscore);
		Strategy = driver.findElement(By.xpath(prop.getProperty("qlastrategy"))).getText();
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
}
