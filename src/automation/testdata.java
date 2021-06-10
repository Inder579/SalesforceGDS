package automation;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

public class testdata {

	
	@Test()
	public void m1() throws Exception {
		splIncreaseRemInc();
	}
	
	public void splIncreaseRemInc() throws IOException {
		File file = new File(System.getProperty("user.dir")
				+ "\\src\\main\\resources\\Excel\\Credit Decisions Calculations - SPL Increase.xlsx");

		FileInputStream inputStream = new FileInputStream(file);

		
		Workbook workbook = new XSSFWorkbook(inputStream);
		// String IntRate = String.valueOf(ExpInt);

		org.apache.poi.ss.usermodel.Sheet sheet = workbook.getSheet("Maximum Affordable QLA");
		
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

		
		
		int cvScore = 565;
		int BehaviourScore = -1;
		double QLA = 50000;
		
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
