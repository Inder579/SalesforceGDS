package resources;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcel {
	ArrayList results1, results2, results3; 
	
	public void readExcel(String filePath,String fileName,String sheetName) throws IOException{

	    //Create an object of File class to open xlsx file

	    File file =    new File(filePath+"\\"+fileName);

	    //Create an object of FileInputStream class to read excel file

	    FileInputStream inputStream = new FileInputStream(file);

	    Workbook workbook = null;

	    //Find the file extension by splitting file name in substring  and getting only extension name

	    String fileExtensionName = fileName.substring(fileName.indexOf("."));

	    //Check condition if the file is xlsx file

	    if(fileExtensionName.equals(".xlsx")){

	    //If it is xlsx file then create object of XSSFWorkbook class

	    workbook = new XSSFWorkbook(inputStream);

	    }

	    //Check condition if the file is xls file

	    else if(fileExtensionName.equals(".xls")){

	        //If it is xls file then create object of HSSFWorkbook class

	        workbook = new HSSFWorkbook(inputStream);

	    }

	    //Read sheet inside the workbook by its name

	    org.apache.poi.ss.usermodel.Sheet sheet = workbook.getSheet(sheetName);

	    int rowCount = sheet.getLastRowNum()-sheet.getFirstRowNum();

	    //Get CV score and Random number from excel
	    results3 = new ArrayList();
	    for (int i = 3; i < rowCount+1; i++) {

	        Row row = sheet.getRow(i);

	        //Create a loop to print cell values in a row

	        for (int j = 5; j < 6; j++) {

	            //Print Excel data in console
	        	int res = (int) row.getCell(j).getNumericCellValue();
	        	results3.add(res);
	        	
	        }

	    
	    } 
	    //Get CV score and Random number from excel
	    results1 = new ArrayList();
	    for (int i = 3; i < rowCount+1; i++) {

	        Row row = sheet.getRow(i);

	        //Create a loop to print cell values in a row

	        for (int j = 2; j < 3; j++) {

	            //Print Excel data in console
	        	int res = (int) row.getCell(j).getNumericCellValue();
	        	results1.add(res);
	        	
	        }

	    
	    } 
	    
	
	// Get Interest Rates from Excel    
	    results2 = new ArrayList();
	    for (int i = 5; i < rowCount+1; i++) {

	        Row row = sheet.getRow(i);

	        //Create a loop to print cell values in a row

	        for (int j = 6; j < 7; j++) {

	            //Print Excel data in console
	        	double res = row.getCell(j).getNumericCellValue();
	        	results2.add(res);
	        	

	        }

	       
	    }
	    }  
	
	public ArrayList getlist1()
	{
		return results1;
	}

	public ArrayList getlist2()
	{
		return results2;
	}
	public ArrayList getlist3()
	{
		return results3;
	}

}
