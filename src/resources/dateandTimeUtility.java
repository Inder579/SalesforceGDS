package resources;

import java.text.SimpleDateFormat;
import java.util.Date;

public class dateandTimeUtility {

	
		public static String getCurrentTimeStamp() {
		    SimpleDateFormat sdfDate = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");//dd/MM/yyyy
		    Date now = new Date();
		    String strDate = sdfDate.format(now);
		    return strDate;
		}
		
		
		public static String getCurrentTimeStampF1() {
		    SimpleDateFormat sdfDate = new SimpleDateFormat("yyyy-MM-dd-HH_mm");//dd/MM/yyyy
		    Date now = new Date();
		    String strDate = sdfDate.format(now);
		    return strDate;
		}
		
		
		
		public static String getCurrentDate() {
		    SimpleDateFormat sdfDate = new SimpleDateFormat("yyyy-MM-dd");//dd/MM/yyyy
		    Date now = new Date();
		    String strDate = sdfDate.format(now);
		    return strDate;
		}
	    
		public static String getCurrentTime() {
		    SimpleDateFormat sdfDate = new SimpleDateFormat("HH-mm-a");//dd/MM/yyyy
		    Date now = new Date();
		    String strDate = sdfDate.format(now);
		    return strDate;
		}
	}
	

