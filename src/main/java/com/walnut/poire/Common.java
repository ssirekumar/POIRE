package com.walnut.poire;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.PrintWriter;
import java.io.StringWriter;
import java.io.Writer;
import java.math.BigInteger;
import java.security.SecureRandom;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.Properties;
import java.util.Random;
import org.apache.log4j.Logger;


public class Common {

	private static final Logger _log = Logger.getLogger(LogsBeans.class);
    private static String dateFormatedValue = "";
    private static SecureRandom random = new SecureRandom();
			
	/**
	 * Function to load the property file
	 * @author Siri Kumar Puttagunta
	 * @param String - path to the property file
	 * @return Properties object if success else null*/
	public static Properties loadPropertyFile(String path){
		LoggerPoire.setLoggerObject(Common._log);
		Properties _props = new Properties();
		try{
			try {
				_props.load(new FileInputStream(path));
			} catch (IOException io) {
				InputStream in = ClassLoader.getSystemResourceAsStream(path);
				_props.load(in);
			}
		} catch(IOException ioe){
			System.err.println(ioe.getMessage());
			LoggerPoire.log(Globals.Log.ERROR, ioe);
		}
		return _props;
	}

	/**
	 * Function to retrieve value from property file
	 * @author Siri Kumar Puttagunta
	 * @param Properties - object, String key
	 * @return String value if found else null*/
	public static String getValueFromProperty(Properties _props, String key){
		LoggerPoire.setLoggerObject(Common._log);
		String value = null;
		if(_props != null){
			value = _props.getProperty(key);
		} else {
			System.err.println("Inside getValueFromProperty(_props,key) method : Object is NULL");
			LoggerPoire.log(Globals.Log.ERROR,	"Inside getValueFromProperty(_props,key) method : Object is NULL");
		}
		return value;
	}
	
	/**
	 * Method to add the exception logger log to file with printStackTrace
	 * @author 	Siri Kumar Puttagunta
	 * @param  	Exception - exception object
	 * @return 	void 
	 * */
	public static void addExceptionLogger(Exception e){
		LoggerPoire.setLoggerObject(Common._log);
		String s = "";
		Writer writer = new StringWriter();
		PrintWriter printWriter = new PrintWriter(writer);
		e.printStackTrace(printWriter);
		s = writer.toString();
		System.err.println("ERROR: ********* Exception raised check the POIRE Log file for exception stack trace *********");
		System.err.println("****************************** POIRE Exception Starts******************************");
		System.err.println(e.getMessage());
		System.err.println("****************************** POIRE Exception End*********************************");
		LoggerPoire.log(Globals.Log.ERROR, "**************** Exception stack trace starts");
		LoggerPoire.log(Globals.Log.ERROR, e.getMessage());
		LoggerPoire.log(Globals.Log.ERROR, s);
	}
	
	/**
	 * Method to add the exception logger to log file with printStackTrace
	 * @author 	Siri Kumar Puttagunta
	 * @param  	Exception - exception object
	 * @param  	methodName - Add method name where it raised.
	 * @return 	void 
	 * */
	public static void addExceptionLogger(Exception e, String methodName){
		LoggerPoire.setLoggerObject(Common._log);
		String s = "";
		Writer writer = new StringWriter();
		PrintWriter printWriter = new PrintWriter(writer);
		e.printStackTrace(printWriter);
		s = writer.toString();
		System.err.println("ERROR: ********* Exception raised check the POIRE Log file for exception stack trace *********");
		System.err.println("********** POIRE Exception Starts**********");
		System.err.println(e.getMessage());
		System.err.println("********** POIRE Exception Ends**********");
		LoggerPoire.log(Globals.Log.ERROR, "****************Exception stacktrace starts:");
		LoggerPoire.log(Globals.Log.ERROR, "Exception arised at Method in POIRE: "+methodName);
		LoggerPoire.log(Globals.Log.ERROR, e.getMessage());
		LoggerPoire.log(Globals.Log.ERROR, s);
	}
    
	/**
	 * It is used to set the Date format as per java
	 * Ex: dd/M/yyyy , dd-M-yyyy hh:mm:ss, yyyy/MM/dd, yyyy/MM/dd HH:mm:ss, dd-M-yyyy
	 * @author 	Siri Kumar Puttagunta
	 * @param  	formatString - String data of the date format 
	 *          Ex: dd/M/yyyy , dd-M-yyyy hh:mm:ss, yyyy/MM/dd, yyyy/MM/dd HH:mm:ss, dd-M-yyyy
	 * @return 	void 
	 * */
	public static void setdateFormatString(String formatString){
		dateFormatedValue = formatString;
	}
	
	/**
	 * It is used to get the Date format according to the setdateFormatString() method.
	 * @author 	Siri Kumar Puttagunta
	 * @return 	String 
	 * */
	public static String getdateFormatString(){
		return dateFormatedValue;
	}
	
	/**
	 * It is used to get the collection as per the filter contain value
	 * In this if the collection {'ada','ade','abc','dad','gee','ewerwe','abc','isk'} list having 'abc'
	 * then if you give a parameter as abc from the this the next values will be returned as arrayList.
	 * the contain value is at first occurrence. 
	 * Ex: {'ada','ade','abc','dad','gee','ewerwe','abc','isk'} out put as {'dad','gee','ewerwe','abc','isk'}
	 * @author 	Siri Kumar Puttagunta
	 * @param  	listContents - arrayList object
	 * @param   containValue - string value in the list         
	 * @return 	ArrayList 
	 * */
	public static ArrayList<String> getlistDataFromSpecifiedValue(ArrayList<String> listContents, String containValue){
		int indexOfContain = 0;
		ArrayList<String> _filteredData = null;
		try {
			for (int i = 0; i < listContents.size(); i++) {
				if(listContents.get(i).equalsIgnoreCase(containValue)){
					indexOfContain = i;
					break;
				}
			}
			_filteredData = new ArrayList<String>();
			for (int i = indexOfContain+1; i < listContents.size(); i++) {
				_filteredData.add(listContents.get(i));
			}
		} catch(Exception e){
			addExceptionLogger(e);
		}
		return _filteredData;
	}
    	
	/**
	 * It will be used to search in boolean array contain target value or not
	 * @author 	Siri Kumar Puttagunta
	 * @param  	array - Array variable.
	 * @param   target - to check the contain value in array true/false.         
	 * @return 	boolean -  if target value present return true or else false.
	 * */
	public static boolean contains(boolean[] array, boolean target)
	  {
	    for (boolean value : array) {
	      if (value == target) {
	        return true;
	      }
	    }
	    return false;
	  }
	
	/**
	 * Function to return random String
	 * @author 	Siri Kumar Puttagunta
	 * @param  	int - length for the String to be returned
	 * @return 	random string with length equals to the parameter passed*/
	public static String getRandomString(int length) {
		LoggerPoire.setLoggerObject(Common._log);
		String result = "";
		try {
			result = new BigInteger(Long.SIZE * length, random).toString(32);
		} catch (Exception e) {
			LoggerPoire.log(Globals.Log.ERROR, e.getMessage());
			LoggerPoire.log(Globals.Log.ERROR, e);
		}
		return result.substring(0, length);
	}
	
	/**
	 * Method to give Date according to format.
	 * @author Siri Kumar Puttagunta
	 * @param String - Format of the date java format(MM/dd/yyyy)
	 * @param int difference - difference of The numbers(-Negative means past, +Positive Future date)
	 * @param String type - type of the days(days, months, years)
	 * @return String  Date-format string()
	 * Ex: "MM/dd/yyyy", -2, "days"
	 * */
	public static String getDate(String dateFormat, int difference, String type) {
		LoggerPoire.setLoggerObject(Common._log);
		Calendar cal = Calendar.getInstance();
		try {
			if (difference != 0) {
				if (type.equalsIgnoreCase("days")) {
					cal.add(Calendar.DATE, difference);
				} else if (type.equalsIgnoreCase("months")) {
					cal.add(Calendar.MONTH, difference);
				} else if (type.equalsIgnoreCase("years")) {
					cal.add(Calendar.YEAR, difference);
				}
			}
			Date updatedDate = cal.getTime();
			SimpleDateFormat format = new SimpleDateFormat(dateFormat);
			format.setLenient(false);
			return format.format(updatedDate);
		} catch (IllegalArgumentException e) {
			LoggerPoire.log(Globals.Log.ERROR, e);
			return null;
		}
	}
	
	/**
	 * Function to return random number
	 * @author 	Siri Kumar Puttagunta
	 * @param 	int - number of digits for the random number
	 * @return 	long random number with number of digits passed in parameter*/
	public static long getRandomNumber(int numberOfDigits) {
		LoggerPoire.setLoggerObject(Common._log);
		long accumulator = 0L;
		try {
			Random randomGenerator = new Random();
			accumulator = 1 + randomGenerator.nextInt(9);
			for (int i = 0; i < (numberOfDigits - 1); i++) {
				accumulator *= 10L;
				accumulator += randomGenerator.nextInt(10);
			}
		} catch (Exception e) {
			LoggerPoire.log(Globals.Log.ERROR, e);
		}
		return accumulator;
	}
    
	/**
	 * Function to return Array of random string
	 * @author 	Siri Kumar Puttagunta
	 * @param 	int - length of the array, int length of each element
	 * @return 	ArrayList as per parameters passed*/
	public static ArrayList<String> getArrayOfRandomString(int lengthOfArray, int lengthOfElements) {
		ArrayList<String> randomArray = null;
		try {
			randomArray = new ArrayList<String>();
			for (int i = 0; i < lengthOfArray; i++) {
				randomArray.add(new BigInteger(Long.SIZE * lengthOfElements, random)
				.toString(32).substring(0, lengthOfElements)
				.toUpperCase());
			}
		} catch (Exception e) {
			LoggerPoire.log(Globals.Log.ERROR, e);
		}
		return randomArray;
	}
    
	/**
	 * Method to check if a package exists or not
	 * @author Siri Kumar Puttagunta
	 * @param String - name of the package
	 * @return true if exists else false
	 * */
	public static boolean doesJavaPackageExists(String packageName) {
		if (!(new File("bin/" + packageName.replace('.', '/'))).exists()) {
			return false;
		} else {
			return true;
		}
	}
	

}
