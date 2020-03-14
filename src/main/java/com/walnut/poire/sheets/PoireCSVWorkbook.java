package com.walnut.poire.sheets;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.sql.Date;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.ListIterator;
import com.walnut.poire.Common;
import com.walnut.poire.Globals;
import com.walnut.poire.LoggerPoire;
import org.apache.log4j.Logger;
import com.opencsv.CSVReader;
import com.opencsv.CSVWriter;

/**
 * @author Ssire kumar Puttagunta</br>
 * @version 2.0.0</br>
 *          <p>
 *          The below class methods are Designed with the help of
 *          OpenCSV-2.3.jar. These methods are useful for doing operations on
 *          the CSV Files The whole class methods are in static only and these
 *          method are useful in all the CSV file formats
 *          </p>
 */

public class PoireCSVWorkbook {
	
	 private static Logger _log = Logger.getLogger(PoireCSVWorkbook.class);
	  
	 /**<p>Get complete CSV data into {@code ArrayList<ArrayList<String>>}</p>
		 * @author <a href="mailto:ssirekumar@gmail.com">Ssire Kumar {@literal <ssirekumar@gmail.com>}</a> </br>
		 * @return {@code ArrayList<ArrayList<String>>}
		 * @param  filePath - File path of the CSV file.
		 * @param  skipLines - Number for how many lines to skip and start reading data.
		 * */
	  public static ArrayList<ArrayList<String>> getCSVData(String filePath, int skipLines){
		  LoggerPoire.setLoggerObject(PoireCSVWorkbook._log);
		  List<String[]> contents = null;
		  String[] rowContents = null;
		  ArrayList<String> row = null;
		  ListIterator<String[]> listIterate = null;
		  ArrayList<ArrayList<String>> csvData = null;
		  try{
			  CSVReader csvReader = new CSVReader(new FileReader(filePath));
			  while (skipLines > 0){
				  csvReader.readNext();
				  skipLines--;
			  }
			  contents = csvReader.readAll();
			  csvData = new ArrayList<ArrayList<String>>();
			  listIterate = contents.listIterator();
			  while (listIterate.hasNext()){
				  rowContents = (String[])listIterate.next();
				  row = new  ArrayList<String>(Arrays.asList(rowContents));
				  csvData.add(row);
			  }
			  csvReader.close();
		  }catch (FileNotFoundException fne){
			  Common.addExceptionLogger(fne);
		  }catch (IOException ioe){
			  Common.addExceptionLogger(ioe);
		  }
		  return csvData;
	  }
	  
	  /**<p>Get CSV row data with given row number </p>
		 * @author <a href="mailto:ssirekumar@gmail.com">Ssire Kumar {@literal <ssirekumar@gmail.com>}</a> </br>
		 * @return {@code ArrayList<String>}
		 * @param  filePath - File path of the CSV file.
		 * @param  rowNumber - Row number.
	  * */
	  public static ArrayList<String> getCSVRow(String filePath, int rowNumber){
		  LoggerPoire.setLoggerObject(PoireCSVWorkbook._log);
		  String[] rowContents = null;
		  ArrayList<String> row = null;
		  int increment = 0;
		  try{
			  CSVReader csvReader = new CSVReader(new FileReader(filePath));
			  if (rowNumber > 0){
				  while (rowNumber != increment) {
					  rowContents = csvReader.readNext();
					  increment++;
				  }
				  row = new  ArrayList<String>(Arrays.asList(rowContents));
			  }else{
				  rowContents = csvReader.readNext();
				  row = new  ArrayList<String>(Arrays.asList(rowContents));
			  }
			  csvReader.close();
		  }catch (FileNotFoundException fne){
			  Common.addExceptionLogger(fne);
		  }catch (IOException ioe){
			  Common.addExceptionLogger(ioe);
		  }
		  return row;
	  }
	  
	  /**<p>Get CSV column data with given Column number </p>
		 * @author <a href="mailto:ssirekumar@gmail.com">Ssire Kumar {@literal <ssirekumar@gmail.com>}</a> </br>
		 * @return {@code ArrayList<String>}
		 * @param  filePath - File path of the CSV file.
		 * @param  colNumber - Column number.
	  * */
	  public static ArrayList<String> getCSVColumn(String filePath, int colNumber){
			ArrayList<String> colContents = null;
			String[] rowContents = null;
			List<String[]> contents = null;
			ListIterator<String[]> listIterate = null;
			try {
				CSVReader csvReader = new CSVReader(new FileReader(filePath));
				if (colNumber > 0) {
					contents = csvReader.readAll();
					listIterate = contents.listIterator();
					colContents = new ArrayList<String>();
					while (listIterate.hasNext()) {
						rowContents = (String[]) listIterate.next();
						colContents.add(rowContents[(colNumber - 1)]);
					}
				}
				csvReader.close();
			} catch (FileNotFoundException fne) {
				Common.addExceptionLogger(fne);
			} catch (IOException ioe) {
				Common.addExceptionLogger(ioe);
			}
			return colContents;
	  }
	  
	  /**<p>Get CSV column data As per the header name</p>
		 * @author <a href="mailto:ssirekumar@gmail.com">Ssire Kumar {@literal <ssirekumar@gmail.com>}</a> </br>
		 * @return {@code ArrayList<String>}
		 * @param  filePath - File path of the CSV file.
		 * @param  Header - Column header name.
	  * */
	  public static ArrayList<String> getCSVColumnPerHeader(String filePath, String header){
		  LoggerPoire.setLoggerObject(PoireCSVWorkbook._log);
		  ArrayList<String> colContents = null;
		  String[] rowContents = null;
		  List<String[]> contents = null;
		  ListIterator<String[]> listIterate = null;
		  String[] headers = null;
		  int colNumber = -1;
		  try {
			  CSVReader csvReader = new CSVReader(new FileReader(filePath));
			  headers = csvReader.readNext();
			  for (int i = 0; i < headers.length; i++) {
				  if (headers[i].equalsIgnoreCase(header)) {
					  colNumber = i;
					  break;
				  }
			  }
			  if (colNumber != -1) {
				  contents = csvReader.readAll();
				  listIterate = contents.listIterator();
				  colContents = new ArrayList<String>();
				  while (listIterate.hasNext()) {
					  rowContents = (String[]) listIterate.next();
					  colContents.add(rowContents[colNumber]);
				  }
			  }
			  csvReader.close();
		  } catch (FileNotFoundException fne) {
			  Common.addExceptionLogger(fne);
		  } catch (IOException ioe) {
			  Common.addExceptionLogger(ioe);
		  }
		  return colContents;
	  }
	  
	  /**<p>Get CSV row data As per the header name</p>
		 * @author <a href="mailto:ssirekumar@gmail.com">Ssire Kumar {@literal <ssirekumar@gmail.com>}</a> </br>
		 * @return {@code ArrayList<String>}
		 * @param  filePath - File path of the CSV file.
		 * @param  Header - Row header name.
	  * */
	  public static ArrayList<String> getCSVRowAsPerHeader(String filePath,int columnNumber, String Header) {
		  LoggerPoire.setLoggerObject(PoireCSVWorkbook._log);
		  ArrayList<String> rowContents = null;
		  LoggerPoire.log(Globals.Log.DEBUG,"Entered inside method getCSVRowAsPerHeader");
		  LoggerPoire.log(Globals.Log.DEBUG, "Going to read file : " + filePath);
		  ArrayList<String> columns = getCSVColumn(filePath, columnNumber);
		  LoggerPoire.log(Globals.Log.DEBUG, "Data present in Column Number : " + columnNumber + " is : " + columns);
		  int rowNumber = -1;
		  boolean found = false;
		  try {
			  for (String column : columns) {
				  rowNumber++;
				  if (column.equalsIgnoreCase(Header)) {
					  found = true;
					  break;
				  }
			  }
			  if (found) {
				  rowContents = getCSVRow(filePath, rowNumber + 1);
			  } else {
				  System.err.println("No Such Data Found");
				  return null;
			  }
			  System.out.println(rowContents);
		  } catch (NullPointerException npe) {
			  Common.addExceptionLogger(npe);
		  }
		  LoggerPoire.log(Globals.Log.DEBUG, "Row Data Returned as per Header : '" + Header + "' is : " + rowContents);
		  return rowContents;
	  }
	  
	  /**<p></p>
		 * @author <a href="mailto:ssirekumar@gmail.com">Ssire Kumar {@literal <ssirekumar@gmail.com>}</a> </br>
		 * @return {@code ArrayList<String>}
		 * @param  filePath - File path of the CSV file.
		 * @param  columnNumber - 
		 * @param  Header   - 
		 * @param  numberOfRows - .
	  * */
	  public static ArrayList<ArrayList<String>> getMultipleCSVRowsAsPerHeader(String filePath, int columnNumber, String Header, int numberOfRows) {
		  LoggerPoire.setLoggerObject(PoireCSVWorkbook._log);
		  ArrayList<ArrayList<String>> rowContents = new ArrayList<ArrayList<String>>();
		  LoggerPoire.log(Globals.Log.DEBUG,"Entered inside method getCSVRowAsPerHeader");
		  LoggerPoire.log(Globals.Log.DEBUG, "Going to read file : " + filePath);
		  ArrayList<String> columns = getCSVColumn(filePath, columnNumber);
		  LoggerPoire.log(Globals.Log.DEBUG, "Data present in Column Number : " + columnNumber + " is : " + columns);
		  int rowNumber = -1;
		  int count = 0;
		  boolean found = false;
		  try {
			  for (String column : columns) {
				  rowNumber++;
				  if (column.equalsIgnoreCase(Header)) {
					  found = true;
					  break;
				  }
			  }
			  if (found) {
				  while (count < numberOfRows) {
					  rowContents.add(getCSVRow(filePath, rowNumber + 1 + count));
					  count++;
				  }
			  } else {
				  System.out.println("No Such Data Found");
				  return null;
			  }
			  System.out.println(rowContents);
		  } catch (NullPointerException npe) {
			  System.err.println(npe.getMessage());
			  Common.addExceptionLogger(npe);
		  }
		  LoggerPoire.log(Globals.Log.DEBUG, "Row Data Returned as per Header : '" + Header + "' is : " + rowContents);
		  return rowContents;
	  }
	  
	  /**<p> </p>
		 * @author <a href="mailto:ssirekumar@gmail.com">Ssire Kumar {@literal <ssirekumar@gmail.com>}</a> </br>
		 * @param  filePath - File path of the CSV file.
		 * @param  headers - 
		 * @param  dataTypes   - 
		 * @param  totalRows - .
	  * */
	  public static void createRandomTestDataFile(String filePath, String[] headers, Class<?>[] dataTypes, int totalRows){
		  LoggerPoire.setLoggerObject(PoireCSVWorkbook._log);
		  try {
			  CSVWriter writer = new CSVWriter(new FileWriter(filePath));
			  ArrayList<String> row = null;
			  String[] rowContents = null;
			  if (headers.length == dataTypes.length) {
				  List<String[]> contents = new ArrayList<String[]>();
				  contents.add(headers);
				  while (totalRows > 0) {
					  row = new  ArrayList<String>();
					  for (Class<?> element : dataTypes) {
						  if (element == String.class) {
							  row.add(Common.getRandomString(11).toUpperCase());
						  } else if ((element == Long.class)
								  || (element == Integer.class)) {
							  row.add(Long.toString(Common.getRandomNumber(11)));
						  } else if (element == Date.class) {
							  row.add(Common.getDate("MM/dd/yyyy", 0, "days"));
						  }
					  }
					  rowContents = (String[]) row.toArray(new String[row.size()]);
					  contents.add(rowContents);
					  totalRows--;
					  row.clear();
				  }
				  if (contents != null) {
					  writer.writeAll(contents);
					  writer.close();
					  contents.clear();
				  }
			  } else {
				  System.err.println("Length Mismatch : Length of headers and DataTypes are not same");
				  LoggerPoire.log(Globals.Log.ERROR,"Length Mismatch : Length of headers and DataTypes are not same");
			  }
		  } catch (Exception e) {
			  Common.addExceptionLogger(e);
		  }
	  }
	  
	  /**<p> </p>
		 * @author <a href="mailto:ssirekumar@gmail.com">Ssire Kumar {@literal <ssirekumar@gmail.com>}</a> </br>
		 * @param  filePath - File path of the CSV file.
		 * @param  columnPositions - 
		 * @param  skipLines   - 
	  * */
	  public static void updateTestDataInFile(String filePath, int[] columnPositions, int skipLines){
		  LoggerPoire.setLoggerObject(PoireCSVWorkbook._log);
		  try {
			  CSVReader csvReader = new CSVReader(new FileReader(filePath));
			  ArrayList<String> randomArray = null;
			  List<String[]> updatedData = new ArrayList<String[]>();
			  while (skipLines > 0) {
				  updatedData.add(csvReader.readNext());
				  skipLines--;
			  }
			  List<String[]> contents = csvReader.readAll();
			  csvReader.close();
			  CSVWriter writer = new CSVWriter(new FileWriter(filePath));
			  for (String[] row : contents) {
				  randomArray = Common.getArrayOfRandomString(columnPositions.length, 11);
				  int increment = 0;
				  for (int pos : columnPositions) {
					  row[pos] = ((String) randomArray.get(increment));
					  increment++;
				  }
				  updatedData.add(row);
			  }
			  writer.writeAll(updatedData);
			  writer.close();
		  } catch (FileNotFoundException fne) {
			 Common.addExceptionLogger(fne);
		  } catch (IOException ioe) {
			  Common.addExceptionLogger(ioe);
		  }
	  }
	  
	  /**<p> </p>
		 * @author <a href="mailto:ssirekumar@gmail.com">Ssire Kumar {@literal <ssirekumar@gmail.com>}</a> </br>
		 * @return {@code ArrayList<String>}
		 * @param  filePath - File path of the CSV file.
		 * @param  columnPositions - 
		 * @param  skipLines   - 
	  * */
	  public static ArrayList<String> updateTestDataInFileWithConfig(String filePath, int[] columnPositions, int skipLines){
		  LoggerPoire.setLoggerObject(PoireCSVWorkbook._log);
		  LoggerPoire.log(Globals.Log.DEBUG,"Entered inside updateTestDataInFileWithConfig method");
		  ArrayList<String> randomDataUpdated = new ArrayList<String>();
		  String randomString = null;
		  String colPos = "";
		  boolean randomDataGenFlag = true;
		  int randomDataGenLen = 8;
		  LoggerPoire.log(Globals.Log.DEBUG, "Going to update file : " + filePath);
		  for (int colPons : columnPositions) {
			  colPos = colPos + colPons + "  ";
		  }
		  if (randomDataGenFlag) {
			  try {
				  CSVReader csvReader = new CSVReader(new FileReader(filePath));
				  List<String[]> updatedData = new ArrayList<String[]>();
				  while (skipLines > 0) {
					  updatedData.add(csvReader.readNext());
					  skipLines--;
				  }
				  List<String[]> contents = csvReader.readAll();
				  csvReader.close();
				  CSVWriter writer = new CSVWriter(new FileWriter(filePath));
				  for (String[] row : contents) {
					  randomString = Common.getRandomString(randomDataGenLen);
					  randomDataUpdated.add(randomString.toUpperCase());
					  for (int pos : columnPositions) {
						  if (row[pos].contains("$")) {
							  row[pos] = row[pos].replace("$", randomString.toUpperCase());
						  } else {
							  row[pos] = randomString.toUpperCase();
						  }
					  }
					  updatedData.add(row);
				  }
				  writer.writeAll(updatedData);
				  writer.close();
			  } catch (FileNotFoundException fne) {
				 Common.addExceptionLogger(fne);
			  } catch (IOException ioe) {
				 Common.addExceptionLogger(ioe);
			  }
		  }
		  return randomDataUpdated;
	  }
	  
	  public static void updateTestDataInFileWithConfig(String filePath, ArrayList<String> randomData, int[] colPos, int skipLines){ 
		  LoggerPoire.setLoggerObject(PoireCSVWorkbook._log);
		  if (true) {
			  String colPosStr = "";
			  String randomString = "";
			  int count = 0;
			  for (int colPons : colPos) {
				  colPosStr = colPosStr + colPons + "  ";
			  }
			  ArrayList<ArrayList<String>> entireData = getCSVData(filePath, skipLines);
			  if (entireData.size() == randomData.size()) {
				  try {
					  CSVReader csvReader = new CSVReader(new FileReader(filePath));
					  List<String[]> updatedData = new ArrayList<String[]>();
					  while (skipLines > 0) {
						  updatedData.add(csvReader.readNext());
						  skipLines--;
					  }
					  List<String[]> contents = csvReader.readAll();
					  csvReader.close();
					  CSVWriter writer = new CSVWriter(new FileWriter(filePath));
					  for (String[] row : contents) {
						  randomString = (String) randomData.get(count);
						  count++;
						  for (int pos : colPos) {
							  if (row[pos].contains("$")) {
								  row[pos] = row[pos].replace("$",
										  randomString.toUpperCase());
							  } else {
								  row[pos] = randomString.toUpperCase();
							  }
						  }
						  updatedData.add(row);
					  }
					  writer.writeAll(updatedData);
					  writer.close();
				  } catch (FileNotFoundException fne) {
					  Common.addExceptionLogger(fne);
				  } catch (IOException ioe) {
					  Common.addExceptionLogger(ioe);
				  }
			  } else {
				  System.err.println("No. of rows are not equal to Size of Random ArrayList Passed");
				  LoggerPoire.log(Globals.Log.ERROR,"No. of rows are not equal to Size of Random ArrayList Passed. Data not updated in file.");
			  }
		  }
	  }
	  
	  public static void updateTestDataInFileAsPerRowHeader(String filePath, int columnNumber, String Header, int[] columnPos, ArrayList<String> data){
		  LoggerPoire.setLoggerObject(PoireCSVWorkbook._log);
		  String pos = "";
		  String[] line = null;
		  String[] updatedLine = null;
		  for (int col : columnPos) {
			  pos = pos + col + " ";
		  }
		  LoggerPoire.log(Globals.Log.DEBUG, "TestDataEngine: Entered inside method updateTestDataInFileAsPerRowHeader method with Parameters : "
				  + filePath
				  + ", "
				  + columnNumber
				  + ", "
				  + Header + ", " + pos + ", " + data);
		  ArrayList<String> rowContents = null;
		  LoggerPoire.log(Globals.Log.DEBUG, "Going to read file : " + filePath);
		  ArrayList<String> columns = getCSVColumn(filePath, columnNumber);
		  LoggerPoire.log(Globals.Log.DEBUG, "Data present in Column Number : " + columnNumber + " is : " + columns);
		  int rowNumber = -1;
		  boolean found = false;
		  try {
			  for (String column : columns) {
				  rowNumber++;
				  if (column.equalsIgnoreCase(Header)) {
					  found = true;
					  break;
				  }
			  }
			  if (found) {
				  if (columnPos.length == data.size()) {
					  rowContents = getCSVRow(filePath, rowNumber + 1);
					  for (int i = 0; i < columnPos.length; i++) {
						  rowContents.set(columnPos[i], data.get(i));
					  }
					  CSVReader reader = new CSVReader(new FileReader(filePath));
					  ArrayList<String[]> updatedData = new ArrayList<String[]>();
					  while (rowNumber > 0) {
						  updatedData.add(reader.readNext());
						  rowNumber--;
					  }
					  updatedLine = new String[rowContents.size()];
					  updatedLine = (String[]) rowContents.toArray(updatedLine);
					  updatedData.add(updatedLine);
					  reader.readNext();
					  while ((line = reader.readNext()) != null) {
						  updatedData.add(line);
					  }
					  reader.close();
					  CSVWriter writer = new CSVWriter(new FileWriter(filePath));
					  writer.writeAll(updatedData);
					  writer.close();
				  } else {
					  System.out.println("MISMATCH : Length of Column Positions & Data to be updated");
					  LoggerPoire.log(Globals.Log.ERROR,"MISMATCH : Length of Column Positions & Data to be updated");
				  }
			  } else {
				  System.err.println("No Such Data Found");
				  LoggerPoire.log(Globals.Log.DEBUG, "Data with Header - " + Header + " is not present");
			  }
		  } catch (NullPointerException npe) {
			 Common.addExceptionLogger(npe);
		  } catch (ArrayIndexOutOfBoundsException aie) {
			  Common.addExceptionLogger(aie);
		  } catch (FileNotFoundException fnfe) {
			  Common.addExceptionLogger(fnfe);
		  } catch (IOException ioe) {
			  Common.addExceptionLogger(ioe);
		  }
		  LoggerPoire.log(Globals.Log.DEBUG,"TestDataEngine: Exiting method updateTestDataInFileAsPerRowHeader");
	  }
	      
	  public static ArrayList<Long> updateTestDataInFileInteger(String filePath, int[] columnPos, int skipLines, int length){
			LoggerPoire.setLoggerObject(PoireCSVWorkbook._log);
			String pos = "";
			for (int p : columnPos) {
				pos = pos + p + " ";
			}
			LoggerPoire.log(Globals.Log.DEBUG, "TestDataEngine: Entered inside method updateTestDataInFileInteger with Parameters : "
									+ filePath + ", " + pos + ", " + skipLines);
			ArrayList<Long> randomDataUpdated = new ArrayList<Long>();
			List<String[]> contents = null;List<String[]> updatedData = new ArrayList<String[]>();
			try {
				CSVReader reader = new CSVReader(new FileReader(filePath));
				CSVWriter writer = null;
				while (skipLines > 0) {
					updatedData.add(reader.readNext());
					skipLines--;
				}
				if (length > 0) {
					contents = reader.readAll();
					reader.close();
					for (String[] row : contents) {
						Long randomNumber = Long.valueOf(Common.getRandomNumber(length));
						for (int p : columnPos) {
							row[p] = String.valueOf(randomNumber);
						}
						updatedData.add(row);
						randomDataUpdated.add(randomNumber);
					}
					writer = new CSVWriter(new FileWriter(filePath));
					writer.writeAll(updatedData);
					writer.close();
					System.out.println("TestData file successfully updated!!");
					LoggerPoire.log(Globals.Log.DEBUG, "TestData file successfully updated!!");
				}
			} catch (FileNotFoundException fne) {
				Common.addExceptionLogger(fne);
			} catch (IOException ioe) {
				Common.addExceptionLogger(ioe);
			} catch (ArrayIndexOutOfBoundsException ae) {
				Common.addExceptionLogger(ae);
			} catch (NullPointerException npe) {
				Common.addExceptionLogger(npe);
			}
			LoggerPoire.log(Globals.Log.DEBUG, "TestDataEngine: Exiting method updateTestDataInFileInteger - Returned : " + randomDataUpdated);
			return randomDataUpdated;
	  }
	  
	  public static void updateTestDataInFileInteger(String filePath, int[] columnPos, ArrayList<Long> data, int skipLines){
		  LoggerPoire.setLoggerObject(PoireCSVWorkbook._log);
		  String pos = "";
		  int count = 0;
		  for (int p : columnPos) {
			  pos = pos + p + " ";
		  }
		  LoggerPoire
		  .log(Globals.Log.DEBUG,"TestDataEngine: Entered inside method updateTestDataInFileInteger with Parameters : "
				  + filePath
				  + ", "
				  + pos
				  + ", "
				  + data
				  + ", "
				  + skipLines);
		  LoggerPoire.log(Globals.Log.DEBUG, "TestDataEngine: Reading file " + filePath);
		  try {
			  List<String[]> updatedData = new ArrayList<String[]>();
			  List<String[]> contents = null;
			  CSVReader reader = new CSVReader(new FileReader(filePath));
			  while (skipLines > 0) {
				  updatedData.add(reader.readNext());
				  skipLines--;
			  }
			  contents = reader.readAll();
			  reader.close();
			  if (contents.size() == data.size()) {
				  for (String[] row : contents) {
					  for (int p : columnPos) {
						  row[p] = String.valueOf(data.get(count));
					  }
					  count++;
					  updatedData.add(row);
				  }
				  CSVWriter writer = new CSVWriter(new FileWriter(filePath));
				  writer.writeAll(updatedData);
				  writer.close();
				  System.out.println("TestData file successfully updated!!");
				  LoggerPoire.log(Globals.Log.DEBUG,"TestData file successfully updated!!");
			  } else {
				  System.err.println("MISMATCH: Number of rows in TestData File is not equal to number of elements to be updated!!");
				  LoggerPoire.log(Globals.Log.ERROR,"MISMATCH: Number of rows in TestData File is not equal to number of elements to be updated!!");
			  }
		  } catch (FileNotFoundException fne) {
			  Common.addExceptionLogger(fne);
		  } catch (IOException ioe) {
			 Common.addExceptionLogger(ioe);
		  } catch (NullPointerException npe) {
			  Common.addExceptionLogger(npe);
		  } catch (ArrayIndexOutOfBoundsException ae) {
			  Common.addExceptionLogger(ae);
		  }
		  LoggerPoire.log(Globals.Log.DEBUG,"TestDataEngine: Exiting method updateTestDataInFileInteger");
	  }
	  
	  public static void updateTestDataInFileRowWise(String filePath, int[] columnPositions, int rowNumber){
		  LoggerPoire.setLoggerObject(PoireCSVWorkbook._log);
		  try {
			  CSVReader csvReader = new CSVReader(new FileReader(filePath));
			  ArrayList<String> randomArray = null;
			  List<String[]> updatedData = new ArrayList<String[]>();
			  String[] row = null;
			  int increment = 0;
			  while (rowNumber > 1) {
				  updatedData.add(csvReader.readNext());
				  rowNumber--;
			  }
			  row = csvReader.readNext();
			  if (columnPositions.length <= row.length) {
				  randomArray = Common.getArrayOfRandomString(columnPositions.length, 8);
				  for (int pos : columnPositions) {
					  row[pos] = ((String) randomArray.get(increment));
					  increment++;
				  }
				  updatedData.add(row);
				  List<String[]> contents = csvReader.readAll();
				  csvReader.close();
				  CSVWriter writer = new CSVWriter(new FileWriter(filePath));
				  for (String[] rows : contents) {
					  updatedData.add(rows);
				  }
				  writer.writeAll(updatedData);
				  writer.close();
			  } else {
				  System.err.println("Column Postions mentioned doesn't exists in CSV file");
			  }
		  } catch (FileNotFoundException fne) {
			  Common.addExceptionLogger(fne);
		  } catch (IOException ioe) {
			  Common.addExceptionLogger(ioe);
		  }
	  }
	  
	  public static void updateTestDataInFileRowWise(String filePath, ArrayList<String> rowData, int[] columnPositions, int rowNumber){
		  try {
			  CSVReader csvReader = new CSVReader(new FileReader(filePath));
			  List<String[]> updatedData = new ArrayList<String[]>();
			  String[] row = null;
			  int increment = 0;
			  while (rowNumber > 1) {
				  updatedData.add(csvReader.readNext());
				  rowNumber--;
			  }
			  row = csvReader.readNext();
			  if (columnPositions.length == rowData.size()) {
				  for (int pos : columnPositions) {
					  row[pos] = ((String) rowData.get(increment));
					  increment++;
				  }
				  updatedData.add(row);
				  List<String[]> contents = csvReader.readAll();
				  csvReader.close();
				  CSVWriter writer = new CSVWriter(new FileWriter(filePath));
				  for (String[] rows : contents) {
					  updatedData.add(rows);
				  }
				  writer.writeAll(updatedData);
				  writer.close();
			  }
		  } catch (FileNotFoundException fne) {
			  Common.addExceptionLogger(fne);
		  } catch (IOException ioe) {
			  Common.addExceptionLogger(ioe);
		  }
	  }
	    
	  public static void writeTestDataInFile(String filePath, ArrayList<ArrayList<String>> contents, boolean appendMode){
		  LoggerPoire.setLoggerObject(PoireCSVWorkbook._log);
		  List<String[]> writeContents = new ArrayList<String[]>();
		  try {
			  for (ArrayList<String> row : contents) {
				  writeContents.add(row.toArray(new String[row.size()]));
			  }
			  CSVWriter csvWriter = new CSVWriter(new FileWriter(filePath,appendMode));
			  csvWriter.writeAll(writeContents);
			  csvWriter.close();
		  } catch (FileNotFoundException fne) {
			  Common.addExceptionLogger(fne);
		  } catch (IOException ioe) {
			  Common.addExceptionLogger(ioe);
		  } catch (NullPointerException npe) {
			  Common.addExceptionLogger(npe);
		  } catch (ArrayIndexOutOfBoundsException ae) {
			  Common.addExceptionLogger(ae);
		  }
	  }
	  
	 /* public static void writeTestDataInFileRowWise(String filePath, ArrayList<String> contents, int rowNumber){
		  LoggerPoire.setLoggerObject(TestDataEngine._log);
		  List<String[]> writeContents = new ArrayList<String[]>();
		  String[] rowContents = null;
		  ArrayList<String> arrayList = new ArrayList<String>();
		  int[] test = new int[100];
		  ArrayList<String> _empty = null;
		  int increment = 0;
		  try {
			  _empty = new ArrayList<String>();
			  ArrayList<String> _csvRow = TestDataEngine.getCSVRow(filePath, 0);
			  for (String cellVal : _csvRow) {
				    cellVal = "";
				  _empty.add(cellVal);
			  }
			  for (int i = 0; i < contents.size(); i++) {
				  test[i] = i;
			  }
			 
			  int[] abcd = Arrays.copyOf(test, contents.size());
			  System.out.println(new int[]{});
			  CSVReader csvReader = new CSVReader(new FileReader(filePath));
			  CSVWriter csvWriter = new CSVWriter(new FileWriter(filePath,true));
			  if (rowNumber > 0){
				  while (increment <= rowNumber) {
					  rowContents = csvReader.readNext();
					  if(rowContents!=null){
						  writeContents.add(increment, rowContents);
					  }else{
						  //writeContents.add(_empty.toArray(new String[_empty.size()]));
						  csvWriter.writeNext(_empty.toArray(new String[_empty.size()]));
						  //csvWriter.writeAll(writeContents);
					  }
					  if(increment == rowNumber){
						  break;
					  }
					  increment++;
				  }
				  if(increment == rowNumber){
					  //update data row
					  csvWriter.writeNext(_empty.toArray(new String[_empty.size()]));
					  TestDataEngine.updateTestDataInFileRowWise(filePath, contents, abcd, rowNumber+1);
				  }
				  csvWriter.close();
				  writeContents.add(contents.toArray(new String[contents.size()]));
				  csvWriter.writeAll(writeContents);
				  csvWriter.close();
			  }else{
				  CSVWriter csvWriter2 = new CSVWriter(new FileWriter(filePath,false));
				  writeContents.add(contents.toArray(new String[contents.size()]));
				  csvWriter2.writeAll(writeContents);
				  csvWriter2.close(); 
			  }
		  } catch (FileNotFoundException fne) {
			  System.err.println(fne.getMessage());
			 SeleCommon.addExceptionLogger(fne);
		  } catch (IOException ioe) {
			  System.err.println(ioe.getMessage());
			  SeleCommon.addExceptionLogger(ioe);
		  } catch (NullPointerException npe) {
			  System.err.println(npe.getMessage());
			  SeleCommon.addExceptionLogger(npe);
		  } catch (ArrayIndexOutOfBoundsException ae) {
			  System.err.println(ae.getMessage());
			  SeleCommon.addExceptionLogger(ae);
		  }
	  } */
	 
	  public static void createExcutionFilesForNewPackage(String packageName){
		  LoggerPoire.setLoggerObject(PoireCSVWorkbook._log);
		  CSVWriter writer = null;
		  String[] testSuitsHeader = { "Test Case Id (M)", "Description (O)",
				  "Complete Classname (M-CS)", "Method's Name (M-CS)",
				  "Parameter Types (M-IA)", "Parameter Values (M-IA)",
				  "Priority (M)", "GTS Case Ids" };
		  try {
			  File testSuits = new File("Execution/" + packageName.replace('.', '/') + "/TestCases.csv");
			  testSuits.getParentFile().mkdirs();
			  if (!testSuits.exists()) {
				  writer = new CSVWriter(new FileWriter("Execution/" + packageName.replace('.', '/') + "/TestCases.csv"));
				  writer.writeNext(testSuitsHeader);
				  writer.close();
			  } else {
				  System.err.println("TestSuits.csv File already exists for package : " + packageName);
			  }
			  File testCases = new File("Execution/" + packageName.replace('.', '/') + "/TestSuits.csv");
			  String[] testCasesHeader = new String[0];
			  if (!testCases.exists()) {
				  writer = new CSVWriter(new FileWriter("Execution/" + packageName.replace('.', '/') + "/TestSuits.csv"));
				  writer.writeNext(testCasesHeader);
				  writer.close();
			  } else {
				  System.err.println("TestCases.csv File already exists for package : " + packageName);
			  }
		  } catch (IOException ioe) {
			  Common.addExceptionLogger(ioe);
		  } catch (NullPointerException npe) {
			  Common.addExceptionLogger(npe);
		  }
	  }
	  
	  public static void createFolderStructureAsPerPackageName(String packageName){
		  LoggerPoire.setLoggerObject(PoireCSVWorkbook._log);
		  if (Common.doesJavaPackageExists(packageName)) {
			  try {
				  new File("TestData/" + packageName.replace('.', '/')).getCanonicalFile().mkdirs();
			  } catch (IOException ioe) {
				 System.err.println(ioe.getMessage());
				 Common.addExceptionLogger(ioe);
			  }
		  } else {
			  System.err.println("The package : " + packageName + " does not exists");
			  LoggerPoire.log(Globals.Log.ERROR, "The package : " + packageName + " does not exists");
		  }
	  }
}
