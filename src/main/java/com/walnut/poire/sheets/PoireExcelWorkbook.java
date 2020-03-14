/**This is a class will be useful for all the excel Operation on with java  
 * Programming. This is useful for any Data driven framework(DDT) of testing
 * 
 * **/
package com.walnut.poire.sheets;

import com.walnut.poire.Common;
import com.walnut.poire.Globals.Log;
import com.walnut.poire.LoggerPoire;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.math.BigInteger;
import java.security.SecureRandom;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.Date;
import java.util.HashSet;
import java.util.List;
import java.util.ListIterator;
import java.util.Random;
import java.util.Set;
import java.util.regex.Pattern;
import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbookFactory;

import com.opencsv.CSVReader;
import com.opencsv.CSVWriter;

/**
 * @author Siri Kumar</br>
 * @version 2.0.0</br> 
 * <p>The below class methods are Designed with the help of POI. These methods are useful for doing operations on the Excel
 * File formats. The whole class methods are in static only and these method are useful in all the excel file formats like 
 * .xls and .xlsx</p>
 * */

public class PoireExcelWorkbook {
	private static Logger _log = Logger.getLogger(PoireExcelWorkbook.class.getPackage().getName());
	private static SecureRandom random = new SecureRandom();
	private boolean xlsx = false;
	private File excelFile;
	
	
	/**
	 * Create PoireExcelWorkbook instance with specified excel file path.
	 * @param excelfilePath - Microsoft Excel file path
	 */
	public PoireExcelWorkbook(String excelfilePath) {
		this(new File(excelfilePath));
	}
	
	/**
	 * Create PoireExcelWorkbook instance with specified excel File.
	 * @param excelFile - File object of the excel file
	 */
	public PoireExcelWorkbook(File excelFile) {
		super();
		if(excelFile.isFile() && excelFile.exists()) {
			boolean xlsxb = PoireExcelWorkbook.excelFileType(excelFile).equalsIgnoreCase("xlsx");
			if(xlsxb) {
				this.excelFile = excelFile;
				this.xlsx = xlsxb;
			}else if(PoireExcelWorkbook.excelFileType(excelFile).equalsIgnoreCase("xls")){
				this.excelFile = excelFile;
			}else {
				this.excelFile = null;
			}
		}
	}
	
	/**
	 * Default PoireExcelWorkbook Constructor 
	 */
	public PoireExcelWorkbook() {
		super();
		this.xlsx = false;
		this.excelFile = null;
	}
	
	
	
	/**
	 * Set excel file path for PoireExcelWorkbook class object.
	 * @param excelfilePath - file path of the the excel.
	 */
	public void setExcelFilePath(String excelfilePath) {
		File excelFile = new File(excelfilePath);
		if(excelFile.isFile() && excelFile.exists()) {
			String sheetType = PoireExcelWorkbook.excelFileType(excelFile);
			if(!sheetType.isEmpty()) {
				this.excelFile = sheetType.equalsIgnoreCase("xlsx") ? excelFile : excelFile;
			}else {
				this.excelFile = null;
			}
		}
	}
	
	/**
	 * Set excel file path for PoireExcelWorkbook class object.
	 * @param excelPath - path of the excel directory
	 * @param fileName - excel file name.
	 * @throws IOException
	 * @throws FileNotFoundException
	 */

	public void setExcelFilePath(String excelPath, String fileName) throws IOException, FileNotFoundException{
		File excelFilePath = new File(excelPath);
		if(excelFilePath.isDirectory() && excelFilePath.exists()) {
			File actualExcelfilePath = new File(new String(excelFilePath.getCanonicalPath() + File.separator + fileName));
			if(!actualExcelfilePath.isFile()) {
				throw new FileNotFoundException(new StringBuffer()
						.append("File was not found on specified path ")
						.append(excelPath)
						.toString());
			}
			String sheetType = PoireExcelWorkbook.excelFileType(actualExcelfilePath);
			if(!sheetType.isEmpty()) {
				this.excelFile = sheetType.equalsIgnoreCase("xlsx") ? actualExcelfilePath : actualExcelfilePath;
			}else {
				this.excelFile = null;
			}
		}
	}
	/**
	 * Set excel file path for PoireExcelWorkbook class object.
	 * @param excelFile - File object of excel.
	 * @throws IOException
	 * @throws FileNotFoundException
	 */
	public void setExcelFilePath(File excelFile) throws IOException, FileNotFoundException{
		File excelFilePath = excelFile;
		if(excelFilePath.isFile() && excelFilePath.exists()) {
			String sheetType = PoireExcelWorkbook.excelFileType(excelFilePath);
			if(!sheetType.isEmpty()) {
				this.excelFile = sheetType.equalsIgnoreCase("xlsx") ? excelFilePath : excelFilePath;
			}else {
				this.excelFile = null;
			}
		}else {
			throw new FileNotFoundException(new StringBuffer()
					.append("File was not found on specified path ")
					.append(excelFilePath.getCanonicalPath())
					.toString());
		}
	}
	
	/**
	 * Set excel file path for PoireExcelWorkbook class object.
	 * @param excelFile - File object of excel path.
	 * @param fileName - excel file name
	 * @throws IOException
	 * @throws FileNotFoundException
	 */
	public void setExcelFilePath(File excelFile, String fileName) throws IOException, FileNotFoundException{
		File excelFilePath = excelFile;
		if(excelFilePath.isDirectory() && excelFilePath.exists()) {
			File actualExcelfilePath = new File(new String(excelFilePath.getCanonicalPath() + File.separator + fileName));
			String sheetType = PoireExcelWorkbook.excelFileType(actualExcelfilePath);
			if(!sheetType.isEmpty()) {
				this.excelFile = sheetType.equalsIgnoreCase("xlsx") ? actualExcelfilePath : actualExcelfilePath;
			}else {
				this.excelFile = null;
			}
		}else {
			throw new FileNotFoundException(new StringBuffer()
					.append("File was not found on specified path ")
					.append(excelFilePath.getCanonicalPath())
					.toString());
		}
	}
	
	/**
	 * Get the excel file path
	 * @return
	 * @throws IOException
	 */
	public String getExcelFilePath() throws IOException {
		String returnValue = "";   
		if (excelFile != null) {
			returnValue = excelFile.getCanonicalPath();
		}
		return returnValue;
	}
	
	private ArrayList<String> getColumnDataByIndex(Workbook wb, int colNumber, int sheetNumberIndex, boolean blankCells, int startRead) throws IllegalArgumentException {
		int iLastRowNumber = 0;
		ArrayList<String> colContents = new ArrayList<String>();
		Row row = null;
		Cell cell = null;
		if(((sheetNumberIndex >= 0) && (sheetNumberIndex <= wb.getNumberOfSheets()-1)) && (colNumber >= 0) ) {
			Sheet sh = wb.getSheetAt(sheetNumberIndex);
			if((iLastRowNumber = sh.getLastRowNum())>0) {
				for (int i=startRead; i<=iLastRowNumber; i++) {
					row = sh.getRow(i);
					if(row == null){
						if(blankCells){
							colContents.add("");
						}
						continue;
					}
					cell = row.getCell(colNumber, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
					colContents.add(PoireExcelWorkbook.cellFormatedDataValue(cell).toString());
					if(!blankCells){//Condition for blank cell data or nor with boolean value.
						if(cell.getCellType() == CellType.BLANK){
							colContents.remove(cell.toString());	
						}	
					}
				}
			}
		}else {
			throw new IllegalArgumentException(
					"Sheet number index should be (index >= 0), "
					+ "Column number should be (index >= 0), "
					+ "Given Sheet number is: " +sheetNumberIndex
			        + " Given Column number is: " +colNumber);
		}
		return colContents;
	}
	
    private ArrayList<String> getColumnDataByName(Workbook wb, int colNumber, String sheetName, boolean blankCells, int startRead) throws IllegalArgumentException{
    	int iLastRowNumber = 0;
		ArrayList<String> colContents = new ArrayList<String>();
		Row row = null;
		Cell cell = null;
		if(colNumber >= 0) {
			Sheet sh = wb.getSheet(sheetName);
			if((iLastRowNumber = sh.getLastRowNum())>0) {
				for (int i=startRead; i<=iLastRowNumber; i++) {
					row = sh.getRow(i);
					if(row == null){
						if(blankCells){
							colContents.add("");
						}
						continue;
					}
					cell = row.getCell(colNumber, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
					colContents.add(PoireExcelWorkbook.cellFormatedDataValue(cell).toString());
					if(!blankCells){//Condition for blank cell data or nor with boolean value.
						if(cell.getCellType() == CellType.BLANK){
							colContents.remove(cell.toString());	
						}	
					}
				}
			}
		}else {
			throw new IllegalArgumentException(
					"Column number should be (index >= 0), "
			        + " Given Column number is: " +colNumber);
		}
		return colContents;
	}

    private ArrayList<String> getColumnDataByName(Workbook wb, String columnHeaderName, String sheetName, boolean blankCells, int startRead) {
    	int iLastRowNumber = 0;
		ArrayList<String> colContents = new ArrayList<String>();
		Row row = null;
		Cell cell = null;
		Sheet sh = null;
		int bHeader = 0;
		boolean headerfound = false;
    	sh = wb.getSheet(sheetName);
		if(sh != null && (row = sh.getRow(0)) != null) {
			short iLastCellNumber = row.getLastCellNum();
			short iFirstCellNumber = row.getFirstCellNum();
			int iMiddleCellNumber = iFirstCellNumber +(iLastCellNumber - iFirstCellNumber)/2;
			int[] iFirstLastMiddleCelNumber = new int[]{iFirstCellNumber,iMiddleCellNumber,iLastCellNumber};
			for (int i : iFirstLastMiddleCelNumber) {
				cell = row.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
				String cellData = PoireExcelWorkbook.cellFormatedDataValue(cell);
				if(cellData.toString().equalsIgnoreCase(columnHeaderName)){
					bHeader = i;
					headerfound = true;
					break;
				}
			}
			if(!headerfound) {
				for (int j=startRead;j<iLastCellNumber; j++) {
					if(j== iFirstCellNumber || j== iMiddleCellNumber || j==iLastCellNumber) {continue;}
					cell = row.getCell(j, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
					String cellData = PoireExcelWorkbook.cellFormatedDataValue(cell);
					if(cellData.toString().equalsIgnoreCase(columnHeaderName)){
						bHeader = j;
						break;
					}
				}
			}
			colContents = getColumnDataByName(wb, bHeader, sheetName, blankCells,1);
		}else {
			throw new IllegalArgumentException(
					"Invalid sheet/row header name,"
			        + " Given sheet name as: " +sheetName
			        + " Given header name as: " +columnHeaderName);
		}
		return colContents;
    }
    
    private ArrayList<String> getExcelRowWithSheetIndex(Workbook wb, int rowNumber, int sheetNumberIndex, boolean blankCells, int startRead) throws IllegalArgumentException{
    	int iLastRowNumber = 0;
		ArrayList<String> colContents = new ArrayList<>();
		Row row = null;
		Cell cell = null;
		if(((sheetNumberIndex >= 0) && (sheetNumberIndex <= wb.getNumberOfSheets()-1)) && (rowNumber >= 0)) {
			Sheet sh = wb.getSheetAt(sheetNumberIndex);
			row = sh.getRow(rowNumber);
			if(row == null){
				return colContents;
			}else{
				short lastCellNumCount = row.getLastCellNum();
				for (int j=startRead;j<lastCellNumCount;j++) {
					cell = row.getCell(j, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
					colContents.add(PoireExcelWorkbook.cellFormatedDataValue(cell).toString());
					if(!blankCells){    //Condition for blank cell data or nor with boolean value.
						if(cell.getCellType() == CellType.BLANK){
							colContents.remove(cell.toString());	
						}	
					}
				}
			}
		}else {
			throw new IllegalArgumentException(
					"Sheet number index should be (index >= 0), "
					+ "Column number should be (index >= 0), "
					+ "Given Sheet number is: " +sheetNumberIndex
			        + " Given Column number is: " +rowNumber);
		}
		return colContents;
    }
    
    private ArrayList<String> getRowWithHeaderName(Workbook wb, int rowNumber, int sheetNumberIndex, boolean blankCells, int startRead) throws IllegalArgumentException{
    	int iLastRowNumber = 0;
		ArrayList<String> colContents = new ArrayList<>();
		Row row = null;
		Cell cell = null;
		if(((sheetNumberIndex >= 0) && (sheetNumberIndex <= wb.getNumberOfSheets()-1)) && (rowNumber >= 0)) {
			Sheet sh = wb.getSheetAt(sheetNumberIndex);
			row = sh.getRow(rowNumber);
			if(row == null){
				return colContents;
			}else{
				short lastCellNumCount = row.getLastCellNum();
				for (int j=startRead;j<lastCellNumCount;j++) {
					cell = row.getCell(j, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
					colContents.add(PoireExcelWorkbook.cellFormatedDataValue(cell).toString());
					if(!blankCells){    //Condition for blank cell data or nor with boolean value.
						if(cell.getCellType() == CellType.BLANK){
							colContents.remove(cell.toString());	
						}	
					}
				}
			}
		}else {
			throw new IllegalArgumentException(
					"Sheet number index should be (index >= 0), "
					+ "Column number should be (index >= 0), "
					+ "Given Sheet number is: " +sheetNumberIndex
			        + " Given Column number is: " +rowNumber);
		}
		return colContents;
    }
    
    
	/**
	 * <p>Get specified excel column data as a {@code java.util.ArrayList<String>} by it's sheet number index {index starts from 0}.</p>
	 * @author  <a href="mailto:ssirekumar@gmail.com">Siri Kumar {@literal <ssirekumar@gmail.com>}</a> </br>
	 * @return  Column data as a {@code java.util.ArrayList<String>} or null
	 * @param   colNumber - Column number of the sheet/workbook {index starts from 0}
	 * @param   sheetNumberIndex - Sheet/Workbook number {index starts from 0}
	 * @param   blankCells - collection of return data, its depends on true/false.
	 * 		    <p>if blankCells as true  Ex: {1,2,3, ,5, , , ,5,7} outPut would be {1,2,3, ,5, , , ,5,7},</br>
	 *          if blankCells as false Ex: {1,2,3, ,5, , , ,5,7} outPut would be {1,2,3,5,5,7}</p>
	 **/
	public ArrayList<String> getExcelColumnWithSheetIndex(int colNumber, int sheetNumberIndex, boolean blankCells){
		ArrayList<String> colContents = new ArrayList<String>();
		File sFilePathObj = null;
		OPCPackage _OPCXSSFW = null;
		POIFSFileSystem  _NFSHSSFW = null;
		Row row = null;
		Cell cell = null;
		boolean bStreenClose = false;
		int iLastRowNumber;
		try {
			if(excelFile != null) {
				sFilePathObj = excelFile.getCanonicalFile();
			}else {
				return colContents;
			}
			if(xlsx){
				_OPCXSSFW = OPCPackage.open(sFilePathObj);
			}else{
				_NFSHSSFW = new POIFSFileSystem(sFilePathObj);
			}
			Workbook wb = xlsx ? XSSFWorkbookFactory.create(_OPCXSSFW) : WorkbookFactory.create(_NFSHSSFW);
			colContents = getColumnDataByIndex(wb, colNumber, sheetNumberIndex, blankCells,0);
		}catch (IllegalArgumentException | InvalidFormatException | IOException e) {
			Common.addExceptionLogger(e);
			if(colContents.size()>0) {
				colContents.removeAll(colContents);
			}
		}finally{
			if(xlsx){
				try {
					if(_OPCXSSFW != null){
						_OPCXSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				} 
			}else{
				try {
					if(_NFSHSSFW != null){
						_NFSHSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		return colContents;
	}
     
	/**<p>Get excel column data as a {@code java.util.ArrayList<String>} by it's sheet name. 
	 * @author <a href="mailto:ssirekumar@gmail.com">Siri Kumar {@literal <ssirekumar@gmail.com>}</a> </br>
	 * @return {@code ArrayList<String>}
	 * @param  colNumber  - Column number of the sheet/workbook {index starts from 0}
	 * @param  sheetName  - Name of the sheet/workbook
	 * @param  blankCells - Collection of return data, its depends on true/false.
	 * 		   <p>if blankCells as true  Ex: {1,2,3, ,5, , , ,5,7} output is {1,2,3, ,5, , , ,5,7}<br>
	 *         if blankCells as false Ex: {1,2,3, ,5, , , ,5,7} output is {1,2,3,5,5,7}.
	 *  */
	public ArrayList<String> getExcelColumnWithSheetName(int colNumber, String sheetName, boolean blankCells){
		ArrayList<String> colContents = null;
		File sFilePathObj = null;
		OPCPackage _OPCXSSFW = null;
		POIFSFileSystem  _NFSHSSFW = null;
		Row row = null;
		Cell cell = null;
		boolean bStreenClose = false;
		int iLastRowNumber = 0;
		try {
			if(excelFile != null) {
				sFilePathObj = excelFile.getCanonicalFile();
			}else {
				return colContents;
			}
			if(xlsx){
				_OPCXSSFW = OPCPackage.open(sFilePathObj);
			}else{
				_NFSHSSFW = new POIFSFileSystem(sFilePathObj);
			}
			Workbook wb = xlsx ? XSSFWorkbookFactory.create(_OPCXSSFW) : WorkbookFactory.create(_NFSHSSFW);
			colContents = getColumnDataByName(wb, colNumber, sheetName, blankCells,0);
		}catch(IllegalArgumentException | InvalidFormatException | IOException e){
			Common.addExceptionLogger(e);
			if(colContents.size()>0) {
				colContents.removeAll(colContents);
			}
		}finally{
			if(xlsx){
				try {
					if(_OPCXSSFW != null){
						_OPCXSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				} 
			}else{
				try {
					if(_NFSHSSFW != null){
						_NFSHSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		return colContents;
	}
	
	/**<p>Get excel named header column data as a {@code java.util.ArrayList<String>} by it's sheet name.</p> 
	 * @author <a href="mailto:ssirekumar@gmail.com">Siri Kumar {@literal <ssirekumar@gmail.com>}</a> </br>
	 * @return {@code java.util.ArrayList<String>}
	 * @param  columnHeaderName    -  Name of the column header name(By default first row as a header names).
	 * @param  sheetName  - Name of the sheet/workbook
	 * @param  blankCells - Read data including blank cells with flag true/false.
	 * 		   <p>if blankCells as true  Ex: {1,2,3, ,5, , , ,5,7} output is {1,2,3, ,5, , , ,5,7}<br>
	 *          if blankCells as false Ex: {1,2,3, ,5, , , ,5,7} output is {1,2,3,5,5,7}</p> 
	 * 
	 * 
	 *  */
	public ArrayList<String> getExcelColumnWithHeaderName(String columnHeaderName, String sheetName, boolean blankCells){
		ArrayList<String> colContents = new ArrayList<String>();
		File sFilePathObj = null;
		OPCPackage _OPCXSSFW = null;
		POIFSFileSystem  _NFSHSSFW = null;
		Row row = null;
		Cell cell = null;
		Sheet sh = null;
		boolean bStreenClose = false;
		boolean headerfound = false;
		int bHeader = 0;
		try {
			if(excelFile != null) {
				sFilePathObj = excelFile.getCanonicalFile();
			}else {
				return colContents;
			}
			if(xlsx){
				_OPCXSSFW = OPCPackage.open(sFilePathObj);
			}else{
				_NFSHSSFW = new POIFSFileSystem(sFilePathObj);
			}
			Workbook wb = xlsx ? XSSFWorkbookFactory.create(_OPCXSSFW) : WorkbookFactory.create(_NFSHSSFW);
			colContents = getColumnDataByName(wb, columnHeaderName, sheetName, blankCells,1);
		}catch(IllegalArgumentException | InvalidFormatException | IOException e){
			Common.addExceptionLogger(e);
			if(colContents.size()>0) {
				colContents.removeAll(colContents);
			}
		} finally{
			if(xlsx){
				try {
					if(_OPCXSSFW != null){
						_OPCXSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				} 
			}else{
				try {
					if(_NFSHSSFW != null){
						_NFSHSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		return colContents;
	}

	/**<p>Get the excel column data as a {@code java.util.ArrayList<String>} by it's column header name and its sheet number. 
	 * @author <a href="mailto:ssirekumar@gmail.com">Siri Kumar {@literal <ssirekumar@gmail.com>}</a> </br>
	 * @return {@code java.util.ArrayList<String>}
	 * @param  columnHeaderName    - Column Header name (By default first row as a header names).
	 * @param  sheetNumberIndex  - Index number of the sheet {index starts from 0}
	 * @param  blankCells - Method which return the data as blank cells the blankCells value as true.
	 * 	<p>if blankCells as true  Ex: {1,2,3, ,5, , , ,5,7} output is {1,2,3, ,5, , , ,5,7}<br>
	 *  if blankCells as false Ex: {1,2,3, ,5, , , ,5,7} output is {1,2,3,5,5,7}</p>
	 *  */
	public ArrayList<String> getExcelColumnWithHeaderSheetIndex(String columnHeaderName, int sheetNumberIndex, boolean blankCells){
		ArrayList<String> colContents = new ArrayList<String>();
		File sFilePathObj = null;
		OPCPackage _OPCXSSFW = null;
		POIFSFileSystem  _NFSHSSFW = null;
		Row row = null;
		Cell cell = null;
		boolean bStreenClose = false;
		int bHeader = 0;
		try {
			if(excelFile != null) {
				sFilePathObj = excelFile.getCanonicalFile();
			}else {
				return colContents;
			}
			if(xlsx){
				_OPCXSSFW = OPCPackage.open(sFilePathObj);
			}else{
				_NFSHSSFW = new POIFSFileSystem(sFilePathObj);
			}
			Workbook wb = xlsx ? XSSFWorkbookFactory.create(_OPCXSSFW) : WorkbookFactory.create(_NFSHSSFW);
			Sheet sh = wb.getSheetAt(sheetNumberIndex);
			colContents = getColumnDataByName(wb, columnHeaderName, sh.getSheetName(), blankCells,1);
		}catch(IllegalArgumentException | InvalidFormatException | IOException e){
			Common.addExceptionLogger(e);
			if(colContents.size()>0) {
				colContents.removeAll(colContents);
			}
		}finally{
			if(xlsx){
				try {
					if(_OPCXSSFW != null){
						_OPCXSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				} 
			}else{
				try {
					if(_NFSHSSFW != null){
						_NFSHSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		return colContents;
	}
		
	/**<p>Get the excel row data as a {@code java.util.ArrayList<String>} by it's sheet number index. 
	 * @author <a href="mailto:ssirekumar@gmail.com">Siri Kumar {@literal <ssirekumar@gmail.com>}</a> </br>
	 * @return return row values as {@code java.util.ArrayList<String>} of each cell, or empty if row doesn't have data.
	 * @param  rowNumber - integer number of the row {index start from 0}.
	 * @param  sheetNumberIndex  - excel sheet number.
	 * @param  blankCells  - read data including blank cells with flag true/false. <br>
	 *         <p>if blankCells as true  Ex: {1,2,3, ,5, , , ,5,7} OutPut is {1,2,3, ,5, , , ,5,7}<br>
	 *         if blankCells as false Ex: {1,2,3, ,5, , , ,5,7} OutPut is {1,2,3,5,5,7}</p>
	 *  */
	public ArrayList<String> getExcelRowWithSheetIndex(int rowNumber, int sheetNumberIndex, boolean blankCells){
		ArrayList<String> colContents = new ArrayList<>();
		File sFilePathObj = null;
		OPCPackage _OPCXSSFW = null;
		POIFSFileSystem  _NFSHSSFW = null;
		Row row = null;
		Cell cell = null;
		boolean bStreenClose = false;
		try {
			if(excelFile != null) {
				sFilePathObj = excelFile.getCanonicalFile();
			}else {
				return colContents;
			}
			if(xlsx){
				_OPCXSSFW = OPCPackage.open(sFilePathObj);
			}else{
				_NFSHSSFW = new POIFSFileSystem(sFilePathObj);
			}
			Workbook wb = xlsx ? XSSFWorkbookFactory.create(_OPCXSSFW) : WorkbookFactory.create(_NFSHSSFW);
			colContents = getExcelRowWithSheetIndex(wb, rowNumber, sheetNumberIndex, blankCells, 0);
		}catch(IllegalArgumentException | InvalidFormatException | IOException e){
			Common.addExceptionLogger(e);
			if(colContents.size()>0) {
				colContents.removeAll(colContents);
			}
		}finally{
			if(xlsx){
				try {
					if(_OPCXSSFW != null){
						_OPCXSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				} 
			}else{
				try {
					if(_NFSHSSFW != null){
						_NFSHSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		return colContents;
	 }
    
	/**<p>It will used to get the Excel row data as a {@code ArrayList<String>} With specified row Number and its sheet Name.
	 * @author <a href="mailto:ssirekumar@gmail.com">Siri Kumar {@literal <ssirekumar@gmail.com>}</a> </br>
	 * @return {@code ArrayList<String>}
	 * @param  filePath  - File path of the .xls or .xlsx file.
	 * @param  rowNumber - integer number of the row start from 0 index.
	 * @param  sheetName  - Excel Sheet Name.
	 * @param  blankCells - Method which return the data as blank cells if the blankCells value as true.<br>
	 *                       if blankCells as true  Ex{1,2,3, ,5, , , ,5,7} OutPut is {1,2,3, ,5, , , ,5,7}<br>
	 *                       if blankCells as false Ex{1,2,3, ,5, , , ,5,7} OutPut is {1,2,3,5,5,7}
	 * 
	 *                       
	 *  */
	public ArrayList<String> getExcelRowWithSheetName(String filePath, int rowNumber, String sheetName, boolean blankCells){
		 ArrayList<String> colContents = null;
			File sFilePathObj = null;
			OPCPackage _OPCXSSFW = null;
			POIFSFileSystem  _NFSHSSFW = null;
			Row row = null;
			Cell cell = null;
			boolean bStreenClose = false;
			try {
				sFilePathObj = new File(filePath);
				xlsx = PoireExcelWorkbook.excelFileType(sFilePathObj).equalsIgnoreCase("xlsx");
				if(xlsx){
					_OPCXSSFW = OPCPackage.open(sFilePathObj);
				}else{
					_NFSHSSFW = new POIFSFileSystem(sFilePathObj);
				}
				Workbook wb = xlsx ? XSSFWorkbookFactory.create(_OPCXSSFW) : WorkbookFactory.create(_NFSHSSFW);
				Sheet sh = wb.getSheet(sheetName);
				colContents = new ArrayList<String>();
				row = sh.getRow(rowNumber);
				if(row == null){
					colContents.add(0,"");
					return colContents;
				}else{
					short lastCellNumCount = row.getLastCellNum();
					for (int j=0;j<=lastCellNumCount;j++) {
						cell = row.getCell(j, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
						colContents.add(PoireExcelWorkbook.cellFormatedDataValue(cell).toString());
						if(!blankCells){    //Condition for blank cell data or nor with boolean value.
							if(cell.getCellType() == CellType.BLANK){
								colContents.remove(cell.toString());	
							}	
						}
					}
				}
			}catch(FileNotFoundException fe){
				Common.addExceptionLogger(fe);
			} catch (Exception e) {
				Common.addExceptionLogger(e);
			}finally{
				if(xlsx){
					try {
						if(_OPCXSSFW != null){
							_OPCXSSFW.close();
						}
					} catch (IOException e) {
						e.printStackTrace();
					} 
				}else{
					try {
						if(_NFSHSSFW != null){
							_NFSHSSFW.close();
						}
					} catch (IOException e) {
						e.printStackTrace();
					}
				}
			}
			return colContents;
	 }
	
	/**<p>It will used to get the Excel row data as a {@code ArrayList<String>} With specified row Header and its sheet Name.
	 * @author <a href="mailto:ssirekumar@gmail.com">Siri Kumar {@literal <ssirekumar@gmail.com>}</a> </br>
	 * @return {@code ArrayList<String>}
	 * @param  filePath  - File path of the .xls or .xlsx file.
	 * @param  rowHeaderName - Name of the string value which is in index 0  column(by default it will take column values at first index 0) .
	 * @param  sheetName  - Excel Sheet Name.
	 * @param  blankCells - Method which return the data as blank cells if the blankCells value as true.<br>
	 *                       if blankCells as true  Ex{1,2,3, ,5, , , ,5,7} OutPut is {1,2,3, ,5, , , ,5,7}<br>
	 *                       if blankCells as false Ex{1,2,3, ,5, , , ,5,7} OutPut is {1,2,3,5,5,7}
	 *  */
	public static ArrayList<String> getExcelRowWithHeaderName(String filePath, String rowHeaderName, String sheetName, boolean blankCells){
		 ArrayList<String> colContents = null;
			File sFilePathObj = null;
			OPCPackage _OPCXSSFW = null;
			POIFSFileSystem  _NFSHSSFW = null;
			Row row = null;
			Cell cell = null;
			boolean bStreenClose = false;
			int bHeader = 0, loopValue = 0;
			try {
				ArrayList<String> sFirstIndexColValues = PoireExcelWorkbook.getExcelColumnWithSheetName(filePath, 0, sheetName, true);
				if(sFirstIndexColValues.size()>0 || sFirstIndexColValues != null){
					for (String zeroIndexcellValue : sFirstIndexColValues) {
						if(zeroIndexcellValue.equalsIgnoreCase(rowHeaderName)){
							bHeader = loopValue;
							break;
						}
						loopValue++;
					}
				}
				sFilePathObj = new File(filePath);
				xlsx = PoireExcelWorkbook.excelFileType(sFilePathObj).equalsIgnoreCase("xlsx");
				if(xlsx){
					_OPCXSSFW = OPCPackage.open(sFilePathObj);
				}else{
					_NFSHSSFW = new POIFSFileSystem(sFilePathObj);
				}
				Workbook wb = xlsx ? XSSFWorkbookFactory.create(_OPCXSSFW) : WorkbookFactory.create(_NFSHSSFW);
				Sheet sh = wb.getSheet(sheetName);
				colContents = new ArrayList<String>();
				row = sh.getRow(bHeader);
				if(row == null){
					colContents.add(0,"");
					return colContents;
				}else{
					short lastCellNumCount = row.getLastCellNum();
					for (int j=1;j<=lastCellNumCount;j++) {
						cell = row.getCell(j, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
						colContents.add(cell.toString());
						if(!blankCells){    //Condition for blank cell data or nor with boolean value.
							if(cell.getCellType() == CellType.BLANK){
								colContents.remove(cell.toString());	
							}	
						}
					}
				}
			}catch(FileNotFoundException fe){
				Common.addExceptionLogger(fe);
			} catch (Exception e) {
				Common.addExceptionLogger(e);
			}finally{
				if(xlsx){
					try {
						if(_OPCXSSFW != null){
							_OPCXSSFW.close();
						}
					} catch (IOException e) {
						e.printStackTrace();
					} 
				}else{
					try {
						if(_NFSHSSFW != null){
							_NFSHSSFW.close();
						}
					} catch (IOException e) {
						e.printStackTrace();
					}
				}
			}
			return colContents;
	 }
	
	/**<p>Get excel row data as a {@code java.util.ArrayList<String>} by it's headername & sheet number index
	 * @author <a href="mailto:ssirekumar@gmail.com">Siri Kumar {@literal <ssirekumar@gmail.com>}</a> </br>
	 * @return {@code java.util.ArrayList<String>}
	 * @param  rowHeaderName - name of the row header name (by default it will take column values at first index 0)
	 * @param  sheetNumberIndex  - excel sheet number {index from the 0}
	 * @param  blankCells - Method which return the data as blank cells if the blankCells value as true.<br>
	 *         <p>if blankCells as true  Ex: {1,2,3, ,5, , , ,5,7} output is {1,2,3, ,5, , , ,5,7}<br>
	 *         if blankCells as false Ex: {1,2,3, ,5, , , ,5,7} output is {1,2,3,5,5,7}
	 *  */
	public ArrayList<String> getExcelRowWithHeaderName(String rowHeaderName, int sheetNumberIndex, boolean blankCells){
		ArrayList<String> colContents = new ArrayList<>();
		File sFilePathObj = null;
		OPCPackage _OPCXSSFW = null;
		POIFSFileSystem  _NFSHSSFW = null;
		Row row = null;
		Cell cell = null;
		boolean bStreenClose = false;
		int bHeader = 0, loopValue = 0;
		try {
			ArrayList<String> sFirstIndexColValues = this.getExcelColumnWithSheetIndex(0, sheetNumberIndex, true);
			if(sFirstIndexColValues.size()>0 || sFirstIndexColValues != null){
				for (String zeroIndexcellValue : sFirstIndexColValues) {
					if(zeroIndexcellValue.equalsIgnoreCase(rowHeaderName)){
						bHeader = loopValue;
						break;
					}
					loopValue++;
				}
			}
			if(excelFile != null) {
				sFilePathObj = excelFile.getCanonicalFile();
			}else {
				return colContents;
			}
			if(xlsx){
				_OPCXSSFW = OPCPackage.open(sFilePathObj);
			}else{
				_NFSHSSFW = new POIFSFileSystem(sFilePathObj);
			}
			Workbook wb = xlsx ? XSSFWorkbookFactory.create(_OPCXSSFW) : WorkbookFactory.create(_NFSHSSFW);
			Sheet sh = wb.getSheetAt(sheetNumberIndex);
			colContents = new ArrayList<String>();
			row = sh.getRow(bHeader);
			if(row == null){
				colContents.add(0,"");
				return colContents;
			}else{
				short lastCellNumCount = row.getLastCellNum();
				for (int j=1;j<=lastCellNumCount;j++) {
					cell = row.getCell(j, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
					colContents.add(cell.toString());
					if(!blankCells){    //Condition for blank cell data or nor with boolean value.
						if(cell.getCellType() == CellType.BLANK){
							colContents.remove(cell.toString());	
						}	
					}
				}
			}
		}catch(FileNotFoundException fe){
			Common.addExceptionLogger(fe);
		} catch (Exception e) {
			Common.addExceptionLogger(e);
		}finally{
			if(xlsx){
				try {
					if(_OPCXSSFW != null){
						_OPCXSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				} 
			}else{
				try {
					if(_NFSHSSFW != null){
						_NFSHSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		return colContents;
	}
	
	/**<p>It will used to Update random data in the existing Excel row with specified columnPositions values.</p>
	 * @author <a href="mailto:ssirekumar@gmail.com">Siri Kumar {@literal <ssirekumar@gmail.com>}</a> </br>
	 * @return boolean
	 * @param  filePath  - File path of the .xls or .xlsx file.
	 * @param  columnPositions - Array of Column index values which is starts from 0. Exp new int{0,2,4,5,8}
	 * @param  skipLines  - How many lines to skip for updating the row with column index.
	 * @param  sheetNumberIndex - Sheet number index it will start from 0.
	 * @param  lengthOfRandomString - integer value for length of the random string. 
	 * @param  randomNumberString   - boolean true/false for only random number if its true or else its random string only.
	 * 
	 * */
	public static boolean updateRandomDataInExcelFile(String filePath, int[] columnPositions, int skipLines, int sheetNumberIndex, int lengthOfRandomString, boolean randomNumberString){
		File sFilePathObj = null;
		OPCPackage _OPCXSSFW = null;
		POIFSFileSystem  _NFSHSSFW = null;
		Row row = null;
		Row cellCountRow = null;
		Cell cell = null;
		FileOutputStream fileOut = null;
		boolean _sReturnObj = false;
		boolean bStreenClose = false;
		boolean lengthOfCell = false;
		ArrayList<String> randomArray = null;
		try {
			sFilePathObj = new File(filePath);
			xlsx = PoireExcelWorkbook.excelFileType(sFilePathObj).equalsIgnoreCase("xlsx");
			if(xlsx){
				_OPCXSSFW = OPCPackage.open(new FileInputStream(sFilePathObj));
			}else{
				_NFSHSSFW = new POIFSFileSystem(new FileInputStream(sFilePathObj));
			}
			Workbook wb = xlsx ? XSSFWorkbookFactory.create(_OPCXSSFW) : WorkbookFactory.create(_NFSHSSFW);
			Sheet sh = wb.getSheetAt(sheetNumberIndex);
			int dataCellNum = 0;
			if(sh.getLastRowNum()>skipLines){
				/*int loopcount = skipLines; 
				for(int rowCellv = 0;loopcount>=rowCellv;loopcount--){
					cellCountRow = sh.getRow(loopcount);
				    ArrayList<String> rowOfExcel = PoireExcelWorkbook.getExcelRowWithSheetIndex(filePath, 0, sheetNumberIndex, true);
					if(rowOfExcel == null || rowOfExcel.size() == 0){
						continue;
					}else{
						dataCellNum = rowOfExcel.size();//cellCountRow.getLastCellNum();
						break;
					}
				}*/
				cellCountRow = sh.getRow(0);
				dataCellNum = cellCountRow.getLastCellNum();
				int totalSkiplines = skipLines+1;
				row = sh.getRow(totalSkiplines);
				if(row == null){
					row = sh.createRow(totalSkiplines);
					for (int k=0;k<dataCellNum;k++) {
						row.createCell(k, CellType.BLANK);
					}
					if(randomNumberString){
						randomArray = PoireExcelWorkbook.getArrayOfRandomNumber(columnPositions.length, lengthOfRandomString);
					}else{
						randomArray = PoireExcelWorkbook.getArrayOfRandomString(columnPositions.length, lengthOfRandomString);
					}
					for (int indexArr : columnPositions) {
						if(indexArr>=dataCellNum){
							lengthOfCell = true;
						}
					}
					int increment = 0;
                    if(!lengthOfCell){
                    	for (int pos : columnPositions) {
    						cell = row.createCell(pos, CellType.STRING);
    						cell.setCellValue((String) randomArray.get(increment));
    						increment++;
    					}
    					fileOut = new FileOutputStream(filePath);
    					wb.write(fileOut);
    					fileOut.flush();
    					fileOut.close();
    					bStreenClose = true;
    					_sReturnObj = true;
                    }else{
                    	System.err.println("WARRING: Column Positions index values are greater then the last Data Cell index");
						System.out.println("\tPlease change the Column Positions index with appropreate Data cells index");
                    }
				}else{
					//row.setRowNum(totalSkiplines);
					if(randomNumberString){
						randomArray = PoireExcelWorkbook.getArrayOfRandomNumber(columnPositions.length, lengthOfRandomString);
					}else{
						randomArray = PoireExcelWorkbook.getArrayOfRandomString(columnPositions.length, lengthOfRandomString);
					}
					for (int indexArr : columnPositions) {
						if(indexArr>=dataCellNum){
							lengthOfCell = true;
						}
					}
					int increment = 0;
					if(!lengthOfCell){
						for (int pos : columnPositions) {
							cell = row.createCell(pos, CellType.STRING);
							cell.setCellValue((String) randomArray.get(increment));
							increment++;
						}
						fileOut = new FileOutputStream(filePath);
						wb.write(fileOut);
						fileOut.flush();
						fileOut.close();
						bStreenClose = true;
						_sReturnObj = true;
					}else{
						System.err.println("WARRING: Column Positions index values are greater then the last Data Cell index");
						System.out.println("\tPlease change the Column Positions index with appropreate Data cells index");
					}
				}
			}else{
				_sReturnObj = false;
				System.err.println("SkipLines value is greter than the Data present in the sheet with index:"+ sheetNumberIndex);
			}
		}catch(FileNotFoundException fe){
			Common.addExceptionLogger(fe);
		} catch (Exception e) {
			Common.addExceptionLogger(e);
		}finally{
			if(xlsx){
				try {
					if(_OPCXSSFW != null){
						_OPCXSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				} 
			}else{
				try {
					if(_NFSHSSFW != null){
						_NFSHSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		return _sReturnObj;		
	}
	
	/**<p>It will used to Update the existing Excel row data with columnPositions values.
	 * @author <a href="mailto:ssirekumar@gmail.com">Siri Kumar {@literal <ssirekumar@gmail.com>}</a> </br>
	 * @return boolean
	 * @param  filePath  - File path of the .xls or .xlsx file.
	 * @param  columnPositions - Array of Column index values which is starts from 0. Exp new int{0,2,4,5,8}
	 * @param  skipLines  - How many lines to skip for updating the row with column index.
	 * @param  sheetName - Name of the sheet.
	 * @param  lengthOfRandomString - integer value for length of the random string. 
	 * @param  randomNumberString   - boolean true/false for only random number if its true or else its random string only.
	 * 
	 * 
	 * */
	public static boolean updateRandomDataInExcelFile(String filePath, int[] columnPositions, int skipLines, String sheetName, int lengthOfRandomString, boolean randomNumberString){
		File sFilePathObj = null;
		OPCPackage _OPCXSSFW = null;
		POIFSFileSystem  _NFSHSSFW = null;
		Row row = null;
		Row cellCountRow = null;
		Cell cell = null;
		FileOutputStream fileOut = null;
		boolean _sReturnObj = false;
		boolean bStreenClose = false;
		boolean lengthOfCell = false;
		ArrayList<String> randomArray = null;
		try {
			sFilePathObj = new File(filePath);
			xlsx = PoireExcelWorkbook.excelFileType(sFilePathObj).equalsIgnoreCase("xlsx");
			if(xlsx){
				_OPCXSSFW = OPCPackage.open(new FileInputStream(sFilePathObj));
			}else{
				_NFSHSSFW = new POIFSFileSystem(new FileInputStream(sFilePathObj));
			}
			Workbook wb = xlsx ? XSSFWorkbookFactory.create(_OPCXSSFW) : WorkbookFactory.create(_NFSHSSFW);
			Sheet sh = wb.getSheet(sheetName);
			int dataCellNum = 0;
			if(sh.getLastRowNum()>skipLines){
				/*int loopcount = skipLines; 
				for(int rowCellv = 0;loopcount>=rowCellv;loopcount--){
					cellCountRow = sh.getRow(loopcount);
				    ArrayList<String> rowOfExcel = PoireExcelWorkbook.getExcelRowWithSheetIndex(filePath, 0, sheetNumberIndex, true);
					if(rowOfExcel == null || rowOfExcel.size() == 0){
						continue;
					}else{
						dataCellNum = rowOfExcel.size();//cellCountRow.getLastCellNum();
						break;
					}
				}*/
				cellCountRow = sh.getRow(0);
				dataCellNum = cellCountRow.getLastCellNum();
				int totalSkiplines = skipLines+1;
				row = sh.getRow(totalSkiplines);
				if(row == null){
					row = sh.createRow(totalSkiplines);
					for (int k=0;k<dataCellNum;k++) {
						row.createCell(k, CellType.BLANK);
					}
					if(randomNumberString){
						randomArray = PoireExcelWorkbook.getArrayOfRandomNumber(columnPositions.length, lengthOfRandomString);
					}else{
						randomArray = PoireExcelWorkbook.getArrayOfRandomString(columnPositions.length, lengthOfRandomString);
					}
					for (int indexArr : columnPositions) {
						if(indexArr>=dataCellNum){
							lengthOfCell = true;
						}
					}
					int increment = 0;
                    if(!lengthOfCell){
                    	for (int pos : columnPositions) {
    						cell = row.createCell(pos, CellType.STRING);
    						cell.setCellValue((String) randomArray.get(increment));
    						increment++;
    					}
    					fileOut = new FileOutputStream(filePath);
    					wb.write(fileOut);
    					fileOut.flush();
    					fileOut.close();
    					bStreenClose = true;
    					_sReturnObj = true;
                    }else{
                    	System.err.println("WARRING: Column Positions index values are greater then the last Data Cell index");
						System.out.println("\tPlease change the Column Positions index with appropreate Data cells index");
                    }
				}else{
					//row.setRowNum(totalSkiplines);
					if(randomNumberString){
						randomArray = PoireExcelWorkbook.getArrayOfRandomNumber(columnPositions.length, lengthOfRandomString);
					}else{
						randomArray = PoireExcelWorkbook.getArrayOfRandomString(columnPositions.length, lengthOfRandomString);
					}
					for (int indexArr : columnPositions) {
						if(indexArr>=dataCellNum){
							lengthOfCell = true;
						}
					}
					int increment = 0;
					if(!lengthOfCell){
						for (int pos : columnPositions) {
							cell = row.createCell(pos, CellType.STRING);
							cell.setCellValue((String) randomArray.get(increment));
							increment++;
						}
						fileOut = new FileOutputStream(filePath);
						wb.write(fileOut);
						fileOut.flush();
						fileOut.close();
						bStreenClose = true;
						_sReturnObj = true;
					}else{
						System.err.println("WARRING: Column Positions index values are greater then the last Data Cell index");
						System.out.println("\tPlease change the Column Positions index with appropreate Data cells index");
					}
				}
			}else{
				_sReturnObj = false;
				System.err.println("SkipLines value is greter than the Data present in the sheet Name:"+ sheetName);
			}
		}catch(FileNotFoundException fe){
			Common.addExceptionLogger(fe);
		} catch (Exception e) {
			Common.addExceptionLogger(e);
		}finally{
			if(xlsx){
				try {
					if(_OPCXSSFW != null){
						_OPCXSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				} 
			}else{
				try {
					if(_NFSHSSFW != null){
						_NFSHSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		return _sReturnObj;
		
	}
	
	/**<p>It will used to Update the data with specified arrayList in existing Excel row with columnPositions values.
	 * @author <a href="mailto:ssirekumar@gmail.com">Siri Kumar {@literal <ssirekumar@gmail.com>}</a> </br>
	 * @return boolean
	 * @param  filePath  - File path of the .xls or .xlsx file.
	 * @param  rowData   - Arraylist of row data for specified column
	 * @param  columnPositions - Array of Column index values which is starts from 0. Exp new int{0,2,4,5,8}
	 * @param  rowNumber  - How many lines to skip for updating the row with column index.
	 * @param  sheetName - Name of the sheet.
	 * 
	 * */
	public static boolean updateListDataInExcelFile(String filePath, ArrayList<String> rowData, int[] columnPositions, int rowNumber, String sheetName){
		File sFilePathObj = null;
		OPCPackage _OPCXSSFW = null;
		POIFSFileSystem  _NFSHSSFW = null;
		Row row = null;
		Cell cell = null;
		FileOutputStream fileOut = null;
		boolean _sReturnObj = false;
		boolean bStreenClose = false;
		boolean lengthOfCell = false;
		try {
			sFilePathObj = new File(filePath);
			xlsx = PoireExcelWorkbook.excelFileType(sFilePathObj).equalsIgnoreCase("xlsx");
			if(xlsx){
				_OPCXSSFW = OPCPackage.open(new FileInputStream(sFilePathObj));
			}else{
				_NFSHSSFW = new POIFSFileSystem(new FileInputStream(sFilePathObj));
			}
			Workbook wb = xlsx ? XSSFWorkbookFactory.create(_OPCXSSFW) : WorkbookFactory.create(_NFSHSSFW);
			Sheet sh = wb.getSheet(sheetName);
			int dataCellNum = 0;
			if(sh.getLastRowNum()>=rowNumber){
				dataCellNum = PoireExcelWorkbook.getColumnCount(filePath, sheetName);
				row = sh.getRow(rowNumber);
				if(row == null){
					row = sh.createRow(rowNumber);
					for (int k=0;k<dataCellNum;k++) {
						row.createCell(k, CellType.BLANK);
					}
					for (int indexArr : columnPositions) {
						if(indexArr>=dataCellNum){
							lengthOfCell = true;
						}
					}
					int increment = 0;
					if(!lengthOfCell){
						if(rowData.size() == columnPositions.length){
							for (int pos : columnPositions) {
								cell = row.getCell(pos);
								if(cell.toString() == ""){
									cell.setCellValue((String) rowData.get(increment));
								}
								increment++;
							}
							fileOut = new FileOutputStream(filePath);
							wb.write(fileOut);
							fileOut.flush();
							fileOut.close();
							bStreenClose = true;
							_sReturnObj = true;
						}else{
							System.err.println("WARRING: columnPositions size and rowData ArrayList size both are not equal");
						}
					}else{
						System.err.println("WARRING: Column Positions index values are greater than the last Data Cell index");
						System.out.println("\tPlease change the Column Positions index with appropreate Data cells index");
					}
				}else{
					for (int indexArr : columnPositions) {
						if(indexArr>=dataCellNum){
							lengthOfCell = true;
						}
					}
					int increment = 0;
					if(!lengthOfCell){
						if(rowData.size() == columnPositions.length){
							for (int pos : columnPositions) {
								cell = row.createCell(pos, CellType.STRING);
								cell.setCellValue((String) rowData.get(increment));
								increment++;
							}
							fileOut = new FileOutputStream(filePath);
							wb.write(fileOut);
							fileOut.flush();
							fileOut.close();
							bStreenClose = true;
							_sReturnObj = true;
						}else{
							System.err.println("WARRING: columnPositions size and rowData size both are not equal");
						}
					}else{
						System.err.println("WARRING: Column Positions index values are greater then the last Data Cell index");
						System.out.println("\tPlease change the Column Positions index with appropreate Data cells index");
					}
				}
			}else{
				_sReturnObj = false;
				System.err.println("RowNumber value is greter than the Data present in the sheet with Name:"+ sheetName);
			}
		}catch(FileNotFoundException fe){
			Common.addExceptionLogger(fe);
		} catch (Exception e) {
			Common.addExceptionLogger(e);
		}finally{
			if(xlsx){
				try {
					if(_OPCXSSFW != null){
						_OPCXSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				} 
			}else{
				try {
					if(_NFSHSSFW != null){
						_NFSHSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		return _sReturnObj;
	}
		
	/**<p>It will used to Update the data with specified arrayList in existing Excel row data with columnPositions values.
	 * @author <a href="mailto:ssirekumar@gmail.com">Siri Kumar {@literal <ssirekumar@gmail.com>}</a> </br>
	 * @return boolean
	 * @param  filePath  - File path of the .xls or .xlsx file.
	 * @param  rowData   - Arraylist of row data for specified column
	 * @param  columnPositions - Array of Column index values which is starts from 0. Exp new int{0,2,4,5,8}
	 * @param  rowNumber  - How many lines to skip for updating the row with column index.
	 * @param  sheetNumberIndex - Number of the sheet it will start from index 0.
	 * 
	 * 
	 * */
	public static boolean updateListDataInExcelFile(String filePath, ArrayList<String> rowData, int[] columnPositions, int rowNumber, int sheetNumberIndex){
		File sFilePathObj = null;
		OPCPackage _OPCXSSFW = null;
		POIFSFileSystem  _NFSHSSFW = null;
		Row row = null;
		Cell cell = null;
		FileOutputStream fileOut = null;
		boolean _sReturnObj = false;
		boolean bStreenClose = false;
		boolean lengthOfCell = false;
		try {
			sFilePathObj = new File(filePath);
			xlsx = PoireExcelWorkbook.excelFileType(sFilePathObj).equalsIgnoreCase("xlsx");
			if(xlsx){
				_OPCXSSFW = OPCPackage.open(new FileInputStream(sFilePathObj));
			}else{
				_NFSHSSFW = new POIFSFileSystem(new FileInputStream(sFilePathObj));
			}
			Workbook wb = xlsx ? XSSFWorkbookFactory.create(_OPCXSSFW) : WorkbookFactory.create(_NFSHSSFW);
			Sheet sh = wb.getSheetAt(sheetNumberIndex);
			int dataCellNum = 0;
			if(sh.getLastRowNum()>=rowNumber){
				dataCellNum = PoireExcelWorkbook.getColumnCount(filePath, sheetNumberIndex);
				row = sh.getRow(rowNumber);
				if(row == null){
					row = sh.createRow(rowNumber);
					for (int k=0;k<dataCellNum;k++) {
						row.createCell(k, CellType.BLANK);
					}
					for (int indexArr : columnPositions) {
						if(indexArr>=dataCellNum){
							lengthOfCell = true;
						}
					}
					int increment = 0;
					if(!lengthOfCell){
						if(rowData.size() == columnPositions.length){
							for (int pos : columnPositions) {
								cell = row.getCell(pos);
								if(cell.toString() == ""){
									cell.setCellValue((String) rowData.get(increment));
								}
								increment++;
							}
							fileOut = new FileOutputStream(filePath);
							wb.write(fileOut);
							fileOut.flush();
							fileOut.close();
							bStreenClose = true;
							_sReturnObj = true;
						}else{
							System.err.println("WARRING: columnPositions size and rowData ArrayList size both are not equal");
						}
					}else{
						System.err.println("WARRING: Column Positions index values are greater then the last Data Cell index");
						System.out.println("\tPlease change the Column Positions index with appropreate Data cells index");
					}
				}else{
					for (int indexArr : columnPositions) {
						if(indexArr>=dataCellNum){
							lengthOfCell = true;
						}
					}
					int increment = 0;
					if(!lengthOfCell){
						if(rowData.size() == columnPositions.length){
							for (int pos : columnPositions) {
								cell = row.createCell(pos, CellType.STRING);
								cell.setCellValue((String) rowData.get(increment));
								increment++;
							}
							fileOut = new FileOutputStream(filePath);
							wb.write(fileOut);
							fileOut.flush();
							fileOut.close();
							bStreenClose = true;
							_sReturnObj = true;
						}else{
							System.err.println("WARRING: columnPositions size and rowData size both are not equal");
						}
					}else{
						System.err.println("WARRING: Column Positions index values are greater then the last Data Cell index");
						System.out.println("\tPlease change the Column Positions index with appropreate Data cells index");
					}
				}
			}else{
				_sReturnObj = false;
				System.err.println("RowNumber value is greter than the Data present in the sheet Number:"+ sheetNumberIndex);
			}
		}catch(FileNotFoundException fe){
			Common.addExceptionLogger(fe);
		} catch (Exception e) {
			Common.addExceptionLogger(e);
		}finally{
			if(xlsx){
				try {
					if(_OPCXSSFW != null){
						_OPCXSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				} 
			}else{
				try {
					if(_NFSHSSFW != null){
						_NFSHSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		return _sReturnObj;
	}
		
	/**<p>It will used to Create/Update the data with specified arrayList based on sheet number insed in existing Excel.
	 * @author <a href="mailto:ssirekumar@gmail.com">Siri Kumar {@literal <ssirekumar@gmail.com>}</a> </br>
	 * @return boolean
	 * @param  filePath  - File path of the .xls or .xlsx file.
	 * @param  rowData   - Arraylist of data for creating data row.
	 * @param  rowNumber  - Positive Number of the row to insert data. this index will start from 0.
	 * @param  sheetNumberIndex - Number of the sheet it will start from index 0.
	 * 
	 * 
	 * */
	public static boolean createRowListDataInExcelFile(String filePath, ArrayList<String> rowData, int rowNumber, int sheetNumberIndex){ 
		File sFilePathObj = null;
		OPCPackage _OPCXSSFW = null;
		POIFSFileSystem  _NFSHSSFW = null;
		Row row = null;
		Cell cell = null;
		FileOutputStream fileOut = null;
		boolean _sReturnObj = false;
		boolean bStreenClose = false;
		try {
			sFilePathObj = new File(filePath);
			xlsx = PoireExcelWorkbook.excelFileType(sFilePathObj).equalsIgnoreCase("xlsx");
			if(xlsx){
				_OPCXSSFW = OPCPackage.open(new FileInputStream(sFilePathObj));
			}else{
				_NFSHSSFW = new POIFSFileSystem(new FileInputStream(sFilePathObj));
			}
			Workbook wb = xlsx ? XSSFWorkbookFactory.create(_OPCXSSFW) : WorkbookFactory.create(_NFSHSSFW);
			Sheet sh = wb.getSheetAt(sheetNumberIndex);
			int dataCellNum = 0;
			dataCellNum = rowData.size()-1;
			if (rowNumber<0) {
				System.err.println("WARRING: Given rowNumber is not appropreate it should be a positive Number");
				return false;
			} 
			if(sh.getLastRowNum()>=rowNumber){  
				row = sh.getRow(rowNumber);
				if(row == null){
					row = sh.createRow(rowNumber);
					for (int k=0;k<=dataCellNum;k++) {
						cell = row.createCell(k, CellType.STRING);
						cell.setCellValue((String) rowData.get(k));
					}
				}else{
					int increment = 0;
					for (String posValue : rowData) {
						cell = row.createCell(increment, CellType.STRING);
						cell.setCellValue((String) posValue);
						increment++;
					}
				}
				fileOut = new FileOutputStream(filePath);
				wb.write(fileOut);
				fileOut.flush();
				fileOut.close();
				bStreenClose = true;
				_sReturnObj = true;
			}else if(sh.getLastRowNum()<rowNumber){
				row = sh.getRow(rowNumber);
				if(row == null){
					row = sh.createRow(rowNumber);
					for (int k=0;k<=dataCellNum;k++) {
						cell = row.createCell(k, CellType.STRING);
						cell.setCellValue((String) rowData.get(k));
					}
				}else{
					int increment = 0;
					for (String posValue : rowData) {
						cell = row.createCell(increment, CellType.STRING);
						cell.setCellValue((String) posValue);
						increment++;
					}
				}
				fileOut = new FileOutputStream(filePath);
				wb.write(fileOut);
				fileOut.flush();
				fileOut.close();
				bStreenClose = true;
				_sReturnObj = true;
			}else{
				System.err.println("WARRING: Given rowNumber is not appropreate it should be a positive Number");
				System.out.println("\tPlease change the rowNumber according to method description");
			}
		}catch(FileNotFoundException fe){
			Common.addExceptionLogger(fe);
		} catch (Exception e) {
			Common.addExceptionLogger(e);
		}finally{
			if(xlsx){
				try {
					if(_OPCXSSFW != null){
						_OPCXSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				} 
			}else{
				try {
					if(_NFSHSSFW != null){
						_NFSHSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		return _sReturnObj;
	}
	
	/**<p>It will used to Create/Update the row with specified arrayList based on the sheet name in existing Excel.
	 * @author <a href="mailto:ssirekumar@gmail.com">Siri Kumar {@literal <ssirekumar@gmail.com>}</a> </br>
	 * @return boolean
	 * @param  filePath  - File path of the .xls or .xlsx file.
	 * @param  rowData   - Arraylist of data for creating data row.
	 * @param  rowNumber  - Positive Number of the row to insert data. this index will start from 0.
	 * @param  sheetName - Name of the sheet.
	 * 
	 * 
	 * */
	public static boolean createRowListDataInExcelFile(String filePath, ArrayList<String> rowData, int rowNumber, String sheetName){
		File sFilePathObj = null;
		OPCPackage _OPCXSSFW = null;
		POIFSFileSystem  _NFSHSSFW = null;
		Row row = null;
		Cell cell = null;
		FileOutputStream fileOut = null;
		boolean _sReturnObj = false;
		boolean bStreenClose = false;
		try {
			sFilePathObj = new File(filePath);
			xlsx = PoireExcelWorkbook.excelFileType(sFilePathObj).equalsIgnoreCase("xlsx");
			if(xlsx){
				_OPCXSSFW = OPCPackage.open(new FileInputStream(sFilePathObj));
			}else{
				_NFSHSSFW = new POIFSFileSystem(new FileInputStream(sFilePathObj));
			}
			Workbook wb = xlsx ? XSSFWorkbookFactory.create(_OPCXSSFW) : WorkbookFactory.create(_NFSHSSFW);
			Sheet sh = wb.getSheet(sheetName);
			int dataCellNum = 0;
			dataCellNum = rowData.size()-1;
			if (rowNumber<0) {
				System.err.println("WARRING: Given rowNumber is not appropreate it should be a positive Number");
				return false;
			} 
			if(sh.getLastRowNum()>=rowNumber){  
				row = sh.getRow(rowNumber);
				if(row == null){
					row = sh.createRow(rowNumber);
					for (int k=0;k<=dataCellNum;k++) {
						cell = row.createCell(k, CellType.STRING);
						cell.setCellValue((String) rowData.get(k));
					}
				}else{
					int increment = 0;
					for (String posValue : rowData) {
						cell = row.createCell(increment, CellType.STRING);
						cell.setCellValue((String) posValue);
						increment++;
					}
				}
				fileOut = new FileOutputStream(filePath);
				wb.write(fileOut);
				fileOut.flush();
				fileOut.close();
				bStreenClose = true;
				_sReturnObj = true;
			}else if(sh.getLastRowNum()<rowNumber){
				row = sh.getRow(rowNumber);
				if(row == null){
					row = sh.createRow(rowNumber);
					for (int k=0;k<=dataCellNum;k++) {
						cell = row.createCell(k, CellType.STRING);
						cell.setCellValue((String) rowData.get(k));
					}
				}else{
					int increment = 0;
					for (String posValue : rowData) {
						cell = row.createCell(increment, CellType.STRING);
						cell.setCellValue((String) posValue);
						increment++;
					}
				}
				fileOut = new FileOutputStream(filePath);
				wb.write(fileOut);
				fileOut.flush();
				fileOut.close();
				bStreenClose = true;
				_sReturnObj = true;
			}else{
				System.err.println("WARRING: Given rowNumber is not appropreate it should be a positive Number");
				System.out.println("\tPlease change the rowNumber according to method description");
			}
		}catch(FileNotFoundException fe){
			Common.addExceptionLogger(fe);
		} catch (Exception e) {
			Common.addExceptionLogger(e);
		}finally{
			if(xlsx){
				try {
					if(_OPCXSSFW != null){
						_OPCXSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				} 
			}else{
				try {
					if(_NFSHSSFW != null){
						_NFSHSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		return _sReturnObj;
	}
	
	/**<p>It will used to Create/Update the row from column position with specified arrayList based on the sheet name in existing Excel.
	 * @author <a href="mailto:ssirekumar@gmail.com">Siri Kumar {@literal <ssirekumar@gmail.com>}</a> </br>
	 * @return boolean
	 * @param  filePath  - File path of the .xls or .xlsx file.
	 * @param  rowData   - Arraylist of data for creating data row.
	 * @param  rowNumber  - Positive Number of the row to insert data. this index will start from 0.
	 * @param  sheetName - Name of the sheet.
	 * @param  fromColumnPosition - From which column position to create or update the row data. 
	 * 
	 * */
	public static boolean createRowListDataFromColPositionInExcelFile(String filePath, ArrayList<String> rowData, int rowNumber, String sheetName, int fromColumnPosition) {
		File sFilePathObj = null;
		OPCPackage _OPCXSSFW = null;
		POIFSFileSystem  _NFSHSSFW = null;
		Row row = null;
		Cell cell = null;
		FileOutputStream fileOut = null;
		boolean _sReturnObj = false;
		boolean bStreenClose = false;
		try {
			sFilePathObj = new File(filePath);
			xlsx = PoireExcelWorkbook.excelFileType(sFilePathObj).equalsIgnoreCase("xlsx");
			if(xlsx){
				_OPCXSSFW = OPCPackage.open(new FileInputStream(sFilePathObj));
			}else{
				_NFSHSSFW = new POIFSFileSystem(new FileInputStream(sFilePathObj));
			}
			Workbook wb = xlsx ? XSSFWorkbookFactory.create(_OPCXSSFW) : WorkbookFactory.create(_NFSHSSFW);
			Sheet sh = wb.getSheet(sheetName);
			int dataCellNum = 0;
			dataCellNum = rowData.size()-1;
			if (rowNumber<0) {
				System.err.println("WARRING: Given rowNumber is not appropreate it should be a positive Number");
				return false;
			} 
			if(sh.getLastRowNum()>=rowNumber){  
				row = sh.getRow(rowNumber);
				if(row == null){
					row = sh.createRow(rowNumber);
					for (int i = 0; i < fromColumnPosition; i++) {
						cell = row.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
						String _cellData = PoireExcelWorkbook.cellFormatedDataValue(cell);
						if(_cellData == ""){
							continue;
						}
					}
					for (int k=0;k<=dataCellNum;k++) {
						cell = row.createCell(fromColumnPosition+k, CellType.STRING);
						cell.setCellValue((String) rowData.get(k));
					}
				}else{
					int increment = 0;
					for (int i = 0; i < fromColumnPosition; i++) {
						cell = row.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
						String _cellData = PoireExcelWorkbook.cellFormatedDataValue(cell);
						if(_cellData == ""){
							continue;
						}
					}
					for (String posValue : rowData) {
						cell = row.createCell(increment+fromColumnPosition, CellType.STRING);
						cell.setCellValue((String) posValue);
						increment++;
					}
				}
				fileOut = new FileOutputStream(filePath);
				wb.write(fileOut);
				fileOut.flush();
				fileOut.close();
				bStreenClose = true;
				_sReturnObj = true;
			}else if(sh.getLastRowNum()<rowNumber){
				row = sh.getRow(rowNumber);
				if(row == null){
					row = sh.createRow(rowNumber);
					for (int k=0;k<=dataCellNum;k++) {
						cell = row.createCell(k, CellType.STRING);
						cell.setCellValue((String) rowData.get(k));
					}
				}else{
					int increment = 0;
					for (String posValue : rowData) {
						cell = row.createCell(increment, CellType.STRING);
						cell.setCellValue((String) posValue);
						increment++;
					}
				}
				fileOut = new FileOutputStream(filePath);
				wb.write(fileOut);
				fileOut.flush();
				fileOut.close();
				bStreenClose = true;
				_sReturnObj = true;
			}else{
				System.err.println("WARRING: Given rowNumber is not appropreate it should be a positive Number");
				System.out.println("\tPlease change the rowNumber according to method description");
			}
			if(xlsx){
				_OPCXSSFW.close();
				bStreenClose = true;
			}else{
				_NFSHSSFW.close();
				bStreenClose = true;
			}
		}catch(FileNotFoundException fe){
			Common.addExceptionLogger(fe);
		} catch (Exception e) {
			Common.addExceptionLogger(e);
		}finally{
			if(xlsx){
				try {
					if(_OPCXSSFW != null){
						_OPCXSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				} 
			}else{
				try {
					if(_NFSHSSFW != null){
						_NFSHSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		return _sReturnObj;
	}
	
	/**<p>It will used to Create/Update the row from column position with specified arrayList based on the sheet Number Index in existing Excel.
	 * @author <a href="mailto:ssirekumar@gmail.com">Siri Kumar {@literal <ssirekumar@gmail.com>}</a> </br>
	 * @return boolean
	 * @param  filePath  - File path of the .xls or .xlsx file.
	 * @param  rowData   - Arraylist of data for creating data row.
	 * @param  rowNumber  - Positive Number of the row to insert data. this index will start from 0.
	 * @param  sheetNumberIndex - Number of the sheet index. which is start from 0.
	 * @param  fromColumnPosition - From which column position to create or update the row data. 
	 * 
	 * */
	public static boolean createRowListDataFromColPositionInExcelFile(String filePath, ArrayList<String> rowData, int rowNumber, int sheetNumberIndex, int fromColumnPosition) {
		File sFilePathObj = null;
		OPCPackage _OPCXSSFW = null;
		POIFSFileSystem  _NFSHSSFW = null;
		Row row = null;
		Cell cell = null;
		FileOutputStream fileOut = null;
		boolean _sReturnObj = false;
		boolean bStreenClose = false;
		try {
			sFilePathObj = new File(filePath);
			xlsx = PoireExcelWorkbook.excelFileType(sFilePathObj).equalsIgnoreCase("xlsx");
			if(xlsx){
				_OPCXSSFW = OPCPackage.open(new FileInputStream(sFilePathObj));
			}else{
				_NFSHSSFW = new POIFSFileSystem(new FileInputStream(sFilePathObj));
			}
			Workbook wb = xlsx ? XSSFWorkbookFactory.create(_OPCXSSFW) : WorkbookFactory.create(_NFSHSSFW);
			Sheet sh = wb.getSheetAt(sheetNumberIndex);
			int dataCellNum = 0;
			dataCellNum = rowData.size()-1;
			if (rowNumber<0) {
				System.err.println("WARRING: Given rowNumber is not appropreate it should be a positive Number");
				return false;
			} 
			if(sh.getLastRowNum()>=rowNumber){  
				row = sh.getRow(rowNumber);
				if(row == null){
					row = sh.createRow(rowNumber);
					for (int i = 0; i < fromColumnPosition; i++) {
						cell = row.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
						String _cellData = PoireExcelWorkbook.cellFormatedDataValue(cell);
						if(_cellData == ""){
							continue;
						}
					}
					for (int k=0;k<=dataCellNum;k++) {
						cell = row.createCell(fromColumnPosition+k, CellType.STRING);
						cell.setCellValue((String) rowData.get(k));
					}
				}else{
					int increment = 0;
					for (int i = 0; i < fromColumnPosition; i++) {
						cell = row.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
						String _cellData = PoireExcelWorkbook.cellFormatedDataValue(cell);
						if(_cellData == ""){
							continue;
						}
					}
					for (String posValue : rowData) {
						cell = row.createCell(increment+fromColumnPosition, CellType.STRING);
						cell.setCellValue((String) posValue);
						increment++;
					}
				}
				fileOut = new FileOutputStream(filePath);
				wb.write(fileOut);
				fileOut.flush();
				fileOut.close();
				bStreenClose = true;
				_sReturnObj = true;
			}else if(sh.getLastRowNum()<rowNumber){
				row = sh.getRow(rowNumber);
				if(row == null){
					row = sh.createRow(rowNumber);
					for (int k=0;k<=dataCellNum;k++) {
						cell = row.createCell(k, CellType.STRING);
						cell.setCellValue((String) rowData.get(k));
					}
				}else{
					int increment = 0;
					for (String posValue : rowData) {
						cell = row.createCell(increment, CellType.STRING);
						cell.setCellValue((String) posValue);
						increment++;
					}
				}
				fileOut = new FileOutputStream(filePath);
				wb.write(fileOut);
				fileOut.flush();
				fileOut.close();
				bStreenClose = true;
				_sReturnObj = true;
			}else{
				System.err.println("WARRING: Given rowNumber is not appropreate it should be a positive Number");
				System.out.println("\tPlease change the rowNumber according to method description");
			}
			if(xlsx){
				_OPCXSSFW.close();
				bStreenClose = true;
			}else{
				_NFSHSSFW.close();
				bStreenClose = true;
			}
		}catch(FileNotFoundException fe){
			Common.addExceptionLogger(fe);
		} catch (Exception e) {
			Common.addExceptionLogger(e);
		}finally{
			if(xlsx){
				try {
					if(_OPCXSSFW != null){
						_OPCXSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				} 
			}else{
				try {
					if(_NFSHSSFW != null){
						_NFSHSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		return _sReturnObj;
	}
	
	/**<p>It will used to Create/Update the row with same data from column position with specified rowdata based on the sheet Number Index in existing Excel.
	 * @author <a href="mailto:ssirekumar@gmail.com">Siri Kumar {@literal <ssirekumar@gmail.com>}</a> </br>
	 * @return boolean
	 * @param  filePath  - File path of the .xls or .xlsx file.
	 * @param  rowData   - String of data for creating same data in cell up to uptoColumnNumber value.
	 * @param  rowNumber  - Positive Number of the row to insert data. this index will start from 0.
	 * @param  sheetNumberIndex - Number of the sheet index. which is start from 0.
	 * @param  fromColumnPosition - From which column position to create or update the row data. 
	 * @param  uptoColumnNumber   - int number of the column value.
	 * */
	public static boolean createRowDataFromColPositionInExcelFile(String filePath, String rowData, int rowNumber, int sheetNumberIndex, int fromColumnPosition, int uptoColumnNumber) {
		File sFilePathObj = null;
		OPCPackage _OPCXSSFW = null;
		POIFSFileSystem  _NFSHSSFW = null;
		Row row = null;
		Cell cell = null;
		FileOutputStream fileOut = null;
		boolean _sReturnObj = false;
		boolean bStreenClose = false;
		try {
			sFilePathObj = new File(filePath);
			xlsx = PoireExcelWorkbook.excelFileType(sFilePathObj).equalsIgnoreCase("xlsx");
			if(xlsx){
				_OPCXSSFW = OPCPackage.open(new FileInputStream(sFilePathObj));
			}else{
				_NFSHSSFW = new POIFSFileSystem(new FileInputStream(sFilePathObj));
			}
			Workbook wb = xlsx ? XSSFWorkbookFactory.create(_OPCXSSFW) : WorkbookFactory.create(_NFSHSSFW);
			Sheet sh = wb.getSheetAt(sheetNumberIndex);
			int dataCellNum = 0;
			dataCellNum = uptoColumnNumber;
			if (rowNumber<0) {
				System.err.println("WARRING: Given rowNumber is not appropreate it should be a positive Number");
				return false;
			} 
			if(sh.getLastRowNum()>=rowNumber){  
				row = sh.getRow(rowNumber);
				if(row == null){
					row = sh.createRow(rowNumber);
					for (int i = 0; i < fromColumnPosition; i++) {
						cell = row.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
						String _cellData = PoireExcelWorkbook.cellFormatedDataValue(cell);
						if(_cellData == ""){
							continue;
						}
					}
					for (int k=0;k<=dataCellNum;k++) {
						cell = row.createCell(fromColumnPosition+k, CellType.STRING);
						cell.setCellValue((String) rowData);
					}
				}else{
					for (int i = 0; i < fromColumnPosition; i++) {
						cell = row.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
						String _cellData = PoireExcelWorkbook.cellFormatedDataValue(cell);
						if(_cellData == ""){
							continue;
						}
					}
					for (int k=0;k<=dataCellNum;k++) {
						cell = row.createCell(fromColumnPosition+k, CellType.STRING);
						cell.setCellValue((String) rowData);
					}
				}
				fileOut = new FileOutputStream(filePath);
				wb.write(fileOut);
				fileOut.flush();
				fileOut.close();
				bStreenClose = true;
				_sReturnObj = true;
			}else if(sh.getLastRowNum()<rowNumber){
				row = sh.getRow(rowNumber);
				if(row == null){
					row = sh.createRow(rowNumber);
					for (int k=0;k<=dataCellNum;k++) {
						cell = row.createCell(k, CellType.STRING);
						cell.setCellValue((String) rowData);
					}
				}else{
					for (int k=0;k<=dataCellNum;k++) {
						cell = row.createCell(fromColumnPosition+k, CellType.STRING);
						cell.setCellValue((String) rowData);
					}
				}
				fileOut = new FileOutputStream(filePath);
				wb.write(fileOut);
				fileOut.flush();
				fileOut.close();
				bStreenClose = true;
				_sReturnObj = true;
			}else{
				System.err.println("WARRING: Given rowNumber is not appropreate it should be a positive Number");
				System.out.println("\tPlease change the rowNumber according to method description");
			}
			if(xlsx){
				_OPCXSSFW.close();
				bStreenClose = true;
			}else{
				_NFSHSSFW.close();
				bStreenClose = true;
			}
		}catch(FileNotFoundException fe){
			Common.addExceptionLogger(fe);
		} catch (Exception e) {
			Common.addExceptionLogger(e);
		}finally{
			if(xlsx){
				try {
					if(_OPCXSSFW != null){
						_OPCXSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				} 
			}else{
				try {
					if(_NFSHSSFW != null){
						_NFSHSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		return _sReturnObj;
	}
	
	/**<p>It will used to Create/Update the row with same data from column position with specified rowdata based on the sheet name in existing Excel.
	 * @author <a href="mailto:ssirekumar@gmail.com">Siri Kumar {@literal <ssirekumar@gmail.com>}</a> </br>
	 * @return boolean
	 * @param  filePath  - File path of the .xls or .xlsx file.
	 * @param  rowData   - String of data for creating same data in cell up to uptoColumnNumber value.
	 * @param  rowNumber  - Positive Number of the row to insert data. this index will start from 0.
	 * @param  sheetName - Number of the sheet index. which is start from 0.
	 * @param  fromColumnPosition - From which column position to create or update the row data. 
	 * @param  uptoColumnNumber   - int number of the column value.
	 * */
	public static boolean createRowDataFromColPositionInExcelFile(String filePath, String rowData, int rowNumber, String sheetName, int fromColumnPosition, int uptoColumnNumber) {
		File sFilePathObj = null;
		OPCPackage _OPCXSSFW = null;
		POIFSFileSystem  _NFSHSSFW = null;
		Row row = null;
		Cell cell = null;
		FileOutputStream fileOut = null;
		boolean _sReturnObj = false;
		boolean bStreenClose = false;
		try {
			sFilePathObj = new File(filePath);
			xlsx = PoireExcelWorkbook.excelFileType(sFilePathObj).equalsIgnoreCase("xlsx");
			if(xlsx){
				_OPCXSSFW = OPCPackage.open(new FileInputStream(sFilePathObj));
			}else{
				_NFSHSSFW = new POIFSFileSystem(new FileInputStream(sFilePathObj));
			}
			Workbook wb = xlsx ? XSSFWorkbookFactory.create(_OPCXSSFW) : WorkbookFactory.create(_NFSHSSFW);
			Sheet sh = wb.getSheet(sheetName);
			int dataCellNum = 0;
			dataCellNum = uptoColumnNumber;
			if (rowNumber<0) {
				System.err.println("WARRING: Given rowNumber is not appropreate it should be a positive Number");
				return false;
			} 
			if(sh.getLastRowNum()>=rowNumber){  
				row = sh.getRow(rowNumber);
				if(row == null){
					row = sh.createRow(rowNumber);
					for (int i = 0; i < fromColumnPosition; i++) {
						cell = row.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
						String _cellData = PoireExcelWorkbook.cellFormatedDataValue(cell);
						if(_cellData == ""){
							continue;
						}
					}
					for (int k=0;k<=dataCellNum;k++) {
						cell = row.createCell(fromColumnPosition+k, CellType.STRING);
						cell.setCellValue((String) rowData);
					}
				}else{
					for (int i = 0; i < fromColumnPosition; i++) {
						cell = row.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
						String _cellData = PoireExcelWorkbook.cellFormatedDataValue(cell);
						if(_cellData == ""){
							continue;
						}
					}
					for (int k=0;k<=dataCellNum;k++) {
						cell = row.createCell(fromColumnPosition+k, CellType.STRING);
						cell.setCellValue((String) rowData);
					}
				}
				fileOut = new FileOutputStream(filePath);
				wb.write(fileOut);
				fileOut.flush();
				fileOut.close();
				bStreenClose = true;
				_sReturnObj = true;
			}else if(sh.getLastRowNum()<rowNumber){
				row = sh.getRow(rowNumber);
				if(row == null){
					row = sh.createRow(rowNumber);
					for (int k=0;k<=dataCellNum;k++) {
						cell = row.createCell(k, CellType.STRING);
						cell.setCellValue((String) rowData);
					}
				}else{
					for (int k=0;k<=dataCellNum;k++) {
						cell = row.createCell(fromColumnPosition+k, CellType.STRING);
						cell.setCellValue((String) rowData);
					}
				}
				fileOut = new FileOutputStream(filePath);
				wb.write(fileOut);
				fileOut.flush();
				fileOut.close();
				bStreenClose = true;
				_sReturnObj = true;
			}else{
				System.err.println("WARRING: Given rowNumber is not appropreate it should be a positive Number");
				System.out.println("\tPlease change the rowNumber according to method description");
			}
		}catch(FileNotFoundException fe){
			Common.addExceptionLogger(fe);
		} catch (Exception e) {
			Common.addExceptionLogger(e);
		}finally{
			if(xlsx){
				try {
					if(_OPCXSSFW != null){
						_OPCXSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				} 
			}else{
				try {
					if(_NFSHSSFW != null){
						_NFSHSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		return _sReturnObj;
	}
		
	/**<p>It will used to Create/Update the Column with specified arrayList based on the sheet name in existing Excel.
	 * Genrally this will place the data with in the specified column from 0 index on words.
	 * @author <a href="mailto:ssirekumar@gmail.com">Siri Kumar {@literal <ssirekumar@gmail.com>}</a> </br>
	 * @return boolean
	 * @param  filePath  - File path of the .xls or .xlsx file.
	 * @param  rowData   - Arraylist of data for creating data row.
	 * @param  rowNumber  - Positive Number of the row to insert data. this index will start from 0.
	 * @param  sheetName - Name of the sheet.
	 * */
	public static boolean createColumnListDataInExcelFile(String filePath, ArrayList<String> rowData, int colNumber, String sheetName) {
		File sFilePathObj = null;
		OPCPackage _OPCXSSFW = null;
		POIFSFileSystem  _NFSHSSFW = null;
		Row row = null;
		Cell cell = null;
		FileOutputStream fileOut = null;
		boolean _sReturnObj = false;
		boolean bStreenClose = false;
		try {
			sFilePathObj = new File(filePath);
			xlsx = PoireExcelWorkbook.excelFileType(sFilePathObj).equalsIgnoreCase("xlsx");
			if(xlsx){
				_OPCXSSFW = OPCPackage.open(new FileInputStream(sFilePathObj));
			}else{
				_NFSHSSFW = new POIFSFileSystem(new FileInputStream(sFilePathObj));
			}
			Workbook wb = xlsx ? XSSFWorkbookFactory.create(_OPCXSSFW) : WorkbookFactory.create(_NFSHSSFW);
			Sheet sh = wb.getSheet(sheetName);
			int dataCellNum = 0;
			dataCellNum = rowData.size()-1;
			if (colNumber < 0) {
				System.err.println("WARRING: Given column number is not appropreate it should be a positive Number");
				return false;
			} 
			for (int i = 0; i <=dataCellNum; i++) {
				row = sh.getRow(i);
				if(row == null){
					row = sh.createRow(i);
					for (int k=0;k<=colNumber;k++) {
						cell = row.createCell(k, CellType.STRING);
					}
					cell.setCellValue((String) rowData.get(i));
				}else{
					for (int k=0;k<=colNumber;k++) {
						cell = row.getCell(k, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
						String _cellData = PoireExcelWorkbook.cellFormatedDataValue(cell);
						if(_cellData == ""){
							continue;
						}
					}
					cell.setCellValue((String) rowData.get(i));
				}
			}
			fileOut = new FileOutputStream(filePath);
			wb.write(fileOut);
			fileOut.flush();
			fileOut.close();
			bStreenClose = true;
			_sReturnObj = true;
		}catch(FileNotFoundException fe){
			Common.addExceptionLogger(fe);
		} catch (Exception e) {
			Common.addExceptionLogger(e);
		}finally{
			if(xlsx){
				try {
					if(_OPCXSSFW != null){
						_OPCXSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				} 
			}else{
				try {
					if(_NFSHSSFW != null){
						_NFSHSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		return _sReturnObj;
	}
	
	/**<p>It will used to Create/Update the Column with specified arrayList based on the sheet Number Index in existing Excel.
	 * Genrally this will place the data with in the specified column from 0 index on words.
	 * @author <a href="mailto:ssirekumar@gmail.com">Siri Kumar {@literal <ssirekumar@gmail.com>}</a> </br>
	 * @return boolean
	 * @param  filePath  - File path of the .xls or .xlsx file.
	 * @param  rowData   - Arraylist of data for creating data row.
	 * @param  rowNumber  - Positive Number of the row to insert data. this index will start from 0.
	 * @param  sheetNumberIndex - Number of the sheet it will start from index 0.
	 * */
	public static boolean createColumnListDataInExcelFile(String filePath, ArrayList<String> rowData, int colNumber, int sheetNumberIndex) {
		File sFilePathObj = null;
		OPCPackage _OPCXSSFW = null;
		POIFSFileSystem  _NFSHSSFW = null;
		Row row = null;
		Cell cell = null;
		FileOutputStream fileOut = null;
		boolean _sReturnObj = false;
		boolean bStreenClose = false;
		try {
			sFilePathObj = new File(filePath);
			xlsx = PoireExcelWorkbook.excelFileType(sFilePathObj).equalsIgnoreCase("xlsx");
			if(xlsx){
				_OPCXSSFW = OPCPackage.open(new FileInputStream(sFilePathObj));
			}else{
				_NFSHSSFW = new POIFSFileSystem(new FileInputStream(sFilePathObj));
			}
			Workbook wb = xlsx ? XSSFWorkbookFactory.create(_OPCXSSFW) : WorkbookFactory.create(_NFSHSSFW);
			Sheet sh = wb.getSheetAt(sheetNumberIndex);
			int dataCellNum = 0;
			dataCellNum = rowData.size()-1;
			if (colNumber < 0) {
				System.err.println("WARRING: Given column number is not appropreate it should be a positive Number");
				return false;
			} 
			for (int i = 0; i <=dataCellNum; i++) {
				row = sh.getRow(i);
				if(row == null){
					row = sh.createRow(i);
					for (int k=0;k<=colNumber;k++) {
						cell = row.createCell(k, CellType.STRING);
					}
					cell.setCellValue((String) rowData.get(i));
				}else{
					for (int k=0;k<=colNumber;k++) {
						cell = row.getCell(k, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
						String _cellData = PoireExcelWorkbook.cellFormatedDataValue(cell);
						if(_cellData == ""){
							continue;
						}
					}
					cell.setCellValue((String) rowData.get(i));
				}
			}
			fileOut = new FileOutputStream(filePath);
			wb.write(fileOut);
			fileOut.flush();
			fileOut.close();
			bStreenClose = true;
			_sReturnObj = true;
			if(xlsx){
				_OPCXSSFW.close();
				bStreenClose = true;
			}else{
				_NFSHSSFW.close();
				bStreenClose = true;
			}
		}catch(FileNotFoundException fe){
			Common.addExceptionLogger(fe);
		} catch (Exception e) {
			Common.addExceptionLogger(e);
		}finally{
			if(xlsx){
				try {
					if(!bStreenClose){
						_OPCXSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				} 
			}else{
				try {
					if(!bStreenClose){
						_NFSHSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		return _sReturnObj;
	}
		
	/**<p>It will used to Create/Update the Column with specified arrayList and its row and column position based on the sheet Name in existing Excel.
	 * @author <a href="mailto:ssirekumar@gmail.com">Siri Kumar {@literal <ssirekumar@gmail.com>}</a> </br>
	 * @return boolean
	 * @param  filePath  - File path of the .xls or .xlsx file.
	 * @param  rowData   - Arraylist of data for creating data row.
	 * @param  colNumber  - Positive Number of the column to insert data. this index will start from 0.
	 * @param  rowNumber  - Positive Number of the row to insert data. this index will start from 0.
	 * @param  sheetName - name of the sheet.
	 * */
	public static boolean createColumnListDataFromRowColPositionInExcelFile(String filePath, ArrayList<String> rowData, int rowNumber, int colNumber, String sheetName) {
		File sFilePathObj = null;
		OPCPackage _OPCXSSFW = null;
		POIFSFileSystem  _NFSHSSFW = null;
		Row row = null;
		Cell cell = null;
		FileOutputStream fileOut = null;
		boolean _sReturnObj = false;
		boolean bStreenClose = false;
		try {
			sFilePathObj = new File(filePath);
			xlsx = PoireExcelWorkbook.excelFileType(sFilePathObj).equalsIgnoreCase("xlsx");
			if(xlsx){
				_OPCXSSFW = OPCPackage.open(new FileInputStream(sFilePathObj));
			}else{
				_NFSHSSFW = new POIFSFileSystem(new FileInputStream(sFilePathObj));
			}
			Workbook wb = xlsx ? XSSFWorkbookFactory.create(_OPCXSSFW) : WorkbookFactory.create(_NFSHSSFW);
			Sheet sh = wb.getSheet(sheetName);
			int dataCellNum = 0;
			dataCellNum = rowData.size()-1;
			if (colNumber < 0) {
				System.err.println("WARRING: Given column number is not appropreate it should be a positive Number");
				return false;
			} 
			int dataIncre = 0;
			for (int i = rowNumber; i <=dataCellNum+rowNumber; i++) {
				row = sh.getRow(i);
				if(row == null){
					row = sh.createRow(i);
					for (int k=0;k<=colNumber;k++) {
						cell = row.createCell(k, CellType.STRING);
					}
					cell.setCellValue((String) rowData.get(dataIncre));
				}else{
					for (int k=0;k<=colNumber;k++) {
						cell = row.getCell(k, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
						String _cellData = PoireExcelWorkbook.cellFormatedDataValue(cell);
						if(_cellData == ""){
							continue;
						}
					}
					cell.setCellValue((String) rowData.get(dataIncre));
				}
				dataIncre++;
			}
			fileOut = new FileOutputStream(filePath);
			wb.write(fileOut);
			fileOut.flush();
			fileOut.close();
			bStreenClose = true;
			_sReturnObj = true;
			if(xlsx){
				_OPCXSSFW.close();
				bStreenClose = true;
			}else{
				_NFSHSSFW.close();
				bStreenClose = true;
			}
		}catch(FileNotFoundException fe){
			Common.addExceptionLogger(fe);
		} catch (Exception e) {
			Common.addExceptionLogger(e);
		}finally{
			if(xlsx){
				try {
					if(!bStreenClose){
						_OPCXSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				} 
			}else{
				try {
					if(!bStreenClose){
						_NFSHSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		return _sReturnObj;
	}
	
	/**<p>It will used to Create/Update the Column with specified arrayList and its row and column position based on the sheet Name in existing Excel.
	 * @author Siri Kumar Puttagunta
	 * @return boolean
	 * @param  filePath  - File path of the .xls or .xlsx file.
	 * @param  rowData   - Arraylist of data for creating data row.
	 * @param  colNumber  - Positive Number of the column to insert data. this index will start from 0.
	 * @param  rowNumber  - Positive Number of the row to insert data. this index will start from 0.
	 * @param  sheetNumberIndex - int Number of the sheet index. index its start from 0.
	 * */
	public static boolean createColumnListDataFromRowColPositionInExcelFile(String filePath, ArrayList<String> rowData, int rowNumber, int colNumber, int sheetNumberIndex) {
		File sFilePathObj = null;
		OPCPackage _OPCXSSFW = null;
		POIFSFileSystem  _NFSHSSFW = null;
		Row row = null;
		Cell cell = null;
		FileOutputStream fileOut = null;
		boolean _sReturnObj = false;
		boolean bStreenClose = false;
		try {
			sFilePathObj = new File(filePath);
			xlsx = PoireExcelWorkbook.excelFileType(sFilePathObj).equalsIgnoreCase("xlsx");
			if(xlsx){
				_OPCXSSFW = OPCPackage.open(new FileInputStream(sFilePathObj));
			}else{
				_NFSHSSFW = new POIFSFileSystem(new FileInputStream(sFilePathObj));
			}
			Workbook wb = xlsx ? XSSFWorkbookFactory.create(_OPCXSSFW) : WorkbookFactory.create(_NFSHSSFW);
			Sheet sh = wb.getSheetAt(sheetNumberIndex);
			int dataCellNum = 0;
			dataCellNum = rowData.size()-1;
			if (colNumber < 0) {
				System.err.println("WARRING: Given column number is not appropreate it should be a positive Number");
				return false;
			} 
			int dataIncre = 0;
			for (int i = rowNumber; i <=dataCellNum+rowNumber; i++) {
				row = sh.getRow(i);
				if(row == null){
					row = sh.createRow(i);
					for (int k=0;k<=colNumber;k++) {
						cell = row.createCell(k, CellType.STRING);
					}
					cell.setCellValue((String) rowData.get(dataIncre));
				}else{
					for (int k=0;k<=colNumber;k++) {
						cell = row.getCell(k, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
						String _cellData = PoireExcelWorkbook.cellFormatedDataValue(cell);
						if(_cellData == ""){
							continue;
						}
					}
					cell.setCellValue((String) rowData.get(dataIncre));
				}
				dataIncre++;
			}
			fileOut = new FileOutputStream(filePath);
			wb.write(fileOut);
			fileOut.flush();
			fileOut.close();
			bStreenClose = true;
			_sReturnObj = true;
			if(xlsx){
				_OPCXSSFW.close();
				bStreenClose = true;
			}else{
				_NFSHSSFW.close();
				bStreenClose = true;
			}
		}catch(FileNotFoundException fe){
			Common.addExceptionLogger(fe);
		} catch (Exception e) {
			Common.addExceptionLogger(e);
		}finally{
			if(xlsx){
				try {
					if(!bStreenClose){
						_OPCXSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				} 
			}else{
				try {
					if(!bStreenClose){
						_NFSHSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		return _sReturnObj;
	}
	
	/**<p>It will used to Create/Update the Column with specified string data and its row and column position based on the sheet Name in existing Excel.
	 * @author Siri Kumar Puttagunta
	 * @return boolean
	 * @param  filePath  - File path of the .xls or .xlsx file.
	 * @param  rowData   - string of data for creating data row.
	 * @param  colNumber  - Positive Number of the column to insert data. this index will start from 0.
	 * @param  rowNumber  - Positive Number of the row to insert data. this index will start from 0.
	 * @param  sheetNumberIndex - int Number of the sheet index. index its start from 0.
	 * @param  uptoRowNumber    - int number of the row to update/create the same string data.
	 * */
	public static boolean createColumnDataFromRowColPositionInExcelFile(String filePath, String rowData, int rowNumber, int colNumber, int sheetNumberIndex, int uptoRowNumber) {
		File sFilePathObj = null;
		OPCPackage _OPCXSSFW = null;
		POIFSFileSystem  _NFSHSSFW = null;
		Row row = null;
		Cell cell = null;
		FileOutputStream fileOut = null;
		boolean _sReturnObj = false;
		boolean bStreenClose = false;
		try {
			sFilePathObj = new File(filePath);
			xlsx = PoireExcelWorkbook.excelFileType(sFilePathObj).equalsIgnoreCase("xlsx");
			if(xlsx){
				_OPCXSSFW = OPCPackage.open(new FileInputStream(sFilePathObj));
			}else{
				_NFSHSSFW = new POIFSFileSystem(new FileInputStream(sFilePathObj));
			}
			Workbook wb = xlsx ? XSSFWorkbookFactory.create(_OPCXSSFW) : WorkbookFactory.create(_NFSHSSFW);
			Sheet sh = wb.getSheetAt(sheetNumberIndex);
			int dataCellNum = 0;
			dataCellNum = uptoRowNumber;
			if (colNumber < 0) {
				System.err.println("WARRING: Given column number is not appropreate it should be a positive Number");
				return false;
			} 
			for (int i = rowNumber; i <=dataCellNum; i++) {
				row = sh.getRow(i);
				if(row == null){
					row = sh.createRow(i);
					for (int k=0;k<=colNumber;k++) {
						cell = row.createCell(k, CellType.STRING);
					}
					cell.setCellValue((String) rowData);
				}else{
					for (int k=0;k<=colNumber;k++) {
						cell = row.getCell(k, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
						String _cellData = PoireExcelWorkbook.cellFormatedDataValue(cell);
						if(_cellData == ""){
							continue;
						}
					}
					cell.setCellValue((String) rowData);
				}
			}
			fileOut = new FileOutputStream(filePath);
			wb.write(fileOut);
			fileOut.flush();
			fileOut.close();
			bStreenClose = true;
			_sReturnObj = true;
			if(xlsx){
				_OPCXSSFW.close();
				bStreenClose = true;
			}else{
				_NFSHSSFW.close();
				bStreenClose = true;
			}
		}catch(FileNotFoundException fe){
			Common.addExceptionLogger(fe);
		} catch (Exception e) {
			Common.addExceptionLogger(e);
		}finally{
			if(xlsx){
				try {
					if(!bStreenClose){
						_OPCXSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				} 
			}else{
				try {
					if(!bStreenClose){
						_NFSHSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		return _sReturnObj;
	}
	
	/**<p>It will used to Create/Update the Column with specified string data and its row and column position based on the sheet Name in existing Excel.
	 * @author Siri Kumar Puttagunta
	 * @return boolean
	 * @param  filePath  - File path of the .xls or .xlsx file.
	 * @param  rowData   - string of data for creating data row.
	 * @param  colNumber  - Positive Number of the column to insert data. this index will start from 0.
	 * @param  rowNumber  - Positive Number of the row to insert data. this index will start from 0.
	 * @param  sheetNumberIndex - int Number of the sheet index. index its start from 0.
	 * @param  uptoRowNumber - int number of the row to update/create the same string data.
	 * Ex: 
	 * */
	public static boolean createColumnDataFromRowColPositionInExcelFile(String filePath, String rowData, int rowNumber, int colNumber, String sheetName, int uptoRowNumber) {
		File sFilePathObj = null;
		OPCPackage _OPCXSSFW = null;
		POIFSFileSystem  _NFSHSSFW = null;
		Row row = null;
		Cell cell = null;
		FileOutputStream fileOut = null;
		boolean _sReturnObj = false;
		boolean bStreenClose = false;
		try {
			sFilePathObj = new File(filePath);
			xlsx = PoireExcelWorkbook.excelFileType(sFilePathObj).equalsIgnoreCase("xlsx");
			if(xlsx){
				_OPCXSSFW = OPCPackage.open(new FileInputStream(sFilePathObj));
			}else{
				_NFSHSSFW = new POIFSFileSystem(new FileInputStream(sFilePathObj));
			}
			Workbook wb = xlsx ? XSSFWorkbookFactory.create(_OPCXSSFW) : WorkbookFactory.create(_NFSHSSFW);
			Sheet sh = wb.getSheet(sheetName);
			int dataCellNum = 0;
			dataCellNum = uptoRowNumber;
			if (colNumber < 0) {
				System.err.println("WARRING: Given column number is not appropreate it should be a positive Number");
				return false;
			} 
			for (int i = rowNumber; i <=dataCellNum; i++) {
				row = sh.getRow(i);
				if(row == null){
					row = sh.createRow(i);
					for (int k=0;k<=colNumber;k++) {
						cell = row.createCell(k, CellType.STRING);
					}
					cell.setCellValue((String) rowData);
				}else{
					for (int k=0;k<=colNumber;k++) {
						cell = row.getCell(k, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
						String _cellData = PoireExcelWorkbook.cellFormatedDataValue(cell);
						if(_cellData == ""){
							continue;
						}
					}
					cell.setCellValue((String) rowData);
				}
			}
			fileOut = new FileOutputStream(filePath);
			wb.write(fileOut);
			fileOut.flush();
			fileOut.close();
			bStreenClose = true;
			_sReturnObj = true;
			if(xlsx){
				_OPCXSSFW.close();
				bStreenClose = true;
			}else{
				_NFSHSSFW.close();
				bStreenClose = true;
			}
		}catch(FileNotFoundException fe){
			Common.addExceptionLogger(fe);
		} catch (Exception e) {
			Common.addExceptionLogger(e);
		}finally{
			if(xlsx){
				try {
					if(!bStreenClose){
						_OPCXSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				} 
			}else{
				try {
					if(!bStreenClose){
						_NFSHSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		return _sReturnObj;
	}
	
	/**<p>It will be used to Create the excel file as data with specified  ArrayList<ArrayList<String>>
	 * @author Siri Kumar Puttagunta
	 * @return boolean
	 * @param  filePath  - File path of the .xls or .xlsx file.
	 * @param  excelData   - ArrayList<ArrayList<String>> of data to write in excel file.
	 * @param  sheetName  - Name of the sheet for which data to be placed in sheet.
	 * */
	public static boolean createExcelFileWithData(String filePath, ArrayList<ArrayList<String>> excelData, String sheetName) {
		boolean _sReturnObject = false;
		try{
			if(!PoireExcelWorkbook.isSheetExist(filePath, sheetName)){
				int index = 0;
				PoireExcelWorkbook.addSheet(filePath, sheetName);
				for (ArrayList<String> row : excelData) {
					PoireExcelWorkbook.createRowListDataInExcelFile(filePath, row, index, sheetName);
					index++;
				}
				_sReturnObject = true;
			}else{
				int index = 0;
				PoireExcelWorkbook.removeSheet(filePath, sheetName);
				PoireExcelWorkbook.addSheet(filePath, sheetName);
				for (ArrayList<String> row : excelData) {
					PoireExcelWorkbook.createRowListDataInExcelFile(filePath, row, index, sheetName);
					index++;
				}
				_sReturnObject = true;
			}
		}catch(Exception e){
			Common.addExceptionLogger(e);
		}
		return _sReturnObject;
	}
		
	/**<p>It will used to Update the same data value in existing Excel sheet with specified columnPositions values.
	 * @author <a href="mailto:ssirekumar@gmail.com">Siri Kumar {@literal <ssirekumar@gmail.com>}</a> </br>
	 * @return boolean
	 * @param  filePath  - File path of the .xls or .xlsx file.
	 * @param  data   - String datavalue for specified columns
	 * @param  columnPositions - Array of Column index values which is starts from 0. Exp new int{0,2,4,5,8}
	 * @param  rowNumber  - How many lines to skip for updating the row with column index.
	 * @param  sheetNumberIndex - Number of the sheet it will start from index 0.
	 * */
	public static boolean updateDataValueInExcelFile(String filePath, String data, int[] columnPositions, int rowNumber, int sheetNumberIndex) {
		File sFilePathObj = null;
		OPCPackage _OPCXSSFW = null;
		POIFSFileSystem  _NFSHSSFW = null;
		Row row = null;
		//Row cellCountRow = null;
		Cell cell = null;
		FileOutputStream fileOut = null;
		boolean _sReturnObj = false;
		boolean bStreenClose = false;
		boolean lengthOfCell = false;
		try {
			sFilePathObj = new File(filePath);
			xlsx = PoireExcelWorkbook.excelFileType(sFilePathObj).equalsIgnoreCase("xlsx");
			if(xlsx){
				_OPCXSSFW = OPCPackage.open(new FileInputStream(sFilePathObj));
			}else{
				_NFSHSSFW = new POIFSFileSystem(new FileInputStream(sFilePathObj));
			}
			Workbook wb = xlsx ? XSSFWorkbookFactory.create(_OPCXSSFW) : WorkbookFactory.create(_NFSHSSFW);
			Sheet sh = wb.getSheetAt(sheetNumberIndex);
			int dataCellNum = 0;
			if(sh.getLastRowNum()>=rowNumber){
				//cellCountRow = sh.getRow(0);
				//dataCellNum = cellCountRow.getLastCellNum();
				dataCellNum = PoireExcelWorkbook.getColumnCount(filePath, sheetNumberIndex);
				row = sh.getRow(rowNumber);
				if(row == null){
					row = sh.createRow(rowNumber);
					for (int k=0;k<dataCellNum;k++) {
						row.createCell(k, CellType.BLANK);
					}
					for (int indexArr : columnPositions) {
						if(indexArr>=dataCellNum){
							lengthOfCell = true;
						}
					}
					if(!lengthOfCell){
						for (int pos : columnPositions) {
							cell = row.getCell(pos);
							if(cell.toString() == ""){
								cell.setCellValue(data);
							}
						}
						fileOut = new FileOutputStream(filePath);
						wb.write(fileOut);
						fileOut.flush();
						fileOut.close();
						bStreenClose = true;
						_sReturnObj = true;
					}else{
						System.err.println("WARRING: Column Positions index values are greater then the last Data Cell index");
						System.out.println("\tPlease change the Column Positions index with appropreate Data cells index");
					}
				}else{
					for (int indexArr : columnPositions) {
						if(indexArr>=dataCellNum){
							lengthOfCell = true;
						}
					}
					if(!lengthOfCell){
						for (int pos : columnPositions) {
							cell = row.createCell(pos, CellType.STRING);
							cell.setCellValue(data);
						}
						fileOut = new FileOutputStream(filePath);
						wb.write(fileOut);
						fileOut.flush();
						fileOut.close();
						bStreenClose = true;
						_sReturnObj = true;
					}else{
						System.err.println("WARRING: Column Positions index values are greater then the last Data Cell index");
						System.out.println("\tPlease change the Column Positions index with appropreate Data cells index");
					}
				}
			}else{
				_sReturnObj = false;
				System.err.println("RowNumber value is greter than the Data present in the sheet Number:"+ sheetNumberIndex);
			}
		}catch(FileNotFoundException fe){
			Common.addExceptionLogger(fe);
		} catch (Exception e) {
			Common.addExceptionLogger(e);
		}finally{
			if(xlsx){
				try {
					if(!bStreenClose){
						_OPCXSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				} 
			}else{
				try {
					if(!bStreenClose){
						_NFSHSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		return _sReturnObj;
	}
		
	/**<p>It will used to Update the same data value in existing Excel sheet with specified columnPositions values.
	 * @author <a href="mailto:ssirekumar@gmail.com">Siri Kumar {@literal <ssirekumar@gmail.com>}</a> </br>
	 * @return boolean
	 * @param  filePath  - File path of the .xls or .xlsx file.
	 * @param  data   - String datavalue for specified columns
	 * @param  columnPositions - Array of Column index values which is starts from 0. Exp new int{0,2,4,5,8}
	 * @param  rowNumber  - How many lines to skip for updating the row with column index.
	 * @param  sheetName - Name of the sheet.
	 * */
	public static boolean updateDataValueInExcelFile(String filePath, String data, int[] columnPositions, int rowNumber, String sheetName) {
		File sFilePathObj = null;
		OPCPackage _OPCXSSFW = null;
		POIFSFileSystem  _NFSHSSFW = null;
		Row row = null;
		//Row cellCountRow = null;
		Cell cell = null;
		FileOutputStream fileOut = null;
		boolean _sReturnObj = false;
		boolean bStreenClose = false;
		boolean lengthOfCell = false;
		try {
			sFilePathObj = new File(filePath);
			xlsx = PoireExcelWorkbook.excelFileType(sFilePathObj).equalsIgnoreCase("xlsx");
			if(xlsx){
				_OPCXSSFW = OPCPackage.open(new FileInputStream(sFilePathObj));
			}else{
				_NFSHSSFW = new POIFSFileSystem(new FileInputStream(sFilePathObj));
			}
			Workbook wb = xlsx ? XSSFWorkbookFactory.create(_OPCXSSFW) : WorkbookFactory.create(_NFSHSSFW);
			Sheet sh = wb.getSheet(sheetName);
			int dataCellNum = 0;
			if(sh.getLastRowNum()>=rowNumber){
				//cellCountRow = sh.getRow(0);
				//dataCellNum = cellCountRow.getLastCellNum();
				dataCellNum = PoireExcelWorkbook.getColumnCount(filePath, sheetName);
				row = sh.getRow(rowNumber);
				if(row == null){
					row = sh.createRow(rowNumber);
					for (int k=0;k<dataCellNum;k++) {
						row.createCell(k, CellType.BLANK);
					}
					for (int indexArr : columnPositions) {
						if(indexArr>=dataCellNum){
							lengthOfCell = true;
						}
					}
					if(!lengthOfCell){
						for (int pos : columnPositions) {
							cell = row.getCell(pos);
							if(cell.toString() == ""){
								cell.setCellValue(data);
							}
						}
						fileOut = new FileOutputStream(filePath);
						wb.write(fileOut);
						fileOut.flush();
						fileOut.close();
						bStreenClose = true;
						_sReturnObj = true;
					}else{
						System.err.println("WARRING: Column Positions index values are greater then the last Data Cell index");
						System.out.println("\tPlease change the Column Positions index with appropreate Data cells index");
					}
				}else{
					for (int indexArr : columnPositions) {
						if(indexArr>=dataCellNum){
							lengthOfCell = true;
						}
					}
					if(!lengthOfCell){
						for (int pos : columnPositions) {
							cell = row.createCell(pos, CellType.STRING);
							cell.setCellValue(data);
						}
						fileOut = new FileOutputStream(filePath);
						wb.write(fileOut);
						fileOut.flush();
						fileOut.close();
						bStreenClose = true;
						_sReturnObj = true;
					}else{
						System.err.println("WARRING: Column Positions index values are greater then the last Data Cell index");
						System.out.println("\tPlease change the Column Positions index with appropreate Data cells index");
					}
				}
			}else{
				_sReturnObj = false;
				System.err.println("RowNumber value is greter than the Data present in the sheet Name:"+ sheetName);
			}
			if(xlsx){
				_OPCXSSFW.close();
				bStreenClose = true;
			}else{
				_NFSHSSFW.close();
				bStreenClose = true;
			}
		}catch(FileNotFoundException fe){
			Common.addExceptionLogger(fe);
		} catch (Exception e) {
			Common.addExceptionLogger(e);
		}finally{
			if(xlsx){
				try {
					if(!bStreenClose){
						_OPCXSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				} 
			}else{
				try {
					if(!bStreenClose){
						_NFSHSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		return _sReturnObj;
	}
	
	/**<p>It will used to Update the same data value in existing Excel sheet with specified upto Column Number value.
	 *    with respected to the sheet Name.
	 * @author <a href="mailto:ssirekumar@gmail.com">Siri Kumar {@literal <ssirekumar@gmail.com>}</a> </br>
	 * @return boolean
	 * @param  filePath  - File path of the .xls or .xlsx file.
	 * @param  data   - String datavalue for specified columns
	 * @param  uptoColumnNumber - integer value of column number{0,2,4,5,8}
	 * @param  rowNumber  - Number of the row its starts from 0 index.
	 * @param  sheetName - Name of the sheet.
	 * */
	public static boolean updateDataValueInExcelFile(String filePath, String data, int uptoColumnNumber, int rowNumber, String sheetName) {
		File sFilePathObj = null;
		OPCPackage _OPCXSSFW = null;
		POIFSFileSystem  _NFSHSSFW = null;
		Row row = null;
		//Row cellCountRow = null;
		Cell cell = null;
		FileOutputStream fileOut = null;
		boolean _sReturnObj = false;
		boolean bStreenClose = false;
		try {
			sFilePathObj = new File(filePath);
			xlsx = PoireExcelWorkbook.excelFileType(sFilePathObj).equalsIgnoreCase("xlsx");
			if(xlsx){
				_OPCXSSFW = OPCPackage.open(new FileInputStream(sFilePathObj));
			}else{
				_NFSHSSFW = new POIFSFileSystem(new FileInputStream(sFilePathObj));
			}
			Workbook wb = xlsx ? XSSFWorkbookFactory.create(_OPCXSSFW) : WorkbookFactory.create(_NFSHSSFW);
			Sheet sh = wb.getSheet(sheetName);
			int dataCellNum = 0;
			if (rowNumber<0) {
				System.err.println("WARRING: Given rowNumber is not appropreate it should be a positive Number");
				return false;
			} 
			if(sh.getLastRowNum()>=rowNumber){
				//cellCountRow = sh.getRow(0);
				//dataCellNum = cellCountRow.getLastCellNum();
				dataCellNum = uptoColumnNumber;
				row = sh.getRow(rowNumber);
				if(row == null){
					row = sh.createRow(rowNumber);
					for (int k=0;k<=dataCellNum;k++) {
						cell = row.createCell(k, CellType.STRING);
						cell.setCellValue((String) data);
					}
				}else{
					for (int k=0;k<=dataCellNum;k++) {
						cell = row.createCell(k, CellType.STRING);
						cell.setCellValue((String) data);
					}
				}
			}else if(sh.getLastRowNum()<rowNumber){
				dataCellNum = uptoColumnNumber;
				row = sh.getRow(rowNumber);
				if(row == null){
					row = sh.createRow(rowNumber);
					for (int k=0;k<=dataCellNum;k++) {
						cell = row.createCell(k, CellType.STRING);
						cell.setCellValue((String) data);
					}
				}else{
					for (int k=0;k<=dataCellNum;k++) {
						cell = row.createCell(k, CellType.STRING);
						cell.setCellValue((String) data);
					}
				}
			}
			fileOut = new FileOutputStream(filePath);
			wb.write(fileOut);
			fileOut.flush();
			fileOut.close();
			bStreenClose = true;
			_sReturnObj = true;
			if(xlsx){
				_OPCXSSFW.close();
				bStreenClose = true;
			}else{
				_NFSHSSFW.close();
				bStreenClose = true;
			}
		}catch(FileNotFoundException fe){
			Common.addExceptionLogger(fe);
		} catch (Exception e) {
			Common.addExceptionLogger(e);
		}finally{
			if(xlsx){
				try {
					if(!bStreenClose){
						_OPCXSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				} 
			}else{
				try {
					if(!bStreenClose){
						_NFSHSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		return _sReturnObj;
	}
	
	/**<p>It will used to Update the same data value in existing Excel sheet with specified upto Column Number value.
	 *    with respected to the sheet Number Index
	 * @author <a href="mailto:ssirekumar@gmail.com">Siri Kumar {@literal <ssirekumar@gmail.com>}</a> </br>
	 * @return boolean
	 * @param  filePath  - File path of the .xls or .xlsx file.
	 * @param  data   - String datavalue for specified columns
	 * @param  uptoColumnNumber - integer value of column number
	 * @param  rowNumber  - Number of the row its starts from 0 index.
	 * @param  sheetNumberIndex - Index number of the sheet it will start from 0.
	 * */
	public static boolean updateDataValueInExcelFile(String filePath, String data, int uptoColumnNumber, int rowNumber, int sheetNumberIndex) {
		File sFilePathObj = null;
		OPCPackage _OPCXSSFW = null;
		POIFSFileSystem  _NFSHSSFW = null;
		Row row = null;
		//Row cellCountRow = null;
		Cell cell = null;
		FileOutputStream fileOut = null;
		boolean _sReturnObj = false;
		boolean bStreenClose = false;
		try {
			sFilePathObj = new File(filePath);
			xlsx = PoireExcelWorkbook.excelFileType(sFilePathObj).equalsIgnoreCase("xlsx");
			if(xlsx){
				_OPCXSSFW = OPCPackage.open(new FileInputStream(sFilePathObj));
			}else{
				_NFSHSSFW = new POIFSFileSystem(new FileInputStream(sFilePathObj));
			}
			Workbook wb = xlsx ? XSSFWorkbookFactory.create(_OPCXSSFW) : WorkbookFactory.create(_NFSHSSFW);
			Sheet sh = wb.getSheetAt(sheetNumberIndex);
			int dataCellNum = 0;
			if (rowNumber<0) {
				System.err.println("WARRING: Given rowNumber is not appropreate it should be a positive Number");
				return false;
			} 
			if(sh.getLastRowNum()>=rowNumber){
				//cellCountRow = sh.getRow(0);
				//dataCellNum = cellCountRow.getLastCellNum();
				dataCellNum = uptoColumnNumber;
				row = sh.getRow(rowNumber);
				if(row == null){
					row = sh.createRow(rowNumber);
					for (int k=0;k<=dataCellNum;k++) {
						cell = row.createCell(k, CellType.STRING);
						cell.setCellValue((String) data);
					}
				}else{
					for (int k=0;k<=dataCellNum;k++) {
						cell = row.createCell(k, CellType.STRING);
						cell.setCellValue((String) data);
					}
				}
			}else if(sh.getLastRowNum()<rowNumber){
				dataCellNum = uptoColumnNumber;
				row = sh.getRow(rowNumber);
				if(row == null){
					row = sh.createRow(rowNumber);
					for (int k=0;k<=dataCellNum;k++) {
						cell = row.createCell(k, CellType.STRING);
						cell.setCellValue((String) data);
					}
				}else{
					for (int k=0;k<=dataCellNum;k++) {
						cell = row.createCell(k, CellType.STRING);
						cell.setCellValue((String) data);
					}
				}
			}
			fileOut = new FileOutputStream(filePath);
			wb.write(fileOut);
			fileOut.flush();
			fileOut.close();
			bStreenClose = true;
			_sReturnObj = true;
			if(xlsx){
				_OPCXSSFW.close();
				bStreenClose = true;
			}else{
				_NFSHSSFW.close();
				bStreenClose = true;
			}
		}catch(FileNotFoundException fe){
			Common.addExceptionLogger(fe);
		} catch (Exception e) {
			Common.addExceptionLogger(e);
		}finally{
			if(xlsx){
				try {
					if(!bStreenClose){
						_OPCXSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				} 
			}else{
				try {
					if(!bStreenClose){
						_NFSHSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		return _sReturnObj;
	}
		
	/**<p>It will used to get the row data with specified column positions.
	 * @author <a href="mailto:ssirekumar@gmail.com">Siri Kumar {@literal <ssirekumar@gmail.com>}</a> </br>
	 * @return {@code ArrayList<String>}
	 * @param  filePath  - File path of the .xls or .xlsx file.
	 * @param  columnPositions - Array of Column index values which is starts from 0. Exp new int{0,2,4,5,8}
	 * @param  skipLines  - How many lines to skip for updating the row with column index.
	 * @param  rowNumber - Number of the row.
	 * @param  sheetNumberIndex - integer number of the sheet Number Index it will start from 0. 
	 * */
	public static ArrayList<String> getExcelRowWithSpecifiedColumn(String filePath, int[] columnPositions, int rowNumber, int sheetNumberIndex) {
		File sFilePathObj = null;
		OPCPackage _OPCXSSFW = null;
		POIFSFileSystem  _NFSHSSFW = null;
		Row row = null;
		//Row cellCountRow = null;
		Cell cell = null;
		boolean bStreenClose = false;
		boolean lengthOfCell = false;
		ArrayList<String> listArray = null;
		try {
			sFilePathObj = new File(filePath);
			xlsx = PoireExcelWorkbook.excelFileType(sFilePathObj).equalsIgnoreCase("xlsx");
			if(xlsx){
				_OPCXSSFW = OPCPackage.open(sFilePathObj);
			}else{
				_NFSHSSFW = new POIFSFileSystem(sFilePathObj);
			}
			Workbook wb = xlsx ? XSSFWorkbookFactory.create(_OPCXSSFW) : WorkbookFactory.create(_NFSHSSFW);
			Sheet sh = wb.getSheetAt(sheetNumberIndex);
			int dataCellNum = 0;
			String cellValue = "";
			listArray = new ArrayList<String>();
			if(sh.getLastRowNum()>rowNumber){
				//cellCountRow = sh.getRow(0);
				//dataCellNum = cellCountRow.getLastCellNum();
				dataCellNum = PoireExcelWorkbook.getColumnCount(filePath, sheetNumberIndex);
				row = sh.getRow(rowNumber);
				if(row == null){
					row = sh.createRow(rowNumber);
					for (int k=0;k<=dataCellNum;k++) {
						row.createCell(k, CellType.BLANK);
					}
					for (int indexArr : columnPositions) {
						if(indexArr>=dataCellNum){
							lengthOfCell = true;
						}
					}
                    if(!lengthOfCell){
                    	for (int pos : columnPositions) {
                    		cell = row.getCell(pos);
                    		if(cell.equals(null)){
                    			listArray.add("");
                    			continue;
                    		}
                    		cellValue = PoireExcelWorkbook.cellFormatedDataValue(cell);
    						listArray.add(cellValue);
    					}
                    }else{
                    	System.err.println("WARRING: Column Positions index values are greater then the last Data Cell index");
						System.out.println("\tPlease change the Column Positions index with appropreate Data cells index");
						listArray.add("");
                    }
				}else{
					for (int indexArr : columnPositions) {
						if(indexArr>=dataCellNum){
							lengthOfCell = true;
						}
					}
					if(!lengthOfCell){
						for (int pos : columnPositions) {
							cell = row.getCell(pos);
							if(cell == null){
                    			listArray.add("");
                    			continue;
                    		}
							cellValue = PoireExcelWorkbook.cellFormatedDataValue(cell);
    						listArray.add(cellValue);
						}
					}else{
						System.err.println("WARRING: Column Positions index values are greater then the last Data Cell index");
						System.out.println("\tPlease change the Column Positions index with appropreate Data cells index");
						listArray.add("");
					}
				}
			}else{
				System.err.println("SkipLines value is greter than the Data present in the sheet with index:"+ sheetNumberIndex);
			}
			if(xlsx){
				_OPCXSSFW.close();
				bStreenClose = true;
			}else{
				_NFSHSSFW.close();
				bStreenClose = true;
			}
		}catch(FileNotFoundException fe){
			Common.addExceptionLogger(fe);
		} catch (Exception e) {
			Common.addExceptionLogger(e);
		}finally{
			if(xlsx){
				try {
					if(!bStreenClose){
						_OPCXSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				} 
			}else{
				try {
					if(!bStreenClose){
						_NFSHSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		return listArray;
	}
	
	/**<p>It will used to get the row data with specified column positions.
	 * @author <a href="mailto:ssirekumar@gmail.com">Siri Kumar {@literal <ssirekumar@gmail.com>}</a> </br>
	 * @return {@code ArrayList<String>}
	 * @param  filePath  - File path of the .xls or .xlsx file.
	 * @param  columnPositions - Array of Column index values which is starts from 0. Exp new int{0,2,4,5,8}
	 * @param  skipLines  - How many lines to skip for updating the row with column index.
	 * @param  rowNumber - Number of the row.
	 * @param  sheetNumberIndex - integer number of the sheet Number Index it will start from 0. 
	 * */
	public static ArrayList<String> getExcelRowWithSpecifiedColumn(String filePath, int[] columnPositions, int rowNumber, String sheetName) {
		File sFilePathObj = null;
		OPCPackage _OPCXSSFW = null;
		POIFSFileSystem  _NFSHSSFW = null;
		Row row = null;
		//Row cellCountRow = null;
		Cell cell = null;
		boolean bStreenClose = false;
		boolean lengthOfCell = false;
		ArrayList<String> listArray = null;
		try {
			sFilePathObj = new File(filePath);
			xlsx = PoireExcelWorkbook.excelFileType(sFilePathObj).equalsIgnoreCase("xlsx");
			if(xlsx){
				_OPCXSSFW = OPCPackage.open(sFilePathObj);
			}else{
				_NFSHSSFW = new POIFSFileSystem(sFilePathObj);
			}
			Workbook wb = xlsx ? XSSFWorkbookFactory.create(_OPCXSSFW) : WorkbookFactory.create(_NFSHSSFW);
			Sheet sh = wb.getSheet(sheetName);
			int dataCellNum = 0;
			String cellValue = "";
			listArray = new ArrayList<String>();
			if(sh.getLastRowNum()>rowNumber){
				//cellCountRow = sh.getRow(0);
				//dataCellNum = cellCountRow.getLastCellNum();
				dataCellNum = PoireExcelWorkbook.getColumnCount(filePath, sheetName);
				row = sh.getRow(rowNumber);
				if(row == null){
					row = sh.createRow(rowNumber);
					for (int k=0;k<=dataCellNum;k++) {
						row.createCell(k, CellType.BLANK);
					}
					for (int indexArr : columnPositions) {
						if(indexArr>=dataCellNum){
							lengthOfCell = true;
						}
					}
                    if(!lengthOfCell){
                    	for (int pos : columnPositions) {
                    		cell = row.getCell(pos);
                    		if(cell == null){
                    			listArray.add("");
                    			continue;
                    		}
                    		cellValue = PoireExcelWorkbook.cellFormatedDataValue(cell);
    						listArray.add(cellValue);
    					}
                    }else{
                    	System.err.println("WARRING: Column Positions index values are greater then the last Data Cell index");
						System.out.println("\tPlease change the Column Positions index with appropreate Data cells index");
						listArray.add("");
                    }
				}else{
					for (int indexArr : columnPositions) {
						if(indexArr>=dataCellNum){
							lengthOfCell = true;
						}
					}
					if(!lengthOfCell){
						for (int pos : columnPositions) {
							cell = row.getCell(pos);
							if(cell == null){
                    			listArray.add("");
                    			continue;
                    		}
							cellValue = PoireExcelWorkbook.cellFormatedDataValue(cell);
    						listArray.add(cellValue);
						}
					}else{
						System.err.println("WARRING: Column Positions index values are greater then the last Data Cell index");
						System.out.println("\tPlease change the Column Positions index with appropreate Data cells index");
						listArray.add("");
					}
				}
			}else{
				System.err.println("SkipLines value is greter than the Data present in the sheet with Sheet Name:"+ sheetName);
			}
			if(xlsx){
				_OPCXSSFW.close();
				bStreenClose = true;
			}else{
				_NFSHSSFW.close();
				bStreenClose = true;
			}
		}catch(FileNotFoundException fe){
			Common.addExceptionLogger(fe);
		} catch (Exception e) {
			Common.addExceptionLogger(e);
		}finally{
			if(xlsx){
				try {
					if(!bStreenClose){
						_OPCXSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				} 
			}else{
				try {
					if(!bStreenClose){
						_NFSHSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		return listArray;
	}
	
	/**<p>It will used to get the Excel existing sheet data as a {@code ArrayList<ArrayList<String>>}.
	 * @author <a href="mailto:ssirekumar@gmail.com">Siri Kumar {@literal <ssirekumar@gmail.com>}</a> </br>
	 * @return {@code ArrayList<ArrayList<String>>}
	 * @param  filePath  - File path of the .xls or .xlsx file.
	 * @param  skipLines  - How many lines to skip for geting the data from the row.
	 * @param  sheetNumberIndex - integer number of the sheet Number Index it will start from 0. 
	 * */
	public static ArrayList<ArrayList<String>> getExcelDataWithSheetNumber(String filePath, int skipLines, int sheetNumberIndex) {
		File sFilePathObj = null;
		OPCPackage _OPCXSSFW = null;
		POIFSFileSystem  _NFSHSSFW = null;
		Row row = null;
		String[] rowContents = null;
		//Row cellCountRow = null;
		Cell cell = null;
		boolean bStreenClose = false;
		ArrayList<ArrayList<String>> listArray = null;
		ArrayList<String> rowIndexData = null;
		try {
			sFilePathObj = new File(filePath);
			xlsx = PoireExcelWorkbook.excelFileType(sFilePathObj).equalsIgnoreCase("xlsx");
			if(xlsx){
				_OPCXSSFW = OPCPackage.open(sFilePathObj);
			}else{
				_NFSHSSFW = new POIFSFileSystem(sFilePathObj);
			}
			Workbook wb = xlsx ? XSSFWorkbookFactory.create(_OPCXSSFW) : WorkbookFactory.create(_NFSHSSFW);
			Sheet sh = wb.getSheetAt(sheetNumberIndex);
			int iLastRowNumber = sh.getLastRowNum();
			int dataCellNum = 0;
			/*cellCountRow = sh.getRow(0);
			dataCellNum = cellCountRow.getLastCellNum();*/
			dataCellNum = PoireExcelWorkbook.getColumnCount(filePath, sheetNumberIndex);
			rowIndexData = new ArrayList<String>();
			listArray = new ArrayList<ArrayList<String>>();
			
			if(sh.getLastRowNum()>skipLines){
				int totalSkiplines = skipLines+1;
				rowContents = new String[dataCellNum+1];
				for (int i=totalSkiplines;i<=iLastRowNumber + 1;i++) {
					row = sh.getRow(i);
					if(row == null){
						row = sh.createRow(i);
						for (int k=0;k<=dataCellNum;k++) {
							row.createCell(k, CellType.BLANK);
							cell = row.getCell(k);
							if(cell.toString() == ""){
                    			rowContents[k] = "";
                    		}
						}
						rowIndexData = new ArrayList<String>(Arrays.asList(rowContents));
						listArray.add(rowIndexData);
					}else{
						for (int k=0;k<=dataCellNum;k++) {
							cell = row.getCell(k);
                    		if(cell == null){
                    			rowContents[k] = "";
                    		}else{
                    			rowContents[k] = PoireExcelWorkbook.cellFormatedDataValue(cell);                   			
                    			//rowIndexData.add(cell.toString());
                    		}
						}
						rowIndexData = new ArrayList<String>(Arrays.asList(rowContents));
						listArray.add(rowIndexData);
					}
				}
			}else{
				System.err.println("SkipLines value is greter than the Data present in the sheet with index:"+ sheetNumberIndex);
			}
			
			if(xlsx){
				_OPCXSSFW.close();
				bStreenClose = true;
			}else{
				_NFSHSSFW.close();
				bStreenClose = true;
			}
		}catch(FileNotFoundException fe){
			Common.addExceptionLogger(fe);
		} catch (Exception e) {
			Common.addExceptionLogger(e);
		}finally{
			if(xlsx){
				try {
					if(!bStreenClose){
						_OPCXSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				} 
			}else{
				try {
					if(!bStreenClose){
						_NFSHSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		return listArray;
	}
	
	/**<p>It will used to get the Excel existing sheet data as a {@code java.util.ArrayList<ArrayList<String>>}.
	 * @author <a href="mailto:ssirekumar@gmail.com">Siri Kumar {@literal <ssirekumar@gmail.com>}</a> </br>
	 * @return {@code java.util.ArrayList<ArrayList<String>>}
	 * @param  filePath  - File path of the .xls or .xlsx file.
	 * @param  skipLines  - How many lines to skip for geting the data from the row.
	 * @param  sheetName - String name of the sheet. 
	 * */
	public static ArrayList<ArrayList<String>> getExcelDataWithSheetName(String filePath, int skipLines, String sheetName) {
		File sFilePathObj = null;
		OPCPackage _OPCXSSFW = null;
		POIFSFileSystem  _NFSHSSFW = null;
		Row row = null;
		String[] rowContents = null;
		//Row cellCountRow = null;
		Cell cell = null;
		boolean bStreenClose = false;
		ArrayList<ArrayList<String>> listArray = null;
		ArrayList<String> rowIndexData = null;
		try {
			sFilePathObj = new File(filePath);
			xlsx = PoireExcelWorkbook.excelFileType(sFilePathObj).equalsIgnoreCase("xlsx");
			if(xlsx){
				_OPCXSSFW = OPCPackage.open(sFilePathObj);
			}else{
				_NFSHSSFW = new POIFSFileSystem(sFilePathObj);
			}
			Workbook wb = xlsx ? XSSFWorkbookFactory.create(_OPCXSSFW) : WorkbookFactory.create(_NFSHSSFW);
			Sheet sh = wb.getSheet(sheetName);
			int iLastRowNumber = sh.getLastRowNum();
			int dataCellNum = 0;
			//cellCountRow = sh.getRow(0);
			//dataCellNum = cellCountRow.getLastCellNum();
			dataCellNum = PoireExcelWorkbook.getColumnCount(filePath, sheetName);
			rowIndexData = new ArrayList<String>();
			listArray = new ArrayList<ArrayList<String>>();
			if(sh.getLastRowNum()>skipLines){
				int totalSkiplines = skipLines+1;
				rowContents = new String[dataCellNum+1];
				for (int i=totalSkiplines;i<=iLastRowNumber + 1;i++) {
					row = sh.getRow(i);
					if(row == null){
						row = sh.createRow(i);
						for (int k=0;k<=dataCellNum;k++) {
							row.createCell(k, CellType.BLANK);
							cell = row.getCell(k);
							if(cell.toString() == ""){
                    			rowContents[k] = "";
                    		}
						}
						rowIndexData = new ArrayList<String>(Arrays.asList(rowContents));
						listArray.add(rowIndexData);
					}else{
						for (int k=0;k<=dataCellNum;k++) {
							cell = row.getCell(k);
                    		if(cell == null){
                    			rowContents[k] = "";
                    		}else{
                    			rowContents[k] = PoireExcelWorkbook.cellFormatedDataValue(cell);                   			
                    			//rowIndexData.add(cell.toString());
                    		}
						}
						rowIndexData = new ArrayList<String>(Arrays.asList(rowContents));
						listArray.add(rowIndexData);
					}
				}
			}else{
				System.err.println("SkipLines value is greter than the Data present in the sheet Name:"+ sheetName);
			}
			if(xlsx){
				_OPCXSSFW.close();
				bStreenClose = true;
			}else{
				_NFSHSSFW.close();
				bStreenClose = true;
			}
		}catch(FileNotFoundException fe){
			Common.addExceptionLogger(fe);
		} catch (Exception e) {
			Common.addExceptionLogger(e);
		}finally{
			if(xlsx){
				try {
					if(!bStreenClose){
						_OPCXSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				} 
			}else{
				try {
					if(!bStreenClose){
						_NFSHSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		return listArray;
	}
				
	/**<p>It will used to create the Excel file in given File Format
	 * @author <a href="mailto:ssirekumar@gmail.com">Siri Kumar {@literal <ssirekumar@gmail.com>}</a> </br>
	 * @param  xlsxfileFormat  - Boolean value for .xlsx format true for .xlsx file.
	 * @param  filePath  - Directory location path for to create the excel file.     
	 * @param  fileName  - Name of the file for excel file in the Directory. 
	 * @param  sheetName - Default sheet name while creation of the Excel file.
	 * @return {@code java.io.File}
	 *  */
	public static File createExcelFile(boolean xlsxfileFormat, String filePath, String fileName, String sheetName) {
		FileOutputStream fileOut = null;
		File _sReturnFileObj = null;
		File _sFilePathObj = null;
		try {
			xlsx = xlsxfileFormat;
			System.out.println("Creating the Excel file");
			if(xlsx){
				_sFilePathObj = new File(filePath + "\\"+ fileName + ".xlsx");
				fileOut = new FileOutputStream(filePath + "\\"+ fileName + ".xlsx");
			}else{
				_sFilePathObj = new File(filePath + "\\"+ fileName + ".xls");
				fileOut = new FileOutputStream(filePath + "\\"+ fileName + ".xls");
			}
			Workbook wb = xlsx ? new XSSFWorkbook() : new HSSFWorkbook();
			String safeName = WorkbookUtil.createSafeSheetName(sheetName); // sheetName = "[O'Brien's sales*?]" returns " O'Brien's sales   "
			wb.createSheet(safeName);
			wb.write(fileOut);
			fileOut.flush();
			fileOut.close();
			_sReturnFileObj = _sFilePathObj.getAbsoluteFile();
			if(xlsx){
				System.out.println("Created " + filePath + "\\"+ fileName + ".xlsx" + " Excel file");
			}else{
				System.out.println("Created " + filePath + "\\"+ fileName + ".xls" + " Excel file");
			}
		} catch (IOException Io) {
			Common.addExceptionLogger(Io);
		} catch (Exception e){
			Common.addExceptionLogger(e);
		}finally {
			if (fileOut != null){
				try {
					fileOut.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		return _sReturnFileObj;
	}
	
	/**<p>It will used to create the Excel file with Number of Sheets specified in given File Format
	 * @author <a href="mailto:ssirekumar@gmail.com">Siri Kumar {@literal <ssirekumar@gmail.com>}</a> </br>
	 * @param  xlsxfileFormat  - Boolean value for .xlsx format true for .xlsx file.
	 * @param  filePath  - Directory location path for to create the excel file.     
	 * @param  fileName  - Name of the file for excel file in the Directory. 
	 * @param  sheetNames - Array of sheet names while creation of the Excel file.
	 *  */
	public static File createExcelFileWithSheetNames(boolean xlsxfileFormat, String filePath, String fileName, String[] sheetNames){
		FileOutputStream fileOut = null;
		File _sReturnFileObj = null;
		File _sFilePathObj = null;
		try {
			xlsx = xlsxfileFormat;
			System.out.println("Creating the Excel file");
			if(xlsx){
				_sFilePathObj = new File(filePath + "\\"+ fileName + ".xlsx");
				fileOut = new FileOutputStream(filePath + "\\"+ fileName + ".xlsx");
			}else{
				_sFilePathObj = new File(filePath + "\\"+ fileName + ".xls");
				fileOut = new FileOutputStream(filePath + "\\"+ fileName + ".xls");
			}
			Workbook wb = xlsx ? new XSSFWorkbook() : new HSSFWorkbook();
			for (String sheetName : sheetNames) {
				String safeName = WorkbookUtil.createSafeSheetName(sheetName); // sheetName = "[O'Brien's sales*?]" returns " O'Brien's sales   "
				wb.createSheet(safeName);
			}
			wb.write(fileOut);
			fileOut.flush();
			fileOut.close();
			_sReturnFileObj = _sFilePathObj.getAbsoluteFile();
			if(xlsx){
				System.out.println("Created " + filePath + "\\"+ fileName + ".xlsx" + " Excel file");
			}else{
				System.out.println("Created " + filePath + "\\"+ fileName + ".xls" + " Excel file");
			}
		} catch (IOException Io) {
			Common.addExceptionLogger(Io);
		} catch (Exception e){
			Common.addExceptionLogger(e);
		}finally {
			if (fileOut != null){
				try {
					fileOut.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		return _sReturnFileObj;
	}
				
	/**<p>It will be used to return the number of Data presented rows in the sheet.
	 * @author <a href="mailto:ssirekumar@gmail.com">Siri Kumar {@literal <ssirekumar@gmail.com>}</a> </br>
	 * @return it will return the integer value of the data rows in the sheet.
	 * @param  filePath  - File path of the .xls or .xlsx file.   
	 * @param  sheetIndexNumber - Sheet Number index which is presented from strts from 0.
	 *  */
	public static int getRowCount(String filePath, int sheetIndexNumber) {
		File sFilePathObj = null;
		OPCPackage _OPCXSSFW = null;
		POIFSFileSystem  _NFSHSSFW = null;
		boolean bStreenClose = false;
		int _sReturnObj = 0;
		try {
			sFilePathObj = new File(filePath);
			xlsx = PoireExcelWorkbook.excelFileType(sFilePathObj).equalsIgnoreCase("xlsx");
			if(xlsx){
				_OPCXSSFW = OPCPackage.open(sFilePathObj);
			}else{
				_NFSHSSFW = new POIFSFileSystem(sFilePathObj);
			}
			Workbook wb = xlsx ? XSSFWorkbookFactory.create(_OPCXSSFW) : WorkbookFactory.create(_NFSHSSFW);
			Sheet sh = wb.getSheetAt(sheetIndexNumber);
			int iLastRowNumber = sh.getLastRowNum();
			_sReturnObj = iLastRowNumber + 1;
			if(xlsx){
				_OPCXSSFW.close();
				bStreenClose = true;
			}else{
				_NFSHSSFW.close();
				bStreenClose = true;
			}
		} catch (Exception e) {
			Common.addExceptionLogger(e);
		}finally{
			if(xlsx){
				try {
					if(!bStreenClose){
						_OPCXSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				} 
			}else{
				try {
					if(!bStreenClose){
						_NFSHSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		return _sReturnObj;
	}
	
	/**<p>It will be used to return the number of Data presented rows in the sheet.
	 * @author <a href="mailto:ssirekumar@gmail.com">Siri Kumar {@literal <ssirekumar@gmail.com>}</a> </br>
	 * @return it will return the integer value of the data rows in the sheet
	 * @param  filePath  - File path of the .xls or .xlsx file.   
	 * @param  sheetName - Sheet Name of the file.
	 *  */
	public static int getRowCount(String filePath, String sheetName) {
		File sFilePathObj = null;
		OPCPackage _OPCXSSFW = null;
		POIFSFileSystem  _NFSHSSFW = null;
		boolean bStreenClose = false;
		int _sReturnObj = 0;
		try {
			sFilePathObj = new File(filePath);
			xlsx = PoireExcelWorkbook.excelFileType(sFilePathObj).equalsIgnoreCase("xlsx");
			if(xlsx){
				_OPCXSSFW = OPCPackage.open(sFilePathObj);
			}else{
				_NFSHSSFW = new POIFSFileSystem(sFilePathObj);
			}
			Workbook wb = xlsx ? XSSFWorkbookFactory.create(_OPCXSSFW) : WorkbookFactory.create(_NFSHSSFW);
			Sheet sh = wb.getSheet(sheetName);
			int iLastRowNumber = sh.getLastRowNum();
			_sReturnObj = iLastRowNumber + 1;
			if(xlsx){
				_OPCXSSFW.close();
				bStreenClose = true;
			}else{
				_NFSHSSFW.close();
				bStreenClose = true;
			}
		} catch (Exception e) {
			Common.addExceptionLogger(e);
		}finally{
			if(xlsx){
				try {
					if(!bStreenClose){
						_OPCXSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				} 
			}else{
				try {
					if(!bStreenClose){
						_NFSHSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		return _sReturnObj;
	}
    		
	/**<p>It will return the gretest positive number of Data presented Columns in the sheet.
	 *    This count is from cell value 0 index.
	 * @author <a href="mailto:ssirekumar@gmail.com">Siri Kumar {@literal <ssirekumar@gmail.com>}</a> </br>
	 * @return It will return the positive integer of gretest value from row having a column data. 
	 *         If data columns are not prasent it will return -1. 
	 * @param  filePath  - File path of the .xls or .xlsx file.   
	 * @param  sheetName - Sheet Name of the file.
	 * @ 
	 *  */
	public static int getColumnCount(String filePath, String sheetName) {  
		int _sReturnObj = -1;
		int largerColumnValue;
		ArrayList<Integer> arrayRowCellValue = new ArrayList<Integer>();
		try {
			int rowCount = PoireExcelWorkbook.getRowCount(filePath,sheetName);
			for (int i = 0; i <= rowCount; i++) {
				ArrayList<String> _rowdata = PoireExcelWorkbook.getExcelRowWithSheetName(filePath, i, sheetName, true);
				if(_rowdata.size() != 0 && _rowdata != null){
					arrayRowCellValue.add(_rowdata.size()-2);
				}
			}
			largerColumnValue = Collections.max(arrayRowCellValue).intValue();
			_sReturnObj = largerColumnValue;
		} catch (Exception e) {
			Common.addExceptionLogger(e);
		}
		return _sReturnObj;
	}
	
	/**<p>It will return the gretest number of Data presented Columns in the sheet.
	 * this count is from cell value 0 index.
	 * @author <a href="mailto:ssirekumar@gmail.com">Siri Kumar {@literal <ssirekumar@gmail.com>}</a> </br>
	 * @return It will return the positive integer of gretest value from row having a column data. 
	 *         If data columns are not prasent it will return -1.  
	 * @param  filePath  - File path of the .xls or .xlsx file.   
	 * @param  sheetNumberIndex - Sheet Number index, Index number its starts from 0.
	 * @ 
	 *  */
	public static int getColumnCount(String filePath, int sheetNumberIndex) {  
		int _sReturnObj = -1;
		int largerColumnValue;
		ArrayList<Integer> arrayRowCellValue = new ArrayList<Integer>();
		try {
			int rowCount = PoireExcelWorkbook.getRowCount(filePath,sheetNumberIndex);
			for (int i = 0; i <= rowCount; i++) {
				ArrayList<String> _rowdata = PoireExcelWorkbook.getExcelRowWithSheetIndex(filePath, i, sheetNumberIndex, true);
				if(_rowdata.size() != 0 && _rowdata != null){
					arrayRowCellValue.add(_rowdata.size()-2);
				}
			}
			largerColumnValue = Collections.max(arrayRowCellValue).intValue();
			_sReturnObj = largerColumnValue;
		} catch (Exception e) {
			Common.addExceptionLogger(e);
		}
		return _sReturnObj;
	}
	
	/**<p>It will return the int Data present of the row in the columns.
	 * this count is from cell value 0 index.
	 * @author <a href="mailto:ssirekumar@gmail.com">Siri Kumar {@literal <ssirekumar@gmail.com>}</a> </br>
	 * @return It will return the positive integer of column count with specified row having data columns. 
	 *         If data columns are not prasent it will return -1.  
	 * @param  filePath  - File path of the .xls or .xlsx file.   
	 * @param  rowNumber  - int number of the row
	 * @param  sheetNumberIndex - Sheet Number index, Index number its starts from 0.
	 * @ 
	 *  */
	public static int getColumnCountOfSpecifiedRow(String filePath, int rowNumber, int sheetNumberIndex) {  // Method is not completed.
		File sFilePathObj = null;
		OPCPackage _OPCXSSFW = null;
		POIFSFileSystem  _NFSHSSFW = null;
		boolean bStreenClose = false;
		Row cellCountRow = null;
		short _cellsNumber = -1;
		try {
			sFilePathObj = new File(filePath);
			xlsx = PoireExcelWorkbook.excelFileType(sFilePathObj).equalsIgnoreCase("xlsx");
			if(xlsx){
				_OPCXSSFW = OPCPackage.open(new FileInputStream(sFilePathObj));
			}else{
				_NFSHSSFW = new POIFSFileSystem(new FileInputStream(sFilePathObj));
			}
			Workbook wb = xlsx ? XSSFWorkbookFactory.create(_OPCXSSFW) : WorkbookFactory.create(_NFSHSSFW);
			Sheet sh = wb.getSheetAt(sheetNumberIndex); 
			int iLastRowNumber = sh.getLastRowNum();
			if(iLastRowNumber>=rowNumber){
				cellCountRow = sh.getRow(rowNumber);
				if(cellCountRow != null){
					_cellsNumber = cellCountRow.getLastCellNum();
					return _cellsNumber; 
				}
			}else{
				System.err.println("RowNumber value is greter than the Data present in the sheet with rowNumber:"+ rowNumber);
			}
			if(xlsx){
				_OPCXSSFW.close();
				bStreenClose = true;
			}else{
				_NFSHSSFW.close();
				bStreenClose = true;
			}
		} catch (Exception e) {
			Common.addExceptionLogger(e);
		}finally{
			if(xlsx){
				try {
					if(!bStreenClose){
						_OPCXSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				} 
			}else{
				try {
					if(!bStreenClose){
						_NFSHSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		return _cellsNumber;
	}
	
	/**<p>It will return the sheet object with specified sheet Number Index
	 * @author <a href="mailto:ssirekumar@gmail.com">Siri Kumar {@literal <ssirekumar@gmail.com>}</a> </br>
	 * @return Sheet  - it will return the sheet class object.  
	 * @param  filePath  - File path of the .xls or .xlsx file.   
	 * @param  int sheetNumberIndex - Sheet Number index, Index number its starts from 0.
	 * @ 
	 *  */
	public static Sheet getSheet(String filePath, int sheetNumberIndex) {  
		File sFilePathObj = null;
		OPCPackage _OPCXSSFW = null;
		POIFSFileSystem  _NFSHSSFW = null;
		boolean bStreenClose = false;
		Sheet sh = null;
		try {
			sFilePathObj = new File(filePath);
			xlsx = PoireExcelWorkbook.excelFileType(sFilePathObj).equalsIgnoreCase("xlsx");
			if(xlsx){
				_OPCXSSFW = OPCPackage.open(sFilePathObj);
			}else{
				_NFSHSSFW = new POIFSFileSystem(sFilePathObj);
			}
			Workbook wb = xlsx ? XSSFWorkbookFactory.create(_OPCXSSFW) : WorkbookFactory.create(_NFSHSSFW);
			sh = wb.getSheetAt(sheetNumberIndex);
			if(sh != null){
				return sh;
			}
			if(xlsx){
				_OPCXSSFW.close();
				bStreenClose = true;
			}else{
				_NFSHSSFW.close();
				bStreenClose = true;
			}
		} catch (Exception e) {
			Common.addExceptionLogger(e);
		}finally{
			if(xlsx){
				try {
					if(!bStreenClose){
						_OPCXSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				} 
			}else{
				try {
					if(!bStreenClose){
						_NFSHSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		return sh;
	}
		
	/**<p>It will return the sheet object with specified sheet Name.
	 * @author <a href="mailto:ssirekumar@gmail.com">Siri Kumar {@literal <ssirekumar@gmail.com>}</a> </br>
	 * @return Sheet  - it will return the sheet class object.  
	 * @param  filePath  - File path of the .xls or .xlsx file.   
	 * @param  sheetName - Name of the Sheet.
	 * */
	public static Sheet getSheet(String filePath, String sheetName) {  
		File sFilePathObj = null;
		OPCPackage _OPCXSSFW = null;
		POIFSFileSystem  _NFSHSSFW = null;
		boolean bStreenClose = false;
		Sheet sh = null;
		try {
			sFilePathObj = new File(filePath);
			xlsx = PoireExcelWorkbook.excelFileType(sFilePathObj).equalsIgnoreCase("xlsx");
			if(xlsx){
				_OPCXSSFW = OPCPackage.open(sFilePathObj);
			}else{
				_NFSHSSFW = new POIFSFileSystem(sFilePathObj);
			}
			Workbook wb = xlsx ? XSSFWorkbookFactory.create(_OPCXSSFW) : WorkbookFactory.create(_NFSHSSFW);
			sh = wb.getSheet(sheetName);
			if(sh != null){
				return sh;
			}
			if(xlsx){
				_OPCXSSFW.close();
				bStreenClose = true;
			}else{
				_NFSHSSFW.close();
				bStreenClose = true;
			}
		} catch (Exception e) {
			Common.addExceptionLogger(e);
		}finally{
			if(xlsx){
				try {
					if(!bStreenClose){
						_OPCXSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				} 
			}else{
				try {
					if(!bStreenClose){
						_NFSHSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		return sh;
	}
		
	/**<p>It will return the {@code ArrayList<String>} of sheet names of the given file.
	 * @author <a href="mailto:ssirekumar@gmail.com">Siri Kumar {@literal <ssirekumar@gmail.com>}</a> </br>
	 * @return {@code ArrayList<String>}  - it will return the ArrayList of sheet names.
	 * @param  filePath  - File path of the .xls or .xlsx file.   
	 * @ 
	 *  */
	public ArrayList<String> getSheetNames(String filePath) {
		File sFilePathObj = null;
		OPCPackage _OPCXSSFW = null;
		POIFSFileSystem  _NFSHSSFW = null;
		boolean bStreenClose = false;
		ArrayList<String> _sheetNames = null;
		try {
			sFilePathObj = new File(filePath);
			xlsx = PoireExcelWorkbook.excelFileType(sFilePathObj).equalsIgnoreCase("xlsx");
			if(xlsx){
				_OPCXSSFW = OPCPackage.open(sFilePathObj);
			}else{
				_NFSHSSFW = new POIFSFileSystem(sFilePathObj);
			}
			Workbook wb = xlsx ? XSSFWorkbookFactory.create(_OPCXSSFW) : WorkbookFactory.create(_NFSHSSFW);
			int iLastSheetNumber = wb.getNumberOfSheets();
			_sheetNames = new ArrayList<String>();
			for (int j=0;j<iLastSheetNumber; j++) {
				String sheetName = wb.getSheetName(j);
				_sheetNames.add(sheetName);
			}
		} catch (Exception e) {
			Common.addExceptionLogger(e);
		}finally{
			if(xlsx){
				try {
					if(_OPCXSSFW != null){
						_OPCXSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				} 
			}else{
				try {
					if(_NFSHSSFW != null){
						_NFSHSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		return _sheetNames;
	}
	
	/**<p>Get all the sheet names as a {@code java.util.ArrayList<String>} of given excel file.
	 * @author <a href="mailto:ssirekumar@gmail.com">Siri Kumar {@literal <ssirekumar@gmail.com>}</a> </br>
	 * @return return {@code java.util.ArrayList<String>} of sheet names.
	 *  */
	public ArrayList<String> getSheetNames() {
		File sFilePathObj = null;
		OPCPackage _OPCXSSFW = null;
		POIFSFileSystem  _NFSHSSFW = null;
		ArrayList<String> _sheetNames = new ArrayList<>();
		try {
			if(excelFile != null) {
				sFilePathObj = excelFile.getCanonicalFile();
			}else {
				return _sheetNames;
			}
			if(xlsx){
				_OPCXSSFW = OPCPackage.open(sFilePathObj);
			}else{
				_NFSHSSFW = new POIFSFileSystem(sFilePathObj);
			}
			Workbook wb = xlsx ? XSSFWorkbookFactory.create(_OPCXSSFW) : WorkbookFactory.create(_NFSHSSFW);
			int iLastSheetNumber = wb.getNumberOfSheets();
			for (int j = 0; j < iLastSheetNumber; j++) {
				String sheetName = wb.getSheetName(j);
				_sheetNames.add(sheetName);
			}
		} catch (Exception e) {
			Common.addExceptionLogger(e);
		}finally{
			if(xlsx){
				try {
					if(_OPCXSSFW != null){
						_OPCXSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				} 
			}else{
				try {
					if(_NFSHSSFW != null){
						_NFSHSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		return _sheetNames;
	}
	
	/**<p>It will used to get the sheet cell value with help of column header name.
	 * @author <a href="mailto:ssirekumar@gmail.com">Siri Kumar {@literal <ssirekumar@gmail.com>}</a> </br>
	 * @return {@code String}
	 * @param  filePath  - File path of the .xls or .xlsx file.
	 * @param  sheetName  - Name of the sheet to get the cell value.
	 * @param  rowNumber - row number of the cell.
	 * @param  columnHeaderName - column header name of the cell.
	 * */
	public static String getCellData(String filePath, String sheetName, int rowNumber, String columnHeaderName) {
		File sFilePathObj = null;
		OPCPackage _OPCXSSFW = null;
		POIFSFileSystem  _NFSHSSFW = null;
		Row row = null;
		Row cellCountRow = null;
		Cell cell = null;
		boolean bStreenClose = false;
		String _sReturnObj = null;
		int bHeader = 0;
		try {
			sFilePathObj = new File(filePath);
			xlsx = PoireExcelWorkbook.excelFileType(sFilePathObj).equalsIgnoreCase("xlsx");
			if(xlsx){
				_OPCXSSFW = OPCPackage.open(sFilePathObj);
			}else{
				_NFSHSSFW = new POIFSFileSystem(sFilePathObj);
			}
			Workbook wb = xlsx ? XSSFWorkbookFactory.create(_OPCXSSFW) : WorkbookFactory.create(_NFSHSSFW);
			Sheet sh = wb.getSheet(sheetName);
			int iLastRowNumber = sh.getLastRowNum();
			if(iLastRowNumber>=rowNumber){
				row = sh.getRow(0);
				short iLastCellNumber = row.getLastCellNum();
				for (int j=0;j<=iLastCellNumber; j++) {
					cell = row.getCell(j, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
					if(cell.toString().equalsIgnoreCase(columnHeaderName)){
						bHeader = j;
						break;
					}
				}
				cellCountRow = sh.getRow(rowNumber);
				if(cellCountRow != null){
					if(bHeader != 0){
						cell = cellCountRow.getCell(bHeader, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
						_sReturnObj = PoireExcelWorkbook.cellFormatedDataValue(cell);  
					}else{
						System.err.println("WARRNING: column Header Name is not fund it might be empty or incorrect spell");
					}
				}else{
					System.err.println("WARRNING: Row is null to get the cell data from the specified values with row Number: "+rowNumber);
					_sReturnObj = "";
				}
			}else{
				System.err.println("WARRNING: RowNumber value is greater than the data present in the sheet:"+ sheetName);
				_sReturnObj = "";
			}
			if(xlsx){
				_OPCXSSFW.close();
				bStreenClose = true;
			}else{
				_NFSHSSFW.close();
				bStreenClose = true;
			}
		} catch (Exception e) {
			Common.addExceptionLogger(e);
		}finally{
			if(xlsx){
				try {
					if(!bStreenClose){
						_OPCXSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				} 
			}else{
				try {
					if(!bStreenClose){
						_NFSHSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		return _sReturnObj;
	}

	/**<p>It will used to get the Excell cell value.
	 * @author <a href="mailto:ssirekumar@gmail.com">Siri Kumar {@literal <ssirekumar@gmail.com>}</a> </br>
	 * @return String
	 * @param  filePath  - File path of the .xls or .xlsx file.
	 * @param  sheetName  - Name of the sheet to get the cell value.
	 * @param  columnHeaderName - column header name of the cell. 
	 * @param  rowNumber - row number of the cell.
	 * @ 
	 * */
	public static String getCellData(String filePath, int sheetNumberIndex, int rowNumber, int columnNumber) {
		File sFilePathObj = null;
		OPCPackage _OPCXSSFW = null;
		POIFSFileSystem  _NFSHSSFW = null;
		Row cellCountRow = null;
		Cell cell = null;
		boolean bStreenClose = false;
		String _sReturnObj = null;
		try {
			sFilePathObj = new File(filePath);
			xlsx = PoireExcelWorkbook.excelFileType(sFilePathObj).equalsIgnoreCase("xlsx");
			if(xlsx){
				_OPCXSSFW = OPCPackage.open(sFilePathObj);
			}else{
				_NFSHSSFW = new POIFSFileSystem(sFilePathObj);
			}
			Workbook wb = xlsx ? XSSFWorkbookFactory.create(_OPCXSSFW) : WorkbookFactory.create(_NFSHSSFW);
			Sheet sh = wb.getSheetAt(sheetNumberIndex);
			int iLastRowNumber = sh.getLastRowNum();
			if(iLastRowNumber>=rowNumber){
				cellCountRow = sh.getRow(rowNumber);
				if(cellCountRow != null){
					cell = cellCountRow.getCell(columnNumber, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
					_sReturnObj = PoireExcelWorkbook.cellFormatedDataValue(cell);  
				}else{
					System.err.println("WARRNING: Row is null to get the cell data from the specified values with row Number: "+rowNumber);
					_sReturnObj = "";
				}
			}else{
				System.err.println("RowNumber value is greter than the Data present in the sheet with rowNumber:"+ rowNumber);
				_sReturnObj = "";
			}
			if(xlsx){
				_OPCXSSFW.close();
				bStreenClose = true;
			}else{
				_NFSHSSFW.close();
				bStreenClose = true;
			}
		} catch (Exception e) {
			Common.addExceptionLogger(e);
		}finally{
			if(xlsx){
				try {
					if(!bStreenClose){
						_OPCXSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				} 
			}else{
				try {
					if(!bStreenClose){
						_NFSHSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		return _sReturnObj;
	}
    
	/**<p>It will used to get the Excell cell value.
	 * @author <a href="mailto:ssirekumar@gmail.com">Siri Kumar {@literal <ssirekumar@gmail.com>}</a> </br>
	 * @return String
	 * @param  filePath  - File path of the .xls or .xlsx file.
	 * @param  sheetNumberIndex  - Number of the sheet number index starts from the 0.
	 * @param  columnHeaderName - column header name of the cell. 
	 * @param  rowNumber - row number of the cell.
	 * @ 
	 * */
	public static String getCellData(String filePath, int sheetNumberIndex, int rowNumber, String columnHeaderName) {
		File sFilePathObj = null;
		OPCPackage _OPCXSSFW = null;
		POIFSFileSystem  _NFSHSSFW = null;
		Row row = null;
		Row cellCountRow = null;
		Cell cell = null;
		int bHeader = 0;
		boolean bStreenClose = false;
		String _sReturnObj = null;
		try {
			sFilePathObj = new File(filePath);
			xlsx = PoireExcelWorkbook.excelFileType(sFilePathObj).equalsIgnoreCase("xlsx");
			if(xlsx){
				_OPCXSSFW = OPCPackage.open(sFilePathObj);
			}else{
				_NFSHSSFW = new POIFSFileSystem(sFilePathObj);
			}
			Workbook wb = xlsx ? XSSFWorkbookFactory.create(_OPCXSSFW) : WorkbookFactory.create(_NFSHSSFW);
			Sheet sh = wb.getSheetAt(sheetNumberIndex);
			int iLastRowNumber = sh.getLastRowNum();
			if(iLastRowNumber>=rowNumber){
				row = sh.getRow(0);
				short iLastCellNumber = row.getLastCellNum();
				for (int j=0;j<=iLastCellNumber; j++) {
					cell = row.getCell(j, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
					if(cell.toString().equalsIgnoreCase(columnHeaderName)){
						bHeader = j;
						break;
					}
				}
				cellCountRow = sh.getRow(rowNumber);
				if(cellCountRow != null){
					if(bHeader !=0){
						cell = cellCountRow.getCell(bHeader, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
						_sReturnObj = PoireExcelWorkbook.cellFormatedDataValue(cell);
					}else{
						System.err.println("WARRNING: column Header Name is not fund it might be empty or incorrect spell");
					}
				}else{
					System.err.println("WARRNING: Row is null to get the cell data from the specified values with row Number: "+rowNumber);
					_sReturnObj = "";
				}
			}else{
				System.err.println("RowNumber value is greter than the Data present in the sheet with rowNumber:"+ rowNumber);
				_sReturnObj = "";
			}
			if(xlsx){
				_OPCXSSFW.close();
				bStreenClose = true;
			}else{
				_NFSHSSFW.close();
				bStreenClose = true;
			}
		} catch (Exception e) {
			Common.addExceptionLogger(e);
		}finally{
			if(xlsx){
				try {
					if(!bStreenClose){
						_OPCXSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				} 
			}else{
				try {
					if(!bStreenClose){
						_NFSHSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		return _sReturnObj;
	}
    	
	/**<p>It will used to get the Excell cell value.
	 * @author <a href="mailto:ssirekumar@gmail.com">Siri Kumar {@literal <ssirekumar@gmail.com>}</a> </br>
	 * @return String
	 * @param  filePath  - File path of the .xls or .xlsx file.
	 * @param  sheetName  - Name of the sheet to get the cell value.
	 * @param  columnNumber - column index Number its starts from the 0 index. 
	 * @param  rowNumber - row number of the cell.
	 * @ 
	 * */
	public static String getCellData(String filePath, String sheetName, int rowNumber, int columnNumber) {
		File sFilePathObj = null;
		OPCPackage _OPCXSSFW = null;
		POIFSFileSystem  _NFSHSSFW = null;
		Row cellCountRow = null;
		Cell cell = null;
		boolean bStreenClose = false;
		String _sReturnObj = null;
		try {
			sFilePathObj = new File(filePath);
			xlsx = PoireExcelWorkbook.excelFileType(sFilePathObj).equalsIgnoreCase("xlsx");
			if(xlsx){
				_OPCXSSFW = OPCPackage.open(sFilePathObj);
			}else{
				_NFSHSSFW = new POIFSFileSystem(sFilePathObj);
			}
			Workbook wb = xlsx ? XSSFWorkbookFactory.create(_OPCXSSFW) : WorkbookFactory.create(_NFSHSSFW);
			Sheet sh = wb.getSheet(sheetName);
			int iLastRowNumber = sh.getLastRowNum();
			if(iLastRowNumber>=rowNumber){
				cellCountRow = sh.getRow(rowNumber);
				if(cellCountRow != null){
					cell = cellCountRow.getCell(columnNumber, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
					_sReturnObj = PoireExcelWorkbook.cellFormatedDataValue(cell);
				}else{
					System.err.println("WARRNING: Row is null to get the cell data from the specified values with row Number: "+rowNumber);
					_sReturnObj = "";
				}
			}else{
				System.err.println("RowNumber value is greter than the Data present in the sheet with rowNumber:"+ rowNumber);
				_sReturnObj = "";
			}
			if(xlsx){
				_OPCXSSFW.close();
				bStreenClose = true;
			}else{
				_NFSHSSFW.close();
				bStreenClose = true;
			}
		} catch (Exception e) {
			Common.addExceptionLogger(e);
		}finally{
			if(xlsx){
				try {
					if(!bStreenClose){
						_OPCXSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				} 
			}else{
				try {
					if(!bStreenClose){
						_NFSHSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		return _sReturnObj;
	}
		
	/**<p>It will used to remove the cell data from the given row with respected to the sheet name
	 * @author Siri Kumar Puttagunta
	 * @return boolean
	 * @param  filePath  - File path of the .xls or .xlsx file.
	 * @param  sheetName  - Name of the sheet to get the cell value.
	 * @param  columnNumber - column index Number its starts from the 0 index. 
	 * @param  rowNumber - row number of the cell.
	 * */
	public static boolean removeCellData(String filePath, String sheetName, int rowNumber, int columnNumber){
		File sFilePathObj = null;
		OPCPackage _OPCXSSFW = null;
		POIFSFileSystem  _NFSHSSFW = null;
		Row cellCountRow = null;
		FileOutputStream fileOut = null;
		Cell cell = null;
		boolean bStreenClose = false;
		boolean _sReturnObj = false;
		try {
			sFilePathObj = new File(filePath);
			xlsx = PoireExcelWorkbook.excelFileType(sFilePathObj).equalsIgnoreCase("xlsx");
			if(xlsx){
				_OPCXSSFW = OPCPackage.open(new FileInputStream(sFilePathObj));
			}else{
				_NFSHSSFW = new POIFSFileSystem(new FileInputStream(sFilePathObj));
			}
			Workbook wb = xlsx ? XSSFWorkbookFactory.create(_OPCXSSFW) : WorkbookFactory.create(_NFSHSSFW);
			Sheet sh = wb.getSheet(sheetName);
			int iLastRowNumber = sh.getLastRowNum();
			if(iLastRowNumber>=rowNumber){
				cellCountRow = sh.getRow(rowNumber);
				if(cellCountRow != null){
					cell = cellCountRow.getCell(columnNumber, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
					String cellValue = PoireExcelWorkbook.cellFormatedDataValue(cell);
					if(cellValue != "" || cell != null){
						cell.setCellValue("");
					}
					fileOut = new FileOutputStream(filePath);
					wb.write(fileOut);
					fileOut.flush();
					fileOut.close();
					bStreenClose = true;
					_sReturnObj = true;
				}else{
					System.err.println("WARRNING: Row is completlly null it does't have any data to remove in the given row number: "+rowNumber);
					_sReturnObj = false;
				}
			}else{
				System.err.println("RowNumber value is greter than the Data present in the sheet with rowNumber:"+ rowNumber);
				_sReturnObj = false;
			}
			if(xlsx){
				_OPCXSSFW.close();
				bStreenClose = true;
			}else{
				_NFSHSSFW.close();
				bStreenClose = true;
			}
		} catch (Exception e) {
			Common.addExceptionLogger(e);
		}finally{
			if(xlsx){
				try {
					if(!bStreenClose){
						_OPCXSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				} 
			}else{
				try {
					if(!bStreenClose){
						_NFSHSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		return _sReturnObj;
	}
	
	/**<p>It will used to remove the cell data from the given row with respected to the sheet number index.
	 * @author Siri Kumar Puttagunta
	 * @return boolean
	 * @param  filePath  - File path of the .xls or .xlsx file.
	 * @param  sheetNumberIndex  - Number of the sheet index this index which is start from 0.
	 * @param  columnNumber - column index Number its starts from the 0 index. 
	 * @param  rowNumber - row number of the cell.
	 * */
	public static boolean removeCellData(String filePath, int sheetNumberIndex, int rowNumber, int columnNumber){
		File sFilePathObj = null;
		OPCPackage _OPCXSSFW = null;
		POIFSFileSystem  _NFSHSSFW = null;
		Row cellCountRow = null;
		FileOutputStream fileOut = null;
		Cell cell = null;
		boolean bStreenClose = false;
		boolean _sReturnObj = false;
		try {
			sFilePathObj = new File(filePath);
			xlsx = PoireExcelWorkbook.excelFileType(sFilePathObj).equalsIgnoreCase("xlsx");
			if(xlsx){
				_OPCXSSFW = OPCPackage.open(new FileInputStream(sFilePathObj));
			}else{
				_NFSHSSFW = new POIFSFileSystem(new FileInputStream(sFilePathObj));
			}
			Workbook wb = xlsx ? XSSFWorkbookFactory.create(_OPCXSSFW) : WorkbookFactory.create(_NFSHSSFW);
			Sheet sh = wb.getSheetAt(sheetNumberIndex);
			int iLastRowNumber = sh.getLastRowNum();
			if(iLastRowNumber>=rowNumber){
				cellCountRow = sh.getRow(rowNumber);
				if(cellCountRow != null){
					cell = cellCountRow.getCell(columnNumber, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
					String cellValue = PoireExcelWorkbook.cellFormatedDataValue(cell);
					if(cellValue != "" || cell != null){
						cell.setCellValue("");
					}
					fileOut = new FileOutputStream(filePath);
					wb.write(fileOut);
					fileOut.flush();
					fileOut.close();
					bStreenClose = true;
					_sReturnObj = true;
				}else{
					System.err.println("WARRNING: Row is completlly null it does't have any data to remove in the given row number: "+rowNumber);
					_sReturnObj = false;
				}
			}else{
				System.err.println("RowNumber value is greter than the Data present in the sheet with rowNumber:"+ rowNumber);
				_sReturnObj = false;
			}
			if(xlsx){
				_OPCXSSFW.close();
				bStreenClose = true;
			}else{
				_NFSHSSFW.close();
				bStreenClose = true;
			}
		} catch (Exception e) {
			Common.addExceptionLogger(e);
		}finally{
			if(xlsx){
				try {
					if(!bStreenClose){
						_OPCXSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				} 
			}else{
				try {
					if(!bStreenClose){
						_NFSHSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		return _sReturnObj;
	}
		
	/**<p>It will used to create the new sheet in the existing excel file.
	 * @author <a href="mailto:ssirekumar@gmail.com">Siri Kumar {@literal <ssirekumar@gmail.com>}</a> </br>
	 * @return boolean
	 * @param  filePath  - File path of the .xls or .xlsx file.
	 * @param  sheetName  - sheet to be created with the this name.
	 * */
	public static boolean addSheet(String filePath, String  sheetName){
		FileOutputStream fileOut;
		File sFilePathObj = null;
		OPCPackage _OPCXSSFW = null;
		POIFSFileSystem  _NFSHSSFW = null;
		boolean bStreenClose = false;
		boolean _sReturnObj = false;
		try {
			sFilePathObj = new File(filePath);
			xlsx = PoireExcelWorkbook.excelFileType(sFilePathObj).equalsIgnoreCase("xlsx");
			if(xlsx){
				_OPCXSSFW = OPCPackage.open(new FileInputStream(sFilePathObj));
			}else{
				_NFSHSSFW = new POIFSFileSystem(new FileInputStream(sFilePathObj));
			}
			Workbook wb = xlsx ? XSSFWorkbookFactory.create(_OPCXSSFW) : WorkbookFactory.create(_NFSHSSFW);
			try {
				wb.createSheet(sheetName);
			} catch (IllegalArgumentException e) {
				System.err.println("WARNNING: For creating a sheet with given name is already present");
			}
			fileOut = new FileOutputStream(sFilePathObj.getAbsolutePath());
			wb.write(fileOut);
			fileOut.flush();
			fileOut.close();
			bStreenClose = true;
			_sReturnObj = true;
		}catch(FileNotFoundException fe){
			Common.addExceptionLogger(fe);
		} catch (Exception e) {
			Common.addExceptionLogger(e);
		}finally{
			if(xlsx){
				try {
					if(!bStreenClose){
						_OPCXSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				} 
			}else{
				try {
					if(!bStreenClose){
						_NFSHSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		return _sReturnObj;
	}

	/**<p>It will used to create the new sheets based on array of names in the existing excel file.
	 * @author <a href="mailto:ssirekumar@gmail.com">Siri Kumar {@literal <ssirekumar@gmail.com>}</a> </br>
	 * @return boolean
	 * @param  filePath  - File path of the .xls or .xlsx file.
	 * @param  sheetName  - sheet to be created with the this name.
	 * */
	public static boolean addSheet(String filePath, String[]  sheetNames){	
		FileOutputStream fileOut;
		File sFilePathObj = null;
		OPCPackage _OPCXSSFW = null;
		POIFSFileSystem  _NFSHSSFW = null;
		boolean bStreenClose = false;
		boolean _sReturnObj = false;
		try {
			sFilePathObj = new File(filePath);
			xlsx = PoireExcelWorkbook.excelFileType(sFilePathObj).equalsIgnoreCase("xlsx");
			if(xlsx){
				_OPCXSSFW = OPCPackage.open(new FileInputStream(sFilePathObj));
			}else{
				_NFSHSSFW = new POIFSFileSystem(new FileInputStream(sFilePathObj));
			}
			Workbook wb = xlsx ? XSSFWorkbookFactory.create(_OPCXSSFW) : WorkbookFactory.create(_NFSHSSFW);
			for (String sheetName : sheetNames) {
				try {
					wb.createSheet(sheetName);
				} catch (IllegalArgumentException e) {
					continue;
				}
			}
			fileOut = new FileOutputStream(sFilePathObj.getAbsolutePath());
			wb.write(fileOut);
			fileOut.flush();
			fileOut.close();
			bStreenClose = true;
			_sReturnObj = true;
		}catch(FileNotFoundException fe){
			Common.addExceptionLogger(fe);
		} catch (Exception e) {
			Common.addExceptionLogger(e);
		}finally{
			if(xlsx){
				try {
					if(!bStreenClose){
						_OPCXSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				} 
			}else{
				try {
					if(!bStreenClose){
						_NFSHSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		return _sReturnObj;
	}

	/**<p>It will used to remove the sheet based on the sheet name in the existing excel file.
	 * @author <a href="mailto:ssirekumar@gmail.com">Siri Kumar {@literal <ssirekumar@gmail.com>}</a> </br>
	 * @return boolean
	 * @param  filePath  - File path of the .xls or .xlsx file.
	 * @param  sheetName  - remove the sheet with sheet name.
	 * */
	public static boolean removeSheet(String filePath, String  sheetname){
		FileOutputStream fileOut;
		File sFilePathObj = null;
		OPCPackage _OPCXSSFW = null;
		POIFSFileSystem  _NFSHSSFW = null;
		boolean bStreenClose = false;
		boolean _sReturnObj = false;
		int bHeader = 0;
		try {
			sFilePathObj = new File(filePath);
			xlsx = PoireExcelWorkbook.excelFileType(sFilePathObj).equalsIgnoreCase("xlsx");
			if(xlsx){
				_OPCXSSFW = OPCPackage.open(new FileInputStream(sFilePathObj));
			}else{
				_NFSHSSFW = new POIFSFileSystem(new FileInputStream(sFilePathObj));
			}
			Workbook wb = xlsx ? XSSFWorkbookFactory.create(_OPCXSSFW) : WorkbookFactory.create(_NFSHSSFW);
			int iLastSheetNumber = wb.getNumberOfSheets();
			for (int j=0;j<=iLastSheetNumber; j++) {
				Sheet sh = wb.getSheetAt(j);
				if(sh.getSheetName().toString().equalsIgnoreCase(sheetname)){
					bHeader = j;
					break;
				}
			}
			if (iLastSheetNumber>=bHeader) {
				wb.removeSheetAt(bHeader);
				fileOut = new FileOutputStream(sFilePathObj.getAbsolutePath());
				wb.write(fileOut);
				fileOut.flush();
				fileOut.close();
				bStreenClose = true;
				_sReturnObj = true;
			} else {
					System.err.println("WARRNING: For removing the sheet name is not presented in the workbook");
			}
		}catch(FileNotFoundException fe){
			Common.addExceptionLogger(fe);
		} catch (Exception e) {
			Common.addExceptionLogger(e);
		}finally{
			if(xlsx){
				try {
					if(!bStreenClose){
						_OPCXSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				} 
			}else{
				try {
					if(!bStreenClose){
						_NFSHSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		return _sReturnObj;
	}
	
	/**<p>It will used to remove the sheet based on the sheet Number in the existing excel file.
	 * @author <a href="mailto:ssirekumar@gmail.com">Siri Kumar {@literal <ssirekumar@gmail.com>}</a> </br>
	 * @return boolean
	 * @param  filePath  - File path of the .xls or .xlsx file.
	 * @param  sheetNumber  - remove the sheet with sheet index number.
	 * */
	public static boolean removeSheet(String filePath, int  sheetNumber){
		FileOutputStream fileOut;
		File sFilePathObj = null;
		OPCPackage _OPCXSSFW = null;
		POIFSFileSystem  _NFSHSSFW = null;
		boolean bStreenClose = false;
		boolean _sReturnObj = false;
		try {
			sFilePathObj = new File(filePath);
			xlsx = PoireExcelWorkbook.excelFileType(sFilePathObj).equalsIgnoreCase("xlsx");
			if(xlsx){
				_OPCXSSFW = OPCPackage.open(new FileInputStream(sFilePathObj));
			}else{
				_NFSHSSFW = new POIFSFileSystem(new FileInputStream(sFilePathObj));
			}
			Workbook wb = xlsx ? XSSFWorkbookFactory.create(_OPCXSSFW) : WorkbookFactory.create(_NFSHSSFW);
			int iLastSheetNumber = wb.getNumberOfSheets();
			if(iLastSheetNumber>=sheetNumber){
				wb.removeSheetAt(sheetNumber);
				fileOut = new FileOutputStream(sFilePathObj.getAbsolutePath());
				wb.write(fileOut);
				fileOut.flush();
				fileOut.close();
				bStreenClose = true;
				_sReturnObj = true;
			}else{
				System.err.println("WARRNING: For removing the sheet name is not prented in the workbook");
			}
		}catch(FileNotFoundException fe){
			Common.addExceptionLogger(fe);
		} catch (Exception e) {
			Common.addExceptionLogger(e);
		}finally{
			if(xlsx){
				try {
					if(!bStreenClose){
						_OPCXSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				} 
			}else{
				try {
					if(!bStreenClose){
						_NFSHSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		return _sReturnObj;
	}
	
	/**<p>It will used to remove the sheets based on the array of sheet Number in the existing excel file.
	 * @author <a href="mailto:ssirekumar@gmail.com">Siri Kumar {@literal <ssirekumar@gmail.com>}</a> </br>
	 * @return boolean
	 * @param  filePath  - File path of the .xls or .xlsx file.
	 * @param  sheetNumber  - remove the sheet with array of sheet numbers.
	 * */
	public static boolean removeSheet(String filePath, int[] sheetNumber){
		FileOutputStream fileOut;
		File sFilePathObj = null;
		OPCPackage _OPCXSSFW = null;
		POIFSFileSystem  _NFSHSSFW = null;
		boolean bStreenClose = false;
		boolean _sReturnObj = false;
		try {
			sFilePathObj = new File(filePath);
			xlsx = PoireExcelWorkbook.excelFileType(sFilePathObj).equalsIgnoreCase("xlsx");
			if(xlsx){
				_OPCXSSFW = OPCPackage.open(new FileInputStream(sFilePathObj));
			}else{
				_NFSHSSFW = new POIFSFileSystem(new FileInputStream(sFilePathObj));
			}
			Workbook wb = xlsx ? XSSFWorkbookFactory.create(_OPCXSSFW) : WorkbookFactory.create(_NFSHSSFW);
			for (int sheetName : sheetNumber) {
				Sheet sh = wb.getSheetAt(sheetName);
				if(sh != null){
					wb.removeSheetAt(sheetName);
				}else{
					System.err.println("The Specified Index number of the sheet is not Prasent in excel file: " +sheetName+" ");
				}
			}
			fileOut = new FileOutputStream(sFilePathObj.getAbsolutePath());
			wb.write(fileOut);
			fileOut.flush();
			fileOut.close();
			bStreenClose = true;
			_sReturnObj = true;
		}catch(FileNotFoundException fe){
			Common.addExceptionLogger(fe);
		} catch (Exception e) {
			Common.addExceptionLogger(e);
		}finally{
			if(xlsx){
				try {
					if(!bStreenClose){
						_OPCXSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				} 
			}else{
				try {
					if(!bStreenClose){
						_NFSHSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		return _sReturnObj;
	}
	
	/**<p>It will used to remove the sheet based on the array of names.
	 * @author <a href="mailto:ssirekumar@gmail.com">Siri Kumar {@literal <ssirekumar@gmail.com>}</a> </br>
	 * @return boolean
	 * @param  filePath  - File path of the .xls or .xlsx file.
	 * @param  sheetName  - remove the sheets with array of sheet Names.
	 * */
	public static boolean removeSheet(String filePath, String[]  sheetNames){
		FileOutputStream fileOut;
		File sFilePathObj = null;
		OPCPackage _OPCXSSFW = null;
		POIFSFileSystem  _NFSHSSFW = null;
		boolean bStreenClose = false;
		boolean _sReturnObj = false;
		try {
			sFilePathObj = new File(filePath);
			xlsx = PoireExcelWorkbook.excelFileType(sFilePathObj).equalsIgnoreCase("xlsx");
			if(xlsx){
				_OPCXSSFW = OPCPackage.open(new FileInputStream(sFilePathObj));
			}else{
				_NFSHSSFW = new POIFSFileSystem(new FileInputStream(sFilePathObj));
			}
			Workbook wb = xlsx ? XSSFWorkbookFactory.create(_OPCXSSFW) : WorkbookFactory.create(_NFSHSSFW);
			int iLastSheetNumber = wb.getNumberOfSheets();
			for (int j=0;j<iLastSheetNumber; j++) {
				boolean removeSheet = false;
				Sheet sh = wb.getSheetAt(j);
				for (String sheetName : sheetNames) {
					if(sh.getSheetName().toString().equalsIgnoreCase(sheetName)){
						wb.removeSheetAt(j);
						removeSheet = true;
						break;
					}
				}
				if(removeSheet){
					iLastSheetNumber = wb.getNumberOfSheets();	
					j = 0;
				}
			}
			fileOut = new FileOutputStream(sFilePathObj.getAbsolutePath());
			wb.write(fileOut);
			fileOut.flush();
			fileOut.close();
			bStreenClose = true;
			_sReturnObj = true;
		}catch(FileNotFoundException fe){
			Common.addExceptionLogger(fe);
		} catch (Exception e) {
			Common.addExceptionLogger(e);
		}finally{
			if(xlsx){
				try {
					if(!bStreenClose){
						_OPCXSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				} 
			}else{
				try {
					if(!bStreenClose){
						_NFSHSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		return _sReturnObj;
	}
	
	/**<p>It will used to give a boolean value with the sheet name is present or not on the existing excel file.
	 * @author <a href="mailto:ssirekumar@gmail.com">Siri Kumar {@literal <ssirekumar@gmail.com>}</a> </br>
	 * @return boolean
	 * @param  filePath  - File path of the .xls or .xlsx file.
	 * @param  sheetName  - sheet to be created with the this name.
	 * */
	public static boolean isSheetExist(String filePath, String sheetName){
		FileOutputStream fileOut;
		File sFilePathObj = null;
		OPCPackage _OPCXSSFW = null;
		POIFSFileSystem  _NFSHSSFW = null;
		boolean bStreenClose = false;
		boolean _sReturnObj = false;
		try {
			sFilePathObj = new File(filePath);
			xlsx = PoireExcelWorkbook.excelFileType(sFilePathObj).equalsIgnoreCase("xlsx");
			if(xlsx){
				_OPCXSSFW = OPCPackage.open(new FileInputStream(sFilePathObj));
			}else{
				_NFSHSSFW = new POIFSFileSystem(new FileInputStream(sFilePathObj));
			}
			Workbook wb = xlsx ? XSSFWorkbookFactory.create(_OPCXSSFW) : WorkbookFactory.create(_NFSHSSFW);
			int iLastSheetNumber = wb.getNumberOfSheets();
			for (int j=0;j<iLastSheetNumber; j++) {
				Sheet sh = wb.getSheetAt(j);
				System.out.println(sh.getSheetName());
				System.out.println(sheetName);
				if(sh.getSheetName().toString().equalsIgnoreCase(sheetName)){
					_sReturnObj = true;
					break;
				}
			}
			fileOut = new FileOutputStream(sFilePathObj.getAbsolutePath());
			wb.write(fileOut);
			fileOut.flush();
			fileOut.close();
			bStreenClose = true;
		}catch(FileNotFoundException fe){
			Common.addExceptionLogger(fe);
		} catch (Exception e) {
			Common.addExceptionLogger(e);
		}finally{
			if(xlsx){
				try {
					if(!bStreenClose){
						_OPCXSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				} 
			}else{
				try {
					if(!bStreenClose){
						_NFSHSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		return _sReturnObj;
		
	} 
	
	/*public static int getColumnFilterCount(String filePath, String sheetName, int columnNumber){ // incompleate the method functionality.
		ArrayList<String> colContents = null;
		File sFilePathObj = null;
		OPCPackage _OPCXSSFW = null;
		POIFSFileSystem  _NFSHSSFW = null;
		Row row = null;
		Cell cell = null;
		boolean bStreenClose = false;
		try {
			sFilePathObj = new File(filePath);
			xlsx = PoireExcelWorkbook.excelFileType(sFilePathObj).equalsIgnoreCase("xlsx");
			if(xlsx){
				_OPCXSSFW = OPCPackage.open(sFilePathObj);
			}else{
				_NFSHSSFW = new POIFSFileSystem(sFilePathObj);
			}
			Workbook wb = xlsx ? XSSFWorkbookFactory.create(_OPCXSSFW) : WorkbookFactory.create(_NFSHSSFW);
			Sheet sh = wb.getSheetAt(sheetNumberIndex);
			row = sh.
		}catch(FileNotFoundException fe){
			Common.addExceptionLogger(fe);
		} catch (Exception e) {
			Common.addExceptionLogger(e);
		}finally{
			if(xlsx){
				try {
					if(!bStreenClose){
						_OPCXSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				} 
			}else{
				try {
					if(!bStreenClose){
						_NFSHSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
	}*/
		
	/**<p>It will create the row at end of each data row based on the sheet name.
	 *    And it will return the created row value. Row count is starts from 0 index.
	 * @author <a href="mailto:ssirekumar@gmail.com">Siri Kumar {@literal <ssirekumar@gmail.com>}</a> </br>
	 * @return It will return the positive integer of newly created value. 
	 * @param  filePath  - File path of the .xls or .xlsx file.   
	 * @param  sheetName - Sheet Name of the file.
	 *  */
	public static int createNewRowAtEnd(String filePath, String sheetName){   
		File sFilePathObj = null;
		OPCPackage _OPCXSSFW = null;
		POIFSFileSystem  _NFSHSSFW = null;
		Row row = null;
		Cell cell = null;
		FileOutputStream fileOut = null;
		boolean bStreenClose = false;
		int _sReturnObj = -1;
		int lastrow = -1;
		try {
			sFilePathObj = new File(filePath);
			xlsx = PoireExcelWorkbook.excelFileType(sFilePathObj).equalsIgnoreCase("xlsx");
			if(xlsx){
				_OPCXSSFW = OPCPackage.open(new FileInputStream(sFilePathObj));
			}else{
				_NFSHSSFW = new POIFSFileSystem(new FileInputStream(sFilePathObj));
			}
			Workbook wb = xlsx ? XSSFWorkbookFactory.create(_OPCXSSFW) : WorkbookFactory.create(_NFSHSSFW);
			Sheet sh = wb.getSheet(sheetName);
			int iLastRowNumber = sh.getLastRowNum();
			lastrow = iLastRowNumber + 1;
			row = sh.createRow(lastrow);
			int lastCell = PoireExcelWorkbook.getColumnCount(filePath, sheetName);
			for (int k=0;k<=lastCell;k++) {  
				cell = row.createCell(k, CellType.STRING);
				cell.setCellValue("");
			}
			fileOut = new FileOutputStream(filePath);
			wb.write(fileOut);
			fileOut.flush();
			fileOut.close();
			bStreenClose = true;
			_sReturnObj = lastrow;
			if(xlsx){
				_OPCXSSFW.close();
				bStreenClose = true;
			}else{
				_NFSHSSFW.close();
				bStreenClose = true;
			}
		}catch(FileNotFoundException fe){
			Common.addExceptionLogger(fe);
		} catch (Exception e) {
			Common.addExceptionLogger(e);
		}finally{
			if(xlsx){
				try {
					if(!bStreenClose){
						_OPCXSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				} 
			}else{
				try {
					if(!bStreenClose){
						_NFSHSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		return _sReturnObj;
	} 
    	
	/**<p>It will create the row at end of each data row based on the sheet number index.
	 *    And it will return the created row value. Row count is starts from 0 index.
	 * @author <a href="mailto:ssirekumar@gmail.com">Siri Kumar {@literal <ssirekumar@gmail.com>}</a> </br>
	 * @return It will return the positive integer of newly created value. 
	 * @param  filePath  - File path of the .xls or .xlsx file.   
	 * @param  sheetNumberIndex - Index of the sheet Number it will start from 0.
	 *  */
	public static int createNewRowAtEnd(String filePath, int sheetNumberIndex){   
		File sFilePathObj = null;
		OPCPackage _OPCXSSFW = null;
		POIFSFileSystem  _NFSHSSFW = null;
		Row row = null;
		Cell cell = null;
		FileOutputStream fileOut = null;
		boolean bStreenClose = false;
		int _sReturnObj = -1;
		int lastrow = -1;
		try {
			sFilePathObj = new File(filePath);
			xlsx = PoireExcelWorkbook.excelFileType(sFilePathObj).equalsIgnoreCase("xlsx");
			if(xlsx){
				_OPCXSSFW = OPCPackage.open(new FileInputStream(sFilePathObj));
			}else{
				_NFSHSSFW = new POIFSFileSystem(new FileInputStream(sFilePathObj));
			}
			Workbook wb = xlsx ? XSSFWorkbookFactory.create(_OPCXSSFW) : WorkbookFactory.create(_NFSHSSFW);
			Sheet sh = wb.getSheetAt(sheetNumberIndex);
			int iLastRowNumber = sh.getLastRowNum();
			lastrow = iLastRowNumber + 1;
			row = sh.createRow(lastrow);
			int lastCell = PoireExcelWorkbook.getColumnCount(filePath, sheetNumberIndex);
			for (int k=0;k<=lastCell;k++) {  
				cell = row.createCell(k, CellType.STRING);
				cell.setCellValue("");
			}
			fileOut = new FileOutputStream(filePath);
			wb.write(fileOut);
			fileOut.flush();
			fileOut.close();
			bStreenClose = true;
			_sReturnObj = lastrow;
			if(xlsx){
				_OPCXSSFW.close();
				bStreenClose = true;
			}else{
				_NFSHSSFW.close();
				bStreenClose = true;
			}
		}catch(FileNotFoundException fe){
			Common.addExceptionLogger(fe);
		} catch (Exception e) {
			Common.addExceptionLogger(e);
		}finally{
			if(xlsx){
				try {
					if(!bStreenClose){
						_OPCXSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				} 
			}else{
				try {
					if(!bStreenClose){
						_NFSHSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		return _sReturnObj;
	} 
	
	/**<p>It will give you a list of all the sheet names arrayList in the existing excel file.
	 * @author <a href="mailto:ssirekumar@gmail.com">Siri Kumar {@literal <ssirekumar@gmail.com>}</a> </br>
	 * @return It will return the ArrayList of all the sheet names.
	 * @param  filePath  - File path of the .xls or .xlsx file.   
	 * */
	public static ArrayList<String> getAllSheetNames(String filePath) throws FileNotFoundException, Exception{
		File sFilePathObj = null;
		OPCPackage _OPCXSSFW = null;
		POIFSFileSystem  _NFSHSSFW = null;
		boolean bStreenClose = false;
		ArrayList<String> _addListElements = null;
		try {
			sFilePathObj = new File(filePath);
			xlsx = PoireExcelWorkbook.excelFileType(sFilePathObj).equalsIgnoreCase("xlsx");
			if(xlsx){
				_OPCXSSFW = OPCPackage.open(new FileInputStream(sFilePathObj));
			}else{
				_NFSHSSFW = new POIFSFileSystem(new FileInputStream(sFilePathObj));
			}
			Workbook wb = xlsx ? XSSFWorkbookFactory.create(_OPCXSSFW) : WorkbookFactory.create(_NFSHSSFW);
			_addListElements = new ArrayList<String>();
			int iLastSheetNumber = wb.getNumberOfSheets();
			for (int j=0;j<=iLastSheetNumber; j++) {
				Sheet sh = wb.getSheetAt(j);
				_addListElements.add(sh.getSheetName());
			}
			if(xlsx){
				_OPCXSSFW.close();
				bStreenClose = true;
			}else{
				_NFSHSSFW.close();
				bStreenClose = true;
			}
		}catch(FileNotFoundException fe){
			Common.addExceptionLogger(fe);
			throw new FileNotFoundException(fe.getMessage());
		} catch (Exception e) {
			Common.addExceptionLogger(e);
			throw new Exception(e.getMessage());
		}finally{
			if(xlsx){
				try {
					if(!bStreenClose){
						_OPCXSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				} 
			}else{
				try {
					if(!bStreenClose){
						_NFSHSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		return _addListElements;
	}
		
	/**<p>It will insert the row at between the two row positions
	 * @author <a href="mailto:ssirekumar@gmail.com">Siri Kumar {@literal <ssirekumar@gmail.com>}</a> </br>
	 * @return It will return the ArrayList of all the sheet names.
	 * @param  filePath  - File path of the .xls or .xlsx file.  
	 * @param  sheetName - Name of the sheet in the given existing excel file.
	 * @param  aboveRowNumber - Insertion row above rownumber
	 * @param  belowRowNumber - Insertion row below row number
	 * @param  numberOfRows   - Number of rows to insert in bitween the above and below rows.
	 * 
	 *  */
	public static boolean insertRowAtPosition(String filePath, String sheetName,int aboveRowNumber, int belowRowNumber, int numberOfRows){  //incompleate method
		File sFilePathObj = null;
		OPCPackage _OPCXSSFW = null;
		POIFSFileSystem  _NFSHSSFW = null;
		boolean bStreenClose = false;
		FileOutputStream fileOut = null;
		boolean _sReturnObj = false;
		Row belowRow = null;
		//Row aboveRow = null;
		Cell cell = null;
		try {
			sFilePathObj = new File(filePath);
			xlsx = PoireExcelWorkbook.excelFileType(sFilePathObj).equalsIgnoreCase("xlsx");
			if(xlsx){
				_OPCXSSFW = OPCPackage.open(new FileInputStream(sFilePathObj));
			}else{
				_NFSHSSFW = new POIFSFileSystem(new FileInputStream(sFilePathObj));
			}
			Workbook wb = xlsx ? XSSFWorkbookFactory.create(_OPCXSSFW) : WorkbookFactory.create(_NFSHSSFW);
			Sheet sh = wb.getSheet(sheetName);
			belowRow =  sh.getRow(belowRowNumber);
			//aboveRow = sh.getRow(aboveRowNumber);
			int lastCell = PoireExcelWorkbook.getColumnCount(filePath, sheetName);
			if(belowRow != null){
				sh.shiftRows(aboveRowNumber, sh.getLastRowNum(), numberOfRows, false, false);
				Row newRow = sh.getRow(aboveRowNumber+numberOfRows);
				for (int i = 1; i < numberOfRows; i++) {
					for (int j = 0; j <= lastCell; j++) {
						cell = newRow.createCell(j, CellType.STRING);
						cell.setCellValue("");
					}
				}
				_sReturnObj = true;
			}else{
				belowRow = sh.createRow(belowRowNumber);
				for (int k=0;k<=lastCell;k++) {  
					cell = belowRow.createCell(k, CellType.STRING);
					cell.setCellValue("");
				}
			}
			fileOut = new FileOutputStream(filePath);
			wb.write(fileOut);
			fileOut.flush();
			fileOut.close();
			bStreenClose = true;
			if(xlsx){
				_OPCXSSFW.close();
				bStreenClose = true;
			}else{
				_NFSHSSFW.close();
				bStreenClose = true;
			}
		}catch(FileNotFoundException fe){
			Common.addExceptionLogger(fe);
		} catch (Exception e) {
			Common.addExceptionLogger(e);
		}finally{
			if(xlsx){
				try {
					if(!bStreenClose){
						_OPCXSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				} 
			}else{
				try {
					if(!bStreenClose){
						_NFSHSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		return _sReturnObj;
	}
		
	/**<p>It will get the Cell class object with the help of the rowNumber and columnNumber
	 * @author <a href="mailto:ssirekumar@gmail.com">Siri Kumar {@literal <ssirekumar@gmail.com>}</a> </br>
	 * @return It will return the ArrayList of all the sheet names.
	 * @param  filePath  - File path of the .xls or .xlsx file.  
	 * @param  sheetName - Name of the sheet in the given existing excel file.
	 * @param  rowNumber - int Number of the row
	 * @param  columnNumber - int Number of the column
	 *  */
	public static Cell getCell(String filePath, String sheetName, int rowNumber, int columnNumber){
		File sFilePathObj = null;
		OPCPackage _OPCXSSFW = null;
		POIFSFileSystem  _NFSHSSFW = null;
		boolean bStreenClose = false;
		Row cellCountRow = null;
		Cell cell = null;
		try {
			sFilePathObj = new File(filePath);
			xlsx = PoireExcelWorkbook.excelFileType(sFilePathObj).equalsIgnoreCase("xlsx");
			if(xlsx){
				_OPCXSSFW = OPCPackage.open(new FileInputStream(sFilePathObj));
			}else{
				_NFSHSSFW = new POIFSFileSystem(new FileInputStream(sFilePathObj));
			}
			Workbook wb = xlsx ? XSSFWorkbookFactory.create(_OPCXSSFW) : WorkbookFactory.create(_NFSHSSFW);
			Sheet sh = wb.getSheet(sheetName);
			int iLastRowNumber = sh.getLastRowNum();
			if(iLastRowNumber>=rowNumber){
				cellCountRow = sh.getRow(rowNumber);
				if(cellCountRow != null){
					cell = cellCountRow.getCell(columnNumber, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
					return cell;
				}else{
					System.err.println("WARRNING: Row is null to get the cell data from the specified values with row Number: "+rowNumber);
				}
			}else{
				System.err.println("RowNumber value is greter than the Data present in the sheet with rowNumber:"+ rowNumber);
			}
			if(xlsx){
				_OPCXSSFW.close();
				bStreenClose = true;
			}else{
				_NFSHSSFW.close();
				bStreenClose = true;
			}
		}catch(FileNotFoundException fe){
			Common.addExceptionLogger(fe);
		} catch (Exception e) {
			Common.addExceptionLogger(e);
		}finally{
			if(xlsx){
				try {
					if(!bStreenClose){
						_OPCXSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				} 
			}else{
				try {
					if(!bStreenClose){
						_NFSHSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		return cell;
	}
		
	/**<p>It will get the Row class object with the help of the rowNumber
	 * @author <a href="mailto:ssirekumar@gmail.com">Siri Kumar {@literal <ssirekumar@gmail.com>}</a> </br>
	 * @return It will return the ArrayList of all the sheet names.
	 * @param  filePath  - File path of the .xls or .xlsx file.  
	 * @param  sheetName - Name of the sheet in the given existing excel file.
	 * @param  rowNumber - int Number of the row
	 *  */
	public static Row getRowWithRowNumber(String filePath, String sheetName, int rowNumber){
		File sFilePathObj = null;
		OPCPackage _OPCXSSFW = null;
		POIFSFileSystem  _NFSHSSFW = null;
		boolean bStreenClose = false;
		Row cellCountRow = null;
		try {
			sFilePathObj = new File(filePath);
			xlsx = PoireExcelWorkbook.excelFileType(sFilePathObj).equalsIgnoreCase("xlsx");
			if(xlsx){
				_OPCXSSFW = OPCPackage.open(new FileInputStream(sFilePathObj));
			}else{
				_NFSHSSFW = new POIFSFileSystem(new FileInputStream(sFilePathObj));
			}
			Workbook wb = xlsx ? XSSFWorkbookFactory.create(_OPCXSSFW) : WorkbookFactory.create(_NFSHSSFW);
			Sheet sh = wb.getSheet(sheetName);
			int iLastRowNumber = sh.getLastRowNum();
			if(iLastRowNumber>=rowNumber){
				cellCountRow = sh.getRow(rowNumber);
				if(cellCountRow != null){
					return cellCountRow;
				}else{
					cellCountRow = null;
				}
			}else{
				System.err.println("RowNumber value is greter than the Data present in the sheet with rowNumber:"+ rowNumber);
			}
			if(xlsx){
				_OPCXSSFW.close();
				bStreenClose = true;
			}else{
				_NFSHSSFW.close();
				bStreenClose = true;
			}
		}catch(FileNotFoundException fe){
			Common.addExceptionLogger(fe);
		} catch (Exception e) {
			Common.addExceptionLogger(e);
		}finally{
			if(xlsx){
				try {
					if(!bStreenClose){
						_OPCXSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				} 
			}else{
				try {
					if(!bStreenClose){
						_NFSHSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		return cellCountRow;
	}
	
	/**<p>Get the Row class object with the help of the headerName
	 * @author <a href="mailto:ssirekumar@gmail.com">Siri Kumar {@literal <ssirekumar@gmail.com>}</a> </br>
	 * @return  {@literal org.apache.poi.ss.usermodel.Row} class object
	 * @param  sheetName - Name of the sheet in the given existing excel file.
	 * @param  rowHeaderName - Name of the row header at first column
	 *  */
	public Row getRowWithHeaderName(String sheetName, String rowHeaderName){
		File sFilePathObj = null;
		OPCPackage _OPCXSSFW = null;
		POIFSFileSystem  _NFSHSSFW = null;
		boolean bStreenClose = false;
		Row cellCountRow = null;
		int bHeader = 0, loopValue = 0;
		try {
			ArrayList<String> sFirstIndexColValues = this.getExcelColumnWithSheetName(0, sheetName, true);
			if(sFirstIndexColValues.size() > 0 || sFirstIndexColValues != null){
				for (String zeroIndexcellValue : sFirstIndexColValues) {
					if(zeroIndexcellValue.equalsIgnoreCase(rowHeaderName)){
						bHeader = loopValue;
						break;
					}
					loopValue++;
				}
			}
			if(excelFile != null) {
				sFilePathObj = excelFile.getCanonicalFile();
			}else {
				return cellCountRow;
			}
			if(xlsx){
				_OPCXSSFW = OPCPackage.open(sFilePathObj);
			}else{
				_NFSHSSFW = new POIFSFileSystem(sFilePathObj);
			}
			Workbook wb = xlsx ? XSSFWorkbookFactory.create(_OPCXSSFW) : WorkbookFactory.create(_NFSHSSFW);
			Sheet sh = wb.getSheet(sheetName);
			int iLastRowNumber = sh.getLastRowNum();
			if(iLastRowNumber>=bHeader){
				cellCountRow = sh.getRow(bHeader);
				if(cellCountRow != null){
					return cellCountRow;
				}else{
					cellCountRow = null;
				}
			}else{
				System.err.println("RowNumber value is greter than the Data present in the sheet with rowNumber:"+ bHeader);
			}
		} catch (InvalidFormatException | IOException e) {
			Common.addExceptionLogger(e);
		}finally{
			if(xlsx){
				try {
					if(_OPCXSSFW != null){
						_OPCXSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				} 
			}else{
				try {
					if(_NFSHSSFW != null){
						_NFSHSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		return cellCountRow;
	}
	
	/**
	 * Function to return boolean value of if the cell is filled with color ot not.
	 * @author 	<a href="mailto:ssirekumar@gmail.com">Siri Kumar {@literal <ssirekumar@gmail.com>}</a> </br>
	 * @param 	cellObject - Cell class object
	 **/
	public static boolean isCellColorFilled(Cell cellObject){
		boolean _returnObj = false;
		try {
			CellStyle style = cellObject.getCellStyle();
			Color colorCell = style.getFillBackgroundColorColor();
			if (colorCell != null) {
				_returnObj = true;
			} else {
				_returnObj = false;
			}
		}catch (Exception e) {
			Common.addExceptionLogger(e);
		}
		return _returnObj;
	}
		
	/**<p>It will used to create the excel headre row with specified data and its background colors
	 * @author <a href="mailto:ssirekumar@gmail.com">Siri Kumar {@literal <ssirekumar@gmail.com>}</a> </br>
	 * @return boolean
	 * @param  filePath  - File path of the .xls or .xlsx file.
	 * @param  headerRowList - list of the data for header row cells. 
	 * @param  sheetName  - sheet to be created with the this name.
	 * @param  cellBackgroundColor - color of the cells of header {IndexedColors.DARK_RED} 
	 * @param  cellFontColor - color of the cell value {IndexedColors.WHITE}
	 * */
 	public static boolean createHeaderRow(String filePath, List<String> headerRowList, String sheetName, IndexedColors cellBackgroundColor, IndexedColors cellFontColor){
		File sFilePathObj = null;
		OPCPackage _OPCXSSFW = null;
		POIFSFileSystem  _NFSHSSFW = null;
		Row row = null;
		FileOutputStream fileOut = null;
		Cell cell = null;
		boolean bStreenClose = false;
		boolean _sReturnObj = false;
		try {
			sFilePathObj = new File(filePath);
			xlsx = PoireExcelWorkbook.excelFileType(sFilePathObj).equalsIgnoreCase("xlsx");
			if(xlsx){
				_OPCXSSFW = OPCPackage.open(new FileInputStream(sFilePathObj));
			}else{
				_NFSHSSFW = new POIFSFileSystem(new FileInputStream(sFilePathObj));
			}
			Workbook wb = xlsx ? XSSFWorkbookFactory.create(_OPCXSSFW) : WorkbookFactory.create(_NFSHSSFW);
			Sheet sh = wb.getSheet(sheetName);
			row = sh.createRow(0);
			CellStyle style = wb.createCellStyle();
			Font font = wb.createFont();
			for(int i=0;i<headerRowList.size();i++){
				cell = row.createCell(i, CellType.BLANK);
				style.setFillBackgroundColor(cellBackgroundColor.getIndex());
				style.setFillPattern(CellStyle.ALIGN_FILL);
	            font.setColor(cellFontColor.getIndex());
	            style.setFont(font);
				cell.setCellValue(headerRowList.get(i).toString());
				row.getCell(i).setCellStyle(style);
			}
			fileOut = new FileOutputStream(filePath);
			wb.write(fileOut);
			fileOut.flush();
			fileOut.close();
			bStreenClose = true;
			_sReturnObj = true;
			if(xlsx){
				_OPCXSSFW.close();
				bStreenClose = true;
			}else{
				_NFSHSSFW.close();
				bStreenClose = true;
			}
		}catch(FileNotFoundException fe){
			Common.addExceptionLogger(fe);
		} catch (Exception e) {
			Common.addExceptionLogger(e);
		}finally{
			if(xlsx){
				try {
					if(!bStreenClose){
						_OPCXSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				} 
			}else{
				try {
					if(!bStreenClose){
						_NFSHSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		return _sReturnObj;
		
	}
	
 	/**<p>It will used to get the Mearged region of the cell as a list
	 * @author Siri Kumar Puttagunta
	 * @return List
	 * @param  filePath  - File path of the .xls or .xlsx file.
	 * @param  sheetName  - Sheet to be created with the this name.
	 * */ 	
 	public static List<CellRangeAddress> getMeargedRegions(String filePath, String sheetName){
 		File sFilePathObj = null;
 		OPCPackage _OPCXSSFW = null;
		POIFSFileSystem  _NFSHSSFW = null;
		List<CellRangeAddress> regionsList = null;
		boolean bStreenClose = false;
		try {
			sFilePathObj = new File(filePath);
			xlsx = PoireExcelWorkbook.excelFileType(sFilePathObj).equalsIgnoreCase("xlsx");
			if(xlsx){
				_OPCXSSFW = OPCPackage.open(new FileInputStream(sFilePathObj));
			}else{
				_NFSHSSFW = new POIFSFileSystem(new FileInputStream(sFilePathObj));
			}
			Workbook wb = xlsx ? XSSFWorkbookFactory.create(_OPCXSSFW) : WorkbookFactory.create(_NFSHSSFW);
			Sheet sh = wb.getSheet(sheetName);
			regionsList = new ArrayList<CellRangeAddress>();
			for(int i = 0; i < sh.getNumMergedRegions(); i++) {
			   regionsList.add(sh.getMergedRegion(i));
			   System.out.println("The total cell mearged range is :"+sh.getMergedRegion(i).getNumberOfCells() );
			   System.out.println("The mearged range is :"+ regionsList.get(i).toString());
			}
			if(xlsx){
				_OPCXSSFW.close();
				bStreenClose = true;
			}else{
				_NFSHSSFW.close();
				bStreenClose = true;
			}
		}catch(FileNotFoundException fe){
			Common.addExceptionLogger(fe);
		} catch (Exception e) {
			Common.addExceptionLogger(e);
		}finally{
			if(xlsx){
				try {
					if(!bStreenClose){
						_OPCXSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				} 
			}else{
				try {
					if(!bStreenClose){
						_NFSHSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		return regionsList;
 	}
 	 	
 	/**<p>It will used to check the is cell is meargeg or not with in the range
	 * @author <a href="mailto:ssirekumar@gmail.com">Siri Kumar {@literal <ssirekumar@gmail.com>}</a> </br>
	 * @return boolean
	 * @param  filePath - File path of the .xls or .xlsx file.
	 * @param  sheetName - Sheet Name in the file.
	 * @param  regionRangeList - list of CellRangeAddress object this will come from the
	 * 		   getMeargedRegions method.	 
	 * @param  rowNumber - int number of the row
	 * @param  colNumber   int number of the column
	 * */
 	public static boolean isCellMearged(String filePath, String sheetName, List<CellRangeAddress> regionRangeList, int rowNumber, int colNumber){
 		File sFilePathObj = null;
		OPCPackage _OPCXSSFW = null;
		POIFSFileSystem  _NFSHSSFW = null;
		Cell cell = null;
		boolean bStreenClose = false;
		boolean _sReturnObj = false;
		try {
			sFilePathObj = new File(filePath);
			xlsx = PoireExcelWorkbook.excelFileType(sFilePathObj).equalsIgnoreCase("xlsx");
			if(xlsx){
				_OPCXSSFW = OPCPackage.open(new FileInputStream(sFilePathObj));
			}else{
				_NFSHSSFW = new POIFSFileSystem(new FileInputStream(sFilePathObj));
			}
			Workbook wb = xlsx ? XSSFWorkbookFactory.create(_OPCXSSFW) : WorkbookFactory.create(_NFSHSSFW);
			Sheet sh = wb.getSheet(sheetName);
			for(CellRangeAddress region : regionRangeList) {
				if(region.isInRange(rowNumber, colNumber)) {
					cell = sh.getRow(rowNumber).getCell(colNumber);
					if(cell != null){
						_sReturnObj = true;
					}
				}
			}
			if(xlsx){
				_OPCXSSFW.close();
				bStreenClose = true;
			}else{
				_NFSHSSFW.close();
				bStreenClose = true;
			}
		}catch(FileNotFoundException fe){
			Common.addExceptionLogger(fe);
		} catch (Exception e) {
			Common.addExceptionLogger(e);
		}finally{
			if(xlsx){
				try {
					if(!bStreenClose){
						_OPCXSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				} 
			}else{
				try {
					if(!bStreenClose){
						_NFSHSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		return _sReturnObj;
 		
 	} 
 	
 	/**<p>It will used to get the is cell header name with respected to the row and column number.
	 * @author <a href="mailto:ssirekumar@gmail.com">Siri Kumar {@literal <ssirekumar@gmail.com>}</a> </br>
	 * @return boolean
	 * @param  filePath - File path of the .xls or .xlsx file. 
	 * @param  sheetName - Sheet Name in the file.
	 * @param  rowNumber - int number of the row
	 * @param  colNumber   int number of the column
	 * */
 	public static String getCellHeaderName(String filePath, String sheetName, int rowNumber, int columnNumber){
 		File sFilePathObj = null;
		OPCPackage _OPCXSSFW = null;
		POIFSFileSystem  _NFSHSSFW = null;
		Row cellCountRow = null;
		Cell cell = null;
		boolean bStreenClose = false;
		String _sReturnObj = null;
		try {
			sFilePathObj = new File(filePath);
			xlsx = PoireExcelWorkbook.excelFileType(sFilePathObj).equalsIgnoreCase("xlsx");
			if(xlsx){
				_OPCXSSFW = OPCPackage.open(sFilePathObj);
			}else{
				_NFSHSSFW = new POIFSFileSystem(sFilePathObj);
			}
			Workbook wb = xlsx ? XSSFWorkbookFactory.create(_OPCXSSFW) : WorkbookFactory.create(_NFSHSSFW);
			Sheet sh = wb.getSheet(sheetName);
			int iLastRowNumber = sh.getLastRowNum();
			if(iLastRowNumber>=rowNumber){
				cellCountRow = sh.getRow(0);
				if(cellCountRow != null){
					cell = cellCountRow.getCell(columnNumber, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
					_sReturnObj = PoireExcelWorkbook.cellFormatedDataValue(cell);
				}else{
					System.err.println("WARRNING: Row is null to get the Header cell data");
					_sReturnObj = "";
				}
			}else{
				System.err.println("RowNumber value is greter than the Data present in the sheet with rowNumber:"+ rowNumber);
				_sReturnObj = "";
			}
			if(xlsx){
				_OPCXSSFW.close();
				bStreenClose = true;
			}else{
				_NFSHSSFW.close();
				bStreenClose = true;
			}
		} catch (Exception e) {
			Common.addExceptionLogger(e);
		}finally{
			if(xlsx){
				try {
					if(!bStreenClose){
						_OPCXSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				} 
			}else{
				try {
					if(!bStreenClose){
						_NFSHSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		return _sReturnObj;
 	}
 	
 	/**<p>It will used to get the row number of the given cell data.
	 * @author <a href="mailto:ssirekumar@gmail.com">Siri Kumar {@literal <ssirekumar@gmail.com>}</a> </br>
	 * @return boolean
	 * @param  filePath - File path of the .xls or .xlsx file. 
	 * @param  sheetName - Sheet Name in the file.
	 * @param  rowNumber - int number of the row
	 * @param  toFindString  - string of cell data for get the cell number.
	 * *//*
 	public static int getCellRowNumber(String sheetName, String colName, String cellValue){
 		
 	}*/
 	 	
 	/**<p>It will used to get the cell number with comparison of toFindString data.
	 * @author <a href="mailto:ssirekumar@gmail.com">Siri Kumar {@literal <ssirekumar@gmail.com>}</a> </br>
	 * @return boolean
	 * @param  filePath - File path of the .xls or .xlsx file. 
	 * @param  sheetName - Sheet Name in the file.
	 * @param  rowNumber - int number of the row
	 * @param  toFindString  - string of cell data for get the cell number.
	 * */
 	public static int getCellNumberOfgivenData(String filePath, String sheetName, int rowNumber, String toFindString){
 		File sFilePathObj = null;
		OPCPackage _OPCXSSFW = null;
		POIFSFileSystem  _NFSHSSFW = null;
		Row cellCountRow = null;
		Cell cell = null;
		boolean bStreenClose = false;
		int _sReturnObj = -1;
		String cellData = "";
		try {
			sFilePathObj = new File(filePath);
			xlsx = PoireExcelWorkbook.excelFileType(sFilePathObj).equalsIgnoreCase("xlsx");
			if(xlsx){
				_OPCXSSFW = OPCPackage.open(sFilePathObj);
			}else{
				_NFSHSSFW = new POIFSFileSystem(sFilePathObj);
			}
			Workbook wb = xlsx ? XSSFWorkbookFactory.create(_OPCXSSFW) : WorkbookFactory.create(_NFSHSSFW);
			Sheet sh = wb.getSheet(sheetName);
			int iLastRowNumber = sh.getLastRowNum();
			int colCount = PoireExcelWorkbook.getColumnCount(filePath, sheetName);
			if(iLastRowNumber>=rowNumber){
				cellCountRow = sh.getRow(rowNumber);
				if(cellCountRow != null){
					for (int i = 0; i <= colCount; i++) {
						cell = cellCountRow.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
						cellData = PoireExcelWorkbook.cellFormatedDataValue(cell);
						if(toFindString.equalsIgnoreCase(cellData)){
							_sReturnObj = i;
							break;
						}
					}
				}else{
					System.err.println("WARRNING: Row is null to get the Header cell data");
					_sReturnObj = -1;
				}
			}else{
				System.err.println("RowNumber value is greter than the Data present in the sheet with rowNumber:"+ rowNumber);
				_sReturnObj = -1;
			}
			if(xlsx){
				_OPCXSSFW.close();
				bStreenClose = true;
			}else{
				_NFSHSSFW.close();
				bStreenClose = true;
			}
		} catch (Exception e) {
			Common.addExceptionLogger(e);
		}finally{
			if(xlsx){
				try {
					if(!bStreenClose){
						_OPCXSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				} 
			}else{
				try {
					if(!bStreenClose){
						_NFSHSSFW.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		return _sReturnObj;
 	}
 	
 	/**<p>It will used to Convert the Excel sheet data into csv file with Sheet Name.
	 * @author <a href="mailto:ssirekumar@gmail.com">Siri Kumar {@literal <ssirekumar@gmail.com>}</a> </br>
	 * @return boolean
	 * @param  excelFilePath - File path of the .xls or .xlsx file. 
	 * @param  sheetName - Sheet Name in the file.
	 * @param  outPutFilePath - Csv file location path to store.
	 * @param  outPutFileName  - Csv file name.
 	 * @ 
 	 *  
	 * */
 	public static boolean convertExcelToCsv(String excelFilePath, String sheetName, String outPutPath, String outPutFileName) 
 			throws FileNotFoundException, NullPointerException, ArrayIndexOutOfBoundsException, IOException, Exception{
 		ArrayList<ArrayList<String>> contents = PoireExcelWorkbook.getExcelDataWithSheetName(excelFilePath, -1, sheetName);
 		List<String[]> writeContents = new ArrayList<String[]>();
 		CSVWriter csvWriter = null;
 		boolean _returnObj = false;
 		boolean bStreenClose = false;
 		try {
 			if(contents != null){
 				for (ArrayList<String> row : contents) {
 	 				writeContents.add(row.toArray(new String[row.size()]));
 	 			}
 	 			csvWriter = new CSVWriter(new FileWriter(outPutPath + "\\" + outPutFileName + ".csv",false));
 	 			csvWriter.writeAll(writeContents);
 	 			csvWriter.close();
 	 			_returnObj = true;
 	 			bStreenClose = true;
 			}else{
 				System.err.println("Data of sheet is null");
 	 			System.err.println("Either Sheet data empty or sheet is not present");
 			}
 		} catch (FileNotFoundException fne) {
 			Common.addExceptionLogger(fne);
 			throw new FileNotFoundException(fne.getMessage());
 		} catch (IOException ioe) {
 			Common.addExceptionLogger(ioe);
 			throw new IOException(ioe.getMessage());
 		} catch (NullPointerException npe) {
 			Common.addExceptionLogger(npe);
 			throw new NullPointerException(npe.getMessage());
 		} catch (ArrayIndexOutOfBoundsException ae) {
 			Common.addExceptionLogger(ae);
 			throw new ArrayIndexOutOfBoundsException(ae.getMessage());
 		}catch (Exception e) {
 			Common.addExceptionLogger(e);
 			throw new Exception(e.getMessage());
 		}finally{
 			try {
				if(!bStreenClose){
					csvWriter.close();
				}
			} catch (IOException e) {
				e.printStackTrace();
				Common.addExceptionLogger(e);
			} 
 		}
 		return _returnObj;
 	}
 	
 	/**<p>It will used to Convert the Excel sheet data into csv file with sheet Number.
	 * @author <a href="mailto:ssirekumar@gmail.com">Siri Kumar {@literal <ssirekumar@gmail.com>}</a> </br>
	 * @return boolean
	 * @param  excelFilePath - File path of the .xls or .xlsx file. 
	 * @param  sheetNumberIndex - Sheet Name in the file.
	 * @param  outPutFilePath - Csv file location path to store.
	 * @param  outPutFileName  - Csv file name.
 	 * @ 
 	 *  
	 * */
 	public static boolean convertExcelToCsv(String excelFilePath, int sheetNumberIndex, String outPutPath, String outPutFileName) 
 			throws FileNotFoundException, NullPointerException, ArrayIndexOutOfBoundsException, IOException, Exception{
 		ArrayList<ArrayList<String>> contents = PoireExcelWorkbook.getExcelDataWithSheetNumber(excelFilePath, -1, sheetNumberIndex);
 		List<String[]> writeContents = new ArrayList<String[]>();
 		CSVWriter csvWriter = null;
 		boolean _returnObj = false;
 		boolean bStreenClose = false;
 		try {
 			if(contents != null){
 				for (ArrayList<String> row : contents) {
 	 				writeContents.add(row.toArray(new String[row.size()]));
 	 			}
 	 			csvWriter = new CSVWriter(new FileWriter(outPutPath + "\\" + outPutFileName + ".csv",false));
 	 			csvWriter.writeAll(writeContents);
 	 			csvWriter.close();
 	 			_returnObj = true;
 	 			bStreenClose = true;
 			}else{
 				System.err.println("Data of sheet is null");
 	 			System.err.println("Either Sheet data empty or sheet is not present");
 			}
 		} catch (FileNotFoundException fne) {
 			Common.addExceptionLogger(fne);
 			throw new FileNotFoundException(fne.getMessage());
 		} catch (IOException ioe) {
 			Common.addExceptionLogger(ioe);
 			throw new IOException(ioe.getMessage());
 		} catch (NullPointerException npe) {
 			Common.addExceptionLogger(npe);
 			throw new NullPointerException(npe.getMessage());
 		} catch (ArrayIndexOutOfBoundsException ae) {
 			Common.addExceptionLogger(ae);
 			throw new ArrayIndexOutOfBoundsException(ae.getMessage());
 		}catch(Exception e){
 			Common.addExceptionLogger(e);
 			throw new Exception(e.getMessage());
 		}finally{
 			try {
				if(!bStreenClose){
					csvWriter.close();
				}
			} catch (IOException e) {
				e.printStackTrace();
				Common.addExceptionLogger(e);
			} 
 		}
 		return _returnObj;
 		
 	}
 	 	
 	/**<p>It will used to Convert All the Excel sheet data into csv files with sheetname as a file name of csv.
	 * @author <a href="mailto:ssirekumar@gmail.com">Siri Kumar {@literal <ssirekumar@gmail.com>}</a> </br>
	 * @return boolean
	 * @param  excelFilePath - File path of the .xls or .xlsx file. 
	 * @param  outPutFilePath - Csv file location path to store.
	 * */
 	public static boolean convertExcelSheetsToCsvFiles(String excelFilePath, String outPutFilePath) 
 			throws FileNotFoundException, NullPointerException, ArrayIndexOutOfBoundsException, IOException, Exception{  // need to check at exception handlings
 		LoggerPoire.setLoggerObject(PoireExcelWorkbook._log);
 		List<String[]> writeContents = new ArrayList<String[]>();
 		CSVWriter csvWriter = null;
 		boolean _returnObj = false;
 		boolean bStreenClose = false;
 		boolean caughtException = false;
 		try {
 			ArrayList<String> _sheetNames = null;
 			try {
 				_sheetNames = PoireExcelWorkbook.getAllSheetNames(excelFilePath);
 			} catch (Exception e) {
 				e.printStackTrace();
 				caughtException = true;
 			}
 			if (_sheetNames != null) {
 				for (String sheetName : _sheetNames) {
 	 				ArrayList<ArrayList<String>> contents = null;
 	 				contents = PoireExcelWorkbook.getExcelDataWithSheetName(excelFilePath, -1, sheetName);
 	 				if(contents != null){
 	 					for (ArrayList<String> row : contents) {
 	 						writeContents.add(row.toArray(new String[row.size()]));
 	 					}
 	 					csvWriter = new CSVWriter(new FileWriter(outPutFilePath + "\\" + sheetName + ".csv",false));
 	 					csvWriter.writeAll(writeContents);
 	 					csvWriter.close();
 	 					_returnObj = true;
 	 					bStreenClose = true;
 	 					writeContents.clear();
 	 					contents.clear();
 	 				}else{
 	 					System.err.println("Data of "+sheetName+" is null");
 	 					System.err.println("Either Sheet data empty or sheet is not present");
 	 					LoggerPoire.log(Log.ERROR, "Either Sheet data empty or sheet is not present");
 	 				}
 	 			}
			} else {
				if(caughtException){
					System.err.println("Sheet Names are not found");
				}
				System.err.println("Sheet Names are not found");
				System.err.println("Either Sheet data empty or sheet is not present");	
			}
 		} catch (FileNotFoundException fne) {
 			Common.addExceptionLogger(fne);
 			throw new FileNotFoundException(fne.getMessage());
 		} catch (IOException ioe) {
 			Common.addExceptionLogger(ioe);
 			throw new IOException(ioe.getMessage());
 		} catch (NullPointerException npe) {
 			Common.addExceptionLogger(npe);
 			throw new NullPointerException(npe.getMessage());
 		} catch (ArrayIndexOutOfBoundsException ae) {
 			Common.addExceptionLogger(ae);
 			throw new ArrayIndexOutOfBoundsException(ae.getMessage());
 		}catch(Exception e){
 			Common.addExceptionLogger(e);
 			throw new Exception(e.getMessage());
 		}finally{
 			try {
				if(!bStreenClose){
					csvWriter.close();
				}
			} catch (IOException e) {
				e.printStackTrace();
				Common.addExceptionLogger(e);
			} 
 		}
 		return _returnObj;
 	}
 	
 	/**<p>It will used to Convert All the CSV data file into Excel file with specified sheet name.
	 * @author <a href="mailto:ssirekumar@gmail.com">Siri Kumar {@literal <ssirekumar@gmail.com>}</a> </br>
	 * @return boolean
	 * @param  csvFilePath - Complete File path of the CSV file  
	 * @param  sheetName - Sheet name to create the excel file.
	 * @param  outPutExcelPath    - Directory path of the Excel file need to create.
	 * @param  outPutExcelFileName - File name of the Excel which place on the outPutExcelPath Directory.
	 * @param  xlsxfileFormat      - true for .xlsx file to create false for .xls file to create.
 	 * @ 
	 * */
 	public static boolean convertCsvToExcel(String csvFilePath, String sheetName, String outPutExcelPath, String outPutExcelFileName, boolean xlsxfileFormat) {
 		ArrayList<ArrayList<String>> contents = PoireExcelWorkbook.getCSVData(csvFilePath, 0);
 		boolean _sReturnObject = false;
 		if(contents != null){
 			File _newExceFile = null;
			try {
				_newExceFile = PoireExcelWorkbook.createExcelFile(xlsxfileFormat, outPutExcelPath, outPutExcelFileName, sheetName);
			} catch (Exception e) {
				e.printStackTrace();
			}
 			if(_newExceFile == null){
 				System.err.println("The New Excel file is not created");
 				return _sReturnObject;
 			}else{
 				try {
					PoireExcelWorkbook.createExcelFileWithData(_newExceFile.getAbsolutePath(), contents, sheetName);
					_sReturnObject = true;
				} catch (Exception e) {
					e.printStackTrace();
					_sReturnObject = false;
				}
 			}
 		}else{
 			System.err.println("Data of csv file is null");
 		}
 		return _sReturnObject;
 	}
 	 
 	/**<p>It will used to Convert array of All the CSV files to sheets in single excel file.
	 * @author <a href="mailto:ssirekumar@gmail.com">Siri Kumar {@literal <ssirekumar@gmail.com>}</a> </br>
	 * @return boolean
	 * @param  csvFilePath - array of Complete File path of the CSV files.  
	 * @param  outPutExcelPath    - Directory path of the Excel file need to create.
	 * @param  outPutExcelFileName - File name of the Excel which place on the outPutExcelPath Directory.
	 * @param  xlsxfileFormat      - true for .xlsx file to create false for .xls file to create.
	 * */
 	public static boolean convertCsvsToExcelSheets(String[] csvFilePath, String outPutExcelPath, String outPutExcelFileName, boolean xlsxfileFormat){
 		LoggerPoire.setLoggerObject(PoireExcelWorkbook._log);
 		boolean _sReturnObject = false;
 		boolean[] loopValue = new boolean[csvFilePath.length]; 
 		int loopIndex = 0;
 		String sSheetName = "";
 		File _newExceFile = null;
 		for (String _sCsvFilePath : csvFilePath) {
 			File _sFile = new File(_sCsvFilePath);
 			sSheetName = _sFile.getName().split(".csv")[0];
 			if(loopIndex == 0){
 				try {
					_newExceFile =  PoireExcelWorkbook.createExcelFile(xlsxfileFormat, outPutExcelPath, outPutExcelFileName, sSheetName);
				} catch (Exception e) {
					e.printStackTrace();
				}
 			}
 			if(_newExceFile == null || !((_newExceFile.exists()) && (_newExceFile.isFile()))){
 				System.err.println("The New Excel file is not created");
	 			LoggerPoire.log(Log.ERROR, "The New Excel file is not created");
	 			return false;
 			}else{
 				ArrayList<ArrayList<String>> contents = PoireExcelWorkbook.getCSVData(_sCsvFilePath, 0);
 	 			if(contents != null){
 	 				if(loopIndex != 0){
 	 					PoireExcelWorkbook.addSheet(_newExceFile.getAbsolutePath(), sSheetName);
 	 				}
 	 				try {
						_sReturnObject = PoireExcelWorkbook.createExcelFileWithData(_newExceFile.getAbsolutePath(), contents, sSheetName);
					} catch (Exception e) {
						e.printStackTrace();
					}	
 	 	 		}else{
 	 	 			System.err.println("Data of csv file is null");
 	 	 		}
 			}
 			loopValue[loopIndex] = _sReturnObject;
 			loopIndex++;
		}
 		if(Common.contains(loopValue, false)){
 			_sReturnObject = false;
 			return _sReturnObject;
 		}else{
 			_sReturnObject = true;
 			return _sReturnObject;
 		}
 	}
 	 	
 	/**<p>It will used to Convert Given name of the Sheet to new excel file. 
	 * @author <a href="mailto:ssirekumar@gmail.com">Siri Kumar {@literal <ssirekumar@gmail.com>}</a> </br>
	 * @return File - it will return converted sheet as file object.
	 * @param  excelFilePath - File path of the excel file. 
	 * @param  sheetName - Sheet name to create the excel file.
	 * @param  outPutPath - Directory Path for new excel file creation.
	 * @param  outPutFileName - New excel file name. 
	 * @param  xlsxfileFormat - boolean value for to create the excel file true - .xlsx file creation.
 	 * @ 
	 * */
 	public static File convertSheetToExcelFile(String excelFilePath, String sheetName, String outPutPath, String outPutFileName, boolean xlsxfileFormat) {  
 		File sFilePathObj = null;
		OPCPackage _OPCXSSFW = null;
		OPCPackage _OPCXSSFW2 = null;
		POIFSFileSystem  _NFSHSSFW = null;
		POIFSFileSystem  _NFSHSSFW2 = null;
		FileOutputStream fileOut = null;
		File _newExceFile = null;
		boolean bStreenClose = false;
		File _sReturnObj = null;
		File snewSheetFilePathObj = null;
		boolean xlsxT = false;
		try {
			try {
				sFilePathObj = new File(excelFilePath);
				xlsx = PoireExcelWorkbook.excelFileType(sFilePathObj).equalsIgnoreCase("xlsx");
				if(xlsx){
					_OPCXSSFW = OPCPackage.open(new FileInputStream(sFilePathObj));
				}else{
					_NFSHSSFW = new POIFSFileSystem(new FileInputStream(sFilePathObj));
				}
				Workbook wb = xlsx ? XSSFWorkbookFactory.create(_OPCXSSFW) : WorkbookFactory.create(_NFSHSSFW);
				Sheet sh = wb.getSheet(sheetName);
				Workbook wbs = sh.getWorkbook();
				_newExceFile =  PoireExcelWorkbook.createExcelFile(xlsxfileFormat, outPutPath, outPutFileName, sheetName);
				fileOut = new FileOutputStream(_newExceFile);
				wbs.write(fileOut);
				fileOut.close();
				if(xlsx){
					_OPCXSSFW.close();
					bStreenClose = true;
				}else{
					_NFSHSSFW.close();
					bStreenClose = true;
				}
			} catch (Exception e) {
				Common.addExceptionLogger(e);
				throw new Exception(e.getMessage());
			}finally{
				if(xlsx){
					try {
						if(!bStreenClose){
							_OPCXSSFW.close();
						}
					} catch (IOException e) {
						e.printStackTrace();
					} 
				}else{
					try {
						if(!bStreenClose){
							_NFSHSSFW.close();
						}
					} catch (IOException e) {
						e.printStackTrace();
					}
				}
			}
			if(_newExceFile.exists() && _newExceFile.isFile()){
				snewSheetFilePathObj = new File(_newExceFile.getAbsolutePath());
				xlsxT = PoireExcelWorkbook.excelFileType(snewSheetFilePathObj).equalsIgnoreCase("xlsx");
				if(xlsxT){
					_OPCXSSFW2 = OPCPackage.open(new FileInputStream(snewSheetFilePathObj));
				}else{
					_NFSHSSFW2 = new POIFSFileSystem(new FileInputStream(snewSheetFilePathObj));
				}
				Workbook wb2 = xlsxT ? WorkbookFactory.create(_OPCXSSFW2) : WorkbookFactory.create(_NFSHSSFW2);
				for(int i = wb2.getNumberOfSheets() - 1; i >= 0; i--){
					if(!wb2.getSheetAt(i).getSheetName().equalsIgnoreCase(sheetName)){
						wb2.removeSheetAt(i);
					}
				}
				fileOut = new FileOutputStream(snewSheetFilePathObj);
				wb2.write(fileOut);
				fileOut.close();
			}else{
				System.err.println("The Newly created file is not found");
			}
			if(xlsxT){
				_OPCXSSFW2.close();
				bStreenClose = true;
			}else{
				_NFSHSSFW2.close();
				bStreenClose = true;
			}
			_sReturnObj = snewSheetFilePathObj;
		} catch (Exception e) {
			Common.addExceptionLogger(e);
		}finally{
			if(xlsxT){
				try {
					if(!bStreenClose){
						_OPCXSSFW2.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				} 
			}else{
				try {
					if(!bStreenClose){
						_NFSHSSFW2.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		return _sReturnObj;
 	}
 	
 	/**<p>It will used to Convert Given names of the Sheets to new excel file. 
	 * @author <a href="mailto:ssirekumar@gmail.com">Siri Kumar {@literal <ssirekumar@gmail.com>}</a> </br>
	 * @return File - it will return converted sheet as file object.
	 * @param  excelFilePath - File path of the excel file. 
	 * @param  sheetNames - Selected Array of Sheet names to create the excel file.
	 * @param  outPutPath - Directory Path for new excel file creation.
	 * @param  outPutFileName - New excel file name. 
	 * @param  xlsxfileFormat - boolean value for to create the excel file true - .xlsx file creation.
	 * */
 	public static File convertSheetsToExcelFile(String excelFilePath, String[] sheetNames, String outPutPath, String outPutFileName, boolean xlsxfileFormat) {  
 		File sFilePathObj = null;
		OPCPackage _OPCXSSFW = null;
		OPCPackage _OPCXSSFW2 = null;
		POIFSFileSystem  _NFSHSSFW = null;
		POIFSFileSystem  _NFSHSSFW2 = null;
		FileOutputStream fileOut = null;
		File _newExceFile = null;
		boolean bStreenClose = false;
		File _sReturnObj = null;
		File snewSheetFilePathObj = null;
		boolean xlsxT = false;
		try {
			try {
				sFilePathObj = new File(excelFilePath);
				xlsx = PoireExcelWorkbook.excelFileType(sFilePathObj).equalsIgnoreCase("xlsx");
				if(xlsx){
					_OPCXSSFW = OPCPackage.open(new FileInputStream(sFilePathObj));
				}else{
					_NFSHSSFW = new POIFSFileSystem(new FileInputStream(sFilePathObj));
				}
				Workbook wb = xlsx ? XSSFWorkbookFactory.create(_OPCXSSFW) : WorkbookFactory.create(_NFSHSSFW);
				for (String sheetName : sheetNames) {
					Sheet sh = wb.getSheet(sheetName);
					Workbook wbs = sh.getWorkbook();
					_newExceFile =  PoireExcelWorkbook.createExcelFile(xlsxfileFormat, outPutPath, outPutFileName, sheetName);
					fileOut = new FileOutputStream(_newExceFile);
					wbs.write(fileOut);
				}
				fileOut.close();
				if(xlsx){
					_OPCXSSFW.close();
					bStreenClose = true;
				}else{
					_NFSHSSFW.close();
					bStreenClose = true;
				}
			} catch (Exception e) {
				Common.addExceptionLogger(e);
				throw new Exception(e.getMessage());
			}finally{
				if(xlsx){
					try {
						if(!bStreenClose){
							_OPCXSSFW.close();
						}
					} catch (IOException e) {
						e.printStackTrace();
					} 
				}else{
					try {
						if(!bStreenClose){
							_NFSHSSFW.close();
						}
					} catch (IOException e) {
						e.printStackTrace();
					}
				}
			}
			if(_newExceFile.exists() && _newExceFile.isFile()){
				snewSheetFilePathObj = new File(_newExceFile.getAbsolutePath());
				xlsxT = PoireExcelWorkbook.excelFileType(snewSheetFilePathObj).equalsIgnoreCase("xlsx");
				if(xlsxT){
					_OPCXSSFW2 = OPCPackage.open(new FileInputStream(snewSheetFilePathObj));
				}else{
					_NFSHSSFW2 = new POIFSFileSystem(new FileInputStream(snewSheetFilePathObj));
				}
				Workbook wb2 = xlsxT ? WorkbookFactory.create(_OPCXSSFW2) : WorkbookFactory.create(_NFSHSSFW2);
				Set<Integer> SheetIndex = new HashSet<Integer>();
				for (String sheetName : sheetNames) {
					int i = wb2.getNumberOfSheets() - 1;
					int j = i;
					for(i = j; i >= 0; i--){
						System.out.println(wb2.getSheetAt(i).getSheetName());
						if(wb2.getSheetAt(i).getSheetName().equalsIgnoreCase(sheetName)){
							SheetIndex.add(Integer.valueOf(i));
						}
					}
				}
				for (int k = wb2.getNumberOfSheets() - 1; k >= 0; k--) {
					if(!(SheetIndex.contains(k))){
						wb2.removeSheetAt(k);
					}
				}
				fileOut = new FileOutputStream(snewSheetFilePathObj);
				wb2.write(fileOut);
				fileOut.close();
			}else{
				System.err.println("The Newly created file is not found");
			}
			if(xlsxT){
				_OPCXSSFW2.close();
				bStreenClose = true;
			}else{
				_NFSHSSFW2.close();
				bStreenClose = true;
			}
			_sReturnObj = snewSheetFilePathObj;
		} catch (Exception e) {
			Common.addExceptionLogger(e);
		}finally{
			if(xlsxT){
				try {
					if(!bStreenClose){
						_OPCXSSFW2.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				} 
			}else{
				try {
					if(!bStreenClose){
						_NFSHSSFW2.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		return _sReturnObj;
 	}
 	
	/**
	 * excelFileType(File excelFile)
	 * 
	 */
	private static String excelFileType(File excelFile) {
		String sReturnFileType = "";
		if (excelFile.isFile() && excelFile.exists()) {
			String fileName = excelFile.getName();
			String extension = fileName.substring(fileName.lastIndexOf(".") + 1, fileName.length());
			if (extension.equalsIgnoreCase("xls")) {
				sReturnFileType = "xls";
			} else if (extension.equalsIgnoreCase("xlsx")) {
				sReturnFileType = "xlsx";
			} else {
				System.err.println("The File is not a type of Excel file format");
			}
		}
		return sReturnFileType;
	}

	/**
	 * Function to return Array of random string
	 * @author 	<a href="mailto:ssirekumar@gmail.com">Siri Kumar {@literal <ssirekumar@gmail.com>}</a> </br>
	 * @param 	int - length of the array,
	 * @param   int length of each element
	 * @return 	ArrayList as per parameters passed*/
	private static ArrayList<String> getArrayOfRandomString(int lengthOfArray, int lengthOfElements) {
		ArrayList<String> randomArray = null;
		try {
			randomArray = new ArrayList<String>();
			for (int i = 0; i < lengthOfArray; i++) {
				randomArray.add(new BigInteger(Long.SIZE * lengthOfElements, random)
				.toString(32).substring(0, lengthOfElements)
				.toUpperCase());
			}
		} catch (Exception e) {
			Common.addExceptionLogger(e);
		}
		return randomArray;
	}
    
	/**
	 * Function to return Array of random number upto Long data type range.
	 * @author 	<a href="mailto:ssirekumar@gmail.com">Siri Kumar {@literal <ssirekumar@gmail.com>}</a> </br>
	 * @param 	lengthOfArray  - number of digits for the random number
	 * @param   numberOfDigits - length of each element
	 * @return 	long random number with number of digits passed in parameter*/
	private static ArrayList<String> getArrayOfRandomNumber(int lengthOfArray, int numberOfDigits) {
		long accumulator = 0L;
		ArrayList<String> randomArray = null;
		try {
			randomArray = new ArrayList<String>();
			Random randomGenerator = new Random();
			for (int j=0;j<lengthOfArray;j++) {
				accumulator = 0L;
				accumulator = 1 + randomGenerator.nextInt(9);
				for (int i = 0; i < (numberOfDigits - 1); i++) {
					accumulator *= 10L;
					accumulator += randomGenerator.nextInt(10);
				}
				randomArray.add(String.valueOf(accumulator));
			}
		} catch (Exception e) {
			Common.addExceptionLogger(e);
		}
		return randomArray;
	}
		
	/**<h3>cellFormatedDataValue</h3>
	 * Function to return String of formated data in the Excel Cell value.
	 * @author 	<a href="mailto:ssirekumar@gmail.com">Siri Kumar {@literal <ssirekumar@gmail.com>}</a> </br>
	 * @param 	cellObject  - Cell class object 
	 * @return 	String of the formated datavalue.*/
	public static String cellFormatedDataValue(Cell cellObject){
		String cellValue = "";
		DataFormatter _formater = new DataFormatter();
		try {
			switch (cellObject.getCellType()) {   // this switch case is used for specific cell type data. 
    		case NUMERIC:        // each cell is formated and to get the exact value in the cell as pr the excel data format.
    			cellValue = String.valueOf(_formater.formatCellValue(cellObject));
    			if(PoireExcelWorkbook.cellDateFormat(cellValue)){
    				DateFormat df = null;
    				String dateFormatString = Common.getdateFormatString();
    				if(dateFormatString != ""){
    					 df = new SimpleDateFormat(Common.getdateFormatString());
    				}else{
    					 df = new SimpleDateFormat("MM/dd/yyyy");
    				}
    			    Date today = cellObject.getDateCellValue();       
    			    cellValue = df.format(today);
    			}
    			break;
    		case STRING:
    			cellValue = String.valueOf(_formater.formatCellValue(cellObject));
    			break;
    		case BLANK:
    			cellValue = String.valueOf("");
    			break;
    		case BOOLEAN:
    			cellValue = String.valueOf(_formater.formatCellValue(cellObject));
    			break;
    		case ERROR:
    			cellValue = String.valueOf(_formater.formatCellValue(cellObject));
    			break;
    		case FORMULA:
    			Workbook wb = cellObject.getSheet().getWorkbook();
    			FormulaEvaluator formulaEval = wb.getCreationHelper().createFormulaEvaluator();
    			formulaEval.evaluate(cellObject);
    			cellValue = String.valueOf(_formater.formatCellValue(cellObject,formulaEval));
    			if(PoireExcelWorkbook.cellDateFormat(cellValue)){
    				DateFormat df = null;
    				String dateFormatString = Common.getdateFormatString();
    				if(dateFormatString != ""){
    					df = new SimpleDateFormat(Common.getdateFormatString());
    				}else{
    					df = new SimpleDateFormat("MM/dd/yyyy");
    				}
    				Date today = cellObject.getDateCellValue();       
    				cellValue = df.format(today);
    			}
    			break;
    		default:
    			break;
			}
		} catch (Exception e) {
			Common.addExceptionLogger(e);
		}
		return cellValue;
	}
    
	/**<h3>cellFormatedDataValue</h3>
	 * Function to return boolean value of the given date.
	 * ex: if the date is mm/dd/yy dd/mm/yy and mm/dd/yyy dd/mm/yyyy give you true 
	 * @author 	<a href="mailto:ssirekumar@gmail.com">Siri Kumar {@literal <ssirekumar@gmail.com>}</a> </br>
	 * @param 	dateValue  - Cell class object 
	 * 
	 * **/
	private static boolean cellDateFormat(String dateValue){
		String regexMMDDYYYY="^(0[1-9]|1[012])[- /.](0[1-9]|[12][0-9]|3[01])[- /.](19|20)\\d\\d$";
	    String regexDDMMYYYY="^(0[1-9]|[12][0-9]|3[01])[- /.](0[1-9]|1[012])[- /.](19|20)\\d\\d$";
	    String regexDDMMYY="^[0-3]?[0-9]/[0-3]?[0-9]/(?:[0-9]{2})?[0-9]{2}$";
	    if (Pattern.matches(regexMMDDYYYY, dateValue)) {
			return true;
		} else if(Pattern.matches(regexDDMMYYYY, dateValue)) {
			return true;
		}else if(Pattern.matches(regexDDMMYY, dateValue)){
			return true;
		}else{
			return false;
		} 
	}
	
	/**<h3>cellFormatedDataValue</h3>
	 * It will used to get the CSV data as a ArrayList<ArrayList<String>>
	 * @author 	<a href="mailto:ssirekumar@gmail.com">Siri Kumar {@literal <ssirekumar@gmail.com>}</a> </br>
	 * @param 	filePath  - Cell class object 
	 * @param   skipLines - integer number to skip the lines.
	 * 
	 * **/
	private static ArrayList<ArrayList<String>> getCSVData(String filePath, int skipLines){
		  LoggerPoire.setLoggerObject(PoireExcelWorkbook._log);
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

}
