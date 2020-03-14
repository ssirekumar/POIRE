package com.walnut.poire;


import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;

import org.testng.Assert;
import org.testng.Reporter;
import org.testng.annotations.Test;
import org.testng.asserts.SoftAssert;

import com.walnut.poire.sheets.PoireExcelWorkbook;

/**
 * Unit test for POIRE.
 */
public class PoireTests {
	PoireExcelWorkbook _obj = new PoireExcelWorkbook(new File(System.getProperty("user.dir")+"//sampleSheets//sampleSheets.xlsx"));
	
	//Get column data tests.
	
    @Test(enabled = true, 
    		description = "Validating 'getExcelColumnWithSheetIndex' method"
    )
    public void getExcelColumnWithSheetIndex() {
    	int size = 0;
    	boolean truefalse = false;
    	SoftAssert softAssert = new SoftAssert();
    	ArrayList<String> validationList = new ArrayList<>();
    	ArrayList<String> list = new ArrayList<>(Arrays.asList("SampleData","","","","","","","","","SampleData"));
    	ArrayList<String> list2 = new ArrayList<>(list.subList(0, 2));
    	ArrayList<String> list3 = new ArrayList<>();
    	ArrayList<String> list4 = new ArrayList<>(Arrays.asList("SampleData","SampleData"));
    	list3.add(list.get(0));
    	list3.add(list.get(9));
    	
    	validationList = _obj.getExcelColumnWithSheetIndex(1, 1, true); 
    	softAssert.assertEquals(validationList, Collections.EMPTY_LIST, "Not return EMPTY_LIST");
    	Reporter.log("Asserting for returning EMPTY_LIST or not if sheet doesn't have a data :" + "Actual: "+ validationList + " Expected: "+Collections.EMPTY_LIST);
    	
    	validationList = _obj.getExcelColumnWithSheetIndex(1, 1, false);
    	softAssert.assertEquals(validationList, Collections.EMPTY_LIST, "Not return EMPTY_LIST");
    	Reporter.log("Asserting for returning EMPTY_LIST or not if sheet doesn't have a data with blankCell = false :" + "Actual: "+ validationList + " Expected: "+Collections.EMPTY_LIST);
    	
    	size = _obj.getExcelColumnWithSheetIndex(0, 0, true).size();
    	softAssert.assertEquals(size, list.size());
    	Reporter.log("Asserting for returning data and its should matches the size of 'list' :" + "Actual: "+ size + " Expected: "+list.size());
    	
    	size = _obj.getExcelColumnWithSheetIndex(0, 0, false).size();
    	softAssert.assertEquals(size, list2.size());
    	Reporter.log("Asserting for returning data and its should matches the size of 'list2' : " + "Actual: "+ size + " Expected: "+list2.size());
    	
    	truefalse = _obj.getExcelColumnWithSheetIndex(0, 0, true).containsAll(list);
    	softAssert.assertEquals(truefalse, true);
    	Reporter.log("Asserting for returning data, and its should  == list of all the elements : " + "Actual: "+ truefalse + " Expected: "+ true);
    	
    	truefalse = _obj.getExcelColumnWithSheetIndex(0, 0, false).containsAll(list3);
    	softAssert.assertEquals(truefalse, true);
    	Reporter.log("Asserting for returning data, and its should  == list3 of all the elements : " + "Actual: "+ truefalse + " Expected: "+ true);
    	
    	//print all the data from the sheet with spaces
    	validationList = _obj.getExcelColumnWithSheetIndex(0, 0, true);
    	softAssert.assertEquals(validationList.toString(), list.toString());
    	Reporter.log("Asserting: Print all the data from the sheet " + "Actual: "+ validationList + " Expected: "+ list);
    	
    	//print all the data from the sheet without spaces
    	validationList = _obj.getExcelColumnWithSheetIndex(0, 0, false);
    	softAssert.assertEquals(validationList.toString(), list4.toString());
    	Reporter.log("Asserting: Print all the data from the sheet " + "Actual: "+ validationList + " Expected: "+ list4);
    	
    	//Negative scenarios
    	validationList = _obj.getExcelColumnWithSheetIndex(-1, -2, false);
    	softAssert.assertEquals(validationList, Collections.EMPTY_LIST, "Not return EMPTY_LIST");
    	Reporter.log("Asserting invalid parameters data, expected exception as 'IllegalArgumentException' ");
    	softAssert.assertAll();
    }
    
    @Test(enabled = true, 
    		description = "Validating 'getExcelColumnWithSheetName' method"
    )
    public void getExcelColumnWithSheetName() {
    	int size = 0;
    	boolean truefalse = false;
    	SoftAssert softAssert = new SoftAssert();
    	ArrayList<String> validationList = new ArrayList<>();
    	ArrayList<String> list = new ArrayList<>(Arrays.asList("SampleData","","","","","","","","","SampleData"));
    	ArrayList<String> list2 = new ArrayList<>(list.subList(0, 2));
    	ArrayList<String> list3 = new ArrayList<>();
    	ArrayList<String> list4 = new ArrayList<>(Arrays.asList("SampleData","SampleData"));
    	list3.add(list.get(0));
    	list3.add(list.get(9));
    	
    	validationList = _obj.getExcelColumnWithSheetName(1, "emptySheet", true); 
    	softAssert.assertEquals(validationList, Collections.EMPTY_LIST, "Not return EMPTY_LIST");
    	Reporter.log("Asserting for returning EMPTY_LIST or not if sheet doesn't have a data :" + "Actual: "+ validationList + " Expected: "+Collections.EMPTY_LIST);
    	
    	validationList = _obj.getExcelColumnWithSheetName(1, "emptySheet", false);
    	softAssert.assertEquals(validationList, Collections.EMPTY_LIST, "Not return EMPTY_LIST");
    	Reporter.log("Asserting for returning EMPTY_LIST or not if sheet doesn't have a data with blankCell = false :" + "Actual: "+ validationList + " Expected: "+Collections.EMPTY_LIST);
    	
    	size = _obj.getExcelColumnWithSheetName(0, "DataSheet", true).size();
    	softAssert.assertEquals(size, list.size());
    	Reporter.log("Asserting for returning data and its should matches the size of 'list' :" + "Actual: "+ size + " Expected: "+list.size());
    	
    	size = _obj.getExcelColumnWithSheetName(0, "DataSheet", false).size();
    	softAssert.assertEquals(size, list2.size());
    	Reporter.log("Asserting for returning data and its should matches the size of 'list2' : " + "Actual: "+ size + " Expected: "+list2.size());
    	
    	truefalse = _obj.getExcelColumnWithSheetName(0, "DataSheet", true).containsAll(list);
    	softAssert.assertEquals(truefalse, true);
    	Reporter.log("Asserting for returning data, and its should  == list of all the elements : " + "Actual: "+ truefalse + " Expected: "+ true);
    	
    	truefalse = _obj.getExcelColumnWithSheetName(0, "DataSheet", false).containsAll(list3);
    	softAssert.assertEquals(truefalse, true);
    	Reporter.log("Asserting for returning data, and its should  == list3 of all the elements : " + "Actual: "+ truefalse + " Expected: "+ true);
    	
    	//print all the data from the sheet
    	validationList = _obj.getExcelColumnWithSheetName(0, "DataSheet", true);
    	softAssert.assertEquals(validationList.toString(), list.toString());
    	Reporter.log("Asserting: Print all the data from the sheet " + "Actual: "+ validationList + " Expected: "+ list);
    	
    	//print all the data from the sheet
    	validationList = _obj.getExcelColumnWithSheetName(0, "DataSheet", false);
    	softAssert.assertEquals(validationList.toString(), list4.toString());
    	Reporter.log("Asserting: Print all the data from the sheet " + "Actual: "+ validationList + " Expected: "+ list4);
    	
    	//Negative scenarios
    	validationList = _obj.getExcelColumnWithSheetIndex(-1, -2, false);
    	softAssert.assertEquals(validationList, Collections.EMPTY_LIST, "Not return EMPTY_LIST");
    	Reporter.log("Asserting invalid parameters data, expected exception as 'IllegalArgumentException' ");
    	softAssert.assertAll();
    }
    
    @Test(enabled = true, 
    		description = "Validating 'getExcelColumnWithHeaderName' method"
    )
    public void getExcelColumnWithHeaderName() {
    	int size = 0;
    	boolean truefalse = false;
    	SoftAssert softAssert = new SoftAssert();
    	ArrayList<String> validationList = new ArrayList<>();
    	ArrayList<String> list = new ArrayList<>(Arrays.asList("","Data","","","","","","","data"));
    	ArrayList<String> list2 = new ArrayList<>(list.subList(0, 2));
    	ArrayList<String> list3 = new ArrayList<>();
    	ArrayList<String> list4 = new ArrayList<>(Arrays.asList("Data","data"));
    	list3.add(list.get(1));
    	list3.add(list.get(8));
    	
    	validationList = _obj.getExcelColumnWithHeaderName("HeaderName", "emptySheet", true); 
    	softAssert.assertEquals(validationList, Collections.EMPTY_LIST, "Not return EMPTY_LIST");
    	Reporter.log("Asserting for returning EMPTY_LIST or not if sheet doesn't have a data :" + "Actual: "+ validationList + " Expected: "+Collections.EMPTY_LIST);
    	
    	validationList = _obj.getExcelColumnWithHeaderName("HeaderName", "emptySheet", false);
    	softAssert.assertEquals(validationList, Collections.EMPTY_LIST, "Not return EMPTY_LIST");
    	Reporter.log("Asserting for returning EMPTY_LIST or not if sheet doesn't have a data with blankCell = false :" + "Actual: "+ validationList + " Expected: "+Collections.EMPTY_LIST);
    	
    	size = _obj.getExcelColumnWithHeaderName("HeaderName", "DataSheet", true).size();
    	softAssert.assertEquals(size, list.size());
    	Reporter.log("Asserting for returning data and its should matches the size of 'list' :" + "Actual: "+ size + " Expected: "+list.size());
    	
    	size = _obj.getExcelColumnWithHeaderName("HeaderName", "DataSheet", false).size();
    	softAssert.assertEquals(size, list2.size());
    	Reporter.log("Asserting for returning data and its should matches the size of 'list2' : " + "Actual: "+ size + " Expected: "+list2.size());
    	
    	truefalse = _obj.getExcelColumnWithHeaderName("HeaderName", "DataSheet", true).containsAll(list);
    	softAssert.assertEquals(truefalse, true);
    	Reporter.log("Asserting for returning data, and its should  == list of all the elements : " + "Actual: "+ truefalse + " Expected: "+ true);
    	
    	truefalse = _obj.getExcelColumnWithHeaderName("HeaderName", "DataSheet", false).containsAll(list3);
    	softAssert.assertEquals(truefalse, true);
    	Reporter.log("Asserting for returning data, and its should  == list3 of all the elements : " + "Actual: "+ truefalse + " Expected: "+ true);
    	
    	//print all the data from the sheet
    	validationList = _obj.getExcelColumnWithHeaderName("HeaderName", "DataSheet", true);
    	softAssert.assertEquals(validationList.toString(), list.toString());
    	Reporter.log("Asserting: Print all the data from the sheet " + "Actual: "+ validationList + " Expected: "+ list);
    	
    	//print all the data from the sheet
    	validationList = _obj.getExcelColumnWithHeaderName("HeaderName", "DataSheet", false);
    	softAssert.assertEquals(validationList.toString(), list4.toString());
    	Reporter.log("Asserting: Print all the data from the sheet " + "Actual: "+ validationList + " Expected: "+ list4);
    	
    	//Negative scenarios
    	validationList = _obj.getExcelColumnWithHeaderName("HeaderNam", "DataShee", false);
    	softAssert.assertEquals(validationList, Collections.EMPTY_LIST, "Not return EMPTY_LIST");
    	Reporter.log("Asserting invalid parameters data, expected exception as 'IllegalArgumentException' ");
    	softAssert.assertAll();
    }
    
    @Test(enabled = true, 
    		description = "getExcelColumnWithHeaderSheetIndex"
    )
    public void getExcelColumnWithHeaderSheetIndex() {
    	int size = 0;
    	boolean truefalse = false;
    	SoftAssert softAssert = new SoftAssert();
    	ArrayList<String> validationList = new ArrayList<>();
    	ArrayList<String> list = new ArrayList<>(Arrays.asList("","Data","","","","","","","data"));
    	ArrayList<String> list2 = new ArrayList<>(list.subList(0, 2));
    	ArrayList<String> list3 = new ArrayList<>();
    	ArrayList<String> list4 = new ArrayList<>(Arrays.asList("Data","data"));
    	list3.add(list.get(1));
    	list3.add(list.get(8));
    	
    	validationList = _obj.getExcelColumnWithHeaderSheetIndex("HeaderName", 1, true); 
    	softAssert.assertEquals(validationList, Collections.EMPTY_LIST, "Not return EMPTY_LIST");
    	Reporter.log("Asserting for returning EMPTY_LIST or not if sheet doesn't have a data :" + "Actual: "+ validationList + " Expected: "+Collections.EMPTY_LIST);
    	
    	validationList = _obj.getExcelColumnWithHeaderSheetIndex("HeaderName", 1, false);
    	softAssert.assertEquals(validationList, Collections.EMPTY_LIST, "Not return EMPTY_LIST");
    	Reporter.log("Asserting for returning EMPTY_LIST or not if sheet doesn't have a data with blankCell = false :" + "Actual: "+ validationList + " Expected: "+Collections.EMPTY_LIST);
    	
    	size = _obj.getExcelColumnWithHeaderSheetIndex("HeaderName", 0, true).size();
    	softAssert.assertEquals(size, list.size());
    	Reporter.log("Asserting for returning data and its should matches the size of 'list' :" + "Actual: "+ size + " Expected: "+list.size());
    	
    	size = _obj.getExcelColumnWithHeaderSheetIndex("HeaderName", 0, false).size();
    	softAssert.assertEquals(size, list2.size());
    	Reporter.log("Asserting for returning data and its should matches the size of 'list2' : " + "Actual: "+ size + " Expected: "+list2.size());
    	
    	truefalse = _obj.getExcelColumnWithHeaderSheetIndex("HeaderName", 0, true).containsAll(list);
    	softAssert.assertEquals(truefalse, true);
    	Reporter.log("Asserting for returning data, and its should  == list of all the elements : " + "Actual: "+ truefalse + " Expected: "+ true);
    	
    	truefalse = _obj.getExcelColumnWithHeaderSheetIndex("HeaderName", 0, false).containsAll(list3);
    	softAssert.assertEquals(truefalse, true);
    	Reporter.log("Asserting for returning data, and its should  == list3 of all the elements : " + "Actual: "+ truefalse + " Expected: "+ true);
    	
    	//print all the data from the sheet
    	validationList = _obj.getExcelColumnWithHeaderSheetIndex("HeaderName", 0, true);
    	softAssert.assertEquals(validationList.toString(), list.toString());
    	Reporter.log("Asserting: Print all the data from the sheet " + "Actual: "+ validationList + " Expected: "+ list);
    	
    	//print all the data from the sheet
    	validationList = _obj.getExcelColumnWithHeaderSheetIndex("HeaderName", 0, false);
    	softAssert.assertEquals(validationList.toString(), list4.toString());
    	Reporter.log("Asserting: Print all the data from the sheet " + "Actual: "+ validationList + " Expected: "+ list4);
    	
    	//Negative scenarios
    	validationList = _obj.getExcelColumnWithHeaderSheetIndex("HeaderNam", -1, false);
    	softAssert.assertEquals(validationList, Collections.EMPTY_LIST, "Not return EMPTY_LIST");
    	Reporter.log("Asserting invalid parameters data, expected exception as 'IllegalArgumentException' ");
    	softAssert.assertAll();
    }
    
    @Test(enabled = true, 
    		description = "finding if file not found", 
    		expectedExceptions = {FileNotFoundException.class}
    )
    public void findingIfFileNotfound() throws FileNotFoundException, IOException {
    	_obj.setExcelFilePath(System.getProperty("user.dir")+"//sampleSheets","sampleSheet.xlsx");
    }
    
    @Test(enabled = true, 
    		description = "Wronglly given excel file path"
    )
    public void changeFilePath() throws IOException {
    	_obj.setExcelFilePath(System.getProperty("user.dir")+"//sampleSheets//sampleSheets.xlsx");
    	Assert.assertEquals(_obj.getExcelFilePath(), System.getProperty("user.dir")+"/sampleSheets/sampleSheets.xlsx");
    	Reporter.log("Assert for Wronglly given excel file path");
    }
    
    
    //Get row data test
    @Test(enabled = true, 
    		description = "Validating 'getExcelRowWithSheetIndex' method")
    public void getExcelRowWithSheetIndex() {
    	int size = 0;
    	boolean truefalse = false;
    	SoftAssert softAssert = new SoftAssert();
    	ArrayList<String> validationList = new ArrayList<>();
    	ArrayList<String> list = new ArrayList<>(Arrays.asList("SampleData","HeaderName","","","","","","HeaderName2"));
    	ArrayList<String> list2 = new ArrayList<>(list.subList(0, 3));
    	ArrayList<String> list3 = new ArrayList<>();
    	list3.add(list.get(0));
    	list3.add(list.get(1));
    	list3.add(list.get(7));
    	ArrayList<String> list4 = new ArrayList<>(Arrays.asList("SampleData","HeaderName","HeaderName2"));
    	
    	validationList = _obj.getExcelRowWithSheetIndex(1, 1, true); 
    	softAssert.assertEquals(validationList, Collections.EMPTY_LIST, "Not return EMPTY_LIST");
    	Reporter.log("Asserting for returning EMPTY_LIST or not if sheet doesn't have a data :" + "Actual: "+ validationList + " Expected: "+Collections.EMPTY_LIST);
    	
    	validationList = _obj.getExcelRowWithSheetIndex(1, 1, false);
    	softAssert.assertEquals(validationList, Collections.EMPTY_LIST, "Not return EMPTY_LIST");
    	Reporter.log("Asserting for returning EMPTY_LIST or not if sheet doesn't have a data with blankCell = false :" + "Actual: "+ validationList + " Expected: "+Collections.EMPTY_LIST);
    	
    	size = _obj.getExcelRowWithSheetIndex(0, 0, true).size();
    	softAssert.assertEquals(size, list.size());
    	Reporter.log("Asserting for returning data and its should matches the size of 'list' :" + "Actual: "+ size + " Expected: "+list.size());
    	
    	size = _obj.getExcelRowWithSheetIndex(0, 0, false).size();
    	softAssert.assertEquals(size, list2.size());
    	Reporter.log("Asserting for returning data and its should matches the size of 'list2' : " + "Actual: "+ size + " Expected: "+list2.size());
    	
    	truefalse = _obj.getExcelRowWithSheetIndex(0, 0, true).containsAll(list);
    	softAssert.assertEquals(truefalse, true);
    	Reporter.log("Asserting for returning data, and its should  == list of all the elements : " + "Actual: "+ truefalse + " Expected: "+ true);
    	
    	truefalse = _obj.getExcelRowWithSheetIndex(0, 0, false).containsAll(list3);
    	softAssert.assertEquals(truefalse, true);
    	Reporter.log("Asserting for returning data, and its should  == list3 of all the elements : " + "Actual: "+ truefalse + " Expected: "+ true);
    	
    	//print all the data from the sheet with spaces
    	validationList = _obj.getExcelRowWithSheetIndex(0, 0, true);
    	softAssert.assertEquals(validationList.toString(), list.toString());
    	Reporter.log("Asserting: Print all the data from the sheet " + "Actual: "+ validationList + " Expected: "+ list);
    	
    	//print all the data from the sheet without spaces
    	validationList = _obj.getExcelRowWithSheetIndex(0, 0, false);
    	softAssert.assertEquals(validationList.toString(), list4.toString());
    	Reporter.log("Asserting: Print all the data from the sheet " + "Actual: "+ validationList + " Expected: "+ list4);
    	
    	//Negative scenarios
    	validationList = _obj.getExcelRowWithSheetIndex(-1, -2, false);
    	softAssert.assertEquals(validationList, Collections.EMPTY_LIST, "Not return EMPTY_LIST");
    	Reporter.log("Asserting invalid parameters data, expected exception as 'IllegalArgumentException' ");
    	softAssert.assertAll();
    }
    
    
    
}
