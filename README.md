POIRE (1.1.0)
=====
Author: Ssire Kumar Puttagunta

Its a Microsoft Ms-office(Excel,Ms Word,PowerPoint) files handling with Java with the help of the POI classes. This API was develop on the top of the POI library. Its has so many methods for working with the Ms-office files(Excel,Ms Word,PowerPoint). These methods will work on the both the binary(.xls, .doc, .ppt) and XML(.xlsx, .docx, .pptx). with the help of the NPOIFSFileSystem & OPCPackage classes in the POI. with these classes the reading and writing into file would be faster. In this i included so many user friendly methods with the help of POI Library.

<h4>Note:</h4>In this POIRE (1.1.0) build having some methods for word and powerpoint.

<h2>Usages</h2></br>
<ul>
<li>These methods are mainly useful for the Automation testing if you want to read the data from the Excel on diffrent sheets.</li>
<li>These methods will helpful while making any automation framework(Data driven & Hybride) with the Test data as Excel. </li>
<li>These Methods will be helpful in development when ever it require to read/write the data from the excel.</li>
</ul>


<b>Example Methods on Excel:</b>

```java
1). ArrayList<String> getExcelColumnWithSheetIndex(String filePath, int colNumber, int sheetNumberIndex, boolean blankCells)

2). ArrayList<String> getExcelColumnWithSheetName(String filePath, int colNumber, String sheetName, boolean blankCells)

3). ArrayList<String> getExcelColumnWithHeaderName(String filePath, String columnHeaderName, String sheetName, boolean blankCells)

4). ArrayList<String> getExcelRowWithSheetName(String filePath, int rowNumber, String sheetName, boolean blankCells)

5). updateRandomDataInExcelFile(String filePath, int[] columnPositions, int skipLines, int sheetNumberIndex, int lengthOfRandomString, boolean randomNumberString)

6). updateDataInExcelFile(String filePath, ArrayList<String> rowData int[] columnPositions, int rowNumber, String sheetName)

7). ArrayList<String> getExcelRowWithSpecifiedColumn(String filePath, int[] columnPositions, int rowNumber, int sheetNumberIndex)

8). ArrayList<ArrayList<String>> getExcelDataWithSheetNumber(String filePath, int skipLines, String sheetName)
```
<h2>Features</h2></br>
      <ul>
	<li>Create a new XLS, XLSX excel sheet.</li>
	<li>Get the data as per the excel format.If the cell having a formula it can be formated while retriving the data.</li>
	<li>Get the Excel Column data as a ArrayList&lt;String&gt; With specified column number/column Header and its sheet number/Name.</li>
	<li>Get the Excel row data as a ArrayList&lt;String&gt; With specified row Number/row header and its number/Name.</li>
	<li>Update random data in the existing Excel row with specified columnPositions values. update data may be random string or number can be possible.</li>
	<li>Update the data with specified arrayList in existing Excel row with columnPositions values.</li>
	<li>Update the data with specified arrayList in existing Excel row data with columnPositions values.</li>
	<li>Create/Update the data with specified arrayList based on sheet number insed in existing Excel.</li>
	<li>Create/Update the row with specified arrayList based on the sheet name in existing Excel.</li>
	<li>Create/Update the row from column position with specified arrayList based on the sheet name/sheet Number Index in existing Excel.</li>
	<li>Create/Update the row with same data from column position with specified rowdata based on the sheet name in existing Excel.</li>
</ul>
<p>&nbsp;</p>
<h2>Changes from V1.0.0 to POIRE(1.1.0)</h2></br>
<b>Excel:</b>
<ol type="1">
	<li>Added converted methods from csv to excel and excel to csv.</li>
	<li>Added a methods to convert the sheet to new excel file and same way multiple array of selected sheet
	        from the excel file to new file.</li></br>
  	<li>Added a new overloded method for removeSheet with array of sheet names given.</li>
	<li>Added a method to create the excel file with given array of sheet names.</li>
</ol>
<b>Word:</b></br>
<ol>
	<li>Added a method to set the margin of the document.</li>
  	<li>Added a method to create the doc file.</li>
	<li>Added a method to create a paragraph in the specified document.</li>
</ol>
<b>Powerpoint:</b></br>
<ol>
	<li>Added a method for create a new powerpoint file.</li>
</ol>

<h2>Issues in POIRE(1.1.0)</h2>
<ol type="1">
	<li>While seting the margins for .doc files it will throw some exception.</li>
   	<li>While creating the paragraphs for .doc file one after other it will create as a single paragraph.</li>
</ol>

```java
  File _fileWordObj2 = WordDataEngine.createDocumentFile(false, "C:\\TestExcel", "Test123");
  WordDataEngine.createParagraph(_fileWordObj2.getAbsolutePath(), "Text need to be placed with size and font1");
  WordDataEngine.createParagraph(_fileWordObj2.getAbsolutePath(), "Text need to be placed with size and font2");
  WordDataEngine.createParagraph(_fileWordObj2.getAbsolutePath(), "Text need to be placed with size and font3");
  WordDataEngine.createParagraph(_fileWordObj2.getAbsolutePath(), "Text need to be placed with size and font4");
 ```
Output as 
  Text need to be placed with size and font4Text need to be placed with size and font3
  Text need to be placed with size and font2Text need to be placed with size and font1
