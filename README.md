POIRE
=====
Author: Ssire Kumar Puttagunta

Its a Microsoft Ms-office file handling with Java with the help of the POI classes. This API was develop on the top
of the POI library. Its has so many methods for working with the excel files. These methods will work on the both the 
Excel formats .xls and .xlsx files.

In this i was included so many user friendly methods with the help of POI Library. This API will supprot all the excel formats like .xls and .xlsx with the help of the  NPOIFSFileSystem & OPCPackage classes in the POI. with these classes the reading and writing into excel file would be faster.

These methods are mainly useful for the Automation testing if you want to read the data from the Excel on diffrent sheets.

These methods will helpful while making any automation framework(Data driven & Hybride) with the Test data as Excel.

These Methods will be helpful in development when ever it require to read/write the data from the excel. 



<b>Example Methods on Excel:</b>

1). ArrayList<String> getExcelColumnWithSheetIndex(String filePath, int colNumber, int sheetNumberIndex, boolean blankCells)

2). ArrayList<String> getExcelColumnWithSheetName(String filePath, int colNumber, String sheetName, boolean blankCells)

3). ArrayList<String> getExcelColumnWithHeaderName(String filePath, String columnHeaderName, String sheetName, boolean blankCells)

4). ArrayList<String> getExcelRowWithSheetName(String filePath, int rowNumber, String sheetName, boolean blankCells)

5). updateRandomDataInExcelFile(String filePath, int[] columnPositions, int skipLines, int sheetNumberIndex, int lengthOfRandomString, boolean randomNumberString)

6). updateDataInExcelFile(String filePath, ArrayList<String> rowData int[] columnPositions, int rowNumber, String sheetName)

7). ArrayList<String> getExcelRowWithSpecifiedColumn(String filePath, int[] columnPositions, int rowNumber, int sheetNumberIndex)

8). ArrayList<ArrayList<String>> getExcelDataWithSheetNumber(String filePath, int skipLines, String sheetName)

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



