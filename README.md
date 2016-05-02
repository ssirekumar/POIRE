POIRE v1.1.0
=====
Author: <b>Ssire (Siri) Kumar Puttagunta</b>

Its a Microsoft Ms-office(Excel,Ms Word,PowerPoint) files handling with Java with the help of the POI classes. This API was develop on the top of the POI library. Its has so many methods for working with the Ms-office files(Excel,Ms Word,PowerPoint). These methods will work on the both the binary(.xls, .doc, .ppt) and XML(.xlsx, .docx, .pptx). with the help of the NPOIFSFileSystem & OPCPackage classes in the POI. with these classes the reading and writing into file would be faster. In this i included so many user friendly methods with the help of POI Library.

<h4>Note:</h4>In this POIRE v1.1.0 build having some methods for word and powerpoint.

Download the .jar and its dependent jar from: https://github.com/ssirekumar/POIRE/releases

<h2>Usages</h2></br>
<ul>
<li>These methods are mainly useful for the Automation testing if you want to read the data from the Excel on diffrent sheets.</li>
<li>These methods will helpful while making any automation framework(Data driven & Hybride) with the Test data as Excel. </li>
<li>These Methods will be helpful in development when ever it require to read/write the data from the excel.</li>
</ul>


<b>Example Methods on Excel:</b>

<ol type="1">
   <li>ArrayList<String> getExcelColumnWithSheetIndex(String filePath, int colNumber, int sheetNumberIndex, boolean
   blankCells);</li>
   <li>ArrayList<String> getExcelColumnWithSheetName(String filePath, int colNumber, String sheetName, boolean  
   blankCells);</li>
   <li>ArrayList<String> getExcelColumnWithHeaderName(String filePath, String columnHeaderName, String sheetName, boolean    blankCells);</li>
   <li>ArrayList<String> getExcelRowWithSheetName(String filePath, int rowNumber, String sheetName, boolean blankCells);</li>
   <li>updateRandomDataInExcelFile(String filePath, int[] columnPositions, int skipLines, int sheetNumberIndex, int
   lengthOfRandomString, boolean randomNumberString);</li>
   <li>updateDataInExcelFile(String filePath, ArrayList<String> rowData int[] columnPositions, int rowNumber, String 
   sheetName);</li>
   <li>ArrayList<String> getExcelRowWithSpecifiedColumn(String filePath, int[] columnPositions, int rowNumber, int
   sheetNumberIndex);</li>
   <li>ArrayList<ArrayList<String>> getExcelDataWithSheetNumber(String filePath, int skipLines, String sheetName);</li>
</ol>


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
<h2>Changes from v1.0.0 to v1.1.0</h2></br>
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

<h2>Issues in POIRE v1.1.0</h2>
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
  
  Output as</br> 
  Text need to be placed with size and font4Text need to be placed with size and font3
  Text need to be placed with size and font2Text need to be placed with size and font1
 ```


Software License Agreement (<a href="https://opensource.org/licenses/BSD-2-Clause">BSD License</a>)</br>
<b>Copyright (c) 2014, Ssire Kumar
All rights reserved.</b>

Redistribution and use in source and binary forms, with or without modification, are permitted provided that the following conditions are met:

1. Redistributions of source code must retain the above copyright notice, this list of conditions and the following disclaimer.
2. Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the following disclaimer in the documentation and/or other materials provided with the distribution.
3. Neither the name of the copyright holder nor the names of its contributors may be used to endorse or promote products derived from this software without specific prior written permission.

THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
