POIRE
=====
Author: Ssire Kumar Puttagunta

Its a Microsoft Ms-office file handling with Java with the help of the POI classes. This API was develop on the top
of the POI library. Its has so many methods for working with the excel files. These methods will work on the both the 
Excel formats .xls and .xlsx files.

This excel library API mainly usefull for developmet and mainly useful for Testing Automation.

<b>Example Methods on Excel:</b>

1). ArrayList<String> getExcelColumnWithSheetIndex(String filePath, int colNumber, int sheetNumberIndex, boolean blankCells)

2). ArrayList<String> getExcelColumnWithSheetName(String filePath, int colNumber, String sheetName, boolean blankCells)

3). ArrayList<String> getExcelColumnWithHeaderName(String filePath, String columnHeaderName, String sheetName, boolean blankCells)

4). ArrayList<String> getExcelRowWithSheetName(String filePath, int rowNumber, String sheetName, boolean blankCells)

5). updateRandomDataInExcelFile(String filePath, int[] columnPositions, int skipLines, int sheetNumberIndex, int lengthOfRandomString, boolean randomNumberString)

6). updateDataInExcelFile(String filePath, ArrayList<String> rowData int[] columnPositions, int rowNumber, String sheetName)

7). ArrayList<String> getExcelRowWithSpecifiedColumn(String filePath, int[] columnPositions, int rowNumber, int sheetNumberIndex)

8). ArrayList<ArrayList<String>> getExcelDataWithSheetNumber(String filePath, int skipLines, String sheetName)




