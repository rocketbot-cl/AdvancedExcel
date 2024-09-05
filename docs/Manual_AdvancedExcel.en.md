



# Excel Advanced Options
  
Apply automatic and advanced filters, format cells, add or delete sheets, rows or columns, export to different file formats, unlock and relock sheets, copy and paste special and more with your Excel files.  

*Read this in other languages: [English](Manual_AdvancedExcel.md), [Português](Manual_AdvancedExcel.pr.md), [Español](Manual_AdvancedExcel.es.md)*
  
![banner](imgs/Banner_AdvancedExcel.png)
## How to install this module
  
To install the module in Rocketbot Studio, it can be done in two ways:
1. Manual: __Download__ the .zip file and unzip it in the modules folder. The folder name must be the same as the module and inside it must have the following files and folders: \__init__.py, package.json, docs, example and libs. If you have the application open, refresh your browser to be able to use the new module.
2. Automatic: When entering Rocketbot Studio on the right margin you will find the **Addons** section, select **Install Mods**, search for the desired module and press install.  

## How to use this module
To use this module, you must have Microsoft Excel installed.


## Description of the commands

### Open Without Alerts
  
Open a file preventing MS Excel alerts.
|Parameters|Description|example|
| --- | --- | --- |
|path to XLSX file |Path of the xlsx file to be opened|file.XLSX|
|Password (optional) |Password of the xlsx file|P@ssW0rd|
|Id (optional) |Name or identifier for the file to be opened. Used when you need to open more than one excel. By default is *default*.|id|
|Assign result to variable |Variable where the result will be stored|id|

### Find and Connect
  
Search a Excel Book opened and connect it
|Parameters|Description|example|
| --- | --- | --- |
|XLSX filename ||File.XLSX|
|Id (optional) |Name or identifier for the file to be opened. Used when you need to open more than one excel. By default is *default*.|excel1|

### Maximize
  
Maximize Excel Window
|Parameters|Description|example|
| --- | --- | --- |

### Calculation options
  
Select the way the formula calculation is executed in the workbook.
|Parameters|Description|example|
| --- | --- | --- |
|Calculation options|Select the book calculation method.||
|Calculate now|If this box is checked, the formulas in the book are calculated|True|

### Read cells
  
Read a cell or range of cells
|Parameters|Description|example|
| --- | --- | --- |
|Sheet |Name of the sheet where the data are located|Sheet1|
|Cells |Cell or Range of cells. The syntax must be the same as excel (A1 or A1B1) |A1:B5|
|Format|Select the format to pull cells that contain dates. Select custom to add a custom format|dd-mm-yy|
|Custom format |Custom format. Doc https//docs.python.org/3/library/datetime.html#strftime-and-strptime-format-codes|'%m/%d/%y %I:%M %p'|
|Assign result to variable |Name of the variable where the result will be stored|cells|

### Convert serial date
  
Convert an excel serial number date to a specific date format
|Parameters|Description|example|
| --- | --- | --- |
|Serial date |Excel serial number that represents a specific date, where 1 = 01/01/1900|44927|
|Output format |Date format to convert the serial date to|%d/%m/%y|
|Assign result to variable |Name of the variable where the result will be stored|output_date|

### Count columns
  
Count the columns or return the last column name. It's necessary that the excel is saved to get the last changes
|Parameters|Description|example|
| --- | --- | --- |
|Sheet |Name of the sheet where the data are located|Sheet1|
|Get address |If this box is checked, it will return the letter of the last column|True|
|Assign result to variable |Name of the variable where the result will be stored|count_columns|

### Count Rows
  
Counts all the rows or from a range.
|Parameters|Description|example|
| --- | --- | --- |
|Sheet|Name of the sheet where the data are located|Sheet1|
|Count all rows|Option to count all rows.||
|Column |Column where the rows will be counted|C|
|Assign result to variable |Name of the variable where the result will be stored|length_rows|

### Hide
  
Hides one or more rows, or one or more columns.
|Parameters|Description|example|
| --- | --- | --- |
|Sheet |Name of the sheet|Sheet1|
|Range |For a range of rows use numbers separated by two dots (13). For a range of columns use letters (AB)|1:3|
|Assign result to variable |Name of the variable where the result will be stored|res|

### Show
  
Shows one or more rows, or one or more columns that are hidden
|Parameters|Description|example|
| --- | --- | --- |
|Sheet |Name of the sheet|Sheet1|
|Range |For a range of rows use numbers separated by two dots (13). For a range of columns use letters (AB)|1:3|
|Assign result to variable |Name of the variable where the result will be stored|res|

### Cell color
  
Change color of a cell or range of cells. Can be a default color or custom
|Parameters|Description|example|
| --- | --- | --- |
|Cells |Cell or Range of cells. The syntax must be the same as excel (A1 or A1B1) |A1:B5|
|Sheet |Workbook sheet|Sheet1|
|Entire sheet|If this box is checked, the color will be applied to the entire sheet.||
|RGB color cell |RGB values of the color that will be the cell or cells|250,250,250|
|Select color |Select the color. You can use the previous input to customize the color|red|

### Get Cell Color
  
Get the color of a cell. The funtion will return a list of two elements: Background Color and Font Color in RGB format.
|Parameters|Description|example|
| --- | --- | --- |
|Sheet |Sheet|Sheet1|
|Cell |Cell. The syntax must be the same as excel (A1)|A1|
|Assign to var|Name of the variable where the result will be stored|color|

### Get Cell Formats
  
Get the format of a cell. The function will return a dictionary with the cell properties and the value of each one.
|Parameters|Description|example|
| --- | --- | --- |
|Sheet |Sheet|Sheet1|
|Cell |Cell. The syntax must be the same as excel (A1)|A1|
|Assign to var|Name of the variable where the result will be stored|color|

### Insert Formula
  
Insert formula into cell
|Parameters|Description|example|
| --- | --- | --- |
|Sheet |Sheet|Sheet5|
|Cell |Cell or Range of cells. The syntax must be the same as excel (A1 or A1B1) |A5|
|Write formula |Formula that will be inserted. Must be writen in english. Remember to use *,* to separate the parameters|=SUM(A1:A4)|
|No IIE|If this box is checked, it allows you to send the formula without IIE|True|

### Insert Macro
  
Insert Macro in Excel
|Parameters|Description|example|
| --- | --- | --- |
|Macro path |Path of the bas file to be inserted|Macro.bas|

### Select and copy Cells
  
Select and Copy cells in Excel
|Parameters|Description|example|
| --- | --- | --- |
|Sheet |Name of the sheet to automate|Hoja 1|
|Enter cells to select |Cell or Range of cells. The syntax must be the same as excel (A1 or A1B1) |A1:B3|
|Copy |By checking the checkbox, the values will be copied to the clipboard.|True|

### Get Cell With Currency Format
  
Get cells with currency format
|Parameters|Description|example|
| --- | --- | --- |
|Sheet |Name of the sheet to automate|Hoja 1|
|Enter cells to select |Cell or Range of cells. The syntax must be the same as excel (A1 or A1B1) |A1:B3|
|Assign to var|Name of the variable where the result will be stored|variable|

### Get Cell With Date Format
  
Get cells with date format
|Parameters|Description|example|
| --- | --- | --- |
|Sheet |Name of the sheet to automate|Hoja 1|
|Enter cells to select |Cell or Range of cells. The syntax must be the same as excel (A1 or A1B1) |A1:B3|
|Assign to var|Name of the variable where the result will be stored|variable|

### Copy-Paste
  
Copy range cell to another sheet
|Parameters|Description|example|
| --- | --- | --- |
|Origin Sheet|Name of the sheet to automate|Sheet1|
|Range to copy|Cell or Range of cells to copy. The syntax must be the same as excel (A1 or A1B1) |A1:C4|
|Destiny Sheet|Name of the destiny sheet|Sheet2|
|Range to paste|Cell or Range of cells to paste. The syntax must be the same as excel (A1 or A1B1) |A1:C4|
|Paste Option|Select paste type for the cell or cells range.|Option|
|Paste Operation|Select paste operation for the cell or cells range.|Operation|
|Skip Blanks|Prevents replacing values in the paste area when blank cells are produced in the copy area when this box is selected.||
|Transpose|Rotate the content of copied cells when pasting. Data in rows will be pasted into columns and vice versa.||

### Format Cell
  
Format Cell
|Parameters|Description|example|
| --- | --- | --- |
|Sheet Name |Name of the sheet to automate|Sheet1|
|Cell |Cell or Range of cells. The syntax must be the same as excel (A1 or A1B1) |A1:C4|
|Format|You should select the formatting type for the cell. Select custom to add a customizable format|dd-mm-yy|
|Custom format |Custom format. Doc https//support.microsoft.com/en-us/office/review-guidelines-for-customizing-a-number-format-c0a1d1fa-d3f4-4018-96b7-9c9354dd99f5|00000|
|Text to Value|||

### Clear Contents
  
Clears formulas and values from the selected range, keeping the format.
|Parameters|Description|example|
| --- | --- | --- |
|Sheet |Name of the sheet where the data are located|Sheet1|
|Cell or Range of cells|Range that contains the data to align|A1:D7|

### Create Sheet
  
Create sheet in the end
|Parameters|Description|example|
| --- | --- | --- |
|Sheet name|Name of the sheet to create|Sheet2|
|After|The sheet will be created next to the sheet indicated in this field.|Hoja1|

### Delete Sheet
  
Delete sheet
|Parameters|Description|example|
| --- | --- | --- |
|Sheet name|Name of the sheet to delete|Sheet2|
|Assign result to variable|Name of the variable where the result will be stored|Variable|

### Copy to another excel
  
Copy the range from one Excel file to another. Indicating the file path, it will open excel to copy or paste the data. If you enter the id of an open excel, it will use that instance to copy or paste.
|Parameters|Description|example|
| --- | --- | --- |
|Excel origin (optional)|Path of the source xlsx file|Source file path:|
|Id (optional) |Name or ID of the open source file.|id|
|Sheet name|Name of the source sheet|Sheet1|
|Range to copy|Cell or Range of cells. The syntax must be the same as excel (A1 or A1B1) |A1:D7|
|Excel destination|Path of the destination xlsx file|Target file path:|
|Normal opening|If this checkbox is checked, the destination file opens normally keeping data, formats and objects. Otherwise it only recovers data.|True|
|Only Values|If this checkbox is checked, only the value will be copied|True|
|Sheet name|Name of the sheet to be copied|Sheet1|
|Range to paste (Optional)|Column, Cell or Range of cells where to paste. The syntax must be the same as excel (A, A1 or A1B1) |A1:D7|

### Add/Delete Row
  
Add or Delete a Row
|Parameters|Description|example|
| --- | --- | --- |
|Option|Select Add or Delete a row|Add|
|Sheet Name|Name of the sheet where add the row|Sheet|
|Row Number|Indicate the row(s) to be added or deleted|2|
|Where to Insert|Indicate where add or delete the row|A1:D7|

### Add/Delete Column
  
Add or Delete a Column
|Parameters|Description|example|
| --- | --- | --- |
|Option|Select Add or Delete a column||
|Sheet Name|Name of the sheet where the data are located|Sheet|
|Range|Indicate the column(s) to be added or deleted|B|

### Convert CSV to XLSX
  
Convert a CSV document to XLSX format
|Parameters|Description|example|
| --- | --- | --- |
|CSV file path|Path of the csv file to be converted||
|Delimiter|Delimiter of the csv file||
|Do it have headers?|Check this checkbox if the csv has headers|True|
|Encoding|Type the encoding type of the file. Default is latin-1|utf-8|
|XLSX file path|Path of the xlsx file where will be saved|file.xlsx|

### Export to JSON
  
Export array data to JSON
|Parameters|Description|example|
| --- | --- | --- |
|Data|Data to be exported|[['header1', 'header2', 'header3', 'header4', 'header5', 'header6'], ['data11', 'data12', 'data13', 'data14', 'data15', 'data16']]|
|Json file path|Path of the json file where will be saved|C:/Users/User/Desktop/file.json|

### (Deprecated) Convert XLSX to CSV
  
Convert a xlsx document to csv
|Parameters|Description|example|
| --- | --- | --- |
|XLSX file path|Path of the xlsx file to be converted|C:/Users/User/Desktop/file.xlsx|
|Delimiter|Delimiter of the csv file|,|
|Sheet name|Name of the sheet where the data are located|Sheet0|
|CSV file path|Path of the xlsx file where will be saved|C:/Users/User/Desktop/file.csv|

### Convert XLSX to CSV
  
Convert a xlsx document to csv
|Parameters|Description|example|
| --- | --- | --- |
|XLSX file path|Path of the xlsx file to be converted|C:/Users/User/Desktop/file.xlsx|
|Delimiter|Delimiter of the csv file|,|
|Sheet name|Name of the sheet where the data are located|Sheet0|
|CSV file path|Path of the xlsx file where will be saved|C:/Users/User/Desktop/file.csv|

### Convert XLS to XLSX
  
Convert a xls document to xlsx
|Parameters|Description|example|
| --- | --- | --- |
|XLS file path|Path of the xls file to be converted|C:\Users\User\Desktop\file.xls|
|XLSX file path|Path of the xlsx file where will be saved|C:\Users\User\Desktop\new_file.xlsx|

### Get active cell
  
Get row and column of active cell
|Parameters|Description|example|
| --- | --- | --- |
|Assign result to variable |Name of the variable where the result will be stored|Variable|

### Refresh Pivot table
  
Refresh a pivot table. Deprecated! Use PivotTableExcel module
|Parameters|Description|example|
| --- | --- | --- |
|Sheet |Name of the sheet where the table is located|Sheet1|
|Pivote table name |Name of the pivot table to be updated|Name: |

### Fit cells
  
Adjusts, groups and ungroups a range of cells. You can group/ungroup by rows or columns
|Parameters|Description|example|
| --- | --- | --- |
|Sheet |Name of the sheet where the data are located|Sheet1|
|Range to fit|Cell or Range of cells to fit. The syntax must be the same as excel (A1 or A1B1) |A1:D7|
|Autofit|Automatically fit the cells to display the data||
|Group rows|To mark this checkbox will group the rows in the selected range.||
|Group columns|To mark this checkbox will group the columns in the selected range.||
|Ungroup rows|To mark this checkbox will upgroup the rows in the selected range.||
|Ungroup columns|To mark this checkbox will upgroup the columns in the selected range.||
|Merge cells|Checking this checkbox will merge the cells in the selected range||
|Unmerge cells|Checking this checkbox will unmerge the cells in the selected range||
|Row level|To mark this checkbox will displays the specified number of row levels of an outline|2|
|Column level|To mark this checkbox will displays the specified number of column levels of an outline|2|
|Column width|Width to which the column will fit|20|
|Row Height|Height to which the row will adjust|20|

### Get Formula
  
Get the formula into cell
|Parameters|Description|example|
| --- | --- | --- |
|Cell |Cell where is the formula. The syntax must be the same as excel (A1 or A1B1) |A5|
|Assign result to variable |Name of the variable where the result will be stored|Variable|

### Add Auto Filter
  
Add auto filter to excel table
|Parameters|Description|example|
| --- | --- | --- |
|Sheet |Name of the sheet where the data are located|Sheet1|
|Columns |Columna or Range of columns. The syntax must be the same as excel (A o AB) |A:E |

### Remove Auto Filter
  
Remove auto filter from an excel sheet
|Parameters|Description|example|
| --- | --- | --- |
|Sheet |Name of the sheet where the filter to remove is located|Sheet1|

### Clear Filter
  
Clears every filter applied over an excel sheet
|Parameters|Description|example|
| --- | --- | --- |
|Sheet |Name of the sheet where the filtered data is located|Sheet1|

### Filter
  
Filter an excel table according to the relative value, exact content, background color or font color of the cells. *Examples according the filter type: xlAnd ['>=10'] or ['>=10', '<=20'] | xlOr ['<=10', '>=20'] | xlFilterValues ['10','20', '30'] | xlFilterCellColor (255,0,0) | xlFilterFontColor (255,0,0)*
|Parameters|Description|example|
| --- | --- | --- |
|Sheet |Name of the sheet where the data are located|Sheet1|
|Table start |Column where the table to be filtered begins|A |
|Column |Column where to add the filter|A |
|Filter |Value or list of values, filter of unique criteria or list of two items for double criteria (eg value between A and B). Use "=" to find blank fields, "<>" for non-empty cells and data negation.|['>=10'] or ['>=10', '<=20'], ['10','20', '30'] or (255,0,0)|
|Filter type |Type of filter to apply.||

### Filter by Date
  
Filter a table by the day, month or year of a date indicated
|Parameters|Description|example|
| --- | --- | --- |
|Sheet |Name of the sheet where the data are located|Sheet1|
|Table start |Column where the table to be filtered begins|A |
|Column |Column where to add the filter|A |
|Date Filter |Date or list of dates to set as filter criteria|18/04/2024|
|Filter type |Type of filter to apply.||

### Advanced filter
  
Apply advanced filter to a table
|Parameters|Description|example|
| --- | --- | --- |
|Sheet |Name of the sheet where the data are located|Sheet1|
|Table range |Range of the table to be filtered begins|A1:G500 |
|Criteria range  |Range with the filter criteria to apply|A1:B4 |
|Unique records only|||
|Copy to another place|Paste the resulting table to the target cell||
|Target  |Cell where to paste the result table of the filter|J1 |

### Clear filters
  
Remove filters and show all data
|Parameters|Description|example|
| --- | --- | --- |
|Sheet |Name of the sheet where the data are located|Sheet1|

### Rename sheet
  
Change name to excel sheet
|Parameters|Description|example|
| --- | --- | --- |
|Sheet |Name of the sheet to be renamed|Sheet1|
|New name |New name of the sheet|new_name|

### Text Format
  
Change the Horizontal or Vertical alignment of values in a range of cells
|Parameters|Description|example|
| --- | --- | --- |
|Sheet |Name of the sheet where the data are located|Sheet1|
|Cell or Range of cells|Range that contains the data to align|A1:D7|
|Horizontal Alignment|Selector that contains the horizontal alignment options||
|Vertical Alignment|Selector that contains the vertical alignment options||

### Cell Style
  
This command modifies the formatting of the selected cell or range of cells. You can change the font and borders
|Parameters|Description|example|
| --- | --- | --- |
|Sheet Name |Name of the sheet to automate|Sheet1|
|Cell to format |Cell or Range of cells to format. The syntax must be the same as excel (A1 or A1B1) |A1:C4|
|Border|Border of the cell to be formatted|Contour|
|Style|Border style of the cell to be formatted|_ _ _ _ _ _ _ _ _ _ _|
|Font size |Font size of the cell to be formatted|20|
|Blod|Mark this checkbox to change the font to bold|True|
|Italic|Mark this checkbox to change the font to italic|True|
|Underline|Mark this checkbox to change the font to underline|True|
|Adjust Text||True|
|Horizontal Alignment|Type of horizontal alignment of the cell that you want to format|Alignment|
|Vertical Alignment|Type of vertical alignment of the cell that you want to format|Alignment|

### Paste in Cells
  
Paste data to cells in Excel
|Parameters|Description|example|
| --- | --- | --- |
|Sheet |Name of the sheet to automate|Hoja 1|
|cells where paste |Cell or Range of cells to paste. The syntax must be the same as excel (A1 or A1B1) |A1:B3|
|Only values |If this checkbox is checked, only the value will be pasted|True|

### Disable Cut/Copy Mode
  
Disable Cut/Copy Mode of the active Excel
|Parameters|Description|example|
| --- | --- | --- |

### Remove Duplicates
  
Execute the remove duplicates command of Excel
|Parameters|Description|example|
| --- | --- | --- |
|Sheet|Name of the sheet to automate|Hoja 1|
|Range|Range of cells. The syntax must be the same as excel (A1 or A1B1) |A1:B3|
|Column |Indicate the column where the duplicates are to be searched for|A / ['A', 'B'] |
|Do it have headers?|Check this checkbox if the excel has headers|True|

### Export to advanced PDF
  
Export to PDF with options
|Parameters|Description|example|
| --- | --- | --- |
|Save PDF|Path where to save the .pdf file|/Users/user/Desktop/excel.pdf|
|Sheet |Name of the sheet to export|Sheet1|
|Al sheets|By checking the checkbox, all sheets will be exported||
|Autofit|||
|Zoom|Adjust the zoom of the spreadsheet content.||
|FitToPagesTall|Adjust the height of the spreadsheet content to the defined number of pages.|1|
|FitToPagesWide|Adjust the width of the spreadsheet content to the defined number of pages.|1|
|Orientation|||

### Copy-Move Sheet
  
Copy or move a sheet
|Parameters|Description|example|
| --- | --- | --- |
|Origin Sheet|Name of the source sheet|Sheet1|
|Move/copy before sheet|Name of the sheet to be moved|Sheet2|
|Excel destination|Path of the .xlsx file where move or copy the sheet|C:/path/to/excel.xlsx|
|Password (optional) |Password of the xlsx file|P@ssW0rd|
|Copy |By checking the checkbox, the sheet will be copied||

### Insert Form
  
Insert Form in Excel
|Parameters|Description|example|
| --- | --- | --- |
|Macro path |Path of the frm file to be inserted|Form.frm|

### Read Filtered Cells
  
Read all the content of the filtered cells and apply formatting to date-type data if indicated
|Parameters|Description|example|
| --- | --- | --- |
|Sheet |Name of the sheet where the data are located|Sheet1|
|Range |Cell or Range of cells. The syntax must be the same as excel (A1 or A1B1) |A1:B100 |
|Assign result to variable |Name of the variable where the result will be stored|Variable|
|Add specific formatting to data stored as date |Give an specific format to data stored as date|%m/%d/%Y, %H:%M:%S|
|Rows|||
|More data |||

### Count Filtered Cells
  
Allow count only cells filters 
|Parameters|Description|example|
| --- | --- | --- |
|Sheet |Name of the sheet where the data are located|Sheet1|
|Range |Range of column filtered (A1A100)|A1:A100 |
|Assign result to variable |Name of the variable where the result will be stored|Variable|

### Replace
  
Run replace action to excel 
|Parameters|Description|example|
| --- | --- | --- |
|Sheet |Name of the sheet where the data are located|Sheet1|
|Range |Cell or Range of cells. The syntax must be the same as excel (A1 or A1B1) |A1:B100 |
|Word to replace |Word to be sought to be replaced|10/10/2020|
|New word |Word that will replace the previous word indicated|10-10-2020|

### Order
  
Run replace action to excel 
|Parameters|Description|example|
| --- | --- | --- |
|Sheet |Name of the sheet where the data are located|Sheet 1|
|Range |Cell or Range of cells. The syntax must be the same as excel (A1 or A1B1) |A1:B100 |
|Column|Indicate the column to be sorted|A1:A22|
|Order type |Indicate how the column will be sorted|Ascending|

### Order by multiple levels
  
Order an excel sheet by value, setting multiple levels
|Parameters|Description|example|
| --- | --- | --- |
|Sheet |Name of the sheet where the data are located|Sheet 1|
|Range to sort |Cell or Range of cells. The syntax must be the same as excel (A1 or A1B1) |A1:B100 |
|It has headers|If this option is checked, it will take the first row of the range as headers.||
|Sort fields|||

### Refresh All
  
Refresh all data in Excel
|Parameters|Description|example|
| --- | --- | --- |

### Find
  
Searches a text in the given range and returns the address of the cell of the first occurence. If a value is not found, it will return empty. If the range it is filtered, the search will be performed over the visible cells
|Parameters|Description|example|
| --- | --- | --- |
|Sheet |Name of the sheet where the data are located|Sheet1|
|Range |Cell or Range of cells. The syntax must be the same as excel (A1 or A1B1) |A1:B100 |
|Text to find|Text to be searched in the excel|Lorem|
|Look at (optional)|Indicates the desired type of match the whole searched text or within any part (ba default within any part). ||
|Look in (optional)|Indicates where to do the searching value, formula or comment (default value). ||
|Case sensitive|If this box is checked, it will search the text string case-sensitive.||
|Find all|If this box is checked, it will return a list with every coincidence.||
|Assign result to variable |Name of the variable where the result will be stored|Variable|

### Find data
  
Returns the first cell that matches the search data
|Parameters|Description|example|
| --- | --- | --- |
|Sheet |Name of the sheet where the data are located|Sheet1|
|Range |Cell or Range of cells. The syntax must be the same as excel (A1 or A1B1)|A1:B100 |
|Letter of column with dates (Optional)|Letter of the column/s that contain dates.|A,B|
|Date Format (Optional)|Date format to search for.|%d/%m/%Y|
|Text to find|Text to be searched in the excel|Lorem|
|Not case sensitive|If this box is checked, it will search for the text string without differences between upper and lower letters.||
|Assign result to variable |Name of the variable where the result will be stored|Variable|

### Lock Cells
  
Lock or Unlock cells
|Parameters|Description|example|
| --- | --- | --- |
|Sheet |Name of the sheet where the data are located|Sheet1|
|Range |Cell or Range of cells. The syntax must be the same as excel (A1 or A1B1) |A1:B100 |
|Action|Select if you want to lock or unlock a cell|Lock|

### Add Chart
  
Create a new chart in an excel sheet
|Parameters|Description|example|
| --- | --- | --- |
|Sheet |Name of the sheet where the data are located|Sheet1|
|Type of Chart|Select the type of chart to be inserted in excel|Line|
|Cell position |Cell where will be inserted the chart. The syntax must be the same as excel (A1) |A1|
|Data range |Cell or Range of cells. The syntax must be the same as excel (A1 or A1B1) |Sheet!A1:B100 |

### Remove Password
  
Remove password and save the Excel
|Parameters|Description|example|
| --- | --- | --- |
|Excel with password|Path of the xlsx file to be opened|C:/Users/User/Desktop/test.xlsx|
|Password|Password of the xlsx file|****|
|Excel without password|Path where to save the .xlsx file. Empty for save in the same Excel|C:/Users/User/Desktop/test2.xlsx|

### Insert image
  
Insert an image
|Parameters|Description|example|
| --- | --- | --- |
|Sheet |Name of the sheet where the data are located|Sheet1|
|Cell |Cell where will be inserted the image. The syntax must be the same as excel (A1) |B5|
|Image path |Path of the image file to be inserted|image.png|

### Export Chart
  
Export a chart from index
|Parameters|Description|example|
| --- | --- | --- |
|Sheet |Name of the sheet where the data are located|Sheet1|
|Index |Index of the chart to be exported|1|
|Image path |Path where the image will be saved|/path/to/image.png|

### Not visible mode
  
Open not visible excel.
|Parameters|Description|example|
| --- | --- | --- |
|path to XLSX file |Path of the xlsx file to be opened|File.XLSX|
|Id (optional) |Name or identifier for the file to be opened. Used when you need to open more than one excel. By default is *default*.|default|

### Write array objects
  
Write array object on Excel cells.
|Parameters|Description|example|
| --- | --- | --- |
|Sheet |Name of the sheet where the data are located|Sheet1|
|Cell o Cells Range |Cell or Range of cells. The syntax must be the same as excel (A1 or A1B1) |A1|
|Data |Cell or Range of cells. The syntax must be the same as excel (A1 or A1B1) |[{ 'id',: 1, 'text': 'hello' },{ 'id',: 2, 'text': 'world' }]|

### Copy-Paste Format
  
Copy format range cell to another sheet
|Parameters|Description|example|
| --- | --- | --- |
|Origin Sheet|Name of the source sheet|Sheet1|
|Range to copy||A1:C4|
|Destiny Sheet|Name of the destiny sheet|Sheet2|
|Range to paste||A1:C4|

### Update links
  
Changes a link from one document to another
|Parameters|Description|example|
| --- | --- | --- |
|Path to change|Path of the xlsx file to be updated||
|Updated path|Path of the xlsx file that will replace the link|file.xlsx|

### Unlock book
  
Unlock book with password
|Parameters|Description|example|
| --- | --- | --- |
|Password|Password of the locked sheet|Password|

### Lock book
  
Lock a book with password
|Parameters|Description|example|
| --- | --- | --- |
|Password|Password to lock the book|Password|

### Unlock sheet
  
Unlock sheet with password
|Parameters|Description|example|
| --- | --- | --- |
|Sheet |Name of the sheet to be unlocked|Sheet1|
|Password|Password of the locked sheet|Password|

### Lock sheet
  
Lock a sheet with password
|Parameters|Description|example|
| --- | --- | --- |
|Sheet |Name of the sheet to be locked|Sheet1|
|Password|Password to lock the sheet|Password|

### Convert to .txt
  
Convert to .txt
|Parameters|Description|example|
| --- | --- | --- |
|path to XLSX file |Path of the xlsx file to be converted|file.XLSX|
|Save TXT|Path where to save the .txt file|/Users/user/Desktop/test.txt|

### Text to columns
  
Parses a column of cells that contain text into several columns.
|Parameters|Description|example|
| --- | --- | --- |
|Sheet |Name of the sheet where the data are located|Sheet1|
|Range |Cell or Range of cells. The syntax must be the same as excel (A1 or A1B1) |A1:B100 |
|Select separator |Select the cell separator, it can be fixed width or delimited||
|Select delimiter type |Select the delimiter type||
|Other delimiter or widths |Write the delimiter or fixed width|| or 20,35,22,10|

### Convert Excel time to hours
  
Convert Excel time to hours. Returns the format as hh: mm: ss
|Parameters|Description|example|
| --- | --- | --- |
|Decimal time ||0.296655812|
|Assign result to variable |Name of the variable where the result will be stored|Variable|

### Print sheet
  
Prints a sheet
|Parameters|Description|example|
| --- | --- | --- |
|Sheet |Name of the sheet to be printed|Sheet1|

### Save Excel with password
  
Save a Excel file
|Parameters|Description|example|
| --- | --- | --- |
|Save Excel as|Path where to save the .xlsx file|/Users/user/Desktop/excel.xlsx|
|Password to save with|Password of the xlsx file|password|

### Save Excel
  
Save an Excel file (as '.xlsx', 'xlsm', '.xls' or '.csv') in the indicated path
|Parameters|Description|example|
| --- | --- | --- |
|Save Excel|Path where to save the .xlsx file|/Users/user/Desktop/excel.xlsx|

### Close XLSX
  
Close the workbook opened by Rocketbot
|Parameters|Description|example|
| --- | --- | --- |
|Kill process |If this box is marked, the process will be completely closed.||
