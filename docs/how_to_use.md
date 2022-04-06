## How to use
To use this module, you must have Microsoft Excel.


## Descripción de los comandos

### Abrir sin alertas
  
Open a file without displaying alert banners.
|Parameters|Description|example|
| --- | --- | --- |
|XLSX file path|Path of the xlsx file to be opened|File.XLSX|
|Password (optional)|Password of the xlsx file|P@ssW0rd|
|Identifier (optional)|Name or identifier for the file to open. It is used when you need to open more than one excel. Default is *default*|id|
|Assign result to variable|Variable where the result will be stored|id|

### Count Columns
  
Count the number of columns in the open excel. Excel is required to be saved to take the latest changes
|Parameters|Description|example|
| --- | --- | --- |
|Sheet|Name of the sheet where the data is located|Sheet 1|
|Get Column Name|If this box is checked, it will return the letter of the last column|True|
|Assign result to variable|Name of the variable where to store the result|number_columns|

### Count Rows
  
Count all rows or within a range.
|Parameters|Description|example|
| --- | --- | --- |
|Sheet|Name of the sheet where the data is located|Sheet 1|
|Count all rows|Option to count all rows.||
|Column|Column where the number of rows will be counted|C|
|Assign result to variable|Name of the variable where to save the result|number_rows|

### Cell color
  
Change color of a cell or range of cells. You can select a default value or a custom one
|Parameters|Description|example|
| --- | --- | --- |
|Enter cells |Cell or Range of cells. The syntax must be the same as excel (A1 or A1B1) |A1:B5|
|Enter color in RGB |Rgb values ​​of the color that the cell(s) will have|250,250,250|
|Select Color |Select the color. You can use the above field to customize|network|

### Insert Formula
  
Insert formula on a cell
|Parameters|Description|example|
| --- | --- | --- |
|Enter cell |Cell or Range of cells. The syntax must be the same as excel (A1 or A1B1) |A5|
|Type formula |Formula you want to insert. It must be written in English. Remember to use *,* to separate parameters|=SUM(A1:A4)|

### Insert Macro to Excel
  
Insert a Macro to Excel
|Parameters|Description|example|
| --- | --- | --- |
|Macro Path|Path of the .bas file to be inserted|Macro.bas|

### Select Cells
  
Select cells in Excel
|Parameters|Description|example|
| --- | --- | --- |
|Sheet|Name of the sheet to be automated|Sheet 1|
|Enter cells to select|Cell or Range of cells to select. The syntax must be the same as excel (A1 or A1B1) |A1:B3|
|Copy|Checking the box will copy the values ​​to the clipboard|True|

### Get Currency Format Cell
  
Get cells with currency format
|Parameters|Description|example|
| --- | --- | --- |
|Sheet|Name of the sheet to be automated|Sheet 1|
|Enter cells to select|Cell or Range of cells. The syntax must be the same as excel (A1 or A1B1) |A1:B3|
|Assign to variable|Name of the variable where to store the result|variable|

### Get Cell Date Format
  
Get cells with date format
|Parameters|Description|example|
| --- | --- | --- |
|Sheet|Name of the sheet to be automated|Sheet 1|
|Enter cells to select|Cell or Range of cells. The syntax must be the same as excel (A1 or A1B1) |A1:B3|
|Assign to variable|Name of the variable where to store the result|variable|

### Copy paste
  
Copy a range of cells from one sheet to another
|Parameters|Description|example|
| --- | --- | --- |
|Source sheet |Name of the sheet to be automated|Sheet1|
|Range to copy |Cell or Range of cells to copy. The syntax must be the same as excel (A1 or A1B1) |A1:C4|
|Target sheet |Target sheet name|Sheet2|
|Range to paste|Cell or Range of cells to paste. The syntax must be the same as excel (A1 or A1B1) |A1:C4|

### Format Cell
  
Format Cell
|Parameters|Description|example|
| --- | --- | --- |
|Sheet Name|Name of the sheet to be automated|Sheet1|
|Range to format |Cell or Range of cells to format. The syntax must be the same as excel (A1 or A1B1) |A1:C4|
|Format|The format type for the cell must be selected. Select custom to add a custom format|dd-mm-yy|
|Custom format |Custom format. It should be the same as shown in the custom section of Excel|00000|

### Create Sheet
  
Add a leaf at the end
|Parameters|Description|example|
| --- | --- | --- |
|Sheet name|Name of the sheet to be created|Sheet2|
|After|The sheet will be created next to the sheet indicated in this field|Sheet1|

### Delete Sheet
  
Delete a sheet
|Parameters|Description|example|
| --- | --- | --- |
|Sheet name|Name of the sheet to be deleted|Sheet2|
|Assign result to variable|Name of the variable where to store the result|Variable|

### Copy from one Excel to another
  
Copy a range from one Excel to another, the destination Excel must not be open
|Parameters|Description|example|
| --- | --- | --- |
|Source excel|Source excel file path|Sheet1|
|Source sheet|Source sheet name|Sheet1|
|Range to copy|Cell or Range of cells to copy. The syntax must be the same as excel (A1 or A1B1) |A1:D7|
|Destination excel|Destination excel file path|Sheet1|
|Destination sheet|Name of the sheet where it will be copied|Sheet1|
|Range to paste|Cell or Range of cells to copy. The syntax must be the same as excel (A1 or A1B1) |A1:D7|
|Values ​​Only|If this box is checked, it will copy only the values ​​|True|

### Insert/Delete Row
  
Insert or delete a row
|Parameters|Description|example|
| --- | --- | --- |
|Option|Select Add to add a row or Delete to delete|Add|
|Sheet Name|Name of the sheet where to add the row|Sheet|
|Row Number|Indicate the row or rows you want to add or delete|2|
|Where to Insert|Indicate where to add or delete the row|A1:D7|

### Insert/Delete Column
  
Insert or delete a column
|Parameters|Description|example|
| --- | --- | --- |
|Option|Select Add to add a column or Delete to delete||
|Sheet Name|Name of the sheet where the data is located|Sheet|
|Column|Indicate the column or columns that you want to add or delete|B|

### Convert CSV to XLSX
  
Convert a CSV document to XLSX
|Parameters|Description|example|
| --- | --- | --- |
|CSV file path|Path of the csv file to be converted||
|Delimiter|Csv file separator||
|Does it have headers?|Check this box if the csv has headers|True|
|Encoding|Type the encoding type of the file. Default is latin-1|latin-1|
|XLSX file path|Path of the xlsx file where to save|file.xlsx|

### Convert XLSX to CSV
  
Convert an XLSX document to CSV
|Parameters|Description|example|
| --- | --- | --- |
|XLSX file path|Path of the xlsx file to be converted|C:/Users/User/Desktop/file.xlsx|
|Delimiter|Csv file separator|,|
|Sheet name|Name of the sheet where the data is located|Sheet0|
|CSV file path|Path of the csv file where to save the conversion|C:/Users/User/Desktop/file.csv|

### Convert XLS to XLSX
  
Convert an XLS document to XLSX
|Parameters|Description|example|
| --- | --- | --- |
|XLS file path|Path of the xls file to convert|C:\Users\User\Desktop\file.xls|
|XLSX file path|Path where the xlsx file will be saved|C:\Users\User\Desktop\new_file.xlsx|

### Get active cell
  
Get row and column of active cell
|Parameters|Description|example|
| --- | --- | --- |
|Assign result to variable|Name of the variable where to store the result|Variable|

### Update pivot table
  
Update a pivot table. Obsolete! Use the PivotTableExcel module
|Parameters|Description|example|
| --- | --- | --- |
|Sheet |Name of the sheet where the table is located|Sheet 1|
|PivotTable Name |PivotTable Name to be updated|Name: |

### Wrap cells
  
Adjust, join, group and ungroup a range of cells. You can group/ungroup by rows or columns
|Parameters|Description|example|
| --- | --- | --- |
|Sheet|Name of the sheet where the data is located|Sheet 1|
|Range to fit|Cell or Range of cells to fit. The syntax must be the same as excel (A1 or A1B1) |A1:D7|
|Autofit|Automatically fits the cells to display the data||
|Group Rows|Checking this checkbox will group the rows in the selected range||
|Group columns|By checking this checkbox, the columns in the selected range will be grouped||
|Ungroup Rows|Checking this checkbox will ungroup the rows in the selected range||
|Ungroup columns|Checking this checkbox will ungroup the columns in the selected range||
|Merge cells|Checking this checkbox will merge the cells in the selected range||
|Row Level|Checking this box will display the specified number of row levels|2|
|Column Range|Checking this box will display the specified number of column levels|2|
|Column width|Width to which the column will fit|20|
|Row Height|Height the row will wrap to|20|

### Get Formula
  
Get the formula on a cell
|Parameters|Description|example|
| --- | --- | --- |
|Enter cell |Cell where the formula is. The syntax must be the same as excel (A1 or A1B1) |A5|
|Assign result to variable|Name of the variable where to store the result|Variable|

### Add Automatic Filter
  
Add automatic filter to an excel table
|Parameters|Description|example|
| --- | --- | --- |
|Sheet |Name of the sheet where the data is located|Sheet 1|
|Range |Cell or Range of cells. The syntax must be the same as excel (A1 or A1B1) |A1:E6 |

### Filter
  
Filter to an excel table
|Parameters|Description|example|
| --- | --- | --- |
|Sheet |Name of the sheet where the data is located|Sheet 1|
|Start of table |Column where the table to be filtered begins|A |
|Column |Column where to add the filter|A |
|Filter |Filter or list of filters to add. Use "=" to find blank fields, "<>" for non-blank cells, and data negation|['filter1','filter2', 'filter3']|

### Rename sheet
  
Rename an excel sheet
|Parameters|Description|example|
| --- | --- | --- |
|Sheet |Name of sheet to rename|Sheet 1|
|New Name |New Sheet Name|new_name|

### Cell Style
  
This command modifies the format of the selected cell or range of cells. You can change the font and borders
|Parameters|Description|example|
| --- | --- | --- |
|Sheet Name|Name of the sheet to be automated|Sheet1|
|Range to format |Cell or Range of cells to format. The syntax must be the same as excel (A1 or A1B1) |A1:C4|
|Border|Border of the cell to be formatted|Contour|
|Style|Border style of the cell to format|_ _ _ _ _ _ _ _ _ _ _|
|Font size |Cell font size|20|
|Bold|Check this box to change the font to bold|True|
|Italic|Check this box to change the font to italic|True|
|Underline|Check this box to change the font to underline|True|
|Wrap Text|Check this box to wrap text within the specified range|True|

### Paste into Cells
  
Paste data into cells in Excel
|Parameters|Description|example|
| --- | --- | --- |
|Sheet|Name of the sheet to be automated|Sheet 1|
|Enter cells to paste|Cell or Range of cells to paste. The syntax must be the same as excel (A1 or A1B1) |A1:B3|
|Only values|If this box is checked, only values|True|

### Remove duplicates
  
Run the remove duplicates command in Excel
|Parameters|Description|example|
| --- | --- | --- |
|Sheet|Name of the sheet to be automated|Sheet 1|
|Enter cells to filter|Cell or Range of cells. The syntax must be the same as excel (A1 or A1B1) |A1:B3|
|Column |Indicate the column where the duplicates will be searched|A |
|Does it have headers?|Check this box if the excel has headers|True|

### Close XLSX
  
Close the open book by Rocketbot
|Parameters|Description|example|
| --- | --- | --- |

### Save Excel
  
Save an Excel file in the indicated path
|Parameters|Description|example|
| --- | --- | --- |
|Save Excel|Path where to save the .xlsx file|/Users/user/Desktop/excel.xlsx|

### Save excel with password
  
Save an Excel file
|Parameters|Description|example|
| --- | --- | --- |
|Save Excel to|Path where to save the .xlsx file|/Users/user/Desktop/excel.xlsx|
|Enter the password|Password of the xlsx file|password|

### Export to advanced PDF
  
Export Excel to PDF with options
|Parameters|Description|example|
| --- | --- | --- |
|Save PDF|Path where to save the .pdf file|/Users/user/Desktop/excel.pdf|
|Auto Adjustment|||
|Zoom|||
|Adjust Height|||
|Fit width|||

### Copy-Move Sheet
  
Copy or move a sheet
|Parameters|Description|example|
| --- | --- | --- |
|Source sheet |Source sheet name|Sheet1|
|Move/copy before sheet |Name of sheet to move to|Sheet2|
|Destination Excel|Path of the .xlsx file where to move or copy the sheet|C:/path/to/excel.xlsx|
|Copy|Checking the box will create a copy of the sheet||

### Insert form
  
Insert a Form to Excel
|Parameters|Description|example|
| --- | --- | --- |
|Form Path|Path of the frm file to be inserted|Form.frm|

### Read filtered cells
  
Read only filtered cells
|Parameters|Description|example|
| --- | --- | --- |
|Sheet |Name of the sheet where the data is located|Sheet 1|
|Range to search for |Cell or Range of cells. The syntax must be the same as excel (A1 or A1B1) |A1:B100 |
|Assign result to variable|Name of the variable where to store the result|Variable|
|Extra data|||

### Count filtered cells
  
Count only filtered cells
|Parameters|Description|example|
| --- | --- | --- |
|Sheet |Name of the sheet where the data is located|Sheet 1|
|Range to search for |Filtered column range (A1A100)|A1:A100 |
|Assign result to variable|Name of the variable where to store the result|Variable|

### Replace
  
Execute the replace option of excel
|Parameters|Description|example|
| --- | --- | --- |
|Sheet |Name of the sheet where the data is located|Sheet 1|
|Range to search for |Cell or Range of cells. The syntax must be the same as excel (A1 or A1B1) |A1:B100 |
|Word to replace|Word to search for to be replaced|10/10/2020|
|New word|Word that will replace the previous one indicated|10-10-2020|

### Order
  
Execute the replace option of excel
|Parameters|Description|example|
| --- | --- | --- |
|Sheet |Name of the sheet where the data is located|Sheet 1|
|Range to search for |Cell or Range of cells. The syntax must be the same as excel (A1 or A1B1) |A1:B100 |
|Column|Indicate the column to sort|A1:A22|
|Type of order |Indicate how the column will be sorted|Ascending|

### Update all
  
Update all book fonts
|Parameters|Description|example|
| --- | --- | --- |

### Look for
  
Returns the first cell found
|Parameters|Description|example|
| --- | --- | --- |
|Sheet |Name of the sheet where the data is located|Sheet 1|
|Range to search for |Cell or Range of cells. The syntax must be the same as excel (A1 or A1B1) |A1:B100 |
|Text to search for|Text to search for in excel|Lorem|
|Assign result to variable|Name of the variable where to store the result|Variable|

### Lock cells
  
Lock or unlock cells
|Parameters|Description|example|
| --- | --- | --- |
|Sheet |Name of the sheet where the data is located|Sheet 1|
|Range to search for |Cell or Range of cells. The syntax must be the same as excel (A1 or A1B1) |A1:B100 |
|Action|Select whether you want to lock or unlock a cell |Lock|

### Add Chart
  
Add a new chart on a sheet in excel
|Parameters|Description|example|
| --- | --- | --- |
|Sheet |Name of the sheet where the data is located|Sheet 1|
|Chart Type|Select the type of chart to be inserted in the excel|Line|
|Cell where to insert the graphic |Cell where to insert the graphic. The syntax should be the same as excel (A1) |A1|
|Data range |Cell or Range of cells. The syntax must be the same as excel (A1 or A1B1) |A1:B100 |

### Remove Password
  
Remove the password and save the Excel
|Parameters|Description|example|
| --- | --- | --- |
|Excel with Password|Path of the xlsx file to open|C:/Users/User/Desktop/test.xlsx|
|Password|Password of the xlsx file|****|
|Excel without Password|Path where to save the .xlsx file. Empty to save to the same Excel|C:/Users/User/Desktop/test2.xlsx|

### Insert image
  
insert an image
|Parameters|Description|example|
| --- | --- | --- |
|Sheet |Name of the sheet where the data is located|Sheet 1|
|Cell |Cell where to insert the image. The syntax must be the same as excel (A1) |B5|
|Image path|Path of the image to be inserted|image.png|

### Export chart
  
Export a chart by index
|Parameters|Description|example|
| --- | --- | --- |
|Sheet |Name of the sheet where the data is located|Sheet 1|
|Index |Index of the chart to export|1|
|Image path|Path where the image will be saved|/path/to/image.png|

### Invisible mode
  
Open excel in invisible mode
|Parameters|Description|example|
| --- | --- | --- |
|XLSX file path|Path of the xlsx file to be opened|File.XLSX|
|Identifier (optional)|Name or identifier for the file to open. It is used when you need to open more than one excel. Default is *default*|default|

### Write array of objects
  
Write an array of objects to Excel cells
|Parameters|Description|example|
| --- | --- | --- |
|Sheet |Name of the sheet where the data is located|Sheet 1|
|Cell or Cell Range|Cell or Cell Range. The syntax must be the same as excel (A1 or A1B1) |A1|
|Data to write|Cell or Range of cells. The syntax should be the same as excel (A1 or A1B1) |[{ 'id',: 1, 'text': 'hello' },{ 'id',: 2, 'text': 'world' }]|

### Copy-Paste Format
  
Format a range of cells from one sheet to another
|Parameters|Description|example|
| --- | --- | --- |
|Source sheet |Source sheet name|Sheet1|
|Range to copy ||A1:C4|
|Target sheet |Target sheet name|Sheet2|
|Range where to paste||A1:C4|

### Search and connect
  
Find an open excel and connect to it.
|Parameters|Description|example|
| --- | --- | --- |
|Name of opened XLSX file||File.XLSX|
|Identifier (optional)|Name or identifier for the file to open. It is used when you need to open more than one excel. Default is *default*|excel1|

### Update links
  
Change a link from one document to another
|Parameters|Description|example|
| --- | --- | --- |
|Path to change|Path of the xlsx file to update||
|Updated path|Path of the xlsx file that will replace the link|file.xlsx|

### Unlock sheet
  
Unlock a sheet with password
|Parameters|Description|example|
| --- | --- | --- |
|Sheet|Name of the sheet to be locked|Sheet 1|
|Password|Lock Sheet Password|Password|

### Convert to .txt
  
Convert to .txt
|Parameters|Description|example|
| --- | --- | --- |
|XLSX file path|Path of the xlsx file to be converted|File.XLSX|
|Save TXT|Path where to save the .txt file|/Users/user/Desktop/test.txt|

### Text in column
  
Execute the option text in column of excel
|Parameters|Description|example|
| --- | --- | --- |
|Sheet |Name of the sheet where the data is located|Sheet 1|
|Range to search for |Cell or Range of cells. The syntax must be the same as excel (A1 or A1B1) |A1:B100 |
|Select color |||
|Other delimiter||,|

### Convert Excel time to hours
  
Convert Excel time to hours. Returns the result as hh:mm:ss
|Parameters|Description|example|
| --- | --- | --- |
|Enter the time in decimal format ||0.296655812|
|Assign result to variable|Name of the variable where to store the result|Variable|

### Print sheet
  
print a sheet
|Parameters|Description|example|
| --- | --- | --- |
|Sheet |Name of the sheet to be printed|Sheet 1|

---




