



# Excel Advanced Options
  
Module with advanced options for Excel  

## How to install this module
  
__Download__ and __install__ the content in 'modules' folder in Rocketbot path  


## How to use
To use this module you must have Microsoft Excel.

## Overview

1. Open Without Alerts  
Open a file preventing MS Excel alerts.

2. Find and Connect  
Search a Excel Book opened and connect it

3. Count columns  
Count the columns or return the last column name. It's necessary that the excel is saved to get the last changes

4. Count Rows  
Counts all the rows or from a range.

5. Cell color  
Change color of a cell or range of cells. Can be a default color or custom 

6. Get Cell colors  
Get the colors of a cell.

7. Insert Formula  
Insert formula into cell

8. Insert Macro  
Insert Macro in Excel

9. Select Cells  
Select cells in Excel

10. Get Cell With Currency Format  
Get cells with currency format

11. Get Cell With Date Format  
Get cells with date format

12. Copy-Paste  
Copy range cell to another sheet

13. Format Cell  
Format Cell

14. Create Sheet  
Create sheet in the end

15. Delete Sheet  
Delete sheet

16. Copy to another excel  
Copy range to another Excel in the background

17. Add/Delete Row  
Add or Delete a Row

18. Add/Delete Column  
Add or Delete a Column

19. Convert CSV to XLSX  
Convert a csv document to xlsx

20. Convert XLSX to CSV  
Convert a xlsx document to csv

21. Convert XLS to XLSX  
Convert a xls document to xlsx

22. Get active cell  
Get row and column of active cell

23. Refresh Pivot table  
Refresh a pivot table. Deprecated! Use PivotTableExcel module

24. Fit cells  
Adjusts, groups and ungroups a range of cells. You can group/ungroup by rows or columns

25. Get Formula  
Get the formula into cell

26. Add Auto Filter  
Add auto filter to excel table

27. Filter  
Add filter to excel table

28. Rename sheet  
Change name to excel sheet

29. Text Format  
Change the Horizontal or Vertical alignment of values in a range of cells

30. Cell Style  
This command modifies the formatting of the selected cell or range of cells. You can change the font and borders

31. Paste in Cells  
Paste data to cells in Excel

32. Remove Duplicates  
Execute the remove duplicates command of Excel

33.  Export to advanced PDF  
Export to PDF with options

34. Copy-Move Sheet  
Copy or move a sheet

35. Insert Form  
Insert Form in Excel

36. Read Filtered Cells  
Allow read only cells filters 

37. Count Filtered Cells  
Allow count only cells filters 

38. Replace  
Run replace action to excel 

39. Order  
Run replace action to excel 

40. Refresh All  
Refresh all data in Excel

41. Find  
Return de first found cell 

42. Lock Cells  
Lock or Unlock cells

43. Add Chart  
Create a new chart in an excel sheet

44. Remove Password  
Remove password and save the Excel

45. Insert image  
Insert an image

46. Export Chart  
Export a chart from index

47. Not visible mode  
Open not visible excel.

48. Write array objects  
Write array object on Excel cells.

49. Copy-Paste Format  
Copy format range cell to another sheet

50. Update links  
Changes a link from one document to another

51. Unlock sheet  
Unlock sheet

52. Convert to .txt  
Convert to .txt

53. Text to columns  
Parses a column of cells that contain text into several columns.

54. Convert Excel time to hours  
Convert Excel time to hours. Returns the format as hh: mm: ss

55. Print sheet  
Prints a sheet

56. Save Excel with password  
Save a Excel file

57. Save Excel  
Save a Excel file in the indicated path

58. Close XLSX  
Close the workbook opened by Rocketbot  



### Changes
#### 09-Aug-2022
- Add "Special Paste" options to Copy-Paste and add Get Cell Colors command
#### 22-Jul-2022
- Fix Copy to another excel and Copy-Paste Format
#### 13-Jun-2022
- Format text: Added command to change text alignment
#### 12-May-2022
- Copy to another excel: fixed command to copy from one excel to another
#### 18-Apr-2022
- Text to Column: command fixed to separate text in columns
#### 06-Apr-2022
- Fit Cells: Added merge cells, adjust rows, adjust columns functions
#### 28-Dec-2021
- Count Rows: command fixed to count all rows.
#### 9-Nov-2021
- Order command: Apply multiple orders and clean filters.
#### 13-Oct-2021
- Fix count cells filtered
#### 30-Sep-2021
- Paste command: Update compatibilities
#### 28-Sep-2021
- Fix get filtered cells command. Now returns extended data
#### 06-Jul-2021
- Fix language
#### 01-Jul-2021
- Read Filtered Cells: The command was fixed because it didn't getting all cell range
#### 27-Apr-2021
- Texto to column: Parses a column of cells that contain text into several columns.
#### 18-Mar-2021
- Unlock sheet: Convert XLSX to TXT.
#### 09-Mar-2021
- Unlock sheet: Unlock a sheet by password.
#### 09-Mar-2021
- Update links: Changes a link from one document to another
#### 17-Feb-2021
- Find and Connect: Find opened Excel file and connect it
#### 1-Feb-2021
- Add command Copy-Paste Format. You can copy format cell to another.
#### 25-Jan-2021
- Write array objects: Writes information obtained from an array of objects to excel cells
#### 21-Jan-2021
- Not visible mode: Open background Excel
#### 1-Dec-2020
- Export chart: Export a chart from index.
#### 24-Nov-2020
- Insert image in a cell.
#### 24-Sep-2020
- Open without alerts: Add field 'Password'
#### 16-Sep-2020
- Add chart: Create a new chart on excel sheet 
#### 15-Sep-2020
- Lock Cells: Lock or unlock cells 
#### 2-Sep-2020
- Find: Replicate Excel Find command 
#### 31-Jul-2020
- Order: Replicate Excel Order command 
#### 15-Jul-2020
- Read Filtered Cells: Read cell after execute Filter command
- Replace: Replicate Excel Replace command 
#### 2-Jul-2020
- Insert Form: Rocketbot can insert VBA Form to Excel
#### 30-Jun-2020
- Csv to xlsx: Checkbox header was added to decide if the csv has a header
- Export to Advanced PDF: Rocketbot export to PDF command enhancement
- Copy-Move Sheet: Replicate move/copy sheet command of Excel
#### 17-Jun-2020
- Remove duplicates: Rocketbot can now remove duplicate data on range Excel
#### 5-Jun-2020
- Focus Excel: Rocketbot can now set Excel to the foreground window

----
### OS

- windows
- mac

### Dependencies
- [**xlwings**](https://pypi.org/project/xlwings/)- [**pandas**](https://pypi.org/project/pandas/)
### License
  
![MIT](https://camo.githubusercontent.com/107590fac8cbd65071396bb4d04040f76cde5bde/687474703a2f2f696d672e736869656c64732e696f2f3a6c6963656e73652d6d69742d626c75652e7376673f7374796c653d666c61742d737175617265)  
[MIT](http://opensource.org/licenses/mit-license.ph)
