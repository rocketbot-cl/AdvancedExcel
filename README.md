



# Excel Advanced Options
  
Módulo con opciones avanzadas para Excel  

## How to install this module
  
__Download__ and __install__ the content in 'modules' folder in Rocketbot path  


## How to use this module
In order to use this module, you need to have Microsoft Excel.


## Overview


1. Open Without Alerts  
Open a file preventing MS Excel alerts.

2. Count columns  
Count the columns or return the last column name. It's necessary that the excel is saved to get the last changes

3. Count Rows  
Count Rows

4. Cell color  
Change color of a cell or range of cells. Can be a default color or custom 

5. Insert Formula  
Insert formula into cell

6. Insert Macro  
Insert Macro in Excel

7. Select Cells  
Select cells in Excel

8. Get Cell With Currency Format  
Get cells with currency format

9. Get Cell With Date Format  
Get cells with date format

10. Copy-Paste  
Copy range cell to another sheet

11. Format Cell  
Format Cell

12. Create Sheet  
Create sheet in the end

13. Delete Sheet  
Delete sheet

14. Copy to another excel  
Copy range to another Excel in the background

15. Add/Delete Row  
Add or Delete a Row

16. Add/Delete Column  
Add or Delete a Column

17. Convert CSV to XLSX  
Convert a csv document to xlsx

18. Convert XLSX to CSV  
Convert a xlsx document to csv

19. Convert XLS to XLSX  
Convert a xls document to xlsx

20. Get active cell  
Get row and column of active cell

21. Refresh Pivot table  
Refresh a pivot table. Deprecated! Use PivotTableExcel module

22. Fit cells  
Adjusts, groups and ungroups a range of cells. You can group/ungroup by rows or columns

23. Get Formula  
Get the formula into cell

24. Add Auto Filter  
Add auto filter to excel table

25. Filter  
Add filter to excel table

26. Rename sheet  
Change name to excel sheet

27. Cell Style  
This command modifies the formatting of the selected cell or range of cells. You can change the font and borders

28. Paste in Cells  
Paste data to cells in Excel

29. Remove Duplicates  
Removes duplicate data in range

30. Close XLSX  
Close the workbook opened by Rocketbot

31. Save Excel  
Save a Excel file in the indicated path

32. Save Excel with password  
Save a Excel file

33. Export to advanced PDF  
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

50. Find and Connect  
Search a Excel Book opened and connect it

51. Update links  
Changes a link from one document to another

52. Unlock sheet  
Unlock sheet

53. Convert to .txt  
Convert to .txt

54. Text to columns  
Parses a column of cells that contain text into several columns.

55. Convert Excel time to hours  
Convert Excel time to hours. Returns the format as hh: mm: ss

56. Print sheet  
Prints a sheet

57. Print sheet  
Prints a sheet  



### Updates
#### 9-Nov-2021
- Order command: Apply multiple orders and clean filters.
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