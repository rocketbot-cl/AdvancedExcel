
# Excel Advanced Options
  
Module with advanced options to automate Microsoft Excel

## How to install this module
  
__Download__ and __install__ the content in 'modules' folder in the Rocketbot path
## How to use this module
  
This module needs to be installed Microsoft Excel with active license. If you need work with xlsx files without Excel application, use the Rocketbot commands in the Files secction
## Overview


1. Open Without Alerts  
Open a file preventing MS Excel alerts.

2. Count columns  
Get the number of columns in a sheet

3. Count Rows  
Get the number of rows in a sheet

4. Cell color  
Change color of a cell or range of cells 

5. Insert Formula  
Insert formula into cell. Must be in English

6. Insert Macro  
Insert Macro in Excel from a file

7. Select Cells  
Select cells in Excel. Can copy the selected cells

8. Get Cell With Currency Format  
Get cells with currency format

9. Copy-Paste  
Copy range cell to another sheet

10. Format Cell  
Change the format of a cell or range of cells

11. Create Sheet  
Create a new sheet in the last position

12. Delete Sheet  
Delete a selected sheet

13. Copy to another excel  
Copy range to another Excel in the background

14. Add/Delete Row  
Add or Delete a Row

15. Add/Delete Column  
Add or Delete a Column

16. Convert CSV to XLSX  
Convert a csv document to xlsx

17. Convert XLSX to CSV  
Convert a xlsx document to csv

18. Convert XLS to XLSX  
Convert a xls document to xlsx

19. Get active cell  
Get row and column of active cell

20. Refresh Pivot table  
Refresh a pivot table. Deprecated! Use PivotTableExcel module

21. Fit cells  
Fit a cell or range of cells 

22. Get Formula  
Get formula into cell

23. Add Auto Filter  
Add auto filter to excel table

24. Filter  
Add filter to excel table

25. Rename sheet  
Change name to excel sheet

26. Cell Style  
Change the style of cell or range of cells

27. Paste in Cells  
Paste data to cells in Excel

28. Remove Duplicates  
Removes duplicate data in range

29. Close XLSX  
Close the workbook opened by Rocketbot

30. Save Excel  
Save a opened Excel file

31. Save Excel with password  
Save a opened Excel file with password

32. Export to advanced PDF  
Export to PDF with more options that the native command

33. Copy-Move Sheet  
Copy or move a sheet

34. Insert Form  
Insert Form in Excel

35. Read Filtered Cells  
Allow read only cells filters 

36. Count Filtered Cells  
Allow count only cells filters 

37. Replace  
Run replace action to excel 

38. Order  
Run replace action to excel 

39. Refresh All  
Refresh all data in Excel

40. Find  
Return de first found cell 

41. Lock Cells  
Lock or Unlock cells

42. Add Chart  
Create a new chart in an excel sheet

43. Remove Password  
Remove password and save the Excel

44. Insert image  
Insert an image

45. Export Chart  
Export a chart from index

46. Not visible mode  
Open not visible excel.

47. Write array objects  
Write array object on Excel cells.

48. Copy-Paste Format  
Copy format range cell to another sheet

49. Find and Connect  
Search a Excel Book opened and connect it

50. Update links  
Changes a link from one document to another

51. Unlock sheet  
Unlock sheet

52. Convert to .txt  
Convert a xlsx file to .txt

53. Text to columns  
Parses a column of cells that contain text into several columns.

54. Convert Excel time to hours  
Convert Excel time to hours
### Updates
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