



# Excel Advanced Options
  
Apply automatic and advanced filters, format cells, add or delete sheets, rows or columns, export to different file formats, unlock and relock sheets, copy and paste special and more with your Excel files.  

*Read this in other languages: [English](README.md), [Português](README.pr.md), [Español](README.es.md)*

## How to install this module
  
To install the module in Rocketbot Studio, it can be done in two ways:
1. Manual: __Download__ the .zip file and unzip it in the modules folder. The folder name must be the same as the module and inside it must have the following files and folders: \__init__.py, package.json, docs, example and libs. If you have the application open, refresh your browser to be able to use the new module.
2. Automatic: When entering Rocketbot Studio on the right margin you will find the **Addons** section, select **Install Mods**, search for the desired module and press install.  


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

6. Get Cell Color  
Get the color of a cell. The funtion will return a list of two elements: Background Color and Font Color in RGB format.

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

20. (Deprecated) Convert XLSX to CSV  
Convert a xlsx document to csv

21. Convert XLSX to CSV  
Convert a xlsx document to csv

22. Convert XLS to XLSX  
Convert a xls document to xlsx

23. Get active cell  
Get row and column of active cell

24. Refresh Pivot table  
Refresh a pivot table. Deprecated! Use PivotTableExcel module

25. Fit cells  
Adjusts, groups and ungroups a range of cells. You can group/ungroup by rows or columns

26. Get Formula  
Get the formula into cell

27. Add Auto Filter  
Add auto filter to excel table

28. Filter  
Filter an excel table according to the relative value, exact content, background color or font color of the cells. *Examples according the filter type: xlAnd ['>=10'] or ['>=10', '<=20'] | xlOr ['<=10', '>=20'] | xlFilterValues ['10','20', '30'] | xlFilterCellColor (255,0,0) | xlFilterFontColor (255,0,0)*

29. Advanced filter  
Apply advanced filter to a table

30. Clear filters  
Remove filters and show all data

31. Rename sheet  
Change name to excel sheet

32. Text Format  
Change the Horizontal or Vertical alignment of values in a range of cells

33. Cell Style  
This command modifies the formatting of the selected cell or range of cells. You can change the font and borders

34. Paste in Cells  
Paste data to cells in Excel

35. Remove Duplicates  
Execute the remove duplicates command of Excel

36. Export to advanced PDF  
Export to PDF with options

37. Copy-Move Sheet  
Copy or move a sheet

38. Insert Form  
Insert Form in Excel

39. Read Filtered Cells  
Allow read only cells filters 

40. Count Filtered Cells  
Allow count only cells filters 

41. Replace  
Run replace action to excel 

42. Order  
Run replace action to excel 

43. Refresh All  
Refresh all data in Excel

44. (Deprecated) Find  
Return de first found cell 

45. Find data  
Returns the first cell that matches the search data

46. Lock Cells  
Lock or Unlock cells

47. Add Chart  
Create a new chart in an excel sheet

48. Remove Password  
Remove password and save the Excel

49. Insert image  
Insert an image

50. Export Chart  
Export a chart from index

51. Not visible mode  
Open not visible excel.

52. Write array objects  
Write array object on Excel cells.

53. Copy-Paste Format  
Copy format range cell to another sheet

54. Update links  
Changes a link from one document to another

55. Unlock sheet  
Unlock sheet with password

56. Lock sheet  
Lock a sheet with password

57. Convert to .txt  
Convert to .txt

58. Text to columns  
Parses a column of cells that contain text into several columns.

59. Convert Excel time to hours  
Convert Excel time to hours. Returns the format as hh: mm: ss

60. Print sheet  
Prints a sheet

61. Save Excel with password  
Save a Excel file

62. Save Excel  
Save a Excel file in the indicated path

63. Close XLSX  
Close the workbook opened by Rocketbot  




----
### OS

- windows
- mac

### Dependencies
- [**xlwings**](https://pypi.org/project/xlwings/)- [**pandas**](https://pypi.org/project/pandas/)
### License
  
![MIT](https://camo.githubusercontent.com/107590fac8cbd65071396bb4d04040f76cde5bde/687474703a2f2f696d672e736869656c64732e696f2f3a6c6963656e73652d6d69742d626c75652e7376673f7374796c653d666c61742d737175617265)  
[MIT](http://opensource.org/licenses/mit-license.ph)