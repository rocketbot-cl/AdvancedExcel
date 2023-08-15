



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

7. Get Cell Formats  
Get the format of a cell. The function will return a dictionary with the cell properties and the value of each one.

8. Insert Formula  
Insert formula into cell

9. Insert Macro  
Insert Macro in Excel

10. Select and copy Cells  
Select and Copy cells in Excel

11. Get Cell With Currency Format  
Get cells with currency format

12. Get Cell With Date Format  
Get cells with date format

13. Copy-Paste  
Copy range cell to another sheet

14. Format Cell  
Format Cell

15. Clear Contents  
Clears formulas and values from the selected range, keeping the format.

16. Create Sheet  
Create sheet in the end

17. Delete Sheet  
Delete sheet

18. Copy to another excel  
Copy range from one Excel file to another. Use the current opened one, select one of the opened ones by ID or do everything in the background opening both Excels and closing them at the end.

19. Add/Delete Row  
Add or Delete a Row

20. Add/Delete Column  
Add or Delete a Column

21. Convert CSV to XLSX  
Convert a CSV document to XLSX format

22. (Deprecated) Convert XLSX to CSV  
Convert a xlsx document to csv

23. Convert XLSX to CSV  
Convert a xlsx document to csv

24. Convert XLS to XLSX  
Convert a xls document to xlsx

25. Get active cell  
Get row and column of active cell

26. Refresh Pivot table  
Refresh a pivot table. Deprecated! Use PivotTableExcel module

27. Fit cells  
Adjusts, groups and ungroups a range of cells. You can group/ungroup by rows or columns

28. Get Formula  
Get the formula into cell

29. Add Auto Filter  
Add auto filter to excel table

30. Remove Auto Filter  
Remove auto filter from an excel sheet

31. Clear Filter  
Clears every filter made over an excel sheet

32. Filter  
Filter an excel table according to the relative value, exact content, background color or font color of the cells. *Examples according the filter type: xlAnd ['>=10'] or ['>=10', '<=20'] | xlOr ['<=10', '>=20'] | xlFilterValues ['10','20', '30'] | xlFilterCellColor (255,0,0) | xlFilterFontColor (255,0,0)*

33. Advanced filter  
Apply advanced filter to a table

34. Clear filters  
Remove filters and show all data

35. Rename sheet  
Change name to excel sheet

36. Text Format  
Change the Horizontal or Vertical alignment of values in a range of cells

37. Cell Style  
This command modifies the formatting of the selected cell or range of cells. You can change the font and borders

38. Paste in Cells  
Paste data to cells in Excel

39. Remove Duplicates  
Execute the remove duplicates command of Excel

40. Export to advanced PDF  
Export to PDF with options

41. Copy-Move Sheet  
Copy or move a sheet

42. Insert Form  
Insert Form in Excel

43. Read Filtered Cells  
Allow read only cells filters 

44. Count Filtered Cells  
Allow count only cells filters 

45. Replace  
Run replace action to excel 

46. Order  
Run replace action to excel 

47. Refresh All  
Refresh all data in Excel

48. Find  
Searches a text in the given range and returns the address of the cell of the first occurence. If a value is not found, it will return empty. If the range it is filtered, the search will be performed over the visible cells

49. Find data  
Returns the first cell that matches the search data

50. Lock Cells  
Lock or Unlock cells

51. Add Chart  
Create a new chart in an excel sheet

52. Remove Password  
Remove password and save the Excel

53. Insert image  
Insert an image

54. Export Chart  
Export a chart from index

55. Not visible mode  
Open not visible excel.

56. Write array objects  
Write array object on Excel cells.

57. Copy-Paste Format  
Copy format range cell to another sheet

58. Update links  
Changes a link from one document to another

59. Unlock sheet  
Unlock sheet with password

60. Lock sheet  
Lock a sheet with password

61. Convert to .txt  
Convert to .txt

62. Text to columns  
Parses a column of cells that contain text into several columns.

63. Convert Excel time to hours  
Convert Excel time to hours. Returns the format as hh: mm: ss

64. Print sheet  
Prints a sheet

65. Save Excel with password  
Save a Excel file

66. Save Excel  
Save an Excel file (as '.xlsx', 'xlsm', '.xls' or '.csv') in the indicated path

67. Close XLSX  
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