



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

3. Maximize  
Maximize Excel Window

4. Calculation options  
Select the way the formula calculation is executed in the workbook.

5. Read cells  
Read a cell or range of cells

6. Convert serial date  
Convert an excel serial number date to a specific date format

7. Count columns  
Count the columns or return the last column name. It's necessary that the excel is saved to get the last changes

8. Count Rows  
Counts all the rows or from a range.

9. Hide  
Hides one or more rows, or one or more columns.

10. Show  
Shows one or more rows, or one or more columns that are hidden

11. Cell color  
Change color of a cell or range of cells. Can be a default color or custom

12. Get Cell Color  
Get the color of a cell. The funtion will return a list of two elements: Background Color and Font Color in RGB format.

13. Get Cell Formats  
Get the format of a cell. The function will return a dictionary with the cell properties and the value of each one.

14. Insert Formula  
Insert formula into cell

15. Insert Macro  
Insert Macro in Excel

16. Select and copy Cells  
Select and Copy cells in Excel

17. Get Cell With Currency Format  
Get cells with currency format

18. Get Cell With Date Format  
Get cells with date format

19. Copy-Paste  
Copy range cell to another sheet

20. Format Cell  
Format Cell

21. Clear Contents  
Clears formulas and values from the selected range, keeping the format.

22. Create Sheet  
Create sheet in the end

23. Delete Sheet  
Delete sheet

24. Copy to another excel  
Copy the range from one Excel file to another. Indicating the file path, it will open excel to copy or paste the data. If you enter the id of an open excel, it will use that instance to copy or paste.

25. Add/Delete Row  
Add or Delete a Row

26. Add/Delete Column  
Add or Delete a Column

27. Convert CSV to XLSX  
Convert a CSV document to XLSX format

28. Export to JSON  
Export array data to JSON

29. (Deprecated) Convert XLSX to CSV  
Convert a xlsx document to csv

30. Convert XLSX to CSV  
Convert a xlsx document to csv

31. Convert XLS to XLSX  
Convert a xls document to xlsx

32. Get active cell  
Get row and column of active cell

33. Refresh Pivot table  
Refresh a pivot table. Deprecated! Use PivotTableExcel module

34. Fit cells  
Adjusts, groups and ungroups a range of cells. You can group/ungroup by rows or columns

35. Get Formula  
Get the formula into cell

36. Add Auto Filter  
Add auto filter to excel table

37. Remove Auto Filter  
Remove auto filter from an excel sheet

38. Clear Filter  
Clears every filter applied over an excel sheet

39. Filter  
Filter an excel table according to the relative value, exact content, background color or font color of the cells. *Examples according the filter type: xlAnd ['>=10'] or ['>=10', '<=20'] | xlOr ['<=10', '>=20'] | xlFilterValues ['10','20', '30'] | xlFilterCellColor (255,0,0) | xlFilterFontColor (255,0,0)*

40. Filter by Date  
Filter a table by the day, month or year of a date indicated

41. Advanced filter  
Apply advanced filter to a table

42. Clear filters  
Remove filters and show all data

43. Rename sheet  
Change name to excel sheet

44. Text Format  
Change the Horizontal or Vertical alignment of values in a range of cells

45. Cell Style  
This command modifies the formatting of the selected cell or range of cells. You can change the font and borders

46. Paste in Cells  
Paste data to cells in Excel

47. Disable Cut/Copy Mode  
Disable Cut/Copy Mode of the active Excel

48. Remove Duplicates  
Execute the remove duplicates command of Excel

49. Export to advanced PDF  
Export to PDF with options

50. Copy-Move Sheet  
Copy or move a sheet

51. Insert Form  
Insert Form in Excel

52. Read Filtered Cells  
Read all the content of the filtered cells and apply formatting to date-type data if indicated

53. Count Filtered Cells  
Allow count only cells filters 

54. Replace  
Run replace action to excel 

55. Order  
Run replace action to excel 

56. Order by multiple levels  
Order an excel sheet by value, setting multiple levels

57. Refresh All  
Refresh all data in Excel

58. Find  
Searches a text in the given range and returns the address of the cell of the first occurence. If a value is not found, it will return empty. If the range it is filtered, the search will be performed over the visible cells

59. Find data  
Returns the first cell that matches the search data

60. Lock Cells  
Lock or Unlock cells

61. Add Chart  
Create a new chart in an excel sheet

62. Remove Password  
Remove password and save the Excel

63. Insert image  
Insert an image

64. Export Chart  
Export a chart from index

65. Not visible mode  
Open not visible excel.

66. Write array objects  
Write array object on Excel cells.

67. Copy-Paste Format  
Copy format range cell to another sheet

68. Update links  
Changes a link from one document to another

69. Unlock book  
Unlock book with password

70. Lock book  
Lock a book with password

71. Unlock sheet  
Unlock sheet with password

72. Lock sheet  
Lock a sheet with password

73. Convert to .txt  
Convert to .txt

74. Text to columns  
Parses a column of cells that contain text into several columns.

75. Convert Excel time to hours  
Convert Excel time to hours. Returns the format as hh: mm: ss

76. Print sheet  
Prints a sheet

77. Save Excel with password  
Save a Excel file

78. Save Excel  
Save an Excel file (as '.xlsx', 'xlsm', '.xls' or '.csv') in the indicated path

79. Close XLSX  
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