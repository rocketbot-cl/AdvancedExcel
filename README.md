



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

12. Font color  
Change the text color in a cell or range of cells. You can use a default or custom color

13. Get Cell Color  
Get the color of a cell. The funtion will return a list of two elements: Background Color and Font Color in RGB format.

14. Get Cell Formats  
Get the format of a cell. The function will return a dictionary with the cell properties and the value of each one.

15. Insert Formula  
Insert formula into cell

16. Insert Macro  
Insert Macro in Excel

17. Select and copy Cells  
Select and Copy cells in Excel

18. Get Cell With Currency Format  
Get cells with currency format

19. Get Cell With Date Format  
Get cells with date format

20. Copy-Paste  
Copy range cell to another sheet

21. Format Cell  
Format Cell

22. Clear Contents  
Clears formulas and values from the selected range, keeping the format.

23. Create Sheet  
Create sheet in the end

24. Delete Sheet  
Delete sheet

25. Copy to another excel  
Copy the range from one Excel file to another. Indicating the file path, it will open excel to copy or paste the data. If you enter the id of an open excel, it will use that instance to copy or paste.

26. Add/Delete Row  
Add or Delete a Row

27. Add/Delete Column  
Add or Delete a Column

28. Convert CSV to XLSX  
Convert a CSV document to XLSX format

29. Export to JSON  
Export array data to JSON

30. (Deprecated) Convert XLSX to CSV  
Convert a xlsx document to csv

31. Convert XLSX to CSV  
Convert a xlsx document to csv

32. Convert XLS to XLSX  
Convert a xls document to xlsx

33. Get active cell  
Get row and column of active cell

34. Refresh Pivot table  
Refresh a pivot table. Deprecated! Use PivotTableExcel module

35. Fit cells  
Adjusts, groups and ungroups a range of cells. You can group/ungroup by rows or columns

36. Get Formula  
Get the formula into cell

37. Add Auto Filter  
Add auto filter to excel table

38. Remove Auto Filter  
Remove auto filter from an excel sheet

39. Clear Filter  
Clears every filter applied over an excel sheet

40. Filter  
Filter an excel table according to the relative value, exact content, background color or font color of the cells. *Examples according the filter type: xlAnd ['>=10'] or ['>=10', '<=20'] | xlOr ['<=10', '>=20'] | xlFilterValues ['10','20', '30'] | xlFilterCellColor (255,0,0) | xlFilterFontColor (255,0,0)*

41. Filter by Date  
Filter a table by the day, month or year of a date indicated

42. Advanced filter  
Apply advanced filter to a table

43. Clear filters  
Remove filters and show all data

44. Rename sheet  
Change name to excel sheet

45. Text Format  
Change the Horizontal or Vertical alignment of values in a range of cells

46. Cell Style  
This command modifies the formatting of the selected cell or range of cells. You can change the font and borders

47. Paste in Cells  
Paste data to cells in Excel

48. Disable Cut/Copy Mode  
Disable Cut/Copy Mode of the active Excel

49. Remove Duplicates  
Execute the remove duplicates command of Excel

50. Export to advanced PDF  
Export to PDF with options

51. Copy-Move Sheet  
Copy or move a sheet

52. Insert Form  
Insert Form in Excel

53. Read Filtered Cells  
Read all the content of the filtered cells and apply formatting to date-type data if indicated

54. Count Filtered Cells  
Allow count only cells filters 

55. Replace  
Run replace action to excel 

56. Order  
Run replace action to excel 

57. Order by multiple levels  
Order an excel sheet by value, setting multiple levels

58. Refresh All  
Refresh all data in Excel

59. Find  
Searches a text in the given range and returns the address of the cell of the first occurence. If a value is not found, it will return empty. If the range it is filtered, the search will be performed over the visible cells

60. Find data  
Returns the first cell that matches the search data

61. Lock Cells  
Lock or Unlock cells

62. Add Chart  
Create a new chart in an excel sheet

63. Remove Password  
Remove password and save the Excel

64. Insert image  
Insert an image

65. Export Chart  
Export a chart from index

66. Not visible mode  
Open not visible excel.

67. Write array objects  
Write array object on Excel cells.

68. Copy-Paste Format  
Copy format range cell to another sheet

69. Update links  
Changes a link from one document to another

70. Unlock book  
Unlock book with password

71. Lock book  
Lock a book with password

72. Unlock sheet  
Unlock sheet with password

73. Lock sheet  
Lock a sheet with password

74. Convert to .txt  
Convert to .txt

75. Text to columns  
Parses a column of cells that contain text into several columns.

76. Convert Excel time to hours  
Convert Excel time to hours. Returns the format as hh: mm: ss

77. Combine spreadsheets  
Combine Excel spreadsheets that are in the same folder and have the same headers. It will combine horizontally the sheets of the same spreadsheet and vertically the different spreadsheets.

78. Print sheet  
Prints a sheet

79. Save Excel with password  
Save a Excel file

80. Save Excel  
Save an Excel file (as '.xlsx', 'xlsm', '.xls', '.csv' or '.prn') in the indicated path

81. Close XLSX  
Close the workbook opened by Rocketbot. The behavior of only closing one excel, works if it is opened with the command Open without alerts, otherwise it will close all.

82. Delete Styles  
Removes styles on a sheet

83. Insert link  
Insert link from a cell to a sheet  




----
### OS

- windows
- mac

### Dependencies
- [**xlwings**](https://pypi.org/project/xlwings/)- [**pandas**](https://pypi.org/project/pandas/)
### License
  
![MIT](https://camo.githubusercontent.com/107590fac8cbd65071396bb4d04040f76cde5bde/687474703a2f2f696d672e736869656c64732e696f2f3a6c6963656e73652d6d69742d626c75652e7376673f7374796c653d666c61742d737175617265)  
[MIT](http://opensource.org/licenses/mit-license.ph)