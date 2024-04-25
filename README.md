



# Excel Advanced Options
  
Apply automatic and advanced filters, format cells, add or delete sheets, rows or columns, export to different file formats, unlock and relock sheets, copy and paste special and more with your Excel files.  

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

4. Opciones de calculo  
Selecciona la manera en que se ejecuta el calculo de formulas en el libro.

5. Read cells  
Read a cell or range of cells

6. Convert serial date  
Convert an excel serial number date to a specific date format

7. Count columns  
Count the columns or return the last column name. It's necessary that the excel is saved to get the last changes

8. Count Rows  
Counts all the rows or from a range.

9. Cell color  
Change color of a cell or range of cells. Can be a default color or custom

10. Get Cell Color  
Get the color of a cell. The funtion will return a list of two elements: Background Color and Font Color in RGB format.

11. Get Cell Formats  
Get the format of a cell. The function will return a dictionary with the cell properties and the value of each one.

12. Insert Formula  
Insert formula into cell

13. Insert Macro  
Insert Macro in Excel

14. Select and copy Cells  
Select and Copy cells in Excel

15. Get Cell With Currency Format  
Get cells with currency format

16. Get Cell With Date Format  
Get cells with date format

17. Copy-Paste  
Copy range cell to another sheet

18. Format Cell  
Format Cell

19. Clear Contents  
Clears formulas and values from the selected range, keeping the format.

20. Create Sheet  
Create sheet in the end

21. Delete Sheet  
Delete sheet

22. Copy to another excel  
Copy the range from one Excel file to another. Indicating the file path, it will open excel to copy or paste the data. If you enter the id of an open excel, it will use that instance to copy or paste.

23. Add/Delete Row  
Add or Delete a Row

24. Add/Delete Column  
Add or Delete a Column

25. Convert CSV to XLSX  
Convert a CSV document to XLSX format

26. (Deprecated) Convert XLSX to CSV  
Convert a xlsx document to csv

27. Convert XLSX to CSV  
Convert a xlsx document to csv

28. Convert XLS to XLSX  
Convert a xls document to xlsx

29. Get active cell  
Get row and column of active cell

30. Refresh Pivot table  
Refresh a pivot table. Deprecated! Use PivotTableExcel module

31. Fit cells  
Adjusts, groups and ungroups a range of cells. You can group/ungroup by rows or columns

32. Get Formula  
Get the formula into cell

33. Add Auto Filter  
Add auto filter to excel table

34. Remove Auto Filter  
Remove auto filter from an excel sheet

35. Clear Filter  
Clears every filter applied over an excel sheet

36. Filter  
Filter an excel table according to the relative value, exact content, background color or font color of the cells. *Examples according the filter type: xlAnd ['>=10'] or ['>=10', '<=20'] | xlOr ['<=10', '>=20'] | xlFilterValues ['10','20', '30'] | xlFilterCellColor (255,0,0) | xlFilterFontColor (255,0,0)*

37. Filter by Date  
Filter a table by the day, month or year of a date indicated

38. Advanced filter  
Apply advanced filter to a table

39. Clear filters  
Remove filters and show all data

40. Rename sheet  
Change name to excel sheet

41. Text Format  
Change the Horizontal or Vertical alignment of values in a range of cells

42. Cell Style  
This command modifies the formatting of the selected cell or range of cells. You can change the font and borders

43. Paste in Cells  
Paste data to cells in Excel

44. Disable Cut/Copy Mode  
Disable Cut/Copy Mode of the active Excel

45. Remove Duplicates  
Execute the remove duplicates command of Excel

46. Export to advanced PDF  
Export to PDF with options

47. Copy-Move Sheet  
Copy or move a sheet

48. Insert Form  
Insert Form in Excel

49. Read Filtered Cells  
Allow read only cells filters 

50. Count Filtered Cells  
Allow count only cells filters 

51. Replace  
Run replace action to excel 

52. Order  
Run replace action to excel 

53. Order by multiple levels  
Order an excel sheet by value, setting multiple levels

54. Refresh All  
Refresh all data in Excel

55. Find  
Searches a text in the given range and returns the address of the cell of the first occurence. If a value is not found, it will return empty. If the range it is filtered, the search will be performed over the visible cells

56. Find data  
Returns the first cell that matches the search data

57. Lock Cells  
Lock or Unlock cells

58. Add Chart  
Create a new chart in an excel sheet

59. Remove Password  
Remove password and save the Excel

60. Insert image  
Insert an image

61. Export Chart  
Export a chart from index

62. Not visible mode  
Open not visible excel.

63. Write array objects  
Write array object on Excel cells.

64. Copy-Paste Format  
Copy format range cell to another sheet

65. Update links  
Changes a link from one document to another

66. Unlock book  
Unlock book with password

67. Lock book  
Lock a book with password

68. Unlock sheet  
Unlock sheet with password

69. Lock sheet  
Lock a sheet with password

70. Convert to .txt  
Convert to .txt

71. Text to columns  
Parses a column of cells that contain text into several columns.

72. Convert Excel time to hours  
Convert Excel time to hours. Returns the format as hh: mm: ss

73. Print sheet  
Prints a sheet

74. Save Excel with password  
Save a Excel file

75. Save Excel  
Save an Excel file (as '.xlsx', 'xlsm', '.xls' or '.csv') in the indicated path

76. Close XLSX  
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