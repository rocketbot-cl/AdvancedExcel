# coding: utf-8
"""
Base para desarrollo de modulos externos.
Para obtener el modulo/Funcion que se esta llamando:
     GetParams("module")

Para obtener las variables enviadas desde formulario/comando Rocketbot:
    var = GetParams(variable)
    Las "variable" se define en forms del archivo package.json

Para modificar la variable de Rocketbot:
    SetVar(Variable_Rocketbot, "dato")

Para obtener una variable de Rocketbot:
    var = GetVar(Variable_Rocketbot)

Para obtener la Opcion seleccionada:
    opcion = GetParams("option")


Para instalar librerias se debe ingresar por terminal a la carpeta "libs"

    pip install <package> -t .

"""
# Changing the data types of all strings in the module at once
from __future__ import unicode_literals
import os
import sys
base_path = tmp_global_obj["basepath"]
cur_path = base_path + 'modules' + os.sep + 'AdvancedExcel' + os.sep + 'libs' + os.sep
sys.path.append(cur_path)
print(cur_path)
import pyexcel as p
from xlwt import Workbook
import platform
import xlwings as xw
from xlwings.constants import InsertShiftDirection
import pandas as pd
import io
import decimal


module = GetParams("module")

if module == "Open":
    excel = GetGlobals("excel")
    id_ = GetParams("id")
    file_path = GetParams("path")
    try:
        app = xw.App(add_book=False)
        app.display_alerts = False
        file_path = file_path.replace("/", os.sep)

        wb = app.books.api.Open(file_path, UpdateLinks=False)
        # wb = app.books.open(file_path, UpdateLinks=False)
        excel.actual_id = excel.id_default

        if id_:
            excel.actual_id = id_
        excel.file_[excel.actual_id] = {}
        excel.file_[excel.actual_id]['workbook'] = xw.Book(file_path)
        excel.file_[excel.actual_id]['app'] = excel.file_[excel.actual_id]['workbook'].app
        excel.file_[excel.actual_id]['sheet'] = excel.file_[excel.actual_id]['workbook'].sheets[0]
        excel.file_[excel.actual_id]['path'] = file_path

    except:
        PrintException()

if module == "CellColor":
    excel = GetGlobals("excel")

    range_ = GetParams("range")
    color = GetParams("color")

    if color == "red":
        rgb = (255, 0, 0)
        print("dos")
    elif color == "blue":
        rgb = (0, 0, 255)
    elif color == "green":
        rgb = (0, 255, 0)
    elif color == "grey":
        rgb = (130, 130, 130)
    else:
        rgb = (255, 255, 0)

    try:
        print("En el try")
        xls = excel.file_[excel.actual_id]

        # wb = xls['workbook']
        #         # print(wb)
        xw.Range(range_).color = rgb

        print("salimos")
        # xw.Range('A1:C1').column_width = 23
        # xw.Range('A1').row_height = 12
        # xw.Range('A2').formula = 2+2
        # print(xw.Range('A1'))
    except Exception as e:
        PrintException()
        raise e
if module == "InsertFormula":
    excel = GetGlobals("excel")

    cell = GetParams("cell")
    formula = GetParams("formula")

    xw.Range(cell).formula = formula

if module == "InsertMacro":
    macro = GetParams("macro_path")

    excel = GetGlobals("excel")
    xls = excel.file_[excel.actual_id]
    content_macro = None
    if macro and macro != "ERROR_NOT_VAR":
        if os.path.exists(macro):
            with open(macro, "r", encoding="latin-1") as m:
                content_macro = m.read()
                m.close()
        else:
            raise Exception("No existe el archivo de macro")
    else:
        raise Exception("No existe variable con ruta de macro")
    if content_macro is not None:
        tmp = xls['workbook'].api.VBProject.VBComponents.Add(1)
        tmp.CodeModule.AddFromString(content_macro.strip())

if module == "SelectCells":
    macro = """Sub RocketSelect(Ran as String, cop as Boolean)
                Range(Ran).Select
           
                If cop = True Then
                    Selection.Copy
                End If
                
                Range(Ran).Select
            End Sub
            Sub DeleteAllMacros() 'Excel vba to delete all macros in new workbook.
                Dim otmp As Object

                With ActiveWorkbook.VBProject
                    For Each otmp In .VBComponents
                        If otmp.Type=100 Then
                            otmp.CodeModule.DeleteLines 1, otmp.CodeModule.CountOfLines
                            otmp.CodeModule.CodePane.Window.Close
                        Else: .VBComponents.Remove otmp
                        End If
                    Next otmp
                End With
            End Sub 
            """

    excel = GetGlobals("excel")
    xls = excel.file_[excel.actual_id]
    cells = GetParams("cells")
    copy = GetParams("copy")

    if copy is None:
        copy = False

    print(copy)

    tmp = xls['workbook'].api.VBProject.VBComponents.Add(1)
    tmp.CodeModule.AddFromString(macro.strip())
    xls['workbook'].api.Application.Run("RocketSelect", cells, copy)
    xls['workbook'].api.Application.Run("DeleteAllMacros")

if module == "copyPaste":
    rango1 = GetParams("cell_range1")
    rango2 = GetParams("cell_range2")
    hoja1 = GetParams("sheet_name1")
    hoja2 = GetParams("sheet_name2")

    my_values = xw.sheets[hoja1].range(rango1).options(ndim=2).value

    xw.sheets[hoja2].range(rango2).value = my_values

if module == "formatCell":
    hoja = GetParams("sheet_name")
    rango = GetParams("cell_range")
    formato = GetParams("format_")

    try:
        if len(rango) == 1:
            rango = rango + ':' + rango
        if formato == "text":
            xw.sheets[hoja].range(rango).number_format = '@'

        if formato == "number_":
            numbers = xw.sheets[hoja].range(rango).value
            d = 0
            if type(numbers[0]) != list and len(numbers) == 1:
                numbers = [numbers]
            print(numbers)
            for i in range(len(numbers)):
                element = numbers[i]
                if type(element) == list:
                    for idx in range(len(element)):
                        print(idx)
                        number = element[idx]
                        print("number", number)
                        if type(element[idx]) is str:
                            number = number.split(",")
                            if "." in number[0]:
                                number[0] = number[0].replace(".", "")
                            number = ".".join(number)
                            if d < len(str(number).split(".")[1]):
                                d = len(str(number).split(".")[1])
                        element[idx] = number
                        numbers[i] = element
                else:
                    number = numbers[i]
                    if type(number) is str:
                        number = number.split(",")
                        print("number > ", number)
                        if "." in number[0]:
                            number[0] = number[0].replace(".", "")
                        number = ".".join(number)

                        if d < len(str(number).split(".")[1]):
                            d = len(str(number).split(".")[1])
                    numbers[i] = number

            if rango.split(":")[0][0] == rango.split(":")[1][0]:
                for i in range(len(numbers)):
                    numbers[i] = [numbers[i]]

            xw.sheets[hoja].range(rango).value = numbers
            print("format", xw.sheets[hoja].range(rango).number_format)
            if d == 0:
                xw.sheets[hoja].range(rango).number_format = '0'
            else:
                xw.sheets[hoja].range(rango).number_format = '0,{}'.format('0'*d)

        if formato == "coin_":
            xw.sheets[hoja].range(rango).number_format = '$#.##0'

        if formato == "date1":
            xw.sheets[hoja].range(rango).number_format = 'dd-mm-yyyy'

        if formato == "date2":
            xw.sheets[hoja].range(rango).number_format = 'dd-mm-yy'

        if formato == "date3":
            xw.sheets[hoja].range(rango).number_format = 'yyyy-mm-dd'

        if formato == "decimal1":
            xw.sheets[hoja].range(rango).number_format = '0,0'

        if formato == "decimal2":
            xw.sheets[hoja].range(rango).number_format = '#.##0,0'

        if formato == "long_date":
            xw.sheets[hoja].range(rango).number_format = 'dd/mm/yyyy h:mm:ss'

    except Exception as e:
        PrintException()
        raise e

if module == "createSheet":
    hoja = GetParams("sheet_name")

    res = [a.name for a in xw.sheets]
    last = res[-1]

    xw.sheets.add(name=hoja, after=last)

if module == "deleteSheet":

    hoja = GetParams("sheet_name")
    var_ = GetParams("var_")
    res = False

    for sheet in xw.sheets:
        if hoja in sheet.name:
            sheet.delete()
            res = True

    SetVar(var_,res)
if module == "copy_other":
    try:
        excel1 = GetParams("excel1")
        excel2 = GetParams("excel2")
        hoja1 = GetParams("sheet_name1")
        hoja2 = GetParams("sheet_name2")
        rango1 = GetParams("cell_range1")
        rango2 = GetParams("cell_range2")
        platform_ = platform.system()

        app = xw.App(visible=True)
        wb1 = xw.Book(excel1)
        wb2 = xw.Book(excel2)

        my_values = wb1.sheets[hoja1].range(rango1).options(ndim=2).value

        wb2.sheets[hoja2].range(rango2).value = my_values

        if platform_ == 'Windows':
            wb2.save(excel2)
            wb2.close()
            #wb1.close()

        else:
            wb2.save()
            wb2.close()

        app.quit()
    except:
        PrintException()

if module == "addRow":

    try:
        hoja = GetParams("sheet")
        fila_ = GetParams("row_")
        tipo = GetParams("type_")
        opcion_ = GetParams("option_")
        print(hoja)
        print(tipo)
        platform_ = platform.system()

        if opcion_ == "add_":

            if platform_ == 'Windows':
                if ":" in fila_:

                    if tipo == "down_":
                        fila = fila_.split(':')
                        f1 = fila[0]
                        f1 = int(f1) + 1
                        f2 = fila[1]
                        f2 = int(f2) + 1
                        fila = str(f1) + ':' + str(f2)
                        print('FILA down', fila)
                        # fila_ = int(fila_) + 1
                        xw.sheets[hoja].range(fila).api.Insert(InsertShiftDirection.xlShiftDown)

                    if tipo == "up_":
                        print('FILA up', fila_)
                        xw.sheets[hoja].range(fila_).api.Insert(InsertShiftDirection.xlShiftDown)

                else:
                    if tipo == "down_":
                        fila_ = int(fila_) + 1
                        # print(fila_)
                        fila = str(fila_) + ':' + str(fila_)
                        print('FILA down', fila)

                        xw.sheets[hoja].range(fila).api.Insert(InsertShiftDirection.xlShiftDown)

                    if tipo == "up_":
                        fila = str(fila_) + ':' + str(fila_)
                        print('FILA up', fila)
                        xw.sheets[hoja].range(fila).api.Insert(InsertShiftDirection.xlShiftDown)


            else:
                if ":" in fila_:
                    if tipo == "down_":
                        fila = fila_.split(':')
                        f1 = fila[0]
                        f1 = int(f1) + 1
                        f2 = fila[1]
                        f2 = int(f2) + 1
                        fila = str(f1) + ':' + str(f2)
                        # fila_ = int(fila_) + 1
                        xw.sheets[hoja].api.rows[fila].insert_into_range()
                    if tipo == "up_":
                        xw.sheets[hoja].api.rows[fila_].insert_into_range()

                else:
                    if tipo == "down_":
                        fila_ = int(fila_) + 1
                        # print(fila_)
                        fila = str(fila_) + ':' + str(fila_)

                        xw.sheets[hoja].api.rows[fila].insert_into_range()

                    if tipo == "up_":
                        fila = str(fila_) + ':' + str(fila_)
                        xw.sheets[hoja].api.rows[fila].insert_into_range()

        if opcion_ == "delete_":
            if ":" in fila_:
                xw.Range(fila_).api.delete()
            else:
                fila = str(fila_) + ':' + str(fila_)
                # print(fila)
                xw.Range(fila).api.delete()

    except:
        PrintException()

if module == "addCol":
    try:
        hoja = GetParams("sheet")
        col_ = GetParams("col_")
        opcion_ = GetParams("option_")
        platform_ = platform.system()

        if opcion_ == "add_":

            if platform_ == 'Windows':

                if ":" in col_:
                    xw.sheets[hoja].range(col_).api.Insert(InsertShiftDirection.xlShiftToRight)
                else:
                    col = str(col_) + ':' + str(col_)
                    xw.sheets[hoja].range(col).api.Insert(InsertShiftDirection.xlShiftToRight)


            else:
                if ":" in col_:
                    xw.sheets[hoja].api.columns[col_].insert_into_range()
                else:
                    col = str(col_) + ':' + str(col_)
                    xw.sheets[hoja].api.columns[col].insert_into_range()

        if opcion_ == "delete_":
            if platform_ == 'Windows':
                if ":" in col_:
                    xw.Range(col_).api.Delete()
                else:
                    col = str(col_) + ':' + str(col_)
                    xw.Range(col).api.Delete()

            else:
                if ":" in col_:
                    xw.Range(col_).api.delete()
                else:
                    col = str(col_) + ':' + str(col_)
                    xw.Range(col).api.delete()

    except:
        PrintException()

if module == "csvToxlsx":
    csv_path = GetParams("csv_path")
    xlsx_path = GetParams("xlsx_path")
    sep = GetParams("separator") or ","

    if not csv_path or not xlsx_path:
        raise Exception("Falta una ruta")
    f_ = open(csv_path, 'r', enconding='latin-1')
    df = pd.read_csv(f_, sep=sep)
    df.to_excel(xlsx_path, index=None)
    f_.close()

if module == "countColumns":

    excel = GetGlobals("excel")

    sheet = GetParams("sheet")
    result = GetParams("var_")

    try:
        excel_path = excel.file_["default"]["path"]
        print(excel_path)
        df = pd.read_excel(excel_path, sheetname=sheet)
        print(df)
        col = df.shape[1]

        if result:
            SetVar(result, col)

    except Exception as e:
        PrintException()
        raise e

if module == "countRows":

    excel = GetGlobals("excel")

    sheet = GetParams("sheet")
    row_ = GetParams("row_")
    result = GetParams("var_")

    if not sheet:
        sheet = 0
    if not row_:
        row_ = 'A'

    try:
        #excel_path = excel.file_["default"]["path"]
        #print(excel_path)
        total = xw.sheets[sheet].range(row_ + str(xw.sheets[sheet].cells.last_cell.row)).end('up').row
        #print(total)

        if result:
            SetVar(result, total)

    except Exception as e:
        PrintException()
        raise e

if module == "xlsToxlsx":

    xls_path = GetParams('xls_path')
    xlsx_path = GetParams('xlsx_path')
    print(xls_path,xlsx_path)

    try:
        try:
            p.save_book_as(file_name=xls_path,
                            dest_file_name=xlsx_path)
        except:

            filename = xls_path
            # Opening the file using 'utf-16' encoding
            file1 = io.open(filename, "r", encoding="utf-16")
            data = file1.readlines()

            # Creating a workbook object
            xldoc = Workbook()
            # Adding a sheet to the workbook object
            sheet = xldoc.add_sheet("Sheet1", cell_overwrite_ok=True)
            # Iterating and saving the data to sheet
            for i, row in enumerate(data):
                # Two things are done here
                # Removeing the '\n' which comes while reading the file using io.open
                # Getting the values after splitting using '\t'
                for j, val in enumerate(row.replace('\n', '').split('\t')):
                    sheet.write(i, j, val)

            # Saving the file as an excel file
            xldoc.save(xls_path)

            p.save_book_as(file_name=xls_path,
                           dest_file_name=xlsx_path)
    except Exception as e:
        PrintException()
        raise e

if module == "getActiveCell":
    excel = GetGlobals("excel")
    result = GetParams("result")

    xls = excel.file_[excel.actual_id]
    wb = xls['workbook']

    abc = ['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k', 'l', 'm', 'n', 'o', 'p', 'q', 'r', 's', 't', 'u', 'v',
           'w', 'x', 'y', 'z']

    try:
        col = int(wb.app.selection.column)
        row = wb.app.selection.row

        length = len(abc)
        if col > length:
            excess = col // length
            mod = col % length

            col = abc[excess - 1] + abc[mod - 1]
        else:
            col = abc[col - 1]

        print(row, "******")
        ans = col + str(row)

        SetVar(result, ans)
    except Exception as e:
        PrintException()
        raise e

if module == "updatePivot":
    sheet = GetParams("sheet")
    pivotTableName = GetParams("table")
    excel = GetGlobals("excel")

    xls = excel.file_[excel.actual_id]
    wb = xls['workbook']
    wb.sheets[sheet].select()
    wb.api.ActiveSheet.PivotTables(pivotTableName).PivotCache().refresh()

if module == "filter":

    data = GetParams("data")
    print("*"*1000 + "\n", data)
    data = eval(data)
    col = GetParams("col").lower()
    type_filter_col = GetParams("type_filter_col")
    filter_col = GetParams("filter_col")
    var_ = GetParams("var_")
    list = []
    cont = 0


    try:
        abc = ['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k', 'l', 'm', 'n', 'o', 'p', 'q', 'r', 's', 't', 'u',
               'v', 'w', 'x', 'y', 'z']

        around_abc = len(col) - 1
        col = col[-1]
        col_index = around_abc * len(abc) + abc.index(col)
        col_index = int(col_index)

        if col:
            for d in data:
                #print(d, "******")
                print('data',d)
                print('f',filter_col)
                if type_filter_col == "equal":
                    if d[col_index] == filter_col:
                        print('equal',d[col_index],filter_col)
                        list.append(d)
                        print('LIST',list)
                if type_filter_col == "not_equal":
                    print('not',d[col_index], "\n")
                    if d[col_index] != filter_col:
                        list.append(d)

        SetVar(var_, list)
    except:
        PrintException()
