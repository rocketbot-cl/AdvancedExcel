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
from xlsx2csv import Xlsx2csv


def get_date_with_format(xl_date):
    import xlrd
    datetime_date = xlrd.xldate_as_datetime(xl_date, 0)
    date_object = datetime_date.date()
    return date_object.isoformat()


module = GetParams("module")

if module == "Open":
    excel = GetGlobals("excel")
    id_ = GetParams("id")
    file_path = GetParams("path")
    password = GetParams("password")
    visible = GetParams("visible")
    try:

        app = xw.App(add_book=False)

        app.api.DisplayAlerts = False

        file_path = file_path.replace("/", os.sep)

        wb = app.api.Workbooks.Open(file_path, False, None, None, password, password, IgnoreReadOnlyRecommended=True,
                                    CorruptLoad=2)
        # wb = app.books.open(file_path, UpdateLinks=False)
        excel.actual_id = excel.id_default

        if id_:
            excel.actual_id = id_
        excel.file_[excel.actual_id] = {}
        excel.file_[excel.actual_id]['workbook'] = app.books[0]
        excel.file_[excel.actual_id]['app'] = excel.file_[excel.actual_id]['workbook'].app
        excel.file_[excel.actual_id]['sheet'] = excel.file_[excel.actual_id]['workbook'].sheets[0]
        excel.file_[excel.actual_id]['path'] = file_path

    except Exception as e:
        print("\x1B[" + "31;40mError\x1B[" + "0m")
        PrintException()
        raise e

if module == "CellColor":
    excel = GetGlobals("excel")

    range_ = GetParams("range")
    color = GetParams("color")
    custom = GetParams("custom")

    try:
        if color == "red":
            rgb = (255, 0, 0)
            print("dos")
        elif color == "blue":
            rgb = (0, 0, 255)
        elif color == "green":
            rgb = (0, 255, 0)
        elif color == "grey":
            rgb = (130, 130, 130)
        elif color == "yell":
            rgb = (255, 255, 0)
        else:
            rgb = eval(custom)

        xls = excel.file_[excel.actual_id]

        # wb = xls['workbook']
        #         # print(wb)
        xw.Range(range_).color = rgb

        # print("salimos")
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

    excel = GetGlobals("excel")
    xls = excel.file_[excel.actual_id]
    cells = GetParams("cells")
    copy = GetParams("copy")
    sheet = GetParams("sheet_name")

    if copy is None:
        copy = False

    try:

        wb = xls['workbook']
        if not sheet in [sh.name for sh in wb.sheets]:
            raise Exception(f"The name {sheet} does not exist in the book")
        wb.sheets[sheet].select()
        wb.sheets[sheet].range(cells).select()

        if copy:
            wb.sheets[sheet].api.Range(cells).Copy()
    except Exception as e:
        print("\x1B[" + "31;40mError\x1B[" + "0m")
        PrintException()
        raise e

if module == "copyPaste":
    rango1 = GetParams("cell_range1")
    rango2 = GetParams("cell_range2")
    hoja1 = GetParams("sheet_name1")
    hoja2 = GetParams("sheet_name2")

    if not hoja1 in [sh.name for sh in xw.sheets]:
        raise Exception(f"The name {hoja1} does not exist in the book")
    if not hoja2 in [sh.name for sh in xw.sheets]:
        raise Exception(f"The name {hoja2} does not exist in the book")
    my_values = xw.sheets[hoja1].range(rango1).options(ndim=2).value

    xw.sheets[hoja2].range(rango2).value = my_values

if module == "formatCell":
    hoja = GetParams("sheet_name")
    rango = GetParams("cell_range")
    formato = GetParams("format_")
    custom = GetParams("custom")

    try:
        if not hoja in [sh.name for sh in xw.sheets]:
            raise Exception(f"The name {hoja} does not exist in the book")
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
                        tmp = str(number).split(".")
                        if len(tmp) > 1:
                            if d < len(tmp[1]):
                                d = len(tmp[1])
                    numbers[i] = number

            if rango.split(":")[0][0] == rango.split(":")[1][0]:
                for i in range(len(numbers)):
                    numbers[i] = [numbers[i]]

            xw.sheets[hoja].range(rango).value = numbers
            print("format", xw.sheets[hoja].range(rango).number_format)
            if d == 0:
                xw.sheets[hoja].range(rango).number_format = '0'
            else:
                xw.sheets[hoja].range(rango).number_format = '0,{}'.format('0' * d)

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
        if formato == 'custom':
            xw.sheets[hoja].range(rango).number_format = custom

    except Exception as e:
        PrintException()
        raise e

if module == "createSheet":
    hoja = GetParams("sheet_name")
    last = GetParams("after")
    excel = GetGlobals("excel")

    try:
        xls = excel.file_[excel.actual_id]
        wb = xls['workbook']

        if not last:
            res = [a.name for a in wb.sheets]
            last = res[-1]

        wb.sheets.add(name=hoja, after=last)
    except Exception as e:
        print("\x1B[" + "31;40mError\x1B[" + "0m")
        PrintException()
        raise e

if module == "deleteSheet":

    hoja = GetParams("sheet_name")
    var_ = GetParams("var_")
    res = False

    for sheet in xw.sheets:
        if hoja in sheet.name:
            sheet.delete()
            res = True

    SetVar(var_, res)
if module == "copy_other":
    import datetime
    try:
        excel1 = GetParams("excel1")
        excel2 = GetParams("excel2")
        hoja1 = GetParams("sheet_name1")
        hoja2 = GetParams("sheet_name2")
        rango1 = GetParams("cell_range1")
        rango2 = GetParams("cell_range2")
        platform_ = platform.system()
        excel = GetGlobals("excel")

        xls = excel.file_[excel.actual_id]
        wb = xls['workbook']

        wb1 = wb.app.books.open(excel1)
        if hoja1 not in [sh.name for sh in wb1.sheets]:
            raise Exception(f"The name {hoja1} does not exist in the book {excel1.split('/')[-1]}")

        origin_sheet = wb1.sheets[hoja1]
        my_values = origin_sheet.range(rango1)

        if platform_ == "Windows":
            values = my_values.api.Value2

            password = None
            # print(test)
            wb2 = wb.app.books.api.Open(excel2, False, None, None, password, password, IgnoreReadOnlyRecommended=True,
                                        CorruptLoad=2)
            if hoja2 not in [sh.name for sh in wb2.sheets]:
                raise Exception(f"The name {hoja2} does not exist in the book  {excel2.split('/')[-1]}")
            destiny_sheet = wb2.Sheets(hoja2)

            if ":" not in rango2:
                len_row = len(values) + destiny_sheet.Range(rango2).Row - 1
                len_col = len(values[0]) + destiny_sheet.Range(rango2).Column - 1
                rango2 = rango2 + ":" + destiny_sheet.Cells(len_row, len_col).Address
            
            destiny_sheet.Range(rango2).value = values
            wb2.Save()
            wb2.Close()
        else:
            values = my_values.value
            wb2 = wb.app.books.open(excel2)
            if hoja2 not in [sh.name for sh in wb2.sheets]:
                raise Exception(f"The name {hoja2} does not exist in the book  {excel2.split('/')[-1]}")
            destiny_sheet = wb2.sheets(hoja2)
            destiny_sheet.range(rango2).value = values

            wb2.save()
            wb2.close()

    except Exception as e:
        print("\x1B[" + "31;40mError\x1B[" + "0m")
        PrintException()
        raise e

if module == "addRow":

    try:
        sheet_name = GetParams("sheet")
        row = GetParams("row_")
        tipo = GetParams("type_")
        opcion_ = GetParams("option_")

        excel = GetGlobals("excel")
        xls = excel.file_[excel.actual_id]
        wb = xls['workbook']

        if not sheet_name in [sh.name for sh in wb.sheets]:
            raise Exception(f"The name {sheet_name} does not exist in the book")

        sheet = wb.sheets[sheet_name]
        platform_ = platform.system()

        if opcion_ == "add_":

            if platform_ == 'Windows':
                if ":" in row:
                    if tipo == "down_":
                        fila = row.split(':')
                        f1 = fila[0]
                        f1 = int(f1) + 1
                        f2 = fila[1]
                        f2 = int(f2) + 1
                        fila = str(f1) + ':' + str(f2)
                        print('FILA down', fila)
                        # row = int(row) + 1
                        sheet.range(fila).api.Insert(InsertShiftDirection.xlShiftDown)

                    if tipo == "up_":
                        print('FILA up', row)
                        sheet.range(row).api.Insert(InsertShiftDirection.xlShiftDown)

                else:
                    if tipo == "down_":
                        row = int(row) + 1
                        # print(row)
                        fila = str(row) + ':' + str(row)
                        print('FILA down', fila)

                        sheet.range(fila).api.Insert(InsertShiftDirection.xlShiftDown)

                    if tipo == "up_":
                        fila = str(row) + ':' + str(row)
                        print('FILA up', fila)
                        sheet.range(fila).api.Insert(InsertShiftDirection.xlShiftDown)


            else:
                if ":" in row:
                    if tipo == "down_":
                        fila = row.split(':')
                        f1 = fila[0]
                        f1 = int(f1) + 1
                        f2 = fila[1]
                        f2 = int(f2) + 1
                        fila = str(f1) + ':' + str(f2)
                        # row = int(row) + 1
                        sheet.api.rows[fila].insert_into_range()
                    if tipo == "up_":
                        sheet.api.rows[row].insert_into_range()

                else:
                    if tipo == "down_":
                        row = int(row) + 1
                        # print(row)
                        fila = str(row) + ':' + str(row)

                        sheet.api.rows[fila].insert_into_range()

                    if tipo == "up_":
                        fila = str(row) + ':' + str(row)
                        sheet.api.rows[fila].insert_into_range()

        if opcion_ == "delete_":
            if ":" not in row:
                row = str(row) + ":" + str(row)

            sheet.range(row).api.Delete()


    except Exception as e:
        PrintException()
        raise e


if module == "addCol":
    try:
        hoja = GetParams("sheet")
        col_ = GetParams("col_")
        opcion_ = GetParams("option_")
        platform_ = platform.system()

        if not hoja in [sh.name for sh in xw.sheets]:
            raise Exception(f"The name {hoja} does not exist in the book")

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
    with_header = GetParams("header")

    try:
        if not csv_path or not xlsx_path:
            raise Exception("Falta una ruta")

        import csv
        from openpyxl import Workbook, load_workbook

        platform_ = platform.system()
        if platform_ == "Windows":
            import ctypes as ct
            csv.field_size_limit(int(ct.c_ulong(-1).value // 2))
            limit1 = csv.field_size_limit()
        if sep.startswith("\\t"):
            sep = "\t"
        workbook = openpyxl.Workbook()
        worksheet = workbook.active
        with open(csv_path, "r", encoding="latin-1") as fobj:
            csv_reader = csv.reader(fobj, delimiter=sep)
            for row_index, row in enumerate(csv_reader):
                for col_index, value in enumerate(row):
                    worksheet.cell(row_index + 1, col_index + 1).value = value
        workbook.save(xlsx_path)

    except Exception as e:
        print("\x1B[" + "31;40mAn error occurred\x1B[" + "0m")
        PrintException()
        raise e

if module == "xlsxToCsv":
    csv_path = GetParams("csv_path")
    xlsx_path = GetParams("xlsx_path")
    delimiter = GetParams("delimiter")

    try:
        if not delimiter:
            delimiter = ","

        data_xls = pd.read_excel(xlsx_path, 'Sheet0', index_col=None, header=None)
        data_xls.to_csv(csv_path, encoding='utf-8', index=False, header=False)
        # Xlsx2csv(xlsx_path, outputencoding="utf-8", delimiter=delimiter, floatformat=True).convert(csv_path)
    except Exception as e:
        PrintException()
        raise e

if module == "countColumns":

    excel = GetGlobals("excel")

    sheet = GetParams("sheet")
    result = GetParams("var_")

    try:
        xls = excel.file_[excel.actual_id]
        wb = xls['workbook']
        # excel_path = excel.file_["default"]["path"]
        excel_path = wb.fullname
        print(excel_path)

        df = pd.read_excel(excel_path, sheet_name=sheet)
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
        # excel_path = excel.file_["default"]["path"]
        # print(excel_path)
        total = xw.sheets[sheet].range(row_ + str(xw.sheets[sheet].cells.last_cell.row)).end('up').row
        # print(total)

        if result:
            SetVar(result, total)

    except Exception as e:
        PrintException()
        raise e

if module == "xlsToxlsx":

    xls_path = GetParams('xls_path')
    xlsx_path = GetParams('xlsx_path')
    print(xls_path, xlsx_path)

    try:
        try:
            p.save_book_as(file_name=xls_path,
                           dest_file_name=xlsx_path)
        except:

            filename = xls_path
            # Opening the file using 'utf-16' encoding
            file1 = io.open(filename, "r", encoding="utf-16")
            print(dir(file1), file1.read())
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
        print("\x1B[" + "31;40mAn error occurred\x1B[" + "0m")
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

if module == "refreshPivot":
    sheet = GetParams("sheet")
    pivotTableName = GetParams("table")
    excel = GetGlobals("excel")

    xls = excel.file_[excel.actual_id]
    wb = xls['workbook']
    if not sheet in [sh.name for sh in wb.sheets]:
        raise Exception(f"The name {sheet} does not exist in the book")
    wb.sheets[sheet].select()
    print(dir(wb.api.ActiveSheet.PivotTables(pivotTableName)))
    wb.api.ActiveSheet.PivotTables(pivotTableName).PivotCache().refresh()

if module == "fitCells":
    sheet = GetParams("sheet")
    range_cell = GetParams("cell_range")
    excel = GetGlobals("excel")

    xls = excel.file_[excel.actual_id]
    wb = xls['workbook']
    if not sheet in [sh.name for sh in wb.sheets]:
        raise Exception(f"The name {sheet} does not exist in the book")
    sh = wb.sheets[sheet].autofit()

if module == "CloseExcel":
    excel = GetGlobals("excel")
    xls = excel.file_[excel.actual_id]
    wb = xls['workbook']
    xw.books.active.close()

if module == "getFormula":
    excel = GetGlobals("excel")

    cell = GetParams("cell")
    result = GetParams("var_")

    try:
        formula = xw.Range(cell).formula
        SetVar(result, formula)
    except Exception as e:
        print("\x1B[" + "31;40mError\x1B[" + "0m")
        PrintException()
        raise e

if module == "AutoFilter":
    sheet = GetParams("sheet")
    columns = GetParams("columns")
    excel = GetGlobals("excel")

    try:
        xls = excel.file_[excel.actual_id]
        wb = xls['workbook']
        if not sheet in [sh.name for sh in wb.sheets]:
            raise Exception(f"The name {sheet} does not exist in the book")
        wb.sheets[sheet].select()
        wb.sheets[sheet].api.Range(columns).AutoFilter()

    except Exception as e:
        print("\x1B[" + "31;40mError\x1B[" + "0m")
        PrintException()
        raise e

if module == "Filter":

    try:
        sheet = GetParams("sheet")
        start = GetParams("start")
        column = GetParams("column")
        data = GetParams("filter")
        result = GetParams("var_")
        excel = GetGlobals("excel")

        xls = excel.file_[excel.actual_id]
        wb = xls['workbook']
        if not sheet in [sh.name for sh in wb.sheets]:
            raise Exception(f"The name {sheet} does not exist in the book")
        wb.sheets[sheet].select()
        if ":" in start:
            range_ = start
            start = start.split(":")[0]
        else:
            start = start + str(1)
            range_ = column + str(1)

        n_start = wb.sheets[sheet].range(start).column
        n_end = wb.sheets[sheet].range(column + str(1)).column

        filter_column = n_end - n_start + 1
        if data.startswith("["):
            data = eval(data)

        wb.sheets[sheet].api.Range(range_).AutoFilter(filter_column, data, 7)

    except Exception as e:
        print("\x1B[" + "31;40mError\x1B[" + "0m")
        PrintException()
        raise e

if module == "rename_sheet":
    sheet = GetParams("sheet")
    name = GetParams("name")
    excel = GetGlobals("excel")

    try:
        xls = excel.file_[excel.actual_id]
        wb = xls['workbook']
        if not sheet in [sh.name for sh in wb.sheets]:
            raise Exception(f"The name {sheet} does not exist in the book")
        wb.sheets[sheet].select()
        wb.sheets[sheet].name = name

    except Exception as e:
        print("\x1B[" + "31;40mError\x1B[" + "0m")
        PrintException()
        raise e

if module == "style_cells":
    sheet = GetParams("sheet_name")
    range_ = GetParams("cell_range")
    position = GetParams("position")
    line_style = GetParams("lineStyle")

    font_size = GetParams("size")
    bold = GetParams("bold")
    underline = GetParams("underline")
    italic = GetParams("italic")

    excel = GetGlobals("excel")

    try:
        xls = excel.file_[excel.actual_id]
        wb = xls['workbook']
        print(range_)
        if not sheet in [sh.name for sh in wb.sheets]:
            raise Exception(f"The name {sheet} does not exist in the book")
        rng = wb.sheets[sheet].api.Range(range_)
        if line_style:
            line_style = int(line_style)
            if position == "all":
                for i in range(7, 13):
                    rng.Borders(i).LineStyle = line_style
            elif position == "contour":
                for i in range(7, 11):
                    rng.Borders(i).LineStyle = line_style
            else:
                position = int(position)
                print(position)
                rng.Borders(position).LineStyle = line_style

        if font_size and font_size.isnumeric:
            rng.Font.Size = int(font_size)
        if underline:
            rng.Font.Underline = 2
        if bold:
            rng.Font.Bold = True
        if italic:
            rng.Font.Italic = True



    except Exception as e:
        print("\x1B[" + "31;40mError\x1B[" + "0m")
        PrintException()
        raise e

if module == "Paste":

    excel = GetGlobals("excel")
    xls = excel.file_[excel.actual_id]
    sheet = GetParams("sheet_name")
    values = GetParams("values")
    cells = GetParams("cells")

    try:

        wb = xls['workbook']
        if not sheet in [sh.name for sh in wb.sheets]:
            raise Exception(f"The name {sheet} does not exist in the book")
        wb.sheets[sheet].select()
        wb.sheets[sheet].api.Range(cells).PasteSpecial(Paste=12 if values else 7)

    except Exception as e:
        print("\x1B[" + "31;40mError\x1B[" + "0m")
        PrintException()
        raise e

if module == "focus":
    try:
        from time import sleep
        from uiautomation import uiautomation as auto

        excel = GetGlobals("excel")
        xls = excel.file_[excel.actual_id]
        wb = xls['workbook']
        sleep(1)
        print(wb.app.impl.hwnd)
        name = f'\u202a{xls["path"].split(os.sep)[-1]}\u202c  -  Excel'
        control = auto.TitleBarControl(Name=name)
        control.SetFocus()
    except Exception as e:
        if e.text != 'Error no especificado':
            print("\x1B[" + "31;40mError\x1B[" + "0m")
            PrintException()
            raise e

if module == "remove_duplicate":
    sheet = GetParams("sheet_name")
    range_ = GetParams("range")
    column = GetParams("column")
    with_header = GetParams("header")
    excel = GetGlobals("excel")
    xls = excel.file_[excel.actual_id]

    try:
        wb = xls['workbook']
        if not sheet in [sh.name for sh in wb.sheets]:
            raise Exception(f"The name {sheet} does not exist in the book")
        sheet_selected = wb.sheets[sheet]
        sheet_selected.select()
        column = eval(column) if column.startswith("[") else [column]
        column_choice = []
        for col in column:
            column_choice.append(wb.sheets[sheet].api.Range(col + "1").column)
        sheet_selected.api.Range(range_).RemoveDuplicates(Columns=column_choice, Header=int(bool(with_header)))
    except Exception as e:
        print("\x1B[" + "31;40mError\x1B[" + "0m")
        PrintException()
        raise e

if module == "save_mac":
    excel = GetGlobals("excel")
    path_file = GetParams('path_file')
    xls = excel.file_[excel.actual_id]

    wb = xls['workbook']
    if not path_file:
        path_file = xls["path"]
    wb.save(path_file)

if module == "copyMove":

    excel = GetGlobals("excel")
    sheet1 = GetParams('sheet_name1')
    sheet2 = GetParams('sheet_name2')
    book = GetParams("book")
    copy_ = GetParams("copy")
    xls = excel.file_[excel.actual_id]

    wb = xls['workbook']
    try:
        if not sheet1 in [sh.name for sh in wb.sheets]:
            raise Exception(f"The name {sheet1} does not exist in the book")
        sheet_selected = wb.sheets[sheet1]
        # sheet_selected.select()
        if not sheet2:
            sheet2 = "tmp"

        if book:
            wb2 = wb.app.books.open(book)
            if sheet2 == "tmp":
                wb2.sheets.add(name=sheet2, after=wb2.sheets[-1])
            destiny = wb2.api.Sheets(sheet2)
        else:
            destiny = wb.api.Sheets(sheet2)
            wb.sheets.add(name=sheet2, after=wb.sheets[-1])

        if copy_:
            sheet_selected.api.Copy(Before=destiny)
        else:
            sheet_selected.api.Move(Before=destiny)

        try:
            wb2.sheets["tmp"].select() if book else wb.sheets["tmp"].select()
            wb2.sheets["tmp"].delete() if book else wb.sheets["tmp"].delete()
            wb2.save()
            wb2.close()
        except:
            pass


    except Exception as e:
        print("\x1B[" + "31;40mError\x1B[" + "0m")
        PrintException()
        raise e

if module == "exportPDF":
    excel = GetGlobals("excel")
    xls = excel.file_[excel.actual_id]
    wb = xls['workbook']

    path_file = GetParams('path_file')
    option = GetParams('option')
    check_zoom = GetParams('check_zoom')
    check_tall = GetParams('check_tall')
    check_wide = GetParams('check_wide')

    excel = GetGlobals("excel")
    xls = excel.file_[excel.actual_id]
    wb = xls['workbook']
    sh = xls['sheet']

    try:
        if option:
            if option == "all":
                sh.autofit()

            if option == "columns":
                sh.autofit('c')

            if option == "rows":
                sh.autofit('r')

        if check_zoom:
            sh.api.PageSetup.Zoom = False
        if check_tall:
            sh.api.PageSetup.FitToPagesTall = False
        if check_wide:
            sh.api.PageSetup.FitToPagesWide = 1

        wb.api.ActiveSheet.ExportAsFixedFormat(0, path_file.replace("/", os.sep))

    except Exception as e:
        PrintException()
        raise e

if module == "ImportForm":
    form_path = GetParams('form_path')
    excel = GetGlobals("excel")
    xls = excel.file_[excel.actual_id]
    wb = xls['workbook']

    try:
        wb.api.VBProject.VBComponents.Import(form_path)

    except Exception as e:
        PrintException()
        raise e


if module == "GetCells":
    sheet = GetParams("sheet")
    range_ = GetParams("range")
    result = GetParams("var_")
    extends = GetParams("more_data")
    excel = GetGlobals("excel")

    xls = excel.file_[excel.actual_id]
    wb = xls['workbook']
    try:
        if not sheet in [sh.name for sh in wb.sheets]:
            raise Exception(f"The name {sheet} does not exist in the book")

        sheet_selected_api = wb.sheets[sheet].api
        filtered_cells = sheet_selected_api.Range(range_).SpecialCells(12)
        cell_values = []

        for r in filtered_cells.Address.split(","):
            range_cell = []
            for ro in wb.sheets[sheet].api.Range(r).Rows:
                if isinstance(ro.Value, list) or isinstance(ro.Value, tuple):
                    cells = []
                    for cell in ro.Cells:
                        if isinstance(cell.Value, datetime.datetime):
                            cells.append(get_date_with_format(cell.Value2))
                        else:
                            cells.append(cell.Value2)

                    range_cell.append(cells)
                else:
                    range_cell.append([ro.Value])
            try:
                extends = eval(extends)
            except TypeError:
                pass
            if extends:
                info = {"range": r.replace("$", ""), "data": range_cell}
                cell_values.append(info)
            else:
                cell_values = cell_values + range_cell if len(cell_values) > 0 else range_cell

        if result:
            SetVar(result, cell_values)

    except Exception as e:
        print("\x1B[" + "31;40mError\x1B[" + "0m")
        PrintException()
        raise e

if module == "Replace":
    sheet = GetParams("sheet")
    range_ = GetParams("range")
    what = GetParams("what")
    replacement = GetParams("replace")
    excel = GetGlobals("excel")
    xls = excel.file_[excel.actual_id]
    wb = xls['workbook']

    try:
        wb.sheets[sheet].range(range_).api.Replace(what, replacement)

    except Exception as e:
        print("\x1B[" + "31;40mError\x1B[" + "0m")
        PrintException()
        raise e

if module == "Order":
    sheet_name= GetParams("sheet")
    range_ = GetParams("range")
    column = GetParams("column")
    order = GetParams("order")
    excel = GetGlobals("excel")
    xls = excel.file_[excel.actual_id]
    wb = xls['workbook']

    try:
        if not sheet_name in [sh.name for sh in wb.sheets]:
            raise Exception(f"The name {sheet_name} does not exist in the book")
        sheet = wb.sheets[sheet_name]
        if order:
            order = int(order)
        else:
            order = 1
        sheet.api.Range(range_).Sort(Key1=sheet.api.Range(column), Order1=order)

    except Exception as e:
        print("\x1B[" + "31;40mError\x1B[" + "0m")
        PrintException()
        raise e

if module == "refreshAll":
    excel = GetGlobals("excel")
    xls = excel.file_[excel.actual_id]
    wb = xls['workbook']

    try:
        wb.api.RefreshAll()
    except Exception as e:
        PrintException()
        raise e


if module == "find":
    sheet_name= GetParams("sheet")
    range_ = GetParams("range")
    text = GetParams("text")
    var_ = GetParams("var_")

    excel = GetGlobals("excel")
    xls = excel.file_[excel.actual_id]
    wb = xls['workbook']

    try:
        if not sheet_name in [sh.name for sh in wb.sheets]:
            raise Exception(f"The name {sheet_name} does not exist in the book")
        sheet = wb.sheets[sheet_name]
        result = sheet.api.Range(range_).Find(text)
        print(sheet.api.Range(range_).Find, result)
        if result:
            SetVar(var_, result.address)


    except Exception as e:
        print("\x1B[" + "31;40mError\x1B[" + "0m")
        PrintException()
        raise e

if module == "LockCells":
    sheet_name= GetParams("sheet")
    range_ = GetParams("range")
    locked = GetParams("locked")

    excel = GetGlobals("excel")
    xls = excel.file_[excel.actual_id]
    wb = xls['workbook']

    try:
        if not sheet_name in [sh.name for sh in wb.sheets]:
            raise Exception(f"The name {sheet_name} does not exist in the book")

        locked = eval(locked) if locked else False

        sheet = wb.sheets[sheet_name]
        result = sheet.api.Range(range_).Locked = locked

    except Exception as e:
        print("\x1B[" + "31;40mError\x1B[" + "0m")
        PrintException()
        raise e

if module == "add_chart":

    sheet_name= GetParams("sheet")
    range_ = GetParams("range")
    cell = GetParams("cell")
    type_ = GetParams("type")

    excel = GetGlobals("excel")
    xls = excel.file_[excel.actual_id]
    wb = xls['workbook']
    try:
        if not sheet_name in [sh.name for sh in wb.sheets]:
            raise Exception(f"The name {sheet_name} does not exist in the book")

        if not type_ or not bool(type_):
            raise Exception("The type of chart has not been selected")

        type_ = int(type_)
        # compatibilidad con versión anterior (antes usaba la api para windows)
        types_charts = {
            4: "line",
            5: "pie",
            51: "column_clustered",
            72: "xy_scatter_smooth",
            -4169: "xy_scatter"

        }
        if type_ in types_charts:
            type_ = types_charts[type_]

        print(type_)
        sheet = wb.sheets[sheet_name]

        cell = sheet.range(cell)
        range_ = sheet.range(range_)

        active_chart = sheet.charts.add(cell.left, cell.top)
        active_chart.set_source_data(range_)
        active_chart.chart_type = type_

    except Exception as e:
        print("\x1B[" + "31;40mError\x1B[" + "0m")
        PrintException()
        raise e


if module == 'removePass':

    try:
        excel_file = GetParams('excel_file')
        new_excel_file = GetParams('new_excel_file')
        if new_excel_file:
            new_excel_file = new_excel_file.replace('/', '\\')
        excel_file = excel_file.replace('/', '\\')
        pass_excel = GetParams('pass_excel')

        import win32com.client

        if not new_excel_file:
            new_excel_file = excel_file

        xcl = win32com.client.Dispatch("Excel.Application")
        wb = xcl.Workbooks.Open(excel_file, False, False, None, pass_excel)
        xcl.DisplayAlerts = False
        wb.SaveAs(f'{new_excel_file}', None, '', '')

        excel = GetGlobals("excel")
        excel.actual_id = excel.id_default

        excel.file_[excel.actual_id] = {}
        excel.file_[excel.actual_id]['workbook'] = xw.Book(new_excel_file)
        excel.file_[excel.actual_id]['app'] = excel.file_[excel.actual_id]['workbook'].app
        excel.file_[excel.actual_id]['sheet'] = excel.file_[excel.actual_id]['workbook'].sheets[0]
        excel.file_[excel.actual_id]['path'] = new_excel_file

        xls = excel.file_[excel.actual_id]
        wb = xls['workbook']
        xw.books.active.close()
    except Exception as e:
        PrintException()
        raise e


if module == "insertImage":
    excel = GetGlobals("excel")
    sheet_name = GetParams("sheet")
    image_path = GetParams("image_path")
    image_path = image_path.replace("/", os.sep)
    cell_position = GetParams("cell_position")
    try:
        xls = excel.file_[excel.actual_id]
        wb = xls['workbook']
        if not sheet_name in [sh.name for sh in wb.sheets]:
            raise Exception(f"The name {sheet_name} does not exist in the book")
        sheet = wb.sheets[sheet_name]
        cell = sheet.range(cell_position)
        sheet.pictures.add(image_path, top=cell.top, left=cell.left)
    except Exception as e:
        print("\x1B[" + "31;40mAn error occurred\x1B[" + "0m")
        PrintException()
        raise e

if module == "ExportChart":
    excel = GetGlobals("excel")
    sheet_name = GetParams("sheet")
    index = GetParams("index")
    path = GetParams("path")
    try:
        xls = excel.file_[excel.actual_id]
        wb = xls['workbook']
        if not sheet_name in [sh.name for sh in wb.sheets]:
            raise Exception(f"The name {sheet_name} does not exist in the book")
        sheet = wb.sheets[sheet_name]
        chart = sheet.api.ChartObjects(int(index))
        chart = chart.Chart
        chart.Export(Filename=path, FilterName="PNG")

    except Exception as e:
        print("\x1B[" + "31;40mAn error occurred\x1B[" + "0m")
        PrintException()
        raise e

if module == "headless":
    excel = GetGlobals("excel")
    path = GetParams("path")
    id_ = GetParams("id")

    try:
        app = xw.App(add_book=False, visible=False)

        if path:
           wb = app.books.open(path)
        else:
            wb = app.books.add()
            path = ""
        print(app.books)

        excel.actual_id = excel.id_default

        if id_:
            excel.actual_id = id_
        excel.file_[excel.actual_id] = {}
        excel.file_[excel.actual_id]['workbook'] = wb
        excel.file_[excel.actual_id]['app'] = excel.file_[excel.actual_id]['workbook'].app
        excel.file_[excel.actual_id]['sheet'] = excel.file_[excel.actual_id]['workbook'].sheets[0]
        excel.file_[excel.actual_id]['path'] = path

    except Exception as e:
        print("\x1B[" + "31;40mAn error occurred\x1B[" + "0m")
        PrintException()
        raise e
        