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
import os
import platform
import xlwings as xw
from xlwings.constants import InsertShiftDirection
import pandas as pd

module = GetParams("module")

if module == "CellColor":
    print("uno")
    excel = GetGlobals("excel")

    range_ = GetParams("range")
    color = GetParams("color")

    if color == "red":
        rgb = (255,0,0)
        print("dos")
    elif color == "blue":
        rgb = (0,0,255)
    elif color == "green":
        rgb = (0,255,0)
    elif color == "grey":
        rgb = (130,130,130)
    else:
        rgb = (255,255,0)

    try:
        print("En el try")
        xls = excel.file_[excel.actual_id]


        # wb = xls['workbook']
        #         # print(wb)
        xw.Range(range_).color= rgb

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

    if len(rango) == 1:
        rango = rango + ':' + rango

    if formato == "text":
        xw.sheets[hoja].range(rango).number_format = '@'

    if formato == "number_":
        xw.sheets[hoja].range(rango).number_format = '0'

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

if module == "createSheet":
    hoja = GetParams("sheet_name")

    res = [a.name for a in xw.sheets]
    last = res[-1]

    xw.sheets.add(name=hoja, after=last)


if module == "copy_other":
    try:
        excel1 = GetParams("excel1")
        excel2 = GetParams("excel2")
        hoja1 = GetParams("sheet_name1")
        hoja2 = GetParams("sheet_name2")
        rango1 = GetParams("cell_range1")
        rango2 = GetParams("cell_range2")
        platform_ = platform.system()

        app = xw.App(visible=False)
        wb1 = xw.Book(excel1)
        wb2 = xw.Book(excel2)

        my_values = wb1.sheets[hoja1].range(rango1).options(ndim=2).value

        wb2.sheets[hoja2].range(rango2).value = my_values

        if platform_ == 'Windows':
            wb2.save(excel2)
            wb2.close()
            # wb1.close()
            
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
                if len(fila_) == 1:

                    if tipo == "down_":
                        fila_ = int(fila_) + 1
                        # print(fila_)
                        fila = str(fila_) + ':' + str(fila_)

                        xw.sheets[hoja].range(fila).api.Insert(InsertShiftDirection.xlShiftDown)

                    if tipo == "up_":
                        fila = str(fila_) + ':' + str(fila_)
                        xw.sheets[hoja].range(fila).api.Insert(InsertShiftDirection.xlShiftDown)

                else:
                    if len(fila_) > 1:
                        if tipo == "down_":
                            fila = fila_.split(':')
                            f1 = fila[0]
                            f1 = int(f1) + 1
                            f2 = fila[1]
                            f2 = int(f2) + 1
                            fila = str(f1) + ':' + str(f2)
                            # fila_ = int(fila_) + 1
                            xw.sheets[hoja].range(fila).api.Insert(InsertShiftDirection.xlShiftDown)
                        if tipo == "up_":
                            xw.sheets[hoja].range(fila_).api.Insert(InsertShiftDirection.xlShiftDown)

            else:
                if len(fila_) == 1:

                    if tipo == "down_":
                        fila_ = int(fila_) + 1
                        # print(fila_)
                        fila = str(fila_) + ':' + str(fila_)

                        xw.sheets[hoja].api.rows[fila].insert_into_range()

                    if tipo == "up_":
                        fila = str(fila_) + ':' + str(fila_)
                        xw.sheets[hoja].api.rows[fila].insert_into_range()

                if len(fila_) > 1:
                    if tipo == "down_":
                        fila = fila_.split(':')
                        f1 = fila[0]
                        f1 = int(f1) + 1
                        f2 = fila[1]
                        f2 = int(f2) + 1
                        fila = str(f1) + ':' + str(f2)
                        #fila_ = int(fila_) + 1
                        xw.sheets[hoja].api.rows[fila].insert_into_range()
                    if tipo == "up_":
                        xw.sheets[hoja].api.rows[fila_].insert_into_range()

        if opcion_ == "delete_":
            if platform_ == 'Windows':
                if len(fila_) == 1:
                    fila = str(fila_) + ':' + str(fila_)
                    xw.Range(fila).api.Delete()
                else:
                    xw.Range(fila_).api.Delete()

            else:
                if len(fila_) == 1:
                    fila = str(fila_) + ':' + str(fila_)
                    xw.Range(fila).api.delete()
                else:
                    xw.Range(fila_).api.delete()

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

                if len(str(col_)) == 1:
                    col = str(col_) + ':' + str(col_)
                    xw.sheets[hoja].range(col).api.Insert(InsertShiftDirection.xlShiftToRight)

                else:
                    xw.sheets[hoja].range(col_).api.Insert(InsertShiftDirection.xlShiftToRight)

            else:
                if len(str(col_)) == 1:
                    col = str(col_) + ':' + str(col_)
                    xw.sheets[hoja].api.columns[col].insert_into_range()

                else:
                    xw.sheets[hoja].api.columns[col_].insert_into_range()

        if opcion_ == "delete_":
            if platform_ == 'Windows':
                if len(col_) == 1:
                    col = str(col_) + ':' + str(col_)
                    xw.Range(col).api.Delete()
                else:
                    xw.Range(col_).api.Delete()

            else:
                if len(col_) == 1:
                    col = str(col_) + ':' + str(col_)
                    xw.Range(col).api.delete()
                else:
                    xw.Range(col_).api.delete()
    except:
        PrintException()


if module == "csvToxlsx":
    csv_path = GetParams("csv_path")
    xlsx_path = GetParams("xlsx_path")

    if not csv_path or not xlsx_path:
        raise Exception("Falta una ruta")

    df = pd.read_csv(csv_path)
    df.to_excel(xlsx_path, index=None)