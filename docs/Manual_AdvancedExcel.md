



# Opciones avanzadas para Excel
  
Aplique filtros automaticos y avanzados, de formato a las celdas, añada o elimine hojas, filas o columnas, exporte a diferentes formatos de archivo, desbloquee y vuelva a bloquear hojas, copie y realice pegado especial y mas con sus archivos de Excel.   

*Read this in other languages: [English](Manual_AdvancedExcel.md), [Português](Manual_AdvancedExcel.pr.md), [Español](Manual_AdvancedExcel.es.md)*
  
![banner](imgs/Banner_AdvancedExcel.png)
## Como instalar este módulo
  
Para instalar el módulo en Rocketbot Studio, se puede hacer de dos formas:
1. Manual: __Descargar__ el archivo .zip y descomprimirlo en la carpeta modules. El nombre de la carpeta debe ser el mismo al del módulo y dentro debe tener los siguientes archivos y carpetas: \__init__.py, package.json, docs, example y libs. Si tiene abierta la aplicación, refresca el navegador para poder utilizar el nuevo modulo.
2. Automática: Al ingresar a Rocketbot Studio sobre el margen derecho encontrara la sección de **Addons**, seleccionar **Install Mods**, buscar el modulo deseado y presionar install.  



## Como usar este modulo
Para usar este modulo debe tener Microsoft Excel instalado.


## Descripción de los comandos

### Abrir sin alertas
  
Abre un archivo sin mostrar carteles de alerta.
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Ruta archivo XLSX|Ruta del archivo xlsx que se quiere abrir|Archivo.XLSX|
|Password (opcional)|Contraseña del archivo xlsx|P@ssW0rd|
|Identificador (opcional)|Nombre o identificador para el archivo que se abrirá. Se utiliza cuando se necesita abrir más de un excel. Por defecto es *default*|id|
|Asignar resultado a variable|Variable donde se almacenara el resultado|id|

### Buscar y conectar
  
Busca un excel abierto y se conecta a este.
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Nombre del archivo XLSX abierto||Archivo.XLSX|
|Identificador (opcional)|Nombre o identificador para el archivo que se abrirá. Se utiliza cuando se necesita abrir más de un excel. Por defecto es *default*|excel1|

### Maximizar
  
Maximizar Ventana de Excel
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |

### Opciones de calculo
  
Selecciona la manera en que se ejecuta el calculo de formulas en el libro.
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Opciones de calculo|Seleccionar la manera de cálculo del libro.||
|Calcular ahora|Si se marca esta casilla, se calculan las formulas del libro|True|

### Leer celdas
  
Lee una celda o rango de celdas
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja|Nombre de la hoja donde se encuentran los datos|Hoja 1|
|Ingrese celdas |Celda o Rango de celdas. La sintaxis debe ser la misma de excel (A1 o A1B1) |A1:B5|
|Formato|Seleccionar el formato a traer las celdas que contengan fechas. Seleccione custom para adicionar un formato personalizado|dd-mm-yy|
|Formato personalizado |Formato personalizado. Doc https//docs.python.org/3/library/datetime.html#strftime-and-strptime-format-codes|'%m/%d/%y %I:%M %p'|
|Asignar resultado a variable|Nombre de la variable donde guardar el resultado|cells|

### Convertir fecha serial
  
Convierte una fecha numero serial excel a un formato de fecha especifico
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Fecha serial |Numero serial de excel que representa una fecha especifica, siendo 1 = 01/01/1900|44927|
|Formato de salida |Formato de fecha al cual convertir la fecha serial|%d/%m/%y|
|Asignar resultado a variable |Nombre de la variable donde guardar el resultado|output_date|

### Contar Columnas
  
Cuenta el número de columnas del excel abierto. Se requiere que el excel esté guardado para tomar los últimos cambios
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja|Nombre de la hoja donde se encuentran los datos|Hoja 1|
|Obtener nombre de columna|Si se marca esta casilla, devolverá la letra de la última columna|True|
|Asignar resultado a variable|Nombre de la variable donde guardar el resultado|numero_columnas|

### Contar Filas
  
Cuenta todas las filas o dentro de un rango.
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja|Nombre de la hoja donde se encuentran los datos|Hoja 1|
|Contar todas las filas|Opción para contar todas las filas.||
|Columna|Columna donde se contará la cantidad de filas|C|
|Asignar resultado a variable|Nombre de la variable donde guardar el resultado|numero_filas|

### Ocultar
  
Oculta una o varias filas, o una o varias columnas.
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja|Nombre de la hoja|Hoja1|
|Rango|Para un rango de filas utilizar números separados por dos puntos (13) Para rango de columnas utilizar letras(AB)|1:3|
|Asignar resultado a variable|Nombre de la variable donde guardar el resultado|res|

### Mostrar
  
Muestra una o varias filas, o una o varias columnas que estén ocultas
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja|Nombre de la hoja|Hoja1|
|Rango|Para un rango de filas utilizar números separados por dos puntos (13) Para rango de columnas utilizar letras(AB).|1:3|
|Asignar resultado a variable|Nombre de la variable donde guardar el resultado|res|

### Color celda
  
Cambia color de una celda o rango de celdas. Puedes seleccionar un valor por defecto o uno personalizado
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Ingrese celdas |Celda o Rango de celdas. La sintaxis debe ser la misma de excel (A1 o A1B1) |A1:B5|
|Hoja |Hoja del libro|Hoja1|
|Toda la hoja|Si se marca esta casilla, el color se aplicara a toda la hoja.||
|Ingrese color en RGB |Valores rgb del color que tendrá la celda o celdas|250,250,250|
|Seleccione color |Seleccione el color. Puede usar el campo anterior para personalizar|red|

### Obtener color de celda
  
Obtener el color de una celda. La función devolverá una lista con dos elementos: Color de fondo y Color de fuente en formato RGB.
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja |Hoja|Hoja1|
|Celda |Celda. La sintaxis debe ser la misma de excel (A1)|A1|
|Asignar a variable|Nombre de la variable donde guardar el resultado.|color|

### Obtener formatos de celda
  
Obtener el formato de una celda. La función devolverá un diccionario con las propiedades de la celda y el valor de cada una.
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja |Hoja|Hoja1|
|Celda |Celda. La sintaxis debe ser la misma de excel (A1)|A1|
|Asignar a variable|Nombre de la variable donde guardar el resultado.|color|

### Insertar Formula
  
Inserta formula sobre una celda 
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Ingrese hoja |Hoja|Hoja1|
|Ingrese celda |Celda o Rango de celdas. La sintaxis debe ser la misma de excel (A1 o A1B1) |A5|
|Escriba fórmula |Formula que se quiere insertar. Debe ser escrita en inglés. Recuerda usar *,* para separar los parámetros|=SUM(A1:A4)|
|No IIE|Si se marca esta casilla, permite enviar la formula sin IIE|True|

### Insertar Macro a Excel
  
Inserta una Macro a Excel 
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Ruta de la Macro|Ruta del archivo .bas que se quiere insertar|Macro.bas|

### Seleccionar y copiar Celdas
  
Selecciona y copia celdas en Excel
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja|Nombre de la hoja que se quiere automatizar|Hoja 1|
|Ingrese celdas a seleccionar|Celda o Rango de celdas para seleccionar. La sintaxis debe ser la misma de excel (A1 o A1B1) |A1:B3|
|Copiar|Al marcar la casilla, se copiarán los valores en el portapapeles|True|

### Obtener Celda Formato Moneda
  
Obtiene celdas con formato moneda
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja|Nombre de la hoja que se quiere automatizar|Hoja 1|
|Ingrese celdas a seleccionar|Celda o Rango de celdas. La sintaxis debe ser la misma de excel (A1 o A1B1) |A1:B3|
|Asignar a variable|Nombre de la variable donde guardar el resultado|variable|

### Obtener Celda Formato Fecha
  
Obtiene celdas con formato fecha
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja|Nombre de la hoja que se quiere automatizar|Hoja 1|
|Ingrese celdas a seleccionar|Celda o Rango de celdas. La sintaxis debe ser la misma de excel (A1 o A1B1) |A1:B3|
|Asignar a variable|Nombre de la variable donde guardar el resultado|variable|

### Copiar-Pegar
  
Copia un rango de celdas desde una hoja a otra 
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja origen |Nombre de la hoja que se quiere automatizar|Sheet1|
|Rango a copiar |Celda o Rango de celdas a copiar. La sintaxis debe ser la misma de excel (A1 o A1B1) |A1:C4|
|Hoja destino |Nombre de la hoja de destino|Sheet2|
|Rango donde pegar|Celda o Rango de celdas donde pegar. La sintaxis debe ser la misma de excel (A1 o A1B1) |A1:C4|
|Opción de Pegado|Seleccionar tipo de pegado para la celda o rango de celdas.|Opcion|
|Operación de Pegado|Seleccionar operación de pegado para la celda o rango de celdas.|Operación|
|Saltar Blancos|Evita reemplazar valores en el área de pegado cuando se producen celdas en blanco en el área de copia cuando se selecciona esta casilla.||
|Transponer|Gira el contenido de celdas copiadas al pegar. Los datos en filas se pegarán en columnas y viceversa.||

### Formatear Celda
  
Formatear Celda
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Nombre de Hoja|Nombre de la hoja que se quiere automatizar|Sheet1|
|Rango a formatear |Celda o Rango de celdas a formatear. La sintaxis debe ser la misma de excel (A1 o A1B1) |A1:C4|
|Formato|Se debe seleccionar el tipo de formato para la celda. Seleccione custom para adicionar un formato personalizado|dd-mm-yy|
|Formato personalizado |Formato personalizado. Doc https//support.microsoft.com/es-es/office/revisar-las-instrucciones-para-personalizar-un-formato-de-n%C3%BAmero-c0a1d1fa-d3f4-4018-96b7-9c9354dd99f5|00000|
|Texto a Valor|||

### Borrar contenido
  
Borra fórmulas y valores del rango seleccionado, manteniendo el formato.
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja|Nombre de la hoja donde se encuentran los datos|Hoja 1|
|Celda o Rango de celdas|Rango que contiene los datos a alinear|A1:D7|

### Crear Hoja
  
Añade una hoja al final
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Nombre de la hoja|Nombre de la hoja que se quiere crear|Sheet2|
|Despues de|La hoja se creará al lado de la hoja indicada en este campo|Hoja1|

### Eliminar Hoja
  
Elimina una hoja
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Nombre de la hoja|Nombre de la hoja que se quiere borrar|Sheet2|
|Asignar resultado a variable|Nombre de la variable donde guardar el resultado|Variable|

### Copiar de un Excel a otro
  
Copia el rango de un archivo de Excel a otro. Indicando la ruta del archivo, abrirá el excel para copiar o pegar los datos. Si ingresas el id de un excel abierto, usará esa instancia para copiar o pegar.
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Excel origen (opcional)|Ruta del archivo excel de origen|Ruta archivo origen:|
|Identificador (opcional)|Nombre o ID del archivo de código abierto.|id|
|Hoja origen|Nombre de la hoja de origen|Sheet1|
|Rango a copiar|Celda o Rango de celdas a copiar. La sintaxis debe ser la misma de excel (A1 o A1B1) |A1:D7|
|Excel destino|Ruta del archivo excel de destino|Ruta archivo destino:|
|Abrir normalmente|Si esta casilla está marcada, el archivo de destino se abre normalmente manteniendo los datos, formatos y objetos. De lo contrario, solo recupera datos.|True|
|Solo valores|Si esta casilla es seleccionada, copiará solo los valores|True|
|Hoja destino|Nombre de la hoja donde se copiará|Sheet1|
|Rango donde pegar (Opcional)|Columna, Celda o Rango de celdas donde pegar. La sintaxis debe ser la misma de excel (A, A1 o A1B1) |A1:D7|

### Insertar/Eliminar Fila
  
Inserta o elimina una fila
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Opción|Seleccione Add para agregar una fila o Delete para borrar|Add|
|Nombre de Hoja|Nombre de la hoja donde agregar la fila|Sheet|
|Número Fila|Indique la fila o filas que se quieren agregar o eliminar|2|
|Dónde Insertar|Indique donde agregar o eliminar la fila|A1:D7|

### Insertar/Eliminar Columna
  
Inserta o elimina una columna
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Opción|Seleccione Add para agregar una columna o Delete para borrar||
|Nombre de Hoja|Nombre de la hoja donde se encuentran los datos|Sheet|
|Rango|Indique la columna o columnas que se quieren agregar o eliminar|B|

### Convertir CSV a XLSX
  
Convierte un documento CSV a formato XLSX
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Ruta archivo CSV|Ruta del archivo csv que se quiere convertir||
|Delimitador|Separador del archivo csv||
|Tiene cabeceras?|Marcar esta casilla si el csv tiene cabeceras|True|
|Codificación|Escriba el tipo de codificación del archivo. Por defecto es latin-1|latin-1|
|Ruta archivo XLSX|Ruta del archivo xlsx donde guardar|file.xlsx|

### Exportar a JSON
  
Exporta un array de datos a un archivo JSON
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Datos|Datos a exportar|[['header1', 'header2', 'header3', 'header4', 'header5', 'header6'], ['data11', 'data12', 'data13', 'data14', 'data15', 'data16']]|
|Ruta archivo json|Ruta del archivo json donde guardar la conversión|C:/Users/User/Desktop/file.json|

### (Deprecado) Convertir XLSX a CSV
  
Convierte un documento XLSX a CSV
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Ruta archivo XLSX|Ruta del archivo xlsx que se quiere convertir|C:/Users/User/Desktop/file.xlsx|
|Delimitador|Separador del archivo csv|,|
|Nombre de la hoja|Nombre de la hoja donde se encuentran los datos|Sheet0|
|Ruta archivo CSV|Ruta del archivo csv donde guardar la conversión|C:/Users/User/Desktop/file.csv|

### Convertir XLSX a CSV
  
Convierte un documento XLSX a CSV
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Ruta archivo XLSX|Ruta del archivo xlsx que se quiere convertir|C:/Users/User/Desktop/file.xlsx|
|Delimitador|Separador del archivo csv|,|
|Nombre de la hoja|Nombre de la hoja donde se encuentran los datos|Sheet0|
|Ruta archivo CSV|Ruta del archivo csv donde guardar la conversión|C:/Users/User/Desktop/file.csv|

### Convertir XLS a XLSX
  
Convierte un documento XLS a XLSX
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Ruta archivo XLS|Ruta del archivo xls que se quiere convertir|C:\Users\User\Desktop\file.xls|
|Ruta archivo XLSX|Ruta donde se guardará el archivo xlsx|C:\Users\User\Desktop\new_file.xlsx|

### Obtener celda activa
  
Obtener fila y columna de una celda activa
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Asignar resultado a variable|Nombre de la variable donde guardar el resultado|Variable|

### Actualizar tabla dinámica
  
Actualiza una tabla dinámica. ¡Obsoleto! Use el módulo PivotTableExcel
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja |Nombre de la hoja donde se encuentra la tabla|Hoja 1|
|Nombre de la tabla dinámica |Nombre de la tabla dinámica que se actualizará|Name: |

### Ajustar celdas
  
Ajusta, une, agrupa y desagrupa un rango de celdas. Puedes agrupar/desagrupar por filas o columnas
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja|Nombre de la hoja donde se encuentran los datos|Hoja 1|
|Rango a ajustar|Celda o Rango de celdas a ajustar. La sintaxis debe ser la misma de excel (A1 o A1B1) |A1:D7|
|Autofit|Ajusta automaticamente las celdas para que se visualicen los datos||
|Agrupar filas|Al marcar este checkbox, se agruparán las filas en el rango seleccionado||
|Agrupar columnas|Al marcar este checkbox, se agruparán las columnas en el rango seleccionado||
|Desagrupar filas|Al marcar este checkbox, se desagruparán las filas en el rango seleccionado||
|Desagrupar columnas|Al marcar este checkbox, se desagruparán las columnas en el rango seleccionado||
|Unir celdas|Al marcar este checkbox, se uniran las celdas en el rango seleccionado||
|Separar celdas|Al marcar este checkbox, se separaran las celdas en el rango seleccionado||
|Nivel de fila|Al marcar esta casilla se mostrará el número especificado de niveles de fila|2|
|Rango de columna|Al marcar esta casilla se mostrará el número especificado de niveles de columna|2|
|Ancho de columna|Ancho al que se ajustara la columna|20|
|Altura de Fila|Altura a la que se ajustara la fila|20|

### Obtener Formula
  
Obtiene la formula sobre una celda 
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Ingrese celda |Celda donde está la formula. La sintaxis debe ser la misma de excel (A1 o A1B1) |A5|
|Asignar resultado a variable|Nombre de la variable donde guardar el resultado|Variable|

### Agregar Filtro Automático
  
Agrega filtro automático a una tabla excel
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja |Nombre de la hoja donde se encuentran los datos|Hoja 1|
|Columnas |Columna o Rango de columnas. La sintaxis debe ser la misma de excel (A o AB) |A:E |

### Eliminar Filtro Automático
  
Eliminar el filtro automático de una hoja de Excel
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja |Nombre de la hoja donde se encuentra el filtro a quitar|Hoja 1|

### Borrar Filtro
  
Borra todos los filtros aplicados sobre una hoja de Excel
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja |Nombre de la hoja donde se encuentran los datos filtrados|Hoja 1|

### Filtrar
  
Filtrar una tabla de excel según el valor relativo, contenido exacto, color de fondo o color de letra de las celdas. *Ejemplos según tipo de filtro: xlAnd ['>=10'] or ['>=10', '<=20'] | xlOr ['<=10', '>=20'] | xlFilterValues ['10','20', '30'] | xlFilterCellColor (255,0,0) | xlFilterFontColor (255,0,0)*
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja |Nombre de la hoja donde se encuentran los datos|Hoja1|
|Inicio de tabla |Columna donde comienza la tabla que se filtrará|A |
|Columna |Columna donde agregar el filtro|A |
|Filtro |Valor o lista de valores, filtro de un criterio o lista de dos items para doble criterio (ej valor entre A y B). Use "=" para encontrar campos en blanco, "<>" para celdas no vacías y negación de datos.|['>=10'] or ['>=10', '<=20'], ['10','20', '30'] or (255,0,0)|
|Tipo de filtro |Tipo de filtro a aplicar.||

### Filtrar por Fecha
  
Filtra una tabla por el día, mes o año de una fecha indicada
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja |Nombre de la hoja donde se encuentran los datos|Hoja1|
|Inicio de tabla |Columna donde comienza la tabla que se filtrará|A |
|Columna |Columna donde agregar el filtro|A |
|Fecha Filtro |Fecha o lista de fechas para establecer como criterio de filtro|18/04/2024|
|Tipo de filtro |Tipo de filtro a aplicar.||

### Filtro avanzado
  
Filtra a una tabla excel
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja |Nombre de la hoja donde se encuentran los datos|Hoja1|
|Rango de tabla |Rango de la tabla que se filtrará|A1:G500 |
|Rango de criterios  |Rango con los criterios del filtro a aplicar|A1:B4 |
|Solo registros únicos|||
|Copiar a otro lugar|Pega la tabla resultante en la celda de destino||
|Destino  |Celda donde pegar la tabla resultado del filtro|J1 |

### Remover Filtros
  
Eliminar filtros y mostrar todos los datos
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja |Nombre de la hoja donde se encuentran los datos|Hoja1|

### Renombrar hoja
  
Cambia el nombre a una hoja de excel
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja |Nombre de la hoja a renombrar|Hoja 1|
|Nuevo nombre |Nuevo Nombre de la hoja|nuevo_nombre|

### Formato de texto
  
Cambia la alineacion Horizontal o Vertical de los valores en un rango de celdas
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja|Nombre de la hoja donde se encuentran los datos|Hoja 1|
|Celda o Rango de celdas|Rango que contiene los datos a alinear|A1:D7|
|Alineacion Horizontal|Selector que contiene las opciones de alineacion horizontal||
|Alineacion Vertical|Selector que contiene las opciones de alineacion vertical||

### Estilo Celda
  
Este comando modifica el formato de la celda o rango de celdas seleccionado. Puedes cambiar la fuente y los bordes
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Nombre de Hoja|Nombre de la hoja que se quiere automatizar|Sheet1|
|Rango a formatear |Celda o Rango de celdas a formatear. La sintaxis debe ser la misma de excel (A1 o A1B1) |A1:C4|
|Borde|Borde de la celda que se quiere formatear|Contour|
|Estilo|Estilo del borde de la celda que se quiere formatear|_ _ _ _ _ _ _ _ _ _ _|
|Tamaño de fuente |Tamaño de la fuente de la celda|20|
|Negrita|Marcar esta casilla para cambiar la fuente a negrita|True|
|Cursiva|Marcar esta casilla para cambiar la fuente a cursiva|True|
|Subrayar|Marcar esta casilla para cambiar la fuente a subrayado|True|
|Ajustar Texto|Marcar esta casilla para ajustar el texto en el rango especificado|True|
|Alineación Horizontal|Tipo de alineado horizontal de la celda que se quiere formatear|Alignment|
|Alineación Vertical|Tipo de alineado verticaal de la celda que se quiere formatear|Alignment|

### Pegar en Celdas
  
Pega datos en celdas en Excel
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja|Nombre de la hoja que se quiere automatizar|Hoja 1|
|Ingrese celdas donde pegar|Celda o Rango de celdas donde pegar. La sintaxis debe ser la misma de excel (A1 o A1B1) |A1:B3|
|Solo valores|Si esta casilla es seleccionada, se pegarán solo los valores|True|

### Deshabilitar modo Copiar/Cortar
  
Deshabilitar el modo Cortar/Copiar del Excel activo
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |

### Eliminar duplicados
  
Ejecuta el comando eliminar duplicados de Excel
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja|Nombre de la hoja que se quiere automatizar|Hoja 1|
|Rango|Rango de celdas. La sintaxis debe ser la misma de excel (A1 o A1B1) |A1:B3|
|Columna |Indicar la columna donde se buscarán los duplicados|A / ['A', 'B']|
|Tiene cabeceras?|Marcar esta casilla si el excel tiene cabeceras|True|

### Exportar a PDF avanzado
  
Exporta Excel a PDF con opciones
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Guardar PDF|Ruta donde guardar el archivo .pdf|/Users/user/Desktop/excel.pdf|
|Hoja |Nombre de la hoja a exportar|Hoja 1|
|Todas las hojas|Al marcar la casilla, se exportaran todas las hojas||
|Ajuste Automatico|||
|Zoom|Ajusta el zoom del contenido de la planilla.||
|Ajustar Alto|Ajusta el alto del contenido de la planilla al numero de carillas definido.|1|
|Ajustar ancho|Ajusta el ancho del contenido de la planilla al numero de carillas definido.|1|
|Orientación|||

### Copiar-Mover Hoja
  
Copia o mueve una hoja
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja origen |Nombre de la hoja de origen|Sheet1|
|Mover/copiar antes de hoja |Nombre de la hoja donde se moverá|Sheet2|
|Excel destino|Ruta del archivo .xlsx donde mover o copiar la hoja|C:/ruta/al/excel.xlsx|
|Password (opcional)|Contraseña del archivo xlsx|P@ssW0rd|
|Copiar|Al marcar la casilla, se creará una copia de la hoja||

### Insertar Formulario
  
Inserta un Formulario a Excel 
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Ruta del Formulario|Ruta del archivo frm que se quiere insertar|Form.frm|

### Leer celdas filtradas
  
Lee todo el contenido de las celdas filtradas y aplica formato a los datos tipo fecha si se indica
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja |Nombre de la hoja donde se encuentran los datos|Hoja 1|
|Rango donde buscar |Celda o Rango de celdas. La sintaxis debe ser la misma de excel (A1 o A1B1) |A1:B100 |
|Asignar resultado a variable|Nombre de la variable donde guardar el resultado|Variable|
|Agregar formato especifico a datos almacenados como fecha |Dar formato especifico a datos almacenados como fecha|%m/%d/%Y, %H:%M:%S|
|Filas|||
|Datos extra|||

### Contar celdas filtradas
  
Cuenta solo las celdas filtradas
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja |Nombre de la hoja donde se encuentran los datos|Hoja 1|
|Rango donde buscar |Rango de columna filtrada (A1A100)|A1:A100 |
|Asignar resultado a variable|Nombre de la variable donde guardar el resultado|Variable|

### Reemplazar
  
Ejecuta la opción de reemplazar de excel
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja |Nombre de la hoja donde se encuentran los datos|Hoja 1|
|Rango donde buscar |Celda o Rango de celdas. La sintaxis debe ser la misma de excel (A1 o A1B1) |A1:B100 |
|Palabra a reemplazar|Palabra que se buscará para ser reemplazada|10/10/2020|
|Nueva palabra|Palabra que va a reemplazar a la anterior indicada|10-10-2020|

### Ordenar
  
Ejecuta la opción de reemplazar de excel
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja |Nombre de la hoja donde se encuentran los datos|Hoja 1|
|Rango donde buscar |Celda o Rango de celdas. La sintaxis debe ser la misma de excel (A1 o A1B1) |A1:B100 |
|Columna|Indicar la columna a ordenar|A1:A22|
|Tipo de orden |Indicar como se ordenará la columna|Ascending|

### Ordenar por múltiples niveles
  
Ordene una hoja de Excel por valor, estableciendo múltiples niveles
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja |Nombre de la hoja donde se encuentran los datos|Hoja 1|
|Rango a ordenar |Celda o Rango de celdas. La sintaxis debe ser la misma de excel (A1 o A1B1) |A1:B100 |
|Posee encabezados|Si se marca esta opción, tomara la primer fila del rango como encabezados.||
|Campos de orden|||

### Actualizar Todo
  
Actualiza todas las fuentes del libro
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |

### Buscar
  
Busca un texto en el rango indicado y retorna la celda donde se encuentra la primera coincidencia. Si no encuentra un valor, retornará vacío. Si el rango elta filtrado, la busqueda sere realizada sobre las celdas visibles.
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja |Nombre de la hoja donde se encuentran los datos|Hoja 1|
|Rango donde buscar |Celda o Rango de celdas. La sintaxis debe ser la misma de excel (A1 o A1B1) |A1:B100 |
|Texto a buscar|Texto que se quiere buscar en el excel|Lorem|
|Buscar en (opcional)|Indica el tipo de coincidencia deseada todo el texto buscado o dentro de cualquier parte (por defecto cualquier parte). ||
|Buscar dentro (opcional)|Indica dónde hacer la búsqueda valor, fórmula o comentario (predeterminado valor).||
|Distinguir mayúsculas y minúsculas|Si se marca esta casilla, buscara la cadena de texto diferenciando entre mayúsculas y minúsculas.||
|Encontrar todos|Si se marca esta casilla, devolvera un listado de todas las coincidencias.||
|Asignar resultado a variable|Nombre de la variable donde guardar el resultado|Variable|

### Encontrar dato
  
Devuelve la primera celda que coincida con el dato buscado
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja |Nombre de la hoja donde se encuentran los datos|Hoja 1|
|Rango donde buscar |Celda o Rango de celdas. La sintaxis debe ser la misma de excel (A1 o A1B1)|A1:B100 |
|Letra de columna con fechas (Opcional)|Letra de la columna/as que contienen fechas.|A,B|
|Formato de Fecha (Opcional)|Formato de la fecha a buscar.|%d/%m/%Y|
|Texto a buscar|Texto que se quiere buscar en el excel|Lorem|
|No distinguir mayúsculas y minúsculas|Si se marca esta casilla, buscara la cadena de texto sin diferencias entre mayúsculas y minúsculas.||
|Asignar resultado a variable|Nombre de la variable donde guardar el resultado|Variable|

### Bloquear celdas
  
Bloquea o desbloquea celdas
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja |Nombre de la hoja donde se encuentran los datos|Hoja 1|
|Rango donde buscar |Celda o Rango de celdas. La sintaxis debe ser la misma de excel (A1 o A1B1) |A1:B100 |
|Acción|Seleccione si desea bloquear o desbloquear una celda |Lock|

### Agregar Gráfico
  
Agrega un nuevo gráfico sobre una hoja en excel
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja |Nombre de la hoja donde se encuentran los datos|Hoja 1|
|Tipo de Gráfico|Seleccione el tipo de gráfico que se insertará en el excel|Line|
|Celda donde insertar gráfico |Celda donde insertar el gráfico. La sintaxis debe ser la misma de excel (A1) |A1|
|Rango de datos |Celda o Rango de celdas. La sintaxis debe ser la misma de excel (A1 o A1B1) |Sheet!A1:B100 |

### Quitar Contraseña
  
Quita la contraseña y guarda el Excel
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Excel con Contraseña|Ruta del archivo xlsx que se quiere abrir|C:/Users/User/Desktop/test.xlsx|
|Contraseña|Contraseña del archivo xlsx|****|
|Excel sin Contraseña|Ruta donde guardar el archivo .xlsx. Vacío para guardar en el mismo Excel|C:/Users/User/Desktop/test2.xlsx|

### Insertar imagen
  
Inserta una imagen
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja |Nombre de la hoja donde se encuentran los datos|Hoja 1|
|Celda |Celda donde insertar la imagen. La sintaxis debe ser la misma de excel (A1) |B5|
|Ruta imagen|Ruta de la imagen que se quiere insertar|imagen.png|

### Exportar gráfico
  
Exporta un gráfico por índice
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja |Nombre de la hoja donde se encuentran los datos|Hoja 1|
|Index |Índice del gráfico a exportar|1|
|Ruta imagen|Ruta dnde se guardará la imagen|/ruta/a/imagen.png|

### Modo no visible
  
Abre excel en modo no visible
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Ruta archivo XLSX|Ruta del archivo xlsx que se quiere abrir|Archivo.XLSX|
|Identificador (opcional)|Nombre o identificador para el archivo que se abrirá. Se utiliza cuando se necesita abrir más de un excel. Por defecto es *default*|default|

### Escribir array de objetos
  
Escribe un array de objetos en las celdas de Excel
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja |Nombre de la hoja donde se encuentran los datos|Hoja 1|
|Celda o Rango de celdas|Celda o Rango de celdas. La sintaxis debe ser la misma de excel (A1 o A1B1) |A1|
|Datos a escribir|Celda o Rango de celdas. La sintaxis debe ser la misma de excel (A1 o A1B1) |[{ 'id',: 1, 'text': 'hola' },{ 'id',: 2, 'text': 'mundo' }]|

### Copiar-Pegar Formato
  
Copia formato de un rango de celdas desde una hoja a otra 
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja origen |Nombre de la hoja de origen|Sheet1|
|Rango a copiar ||A1:C4|
|Hoja destino |Nombre de la hoja de destino|Sheet2|
|Rango donde pegar||A1:C4|

### Actualizar vínculos
  
Cambia un vínculo desde un documento a otro
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Ruta a cambiar|Ruta del archivo xlsx que se quiere actualizar||
|Ruta actualizada|Ruta del archivo xlsx que reemplazará el vinculo|file.xlsx|

### Desbloquear libro
  
Desbloquea un libro con contraseña
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Contraseña|Contraseña de la hoja bloqueada|Contraseña|

### Bloquear libro
  
Bloquear un libro con contraseña
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Contraseña|Contraseña para bloquear el libro|Contraseña|

### Desbloquear hoja
  
Desbloquea una hoja con contraseña
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja|Nombre de la hoja que se quiere desbloquear|Hoja 1|
|Contraseña|Contraseña de la hoja bloqueada|Contraseña|

### Bloquear hoja
  
Bloquear una hoja con contraseña
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja|Nombre de la hoja que se quiere bloquear|Hoja 1|
|Contraseña|Contraseña para bloquear la hoja|Contraseña|

### Convertir a .txt
  
Convierte a .txt
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Ruta archivo XLSX|Ruta del archivo xlsx que se quiere convertir|Archivo.XLSX|
|Guardar TXT|Ruta donde guardar el archivo .txt|/Users/user/Desktop/prueba.txt|

### Texto en columna
  
Ejecuta la opción texto en columna de excel
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja |Nombre de la hoja donde se encuentran los datos|Hoja 1|
|Rango donde buscar |Celda o Rango de celdas. La sintaxis debe ser la misma de excel (A1 o A1B1) |A1:B100 |
|Seleccione separador |Seleccione el separador de celdas, puede ser ancho fijo o delimitado||
|Seleccione tipo de delimitador |Seleccione el tipo de delimitador||
|Otro delimitador o ancho|Escriba el delimitador o ancho fijo|| o 20,35,22,10|

### Convertir tiempo de Excel a horas
  
Convertir tiempo de Excel a horas. Devuelve el resultado como hh:mm:ss
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Ingrese el tiempo en formato decimal ||0.296655812|
|Asignar resultado a variable|Nombre de la variable donde guardar el resultado|Variable|

### Imprimir hoja
  
Imprime una hoja
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja |Nombre de la hoja que se quiere imprimir|Hoja 1|

### Guardar Excel con password
  
Guarda un archivo Excel
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Guardar Excel en|Ruta donde guardar el archivo .xlsx|/Users/user/Desktop/excel.xlsx|
|Ingrese la password|Contraseña del archivo xlsx|password|

### Guardar Excel
  
Guarda un archivo Excel (como '.xlsx', 'xlsm', '.xls' or '.csv') en la ruta indicada
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Guardar Excel|Ruta donde guardar el archivo .xlsx|/Users/user/Desktop/excel.xlsx|

### Cerrar XLSX
  
Cierra el libro abierto por Rocketbot
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Matar proceso|Si se marca esta casillaa, cerrará por completo el proceso.||
