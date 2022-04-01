



# Opciones avanzadas para Excel
  
Módulo con opciones avanzadas para Excel  
  
![banner](imgs/Banner_AdvancedExcel.png)
## Como instalar este módulo
  
__Descarga__ e __instala__ el contenido en la carpeta 'modules' en la ruta de rocketbot.  




## Como usar este módulo
Para usar este módulo, tienes que tener Microsoft Excel.


## Descripción de los comandos

### Abrir sin alertas
  
Abre un archivo sin mostrar carteles de alerta.
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Ruta archivo XLSX|Ruta del archivo xlsx que se quiere abrir|Archivo.XLSX|
|Password (opcional)|Contraseña del archivo xlsx|P@ssW0rd|
|Identificador (opcional)|Nombre o identificador para el archivo que se abrirá. Se utiliza cuando se necesita abrir más de un excel. Por defecto es *default*|id|
|Asignar resultado a variable|Variable donde se almacenara el resultado|id|

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

### Color celda
  
Cambia color de una celda o rango de celdas. Puedes seleccionar un valor por defecto o uno personalizado
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Ingrese celdas |Celda o Rango de celdas. La sintaxis debe ser la misma de excel (A1 o A1B1) |A1:B5|
|Ingrese color en RGB |Valores rgb del color que tendrá la celda o celdas|250,250,250|
|Seleccione color |Seleccione el color. Puede usar el campo anterior para personalizar|red|

### Insertar Formula
  
Inserta formula sobre una celda 
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Ingrese celda |Celda o Rango de celdas. La sintaxis debe ser la misma de excel (A1 o A1B1) |A5|
|Escriba fórmula |Formula que se quiere insertar. Debe ser escrita en inglés. Recuerda usar *,* para separar parámetros|=SUM(A1:A4)|

### Insertar Macro a Excel
  
Inserta una Macro a Excel 
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Ruta de la Macro|Ruta del archivo .bas que se quiere insertar|Macro.bas|

### Seleccionar Celdas
  
Selecciona celdas en Excel
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

### Formatear Celda
  
Formatear Celda
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Nombre de Hoja|Nombre de la hoja que se quiere automatizar|Sheet1|
|Rango a formatear |Celda o Rango de celdas a formatear. La sintaxis debe ser la misma de excel (A1 o A1B1) |A1:C4|
|Formato|Se debe seleccionar el tipo de formato para la celda. Seleccione custom para adicionar un formato personalizado|dd-mm-yy|
|Formato personalizado |Formato personalizado. Debe ser el mismo mostrado en la sección personalizado de Excel|00000|

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
  
Copia un rango desde un Excel a otro, el excel de destino no debe estar abierto
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Excel origen|Ruta del archivo excel de origen|Sheet1|
|Hoja origen|Nombre de la hoja de origen|Sheet1|
|Rango a copiar|Celda o Rango de celdas a copiar. La sintaxis debe ser la misma de excel (A1 o A1B1) |A1:D7|
|Excel destino|Ruta del archivo excel de destino|Sheet1|
|Hoja destino|Nombre de la hoja donde se copiará|Sheet1|
|Rango donde pegar|Celda o Rango de celdas a copiar. La sintaxis debe ser la misma de excel (A1 o A1B1) |A1:D7|
|Solo valores|Si esta casilla es seleccionada, copiará solo los valores|True|

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
|Columna|Indique la columna o columnas que se quieren agregar o eliminar|B|

### Convertir CSV a XLSX
  
Convierte un documento CSV a XLSX
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Ruta archivo CSV|Ruta del archivo csv que se quiere convertir||
|Delimitador|Separador del archivo csv||
|Tiene cabeceras?|Marcar esta casilla si el csv tiene cabeceras|True|
|Codificación|Escriba el tipo de codificación del archivo. Por defecto es latin-1|latin-1|
|Ruta archivo XLSX|Ruta del archivo xlsx donde guardar|file.xlsx|

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
|Rango |Celda o Rango de celdas. La sintaxis debe ser la misma de excel (A1 o A1B1) |A1:E6 |

### Filtrar
  
Filtra a una tabla excel
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja |Nombre de la hoja donde se encuentran los datos|Hoja 1|
|Inicio de tabla |Columna donde comienza la tabla que se filtrará|A |
|Columna |Columna donde agregar el filtro|A |
|Filtro |Filtro o lista de filtros a agregar. Use "=" para encontrar campos en blanco, "<>" para celdas no vacías y negación de datos|['filtro1','filtro2', 'filtro3']|

### Renombrar hoja
  
Cambia el nombre a una hoja de excel
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja |Nombre de la hoja a renombrar|Hoja 1|
|Nuevo nombre |Nuevo Nombre de la hoja|nuevo_nombre|

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

### Pegar en Celdas
  
Pega datos en celdas en Excel
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja|Nombre de la hoja que se quiere automatizar|Hoja 1|
|Ingrese celdas donde pegar|Celda o Rango de celdas donde pegar. La sintaxis debe ser la misma de excel (A1 o A1B1) |A1:B3|
|Solo valores|Si esta casilla es seleccionada, se pegarán solo los valores|True|

### Eliminar duplicados
  
Ejecuta el comando eliminar duplicados de Excel
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja|Nombre de la hoja que se quiere automatizar|Hoja 1|
|Ingrese celdas donde filtrar|Celda o Rango de celdas. La sintaxis debe ser la misma de excel (A1 o A1B1) |A1:B3|
|Columna |Indicar la columna donde se buscarán los duplicados|A |
|Tiene cabeceras?|Marcar esta casilla si el excel tiene cabeceras|True|

### Cerrar XLSX
  
Cierra el libro abierto por Rocketbot
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |

### Guardar Excel
  
Guarda un archivo Excel en la ruta indicada
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Guardar Excel|Ruta donde guardar el archivo .xlsx|/Users/user/Desktop/excel.xlsx|

### Guardar Excel con password
  
Guarda un archivo Excel
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Guardar Excel en|Ruta donde guardar el archivo .xlsx|/Users/user/Desktop/excel.xlsx|
|Ingrese la password|Contraseña del archivo xlsx|password|

### Exportar a PDF avanzado
  
Exporta Excel a PDF con opciones
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Guardar PDF|Ruta donde guardar el archivo .pdf|/Users/user/Desktop/excel.pdf|
|Ajuste Automatico|||
|Zoom|||
|Ajustar Alto|||
|Ajustar ancho|||

### Copiar-Mover Hoja
  
Copia o mueve una hoja
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja origen |Nombre de la hoja de origen|Sheet1|
|Mover/copiar antes de hoja |Nombre de la hoja donde se moverá|Sheet2|
|Excel destino|Ruta del archivo .xlsx donde mover o copiar la hoja|C:/ruta/al/excel.xlsx|
|Copiar|Al marcar la casilla, se creará una copia de la hoja||

### Insertar Formulario
  
Inserta un Formulario a Excel 
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Ruta del Formulario|Ruta del archivo frm que se quiere insertar|Form.frm|

### Leer celdas filtradas
  
Lee solo las celdas filtradas
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja |Nombre de la hoja donde se encuentran los datos|Hoja 1|
|Rango donde buscar |Celda o Rango de celdas. La sintaxis debe ser la misma de excel (A1 o A1B1) |A1:B100 |
|Asignar resultado a variable|Nombre de la variable donde guardar el resultado|Variable|
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

### Actualizar Todo
  
Actualiza todas las fuentes del libro
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |

### Buscar
  
Devuelve la primera celda encontrada
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja |Nombre de la hoja donde se encuentran los datos|Hoja 1|
|Rango donde buscar |Celda o Rango de celdas. La sintaxis debe ser la misma de excel (A1 o A1B1) |A1:B100 |
|Texto a buscar|Texto que se quiere buscar en el excel|Lorem|
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
|Rango de datos |Celda o Rango de celdas. La sintaxis debe ser la misma de excel (A1 o A1B1) |A1:B100 |

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

### Buscar y conectar
  
Busca un excel abrierto y se conecta a este.
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Nombre del archivo XLSX abierto||Archivo.XLSX|
|Identificador (opcional)|Nombre o identificador para el archivo que se abrirá. Se utiliza cuando se necesita abrir más de un excel. Por defecto es *default*|excel1|

### Actualizar vínculos
  
Cambia un vínculo desde un documento a otro
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Ruta a cambiar|Ruta del archivo xlsx que se quiere actualizar||
|Ruta actualizada|Ruta del archivo xlsx que reemplazará el vinculo|file.xlsx|

### Desbloquear hoja
  
Desbloquea una hoja con contraseña
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja|Nombre de la hoja que se quiere bloquear|Hoja 1|
|Contraseña|Contraseña de la hoja bloqueada|Contraseña|

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
|Seleccione color |||
|Otro delimitador||,|

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
