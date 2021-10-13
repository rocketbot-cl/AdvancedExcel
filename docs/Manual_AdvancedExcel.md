
# Opciones avanzadas para Excel
  
Módulo con opciones avanzadas para Excel  
  
![banner](imgs/Banner_AdvancedExcel.png)
## Como instalar este módulo
  
__Descarga__ e __instala__ el contenido en la carpeta 'modules' en la ruta de rocketbot.
## Como usar este module
  
Eiusmod veniam ut nisi minim in. Do et deserunt eiusmod veniam sint aliqua nulla adipisicing laboris voluptate fugiat 
ullamco elit do. Sint amet cillum fugiat excepteur mollit voluptate reprehenderit nisi commodo sint minim.
## Descripción de los comandos

### Abrir sin alertas
  
Abre un archivo sin mostrar carteles de alerta.
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Ruta archivo XLSX||Archivo.XLSX|
|Password (opcional)||P@ssW0rd|
|Identificador (opcional)||id|

### Contar Columnas
  
Contar Columnas
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja||Hoja 1|
|Obtener nombre de columna|||
|Asignar resultado a variable||Variable|

### Contar Filas
  
Contar Filas
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja||Hoja 1|
|Columna||C|
|Asignar resultado a variable||Variable|

### Color celda
  
Cambia color de una celda o rango de celdas 
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Ingrese celdas ||A1:B5|
|Ingrese color en RGB ||250,250,250|
|Seleccione color |||

### Insertar Formula
  
Inserta formula sobre una celda 
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Ingrese celda ||A5|
|Escriba fórmula ||=SUM(A1:A4)|

### Insertar Macro a Excel
  
Inserta una Macro a Excel 
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Ruta de la Macro||Macro.bas|

### Seleccionar Celdas
  
Selecciona celdas en Excel
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja||Hoja 1|
|Ingrese celdas a seleccionar||A1:B3|
|Copiar|||

### Obtener Celda Formato Moneda
  
Obtiene celdas con formato moneda
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja||Hoja 1|
|Ingrese celdas a seleccionar||A1:B3|
|Asignar a variable||variable|

### Copiar-Pegar
  
Copia un rango de celdas desde una hoja a otra 
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja origen ||Sheet1|
|Rango a copiar ||A1:C4|
|Hoja destino ||Sheet2|
|Rango donde pegar||A1:C4|

### Formatear Celda
  
Formatear Celda
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Nombre de Hoja||Sheet1|
|Rango a formatear ||A1:C4|
|Formato||B:C|
|Formato personalizado ||00000|

### Crear Hoja
  
Añade una hoja al final
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Nombre de la hoja||Sheet2|
|Despues de||Hoja1|

### Eliminar Hoja
  
Elimina una hoja
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Nombre de la hoja||Sheet2|
|Asignar resultado a variable||Variable|

### Copiar de un Excel a otro
  
Copia un rango desde un Excel a otro, el excel de destino no debe estar abierto
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Excel origen||Sheet1|
|Hoja origen||Sheet1|
|Rango a copiar||A1:D7|
|Excel destino||Sheet1|
|Hoja destino||Sheet1|
|Rango donde pegar||A1:D7|
|Solo valores|||

### Insertar/Eliminar Fila
  
Inserta o elimina una fila
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Opción|||
|Nombre de Hoja||Sheet|
|Número Fila||2|
|Dónde Insertar||A1:D7|

### Insertar/Eliminar Columna
  
Inserta o elimina una columna
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Opción|||
|Nombre de Hoja||Sheet|
|Columna||B|

### Convertir CSV a XLSX
  
Convierte un documento CSV a XLSX
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Ruta archivo CSV|||
|Delimitador|||
|Tiene cabeceras?|||
|Codificación||latin-1|
|Ruta archivo XLSX||file.xlsx|

### Convertir XLSX a CSV
  
Convierte un documento XLSX a CSV
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Ruta archivo XLSX||C:/Users/User/Desktop/file.xlsx|
|Delimitador||,|
|Nombre de la hoja||Sheet0|
|Ruta archivo CSV||C:/Users/User/Desktop/file.csv|

### Convertir XLS a XLSX
  
Convierte un documento XLS a XLSX
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Ruta archivo XLS||C:\Users\User\Desktop\file.xls|
|Ruta archivo XLSX||C:\Users\User\Desktop\new_file.xlsx|

### Obtener celda activa
  
Obtener fila y columna de una celda activa
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Asignar resultado a variable||Variable|

### Actualizar tabla dinámica
  
Actualiza una tabla dinámica. ¡Obsoleto! Use el módulo PivotTableExcel
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja ||Hoja 1|
|Nombre de la tabla dinámica ||Name: |

### Ajustar celdas
  
Ajusta un rango de celdas
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja||Hoja 1|
|Rango a ajustar||A1:D7|
|Autofit|||
|Agrupar filas|||
|Agrupar columnas|||
|Desagrupar filas|||
|Desagrupar columnas|||
|Nivel de fila||2|
|Rango de columna||2|

### Obtener Formula
  
Obtiene la formula sobre una celda 
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Ingrese celda ||A5|
|Asignar resultado a variable||Variable|

### Agregar Filtro Automático
  
Agrega filtro automático a una tabla excel
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja ||Hoja 1|
|Rango ||A1:E6 |

### Filtrar
  
Filtra a una tabla excel
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja ||Hoja 1|
|Inicio de tabla ||A |
|Columna ||A |
|Filtro ||['filtro1','filtro2', 'filtro3']|

### Renombrar hoja
  
Cambia el nombre a una hoja de excel
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja ||Hoja 1|
|Nuevo nombre ||nuevo_nombre|

### Estilo Celda
  
Formatear Celda
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Nombre de Hoja||Sheet1|
|Rango a formatear ||A1:C4|
|Borde||--Seleccione--|
|Estilo||--Seleccione--|
|Tamaño de fuente ||20|
|Negrita||A1:C4|
|Cursiva||A1:C4|
|Subrayar||A1:C4|

### Pegar en Celdas
  
Pega datos en celdas en Excel
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja||Hoja 1|
|Ingrese celdas donde pegar||A1:B3|
|Solo valores|||

### Eliminar duplicados
  
Ejecuta el comando eliminar duplicados de Excel
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja||Hoja 1|
|Ingrese celdas donde filtrar||A1:B3|
|Columna ||A |
|Tiene cabeceras?|||

### Cerrar XLSX
  
Cierra el libro abierto por Rocketbot
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |

### Guardar Excel
  
Guarda un archivo Excel
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Guardar Excel||/Users/user/Desktop/excel.xlsx|

### Guardar Excel con password
  
Guarda un archivo Excel
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Guardar Excel en||/Users/user/Desktop/excel.xlsx|
|Ingrese la password||password|

### Exportar a PDF avanzado
  
Exporta Excel a PDF con opciones
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Guardar PDF||/Users/user/Desktop/excel.pdf|
|Ajuste Automatico|||
|Zoom|||
|Ajustar Alto|||
|Ajustar ancho|||

### Copiar-Mover Hoja
  
Copia o mueve una hoja
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja origen ||Sheet1|
|Mover/copiar antes de hoja ||Sheet2|
|Excel destino||C:/ruta/al/excel.xlsx|
|Copiar|||

### Insertar Formulario
  
Inserta un Formulario a Excel 
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Ruta del Formulario||Form.frm|

### Leer celdas filtradas
  
Lee solo las celdas filtradas
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja ||Hoja 1|
|Rango donde buscar ||A1:B100 |
|Asignar resultado a variable||Variable|
|Datos extra|||

### Contar celdas filtradas
  
Cuenta solo las celdas filtradas
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja ||Hoja 1|
|Rango donde buscar ||A1:B100 |
|Asignar resultado a variable||Variable|
|Datos extra|||

### Reemplazar
  
Ejecuta la opción de reemplazar de excel
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja ||Hoja 1|
|Rango donde buscar ||A1:B100 |
|Palabra a reemplazar||10/10/2020|
|Nueva palabra||10-10-2020|

### Ordenar
  
Ejecuta la opción de reemplazar de excel
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja ||Hoja 1|
|Rango donde buscar ||A1:B100 |
|Columna||A1:A22|
|Tipo de orden |||

### Actualizar Todo
  
Actualiza todas las fuentes del libro
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |

### Buscar
  
Devuelve la primera celda encontrada
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja ||Hoja 1|
|Rango donde buscar ||A1:B100 |
|Texto a buscar||Lorem|
|Asignar resultado a variable||Variable|

### Bloquear celdas
  
Bloquea o desbloquea celdas
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja ||Hoja 1|
|Rango donde buscar ||A1:B100 |
|Acción||Lorem|

### Agregar Gráfico
  
Agrega un nuevo gráfico sobre una hoja en excel
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja ||Hoja 1|
|Tipo de Gráfico||Lorem|
|Celda donde insertar gráfico ||A1|
|Rango de datos ||A1:B100 |

### Quitar Contraseña
  
Quita la contraseña y guarda el Excel
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Excel con Contraseña||C:/Users/User/Desktop/test.xlsx|
|Contraseña||****|
|Excel sin Contraseña||C:/Users/User/Desktop/test2.xlsx|

### Insertar imagen
  
Inserta una imagen
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja ||Hoja 1|
|Celda ||B5|
|Ruta imagen||imagen.png|

### Exportar gráfico
  
Exporta un gráfico por índice
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja ||Hoja 1|
|Index ||1|
|Ruta imagen||/ruta/a/imagen.png|

### Modo no visible
  
Abre excel en modo no visible
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Ruta archivo XLSX||Archivo.XLSX|
|Identificador (opcional)||id|

### Escribir array de objetos
  
Escribe un array de objetos en las celdas de Excel
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja ||Hoja 1|
|Celda o Rango de celdas||A1|
|Datos a escribir||[{ 'id',: 1, 'text': 'hola' },{ 'id',: 2, 'text': 'mundo' }]|

### Copiar-Pegar Formato
  
Copia formato de un rango de celdas desde una hoja a otra 
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja origen ||Sheet1|
|Rango a copiar ||A1:C4|
|Hoja destino ||Sheet2|
|Rango donde pegar||A1:C4|

### Buscar y conectar
  
Busca un excel abrierto y se conecta a este.
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Nombre del archivo XLSX abierto||Archivo.XLSX|
|Identificador (opcional)||id|

### Actualizar vínculos
  
Cambia un vínculo desde un documento a otro
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Ruta a cambiar|||
|Ruta actualizada||file.xlsx|

### Desbloquear hoja
  
Desbloquea una hoja con contraseña
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja||Hoja 1|
|Contraseña||Contraseña|

### Convertir a .txt
  
Convierte a .txt
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Ruta archivo XLSX||Archivo.XLSX|
|Guardar TXT||/Users/user/Desktop/prueba.txt|

### Texto en columna
  
Ejecuta la opción texto en columna de excel
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja ||Hoja 1|
|Rango donde buscar ||A1:B100 |
|Seleccione color |||
|Otro delimitador||,|

### Convertir tiempo de Excel a horas
  
Convertir tiempo de Excel a horas
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Ingrese el tiempo en formato decimal ||0.296655812|
|Asignar resultado a variable||Variable|
