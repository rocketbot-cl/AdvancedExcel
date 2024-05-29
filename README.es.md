



# Opciones avanzadas para Excel
  
Aplique filtros automaticos y avanzados, de formato a las celdas, añada o elimine hojas, filas o columnas, exporte a diferentes formatos de archivo, desbloquee y vuelva a bloquear hojas, copie y realice pegado especial y mas con sus archivos de Excel.   

*Read this in other languages: [English](README.md), [Português](README.pr.md), [Español](README.es.md)*

## Como instalar este módulo
  
Para instalar el módulo en Rocketbot Studio, se puede hacer de dos formas:
1. Manual: __Descargar__ el archivo .zip y descomprimirlo en la carpeta modules. El nombre de la carpeta debe ser el mismo al del módulo y dentro debe tener los siguientes archivos y carpetas: \__init__.py, package.json, docs, example y libs. Si tiene abierta la aplicación, refresca el navegador para poder utilizar el nuevo modulo.
2. Automática: Al ingresar a Rocketbot Studio sobre el margen derecho encontrara la sección de **Addons**, seleccionar **Install Mods**, buscar el modulo deseado y presionar install.  


## Overview


1. Abrir sin alertas  
Abre un archivo sin mostrar carteles de alerta.

2. Buscar y conectar  
Busca un excel abierto y se conecta a este.

3. Maximizar  
Maximizar Ventana de Excel

4. Calculation options  
Select the way the formula calculation is executed in the workbook.

5. Leer celdas  
Lee una celda o rango de celdas

6. Convertir fecha serial  
Convierte una fecha numero serial excel a un formato de fecha especifico

7. Contar Columnas  
Cuenta el número de columnas del excel abierto. Se requiere que el excel esté guardado para tomar los últimos cambios

8. Contar Filas  
Cuenta todas las filas o dentro de un rango.

9. Color celda  
Cambia color de una celda o rango de celdas. Puedes seleccionar un valor por defecto o uno personalizado

10. Obtener color de celda  
Obtener el color de una celda. La función devolverá una lista con dos elementos: Color de fondo y Color de fuente en formato RGB.

11. Obtener formatos de celda  
Obtener el formato de una celda. La función devolverá un diccionario con las propiedades de la celda y el valor de cada una.

12. Insertar Formula  
Inserta formula sobre una celda 

13. Insertar Macro a Excel  
Inserta una Macro a Excel 

14. Seleccionar y copiar Celdas  
Selecciona y copia celdas en Excel

15. Obtener Celda Formato Moneda  
Obtiene celdas con formato moneda

16. Obtener Celda Formato Fecha  
Obtiene celdas con formato fecha

17. Copiar-Pegar  
Copia un rango de celdas desde una hoja a otra 

18. Formatear Celda  
Formatear Celda

19. Borrar contenido  
Borra fórmulas y valores del rango seleccionado, manteniendo el formato.

20. Crear Hoja  
Añade una hoja al final

21. Eliminar Hoja  
Elimina una hoja

22. Copiar de un Excel a otro  
Copia el rango de un archivo de Excel a otro. Indicando la ruta del archivo, abrirá el excel para copiar o pegar los datos. Si ingresas el id de un excel abierto, usará esa instancia para copiar o pegar.

23. Insertar/Eliminar Fila  
Inserta o elimina una fila

24. Insertar/Eliminar Columna  
Inserta o elimina una columna

25. Convertir CSV a XLSX  
Convierte un documento CSV a formato XLSX

26. (Deprecado) Convertir XLSX a CSV  
Convierte un documento XLSX a CSV

27. Convertir XLSX a CSV  
Convierte un documento XLSX a CSV

28. Convertir XLS a XLSX  
Convierte un documento XLS a XLSX

29. Obtener celda activa  
Obtener fila y columna de una celda activa

30. Actualizar tabla dinámica  
Actualiza una tabla dinámica. ¡Obsoleto! Use el módulo PivotTableExcel

31. Ajustar celdas  
Ajusta, une, agrupa y desagrupa un rango de celdas. Puedes agrupar/desagrupar por filas o columnas

32. Obtener Formula  
Obtiene la formula sobre una celda 

33. Agregar Filtro Automático  
Agrega filtro automático a una tabla excel

34. Eliminar Filtro Automático  
Eliminar el filtro automático de una hoja de Excel

35. Borrar Filtro  
Borra todos los filtros aplicados sobre una hoja de Excel

36. Filtrar  
Filtrar una tabla de excel según el valor relativo, contenido exacto, color de fondo o color de letra de las celdas. *Ejemplos según tipo de filtro: xlAnd ['>=10'] or ['>=10', '<=20'] | xlOr ['<=10', '>=20'] | xlFilterValues ['10','20', '30'] | xlFilterCellColor (255,0,0) | xlFilterFontColor (255,0,0)*

37. Filtrar por Fecha  
Filtra una tabla por el día, mes o año de una fecha indicada

38. Filtro avanzado  
Filtra a una tabla excel

39. Remover Filtros  
Eliminar filtros y mostrar todos los datos

40. Renombrar hoja  
Cambia el nombre a una hoja de excel

41. Formato de texto  
Cambia la alineacion Horizontal o Vertical de los valores en un rango de celdas

42. Estilo Celda  
Este comando modifica el formato de la celda o rango de celdas seleccionado. Puedes cambiar la fuente y los bordes

43. Pegar en Celdas  
Pega datos en celdas en Excel

44. Deshabilitar modo Copiar/Cortar  
Deshabilitar el modo Cortar/Copiar del Excel activo

45. Eliminar duplicados  
Ejecuta el comando eliminar duplicados de Excel

46. Exportar a PDF avanzado  
Exporta Excel a PDF con opciones

47. Copiar-Mover Hoja  
Copia o mueve una hoja

48. Insertar Formulario  
Inserta un Formulario a Excel 

49. Leer celdas filtradas  
Lee solo las celdas filtradas

50. Contar celdas filtradas  
Cuenta solo las celdas filtradas

51. Reemplazar  
Ejecuta la opción de reemplazar de excel

52. Ordenar  
Ejecuta la opción de reemplazar de excel

53. Ordenar por múltiples niveles  
Ordene una hoja de Excel por valor, estableciendo múltiples niveles

54. Actualizar Todo  
Actualiza todas las fuentes del libro

55. Buscar  
Busca un texto en el rango indicado y retorna la celda donde se encuentra la primera coincidencia. Si no encuentra un valor, retornará vacío. Si el rango elta filtrado, la busqueda sere realizada sobre las celdas visibles.

56. Encontrar dato  
Devuelve la primera celda que coincida con el dato buscado

57. Bloquear celdas  
Bloquea o desbloquea celdas

58. Agregar Gráfico  
Agrega un nuevo gráfico sobre una hoja en excel

59. Quitar Contraseña  
Quita la contraseña y guarda el Excel

60. Insertar imagen  
Inserta una imagen

61. Exportar gráfico  
Exporta un gráfico por índice

62. Modo no visible  
Abre excel en modo no visible

63. Escribir array de objetos  
Escribe un array de objetos en las celdas de Excel

64. Copiar-Pegar Formato  
Copia formato de un rango de celdas desde una hoja a otra 

65. Actualizar vínculos  
Cambia un vínculo desde un documento a otro

66. Desbloquear libro  
Desbloquea un libro con contraseña

67. Bloquear libro  
Bloquear un libro con contraseña

68. Desbloquear hoja  
Desbloquea una hoja con contraseña

69. Bloquear hoja  
Bloquear una hoja con contraseña

70. Convertir a .txt  
Convierte a .txt

71. Texto en columna  
Ejecuta la opción texto en columna de excel

72. Convertir tiempo de Excel a horas  
Convertir tiempo de Excel a horas. Devuelve el resultado como hh:mm:ss

73. Imprimir hoja  
Imprime una hoja

74. Guardar Excel con password  
Guarda un archivo Excel

75. Guardar Excel  
Guarda un archivo Excel (como '.xlsx', 'xlsm', '.xls' or '.csv') en la ruta indicada

76. Cerrar XLSX  
Cierra el libro abierto por Rocketbot  




----
### OS

- windows
- mac

### Dependencies
- [**xlwings**](https://pypi.org/project/xlwings/)- [**pandas**](https://pypi.org/project/pandas/)
### License
  
![MIT](https://camo.githubusercontent.com/107590fac8cbd65071396bb4d04040f76cde5bde/687474703a2f2f696d672e736869656c64732e696f2f3a6c6963656e73652d6d69742d626c75652e7376673f7374796c653d666c61742d737175617265)  
[MIT](http://opensource.org/licenses/mit-license.ph)