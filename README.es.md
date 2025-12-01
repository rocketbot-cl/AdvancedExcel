



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

4. Opciones de calculo  
Selecciona la manera en que se ejecuta el calculo de formulas en el libro.

5. Leer celdas  
Lee una celda o rango de celdas

6. Convertir fecha serial  
Convierte una fecha numero serial excel a un formato de fecha especifico

7. Contar Columnas  
Cuenta el número de columnas del excel abierto. Se requiere que el excel esté guardado para tomar los últimos cambios

8. Contar Filas  
Cuenta todas las filas o dentro de un rango.

9. Ocultar  
Oculta una o varias filas, o una o varias columnas.

10. Mostrar  
Muestra una o varias filas, o una o varias columnas que estén ocultas

11. Color celda  
Cambia color de una celda o rango de celdas. Puedes seleccionar un valor por defecto o uno personalizado

12. Color fuente  
Cambie el color de fuente del texto en una celda o rango de celdas. Puede ser un color predeterminado o personalizado

13. Obtener color de celda  
Obtener el color de una celda. La función devolverá una lista con dos elementos: Color de fondo y Color de fuente en formato RGB.

14. Obtener formatos de celda  
Obtener el formato de una celda. La función devolverá un diccionario con las propiedades de la celda y el valor de cada una.

15. Insertar Formula  
Inserta formula sobre una celda 

16. Insertar Macro a Excel  
Inserta una Macro a Excel 

17. Seleccionar y copiar Celdas  
Selecciona y copia celdas en Excel

18. Obtener Celda Formato Moneda  
Obtiene celdas con formato moneda

19. Obtener Celda Formato Fecha  
Obtiene celdas con formato fecha

20. Copiar-Pegar  
Copia un rango de celdas desde una hoja a otra 

21. Formatear Celda  
Formatear Celda

22. Borrar contenido  
Borra fórmulas y valores del rango seleccionado, manteniendo el formato.

23. Crear Hoja  
Añade una hoja al final

24. Eliminar Hoja  
Elimina una hoja

25. Copiar de un Excel a otro  
Copia el rango de un archivo de Excel a otro. Indicando la ruta del archivo, abrirá el excel para copiar o pegar los datos. Si ingresas el id de un excel abierto, usará esa instancia para copiar o pegar.

26. Insertar/Eliminar Fila  
Inserta o elimina una fila

27. Insertar/Eliminar Columna  
Inserta o elimina una columna

28. Convertir CSV a XLSX  
Convierte un documento CSV a formato XLSX

29. Exportar a JSON  
Exporta un array de datos a un archivo JSON

30. (Deprecado) Convertir XLSX a CSV  
Convierte un documento XLSX a CSV

31. Convertir XLSX a CSV  
Convierte un documento XLSX a CSV

32. Convertir XLS a XLSX  
Convierte un documento XLS a XLSX

33. Obtener celda activa  
Obtener fila y columna de una celda activa

34. Actualizar tabla dinámica  
Actualiza una tabla dinámica. ¡Obsoleto! Use el módulo PivotTableExcel

35. Ajustar celdas  
Ajusta, une, agrupa y desagrupa un rango de celdas. Puedes agrupar/desagrupar por filas o columnas

36. Obtener Formula  
Obtiene la formula sobre una celda 

37. Agregar Filtro Automático  
Agrega filtro automático a una tabla excel

38. Eliminar Filtro Automático  
Eliminar el filtro automático de una hoja de Excel

39. Borrar Filtro  
Borra todos los filtros aplicados sobre una hoja de Excel

40. Filtrar  
Filtrar una tabla de excel según el valor relativo, contenido exacto, color de fondo o color de letra de las celdas. *Ejemplos según tipo de filtro: xlAnd ['>=10'] or ['>=10', '<=20'] | xlOr ['<=10', '>=20'] | xlFilterValues ['10','20', '30'] | xlFilterCellColor (255,0,0) | xlFilterFontColor (255,0,0)*

41. Filtrar por Fecha  
Filtra una tabla por el día, mes o año de una fecha indicada

42. Filtro avanzado  
Filtra a una tabla excel

43. Remover Filtros  
Eliminar filtros y mostrar todos los datos

44. Renombrar hoja  
Cambia el nombre a una hoja de excel

45. Formato de texto  
Cambia la alineacion Horizontal o Vertical de los valores en un rango de celdas

46. Estilo Celda  
Este comando modifica el formato de la celda o rango de celdas seleccionado. Puedes cambiar la fuente y los bordes

47. Pegar en Celdas  
Pega datos en celdas en Excel

48. Deshabilitar modo Copiar/Cortar  
Deshabilitar el modo Cortar/Copiar del Excel activo

49. Eliminar duplicados  
Ejecuta el comando eliminar duplicados de Excel

50. Exportar a PDF avanzado  
Exporta Excel a PDF con opciones

51. Copiar-Mover Hoja  
Copia o mueve una hoja

52. Insertar Formulario  
Inserta un Formulario a Excel 

53. Leer celdas filtradas  
Lee todo el contenido de las celdas filtradas y aplica formato a los datos tipo fecha si se indica

54. Contar celdas filtradas  
Cuenta solo las celdas filtradas

55. Reemplazar  
Ejecuta la opción de reemplazar de excel

56. Ordenar  
Ejecuta la opción de reemplazar de excel

57. Ordenar por múltiples niveles  
Ordene una hoja de Excel por valor, estableciendo múltiples niveles

58. Actualizar Todo  
Actualiza todas las fuentes del libro

59. Buscar  
Busca un texto en el rango indicado y retorna la celda donde se encuentra la primera coincidencia. Si no encuentra un valor, retornará vacío. Si el rango elta filtrado, la busqueda sere realizada sobre las celdas visibles.

60. Encontrar dato  
Devuelve la primera celda que coincida con el dato buscado

61. Bloquear celdas  
Bloquea o desbloquea celdas

62. Agregar Gráfico  
Agrega un nuevo gráfico sobre una hoja en excel

63. Quitar Contraseña  
Quita la contraseña y guarda el Excel

64. Insertar imagen  
Inserta una imagen

65. Exportar gráfico  
Exporta un gráfico por índice

66. Modo no visible  
Abre excel en modo no visible

67. Escribir array de objetos  
Escribe un array de objetos en las celdas de Excel

68. Copiar-Pegar Formato  
Copia formato de un rango de celdas desde una hoja a otra 

69. Actualizar vínculos  
Cambia un vínculo desde un documento a otro

70. Desbloquear libro  
Desbloquea un libro con contraseña

71. Bloquear libro  
Bloquear un libro con contraseña

72. Desbloquear hoja  
Desbloquea una hoja con contraseña

73. Bloquear hoja  
Bloquear una hoja con contraseña

74. Convertir a .txt  
Convierte a .txt

75. Texto en columna  
Ejecuta la opción texto en columna de excel

76. Convertir tiempo de Excel a horas  
Convertir tiempo de Excel a horas. Devuelve el resultado como hh:mm:ss

77. Combinar planillas  
Combina planillas de Excel que estén en la misma carpeta y que tengan los mismos encabezados. Combinará horizontalmente las hojas de la misma planilla y verticalmente las planillas diferentes.

78. Imprimir hoja  
Imprime una hoja

79. Guardar Excel con password  
Guarda un archivo Excel

80. Guardar Excel  
Guarda un archivo Excel (como '.xlsx', 'xlsm', '.xls', '.csv' or '.prn') en la ruta indicada

81. Cerrar XLSX  
Cierra el libro abierto por Rocketbot. El comportamiento de que solo mate un excel, funciona si es abierto con el comando Abrir sin alertas, caso contrario cerrara todos.

82. Eliminar Estilos  
Elimina estilos en una hoja

83. Insertar vinculo  
Inserta un vínculo en una celda, ya sea a una URL externa o a una celda dentro de otra hoja de cálculo.  




----
### OS

- windows
- mac

### Dependencies
- [**xlwings**](https://pypi.org/project/xlwings/)- [**pandas**](https://pypi.org/project/pandas/)
### License
  
![MIT](https://camo.githubusercontent.com/107590fac8cbd65071396bb4d04040f76cde5bde/687474703a2f2f696d672e736869656c64732e696f2f3a6c6963656e73652d6d69742d626c75652e7376673f7374796c653d666c61742d737175617265)  
[MIT](http://opensource.org/licenses/mit-license.ph)