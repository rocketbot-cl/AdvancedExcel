



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

3. Contar Columnas  
Cuenta el número de columnas del excel abierto. Se requiere que el excel esté guardado para tomar los últimos cambios

4. Contar Filas  
Cuenta todas las filas o dentro de un rango.

5. Color celda  
Cambia color de una celda o rango de celdas. Puedes seleccionar un valor por defecto o uno personalizado

6. Obtener color de celda  
Obtener el color de una celda. La función devolverá una lista con dos elementos: Color de fondo y Color de fuente en formato RGB.

7. Obtener formatos de celda  
Obtener el formato de una celda. La función devolverá un diccionario con las propiedades de la celda y el valor de cada una.

8. Insertar Formula  
Inserta formula sobre una celda 

9. Insertar Macro a Excel  
Inserta una Macro a Excel 

10. Seleccionar y copiar Celdas  
Selecciona y copia celdas en Excel

11. Obtener Celda Formato Moneda  
Obtiene celdas con formato moneda

12. Obtener Celda Formato Fecha  
Obtiene celdas con formato fecha

13. Copiar-Pegar  
Copia un rango de celdas desde una hoja a otra 

14. Formatear Celda  
Formatear Celda

15. Borrar contenido  
Borra fórmulas y valores del rango seleccionado, manteniendo el formato.

16. Crear Hoja  
Añade una hoja al final

17. Eliminar Hoja  
Elimina una hoja

18. Copiar de un Excel a otro  
Copia un rango desde un Excel a otro, el excel de destino no debe estar abierto

19. Insertar/Eliminar Fila  
Inserta o elimina una fila

20. Insertar/Eliminar Columna  
Inserta o elimina una columna

21. Convertir CSV a XLSX  
Convierte un documento CSV a formato XLSX

22. (Deprecado) Convertir XLSX a CSV  
Convierte un documento XLSX a CSV

23. Convertir XLSX a CSV  
Convierte un documento XLSX a CSV

24. Convertir XLS a XLSX  
Convierte un documento XLS a XLSX

25. Obtener celda activa  
Obtener fila y columna de una celda activa

26. Actualizar tabla dinámica  
Actualiza una tabla dinámica. ¡Obsoleto! Use el módulo PivotTableExcel

27. Ajustar celdas  
Ajusta, une, agrupa y desagrupa un rango de celdas. Puedes agrupar/desagrupar por filas o columnas

28. Obtener Formula  
Obtiene la formula sobre una celda 

29. Agregar Filtro Automático  
Agrega filtro automático a una tabla excel

30. Eliminar Filtro Automático  
Eliminar el filtro automático de una hoja de Excel

31. Filtrar  
Filtrar una tabla de excel según el valor relativo, contenido exacto, color de fondo o color de letra de las celdas. *Ejemplos según tipo de filtro: xlAnd ['>=10'] or ['>=10', '<=20'] | xlOr ['<=10', '>=20'] | xlFilterValues ['10','20', '30'] | xlFilterCellColor (255,0,0) | xlFilterFontColor (255,0,0)*

32. Filtro avanzado  
Filtra a una tabla excel

33. Remover Filtros  
Eliminar filtros y mostrar todos los datos

34. Renombrar hoja  
Cambia el nombre a una hoja de excel

35. Formato de texto  
Cambia la alineacion Horizontal o Vertical de los valores en un rango de celdas

36. Estilo Celda  
Este comando modifica el formato de la celda o rango de celdas seleccionado. Puedes cambiar la fuente y los bordes

37. Pegar en Celdas  
Pega datos en celdas en Excel

38. Eliminar duplicados  
Ejecuta el comando eliminar duplicados de Excel

39. Exportar a PDF avanzado  
Exporta Excel a PDF con opciones

40. Copiar-Mover Hoja  
Copia o mueve una hoja

41. Insertar Formulario  
Inserta un Formulario a Excel 

42. Leer celdas filtradas  
Lee solo las celdas filtradas

43. Contar celdas filtradas  
Cuenta solo las celdas filtradas

44. Reemplazar  
Ejecuta la opción de reemplazar de excel

45. Ordenar  
Ejecuta la opción de reemplazar de excel

46. Actualizar Todo  
Actualiza todas las fuentes del libro

47. Buscar  
Busca un texto en el rango indicado y retorna la celda donde se encuentra la primera coincidencia. Si no encuentra un valor, retornará vacío. Si el rango elta filtrado, la busqueda sere realizada sobre las celdas visibles.

48. Encontrar dato  
Devuelve la primera celda que coincida con el dato buscado

49. Bloquear celdas  
Bloquea o desbloquea celdas

50. Agregar Gráfico  
Agrega un nuevo gráfico sobre una hoja en excel

51. Quitar Contraseña  
Quita la contraseña y guarda el Excel

52. Insertar imagen  
Inserta una imagen

53. Exportar gráfico  
Exporta un gráfico por índice

54. Modo no visible  
Abre excel en modo no visible

55. Escribir array de objetos  
Escribe un array de objetos en las celdas de Excel

56. Copiar-Pegar Formato  
Copia formato de un rango de celdas desde una hoja a otra 

57. Actualizar vínculos  
Cambia un vínculo desde un documento a otro

58. Desbloquear hoja  
Desbloquea una hoja con contraseña

59. Bloquear hoja  
Bloquear una hoja con contraseña

60. Convertir a .txt  
Convierte a .txt

61. Texto en columna  
Ejecuta la opción texto en columna de excel

62. Convertir tiempo de Excel a horas  
Convertir tiempo de Excel a horas. Devuelve el resultado como hh:mm:ss

63. Imprimir hoja  
Imprime una hoja

64. Guardar Excel con password  
Guarda un archivo Excel

65. Guardar Excel  
Guarda un archivo Excel (como '.xlsx', '.xls' or '.csv') en la ruta indicada

66. Cerrar XLSX  
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