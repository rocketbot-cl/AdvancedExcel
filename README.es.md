



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

7. Insertar Formula  
Inserta formula sobre una celda 

8. Insertar Macro a Excel  
Inserta una Macro a Excel 

9. Seleccionar y copiar Celdas  
Selecciona y copia celdas en Excel

10. Obtener Celda Formato Moneda  
Obtiene celdas con formato moneda

11. Obtener Celda Formato Fecha  
Obtiene celdas con formato fecha

12. Copiar-Pegar  
Copia un rango de celdas desde una hoja a otra 

13. Formatear Celda  
Formatear Celda

14. Crear Hoja  
Añade una hoja al final

15. Eliminar Hoja  
Elimina una hoja

16. Copiar de un Excel a otro  
Copia un rango desde un Excel a otro, el excel de destino no debe estar abierto

17. Insertar/Eliminar Fila  
Inserta o elimina una fila

18. Insertar/Eliminar Columna  
Inserta o elimina una columna

19. Convertir CSV a XLSX  
Convierte un documento CSV a XLSX

20. (Deprecado) Convertir XLSX a CSV  
Convierte un documento XLSX a CSV

21. Convertir XLSX a CSV  
Convierte un documento XLSX a CSV

22. Convertir XLS a XLSX  
Convierte un documento XLS a XLSX

23. Obtener celda activa  
Obtener fila y columna de una celda activa

24. Actualizar tabla dinámica  
Actualiza una tabla dinámica. ¡Obsoleto! Use el módulo PivotTableExcel

25. Ajustar celdas  
Ajusta, une, agrupa y desagrupa un rango de celdas. Puedes agrupar/desagrupar por filas o columnas

26. Obtener Formula  
Obtiene la formula sobre una celda 

27. Agregar Filtro Automático  
Agrega filtro automático a una tabla excel

28. Filtrar  
Filtrar una tabla de excel según el valor relativo, contenido exacto, color de fondo o color de letra de las celdas. *Ejemplos según tipo de filtro: xlAnd ['>=10'] or ['>=10', '<=20'] | xlOr ['<=10', '>=20'] | xlFilterValues ['10','20', '30'] | xlFilterCellColor (255,0,0) | xlFilterFontColor (255,0,0)*

29. Filtro avanzado  
Filtra a una tabla excel

30. Remover Filtros  
Eliminar filtros y mostrar todos los datos

31. Renombrar hoja  
Cambia el nombre a una hoja de excel

32. Formato de texto  
Cambia la alineacion Horizontal o Vertical de los valores en un rango de celdas

33. Estilo Celda  
Este comando modifica el formato de la celda o rango de celdas seleccionado. Puedes cambiar la fuente y los bordes

34. Pegar en Celdas  
Pega datos en celdas en Excel

35. Eliminar duplicados  
Ejecuta el comando eliminar duplicados de Excel

36. Exportar a PDF avanzado  
Exporta Excel a PDF con opciones

37. Copiar-Mover Hoja  
Copia o mueve una hoja

38. Insertar Formulario  
Inserta un Formulario a Excel 

39. Leer celdas filtradas  
Lee solo las celdas filtradas

40. Contar celdas filtradas  
Cuenta solo las celdas filtradas

41. Reemplazar  
Ejecuta la opción de reemplazar de excel

42. Ordenar  
Ejecuta la opción de reemplazar de excel

43. Actualizar Todo  
Actualiza todas las fuentes del libro

44. (Deprecado) Buscar  
Devuelve la primera celda encontrada

45. Encontrar dato  
Devuelve la primera celda que coincida con el dato buscado

46. Bloquear celdas  
Bloquea o desbloquea celdas

47. Agregar Gráfico  
Agrega un nuevo gráfico sobre una hoja en excel

48. Quitar Contraseña  
Quita la contraseña y guarda el Excel

49. Insertar imagen  
Inserta una imagen

50. Exportar gráfico  
Exporta un gráfico por índice

51. Modo no visible  
Abre excel en modo no visible

52. Escribir array de objetos  
Escribe un array de objetos en las celdas de Excel

53. Copiar-Pegar Formato  
Copia formato de un rango de celdas desde una hoja a otra 

54. Actualizar vínculos  
Cambia un vínculo desde un documento a otro

55. Desbloquear hoja  
Desbloquea una hoja con contraseña

56. Bloquear hoja  
Bloquear una hoja con contraseña

57. Convertir a .txt  
Convierte a .txt

58. Texto en columna  
Ejecuta la opción texto en columna de excel

59. Convertir tiempo de Excel a horas  
Convertir tiempo de Excel a horas. Devuelve el resultado como hh:mm:ss

60. Imprimir hoja  
Imprime una hoja

61. Guardar Excel con password  
Guarda un archivo Excel

62. Guardar Excel  
Guarda un archivo Excel en la ruta indicada

63. Cerrar XLSX  
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