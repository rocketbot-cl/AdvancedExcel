# Opções avançadas para Excel
  
Módulo com opções avançadas para Excel  

## Como instalar este módulo
  
__Baixe__ e __instale__ o conteúdo na pasta 'modules' no caminho do Rocketbot  



## Como usar este módulo
Para usar este módulo você deve ter o Microsoft Excel.


## Overview


1. Abrir sem alertas  
Abre um arquivo sem mostrar alertas do MS Excel.

2. Buscar e conectar  
Busca um excel aberto e conecta-se a este.

3. Contar Colunas  
Conta o número de colunas do excel aberto. É necessário que o excel esteja salvo para tomar os últimos cambios

4. Contar Linhas  
Conta todas as linhas ou dentro de um intervalo.

5. Cor da célula  
Muda a cor de uma célula ou intervalo de células. Pode ser uma cor por defeito ou uma personalizada

6. Obter cor da célula  
Obter a cor de uma célula. A função retornará uma lista com dois elementos: Background Color e Font Color no formato 
RGB.

7. Insertar Formula  
Inserta formula sobre una celda 

8. Inserir Macro a Excel  
Insere uma Macro a Excel

9. Selecionar Células  
Seleciona células em Excel

10. Obter Célula Formato Moeda  
Obtém células com formato moeda

11. Obter Célula Formato Data  
Obtém células com formato de data

12. Copiar-Colar  
Copia um intervalo de células de uma planilha para outra

13. Formatar Célula  
Formatar Célula

14. Criar Planilha  
Adiciona uma planilha no final

15. Eliminar Planilha  
Elmina uma planilha

16. Copiar de um Excel para outro  
Copia um intervalo de um Excel para outro, o excel de destino não deve estar aberto

17. Adicionar/Eliminar Linha  
Adiciona ou elimina uma linha

18. Adicionar/Excluir Coluna  
Adiciona o exclui uma coluna

19. Converter CSV para XLSX  
Converte um documento CSV para XLSX

20. (Descontinuado) Converter XLSX para CSV  
Converte um documento XLSX para CSV

21. Converter XLSX para CSV  
Converte um documento XLSX para CSV

22. Converter XLS para XLSX  
Converte um documento XLS para XLSX

23. Obter celula activa  
Obter linha e coluna de uma celula activa

24. Atualizar tabela dinâmica  
Atualiza uma tabela dinâmica. Descontinuado! Use o módulo PivotTableExcel

25. Ajustar células  
Ajusta, une, agrupa e desagrupa um intervalo de células. Você pode agrupar/desagrupar por linhas ou colunas

26. Obter Formula  
Obtém a fórmula numa célula

27. Adicionar Filtro Automático  
Adiciona filtro automático a uma tabela excel

28. Filtrar  
Filtra a uma tabela excel

29. Filtro avançado  
Aplicar filtro avançado a uma tabela

30. Remover filtros  
Remova os filtros e mostre todos os dados

31. Renomear planilha  
Muda o nome de uma planilha de excel

32. Formato de texto  
Altere o alinhamento Horizontal ou Vertical de valores em um intervalo de células

33. Estilo de Célula  
Este comando modifica o formata a célula o intervalo de células selecionado. Você pode mudar a fonte e as bordas

34. Colar em Células  
Colar dados em células em Excel

35. Remover duplicatas  
Executa o comando remover duplicatas de Excel

36. Exportar para PDF avançado  
Exporta Excel para PDF com opções

37. Copiar-Mover Planilha  
Copiar ou mover uma planilha

38. Inserir Formulário  
Insere um Formulário no Excel

39. Ler células filtradas  
Ler somente as células filtradas

40. Contar celulas filtradas  
Conta somente as celulas filtradas

41. Replace  
Run replace action to excel 

42. Ordenar  
Executa a ação de substituir de excel

43. Atualizar Tudo  
Atualiza todas as fontes do livro

44. Buscar  
Devuelve a primeira celula encontrada

45. Bloquear celulas  
Bloquea ou desbloqueia celulas

46. Adicionar Gráfico  
Adiciona um novo gráfico sobre uma planilha de excel

47. Remover Senha  
Remove a senha e salva o Excel

48. Inserir imagem  
Inserir uma imagem

49. Exportar gráfico  
Exporta um gráfico por índice

50. Modo não visível  
Abre excel em modo não visível

51. Escrever array de objetos  
Escrever um array de objetos em células de Excel

52. Copiar-Colar Formato  
Copia formato de um intervalo de células de uma planilha para outra

53. Atualizar vínculos  
Muda um vínculo de um documento para outro

54. Desbloquear planilha  
Desbloquea uma folha com senha

55. Bloquear folha  
Bloquear uma folha com senha

56. Converter para .txt  
Converte para .txt

57. Texto em coluna  
Executa a opção texto em coluna de excel

58. Converter tempo de Excel para horas  
Converter tempo de Excel para horas. Retorna o formato como hh: mm: ss

59. Imprimir planilha  
Imprime uma planilha

60. Salvar Excel com senha  
Salva um arquivo Excel

61. Salvar Excel  
Salva um arquivo Excel na ruta indicada

62. Fechar XLSX  
Fecha o arquivo aberto por Rocketbot  

### Changes
### 12-Jan-2023
- Add lock sheet command
### 11-Jan-2023
- Fix filter for mac and open whithout alerts compatibility for Rocket V2022
### 16-Nov-2022
- Can select chart Data Range from different sheet
### 16-Dic-2022
- Improve Get Filtered Cells, Fit Cells and CSV to XLSX
### 18-Nov-2022
- Get Filtered Cells parse cells with dates correctly
### 16-Nov-2022
- Fix compatibility with older versions of Filter command
### 15-Nov-2022
- Add Advanced Filter and Remove Filter commands
### 07-Nov-2022
- Add the possibilitie to save in .xls format
### 02-Nov-2022
- Add Filter, Auto Filter, Read and Write filtered cells for mac
- Add new xslx_to_csv command using xlwings, openpyxl one deprecated
#### 24-Oct-2022
- Add new features Filter command and update Text2Column to rely only on xlwings
#### 19-Oct-2022
- Add Filter, AutoFilter, GetCells and CountCells for MacOS
#### 28-Sep-2022
- Fix Remove Duplicates
#### 09-Aug-2022
- Add "Special Paste" options to Copy-Paste and add Get Cell Colors command
#### 22-Jul-2022
- Fix Copy to another excel and Copy-Paste Format
#### 13-Jun-2022
- Format text: Added command to change text alignment
#### 12-May-2022
- Copy to another excel: fixed command to copy from one excel to another
#### 18-Apr-2022
- Text to Column: command fixed to separate text in columns
#### 06-Apr-2022
- Fit Cells: Added merge cells, adjust rows, adjust columns functions
#### 28-Dec-2021
- Count Rows: command fixed to count all rows.
#### 9-Nov-2021
- Order command: Apply multiple orders and clean filters.
#### 13-Oct-2021
- Fix count cells filtered
#### 30-Sep-2021
- Paste command: Update compatibilities
#### 28-Sep-2021
- Fix get filtered cells command. Now returns extended data
#### 06-Jul-2021
- Fix language
#### 01-Jul-2021
- Read Filtered Cells: The command was fixed because it didn't getting all cell range
#### 27-Apr-2021
- Texto to column: Parses a column of cells that contain text into several columns.
#### 18-Mar-2021
- Unlock sheet: Convert XLSX to TXT.
#### 09-Mar-2021
- Unlock sheet: Unlock a sheet by password.
#### 09-Mar-2021
- Update links: Changes a link from one document to another
#### 17-Feb-2021
- Find and Connect: Find opened Excel file and connect it
#### 1-Feb-2021
- Add command Copy-Paste Format. You can copy format cell to another.
#### 25-Jan-2021
- Write array objects: Writes information obtained from an array of objects to excel cells
#### 21-Jan-2021
- Not visible mode: Open background Excel
#### 1-Dec-2020
- Export chart: Export a chart from index.
#### 24-Nov-2020
- Insert image in a cell.
#### 24-Sep-2020
- Open without alerts: Add field 'Password'
#### 16-Sep-2020
- Add chart: Create a new chart on excel sheet 
#### 15-Sep-2020
- Lock Cells: Lock or unlock cells 
#### 2-Sep-2020
- Find: Replicate Excel Find command 
#### 31-Jul-2020
- Order: Replicate Excel Order command 
#### 15-Jul-2020
- Read Filtered Cells: Read cell after execute Filter command
- Replace: Replicate Excel Replace command 
#### 2-Jul-2020
- Insert Form: Rocketbot can insert VBA Form to Excel
#### 30-Jun-2020
- Csv to xlsx: Checkbox header was added to decide if the csv has a header
- Export to Advanced PDF: Rocketbot export to PDF command enhancement
- Copy-Move Sheet: Replicate move/copy sheet command of Excel
#### 17-Jun-2020
- Remove duplicates: Rocketbot can now remove duplicate data on range Excel
#### 5-Jun-2020
- Focus Excel: Rocketbot can now set Excel to the foreground window

----
### OS

- windows
- mac

### Dependencies
- [**xlwings**](https://pypi.org/project/xlwings/)- [**pandas**](https://pypi.org/project/pandas/)
### License
  
![MIT](https://camo.githubusercontent.com/107590fac8cbd65071396bb4d04040f76cde5bde/687474703a2f2f696d672e736869656c64732e696f2f3a6c6963656e73652d6d69742d626c75652e7376673f7374796c653d666c61742d737175617265)  
[MIT](http://opensource.org/licenses/mit-license.ph)