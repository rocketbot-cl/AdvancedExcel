



# Opções avançadas para Excel
  
Aplique filtros automáticos e avançados, formate células, adicione ou exclua planilhas, linhas ou colunas, exporte para diferentes formatos de arquivo, desbloqueie e bloqueie novamente planilhas, copie e cole especiais e muito mais com seus arquivos do Excel.  

*Read this in other languages: [English](README.md), [Português](README.pr.md), [Español](README.es.md)*

## Como instalar este módulo
  
Para instalar o módulo no Rocketbot Studio, pode ser feito de duas formas:
1. Manual: __Baixe__ o arquivo .zip e descompacte-o na pasta módulos. O nome da pasta deve ser o mesmo do módulo e dentro dela devem ter os seguintes arquivos e pastas: \__init__.py, package.json, docs, example e libs. Se você tiver o aplicativo aberto, atualize seu navegador para poder usar o novo módulo.
2. Automático: Ao entrar no Rocketbot Studio na margem direita você encontrará a seção **Addons**, selecione **Install Mods**, procure o módulo desejado e aperte instalar.  


## Overview


1. Abrir sem alertas  
Abre um arquivo sem mostrar alertas do MS Excel.

2. Buscar e conectar  
Busca um excel aberto e conecta-se a este.

3. Opções de cálculo  
Selecione a forma como o cálculo da fórmula é executado na pasta de trabalho.

4. Ler células  
Ler uma célula ou intervalo de células

5. Converter data de série  
Converter uma data de número de série do Excel em um formato de data específico

6. Contar Colunas  
Conta o número de colunas do excel aberto. É necessário que o excel esteja salvo para tomar os últimos cambios

7. Contar Linhas  
Conta todas as linhas ou dentro de um intervalo.

8. Cor da célula  
Muda a cor de uma célula ou intervalo de células. Pode ser uma cor por defeito ou uma personalizada

9. Obter cor da célula  
Obter a cor de uma célula. A função retornará uma lista com dois elementos: Background Color e Font Color no formato RGB.

10. Obter formato de célula  
Obtenha o formato de uma célula. A função retornará um dicionário com as propriedades da célula e o valor de cada uma.

11. Insertar Formula  
Inserta formula sobre una celda 

12. Inserir Macro a Excel  
Insere uma Macro a Excel

13. Selecionar e copiar Células  
Seleciona e copia células em Excel

14. Obter Célula Formato Moeda  
Obtém células com formato moeda

15. Obter Célula Formato Data  
Obtém células com formato de data

16. Copiar-Colar  
Copia um intervalo de células de uma planilha para outra

17. Formatar Célula  
Formatar Célula

18. Remover conteúdo  
Limpa fórmulas e valores do intervalo selecionado, mantendo o formato

19. Criar Planilha  
Adiciona uma planilha no final

20. Eliminar Planilha  
Elmina uma planilha

21. Copiar de um Excel para outro  
Copie o intervalo de um arquivo Excel para outro. Indicando o caminho do arquivo, abrirá o Excel para copiar ou colar os dados. Se você inserir o id de um Excel aberto, ele usará essa instância para copiar ou colar.

22. Adicionar/Eliminar Linha  
Adiciona ou elimina uma linha

23. Adicionar/Excluir Coluna  
Adiciona o exclui uma coluna

24. Converter CSV para XLSX  
Converte um documento CSV para formato XLSX

25. (Descontinuado) Converter XLSX para CSV  
Converte um documento XLSX para CSV

26. Converter XLSX para CSV  
Converte um documento XLSX para CSV

27. Converter XLS para XLSX  
Converte um documento XLS para XLSX

28. Obter celula activa  
Obter linha e coluna de uma celula activa

29. Atualizar tabela dinâmica  
Atualiza uma tabela dinâmica. Descontinuado! Use o módulo PivotTableExcel

30. Ajustar células  
Ajusta, une, agrupa e desagrupa um intervalo de células. Você pode agrupar/desagrupar por linhas ou colunas

31. Obter Formula  
Obtém a fórmula numa célula

32. Adicionar Filtro Automático  
Adiciona filtro automático a uma tabela excel

33. Remover Filtro Automático  
Remova o filtro automático de uma planilha do Excel

34. Limpa Filtro  
Limpa todos os filtros feitos em uma planilha do Excel

35. Filtrar  
Filtre uma tabela do Excel de acordo com o valor relativo, conteúdo exato, cor de fundo ou cor da fonte das células. *Exemplos por tipo de filtro: xlAnd ['>=10'] or ['>=10', '<=20'] | xlOr ['<=10', '>=20'] | xlFilterValues ['10','20', '30'] | xlFilterCellColor (255,0,0) | xlFilterFontColor (255,0,0)*

36. Filtro avançado  
Aplicar filtro avançado a uma tabela

37. Remover filtros  
Remova os filtros e mostre todos os dados

38. Renomear planilha  
Muda o nome de uma planilha de excel

39. Formato de texto  
Altere o alinhamento Horizontal ou Vertical de valores em um intervalo de células

40. Estilo de Célula  
Este comando modifica o formata a célula o intervalo de células selecionado. Você pode mudar a fonte e as bordas

41. Colar em Células  
Colar dados em células em Excel

42. Remover duplicatas  
Executa o comando remover duplicatas de Excel

43. Exportar para PDF avançado  
Exporta Excel para PDF com opções

44. Copiar-Mover Planilha  
Copiar ou mover uma planilha

45. Inserir Formulário  
Insere um Formulário no Excel

46. Ler células filtradas  
Ler somente as células filtradas

47. Contar celulas filtradas  
Conta somente as celulas filtradas

48. Replace  
Run replace action to excel 

49. Ordenar  
Executa a ação de substituir de excel

50. Ordenar por múltiples niveles  
Ordene uma planilha Excel por valor, definindo vários níveis

51. Atualizar Tudo  
Atualiza todas as fontes do livro

52. Procurar  
Procura um texto no intervalo indicado e retorna a célula onde foi encontrada a primeira correspondência. Se não encontrar um valor, retornará vazio. Se o intervalo for filtrado, a pesquisa será realizada sobre as células visíveis.

53. Encontrar dados  
Retorna a primeira célula que corresponde aos dados da pesquisa

54. Bloquear celulas  
Bloquea ou desbloqueia celulas

55. Adicionar Gráfico  
Adiciona um novo gráfico sobre uma planilha de excel

56. Remover Senha  
Remove a senha e salva o Excel

57. Inserir imagem  
Inserir uma imagem

58. Exportar gráfico  
Exporta um gráfico por índice

59. Modo não visível  
Abre excel em modo não visível

60. Escrever array de objetos  
Escrever um array de objetos em células de Excel

61. Copiar-Colar Formato  
Copia formato de um intervalo de células de uma planilha para outra

62. Atualizar vínculos  
Muda um vínculo de um documento para outro

63. Desbloquear livro  
Desbloquea um livro com senha

64. Bloquear livro  
Bloquear um livro com senha

65. Desbloquear planilha  
Desbloquea uma folha com senha

66. Bloquear folha  
Bloquear uma folha com senha

67. Converter para .txt  
Converte para .txt

68. Texto em coluna  
Executa a opção texto em coluna de excel

69. Converter tempo de Excel para horas  
Converter tempo de Excel para horas. Retorna o formato como hh: mm: ss

70. Imprimir planilha  
Imprime uma planilha

71. Salvar Excel com senha  
Salva um arquivo Excel

72. Salvar Excel  
Salva um arquivo Excel (como '.xlsx', 'xlsm', '.xls' or '.csv')  na ruta indicada

73. Fechar XLSX  
Fecha o arquivo aberto por Rocketbot  




----
### OS

- windows
- mac

### Dependencies
- [**xlwings**](https://pypi.org/project/xlwings/)- [**pandas**](https://pypi.org/project/pandas/)
### License
  
![MIT](https://camo.githubusercontent.com/107590fac8cbd65071396bb4d04040f76cde5bde/687474703a2f2f696d672e736869656c64732e696f2f3a6c6963656e73652d6d69742d626c75652e7376673f7374796c653d666c61742d737175617265)  
[MIT](http://opensource.org/licenses/mit-license.ph)