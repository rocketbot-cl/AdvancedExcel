



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

3. Maximizar  
Maximizar a janela do Excel

4. Opções de cálculo  
Selecione a forma como o cálculo da fórmula é executado na pasta de trabalho.

5. Ler células  
Ler uma célula ou intervalo de células

6. Converter data de série  
Converter uma data de número de série do Excel em um formato de data específico

7. Contar Colunas  
Conta o número de colunas do excel aberto. É necessário que o excel esteja salvo para tomar os últimos cambios

8. Contar Linhas  
Conta todas as linhas ou dentro de um intervalo.

9. Ocultar  
Oculta uma ou várias linhas, ou uma ou várias colunas.

10. Ocultar  
Mostra uma ou mais linhas, ou uma ou mais colunas que estão ocultas

11. Cor da célula  
Muda a cor de uma célula ou intervalo de células. Pode ser uma cor por defeito ou uma personalizada

12. Obter cor da célula  
Obter a cor de uma célula. A função retornará uma lista com dois elementos: Background Color e Font Color no formato RGB.

13. Obter formato de célula  
Obtenha o formato de uma célula. A função retornará um dicionário com as propriedades da célula e o valor de cada uma.

14. Insertar Formula  
Inserta formula sobre una celda 

15. Inserir Macro a Excel  
Insere uma Macro a Excel

16. Selecionar e copiar Células  
Seleciona e copia células em Excel

17. Obter Célula Formato Moeda  
Obtém células com formato moeda

18. Obter Célula Formato Data  
Obtém células com formato de data

19. Copiar-Colar  
Copia um intervalo de células de uma planilha para outra

20. Formatar Célula  
Formatar Célula

21. Remover conteúdo  
Limpa fórmulas e valores do intervalo selecionado, mantendo o formato

22. Criar Planilha  
Adiciona uma planilha no final

23. Eliminar Planilha  
Elmina uma planilha

24. Copiar de um Excel para outro  
Copie o intervalo de um arquivo Excel para outro. Indicando o caminho do arquivo, abrirá o Excel para copiar ou colar os dados. Se você inserir o id de um Excel aberto, ele usará essa instância para copiar ou colar.

25. Adicionar/Eliminar Linha  
Adiciona ou elimina uma linha

26. Adicionar/Excluir Coluna  
Adiciona o exclui uma coluna

27. Converter CSV para XLSX  
Converte um documento CSV para formato XLSX

28. Exportar para JSON  
Exporta um array de dados para um arquivo JSON

29. (Descontinuado) Converter XLSX para CSV  
Converte um documento XLSX para CSV

30. Converter XLSX para CSV  
Converte um documento XLSX para CSV

31. Converter XLS para XLSX  
Converte um documento XLS para XLSX

32. Obter celula activa  
Obter linha e coluna de uma celula activa

33. Atualizar tabela dinâmica  
Atualiza uma tabela dinâmica. Descontinuado! Use o módulo PivotTableExcel

34. Ajustar células  
Ajusta, une, agrupa e desagrupa um intervalo de células. Você pode agrupar/desagrupar por linhas ou colunas

35. Obter Formula  
Obtém a fórmula numa célula

36. Adicionar Filtro Automático  
Adiciona filtro automático a uma tabela excel

37. Remover Filtro Automático  
Remova o filtro automático de uma planilha do Excel

38. Limpa Filtro  
Limpa todos os filtros feitos em uma planilha do Excel

39. Filtrar  
Filtre uma tabela do Excel de acordo com o valor relativo, conteúdo exato, cor de fundo ou cor da fonte das células. *Exemplos por tipo de filtro: xlAnd ['>=10'] or ['>=10', '<=20'] | xlOr ['<=10', '>=20'] | xlFilterValues ['10','20', '30'] | xlFilterCellColor (255,0,0) | xlFilterFontColor (255,0,0)*

40. Filtrar por Data  
Filtrar uma tabela por o dia, mes ou ano de uma data indicada

41. Filtro avançado  
Aplicar filtro avançado a uma tabela

42. Remover filtros  
Remova os filtros e mostre todos os dados

43. Renomear planilha  
Muda o nome de uma planilha de excel

44. Formato de texto  
Altere o alinhamento Horizontal ou Vertical de valores em um intervalo de células

45. Estilo de Célula  
Este comando modifica o formata a célula o intervalo de células selecionado. Você pode mudar a fonte e as bordas

46. Colar em Células  
Colar dados em células em Excel

47. Desativar modo de corte/cópia  
Desative o modo Cortar/Copiar do Excel ativo

48. Remover duplicatas  
Executa o comando remover duplicatas de Excel

49. Exportar para PDF avançado  
Exporta Excel para PDF com opções

50. Copiar-Mover Planilha  
Copiar ou mover uma planilha

51. Inserir Formulário  
Insere um Formulário no Excel

52. Ler células filtradas  
Leia todo o conteúdo das células filtradas e aplique formatação aos dados do tipo data, se indicado

53. Contar celulas filtradas  
Conta somente as celulas filtradas

54. Replace  
Run replace action to excel 

55. Ordenar  
Executa a ação de substituir de excel

56. Ordenar por múltiples niveles  
Ordene uma planilha Excel por valor, definindo vários níveis

57. Atualizar Tudo  
Atualiza todas as fontes do livro

58. Procurar  
Procura um texto no intervalo indicado e retorna a célula onde foi encontrada a primeira correspondência. Se não encontrar um valor, retornará vazio. Se o intervalo for filtrado, a pesquisa será realizada sobre as células visíveis.

59. Encontrar dados  
Retorna a primeira célula que corresponde aos dados da pesquisa

60. Bloquear celulas  
Bloquea ou desbloqueia celulas

61. Adicionar Gráfico  
Adiciona um novo gráfico sobre uma planilha de excel

62. Remover Senha  
Remove a senha e salva o Excel

63. Inserir imagem  
Inserir uma imagem

64. Exportar gráfico  
Exporta um gráfico por índice

65. Modo não visível  
Abre excel em modo não visível

66. Escrever array de objetos  
Escrever um array de objetos em células de Excel

67. Copiar-Colar Formato  
Copia formato de um intervalo de células de uma planilha para outra

68. Atualizar vínculos  
Muda um vínculo de um documento para outro

69. Desbloquear livro  
Desbloquea um livro com senha

70. Bloquear livro  
Bloquear um livro com senha

71. Desbloquear planilha  
Desbloquea uma folha com senha

72. Bloquear folha  
Bloquear uma folha com senha

73. Converter para .txt  
Converte para .txt

74. Texto em coluna  
Executa a opção texto em coluna de excel

75. Converter tempo de Excel para horas  
Converter tempo de Excel para horas. Retorna o formato como hh: mm: ss

76. Imprimir planilha  
Imprime uma planilha

77. Salvar Excel com senha  
Salva um arquivo Excel

78. Salvar Excel  
Salva um arquivo Excel (como '.xlsx', 'xlsm', '.xls' or '.csv')  na ruta indicada

79. Fechar XLSX  
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