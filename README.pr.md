



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

12. Cor da fonte  
Altere a cor da fonte do texto em uma célula ou intervalo de células. Pode ser uma cor predefinida ou personalizada

13. Obter cor da célula  
Obter a cor de uma célula. A função retornará uma lista com dois elementos: Background Color e Font Color no formato RGB.

14. Obter formato de célula  
Obtenha o formato de uma célula. A função retornará um dicionário com as propriedades da célula e o valor de cada uma.

15. Insertar Formula  
Inserta formula sobre una celda 

16. Inserir Macro a Excel  
Insere uma Macro a Excel

17. Selecionar e copiar Células  
Seleciona e copia células em Excel

18. Obter Célula Formato Moeda  
Obtém células com formato moeda

19. Obter Célula Formato Data  
Obtém células com formato de data

20. Copiar-Colar  
Copia um intervalo de células de uma planilha para outra

21. Formatar Célula  
Formatar Célula

22. Remover conteúdo  
Limpa fórmulas e valores do intervalo selecionado, mantendo o formato

23. Criar Planilha  
Adiciona uma planilha no final

24. Eliminar Planilha  
Elmina uma planilha

25. Copiar de um Excel para outro  
Copie o intervalo de um arquivo Excel para outro. Indicando o caminho do arquivo, abrirá o Excel para copiar ou colar os dados. Se você inserir o id de um Excel aberto, ele usará essa instância para copiar ou colar.

26. Adicionar/Eliminar Linha  
Adiciona ou elimina uma linha

27. Adicionar/Excluir Coluna  
Adiciona o exclui uma coluna

28. Converter CSV para XLSX  
Converte um documento CSV para formato XLSX

29. Exportar para JSON  
Exporta um array de dados para um arquivo JSON

30. (Descontinuado) Converter XLSX para CSV  
Converte um documento XLSX para CSV

31. Converter XLSX para CSV  
Converte um documento XLSX para CSV

32. Converter XLS para XLSX  
Converte um documento XLS para XLSX

33. Obter celula activa  
Obter linha e coluna de uma celula activa

34. Atualizar tabela dinâmica  
Atualiza uma tabela dinâmica. Descontinuado! Use o módulo PivotTableExcel

35. Ajustar células  
Ajusta, une, agrupa e desagrupa um intervalo de células. Você pode agrupar/desagrupar por linhas ou colunas

36. Obter Formula  
Obtém a fórmula numa célula

37. Adicionar Filtro Automático  
Adiciona filtro automático a uma tabela excel

38. Remover Filtro Automático  
Remova o filtro automático de uma planilha do Excel

39. Limpa Filtro  
Limpa todos os filtros feitos em uma planilha do Excel

40. Filtrar  
Filtre uma tabela do Excel de acordo com o valor relativo, conteúdo exato, cor de fundo ou cor da fonte das células. *Exemplos por tipo de filtro: xlAnd ['>=10'] or ['>=10', '<=20'] | xlOr ['<=10', '>=20'] | xlFilterValues ['10','20', '30'] | xlFilterCellColor (255,0,0) | xlFilterFontColor (255,0,0)*

41. Filtrar por Data  
Filtrar uma tabela por o dia, mes ou ano de uma data indicada

42. Filtro avançado  
Aplicar filtro avançado a uma tabela

43. Remover filtros  
Remova os filtros e mostre todos os dados

44. Renomear planilha  
Muda o nome de uma planilha de excel

45. Formato de texto  
Altere o alinhamento Horizontal ou Vertical de valores em um intervalo de células

46. Estilo de Célula  
Este comando modifica o formata a célula o intervalo de células selecionado. Você pode mudar a fonte e as bordas

47. Colar em Células  
Colar dados em células em Excel

48. Desativar modo de corte/cópia  
Desative o modo Cortar/Copiar do Excel ativo

49. Remover duplicatas  
Executa o comando remover duplicatas de Excel

50. Exportar para PDF avançado  
Exporta Excel para PDF com opções

51. Copiar-Mover Planilha  
Copiar ou mover uma planilha

52. Inserir Formulário  
Insere um Formulário no Excel

53. Ler células filtradas  
Leia todo o conteúdo das células filtradas e aplique formatação aos dados do tipo data, se indicado

54. Contar celulas filtradas  
Conta somente as celulas filtradas

55. Replace  
Run replace action to excel 

56. Ordenar  
Executa a ação de substituir de excel

57. Ordenar por múltiples niveles  
Ordene uma planilha Excel por valor, definindo vários níveis

58. Atualizar Tudo  
Atualiza todas as fontes do livro

59. Procurar  
Procura um texto no intervalo indicado e retorna a célula onde foi encontrada a primeira correspondência. Se não encontrar um valor, retornará vazio. Se o intervalo for filtrado, a pesquisa será realizada sobre as células visíveis.

60. Encontrar dados  
Retorna a primeira célula que corresponde aos dados da pesquisa

61. Bloquear celulas  
Bloquea ou desbloqueia celulas

62. Adicionar Gráfico  
Adiciona um novo gráfico sobre uma planilha de excel

63. Remover Senha  
Remove a senha e salva o Excel

64. Inserir imagem  
Inserir uma imagem

65. Exportar gráfico  
Exporta um gráfico por índice

66. Modo não visível  
Abre excel em modo não visível

67. Escrever array de objetos  
Escrever um array de objetos em células de Excel

68. Copiar-Colar Formato  
Copia formato de um intervalo de células de uma planilha para outra

69. Atualizar vínculos  
Muda um vínculo de um documento para outro

70. Desbloquear livro  
Desbloquea um livro com senha

71. Bloquear livro  
Bloquear um livro com senha

72. Desbloquear planilha  
Desbloquea uma folha com senha

73. Bloquear folha  
Bloquear uma folha com senha

74. Converter para .txt  
Converte para .txt

75. Texto em coluna  
Executa a opção texto em coluna de excel

76. Converter tempo de Excel para horas  
Converter tempo de Excel para horas. Retorna o formato como hh: mm: ss

77. Combinar planilhas  
Combine planilhas do Excel que estão na mesma pasta e que tenhamos o mesmo cabecalho. Combinar horizontalmente as planilhas da mesma planilha e verticalmente as planilhas diferentes.

78. Imprimir planilha  
Imprime uma planilha

79. Salvar Excel com senha  
Salva um arquivo Excel

80. Salvar Excel  
Salva um arquivo Excel (como '.xlsx', 'xlsm', '.xls', '.csv' or '.prn')  na ruta indicada

81. Fechar XLSX  
Fecha o arquivo aberto por Rocketbot. O comportamento de que apenas mate um arquivo, funciona se estiver aberto com o comando Abrir sem alertas, caso contrário, irá fechar todos.

82. Eliminar Estilos  
Remover estilos em uma planilha

83. Inserir link  
Inserir link de uma célula para uma planilha  




----
### OS

- windows
- mac

### Dependencies
- [**xlwings**](https://pypi.org/project/xlwings/)- [**pandas**](https://pypi.org/project/pandas/)
### License
  
![MIT](https://camo.githubusercontent.com/107590fac8cbd65071396bb4d04040f76cde5bde/687474703a2f2f696d672e736869656c64732e696f2f3a6c6963656e73652d6d69742d626c75652e7376673f7374796c653d666c61742d737175617265)  
[MIT](http://opensource.org/licenses/mit-license.ph)