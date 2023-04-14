



# Opções avançadas para Excel
  
Module with advanced options to work with files with Microsoft Excel  

## Como instalar este módulo
  
Para instalar o módulo no Rocketbot Studio, pode ser feito de duas formas:
1. Manual: __Baixe__ o arquivo .zip e 
descompacte-o na pasta módulos. O nome da pasta deve ser o mesmo do módulo e dentro dela devem ter os seguintes arquivos e pastas: \__init__.py, package.json, docs, example e libs. Se você tiver o aplicativo aberto, atualize seu navegador para poder usar o novo módulo.
2. Automático: Ao entrar no Rocketbot Studio na margem direita você encontrará a seção **Addons**, selecione **Install Mods**, procure o módulo desejado e aperte instalar.  




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

44. (Descontinuado) Pesquisar  
Devuelve a primeira celula encontrada

45. Encontrar dados  
Retorna a primeira célula que corresponde aos dados da pesquisa

46. Bloquear celulas  
Bloquea ou desbloqueia celulas

47. Adicionar Gráfico  
Adiciona um novo gráfico sobre uma planilha de excel

48. Remover Senha  
Remove a senha e salva o Excel

49. Inserir imagem  
Inserir uma imagem

50. Exportar gráfico  
Exporta um gráfico por índice

51. Modo não visível  
Abre excel em modo não visível

52. Escrever array de objetos  
Escrever um array de objetos em células de Excel

53. Copiar-Colar Formato  
Copia formato de um intervalo de células de uma planilha para outra

54. Atualizar vínculos  
Muda um vínculo de um documento para outro

55. Desbloquear planilha  
Desbloquea uma folha com senha

56. Bloquear folha  
Bloquear uma folha com senha

57. Converter para .txt  
Converte para .txt

58. Texto em coluna  
Executa a opção texto em coluna de excel

59. Converter tempo de Excel para horas  
Converter tempo de Excel para horas. Retorna o formato como hh: mm: ss

60. Imprimir planilha  
Imprime uma planilha

61. Salvar Excel com senha  
Salva um arquivo Excel

62. Salvar Excel  
Salva um arquivo Excel na ruta indicada

63. Fechar XLSX  
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