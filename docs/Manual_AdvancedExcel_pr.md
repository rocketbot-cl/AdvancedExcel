



# Opções avançadas para Excel
  
Módulo com opções avançadas para Excel

![banner](imgs/Banner_AdvancedExcel.png)
## Como instalar este módulo
  
__Baixe__ e __instale__ o conteúdo na pasta 'modules' no caminho do Rocketbot  

## Descrição do comando

### Abrir sem alertas
  
Abre um arquivo sem mostrar alertas do MS Excel.
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Direcção do arquivo XLSX|Direcação do arquivo xlsx que se deseja abrir|arquivo.XLSX|
|Password (opcional)|Password do arquivo xlsx|P@ssW0rd|
|Identificador (opcional)|Nome ou identificador para o arquivo que se abrirá. É utilizado quando se precisa abrir mais de um excel. Por padrão é *default*.|id|
|Atribuir resultado a variável|Variável onde o resultado será armazenado|id|

### Buscar e conectar
  
Busca um excel aberto e conecta-se a este.
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Nome do arquivo XLSX aberto||Arquivo.XLSX|
|Identificador (opcional)|Nome ou identificador para o arquivo que será aberto. É utilizado quando se precisa abrir mais de um excel. Por padrão é *default*.|excel1|

### Contar Colunas
  
Conta o número de colunas do excel aberto. É necessário que o excel esteja salvo para tomar os últimos cambios
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Planilha |Nome da planilha onde os dados estão localizados|Planilha 1|
|Obter nome da coluna|Se marcar esta caixa, devolverá a letra da última coluna|True|
|Atribuir resultado a variável |Nome da variável para armazenar o resultado|numero_colunas|

### Contar Linhas
  
Conta todas as linhas ou dentro de um intervalo.
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Planilha|Nome da planilha onde os dados estão localizados|Planilha 1|
|Contar todas as linhas|Opção para contar todas as linhas.||
|Coluna|Coluna onde as linhas serão contadas|C|
|Atribuir resultado a variável|Nome da variável para armazenar o resultado|numero_linhas|

### Cor da célula
  
Muda a cor de uma célula ou intervalo de células. Pode ser uma cor por defeito ou uma personalizada
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Células |Célula ou intervalo de células. A sintaxe deve ser a mesma do excel (A1 ou A1B1) |A1:B5|
|Cor da célula em RGB |Valores rgb da cor que terá a célula ou células|250,250,250|
|Seleccione cor |Seleccione a cor. Pode usar o campo anterior para personalizar a cor|red|

### Obter cor da célula
  
Obter a cor de uma célula. A função retornará uma lista com dois elementos: Background Color e Font Color no formato RGB.
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Folha |Folha|Folha1|
|Célula |Célula. A sintaxe deve ser a mesma do excel (A1)|A1|
|Atribuir a variável|Nome da variável para armazenar o resultado|cor|

### Insertar Formula
  
Inserta formula sobre una celda 
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Folha |Folha|Folha5|
|Celda |Célula ou intervalo de células. A sintaxe deve ser a mesma do excel (A1 ou A1B1) |A5|
|Escreva fórmula |Fórmula a ser inserida. Deve ser escrito em inglês. Lembre-se de usar *,* para separar parâmetros|=SUM(A1:A4)|

### Inserir Macro a Excel
  
Insere uma Macro a Excel
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Caminho da Macro|Caminho do arquivo .bas que se quer inserir|Macro.bas|

### Selecionar Células
  
Seleciona células em Excel
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Planilha|Nome da planilha a ser automatizada|Planilha 1|
|Digite células a selecionar|Célula ou intervalo de células. A sintaxe deve ser a mesma do excel (A1 ou A1B1) |A1:B3|
|Copiar|Ao marcar a caixa, os valores serão copiados para a prancheta.|True|

### Obter Célula Formato Moeda
  
Obtém células com formato moeda
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Planilha|Nome da planilha a ser automatizada|Planilha 1|
|Insira células a selecionar|Célula ou intervalo de células. A sintaxe deve ser a mesma do excel (A1 ou A1B1) |A1:B3|
|Atribuir a variável|Nome da variável para armazenar o resultado|variável|

### Obter Célula Formato Data
  
Obtém células com formato de data
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Planilha|Nome da planilha a ser automatizada|Planilha 1|
|Entre as celulas a selecionar|Célula ou intervalo de células. A sintaxe deve ser a mesma do excel (A1 ou A1B1) |A1:B3|
|Atribuir a variável|Nome da variável para armazenar o resultado|variável|

### Copiar-Colar
  
Copia um intervalo de células de uma planilha para outra
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Planilha de origem|Nome da planilha a ser automatizada|Folha1|
|Intervalo a copiar|Célula ou intervalo de células para copiar. A sintaxe deve ser a mesma do excel (A1 ou A1B1) |A1:C4|
|Planilha de destino|Nome da planilha de destino|Folha2|
|Intervalo para colar|Célula ou intervalo de células para colar. A sintaxe deve ser a mesma do excel (A1 ou A1B1) |A1:C4|
|Opção de Colar|Selecionar tipo de colagem para a célula ou intervalo de células.|Opção|
|Operação de Colar|Selecione a operação de colagem para a célula ou intervalo de células.|Operação|
|Pular espaços em branco||Impede a substituição de valores na área de colagem quando células em branco são produzidas na área de cópia quando esta caixa é selecionada.|
|Transpor||Rotate the content of copied cells when pasting. Data in rows will be pasted into columns and vice versa.|

### Formatar Célula
  
Formatar Célula
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Nome da Planilha|Nome da planilha a ser automatizada|Folha1|
|Célula a formatar|Célula ou intervalo de células. A sintaxe deve ser a mesma do excel (A1 ou A1B1) |A1:C4|
|Formato|O tipo de formato para a célula deve ser selecionado. Selecione o formato personalizado para adicionar um formato personalizado|dd-mm-yy|
|Formato personalizado|Formato personalizado. Deve ser o mesmo que mostrado na seção personalizada do Excel.|00000|
|Texto para valor|||

### Criar Planilha
  
Adiciona uma planilha no final
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Nome da planilha|Nome da planilha a ser criada|Folha2|
|Depois de|A planilha será criada ao lado da planilha indicada neste campo.|Folha1|

### Eliminar Planilha
  
Elmina uma planilha
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Nome da planilha|Nome da planilha a ser excluida|Folha2|
|Atribuir resultado a variável|Nome da variável para armazenar o resultado|Variável|

### Copiar de um Excel para outro
  
Copia um intervalo de um Excel para outro, o excel de destino não deve estar aberto
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Excel de origem|Caminho do arquivo xlsx de origen|Folha1|
|Planilha de origem|Nome da planilha de origen|Folha1|
|Intervalo a copiar|Célula ou intervalo de células. A sintaxe deve ser a mesma do excel (A1 ou A1B1) |A1:D7|
|Excel de destino|Caminho do arquivo xlsx de destino|Folha1|
|Planilha de destino|Nome da planilha onde vai ser colada|Folha1|
|Intervalo onde colar|Célula ou intervalo de células. A sintaxe deve ser a mesma do excel (A1 ou A1B1) |A1:D7|
|Sólo valores|Se esta caixa foi marcada, copiará apenas os valores|True|

### Adicionar/Eliminar Linha
  
Adiciona ou elimina uma linha
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Opção|Selecione Add para adicionar uma linha ou Delete para excluir uma linha.|Adicionar|
|Nome da Planilha|Nome da planilha onde acrescentar a fila|Planilha|
|Número da Linha|Indicar a(s) linha(s) a ser(em) adicionada(s) ou deletada(s)|2|
|Onde Inserir|Indicar onde adicionar o excluir a linha|A1:D7|

### Adicionar/Excluir Coluna
  
Adiciona o exclui uma coluna
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Opção|Selecione Add para adicionar uma coluna ou Delete para excluir uma coluna.||
|Nome da Planilha|Nome da planilha onde os dados estão localizados|Planilha|
|Coluna|Indicar a(s) coluna(s) a ser(em) adicionada(s) ou deletada(s)|B|

### Converter CSV para XLSX
  
Converte um documento CSV para XLSX
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Caminho do arquivo CSV|Direcação do arquivo csv que se quer converter||
|Delimitador|Delimitador da arquivo csv||
|Tem cabeçeras?|marque esta caixa se o csv tiver cabeçalhos|True|
|Codificação|Digite o tipo de codificação do arquivo. O padrão é latino-1|utf-8|
|Caminho do arquivo XLSX|Direcação do arquivo xlsx onde será salvo|file.xlsx|

### Converter XLSX para CSV
  
Converte um documento XLSX para CSV
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Caminho do arquivo XLSX|Caminho do arquivo xlsx que se quer converter|C:/Users/User/Desktop/file.xlsx|
|Delimitador|Delimitador da arquivo csv|,|
|Nome da planilha|Nome da planilha onde os dados estão localizados|Sheet0|
|Direcação do arquivo CSV|Direção do arquivo csv onde será salvo|C:/Users/User/Desktop/file.csv|

### Converter XLS para XLSX
  
Converte um documento XLS para XLSX
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Direcão do arquivo XLS|Direção do arquivo xls que se quer converter|C:\Users\User\Desktop\file.xls|
|Direção do arquivo XLSX|Direção onde se guardará o arquivo xlsx|C:\Users\User\Desktop\new_file.xlsx|

### Obter celula activa
  
Obter linha e coluna de uma celula activa
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Atribuir resultado a variável |Nome da variável para armazenar o resultado|Variável|

### Atualizar tabela dinâmica
  
Atualiza uma tabela dinâmica. Descontinuado! Use o módulo PivotTableExcel
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Planilha |Nome da planilha onde a tabela está localizada|Planilha 1|
|Nome da tabela dinâmica |Nome da tabela dinámica que vai ser actualizada|Nome: |

### Ajustar células
  
Ajusta, une, agrupa e desagrupa um intervalo de células. Você pode agrupar/desagrupar por linhas ou colunas
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Planilha |Nome da planilha onde os dados estão localizados|Planilha 1|
|Intervalo a ajustar|Célula ou intervalo de células. A sintaxe deve ser a mesma do excel (A1 ou A1B1) |A1:D7|
|Autofit|Ajusta automaticamente as células para exibir os dados||
|Agrupar linhas|Ao marcar esta opção, as linhas serão agrupadas na faixa selecionada.||
|Agrupar colunas|Ao marcar esta opção, as colunas serão agrupadas na faixa selecionada.||
|Desagrupar linhas|Ao marcar esta opção, as linhas serão desagrupadas na faixa selecionada.||
|Desagrupar colunas|Ao marcar esta opção, as colunas serão desagrupadas na faixa selecionada.||
|Mesclar células|Marcar esta caixa de seleção mesclará as células no intervalo selecionado||
|Nível de linha|Ao marcar esta caixa, será exibido o número especificado de níveis de linha.|2|
|Faixa de coluna|Ao marcar esta caixa, será exibido o número especificado de níveis de coluna.|2|
|Largura da coluna|Largura na qual a coluna se ajustará|20|
|Altura da linha|Altura à qual a linha se ajustará|20|

### Obter Formula
  
Obtém a fórmula numa célula
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Célula |Célula onde fica a formula. A sintaxe deve ser a mesma do excel (A1 ou A1B1) |A5|
|Atribuir resultado a variável |Nome da variável para armazenar o resultado|Variável|

### Adicionar Filtro Automático
  
Adiciona filtro automático a uma tabela excel
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Planilha |Nome da planilha onde os dados estão localizados|Planilha 1|
|Intervalo |Célula ou intervalo de células. A sintaxe deve ser a mesma do excel (A1 ou A1B1) |A1:E6 |

### Filtrar
  
Filtra a uma tabela excel
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Planilha |Nome da planilha onde os dados estão localizados|Planilha 1|
|Início da tabela |Coluna onde começa a tabela a ser filtrada|A |
|Coluna |Coluna onde adicionar o filtro|A |
|Filtro |Filtro ou lista de filtros a adicionar. Use "=" para encontrar campos em branco, "<>" para células não vazias e negação de dados.|['filtro1','filtro2', 'filtro3']|

### Renomear planilha
  
Muda o nome de uma planilha de excel
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Planilha |Nome da planilha a ser renomeada|Planilha 1|
|Novo nome |Novo nome da planilha|novo_nome|

### Formato de texto
  
Altere o alinhamento Horizontal ou Vertical de valores em um intervalo de células
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Planilha |Nome da planilha onde os dados estão localizados|Planilha 1|
|Célula o intervalo de células|Intervalo que contém os dados a alinhar|A1:D7|
|Alinhamento horizontal|Selector que contém as opções de alinhamento horizontal||
|Alinhamento Vertical|Selector que contém as opções de alinhamento vertical||

### Estilo de Célula
  
Este comando modifica o formata a célula o intervalo de células selecionado. Você pode mudar a fonte e as bordas
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Nome da Planilha|Nome da planilha a ser automatizada|Planilha1|
|Intervalo a formatar|Célula ou intervalo de células. A sintaxe deve ser a mesma do excel (A1 ou A1B1) |A1:C4|
|Borda|Borda da célula a ser formatada|Contour|
|Estilo|Estilo da borda da célula a ser formatada|_ _ _ _ _ _ _ _ _ _ _|
|Tamanho da fonte|Tamanho da fonte da célula|20|
|Negrita|Seleccione esta caixa para cambiar o texto em negrito|True|
|Cursiva|Seleccione esta caixa para colocar o texto em itálico|True|
|Sublinhar|Seleccione esta caixa para sublinhar o texto|True|
|Ajustar Texto||True|

### Colar em Células
  
Colar dados em células em Excel
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Planilha|Nome da planilha a ser automatizada|Planilha 1|
|Ingrese células onde colar|Célula ou intervalo de células onde vai colar. A sintaxe deve ser a mesma do excel (A1 ou A1B1) |A1:B3|
|Só valores|Se esta caixa foi marcada, vai colar apenas os valores|True|

### Remover duplicatas
  
Executa o comando remover duplicatas de Excel
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Planilha|Nome da planilha a ser automatizada|Planilha 1|
|Ingrese células onde filtrar|Célula ou intervalo de células. A sintaxe deve ser a mesma do excel (A1 ou A1B1) |A1:B3|
|Coluna |Indicar a coluna onde as duplicatas serão procuradas.|A |
|Tem cabeçeras?|marque esta caixa se o excel tiver cabeçalhos|True|

### Exportar para PDF avançado
  
Exporta Excel para PDF com opções
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Salvar PDF|Caminho onde salvar o arquivo .pdf|/Users/user/Desktop/excel.pdf|
|Ajuste Automático|||
|Zoom|||
|Ajustar Altura|||
|Ajustar Largura|||

### Copiar-Mover Planilha
  
Copiar ou mover uma planilha
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Planilha origem|Nome da planilha de origen|Sheet1|
|Mover/copiar antes da planilha|Nome da planilha onde vai ser movida|Sheet2|
|Excel destino|Caminho do arquivo .xlsx onde mover ou copiar a planilha|C:/ruta/para/excel.xlsx|
|Copy |Ao marcar a caixa, a planilha vai ser copiada.||

### Inserir Formulário
  
Insere um Formulário no Excel
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Caminho do Formulário |Direcação do arquivo frm que se deseja inserir|Form.frm|

### Ler células filtradas
  
Ler somente as células filtradas
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Planilha |Nome da planilha onde os dados estão localizados|Sheet1|
|Intervalo |Célula ou intervalo de células. A sintaxe deve ser a mesma do excel (A1 ou A1B1) |A1:B100 |
|Atribuir resultado a variável |Nome da variável para armazenar o resultado|Variável|
|Mais dados |||

### Contar celulas filtradas
  
Conta somente as celulas filtradas
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Planilha |Nome da planilha onde os dados estão localizados|Planilha 1|
|Intervalo onde buscar |Intervalo de coluna filtrada (A1A100)|A1:A100 |
|Atribuir resultado a variável |Nome da variável para armazenar o resultado|Variável|

### Replace
  
Run replace action to excel 
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Planilha |Nome da planilha onde os dados estão localizados|Folha1|
|Intervalo onde buscar |Célula ou intervalo de células. A sintaxe deve ser a mesma do excel (A1 ou A1B1) |A1:B100 |
|Palavra a substituir |Palavra a ser procurada para ser substituída|10/10/2020|
|Nova palavra |Palavra que substituirá a anterior indicada|10-10-2020|

### Ordenar
  
Executa a ação de substituir de excel
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Planilha |Nome da planilha onde os dados estão localizados|Planilha 1|
|Intervalo onde buscar |Célula ou intervalo de células. A sintaxe deve ser a mesma do excel (A1 ou A1B1) |A1:B100 |
|Coluna|Indicar a coluna que vai ser classificada|A1:A22|
|Tipo de ordem |Indicar como a coluna vai ser classificada|Ascendente|

### Atualizar Tudo
  
Atualiza todas as fontes do livro
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |

### Buscar
  
Devuelve a primeira celula encontrada
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Planilha |Nome da planilha onde os dados estão localizados|Folha1 |
|Intervalo onde buscar |Célula ou intervalo de células. A sintaxe deve ser a mesma do excel (A1 ou A1B1) |A1:B100 |
|Texto a buscar|Texto a ser procurado no excel|Lorem|
|Atribuir resultado a variável |Nome da variável para armazenar o resultado|Variável|

### Bloquear celulas
  
Bloquea ou desbloqueia celulas
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Sheet |Nome da planilha onde os dados estão localizados|Sheet1|
|Intervalo onde buscar |Célula ou intervalo de células. A sintaxe deve ser a mesma do excel (A1 ou A1B1) |A1:B100 |
|Action|Selecione se você deseja travar ou destravar uma célula|Lock|

### Adicionar Gráfico
  
Adiciona um novo gráfico sobre uma planilha de excel
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Planilha |Nome da planilha onde os dados estão localizados|Planilha 1|
|Tipo de Gráfico|Selecione o tipo de gráfico a ser inserido no Excel|Linha|
|Célula onde inserir gráfico |Célula onda vai ser inserido o gráfico. A sintaxe deve ser a mesma do excel (A1) |A1|
|Intervalo de dados |Célula ou intervalo de células. A sintaxe deve ser a mesma do excel (A1 ou A1B1) |A1:B100 |

### Remover Senha
  
Remove a senha e salva o Excel
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Excel com senha|Caminho do arquivo xlsx que se deseja abrir|C:/Users/User/Desktop/test.xlsx|
|Senha|Senha do arquivo xlsx|****|
|Excel sem senha|Caminho onde salvar o arquivo .xlsx|C:/Users/User/Desktop/test2.xlsx|

### Inserir imagem
  
Inserir uma imagem
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Planilha |Nome da planilha onde os dados estão localizados|Planilha 1|
|Célula|Célula onda vai ser inserido a imagem. A sintaxe deve ser a mesma do excel (A1) |B5|
|Direcação da imagem |Direção do arquivo de imagem que se quer inserir|imagem.png|

### Exportar gráfico
  
Exporta um gráfico por índice
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Planilha |Nome da planilha onde os dados estão localizados|Planilha 1|
|Index |Índice do gráfico a ser exportado|1|
|Direcação da imagem |Direção onde a imagem será salva|/direção/para/imagem.png|

### Modo não visível
  
Abre excel em modo não visível
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Direcação do arquivo XLSX|Direção do arquivo xlsx que se deseja abrir|Arquivo.XLSX|
|Identificador (opcional)|Nome ou identificador para o arquivo que será aberto. É utilizado quando se precisa abrir mais de um excel. Por padrão é *default*.|default|

### Escrever array de objetos
  
Escrever um array de objetos em células de Excel
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Planilha |Nome da planilha onde os dados estão localizados|Planilha 1|
|Célula o Rango de Células |Célula ou intervalo de células. A sintaxe deve ser a mesma do excel (A1 ou A1B1) |A1|
|Dados a escrever|Célula ou intervalo de células. A sintaxe deve ser a mesma do excel (A1 ou A1B1) |[{ 'id',: 1, 'text': 'olá' },{ 'id',: 2, 'text': 'mundo' }]|

### Copiar-Colar Formato
  
Copia formato de um intervalo de células de uma planilha para outra
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Planilha de origem|Nome da planilha de origem|Folha1|
|Intervalo a copiar||A1:C4|
|Planilha de destino|Nome da planilha do destino|Folha2|
|Intervalo onde colar||A1:C4|

### Atualizar vínculos
  
Muda um vínculo de um documento para outro
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Direcação|Direcação do arquivo xlsx que se quere atualizar||
|Direcação atualizada|Direcação do arquivo xlsx que substituirá o vínculo|file.xlsx|

### Desbloquear planilha
  
Desbloquea uma planilha com senha
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Planilha|Nome da planilha a ser bloqueada|Planilha 1|
|Senha|Senha da planilha bloqueada|Senha|

### Converter para .txt
  
Converte para .txt
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Direcação do arquivo XLSX|Direcção do arquivo xlsx que se quer converter|Arquivo.XLSX|
|Salvar TXT|Caminho onde salvar o arquivo .txt|/Users/user/Desktop/test.txt|

### Texto em coluna
  
Executa a opção texto em coluna de excel
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Planilha |Nome da planilha onde os dados estão localizados|Planilha 1|
|Intervalo |Célula ou intervalo de células. A sintaxe deve ser a mesma do excel (A1 ou A1B1) |A1:B100 |
|Seleciona separador |Seleciona o separador de células, pode ser largura fixa ou delimitado||
|Selecione o tipo de separador |Seleciona o tipo de separador||
|Outro delimitador ou largura |Escreva o delimitador ou largura fixa|| ou 20,35,22,10|

### Converter tempo de Excel para horas
  
Converter tempo de Excel para horas. Retorna o formato como hh: mm: ss
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Insere o tempo no formato decimal ||0.296655812|
|Atribuir resultado a variável |Nome da variável para armazenar o resultado|Variável|

### Imprimir planilha
  
Imprime uma planilha
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Planilha |Nome da planilha a ser impresso|Planilha 1|

### Salvar Excel com senha
  
Salva um arquivo Excel
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Salvar Excel como|Caminho onde salvar o arquivo .xlsx|/Users/user/Desktop/excel.xlsx|
|Digite a senha|Senha do arquivo xlsx|password|

### Salvar Excel
  
Salva um arquivo Excel na ruta indicada
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Salvar Excel|Caminho onde salvar o arquivo .xlsx|/Users/user/Desktop/excel.xlsx|

### Fechar XLSX
  
Fecha o arquivo aberto por Rocketbot
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
