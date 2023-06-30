# Detalhes Técnicos de Implementação

O objetivo aqui e dar uma explicação um pouco mais aprofundada de alguns processos realizados pelo Add-In.

## 1. Como o Add-In grava na memória a base de dados que deve ser conectada

O caminho do arquivo Access é gravado como uma propriedade personalizada no excel.

Quando realizamos a primeira conexão, onde selecionamos o arquivo access que queremos conectar, o programa grava o caminho do arquivo em uma propriedade personalizada chamada “LinktoAccess”. Após realizar a conexão, é possível ver essa propriedade indo em Arquivo > Informações > Propriedades > Propriedades Avançadas. Na tela, ir na aba Personalizar, e poderemos ver a propriedade gravada nesta planilha:

![LinkToAccess](imgs_detalhes\LinkToAccess.png)

Deste modo, quando o programa não encontra essa propriedade, ele irá perguntar se deseja conectar com uma base de dados, como foi é feito no ínicio do procedimento. Uma vez gravado, o programa busca o valor desta propriedade (No caso, o caminho do arquivo access) e tenta realizar a conexão.

Se caso algum nome de pasta ou nome do arquivo access seja alterado, a conexão precisará ser refeita. Desta forma o programa irá atualizar o valor desta propriedade para o caminho atualizado.

## 2. Como o caminho dos arquivos ou da pasta são gravados para realizar as atualizações

Para gravar o local da pasta ou do arquivo, o programa também grava o caminho em uma propriedade, porém em uma propriedade chamada “TExto de validação” da tabela criada no access.

É possível verificar essa propriedade abrindo o arquivo access, clicando com o botão direito do mouse na tabela e selecionando o “Modo Design”. Por fim, clicando na opção folha de propriedades, é possível ver o caminho gravado na tabela. Segue exemplo:

![Texto Validação](imgs_detalhes\caminho_tabela_access.png)

Quando é realizada a primeira conexão, o programa cria a tabela no access e grava o caminho do arquivo ou pasta nesta propriedade. Ao selecionar a opção “Atualizar dados já conectados”, o programa irá fazer a leitura de todas as tabelas no access, e para cada tabela, buscará o caminho de origem dos dados. Com isso, ela faz a atualização excluindo a tabela e recriando-a com o mesmo nome e dados com os respectivos dados do caminho gravando nas propriedades.