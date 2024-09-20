# Gerenciamento de Transações - XYZ Administradora de Cartões

## Descrição
Este projeto foi desenvolvido como parte de um desafio técnico para a XYZ Administradora de Cartões de Crédito. Ele visa gerenciar transações de cartões de crédito de forma eficiente, utilizando uma aplicação em **VB6** e **SQL Server** para criar, editar, excluir e consultar transações. Além disso, inclui funcionalidades de geração de relatórios e cálculo de totais via stored procedures e funções no banco de dados.

## Funcionalidades
A aplicação desenvolvida oferece as seguintes funcionalidades:

### CRUD de Transações
- **Inserção de Transações**: Permite registrar novas transações de cartões de crédito com os seguintes campos: `Id_Transacao`, `Numero_Cartao`, `Valor_Transacao`, `Data_Transacao`, `Descricao`.
- **Edição de Transações**: Modificação de transações já existentes.
- **Exclusão de Transações**: Remoção de transações específicas.
- **Consulta de Transações**: Permite filtrar transações por `Numero_Cartao`, `Data_Transacao` e `Valor_Transacao` usando uma interface simples com caixas de texto e uma DataGrid para exibição.

### Stored Procedure - Cálculo de Totais
Uma stored procedure foi criada no SQL Server para calcular o total de transações em um determinado período, agrupando os resultados por `Numero_Cartao`. A procedure recebe dois parâmetros: `Data_Inicial` e `Data_Final`, e retorna:
- `Numero_Cartao`
- `Valor_Total` (total das transações no período)
- `Quantidade_Transacoes` (número de transações no período)

### Função SQL - Categorizar Transações
Uma função no SQL Server foi desenvolvida para categorizar as transações de acordo com o valor:
- **Alta**: Transações acima de R$1000.
- **Média**: Transações entre R$500 e R$1000.
- **Baixa**: Transações abaixo de R$500.

### View - Informações Combinadas
Uma view no SQL Server foi criada para combinar as informações de **Transações** e **Clientes**, retornando:
- `Nome_Cliente`
- `Numero_Cartao`
- `Valor_Transacao`
- `Data_Transacao`
- `Categoria` da transação, usando a função criada anteriormente.

### Relatório em Excel
A aplicação inclui uma funcionalidade para exportar as transações do último mês em um arquivo Excel, que contém as colunas:
- `Numero_Cartao`
- `Valor_Transacao`
- `Data_Transacao`
- `Descricao`
- `Categoria` da transação

O arquivo Excel é salvo em um diretório informado pelo usuário através de uma caixa de diálogo.

## Tecnologias Utilizadas
- **VB6**: Utilizado para o desenvolvimento da interface e funcionalidades do CRUD.
- **SQL Server**: Utilizado para armazenar os dados e implementar as stored procedures, funções e views.
- **Excel**: Geração de relatórios com exportação das transações financeiras.

## Estrutura do Projeto
O projeto é composto pelos seguintes módulos principais:

### Interface de Usuário
- Interface desenvolvida em **VB6** com caixas de texto para inserção e consulta de dados, botões para as operações de CRUD, e uma **DataGrid** para exibição dos resultados.

### Banco de Dados
- **Tabela de Transações**: Armazena as informações de cada transação.
- **Stored Procedures e Funções**: Implementadas para realizar cálculos e categorizações de transações de forma eficiente.
- **View**: Combina dados de transações e clientes, categorizando as transações.

### Relatório
- Relatório exportado para Excel, com transações filtradas pelo último mês.

## Requisitos de Avaliação
- **Qualidade do Código**: Segui boas práticas de programação tanto no VB6 quanto nas queries SQL.
- **Eficiência**: As queries SQL foram otimizadas para lidar com grandes volumes de dados.
- **Interface Simples e Funcional**: A interface da aplicação é clara e de fácil uso, permitindo o gerenciamento de transações de forma eficiente.
- **Relatórios Excel**: Correta exportação e formatação dos dados para arquivo Excel.

## Conclusão
Este projeto entrega uma solução completa para o gerenciamento de transações financeiras, com funcionalidades robustas e integração eficiente com o SQL Server, além de permitir exportação de relatórios em Excel.


