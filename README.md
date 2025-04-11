# Transferência de Excel para PostgreSQL

Este programa permite automatizar a transferência de dados de múltiplos arquivos Excel para um banco de dados PostgreSQL. É especialmente útil para situações onde você precisa importar regularmente planilhas para um banco de dados PostgreSQL.

## Índice
1. [Requisitos](#requisitos)
2. [Instalação](#instalação)
3. [Configuração](#configuração)
4. [Uso](#uso)
5. [Personalização](#personalização)
6. [Estrutura do Código](#estrutura-do-código)
7. [Tratamento de Erros](#tratamento-de-erros)
8. [Relatório de Processamento](#relatório-de-processamento)

## Requisitos

- Python 3.6 ou superior
- PostgreSQL instalado e em execução
- Pacotes Python:
  - pandas
  - psycopg2
  - sqlalchemy
  - openpyxl (para arquivos .xlsx)
  - xlrd (para arquivos .xls antigos)

## Instalação

1. Clone ou baixe o script de transferência
2. Instale as dependências necessárias:

```bash
pip install pandas psycopg2-binary sqlalchemy openpyxl xlrd
```

## Configuração

Antes de utilizar o programa, você precisará configurar os parâmetros de conexão com o PostgreSQL e personalizar o tratamento dos dados conforme suas necessidades.

### 1. Parâmetros do Banco de Dados

No início do script, localize e preencha as configurações do banco de dados:

```python
CONFIG_BD = {
    "usuario": "postgres",    # Seu usuário PostgreSQL
    "senha": "",              # Sua senha PostgreSQL
    "host": "",               # Host (normalmente "localhost")
    "porta": "",              # Porta (normalmente "5432")
    "banco": ""               # Nome do banco de dados desejado
}
```

### 2. Definição da Tabela

A função `criar_tabela()` contém o SQL para criar a tabela que receberá os dados. **Você deve modificar essa parte** para definir a estrutura da sua tabela:

```python
def criar_tabela():
    # ...
    sql = """
    CREATE TABLE IF NOT EXISTS nome_da_sua_tabela (  
        # SUBSTITUA ESTA PARTE PELO ESQUEMA DA SUA TABELA
        # Exemplo:
        # id SERIAL PRIMARY KEY,
        # nome VARCHAR(100),
        # data_nascimento DATE,
        # salario NUMERIC(10,2)
    )
    """
    # ...
```

Substitua `nome_da_sua_tabela` pelo nome real da sua tabela e defina as colunas conforme o seu esquema de dados.

### 3. Mapeamento das Colunas

Na função `tratar_dados(df)`, configure o mapeamento entre os nomes das colunas do Excel e os nomes das colunas no PostgreSQL:

```python
def tratar_dados(df):
    mapeamento_colunas = {
        'NOME COLUNA': 'NOME_COLUNA',
        # Adicione mais mapeamentos:
        # 'Nome Original na Planilha': 'nome_coluna_no_banco',
    }
    # ...
```

### 4. Tratamento de Tipos de Dados

Ainda na função `tratar_dados(df)`, especifique quais colunas contêm datas e quais contêm valores numéricos:

```python
# Para colunas que contêm datas
for col in ['ADICIONE O NOME DAS COLUNAS QUE CONTEM DATAS']:
    # ...

# Para colunas que contêm valores numéricos
colunas_numericas = ['ADICIONE O NOME DAS COLUNAS QUE CONTEM VALORES NÚMERICOS']
# ...
```

### 5. Inserção de Dados

Na função `inserir_dados_psycopg2(df)`, verifique se a tabela mencionada corresponde à que você criou:

```python
sql = f"""
INSERT INTO nome_da_sua_tabela ({', '.join(colunas)})
VALUES ({', '.join(placeholders)})
"""
```

## Uso

1. Execute o script:
```bash
python nome_do_script.py
```

2. Confirme as configurações do banco de dados quando solicitado
3. Digite o caminho para o diretório que contém os arquivos Excel
4. O programa processará todos os arquivos .xlsx e .xls no diretório especificado

## Personalização

### Manipulação de Dados Específicos

Você pode ampliar a função `tratar_dados(df)` para realizar transformações específicas nos seus dados, como:

- Normalização de valores
- Preenchimento de dados ausentes
- Validação de formatos
- Conversão de unidades

Exemplo de personalização:

```python
def tratar_dados(df):
    # Código existente...
    
    # Exemplo: normalizar nomes para maiúsculas
    if 'nome' in df.columns:
        df['nome'] = df['nome'].str.upper()
    
    # Exemplo: tratar valores monetários (remover R$ e converter para decimal)
    if 'valor' in df.columns:
        df['valor'] = df['valor'].astype(str).str.replace('R$', '').str.replace('.', '').str.replace(',', '.').astype(float)
    
    return df
```

## Estrutura do Código

O programa está organizado nas seguintes seções principais:

1. **Configuração da conexão** - Parâmetros para conectar ao PostgreSQL
2. **Criação do banco e tabela** - Funções para preparar a estrutura de dados
3. **Processamento de dados** - Leitura, tratamento e inserção dos dados
4. **Relatório** - Geração de estatísticas sobre o processamento
5. **Programa principal** - Fluxo de execução e interação com o usuário

## Tratamento de Erros

O script inclui tratamento de erros em vários níveis:

- Verificação da existência do diretório de entrada
- Tratamento de erros de conexão com o banco de dados
- Tratamento de erros durante a criação de tabelas
- Tratamento de erros durante a leitura e processamento de planilhas
- Relatório detalhado de sucessos e falhas

## Relatório de Processamento

Ao final da execução, o programa gera um relatório detalhado que inclui:

- Arquivos processados com sucesso e quantidade de registros inseridos
- Arquivos que falharam durante o processamento e o motivo da falha
- Estatísticas gerais (total de arquivos, taxa de sucesso)

O relatório ajuda a identificar quais arquivos precisam de atenção manual ou reprocessamento.

---

## Resumo das Personalizações Necessárias

Para utilizar o programa, você deve:

1. ✅ Configurar os parâmetros de conexão ao PostgreSQL (`CONFIG_BD`)
2. ✅ Definir a estrutura da tabela na função `criar_tabela()`
3. ✅ Configurar o mapeamento de colunas na função `tratar_dados(df)`
4. ✅ Especificar colunas de data e numéricas na função `tratar_dados(df)`
5. ✅ Verificar o nome da tabela na função `inserir_dados_psycopg2(df)`

Após essas personalizações, o script estará pronto para processar seus arquivos Excel e inseri-los no PostgreSQL.
