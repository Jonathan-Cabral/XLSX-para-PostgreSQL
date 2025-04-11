import os
import pandas as pd
import psycopg2
from sqlalchemy import create_engine, text

CONFIG_BD = {
    "usuario": "postgres",
    "senha": "",
    "host": "",
    "porta": "",
    "banco": ""
}

def conectar_psycopg2(banco="postgres"):
    try:
        conn = psycopg2.connect(
            user=CONFIG_BD['usuario'],
            password=CONFIG_BD['senha'],
            host=CONFIG_BD['host'],
            port=CONFIG_BD['porta'],
            database=banco
        )
        return conn
    except Exception as e:
        print(f"Erro ao conectar ao PostgreSQL usando psycopg2: {str(e)}")
        raise

def criar_banco_dados():
    try:
        conn = conectar_psycopg2("postgres")
        conn.autocommit = True
        cursor = conn.cursor()
        
        cursor.execute(f"SELECT 1 FROM pg_database WHERE datname = '{CONFIG_BD['banco']}'")
        exists = cursor.fetchone()
        
        if not exists:
            print(f"Criando banco de dados '{CONFIG_BD['banco']}'...")
            cursor.execute(f"CREATE DATABASE {CONFIG_BD['banco']}")
            print(f"Banco de dados '{CONFIG_BD['banco']}' criado com sucesso!")
        else:
            print(f"Banco de dados '{CONFIG_BD['banco']}' já existe.")
        
        cursor.close()
        conn.close()
    except Exception as e:
        print(f"Erro ao criar banco de dados: {str(e)}")
        raise

def conectar_sqlalchemy():
    try:
        conn_string = f"postgresql://{CONFIG_BD['usuario']}:{CONFIG_BD['senha']}@{CONFIG_BD['host']}:{CONFIG_BD['porta']}/{CONFIG_BD['banco']}"
        engine = create_engine(conn_string)
        
        with engine.connect() as conn:
            conn.execute(text("SELECT 1"))
            
        print(f"Conexão com o banco de dados '{CONFIG_BD['banco']}' estabelecida com sucesso!")
        return engine
    except Exception as e:
        print(f"\nErro detalhado de conexão: {str(e)}")
        print("\nVerifique se:")
        print("1. O serviço PostgreSQL está em execução")
        print("2. As credenciais configuradas no início do script estão corretas")
        print("3. O banco de dados especificado existe")
        print("4. O host e porta estão acessíveis")
        print("5. As permissões de usuário são adequadas\n")
        raise

def criar_tabela():
    try:
        conn = conectar_psycopg2(CONFIG_BD['banco'])
        conn.autocommit = True
        cursor = conn.cursor()
        
        sql = """
        CREATE TABLE IF NOT EXISTS nome_da_sua_tabela (  

        )
        """
        
        cursor.execute(sql)
        print("Tabela 'nome_da_sua_tabela' criada ou já existente!")
        
        cursor.close()
        conn.close()
    except Exception as e:
        print(f"Erro ao criar tabela: {str(e)}")
        raise

def tratar_dados(df): #converte o nome das tabelas para o padrão do postgre
    mapeamento_colunas = {
        'NOME COLUNA': 'NOME_COLUNA', #Exemplo de conversão
    }
    
    for coluna_original, coluna_nova in mapeamento_colunas.items():
        if coluna_original in df.columns:
            df = df.rename(columns={coluna_original: coluna_nova})
    
    for col in ['ADICIONE O NOME DAS COLUNAS QUE CONTEM DATAS']:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors='coerce')
    
    colunas_numericas = ['ADICIONE O NOME DAS COLUNAS QUE CONTEM VALORES NÚMERICOS']
    for col in colunas_numericas:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')
    
    return df

def inserir_dados_psycopg2(df):
    conn = None
    cursor = None
    registros_inseridos = 0
    
    try:
        conn = conectar_psycopg2(CONFIG_BD['banco'])
        cursor = conn.cursor()
        
        for _, row in df.iterrows():
            colunas = []
            placeholders = []
            valores = []
            
            for coluna, valor in row.items():
                if pd.isna(valor):
                    continue
                
                colunas.append(coluna)
                placeholders.append("%s")
                valores.append(valor)
            
            if not colunas:
                continue
                
            sql = f"""
            INSERT INTO nome_da_sua_tabela ({', '.join(colunas)})
            VALUES ({', '.join(placeholders)})
            """
            
            cursor.execute(sql, valores)
            registros_inseridos += 1
        
        conn.commit()
        return registros_inseridos
        
    except Exception as e:
        if conn:
            conn.rollback()
        print(f"Erro ao inserir dados: {str(e)}")
        raise
    finally:
        if cursor:
            cursor.close()
        if conn:
            conn.close()

def processar_planilhas(diretorio):
    resultados = {
        'sucesso': [],
        'falha': []
    }
    
    arquivos_excel = [f for f in os.listdir(diretorio) if f.endswith(('.xlsx', '.xls'))]
    total_arquivos = len(arquivos_excel)
    
    if total_arquivos == 0:
        print("Nenhum arquivo Excel (.xlsx ou .xls) encontrado no diretório!")
        return resultados
    
    print(f"Encontrados {total_arquivos} arquivos Excel para processar.")
    
    for i, arquivo in enumerate(arquivos_excel, 1):
        caminho_completo = os.path.join(diretorio, arquivo)
        try:
            print(f"\nProcessando arquivo {i} de {total_arquivos}: {arquivo}")
            
            df = pd.read_excel(caminho_completo)
            
            if df.empty:
                print(f"  Arquivo vazio: {arquivo}")
                resultados['falha'].append(f"{arquivo} - Arquivo vazio")
                continue
            
            print(f"  Tratando e padronizando os dados...")
            df_tratado = tratar_dados(df)
            
            print(f"  Inserindo dados no PostgreSQL...")
            registros_inseridos = inserir_dados_psycopg2(df_tratado)
            
            resultados['sucesso'].append(f"{arquivo} - {registros_inseridos} registros inseridos")
            print(f"  Concluído: {arquivo} - {registros_inseridos} registros inseridos")
            
        except Exception as e:
            resultados['falha'].append(f"{arquivo} - Erro: {str(e)}")
            print(f"  Erro ao processar {arquivo}: {str(e)}")
    
    return resultados

def gerar_relatorio(resultados):
    print("\n" + "="*50)
    print("RELATÓRIO DE PROCESSAMENTO")
    print("="*50)
    
    print("\nArquivos processados com sucesso:")
    if resultados['sucesso']:
        for item in resultados['sucesso']:
            print(f"✓ {item}")
    else:
        print("Nenhum arquivo processado com sucesso.")
    
    print("\nArquivos com falha:")
    if resultados['falha']:
        for item in resultados['falha']:
            print(f"✗ {item}")
    else:
        print("Nenhum arquivo falhou no processamento.")
    
    total_sucesso = len(resultados['sucesso'])
    total_falha = len(resultados['falha'])
    total = total_sucesso + total_falha
    
    print("\nResumo:")
    print(f"Total de arquivos: {total}")
    if total > 0:
        print(f"Processados com sucesso: {total_sucesso} ({(total_sucesso/total*100):.2f}%)")
        print(f"Falhas: {total_falha} ({(total_falha/total*100):.2f}%)")

if __name__ == "__main__":
    print("")
    print("\n===== TRANSFERÊNCIA DE EXCEL PARA POSTGRESQL =====\n")
    print(" ")
    
    print("Configurações do banco de dados:")
    print(f"Host: {CONFIG_BD['host']}")
    print(f"Porta: {CONFIG_BD['porta']}")
    print(f"Banco de dados: {CONFIG_BD['banco']}")
    print(f"Usuário: {CONFIG_BD['usuario']}")
    print(f"Senha: {'*' * len(CONFIG_BD['senha'])}")
    
    confirmar = input("\nAs configurações acima estão corretas? (s/n): ")
    if confirmar.lower() != 's':
        print("\nPor favor, edite as configurações no início do script e execute novamente.")
        exit(0)
    
    diretorio_excel = input("\nDigite o caminho do diretório com os arquivos Excel: ")
    
    if not os.path.exists(diretorio_excel):
        print(f"Diretório não encontrado: {diretorio_excel}")
        exit(1)
    
    try:
        print("\n--- Criando/verificando banco de dados ---")
        criar_banco_dados()
    except Exception as e:
        print(f"Não foi possível verificar/criar o banco de dados: {str(e)}")
        continuar = input("Deseja tentar continuar mesmo assim? (s/n): ")
        if continuar.lower() != 's':
            exit(1)
    
    try:
        print("\n--- Criando/verificando tabela ---")
        criar_tabela()
    except Exception as e:
        print(f"Erro ao criar tabela: {str(e)}")
        continuar = input("Deseja tentar continuar mesmo assim? (s/n): ")
        if continuar.lower() != 's':
            exit(1)
    
    try:
        print("\n--- Processando planilhas ---")
        resultados = processar_planilhas(diretorio_excel)
        
        gerar_relatorio(resultados)
        
    except Exception as e:
        print(f"Erro ao processar planilhas: {str(e)}")
        exit(1)
    
    print("\nProcesso concluído!")