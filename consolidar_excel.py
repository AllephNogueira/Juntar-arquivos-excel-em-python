import pandas as pd
import os
from glob import glob
from datetime import datetime

# --------------------------------------------------------
# 🚀 Consolidador de Arquivos Excel com Python/Pandas
# --------------------------------------------------------

# 1. Defina o caminho da pasta
# Mantenha o 'r' antes do caminho para evitar problemas com barras invertidas.
# Substitua pelo seu caminho real:
CAMINHO_PASTA = r"C:\Users\alleph.oliveira\Downloads\Outros"

# 2. Defina o padrão de busca (todos os arquivos .xlsx)
# Exclui arquivos que comecem com "~$" (arquivos temporários abertos do Excel)
PADRAO_ARQUIVOS = os.path.join(CAMINHO_PASTA, "[!~]*.xlsx")

# 3. Lista todos os arquivos que correspondem ao padrão
# glob() é uma função poderosa para listar arquivos
arquivos_excel = glob(PADRAO_ARQUIVOS)

# 4. Define o nome do arquivo de saída
data_hora = datetime.now().strftime("%Y%m%d_%H%M")
nome_saida = f"Dados_Consolidados_{data_hora}.xlsx"
caminho_saida = os.path.join(CAMINHO_PASTA, nome_saida)

# 5. Inicializa uma lista para armazenar os DataFrames de cada arquivo
lista_dataframes = []

print(f"Iniciando a consolidação na pasta: {CAMINHO_PASTA}")
print(f"Encontrados {len(arquivos_excel)} arquivos para processar.")

# 6. Loop para ler cada arquivo e adicionar à lista
for arquivo in arquivos_excel:
    nome_arquivo = os.path.basename(arquivo)
    
    # Ignora o arquivo de saída, caso ele já exista
    if nome_arquivo == nome_saida:
        continue

    print(f"-> Lendo arquivo: {nome_arquivo}")

    try:
        # Lê o conteúdo do Excel. 
        # O parâmetro sheet_name=0 lê a primeira planilha.
        df = pd.read_excel(arquivo, sheet_name=0)
        
        # Opcional: Adiciona uma coluna com o nome da fonte original
        df['Arquivo_Fonte'] = nome_arquivo 
        
        lista_dataframes.append(df)
        
    except Exception as e:
        print(f"--- ERRO ao processar o arquivo {nome_arquivo}: {e}")

# 7. Concatena todos os DataFrames em um único DataFrame
if lista_dataframes:
    try:
        df_consolidado = pd.concat(lista_dataframes, ignore_index=True)
        
        print("\n✅ Consolidação de DataFrames concluída. Exportando...")
        
        # 8. Exporta o DataFrame consolidado para um novo arquivo Excel
        # engine='openpyxl' garante compatibilidade com .xlsx
        df_consolidado.to_excel(
            caminho_saida, 
            sheet_name='Consolidado', 
            index=False, # Não inclui o índice do DataFrame como coluna
            engine='openpyxl'
        )
        
        print(f"\n✅ SUCESSO! Dados exportados para:")
        print(f"   {caminho_saida}")
        print(f"   Total de linhas consolidadas: {len(df_consolidado)}")

    except Exception as e:
        print(f"\n--- ERRO ao concatenar ou exportar: {e}")

else:
    print("\n⚠️ Nenhum dado válido foi encontrado para consolidação.")