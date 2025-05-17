import pandas as pd
from glob import glob
import os

# Pasta de arquivos para leitura
pasta_arquivos = input('Digite ou cole o caminho da pasta com arquivos:')


# Nome do arquivo que deseja salvar
nome_arquivo_novo = input('Digite nome do arquivo que deseja salvar:')

# Pasta para salvar arquivo de Excel
local_salvar = input('Digite ou cole local pasta para salvar:')

# ------------------------------------------------------------------------------------#

# Lendo os arquivos e salvando no lugar desejado

# Processar os arquivos
if pasta_arquivos and nome_arquivo_novo and local_salvar:
    arquivos = sorted(glob(os.path.join(pasta_arquivos, '*.xlsx')))
    if arquivos:
        # Concatenando todos os arquivos
        todos_arquivos = pd.concat((pd.read_excel(arquivo) for arquivo in arquivos), 
                                       ignore_index=True)
        # Caminho final do arquivo salvo
        caminho_final = os.path.join(local_salvar, f'{nome_arquivo_novo}.xlsx')

        # Salvando o arquivo
        todos_arquivos.to_excel(caminho_final, index=False)
        
        print(f'Arquivo salvo com sucesso em: {caminho_final}')
    else:
        print('Nenhum arquivo .xlsx encontrado na pasta especificada.')
else:
    print('Por favor, preencha todos os campos antes de salvar.')
        
# ------------------------------------------------------------------------------------#