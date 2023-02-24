import pandas as pd

# abra o arquivo Excel
arquivo_excel = pd.read_excel('Estoque -DAT.xlsx')

# Adicione a coluna ASN e preencha com valores em branco
arquivo_excel['ASN'] = arquivo_excel['ASN'].fillna('')

# Seleciona as linhas que atendem aos critérios e têm valores não nulos na coluna ASN
linhas_selecionadas = arquivo_excel[(arquivo_excel['Código do Setor'].isin([25, 26, 27, 29])) & (arquivo_excel['Status da LPN'] == 'In Transit') & (arquivo_excel['ASN'] != '')]

# Converte a coluna "Data de Criação LPN" para a data com horário
linhas_selecionadas['Data de Criação LPN'] = pd.to_datetime(linhas_selecionadas['Data de Criação LPN'])

# Agrupa as linhas selecionadas por dia de criação da LPN e conta a quantidade de ocorrências únicas em cada data
contagem_por_data_LPN = linhas_selecionadas.groupby(linhas_selecionadas['Data de Criação LPN'].dt.date)['ASN'].nunique()
contagem_por_data_ASN = linhas_selecionadas.groupby(linhas_selecionadas['Data de Criação LPN'].dt.date)['ASN'].count()

# Imprime as informações de data, quantidade de LPNs e quantidade de ASN para cada data
for data in contagem_por_data_LPN.index:
    print(f'{data}: {contagem_por_data_LPN[data]} ASN - {contagem_por_data_ASN[data]} LPNs')

# Calcula a quantidade total de LPNs em trânsito
quantidade_total_LPN = contagem_por_data_ASN.sum()
print(f'Quantidade total de LPNs em trânsito: {quantidade_total_LPN}')

# Calcula a quantidade total de ASN em trânsito
quantidade_total_ASN = contagem_por_data_LPN.sum()
print(f'Quantidade total de ASN em trânsito: {quantidade_total_ASN}')

# Crie uma lista de ASN únicas em trânsito
ASN_unicas = linhas_selecionadas['ASN'].unique().tolist()

# Imprime a lista de ASN únicas em trânsito
print(f'LISTAS DAS ASNs em trânsito: {ASN_unicas}')



