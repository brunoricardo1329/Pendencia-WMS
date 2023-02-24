import pandas as pd

# abra o arquivo Excel
arquivo_excel = pd.read_excel(r'C:\Users\2904894200\Downloads\Estoque -DAT.xlsx')

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
print('Pendência DAT Linha Branca')
print('-' * 50)
for data in contagem_por_data_LPN.index:
    print(f'{data.strftime("%Y-%m-%d"): <12} {contagem_por_data_LPN[data]: >3} ASN - {contagem_por_data_ASN[data]: >3} LPNs')
print('-' * 50)

# Calcula a quantidade total de LPNs em trânsito
quantidade_total_LPN = contagem_por_data_ASN.sum()
print(f'Quantidade total de LPNs em trânsito: {quantidade_total_LPN}')
print('-' * 50)

# Calcula a quantidade total de ASN em trânsito
quantidade_total_ASN = contagem_por_data_LPN.sum()
print(f'Quantidade total de ASN em trânsito: {quantidade_total_ASN}')
print('-' * 50)

# Crie uma lista de ASN únicas em trânsito
ASN_unicas = linhas_selecionadas['ASN'].unique().tolist()

# Crie um dicionário de ASN por data
ASN_por_data = {}

# Itera pelas linhas selecionadas e armazena as ASN em um dicionário
for index, row in linhas_selecionadas.iterrows():
    data = row['Data de Criação LPN'].strftime("%Y-%m-%d")
    ASN = row['ASN']
    if data not in ASN_por_data:
        ASN_por_data[data] = [ASN]
    else:
        ASN_por_data[data].append(ASN)

# Imprime a lista de ASN únicas em trânsito para cada data, ordenada da mais antiga para a mais nova
datas_ordenadas = sorted(ASN_por_data.keys())
for data in datas_ordenadas:
    ASN_lista = ASN_por_data[data]
    ASN_unicas = list(set(ASN_lista))
    print(f'ASN em trânsito {data}: {", ".join(ASN_unicas)}')
