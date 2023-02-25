import pandas as pd

# mapeia cada código do setor ao seu respectivo código de linha
setor_linha = {
    5: 'ELE', 9: 'ELE', 11: 'ELE', 13: 'ELE', 17: 'ELE', 18: 'ELE', 19: 'ELE', 21: 'ELE', 22: 'ELE', 28: 'ELE',
    42: 'ELE', 44: 'ELE', 45: 'ELE', 46: 'ELE', 54: 'ELE', 56: 'ELE', 59: 'ELE',
    33: 'EST', 34: 'EST', 25: 'LBR', 26: 'LBR', 27: 'LBR', 29: 'LBR',
    31: 'MOV', 32: 'MOV', 36: 'MOV', 38: 'MOV', 39: 'MOV', 63: 'MOV', 64: 'MOV',
    1: 'PAR', 4: 'PAR', 6: 'PAR', 10: 'PAR', 14: 'PAR', 15: 'PAR', 20: 'PAR',
    40: 'PAR', 43: 'PAR', 49: 'PAR', 50: 'PAR', 51: 'PAR', 52: 'PAR', 58: 'PAR', 60: 'PAR', 67: 'PAR', 69: 'PAR'
}

# carrega o arquivo Excel em um DataFrame
df = pd.read_excel('Estoque -DAT.xlsx')

# filtra as linhas que possuem o status 'in transit'
in_transit = df[df['Status da LPN'] == 'In Transit']

# adiciona uma coluna com as siglas correspondentes ao código do setor
in_transit.loc[:, 'LINHA'] = in_transit['Código do Setor'].map(setor_linha)

# converte a coluna 'Data de Criação LPN' para o tipo datetime e extrai a data
in_transit.loc[:, 'Data de Criação LPN'] = pd.to_datetime(in_transit['Data de Criação LPN']).dt.date

# agrupa as linhas por data, linha e status, e conta o número de linhas em cada grupo
grouped = in_transit.groupby(['Data de Criação LPN', 'LINHA', 'Status da LPN'], as_index=False).size()

# filtra apenas as linhas com status 'In Transit'
grouped = grouped[grouped['Status da LPN'] == 'In Transit']

# agrupa as linhas novamente por data e linha, e soma o total de in transit para cada grupo
grouped = grouped.groupby(['LINHA', 'Data de Criação LPN'], as_index=False).sum()

# imprime o resultado separado por setor, ordenado por linha e data
setores = sorted(set(grouped['LINHA']))
for linha in setores:
    print(f"\n Pendência {linha}:")
    df = grouped[grouped['LINHA'] == linha].rename(columns={"Data de Criação LPN": "DATA", "size": "Quantidade"})
    total = df['Quantidade'].sum()
    print(df[['DATA', 'Quantidade']].sort_values(['DATA']).to_string(index=False, justify='center'))
    print(f"Falta Receber: {total}")
    print('-'* 50)



# $$$$$$$$$$$$$$$$$$