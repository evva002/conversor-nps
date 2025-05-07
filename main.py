import pandas 

arquivo_entrada = 'dados.xlsx'

colunas_alvo = ['COD_CPF_CGC', 'FONE_1', 'FONE_2', 'FONE_3', 'FONE_4']

planilha = pandas.read_excel(arquivo_entrada)

for coluna in colunas_alvo:
    planilha[coluna] = pandas.to_numeric(planilha[coluna], errors='coerce')

for coluna in colunas_alvo:
    if planilha[coluna].notna().sum() != planilha[coluna].count():
        print(f'Atenção: A coluna {coluna} contém valores inválidos (ex: texto que não é número).')

for col in colunas_alvo:
    planilha[col] = planilha[col].dropna().astype(int)
    planilha[col] = planilha[col].astype('Int64')

planilha.to_csv('dados_para_discador.csv', index=False)

print('Processo concluído!')