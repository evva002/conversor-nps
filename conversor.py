import pandas
import os
import sys
from tkinter import messagebox

colunas_alvo = ['COD_CPF_CGC', 'FONE_1', 'FONE_2', 'FONE_3', 'FONE_4']

def iniciar_processo_conversao():

    if getattr(sys, 'frozen', False):
        pasta_base = os.path.dirname(sys.executable)

    else:
        pasta_base = os.path.dirname(os.path.abspath(__file__))

    pasta_entrada = os.path.join(pasta_base, 'pre-processamento')
    pasta_saida = os.path.join(pasta_base, 'pos-processamento')

    os.makedirs(pasta_entrada, exist_ok=True)
    os.makedirs(pasta_saida, exist_ok=True)

    if not os.path.exists(pasta_entrada) or not os.path.exists(pasta_saida):
        messagebox.showerror("Erro", "As pastas 'pre-processamento' e 'pos-processamento' precisam existir no mesmo diretório do script.")
        return

    arquivos = [f for f in os.listdir(pasta_entrada) if f.lower().endswith(('.xlsx', '.xls'))]
    if not arquivos:
        messagebox.showinfo("Aviso", "Nenhum arquivo Excel encontrado na pasta 'pre-processamento'.")
        return

    mapeamento = [
        ('PESQUISA_NPS', 'PESQUISA_NPS.csv'),
        ('NPS20', 'NPS20.csv'),
        ('NPS19', 'NPS20.csv'),
        ('NPS18', 'NPS18.csv'),
        ('NPS17', 'NPS17.csv'),
        ('NPS16', 'NPS16.csv'),
        ('NPS15', 'NPS15.csv'),
        ('NPS14', 'NPS14.csv'),
        ('NPS13', 'NPS13.csv'),
        ('NPS12', 'NPS12.csv'),
        ('NPS11', 'NPS11.csv'),
        ('NPS10', 'NPS10.csv'),
        ('NPS9', 'NPS9.csv'),
        ('NPS8', 'NPS8.csv'),
        ('NPS7', 'NPS7.csv'),
        ('NPS6', 'NPS6.csv'),
        ('NPS5', 'NPS5.csv'),
        ('NPS4', 'NPS4.csv'),
        ('NPS3', 'NPS3.csv'),
        ('NPS', 'NPS2.csv'),
    ]

    for arquivo in arquivos:
        try:
            caminho_entrada = os.path.join(pasta_entrada, arquivo)
            planilha = pandas.read_excel(caminho_entrada)

            for coluna in colunas_alvo:
                planilha[coluna] = pandas.to_numeric(planilha[coluna], errors='coerce')

            for coluna in colunas_alvo:
                planilha[coluna] = planilha[coluna].dropna().astype(int)
                planilha[coluna] = planilha[coluna].astype('Int64')

            nome_arquivo_base = os.path.splitext(arquivo)[0]
            nome_arquivo_upper = nome_arquivo_base.upper()

            nome_saida = None
            for chave, nome_csv in mapeamento:
                if chave in nome_arquivo_upper:
                    nome_saida = nome_csv
                    break

            if not nome_saida:
                nome_saida = nome_arquivo_base.lower() + '.csv'

            caminho_saida = os.path.join(pasta_saida, nome_saida)
            planilha.to_csv(caminho_saida, index=False)

        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao processar {arquivo}:\n{str(e)}")

    messagebox.showinfo("Concluído", "Todos os arquivos foram processados com sucesso!")