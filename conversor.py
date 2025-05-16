import pandas
import os
import sys
from tkinter import messagebox

colunas_alvo = ['COD_CPF_CGC', 'FONE_1', 'FONE_2', 'FONE_3', 'FONE_4']

def iniciar_processo_conversao():

    # procura o caminho do arquivo caso seja .exe
    if getattr(sys, 'frozen', False):
        pasta_base = os.path.dirname(sys.executable)

    # procura o caminho do arquivo caso seja .py
    else:
        pasta_base = os.path.dirname(os.path.abspath(__file__))

    # define a localizacao e nome das pastas necessarias para rodar o programa
    pasta_entrada = os.path.join(pasta_base, 'pasta-entrada')
    pasta_saida = os.path.join(pasta_base, 'pasta-saida')

    # cria as pastas caso elas nao existam
    os.makedirs(pasta_entrada, exist_ok=True)
    os.makedirs(pasta_saida, exist_ok=True)

    # erro caso a pasta nao exista
    if not os.path.exists(pasta_entrada) or not os.path.exists(pasta_saida):
        messagebox.showerror("Erro", "As pastas 'pasta-entrada' e 'pasta-saida' precisam existir no mesmo diretório do script.")
        return
    
    # percorre a pasta_entrada e joga todos os arquivos .xlsx ou .xls dentro da lista arquivos
    arquivos = [arquivo for arquivo in os.listdir(pasta_entrada) if arquivo.lower().endswith(('.xlsx', '.xls'))]
    if not arquivos:
        messagebox.showinfo("Aviso", "Nenhum arquivo Excel encontrado na pasta 'pasta-entrada'.")
        return

    # Cria uma lista, sendo o primeiro campo o valor que ele vai procurar no nome do arquivo .xlsx 
    # o segundo campo é o nome do novo arquivo .csv
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
        ('NPS2', 'NPS2.csv'),
    ]

    # le arquivo por arquivo na lista arquivos
    for arquivo in arquivos:
        try:
            # acha o caminho do arquivo .xlsx e inicia a leitura do arquivo
            caminho_entrada = os.path.join(pasta_entrada, arquivo)
            planilha = pandas.read_excel(caminho_entrada)

            # altera as colunas escolhidas para o valor numerico
            for coluna in colunas_alvo:
                planilha[coluna] = pandas.to_numeric(planilha[coluna], errors='coerce') # a parte de errors='coerce' trata alguma linha que possui um valor que nao pode ser convertido para numero e substitui por NaN
                                                                                        # pode ser subistituido por um 'ignore' porem ira aparecer posteriormente 

            # percorre cada coluna do arquivo
            for coluna in colunas_alvo:
                planilha[coluna] = planilha[coluna].dropna().astype(int) # remove os NaN e converte a coluna para int
                planilha[coluna] = planilha[coluna].astype('Int64') # converte a coluna para Int64, pertence a biblioteca pandas e permite valores NaN

            # divide o nome do arquivo e pega somente o nome[0] sem a extensao .xlsx
            nome_arquivo_base = os.path.splitext(arquivo)[0]
            nome_arquivo_upper = nome_arquivo_base.upper()

            # pega a lista mapeamento e procura o nome_arquivo_upper que foi feito acima para renomear o arquivo .csv
            nome_saida = None
            for chave, nome_csv in mapeamento:
                if chave in nome_arquivo_upper:
                    nome_saida = nome_csv
                    break
      
            if not nome_saida:
                nome_saida = nome_arquivo_base.lower() + '.csv'

            #salva o arquivo na pasta-saida
            caminho_saida = os.path.join(pasta_saida, nome_saida)
            planilha.to_csv(caminho_saida, index=False)

        except Exception as erro:
            messagebox.showerror("Erro", f"Erro ao processar {arquivo}:\n{str(erro)}")

    messagebox.showinfo("Concluído", "Todos os arquivos foram processados com sucesso!")