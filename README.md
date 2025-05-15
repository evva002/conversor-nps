
CONVERSOR DE PLANILHAS NPS (.XLSX → .CSV)

Este programa converte automaticamente arquivos Excel contendo pesquisas de NPS para arquivos .csv

---------------------------------------
1. O QUE O PROGRAMA FAZ
---------------------------------------

- Lê todos os arquivos .xlsx ou .xls da pasta "pre-processamento"
- Converte colunas de telefone e CPF/CNPJ para formato numérico
- Renomeia os arquivos .csv de saída com base no nome original
- Salva os arquivos convertidos na pasta "pos-processamento"

---------------------------------------
2. ESTRUTURA DE PASTAS NECESSÁRIA
---------------------------------------

O programa espera que estas pastas existam no mesmo local do EXE (ou .py):

📁 sua_pasta_do_programa
├── conversor.exe  (ou main.py)
├── pre-processamento
└── pos-processamento

* Se as pastas não existirem, o programa irá criá-las automaticamente.

---------------------------------------
3. COLUNAS ALTERADAS
---------------------------------------

O programa altera as seguintes colunas (se existirem no arquivo):

- COD_CPF_CGC
- FONE_1
- FONE_2
- FONE_3
- FONE_4

Essas colunas são convertidas para número inteiro (Int64), ignorando valores inválidos.

---------------------------------------
4. RENOMEAÇÃO AUTOMÁTICA DOS ARQUIVOS
---------------------------------------

Baseado no nome do arquivo .xlsx, os nomes dos .csv gerados são:

Se nome conter:           → Será salvo como:
-------------------------    --------------------------
PESQUISA_NPS              → PESQUISA_NPS.csv
NPS20                     → NPS20.csv
NPS19                     → NPS20.csv
NPS18                     → NPS18.csv
NPS17                     → NPS17.csv
NPS16                     → NPS16.csv
NPS15                     → NPS15.csv
NPS14                     → NPS14.csv
NPS13                     → NPS13.csv
NPS12                     → NPS12.csv
NPS11                     → NPS11.csv
NPS10                     → NPS10.csv
NPS9                      → NPS9.csv
NPS8                      → NPS8.csv
NPS7                      → NPS7.csv
NPS6                      → NPS6.csv
NPS5                      → NPS5.csv
NPS4                      → NPS4.csv
NPS3                      → NPS3.csv
NPS                       → NPS2.csv  ← (só "NPS" por último)

Se o nome não combinar com nenhum dos acima, será salvo como:
nome_original_do_arquivo.csv

---------------------------------------
5. COMO USAR
---------------------------------------

1. Coloque seus arquivos .xlsx em "pre-processamento"
2. Execute o programa (conversor.exe)
3. Os .csv processados serão salvos na pasta "pos-processamento"

---------------------------------------
6. REQUISITOS (APENAS SE FOR USAR COM PYTHON)
---------------------------------------

- Python 3.x instalado
- Bibliotecas:
  - pandas
  - openpyxl

Instale com:
pip install pandas openpyxl

---------------------------------------
7. COMPILAR PARA EXE (OPCIONAL)
---------------------------------------

Se quiser transformar em um programa executável:

Use PyInstaller:
pyinstaller --onefile --noconsole main.py

O .exe gerado estará na pasta "dist".

---------------------------------------
DÚVIDAS OU ERROS?
---------------------------------------
- Para somente utilizar o programa e o conversor.exe, copiar a pasta conversor_nps.zip e colar no local desejado
(de preferencia no disco C)!

Se o programa disser que não encontrou arquivos ou pastas:
- Verifique se as pastas estão no mesmo local que o .exe
- Verifique se os arquivos têm extensão .xlsx ou .xls

Todos os arquivos .csv serão sobrescritos se já existirem.
