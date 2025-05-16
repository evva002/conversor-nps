import tkinter 
from conversor import iniciar_processo_conversao

# Interface simples
janela = tkinter.Tk()
janela.title("Conversor de NPS")
janela.geometry("300x150")

tkinter.Button(janela, text="Converter arquivos", command=iniciar_processo_conversao, height=2, width=30).pack(pady=40)

janela.mainloop()