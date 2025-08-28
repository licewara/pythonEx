import pandas as pd
import tkinter as tk
from tkinter import filedialog

def organizarPlanilha():
  caminho = filedialog.askopenfilename(
    title="Selecione a planilha",
  filetypes=[("Arquivos Excel", "*.xlsx *.xls")]
  )

  if caminho:
    df = pd.read_excel(caminho)
    df = df.drop_duplicates() 
    df.to_excel("planilha_organizada.xlsx", index=False)
    
    texto.delete("1.0", tk.END)
    texto.insert(tk.END, df.head().to_string())


janela = tk.Tk()
janela.title("Organizador de Planilhas")

botao = tk.Button(janela, text="Organizar Planilha", command=organizarPlanilha)
botao.pack()

texto = tk.Text(janela, height=1920, width=1080)
texto.pack()

janela.mainloop()

