import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import os

def organizarPlanilha():
  caminho = filedialog.askopenfilename(
    title="Selecione a planilha",
  filetypes=[("Arquivos Excel", "*.xlsx *.xls")]
  )

  if not caminho:
    return
  
  try:
    if caminho.lower().endswith((".xlsx", ".xls")):
      df = pd.read_excel(caminho)

      df = df[["Produto", "Valor Unitário"]]

      pasta = os.path.dirname(caminho)
      saida = os.path.join(pasta, "planilha_organizada.xlsx")
      df.to_excel(saida, index=False)

      texto.delete("1.0", tk.END)
      texto.insert(tk.END, df.head(20).to_string(index=False))

      messagebox.showinfo("Pronto!", f"Planilha salva em:\n{saida}")
    else:
      df = pd.read_csv(caminho)

      df = df[["Produto", "Valor Unitário"]]

      pasta = os.path.dirname(caminho)
      saida = os.path.join(pasta, "planilha_organizada.xlsx")
      df.to_excel(saida, index=False)

      texto.delete("1.0", tk.END)
      texto.insert(tk.END, df.head(20).to_string(index=False))

      messagebox.showinfo("Pronto!", f"Planilha salva em:\n{saida}")

  except KeyError as e:
    messagebox.showerror("Erro",
      f"Coluna não encontrada: {e}\n\n"
      "Confira se os nomes das colunas estão exatamente como 'Produto' e 'Valor Unitário'.")
  
  except Exception as e:
    messagebox.showerror("Erro inesperado", str(e))

if __name__ == "__main__":
  janela = tk.Tk()
  janela.title("Organizador de Planilhas")

  botao = tk.Button(janela, text="Organizar Planilha", command=organizarPlanilha)
  botao.pack(pady=8)

  frame = tk.Frame(janela)
  frame.pack(fill="both", expand=True)

  texto = tk.Text(frame, height=20, width=90, wrap="none")
  texto.pack(side="left", fill="both", expand=True)

  scroll_y = tk.Scrollbar(frame, orient="vertical", command=texto.yview)
  scroll_y.pack(side="right", fill="y")

  texto.configure(yscrollcommand=scroll_y.set)

  janela.mainloop()

janela = tk.Tk()
janela.title("Organizador de Planilhas")

botao = tk.Button(janela, text="Organizar Planilha", command=organizarPlanilha)
botao.pack()

texto = tk.Text(janela, height=1920, width=1080)
texto.pack()

janela.mainloop()

