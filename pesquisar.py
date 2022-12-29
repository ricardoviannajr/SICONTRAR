import tkinter as tk
from tkinter import messagebox
import openpyxl

def pesquisar():
  # Obtém o nome a ser pesquisado
  nome = nome_entry.get()

  # Abre a planilha de dados
  wb = openpyxl.load_workbook('dados.xlsx')
  sheet = wb.active

  # Pesquisa por um determinado valor na planilha
  for row in sheet.iter_rows():
    if row[0].value == nome:
      # Exibe os resultados da pesquisa em um pop-up
      messagebox.showinfo('Resultados da Pesquisa', f'Nome: {row[0].value}\nIdade: {row[1].value}\nPaís: {row[2].value}')
      return

  # Exibe uma mensagem de erro se o resultado da pesquisa não foi encontrado
  messagebox.showerror('Erro', 'Nome não encontrado')

# Cria a janela principal da aplicação
root = tk.Tk()
root.title('Pesquisa de Dados')

# Cria um rótulo e uma caixa de entrada para o nome
nome_label = tk.Label(root, text='Nome:')
nome_entry = tk.Entry(root)
nome_label.pack()
nome_entry.pack()

# Cria um botão de pesquisa
pesquisar_button = tk.Button(root, text='Pesquisar', command=pesquisar)
pesquisar_button.pack()

