import tkinter as tk
from tkinter import messagebox
import openpyxl

def adicionar():
  nome = nome_entry.get()
  idade = idade_entry.get()
  pais = pais_entry.get()

  wb = openpyxl.load_workbook('dados.xlsx')
  sheet = wb.active

  last_row = sheet.max_row + 1

  sheet[f'A{last_row}'] = nome
  sheet[f'B{last_row}'] = idade
  sheet[f'C{last_row}'] = pais

  wb.save('dados.xlsx')

  nome_entry.delete(0, 'end')
  idade_entry.delete(0, 'end')
  pais_entry.delete(0, 'end')

root = tk.Tk()
root.title('Inserção de Dados')

nome_label = tk.Label(root, text='Nome:')
idade_label = tk.Label(root, text='Idade:')
pais_label = tk.Label(root, text='País:')
nome_entry = tk.Entry(root)
idade_entry = tk.Entry(root)
pais_entry = tk.Entry(root)
nome_label.pack()
nome_entry.pack()
idade_label.pack()
idade_entry.pack()
pais_label.pack()
pais_entry.pack()

adicionar_button = tk.Button(root, text='Adicionar', command=adicionar)
adicionar_button.pack()

root.mainloop()
