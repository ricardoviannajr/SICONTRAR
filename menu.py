import tkinter as tk

# Importa as funções das aplicações de inserção e pesquisa de dados
from adicionar import adicionar
from pesquisar import pesquisar

root = tk.Tk()
root.title('Gerenciamento de Dados')

menu = tk.Menu(root)
root.config(menu=menu)

inserir_menu = tk.Menu(menu)
pesquisar_menu = tk.Menu(menu)
menu.add_cascade(label='Inserir', menu=inserir_menu)
menu.add_cascade(label='Pesquisar', menu=pesquisar_menu)
inserir_menu.add_command(label='Inserir Dados', command=adicionar)
pesquisar_menu.add_command(label='Pesquisar Dados', command=pesquisar)

# Remove as chamadas das aplicações de inserção de dados e pesquisa de dados
#adicionar()
#pesquisar()

root.mainloop()
