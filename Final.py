import tkinter as tk
import tkinter.messagebox
import tkinter.filedialog
import csv
import os
import webbrowser
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from unidecode import unidecode
from PIL import Image, ImageTk

# Verifica se o arquivo com o caminho do arquivo "dados.csv" existe
if not os.path.exists('caminho_arquivo.txt'):
    # Pergunta ao usuário de onde importar o arquivo
    caminho_arquivo = tkinter.filedialog.askopenfilename()
    # Salva o caminho do arquivo em um arquivo de texto
    with open('caminho_arquivo.txt', 'w', encoding = 'unicode_escape') as arquivo:
        arquivo.write(caminho_arquivo)
else:
    # Lê o caminho do arquivo do arquivo de texto
    with open('caminho_arquivo.txt', 'r', encoding = 'unicode_escape') as arquivo:
        caminho_arquivo = arquivo.read()


def adicionar():
    data_transferencia = data_transferencia_entry.get()
    numero_documento = numero_documento_entry.get()
    sigla_unidade = sigla_unidade_entry.get()
    nome_unidade = nome_unidade_entry.get()
    numero_cx_escritorio = numero_cx_escritorio_entry.get()
    numero_cx_custodia = numero_cx_custodia_entry.get()
    codigo_classificacao = codigo_classificacao_entry.get()
    datas_limite = datas_limite_entry.get()
    descricao_documentos = descricao_documentos_entry.get()
    prazo_guarda = prazo_guarda_entry.get()
    destinacao = destinacao_entry.get()
    localizacao_unidade = localizacao_unidade_entry.get()
    localizacao_conjunto = localizacao_conjunto_entry.get()
    localizacao_rua = localizacao_rua_entry.get()
    localizacao_estante = localizacao_estante_entry.get()
    localizacao_prateleira = localizacao_prateleira_entry.get()

    # Verifica se todos os campos foram preenchidos
    if not all([data_transferencia, numero_documento, sigla_unidade, nome_unidade,
                numero_cx_escritorio, numero_cx_custodia, codigo_classificacao,
                datas_limite, descricao_documentos, prazo_guarda, destinacao,
                localizacao_unidade, localizacao_conjunto, localizacao_rua,
                localizacao_estante, localizacao_prateleira]):
        tk.messagebox.showerror('Erro', 'Todos os campos são obrigatórios.')
        return

    # Adiciona os dados no arquivo CSV
    with open(caminho_arquivo, 'a', encoding = 'unicode_escape', newline='') as arquivo_csv:
        writer = csv.writer(arquivo_csv)
        writer.writerow([data_transferencia, numero_documento, sigla_unidade, nome_unidade,
                         numero_cx_escritorio, numero_cx_custodia, codigo_classificacao,
                         datas_limite, descricao_documentos, prazo_guarda, destinacao,
                         localizacao_unidade, localizacao_conjunto, localizacao_rua,
                         localizacao_estante, localizacao_prateleira])

    # Limpa os campos de entrada
    data_transferencia_entry.delete(0, 'end')
    numero_documento_entry.delete(0, 'end')
    sigla_unidade_entry.delete(0, 'end')
    nome_unidade_entry.delete(0, 'end')
    numero_cx_escritorio_entry.delete(0, 'end')
    numero_cx_custodia_entry.delete(0, 'end')
    codigo_classificacao_entry.delete(0, 'end')
    datas_limite_entry.delete(0, 'end')
    descricao_documentos_entry.delete(0, 'end')
    prazo_guarda_entry.delete(0, 'end')
    destinacao_entry.delete(0, 'end')
    localizacao_unidade_entry.delete(0, 'end')
    localizacao_conjunto_entry.delete(0, 'end')
    localizacao_rua_entry.delete(0, 'end')
    localizacao_estante_entry.delete(0, 'end')
    localizacao_prateleira_entry.delete(0, 'end')


def pesquisar():
    termo = pesquisa_entry.get()
    resultado = ''

    if not termo:
        tk.messagebox.showerror('Erro', 'O campo de pesquisa é obrigatório.')
        return

    with open(caminho_arquivo, 'r', encoding = 'unicode_escape') as arquivo_csv:
        reader = csv.reader(arquivo_csv)
        for linha in reader:
            if len(linha) >= 16:
                dados = f"{linha[0]} {linha[1]} {linha[2]} {linha[3]} {linha[4]} {linha[5]} {linha[6]} {linha[7]} " \
                    f"{linha[8]} {linha[9]} {linha[10]} {linha[11]} {linha[12]} {linha[13]} {linha[14]} {linha[15]}"
                if termo.lower() in unidecode(dados).lower():
                    resultado += f"Data de transferência: {linha[0]}, " \
                             f"Número do documento de encaminhamento: {linha[1]}, " \
                             f"Unidade produtora (sigla): {linha[2]}, " \
                             f"Unidade produtora (nome): {linha[3]}, " \
                             f"Nº CX escritório: {linha[4]}, " \
                             f"Nº caixa custódia: {linha[5]}, " \
                             f"Código classificação documental: {linha[6]}, " \
                             f"Datas-limite: {linha[7]}, " \
                             f"Descrição dos documentos: {linha[8]}, " \
                             f"Prazo de guarda - arquivo intermediário: {linha[9]}, " \
                             f"Destinação: {linha[10]}, " \
                             f"Localização unidade de arquivo: {linha[11]}, " \
                             f"Localização conjunto: {linha[12]}, " \
                             f"Localização rua: {linha[13]}, " \
                             f"Localização estante: {linha[14]}, " \
                             f"Localização prateleira: {linha[15]}\n"

    # Limpa o campo de entrada
    pesquisa_entry.delete(0, 'end')

    if resultado:
        tk.messagebox.showinfo('Resultado da Pesquisa', resultado)
    else:
        tk.messagebox.showerror('Erro', 'Nenhum registro encontrado.')


def gerar_relatorio():
    # Lê o caminho do arquivo do arquivo de texto
    with open('caminho_arquivo.txt', 'r', encoding = 'unicode_escape') as arquivo:
        caminho_arquivo_csv = arquivo.read()

    # Inicializa uma lista para armazenar os dados
    dados = []

    # Lê todas as linhas da base de dados e adiciona na lista de dados
    with open(caminho_arquivo_csv, 'r', encoding = 'unicode_escape') as arquivo_csv:
        reader = csv.reader(arquivo_csv)
        for i, linha in enumerate(reader):
            if i == 0:
                # Adiciona os títulos das colunas na primeira linha
                dados.append(['Data de transferencia ao arquivo de custodia',
                              'Número do documento de encaminhamento',
                              'Unidade Produtora Sigla',
                              'Unidade Produtora Nome',
                              'Nº CX escritorio',
                              'Nº Caixa Custódia',
                              'Codigo classificaçao documental',
                              'Datas-limite',
                              'Descriçao dos documentos',
                              'Prazo de Guarda - Arquivo Intermediário',
                              'Destinaçao',
                              'Localizaçao Unidade de arquivo',
                              'Localizaçao Conjunto',
                              'Localizaçao Rua',
                              'Localizaçao Estante',
                              'Localizaçao Prateleira'
                              ])
            else:
                dados.append(linha)

    # Inicializa uma lista para armazenar as tuplas de estilo
    estilo = []

    # Adiciona a tupla para alterar a cor da primeira linha para lightgrey
    estilo.append(('BACKGROUND', (0, 0), (-1, 0), colors.cyan))

    # Gera as tuplas de estilo de acordo com o número de linhas da tabela
    for i in range(1, len(dados)):
        if i % 2 == 0:
            # Adiciona a tupla para alterar a cor da linha para lightgrey
            estilo.append(('BACKGROUND', (0, i), (-1, i), colors.lightgrey))
        else:
            # Adiciona a tupla para alterar a cor da linha para white
            estilo.append(('BACKGROUND', (0, i), (-1, i), colors.white))

    # Cria a tabela com os dados e o estilo
    tabela = Table(dados, style=estilo)

    # Adiciona as tuplas para alterar a largura das colunas
    tabela.setStyle(TableStyle([
        ('COLWIDTH', (0, 0), (0, -1), 200),
        ('COLWIDTH', (1, 0), (1, -1), 100)
    ]))

    # Cria o documento PDF
    caminho_arquivo_pdf = '.\\relatorio.pdf'
    doc = SimpleDocTemplate(caminho_arquivo_pdf, pagesize=A4)

    # Adiciona o conteúdo do relatório
    doc.build([tabela])

    # Abre o arquivo PDF
    webbrowser.open(caminho_arquivo_pdf)


def sair():
    root.destroy()


root = tk.Tk()
root.title('Sistema de Controle de Arquivo')

# Logo
logo = Image.open('infralogo.png')
lb = ImageTk.PhotoImage(logo)
label1 = tkinter.Label(image=lb)
label1.image = lb
label1.grid(row=1, column=0, padx=10, pady=10, sticky='nswe')

# Widgets da aplicação
label_titulo = tk.Label(text="SICONTRAR", font=("Helvetica", 14))
label_titulo1 = tk.Label(text="")

# Novos widgets
data_transferencia_label = tk.Label(root, text='Data de transferência:')
data_transferencia_entry = tk.Entry(root)
numero_documento_label = tk.Label(root, text='Número do documento de encaminhamento:')
numero_documento_entry = tk.Entry(root)
sigla_unidade_label = tk.Label(root, text='Unidade produtora (sigla):')
sigla_unidade_entry = tk.Entry(root)
nome_unidade_label = tk.Label(root, text='Unidade produtora (nome):')
nome_unidade_entry = tk.Entry(root)
numero_cx_escritorio_label = tk.Label(root, text='Nº CX escritório:')
numero_cx_escritorio_entry = tk.Entry(root)
numero_cx_custodia_label = tk.Label(root, text='Nº caixa custódia:')
numero_cx_custodia_entry = tk.Entry(root)
codigo_classificacao_label = tk.Label(root, text='Código classificação documental:')
codigo_classificacao_entry = tk.Entry(root)
datas_limite_label = tk.Label(root, text='Datas-limite:')
datas_limite_entry = tk.Entry(root)
descricao_documentos_label = tk.Label(root, text='Descrição dos documentos:')
descricao_documentos_entry = tk.Entry(root)
prazo_guarda_label = tk.Label(root, text='Prazo de guarda - arquivo intermediário:')
prazo_guarda_entry = tk.Entry(root)
destinacao_label = tk.Label(root, text='Destinação:')
destinacao_entry = tk.Entry(root)
localizacao_unidade_label = tk.Label(root, text='Localização unidade de arquivo:')
localizacao_unidade_entry = tk.Entry(root)
localizacao_conjunto_label = tk.Label(root, text='Localização conjunto:')
localizacao_conjunto_entry = tk.Entry(root)
localizacao_rua_label = tk.Label(root, text='Localização rua:')
localizacao_rua_entry = tk.Entry(root)
localizacao_estante_label = tk.Label(root, text='Localização estante:')
localizacao_estante_entry = tk.Entry(root)
localizacao_prateleira_label = tk.Label(root, text='Localização prateleira:')
localizacao_prateleira_entry = tk.Entry(root)

# Botões
inserir_button = tk.Button(root, text='Inserir', command=adicionar)
pesquisa_label = tk.Label(root, text='Pesquisa:')
pesquisa_entry = tk.Entry(root)
pesquisar_button = tk.Button(root, text='Pesquisar', command=pesquisar)
relatorio_button = tk.Button(root, text='Relatório', command=gerar_relatorio)
sair_button = tk.Button(root, text='Sair', command=sair)

# Posicionamento dos novos widgets
data_transferencia_label.grid(row=2, column=0, padx=10, pady=10, sticky='nswe')
data_transferencia_entry.grid(row=2, column=1, padx=10, pady=10, sticky='nswe')
numero_documento_label.grid(row=3, column=0, padx=10, pady=10, sticky='nswe')
numero_documento_entry.grid(row=3, column=1, padx=10, pady=10, sticky='nswe')
sigla_unidade_label.grid(row=4, column=0, padx=10, pady=10, sticky='nswe')
sigla_unidade_entry.grid(row=4, column=1, padx=10, pady=10, sticky='nswe')
nome_unidade_label.grid(row=5, column=0, padx=10, pady=10, sticky='nswe')
nome_unidade_entry.grid(row=5, column=1, padx=10, pady=10, sticky='nswe')
numero_cx_escritorio_label.grid(row=6, column=0, padx=10, pady=10, sticky='nswe')
numero_cx_escritorio_entry.grid(row=6, column=1, padx=10, pady=10, sticky='nswe')
numero_cx_custodia_label.grid(row=7, column=0, padx=10, pady=10, sticky='nswe')
numero_cx_custodia_entry.grid(row=7, column=1, padx=10, pady=10, sticky='nswe')
codigo_classificacao_label.grid(row=8, column=0, padx=10, pady=10, sticky='nswe')
codigo_classificacao_entry.grid(row=8, column=1, padx=10, pady=10, sticky='nswe')
datas_limite_label.grid(row=9, column=0, padx=10, pady=10, sticky='nswe')
datas_limite_entry.grid(row=9, column=1, padx=10, pady=10, sticky='nswe')
descricao_documentos_label.grid(row=10, column=0, padx=10, pady=10, sticky='nswe')
descricao_documentos_entry.grid(row=10, column=1, padx=10, pady=10, sticky='nswe')
prazo_guarda_label.grid(row=11, column=0, padx=10, pady=10, sticky='nswe')
prazo_guarda_entry.grid(row=11, column=1, padx=10, pady=10, sticky='nswe')
destinacao_label.grid(row=12, column=0, padx=10, pady=10, sticky='nswe')
destinacao_entry.grid(row=12, column=1, padx=10, pady=10, sticky='nswe')
localizacao_unidade_label.grid(row=13, column=0, padx=10, pady=10, sticky='nswe')
localizacao_unidade_entry.grid(row=13, column=1, padx=10, pady=10, sticky='nswe')
localizacao_conjunto_label.grid(row=14, column=0, padx=10, pady=10, sticky='nswe')
localizacao_conjunto_entry.grid(row=14, column=1, padx=10, pady=10, sticky='nswe')
localizacao_rua_label.grid(row=15, column=0, padx=10, pady=10, sticky='nswe')
localizacao_rua_entry.grid(row=15, column=1, padx=10, pady=10, sticky='nswe')
localizacao_estante_label.grid(row=16, column=0, padx=10, pady=10, sticky='nswe')
localizacao_estante_entry.grid(row=16, column=1, padx=10, pady=10, sticky='nswe')
localizacao_prateleira_label.grid(row=17, column=0, padx=10, pady=10, sticky='nswe')
localizacao_prateleira_entry.grid(row=17, column=1, padx=10, pady=10, sticky='nswe')

# Posicionamento dos botões
inserir_button.grid(row=18, column=0, padx=10, pady=10, sticky='nswe')
pesquisa_label.grid(row=19, column=0, padx=10, pady=10, sticky='nswe')
pesquisa_entry.grid(row=19, column=1, padx=10, pady=10, sticky='nswe')
pesquisar_button.grid(row=20, column=0, padx=10, pady=10, sticky='nswe')
relatorio_button.grid(row=21, column=0, padx=10, pady=10, sticky='nswe')
sair_button.grid(row=22, column=0, padx=10, pady=10, sticky='nswe')


root.mainloop()