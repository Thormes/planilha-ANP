import os
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, font, Tk
from tkinter import Label, Frame, Button, Entry
from tkinter.constants import *
from read_planilha_pandas import analisa_planilha

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)



window = Tk()
window.title("Análise planilha ANP")
window.iconbitmap(resource_path("icon.ico"))
window.config(padx=10, pady=10)

arquivos_folder = os.getcwd()
planilha_file = ''
categorias_originais = ['ANP RETENÇÃO', 'AQUISIÇÃO DE CONCESSÃO', 'G&G', 'G&A','SÍSMICA','SSMA', 'PERFURAÇÃO']
categorias_pre_selecionadas = ['G&G','SÍSMICA','SSMA','PERFURAÇÃO']



def choose_directory():
    folder = filedialog.askdirectory(initialdir=os.getcwd())
    global arquivos_folder
    if len(folder) > 0:
        pasta_input.delete(0, END)
        pasta_input.insert(0, folder)
        arquivos_folder = folder


def choose_file():
    file = filedialog.askopenfilename(initialdir=os.getcwd(), initialfile="planilha.xlsx")
    global planilha_file
    if len(file) > 0:
        planilha_input.delete(0, END)
        planilha_input.insert(0, file)
        planilha_file = file


def get_palavras(palavras: str):
    if len(palavras) == 0:
        return []
    return set([palavra.strip() for palavra in palavras.split(";") if len(palavra) > 0])

def get_categorias_selecionadas():
    selecionadas = []
    for i in range(len(categorias_originais)):
        if checkbuttons_categorias[i].get():
            selecionadas.append(categorias_originais[i])
    return selecionadas

def analisar_planilha():
    palavras = get_palavras(palavras_input.get())
    cat = get_categorias_selecionadas()
    if not os.path.exists(planilha_file):
        messagebox.showerror("Arquivo de Planilha Não encontrado","Não foi possível encontrar o arquivo da planilha")
        return
    if not os.path.exists(arquivos_folder):
        messagebox.showerror("Pasta de Arquivos Ausente",
            "Não foi possível encontrar a pasta onde estão os arquivos de texto")
        return
    analisa_planilha(planilha_file, cat, palavras, arquivos_folder)
    pla_folder, pla_file = os.path.split(planilha_file)
    messagebox.showinfo("Análise Concluída",f"Análise concluída. Resultado salvo em {os.path.join(pla_folder, 'resultado.xlsx')}")


ipadding = {'ipadx': 5, 'ipady': 5, "padx": 15, "pady": 15}

# escolher pasta arquivos texto
pasta_frame = Frame(window)
pasta_label = Label(pasta_frame, text="Localização Arquivos Texto:", anchor="w")
pasta_label.pack(**ipadding, side=LEFT)
pasta_input = Entry(pasta_frame)
pasta_input.insert(0, arquivos_folder)
pasta_input.pack(**ipadding, side=LEFT, fill=X, expand=True)
directory_button = Button(pasta_frame, text="Escolher Pasta", command=choose_directory)
directory_button.pack(**ipadding, side=LEFT)
pasta_frame.pack(expand=True, fill=BOTH)

# escolher planilha
planilha_frame = Frame(window)
planilha_label = Label(planilha_frame, text="Planilha Original:", anchor="w")
planilha_label.pack(**ipadding, side=LEFT)
planilha_input = Entry(planilha_frame)
planilha_input.insert(0, planilha_file)
planilha_input.pack(**ipadding, side=LEFT, fill=X, expand=True)
arquivo_button = Button(planilha_frame, text="Escolher Arquivo", command=choose_file)
arquivo_button.pack(**ipadding, side=LEFT)
planilha_frame.pack(expand=True, fill=BOTH)

# input palavras
palavras_frame = Frame(window)
palavras_label = Label(palavras_frame, text='Palavras a procurar (separadas por  ";" ):', anchor="w", justify="left")
palavras_label.pack(**ipadding, fill=tk.X, expand=False, side=tk.LEFT)
palavras_input = Entry(palavras_frame, width=45)
palavras_input.insert(0, "")
palavras_input.pack(**ipadding, fill=tk.X, expand=True, side=tk.LEFT)
palavras_input.focus()
palavras_frame.pack(fill=tk.BOTH, expand=True)

# Criação dos checkbuttons para cada categoria
categorias_frame = Frame(window)
checkbuttons_categorias = []
label_categorias = Label(categorias_frame, text="Selecione a(s) categoria(s)",
                          font="Verdana 10 bold")
label_categorias.pack()
for categoria in categorias_originais:
    var = tk.BooleanVar()
    if categoria in categorias_pre_selecionadas:
        var.set(True)
    checkbutton = ttk.Checkbutton(categorias_frame, text=categoria, variable=var)
    checkbutton.pack(anchor=tk.W)
    checkbuttons_categorias.append(var)

categorias_frame.pack(anchor=tk.W)

# iniciar execucao
download_button = Button(text="Iniciar Analise", width=40, command=analisar_planilha)
download_button.pack(**ipadding)

window.mainloop()
