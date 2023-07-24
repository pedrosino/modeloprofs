import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from tkinter import font
import pandas as pd
import numpy as np
from tktooltip import ToolTip

# Variáveis globais
m_unidades = None
m_perfis = None
n_restricoes = None

def check_boxes():
    if minima.get():
        entrada_ch_min['state'] = tk.NORMAL
    else:
        entrada_ch_min['state'] = tk.DISABLED
        
    if maxima.get():
        entrada_ch_max['state'] = tk.NORMAL
    else:
        entrada_ch_max['state'] = tk.DISABLED
    
    if min_total.get():
        entrada_n_min_total['state'] = tk.NORMAL
    else:
        entrada_n_min_total['state'] = tk.DISABLED
        
    if max_total.get():
        entrada_n_max_total['state'] = tk.NORMAL
    else:
        entrada_n_max_total['state'] = tk.DISABLED

def carregar_arquivo():
    global m_unidades, m_perfis
    arquivo = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    # Carregar arquivo da planilha
    df_todas = pd.read_excel(arquivo, sheet_name=['unidades','perfis'])
    m_unidades = df_todas['unidades'].to_numpy()
    m_perfis = df_todas['perfis'].to_numpy()
    m_perfis = np.delete(m_perfis, 0, axis=1)
    botaoExecutar['state'] = tk.NORMAL 
    
#def open_file_dialog():
#    global file_path
#    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])

def executar():
    print(m_unidades)
    print(m_perfis)
    
    # Verifica valores
    resultado.config(text=f"Min: {minima.get()} - {ch_min.get()}, Max: {maxima.get()} - {ch_max.get()}\nTmin: {min_total.get()} - {n_min_total.get()}, Tmax: {max_total.get()} - {n_max_total.get()}\nLimite: {limite.get()}, Restrições: {n_restricoes.get()}")

root = tk.Tk()
root.title("Otimizador de distribuição de professores 1.0 - Pedro Santos Guimarães")
#root.geometry("1200x600")

# Título
textoOpcoes = tk.Label(root, text="Bem vindo.", anchor="w", justify="left", font=font.Font(weight="bold"))
textoOpcoes.grid(sticky='W', row=0, column=0, columnspan=5, padx=10, pady=10)

# Texto instruções
#textoBotao = tk.Label(root, text="Primeiro escolha o arquivo no botão abaixo.\nDepois verifique as opções e clique em Executar.")
#textoBotao.grid(row=0, column=3)

# Texto botão arquivo
textoBotao = tk.Label(root, text="Escolha o arquivo:")
textoBotao.grid(row=1, column=0)

# Botão para selecionar o arquivo
botaoArquivo = ttk.Button(root, text="Abrir arquivo", command=carregar_arquivo)
botaoArquivo.grid(row=1, column=1, padx=10, pady=10)

# Texto opções
textoOpcoes = tk.Label(root, text="Opções:", anchor="w", justify="left", font=font.Font(weight="bold"))
textoOpcoes.grid(sticky='W', row=2, column=0, padx=10, pady=10)

# Checkbox para ch minima
minima = tk.BooleanVar(value=True)
checkbox_minima = tk.Checkbutton(root, text="CH mínima: ", variable=minima, command=check_boxes)
checkbox_minima.grid(row=3, column=0, padx=10, pady=10)

# campo texto
ch_min = tk.IntVar(value=12)
entrada_ch_min = tk.Entry(root, textvariable=ch_min, width=5)
entrada_ch_min.grid(row=3, column=1, padx=10, pady=10)

ToolTip(checkbox_minima, msg="Ativar carga horária média máxima por unidade", delay=0.1)
ToolTip(entrada_ch_min, msg="Valor da carga horária máxima", delay=0.1)

# Checkbox para ch maxima
maxima = tk.BooleanVar(value=True)
checkbox_maxima = tk.Checkbutton(root, text="CH máxima: ", variable=maxima, command=check_boxes)
checkbox_maxima.grid(row=4, column=0, padx=10, pady=10)

# campo texto
ch_max = tk.IntVar(value=16)
entrada_ch_max = tk.Entry(root, textvariable=ch_max, width=5)
entrada_ch_max.grid(row=4, column=1, padx=10, pady=10)

ToolTip(checkbox_maxima, msg="Ativar carga horária média mínima geral (para toda a Universidade)", delay=0.1)
ToolTip(entrada_ch_max, msg="Valor da carga horária mínima", delay=0.1)

# Checkbox para total minimo
min_total = tk.BooleanVar(value=False)
checkbox_min_total = tk.Checkbutton(root, text="Total mínimo: ", variable=min_total, command=check_boxes)
checkbox_min_total.grid(row=3, column=3, padx=10, pady=10)

# campo texto
n_min_total = tk.IntVar()
entrada_n_min_total = tk.Entry(root, textvariable=n_min_total, width=5, state=tk.DISABLED)
entrada_n_min_total.grid(row=3, column=4, padx=10, pady=10)

ToolTip(checkbox_min_total, msg="Ativar número mínimo total de professores", delay=0.1)
ToolTip(entrada_n_min_total, msg="Valor do mínimo total", delay=0.1)

# Checkbox para total maxima
max_total = tk.BooleanVar(value=False)
checkbox_max_total = tk.Checkbutton(root, text="Total máximo: ", variable=max_total, command=check_boxes)
checkbox_max_total.grid(row=4, column=3, padx=10, pady=10)

# campo texto
n_max_total = tk.IntVar()
entrada_n_max_total = tk.Entry(root, textvariable=n_max_total, width=5, state=tk.DISABLED)
entrada_n_max_total.grid(row=4, column=4, padx=10, pady=10)

ToolTip(checkbox_max_total, msg="Ativar número máximo total de professores", delay=0.1)
ToolTip(entrada_n_max_total, msg="Valor do máximo total", delay=0.1)

# Tempo limite
limite = tk.IntVar(value=30)
label_nome_saida = tk.Label(root, text="Tempo limite:")
label_nome_saida.grid(row=5, column=0, padx=10, pady=10)
entrada_tempo_limite = tk.Entry(root, textvariable=limite, width=5)
entrada_tempo_limite.grid(row=5, column=1, padx=10, pady=10)

ToolTip(entrada_tempo_limite, msg="Tempo máximo para procurar a solução ótima", delay=0.1)

# Número de restrições
n_restricoes = tk.IntVar(value=5)
label_restricoes = tk.Label(root, text="Número de restrições:")
label_restricoes.grid(row=6, column=0, padx=10, pady=10)
entrada_restricoes = tk.Entry(root, textvariable=n_restricoes, width=5)
entrada_restricoes.grid(row=6, column=1, padx=10, pady=10)

ToolTip(entrada_tempo_limite, msg="Tempo máximo para procurar a solução ótima", delay=0.1)

# Botão para executar
botaoExecutar = ttk.Button(root, text="Executar", state=tk.DISABLED, command=executar)
botaoExecutar.grid(row=7, column=0, padx=10, pady=10)

# Dummy para ajustar as colunas
resultado = tk.Label(root)
resultado.grid(row=8, column=2, padx=10, pady=10)

root.mainloop()