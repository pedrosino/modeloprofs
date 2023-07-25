"""Modelo de programação linear para distribuição de vagas entre unidades acadêmicas
de uma universidade federal. Desenvolvido na pesquisa do Mestrado Profissional em
Gestão Organizacional, da Faculdade de Gestão e Negócios, Universidade Federal de Uberlândia,
por Pedro Santos Guimarães, em 2023"""
from datetime import datetime
import sys
import math
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from tkinter import font
import pandas as pd
import numpy as np
from tktooltip import ToolTip
from pulp import lpSum, lpDot, LpVariable, LpStatus, PULP_CBC_CMD, LpProblem, LpMaximize, LpMinimize

# Variáveis globais
MATRIZ_UNIDADES = None
MATRIZ_PERFIS = None
MATRIZ_PEQ = None
MATRIZ_TEMPO = None
PESOS = None
NOMES_RESTRICOES = None

# Dados do problema
N_PERFIS = None
N_RESTRICOES = None
N_UNIDADES = None
CONECTORES = None

#Configurações
# Restrição de carga horária média máxima por unidade
LIMITAR_CH_MINIMA = False
CH_MIN = 12
# Restrição de carga horária média mínima por unidade
LIMITAR_CH_MAXIMA = False
CH_MAX = 16
MAX_TOTAL = False  # Restrição do número máximo total
MIN_TOTAL = False # Restrição do número mínimo total
# Restrição quanto aos perfis 5 e 6 - ministram grande número de aulas
# e não participam de pesquisa ou extensão (podem ser considerados como regime de 40h)
LIMITAR_QUARENTA = False
P_QUARENTA = 0.1 # porcentagem máxima aceita
# Restrição quanto aos perfis 7 e 8 - regime de 20h
LIMITAR_VINTE = False
P_VINTE = 0.2 # porcentagem máxima aceita
# Tempo limite para o programa procurar a solução ótima (em segundos)
#esse parâmetro pode ser alterado pelo usuário ao chamar o programa
TEMPO_LIMITE = 30
modo_escolhido = 'todos'

# Vetor com formato do resultado conforme o modo
formato_resultado = dict()
formato_resultado['num'] = 'professores'
formato_resultado['peq'] = 'prof-equivalente'
formato_resultado['tempo'] = 'horas'
formato_resultado['ch'] = 'aulas/prof'
formato_resultado['todos'] = '(na escala de 0 a 1)'

def check_boxes():
    """Habilita ou desabilita os campos de texto conforme o checkbox"""
    if minima.get():
        entrada_CH_MIN['state'] = tk.NORMAL
    else:
        entrada_CH_MIN['state'] = tk.DISABLED

    if maxima.get():
        entrada_CH_MAX['state'] = tk.NORMAL
    else:
        entrada_CH_MAX['state'] = tk.DISABLED

    if MIN_TOTAL.get():
        entrada_n_MIN_TOTAL['state'] = tk.NORMAL
    else:
        entrada_n_MIN_TOTAL['state'] = tk.DISABLED

    if MAX_TOTAL.get():
        entrada_n_MAX_TOTAL['state'] = tk.NORMAL
    else:
        entrada_n_MAX_TOTAL['state'] = tk.DISABLED

def carregar_arquivo():
    """Carrega os dados da planilha"""
    global MATRIZ_UNIDADES, MATRIZ_PERFIS, MATRIZ_PEQ, MATRIZ_TEMPO, N_RESTRICOES, N_UNIDADES, \
           N_PERFIS, PESOS, NOMES_RESTRICOES, CONECTORES
    # Importa dados do arquivo
    arquivo = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])

    df_todas = pd.read_excel(arquivo, sheet_name=['unidades','perfis'])
    MATRIZ_UNIDADES = df_todas['unidades'].to_numpy()
    MATRIZ_PERFIS = df_todas['perfis'].to_numpy()
    # Os nomes das restrições são as linhas da matriz de perfis
    N_PERFIS = len(MATRIZ_PERFIS[0]) - 2
    N_RESTRICOES = len(MATRIZ_PERFIS) - 2
    NOMES_RESTRICOES = MATRIZ_PERFIS[:, 0].transpose()[0:N_RESTRICOES]
    CONECTORES = MATRIZ_PERFIS[:, N_PERFIS+1].transpose()[0:N_RESTRICOES]
    # Ajusta matriz
    MATRIZ_PERFIS = np.delete(MATRIZ_PERFIS, 0, axis=1)
    MATRIZ_PERFIS = np.delete(MATRIZ_PERFIS, N_PERFIS, axis=1)
    N_UNIDADES = len(MATRIZ_UNIDADES)
    MATRIZ_PEQ = MATRIZ_PERFIS[N_RESTRICOES+1]
    MATRIZ_TEMPO = MATRIZ_PERFIS[N_RESTRICOES] # tempo disponível
    # Lê critérios do AHP
    df_criterios = pd.read_excel(arquivo, sheet_name='criterios', usecols="A:B").dropna()
    pesos_lidos = df_criterios.to_numpy()
    PESOS = np.delete(pesos_lidos, 0, axis=1).transpose()

    print(MATRIZ_UNIDADES)
    print(MATRIZ_PERFIS)
    print(NOMES_RESTRICOES)
    print(CONECTORES)
    print(PESOS)

    if(len(MATRIZ_UNIDADES) and len(MATRIZ_PERFIS)):
        botaoExecutar['state'] = tk.NORMAL

def executar():
    """Executa a otimização"""
    print(MATRIZ_UNIDADES)
    print(MATRIZ_PERFIS)

    # Verifica valores
    resultado.config(text=f"Min: {minima.get()} - {CH_MIN.get()}, Max: {maxima.get()} \
                           - {CH_MAX.get()}\nTmin: {MIN_TOTAL.get()} - {n_MIN_TOTAL.get()},\
                              Tmax: {MAX_TOTAL.get()} - {n_MAX_TOTAL.get()}\n\
                                Limite: {limite.get()}, Restrições: {N_RESTRICOES}")

root = tk.Tk()
root.title("Otimizador de distribuição de professores 1.0 - Pedro Santos Guimarães")
#root.geometry("1200x600")

# Título
textoTitulo = tk.Label(root, text="Bem vindo.", anchor="w", justify="left",
                       font=font.Font(weight="bold"))
textoTitulo.grid(sticky='W', row=0, column=0, columnspan=5, padx=10, pady=10)

# Texto instruções
#textoBotao = tk.Label(root, text="Primeiro escolha o arquivo no botão abaixo.\n\
#                                  Depois verifique as opções e clique em Executar.")
#textoBotao.grid(row=0, column=3)

# Texto botão arquivo
textoBotao = tk.Label(root, text="Escolha o arquivo:")
textoBotao.grid(row=1, column=0)

# Botão para selecionar o arquivo
botaoArquivo = ttk.Button(root, text="Abrir arquivo", command=carregar_arquivo)
botaoArquivo.grid(row=1, column=1, padx=10, pady=10)

# Grupo opções
grupo = ttk.LabelFrame(root, text="Opções")
grupo.grid(row=2, column=0, padx=10, pady=10)

# Checkbox para ch minima
minima = tk.BooleanVar(value=True)
checkbox_minima = tk.Checkbutton(grupo, text="CH mínima: ", variable=minima, command=check_boxes)
checkbox_minima.grid(row=3, column=0, padx=10, pady=10, sticky='e')

# campo texto
CH_MIN = tk.IntVar(value=12)
entrada_CH_MIN = tk.Entry(grupo, textvariable=CH_MIN, width=5)
entrada_CH_MIN.grid(row=3, column=1, padx=10, pady=10, sticky='e')

ToolTip(checkbox_minima, msg="Ativar carga horária média máxima por unidade", delay=0.1)
ToolTip(entrada_CH_MIN, msg="Valor da carga horária máxima", delay=0.1)

# Checkbox para ch maxima
maxima = tk.BooleanVar(value=True)
checkbox_maxima = tk.Checkbutton(grupo, text="CH máxima: ", variable=maxima, command=check_boxes)
checkbox_maxima.grid(row=4, column=0, padx=10, pady=10)

# campo texto
CH_MAX = tk.IntVar(value=16)
entrada_CH_MAX = tk.Entry(grupo, textvariable=CH_MAX, width=5)
entrada_CH_MAX.grid(row=4, column=1, padx=10, pady=10)

ToolTip(checkbox_maxima, msg="Ativar carga horária média mínima geral (para toda a Universidade)",
        delay=0.1)
ToolTip(entrada_CH_MAX, msg="Valor da carga horária mínima", delay=0.1)

# Checkbox para total minimo
MIN_TOTAL = tk.BooleanVar(value=False)
checkbox_MIN_TOTAL = tk.Checkbutton(grupo, text="Total mínimo: ",
                                    variable=MIN_TOTAL, command=check_boxes)
checkbox_MIN_TOTAL.grid(row=3, column=3, padx=10, pady=10)

# campo texto
n_MIN_TOTAL = tk.IntVar()
entrada_n_MIN_TOTAL = tk.Entry(grupo, textvariable=n_MIN_TOTAL, width=5, state=tk.DISABLED)
entrada_n_MIN_TOTAL.grid(row=3, column=4, padx=10, pady=10)

ToolTip(checkbox_MIN_TOTAL, msg="Ativar número mínimo total de professores", delay=0.1)
ToolTip(entrada_n_MIN_TOTAL, msg="Valor do mínimo total", delay=0.1)

# Checkbox para total maxima
MAX_TOTAL = tk.BooleanVar(value=False)
checkbox_MAX_TOTAL = tk.Checkbutton(grupo, text="Total máximo: ",
                                    variable=MAX_TOTAL, command=check_boxes)
checkbox_MAX_TOTAL.grid(row=4, column=3, padx=10, pady=10)

# campo texto
n_MAX_TOTAL = tk.IntVar()
entrada_n_MAX_TOTAL = tk.Entry(grupo, textvariable=n_MAX_TOTAL, width=5, state=tk.DISABLED)
entrada_n_MAX_TOTAL.grid(row=4, column=4, padx=10, pady=10)

ToolTip(checkbox_MAX_TOTAL, msg="Ativar número máximo total de professores", delay=0.1)
ToolTip(entrada_n_MAX_TOTAL, msg="Valor do máximo total", delay=0.1)

# Tempo limite
limite = tk.IntVar(value=30)
label_nome_saida = tk.Label(grupo, text="Tempo limite:")
label_nome_saida.grid(row=5, column=0, padx=10, pady=10)
entrada_tempo_limite = tk.Entry(grupo, textvariable=limite, width=5)
entrada_tempo_limite.grid(row=5, column=1, padx=10, pady=10)

ToolTip(entrada_tempo_limite, msg="Tempo máximo para procurar a solução ótima", delay=0.1)

# Botão para executar
botaoExecutar = ttk.Button(root, text="Executar", state=tk.DISABLED, command=executar)
botaoExecutar.grid(row=7, column=0, padx=10, pady=10)

# Dummy para ajustar as colunas
resultado = tk.Label(root)
resultado.grid(row=8, column=2, padx=10, pady=10)

root.mainloop()
