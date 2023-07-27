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

LISTA_MODOS = {
    "Menor número": "num",
    "Menos P-Eq": "peq",
    "Mais tempo": "tempo",
    "Equilíbrio CH": "ch",
    "Todos": "todos"
}

# Vetor de modelos
MODELOS = {}

# Vetor com formato do resultado conforme o modo
FORMATO_RESULTADO = {}
FORMATO_RESULTADO['num'] = 'professores'
FORMATO_RESULTADO['peq'] = 'prof-equivalente'
FORMATO_RESULTADO['tempo'] = 'horas'
FORMATO_RESULTADO['ch'] = 'aulas/prof'
FORMATO_RESULTADO['todos'] = '(na escala de 0 a 1)'

# Restrição de carga horária média máxima por unidade
LIMITAR_CH_MINIMA = False
CH_MIN = 16
# Restrição de carga horária média mínima por unidade
LIMITAR_CH_MAXIMA = False
CH_MAX = 16
# Restrição do número máximo total
MAX_TOTAL = False
N_MAX_TOTAL = None
# Restrição do número mínimo total
MIN_TOTAL = False
N_MIN_TOTAL = None
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
MODO_ESCOLHIDO = 'todos'

def verifica_executar():
    """Habilita ou desabilita o botão Executar"""
    if combo_var.get() and radio_var.get():
        botaoExecutar['state'] = tk.NORMAL
    else:
        botaoExecutar['state'] = tk.DISABLED

def verifica_check_boxes():
    """Habilita ou desabilita os campos de texto conforme o checkbox"""
    if bool_minima.get():
        entrada_CH_MIN['state'] = tk.NORMAL
    else:
        entrada_CH_MIN['state'] = tk.DISABLED

    if bool_maxima.get():
        entrada_CH_MAX['state'] = tk.NORMAL
    else:
        entrada_CH_MAX['state'] = tk.DISABLED

    if bool_min_total.get():
        entrada_N_MIN_total['state'] = tk.NORMAL
    else:
        entrada_N_MIN_total['state'] = tk.DISABLED

    if bool_max_total.get():
        entrada_N_MAX_total['state'] = tk.NORMAL
    else:
        entrada_N_MAX_total['state'] = tk.DISABLED

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

    # Se a importação teve sucesso
    if(len(MATRIZ_UNIDADES) and len(MATRIZ_PERFIS)):
        # Mostra as opções
        grupo.grid(row=2, column=0, padx=10, pady=10)
        verifica_check_boxes()

def executar():
    """Executa a otimização"""
    global LIMITAR_CH_MAXIMA, LIMITAR_CH_MINIMA, CH_MAX, CH_MIN, MAX_TOTAL, MIN_TOTAL,\
        TEMPO_LIMITE, N_MIN_TOTAL, N_MAX_TOTAL, MODO_ESCOLHIDO

    #print(MATRIZ_UNIDADES)
    #print(MATRIZ_PERFIS)
    #print(NOMES_RESTRICOES)
    #print(CONECTORES)
    #print(PESOS)

    # Captura opções e valores escolhidos
    LIMITAR_CH_MAXIMA = bool_maxima.get()
    LIMITAR_CH_MINIMA = bool_minima.get()
    if LIMITAR_CH_MAXIMA:
        CH_MAX = texto_ch_max.get()
    if LIMITAR_CH_MINIMA:
        CH_MIN = texto_ch_min.get()
    MAX_TOTAL = bool_max_total.get()
    MIN_TOTAL = bool_min_total.get()
    if MAX_TOTAL:
        N_MAX_TOTAL = texto_max_total.get()
    if MIN_TOTAL:
        N_MIN_TOTAL = texto_min_total.get()
    TEMPO_LIMITE = val_limite.get()
    MODO_ESCOLHIDO = LISTA_MODOS[combo_var.get()]

    print(MODO_ESCOLHIDO)

    # Verifica modo escolhido
    if MODO_ESCOLHIDO not in ['num', 'peq', 'tempo', 'ch', 'todos']:
        MODO_ESCOLHIDO = 'todos'
        print("O modo escolhido era inválido. Será utilizado o modo 'todos'.")

    # Verifica números totais
    if (MIN_TOTAL and N_MIN_TOTAL < 1) or (MAX_TOTAL and N_MAX_TOTAL < 1):
        print("Os números totais especificados são inválidos. Essas opções foram desativadas.")
        MIN_TOTAL = False
        MAX_TOTAL = False

    # Verifica porcentagens
    #if P_QUARENTA > 1 or P_VINTE > 1 or P_QUARENTA < 0 or P_VINTE < 0:
    #    print("Os percentuais especificados são inválidos. Essas opções foram desativadas.")
    #    LIMITAR_QUARENTA = False
    #    LIMITAR_VINTE = False

    print("Bem vindo!")
    print(f"O modo escolhido foi {MODO_ESCOLHIDO}")
    if MODO_ESCOLHIDO == 'todos':
        print("Primeiramente vamos definir os parâmetros para cada critério/objetivo")

    # Arquivo de texto de saída
    original_stdout = sys.stdout
    filename = f"CBC completo_{MODO_ESCOLHIDO}_N={N_UNIDADES} {datetime.now().strftime('%H-%M-%S')}.txt"
    fileout =  open(filename, 'w', encoding='utf-8')
    sys.stdout = fileout

    # Conforme o modo escolhido, faz só uma otimização ou todas
    if MODO_ESCOLHIDO != 'todos':
        resultado_final, qtdes_final = otimizar(MODO_ESCOLHIDO, fileout, original_stdout, None, None)
    else:
        # Critérios/modos
        modos = np.array(['num', 'peq', 'tempo', 'ch'])

        # Lista dos melhores e piores casos
        melhores = {}
        piores = {}

        # Valores conhecidos a priori
        # número máximo é dado pelo total de aulas dividido pela carga horária mínima
        # ou é definido pelo usuário
        # Para avaliar cada cenário em relação a esses critérios, é necessário estabelecer
        # uma forma de pontuação
        # Essa pontuação varia entre 0 (pior caso) e 1 (melhor caso), e será multiplicada
        # pelo peso de cada critério para obter a pontuação total daquele cenário. Para cada
        # critério é necessário então determinar o pior e o melhor caso. Nos dois primeiros
        # o pior caso é quando o total de professores é número máximo possível, com base na
        # carga horária média mínima. Por exemplo, se são 900 aulas e a carga horária mínima
        # foi definida em 12 aulas por professor, o número máximo possível é 75. Para o
        # critério 1 esse é o valor a ser considerado. Para o critério 2, o pior valor seria
        # 75*1,65, que é o fator do professor 40h-DE. Já o melhor caso não pode ser determinado
        # a priori, pois o número de professores deve atender às restrições de aulas
        # e orientações, por exemplo. Assim, é necessário resolver o modelo com cada critério
        # e registrar o valor ótimo obtido para ser a base da escala. Para esses dois critérios
        # a escala é invertida, ou seja, quanto mais professores, menor a pontuação.
        # No caso do critério 3 a lógica é inversa, quanto mais tempo disponível, melhor o
        # cenário. O pior caso é quando não há tempo nenhum (valor 0), e o melhor caso deve
        # ser determinado resolvendo o modelo usando esse critério.
        # Para o critério 4, o melhor caso é determinado pela solução inicial do modelo com
        # este critério. O pior caso é quando metade das unidades estiver na carga horária
        # máxima e a outra metade na mínima. A média geral seria o valor intermediário entre
        # as duas e o desvio total é dado por N_UNIDADES*(CH_MAX - CH_MIN)/2
        # Exemplo:
        #
        # CH_MAX      = 16  --   d1 = d2 = (CH_MAX - CH_MIN)/2
        #                   |
        #                   d1   desvio total = N_UNIDADES/2*d1 + N_UNIDADES/2*d2
        #                   |
        # média geral = 14  --   desvio total = N_UNIDADES*d1 = N_UNIDADES*(CH_MAX - CH_MIN)/2
        #                   |
        #                   d2
        #                   |
        # CH_MIN      = 12  --
        numero_max = \
            round(np.sum(MATRIZ_UNIDADES[:N_UNIDADES], axis=0)[1] / CH_MIN) if not MAX_TOTAL else N_MAX_TOTAL
        piores['num'] = numero_max
        piores['peq'] = numero_max*1.65
        piores['tempo'] = 0
        piores['ch'] = (CH_MAX - CH_MIN)/2

        # Percorre os critérios
        for modo_usado in modos:
            resultado_modo, qtdes_modo = otimizar(modo_usado, fileout, original_stdout, piores, melhores)

            # Registra o resultado na lista de melhores casos
            melhores[modo_usado] = resultado_modo
            sys.stdout = original_stdout
            print(f"Ok. Resultado: {melhores[modo_usado]}")
            sys.stdout = fileout

            # Imprime resultados
            imprimir_resultados(qtdes_modo)
            print("------------------------------------------------------------")
            print("")

        ##### -----------------Fim da primeira 'rodada'-----------------
        print("Melhores:")
        print(melhores)
        print("Piores:")
        print(piores)
        print()
        print("------------------------------------------------------------")

        ## Agora uma nova rodada do modelo usando os pesos e as listas de melhores e piores casos
        resultado_final, qtdes_final = otimizar('todos', fileout, original_stdout, piores, melhores)

    # ---- Finalização ----
    # Imprime resultados
    imprimir_resultados(qtdes_final)
    # Imprime parâmetros
    imprimir_parametros(qtdes_final)

    if MODO_ESCOLHIDO == 'todos':
        # PESOS
        print(f"PESOS: {PESOS}")
        # Imprime médias e desvios
        for variable in MODELOS['todos'].variables():
            if variable.name[0] == 'p' or variable.name[0] == 'z':
                print(f"{variable.name}: {variable.value():.4f}")

    # Imprime o modelo completo
    print("")
    print("------------------Modelo:------------------")
    print(MODELOS[MODO_ESCOLHIDO])

    # Fecha arquivo texto
    fileout.close()

    #Imprime na tela
    sys.stdout = original_stdout
    print(f"Situação: {MODELOS[MODO_ESCOLHIDO].status}, {LpStatus[MODELOS[MODO_ESCOLHIDO].status]}")
    #objetivo = f"{MODELOS[MODO_ESCOLHIDO].objective.value():.4f}"
    print(f"Objetivo: {resultado_final} {FORMATO_RESULTADO[MODO_ESCOLHIDO]}")
    print(f"Resolvido em {MODELOS[MODO_ESCOLHIDO].solutionTime} segundos")
    print("")
    print(f"Verifique o arquivo {filename} para o relatório completo")
    print("")

    resultado.config(text=f"{qtdes_final}")

    #salva em planilha
    data_frame = pd.DataFrame(qtdes_final, columns=[f'x{i}' for i in range(1, N_PERFIS+1)])
    data_frame.insert(0, "Unidade", MATRIZ_UNIDADES[:, 0])
    data_frame.to_excel('CBC_Completo.xlsx', sheet_name='Resultados', index=False)

def otimizar(modo, arquivo_saida, stdout, piores, melhores):
    """Função que faz a otimização conforme o modo escolhido"""
    #global MODELOS

    #nomes
    nomes = [str(perfil) + "_"
        + MATRIZ_UNIDADES[und][0] for und in range(N_UNIDADES) for perfil in range(1, N_PERFIS+1)]

    # Variáveis de decisão
    var_x = LpVariable.matrix("x", nomes, cat="Integer", lowBound=0)
    saida = np.array(var_x).reshape(N_UNIDADES, N_PERFIS)

    sys.stdout = stdout
    print(f"Modo: {modo}")
    sys.stdout = arquivo_saida
    minima = LIMITAR_CH_MINIMA
    maxima = LIMITAR_CH_MAXIMA
    # Ativa carga horária mínima e máxima, para que todos os modos tenham as mesmas restrições
    if MODO_ESCOLHIDO == 'todos' or modo == 'ch':
        minima = True
        maxima = True
    # No modo tempo é necessário estabelecer a carga horária mínima (ou número máximo)
    if modo == 'tempo':
        minima = True

    # -- Definir o modelo --
    if modo in ('tempo', 'todos'):
        MODELOS[modo] = LpProblem(name=f"Professores-{modo}", sense=LpMaximize)
    else:
        MODELOS[modo] = LpProblem(name=f"Professores-{modo}", sense=LpMinimize)

    # -- Restrições --
    for restricao in range(N_RESTRICOES):
        for unidade in range(N_UNIDADES):
            match CONECTORES[restricao]:
                case ">=":
                    MODELOS[modo] += lpDot(saida[unidade], MATRIZ_PERFIS[restricao]) \
                        >= MATRIZ_UNIDADES[unidade][restricao+1], \
                        NOMES_RESTRICOES[restricao] + " " + MATRIZ_UNIDADES[unidade][0]
                case "==":
                    MODELOS[modo] += lpDot(saida[unidade], MATRIZ_PERFIS[restricao]) \
                        == MATRIZ_UNIDADES[unidade][restricao+1], \
                        NOMES_RESTRICOES[restricao] + " " + MATRIZ_UNIDADES[unidade][0]
                case "<=":
                    MODELOS[modo] += lpDot(saida[unidade], MATRIZ_PERFIS[restricao]) \
                        <= MATRIZ_UNIDADES[unidade][restricao+1], \
                        NOMES_RESTRICOES[restricao] + " " + MATRIZ_UNIDADES[unidade][0]

    # Restrições de regimes de trabalho
    if LIMITAR_QUARENTA:
        for unidade in range(N_UNIDADES):
            MODELOS[modo] += saida[unidade][4] + saida[unidade][5] - P_QUARENTA*lpSum(saida[unidade]) \
                <= 0, f"40 horas {MATRIZ_UNIDADES[unidade][0]} <= 10%"
    if LIMITAR_VINTE:
        for unidade in range(N_UNIDADES):
            MODELOS[modo] += saida[unidade][6] + saida[unidade][7] - P_VINTE*lpSum(saida[unidade]) \
                <= 0, f"20 horas {MATRIZ_UNIDADES[unidade][0]} <= 20%"

    # Restrições de máximo e mínimo por unidade -> carga horária média
    # ------------------
    # Exemplo para 900 aulas, considerando-se carga horária média mínima de 12 aulas e máxima de 16
    # soma <= 900/12 -> soma <= 75
    # soma >= 900/16 -> soma >= 56.25
    # a soma deve estar entre 57 e 75, inclusive
    for unidade in range(N_UNIDADES):
        if maxima:
            MODELOS[modo] += CH_MAX*lpSum(saida[unidade]) >= MATRIZ_UNIDADES[unidade][1], \
                f"{MATRIZ_UNIDADES[unidade][0]}_chmax: {math.ceil(MATRIZ_UNIDADES[unidade][1]/CH_MAX)}"
        # Restrição mínima somente no geral -> permite exceções como o IBTEC em Monte Carmelo,
        # que tem 19 aulas apenas
        #if minima:
        #    MODELOS[modo] += CH_MIN*lpSum(saida[u]) <= MATRIZ_UNIDADES[u][1], \
        # f"{MATRIZ_UNIDADES[u][0]}_chmin: {int(MATRIZ_UNIDADES[u][1]/CH_MIN)}"

    # Restrições no geral
    if maxima:
        MODELOS[modo] += CH_MAX*lpSum(saida) >= np.sum(MATRIZ_UNIDADES[:N_UNIDADES], axis=0)[1], \
            f"chmax: {math.ceil(np.sum(MATRIZ_UNIDADES[:N_UNIDADES], axis=0)[1]/CH_MAX)}"
    if minima:
        MODELOS[modo] += CH_MIN*lpSum(saida) <= np.sum(MATRIZ_UNIDADES[:N_UNIDADES], axis=0)[1], \
            f"chmin: {int(np.sum(MATRIZ_UNIDADES[:N_UNIDADES], axis=0)[1]/CH_MIN)}"

    # Restrição do número total de professores
    if MAX_TOTAL:
        MODELOS[modo] += lpSum(saida) <= N_MAX_TOTAL, f"TotalMax: {N_MAX_TOTAL}"
    if MIN_TOTAL:
        MODELOS[modo] += lpSum(saida) >= N_MIN_TOTAL, f"TotalMin: {N_MIN_TOTAL}"

    # Modo de equilíbrio da carga horária
    if modo in ('ch', 'todos'):
        # variáveis auxiliares
        # Para cada unidade há um valor da média e um desvio em relação à média geral
        desvios = LpVariable.matrix("zd", MATRIZ_UNIDADES[:N_UNIDADES, 0], cat="Continuous", lowBound=0)
        medias = LpVariable.matrix("zm", MATRIZ_UNIDADES[:N_UNIDADES, 0], cat="Continuous", lowBound=0)
        media_geral = LpVariable("zmg", cat="Continuous", lowBound=0)

        # Como a média seria uma função não linear (aulas/professores), foi feita uma
        # aproximação com uma reta que passa pelos dois pontos extremos dados pelos
        # valores de carga horária mínima e máxima.
        # Coeficientes
        # coef = (y1-y0)/(x1-x0) -> y1 = media minima, y0 = media maxima,
        # x1 = numero maximo, x0 = numero minimo
        x_1 = np.sum(MATRIZ_UNIDADES[:N_UNIDADES], axis=0)[1] / CH_MIN
        x_0 = np.sum(MATRIZ_UNIDADES[:N_UNIDADES], axis=0)[1] / CH_MAX
        coef = (CH_MIN - CH_MAX) / (x_1 - x_0)

        # O cálculo da média é inserido no modelo como uma restrição
        MODELOS[modo] += media_geral == coef*lpSum(saida[:N_UNIDADES]) + CH_MIN + CH_MAX, "Media geral"

        for unidade in range(N_UNIDADES):
            x1u = MATRIZ_UNIDADES[unidade][1] / CH_MIN
            x0u = MATRIZ_UNIDADES[unidade][1] / CH_MAX
            coef_u = (CH_MIN - CH_MAX) / (x1u - x0u)
            # Cálculo da media
            MODELOS[modo] += medias[unidade] == coef_u*lpSum(saida[unidade]) + CH_MIN + CH_MAX, \
                f"{MATRIZ_UNIDADES[unidade][0]}_ch_media"
            # Cálculo do desvio
            # O desvio é dado pelo módulo da subtração, porém o PuLP não aceita a função abs()
            # Assim, são colocadas duas restrições, uma usando o valor positivo e outra o negativo
            MODELOS[modo] += desvios[unidade] >= medias[unidade] - media_geral, \
                f"{MATRIZ_UNIDADES[unidade][0]}_up"
            MODELOS[modo] += desvios[unidade] >= -1*(medias[unidade] - media_geral), \
                f"{MATRIZ_UNIDADES[unidade][0]}_low"

    # Modo com todos os critérios
    if modo == 'todos':
        # Variáveis com as pontuações
        pontuacoes = LpVariable.matrix("p", range(4), cat="Continuous", lowBound=0, upBound=1)
        # Restrições/cálculos
        # caso seja dado um número exato, é necessário alterar a pontuação do critério 'num'
        # para evitar a divisão por zero
        if MAX_TOTAL and MIN_TOTAL and N_MAX_TOTAL == N_MIN_TOTAL:
            MODELOS[modo] += pontuacoes[0] == lpSum(saida)/melhores['num'], "Pontuação número"
        else:
            MODELOS[modo] += pontuacoes[0] == (lpSum(saida) - piores['num'])/(melhores['num'] - piores['num']), "Pontuação número"
        MODELOS[modo] += pontuacoes[1] == (lpSum(saida*MATRIZ_PEQ) - piores['peq'])/(melhores['peq'] - piores['peq']), "Pontuação P-Eq"
        MODELOS[modo] += pontuacoes[2] == (lpSum(saida*MATRIZ_TEMPO) - np.sum(MATRIZ_UNIDADES[:N_UNIDADES], axis=0)[2] - piores['tempo'])/(melhores['tempo'] - piores['tempo']), "Pontuação tempo"
        MODELOS[modo] += pontuacoes[3] == (lpSum(desvios[:N_UNIDADES])/N_UNIDADES - piores['ch'])/(melhores['ch'] - piores['ch']), "Pontuação Equilíbrio"

    # -- Função objetivo --
    if modo == 'num':
        MODELOS[modo] += lpSum(saida)
    elif modo == 'peq':
        MODELOS[modo] += lpSum(saida*MATRIZ_PEQ)
    # No modo tempo é necessário deduzir do tempo calculado as horas
    # que serão destinadas às orientações
    elif modo == 'tempo':
        MODELOS[modo] += lpSum(saida*MATRIZ_TEMPO) - np.sum(MATRIZ_UNIDADES[:N_UNIDADES], axis=0)[2]
    # No modo de equilíbrio a medida de desempenho é o desvio médio
    elif modo == 'ch':
        MODELOS[modo] += lpSum(desvios[:N_UNIDADES])/N_UNIDADES
    elif modo == 'todos':
        # O objetivo é maximimizar a pontuação dada pela soma de cada
        # pontuação multiplicada por seu peso
        MODELOS[modo] += lpSum(pontuacoes*PESOS)

    # Imprime os resultados no arquivo txt
    print(f'Modo: {modo}')
    print(f'Unidades: {N_UNIDADES}')
    print(f'CH Maxima: {maxima} {CH_MAX if maxima else ""}')
    print(f'CH Minima: {minima} {CH_MIN if minima else ""}')
    print(f'Total: {N_MIN_TOTAL if MIN_TOTAL else "-"} a {N_MAX_TOTAL if MAX_TOTAL else "-"}')
    print(f'Vinte: {LIMITAR_VINTE}')
    print(f'Quarenta: {LIMITAR_QUARENTA}')
    print()

    # Ajusta limite
    # Para os modos "intermediários" não é necessário um limite maior que 30 segundos
    if MODO_ESCOLHIDO == 'todos' and modo != 'todos' and TEMPO_LIMITE > 30:
        novo_limite = 30
    else:
        novo_limite = TEMPO_LIMITE

    # Resolver o modelo
    MODELOS[modo].solve(PULP_CBC_CMD(msg=0, timeLimit=novo_limite))

    # Resultados
    print(f"Situação: {MODELOS[modo].status}, {LpStatus[MODELOS[modo].status]}")
    # Para cada critério o resultado é em um formato diferente
    if modo == 'ch':
        objetivo = MODELOS[modo].objective.value()
    elif modo == 'peq':
        objetivo = round(MODELOS[modo].objective.value(), 2)
    elif modo == 'num':
        objetivo = int(MODELOS[modo].objective.value())
    elif modo == 'tempo':
        objetivo = MODELOS[modo].objective.value()
    elif modo == 'todos':
        objetivo = round(MODELOS[modo].objective.value(), 4)

    print(f"Objetivo: {objetivo} {FORMATO_RESULTADO[modo]}")
    print(f"Resolvido em {MODELOS[modo].solutionTime} segundos")
    print("")

    # Extrai as quantidades
    qtdes_saida = np.full((N_UNIDADES, N_PERFIS), 0, dtype=int)
    for var in MODELOS[modo].variables():
        if var.name.find('x') == 0:
            _, perfil, unidade = var.name.split("_")
            perfil = int(perfil) - 1
            ind_unidade = np.where(MATRIZ_UNIDADES[:N_UNIDADES, 0] == unidade)[0][0]
            qtdes_saida[ind_unidade][perfil] = int(var.value())

    # Retorna o valor da função objetivo e as quantidades
    return objetivo, qtdes_saida

def imprimir_resultados(qtdes):
    """Imprime resultados da quantidade de cada perfil em cada unidade"""
    print("Resultados:")
    print("---------+" + "-----"*N_PERFIS + "-+-------+---------+----------+------------+")
    print("Unidade  |  " + "  ".join([f"{i: >3}" for i in [f"x{p+1}" for p in range(N_PERFIS)]]) +
          " | Total |   P-Eq  |   Tempo  | Tempo/prof |")
    print("---------+" + "-----"*N_PERFIS + "-+-------+---------+----------+------------+")
    for unidade in range(N_UNIDADES):
        print(f"{MATRIZ_UNIDADES[unidade][0]:6s}   | "
              + " ".join([f"{qtdes[unidade][p]:4d}" for p in range(N_PERFIS)])
        #                             Total                      P-Eq                                    Tempo                 - horas de orientação                                     Tempo/prof
        + f" |  {np.sum(qtdes[unidade]):4d} | {np.sum(qtdes[unidade]*MATRIZ_PEQ):7.2f} |  {np.sum(qtdes[unidade]*MATRIZ_TEMPO) - MATRIZ_UNIDADES[unidade][2]:7.2f} |    {(np.sum(qtdes[unidade]*MATRIZ_TEMPO) - MATRIZ_UNIDADES[unidade][2])/np.sum(qtdes[unidade]):7.3f} |")
    print("---------+" + "-----"*N_PERFIS + "-+-------+---------+----------+------------+")
    print("Total    | " + " ".join([f"{np.sum(qtdes, axis=0)[p]:4d}" for p in range(N_PERFIS)])
    #            Total                      P-Eq                                     Tempo      -  horas de orientação                                               Tempo/prof
    + f" |  {np.sum(qtdes):4d} | {np.sum(qtdes*MATRIZ_PEQ):7.2f} |  {np.sum(qtdes*MATRIZ_TEMPO) - np.sum(MATRIZ_UNIDADES[:N_UNIDADES], axis=0)[2]:7.2f} |    {(np.sum(qtdes*MATRIZ_TEMPO) - np.sum(MATRIZ_UNIDADES[:N_UNIDADES], axis=0)[2])/np.sum(qtdes):7.3f} |")
    print("---------+" + "-----"*N_PERFIS + "-+-------+---------+----------+------------+")
    print("")

#-----------------------------------------------------------------------------------------
def imprimir_parametros(qtdes):
    """Imprime os dados de entrada e os resultados obtidos"""
    print("Parâmetros:")
    print("---------+-------------+--------------------+--------------+-----------+-----------+---------+---------+----------+")
    print("Unidade  |      aulas  |     horas_orient   |  num_orient  |   diretor |   coords. |   40h   |   20h   | ch media |")
    print("---------+-------------+--------------------+--------------+-----------+-----------+---------+---------+----------+")
    # Formatos dos números - tem que ser tudo como float, pois ao importar os valores de
    # professor-equivalente, a MATRIZ_PERFIS fica toda como float
    formatos =          ['4.0f', '7.2f', '4.0f', '3.0f', '3.0f']
    formatosdiferenca = ['3.0f', '7.2f', '4.0f', '2.0f', '2.0f']
    for unidade in range(N_UNIDADES):
        print(f"{MATRIZ_UNIDADES[unidade][0]:6s}   | "
        + " ".join([f"{np.sum(qtdes[unidade]*MATRIZ_PERFIS[p]):{formatos[p]}} (+{(np.sum(qtdes[unidade]*MATRIZ_PERFIS[p]) - MATRIZ_UNIDADES[unidade][p+1]):{formatosdiferenca[p]}}) |" for p in range(N_RESTRICOES)])
        + f"  {((qtdes[unidade][4] + qtdes[unidade][5]) / np.sum(qtdes[unidade]))*100:5.2f}% |" #40h
        + f"  {((qtdes[unidade][6] + qtdes[unidade][7]) / np.sum(qtdes[unidade]))*100:5.2f}% |" #20h
        + f"  {MATRIZ_UNIDADES[unidade][1] / np.sum(qtdes[unidade]):7.3f} |"
        )
    print("---------+-------------+--------------------+--------------+-----------+-----------+---------+---------+----------+")
    print("Total    | "
    + " ".join([f"{np.sum(qtdes*MATRIZ_PERFIS[p]):{formatos[p]}} (+{np.sum(qtdes*MATRIZ_PERFIS[p]) - int(np.sum(MATRIZ_UNIDADES, axis=0)[p+1]):{formatosdiferenca[p]}}) |" for p in range(N_RESTRICOES)])
    + f"  {(np.sum(qtdes, axis=0)[4] + np.sum(qtdes, axis=0)[5]) / np.sum(qtdes)*100:5.2f}% |"
    + f"  {(np.sum(qtdes, axis=0)[6] + np.sum(qtdes, axis=0)[7]) / np.sum(qtdes)*100:5.2f}% |"
    + f"  {np.sum(MATRIZ_UNIDADES, axis=0)[1]/np.sum(qtdes):7.3f} |"
    )
    print("---------+-------------+--------------------+--------------+-----------+-----------+---------+---------+----------+")
    print()

### Fim das funçõoes ###

root = tk.Tk()
root.title("Otimizador de distribuição de professores 1.0 - Pedro Santos Guimarães")
#root.geometry("1200x600")

# Título
textoTitulo = tk.Label(root, text="Bem vindo.", anchor="w", justify="left",
                       font=font.Font(weight="bold"))
textoTitulo.grid(sticky='W', row=0, column=0, padx=10, pady=10)

# Texto instruções
#textoBotao = tk.Label(root, text="Primeiro escolha o arquivo no botão abaixo.\n\
#                                  Depois verifique as opções e clique em Executar.")
#textoBotao.grid(row=0, column=3)

# Grupo arquivo
grupo_arq = ttk.LabelFrame(root)
grupo_arq.grid(row=1, column=0, padx=10, pady=10, sticky='w')

# Texto botão arquivo
textoBotao = tk.Label(grupo_arq, text="Escolha o arquivo:")
textoBotao.grid(row=1, column=0)

# Botão para selecionar o arquivo
botaoArquivo = ttk.Button(grupo_arq, text="Abrir arquivo", command=carregar_arquivo)
botaoArquivo.grid(row=1, column=1, padx=10, pady=10)

# Grupo opções
grupo = ttk.LabelFrame(root, text="Opções")
grupo.grid(row=2, column=0, padx=10, pady=10)

# Checkbox para ch minima
bool_minima = tk.BooleanVar(value=True)
checkbox_minima = tk.Checkbutton(grupo, text="CH mínima: ", variable=bool_minima,
                                 command=verifica_check_boxes)
checkbox_minima.grid(row=3, column=0, padx=10, pady=10)

# campo texto
texto_ch_min = tk.IntVar(value=12)
entrada_CH_MIN = tk.Entry(grupo, textvariable=texto_ch_min, width=5)
entrada_CH_MIN.grid(row=3, column=1, padx=10, pady=10)

ToolTip(checkbox_minima, msg="Ativar carga horária média máxima por unidade", delay=0.1)
ToolTip(entrada_CH_MIN, msg="Valor da carga horária máxima", delay=0.1)

# Checkbox para ch maxima
bool_maxima = tk.BooleanVar(value=True)
checkbox_maxima = tk.Checkbutton(grupo, text="CH máxima: ", variable=bool_maxima,
                                 command=verifica_check_boxes)
checkbox_maxima.grid(row=4, column=0, padx=10, pady=10)

# campo texto
texto_ch_max = tk.IntVar(value=16)
entrada_CH_MAX = tk.Entry(grupo, textvariable=texto_ch_max, width=5)
entrada_CH_MAX.grid(row=4, column=1, padx=10, pady=10)

ToolTip(checkbox_maxima, msg="Ativar carga horária média mínima geral (para toda a Universidade)",
        delay=0.1)
ToolTip(entrada_CH_MAX, msg="Valor da carga horária mínima", delay=0.1)

# Checkbox para total minimo
bool_min_total = tk.BooleanVar(value=False)
checkbox_min_total = tk.Checkbutton(grupo, text="Total mínimo: ",
    variable=bool_min_total, command=verifica_check_boxes)
checkbox_min_total.grid(row=3, column=3, padx=10, pady=10)

# campo texto
texto_min_total = tk.IntVar()
entrada_N_MIN_total = tk.Entry(grupo, textvariable=texto_min_total, width=5)
entrada_N_MIN_total.grid(row=3, column=4, padx=10, pady=10)

ToolTip(checkbox_min_total, msg="Ativar número mínimo total de professores", delay=0.1)
ToolTip(entrada_N_MIN_total, msg="Valor do mínimo total", delay=0.1)

# Checkbox para total maxima
bool_max_total = tk.BooleanVar(value=False)
checkbox_max_total = tk.Checkbutton(grupo, text="Total máximo: ",
    variable=bool_max_total, command=verifica_check_boxes)
checkbox_max_total.grid(row=4, column=3, padx=10, pady=10)

# campo texto
texto_max_total = tk.IntVar()
entrada_N_MAX_total = tk.Entry(grupo, textvariable=texto_max_total, width=5)
entrada_N_MAX_total.grid(row=4, column=4, padx=10, pady=10)

ToolTip(checkbox_max_total, msg="Ativar número máximo total de professores", delay=0.1)
ToolTip(entrada_N_MAX_total, msg="Valor do máximo total", delay=0.1)

# Tempo limite
val_limite = tk.IntVar(value=30)
label_limite = tk.Label(grupo, text="Tempo limite:")
label_limite.grid(row=5, column=0, padx=10, pady=10)
entrada_tempo_limite = tk.Entry(grupo, textvariable=val_limite, width=5)
entrada_tempo_limite.grid(row=5, column=1, padx=10, pady=10)

ToolTip(label_limite, msg="Tempo máximo para procurar a solução ótima", delay=0.1)

# Combobox
label1 = ttk.Label(grupo, text="Critério:")
label1.grid(row=6, column=0, padx=10, pady=10)

combo_var = tk.StringVar()
combobox = ttk.Combobox(grupo, textvariable=combo_var, values=list(LISTA_MODOS.keys()),
                        state="readonly")
combobox.grid(row=6, column=1, padx=10, pady=10)
combobox.bind("<<ComboboxSelected>>", lambda event: verifica_executar())
combobox.set("Todos")
combo_var.set("Todos")

# Radiobuttons
label2 = ttk.Label(grupo, text="Escolha uma opção")
label2.grid(row=7, column=0, padx=10, pady=10)

radio_var = tk.StringVar()

radio1 = ttk.Radiobutton(grupo, text="Planejar", variable=radio_var, value="Planejar",
                         command=verifica_executar)
radio1.grid(row=7, column=1, padx=10, pady=10, sticky='w')

radio2 = ttk.Radiobutton(grupo, text="Operar", variable=radio_var, value="Operar",
                         command=verifica_executar)
radio2.grid(row=8, column=1, padx=10, pady=10, sticky='w')

radio3 = ttk.Radiobutton(grupo, text="Fazer", variable=radio_var, value="Fazer",
                         command=verifica_executar)
radio3.grid(row=9, column=1, padx=10, pady=10, sticky='w')

# Botão para executar
botaoExecutar = ttk.Button(grupo, text="Executar", state=tk.DISABLED, command=executar)
botaoExecutar.grid(row=10, column=0, padx=10, pady=10)

# Inicialmente oculta as opções
grupo.grid_forget()

# Dummy para ajustar as colunas
resultado = tk.Label(root)
resultado.grid(row=8, column=0, padx=10, pady=10)

root.mainloop()
