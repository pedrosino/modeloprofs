"""Modelo de programação linear para distribuição de vagas entre unidades acadêmicas
de uma universidade federal. Desenvolvido na pesquisa do Mestrado Profissional em
Gestão Organizacional, da Faculdade de Gestão e Negócios, Universidade Federal de Uberlândia,
por Pedro Santos Guimarães, em 2023"""
#from datetime import datetime
#mport sys
from datetime import datetime
import math
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from tkinter import font
import pandas as pd
import numpy as np
from tktooltip import ToolTip
from ortools.linear_solver import pywraplp

# Variáveis globais
MATRIZ_UNIDADES = None
MATRIZ_PERFIS = None
MATRIZ_PEQ = None
MATRIZ_TEMPO = None
PESOS = None
NOMES_RESTRICOES = None
QTDES_FINAL = None
DATA_FRAME = None
RELATORIO = ""
RESTRICOES_PERCENTUAIS = {}

# Dados do problema
N_PERFIS = None
N_RESTRICOES = None
N_UNIDADES = None
CONECTORES = None

LISTA_MODOS = {
    "Menor número": "num",
    "Menos P-Eq": "peq",
    "Mais tempo": "tempo",
    "Menos tempo": "tempo-reverso",
    "Equilíbrio CH": "ch",
    "Desequilíbrio CH": "ch-reverso",
    "Todos": "todos"
}

# Vetor de modelos
MODELOS = {}

# Vetor de status
STATUS = {}

# Vetor com formato do resultado conforme o modo
FORMATO_RESULTADO = {}
FORMATO_RESULTADO['num'] = 'professores'
FORMATO_RESULTADO['peq'] = 'prof-equivalente'
FORMATO_RESULTADO['tempo'] = 'horas'
FORMATO_RESULTADO['tempo-reverso'] = 'horas'
FORMATO_RESULTADO['ch'] = 'aulas/prof'
FORMATO_RESULTADO['ch-reverso'] = 'aulas/prof'
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
# Número total de aulas a serem ministradas
TOTAL_AULAS = None
# Tempo limite para procurar a solução ótima
TEMPO_LIMITE = 30
MODO_ESCOLHIDO = 'todos'
FATOR = 1000

def verifica_executar():
    """Habilita ou desabilita o botão Executar"""
    if combo_var.get():
        botao_executar['state'] = tk.NORMAL
        botao_executar.configure(bg="#ddd")
    else:
        botao_executar['state'] = tk.DISABLED


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

    var_nome_arquivo.set(arquivo.split("/")[-1])
    root.update()

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
    PESOS = np.delete(pesos_lidos, 0, axis=1).transpose().reshape(4)

    # Se a importação teve sucesso
    if(len(MATRIZ_UNIDADES) and len(MATRIZ_PERFIS)):
        # Mostra as opções
        grupo_opcoes.grid(row=2, column=0, padx=10, pady=10, rowspan=1, sticky='nw')
        verifica_check_boxes()
        atualiza_tela()


def executar():
    """Executa a otimização"""
    global LIMITAR_CH_MAXIMA, LIMITAR_CH_MINIMA, CH_MAX, CH_MIN, MAX_TOTAL, MIN_TOTAL, \
        TEMPO_LIMITE, N_MIN_TOTAL, N_MAX_TOTAL, MODO_ESCOLHIDO, QTDES_FINAL, RELATORIO, \
        DATA_FRAME, TOTAL_AULAS

    # Mostra resultados
    grupo_resultados.grid(row=1, column=1, padx=10, pady=10, rowspan=2, sticky='nw')
    atualiza_tela()

    # Limpa tabela
    text_aba.delete("1.0", tk.END)

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

    TOTAL_AULAS = np.sum(MATRIZ_UNIDADES[:N_UNIDADES], axis=0)[1]

    # Verifica modo escolhido
    if MODO_ESCOLHIDO not in ['num', 'peq', 'tempo', 'tempo-reverso', 'ch', 'ch-reverso', 'todos']:
        MODO_ESCOLHIDO = 'todos'
        print("O modo escolhido era inválido. Será utilizado o modo 'todos'.")

    # Verifica números totais
    if (MIN_TOTAL and N_MIN_TOTAL < 1) or (MAX_TOTAL and N_MAX_TOTAL < 1):
        print("Os números totais especificados são inválidos. Essas opções foram desativadas.")
        MIN_TOTAL = False
        MAX_TOTAL = False

    # Inicia relatório
    RELATORIO = "Relatório da execução do otimizador" \
        + f"\nData: {datetime.now().strftime('%d/%m/%Y')}\n"
    RELATORIO += f'\nModo escolhido: {MODO_ESCOLHIDO}'
    RELATORIO += f'\nUnidades: {N_UNIDADES}'
    RELATORIO += f'\nCH Maxima: {LIMITAR_CH_MAXIMA} {CH_MAX if LIMITAR_CH_MAXIMA else ""}'
    RELATORIO += f'\nCH Minima: {LIMITAR_CH_MINIMA} {CH_MIN if LIMITAR_CH_MINIMA else ""}'
    RELATORIO += f'\nTotal: {N_MIN_TOTAL if MIN_TOTAL else "-"} a ' \
        f'{N_MAX_TOTAL if MAX_TOTAL else "-"}'
    # Se houver restrições em algum perfil, imprime aqui
    if len(RESTRICOES_PERCENTUAIS) > 0:
        for restricao in RESTRICOES_PERCENTUAIS.values():
            # Calcula coeficientes dos perfis
            percentual = restricao['percentual']
            perfis = restricao['perfis']
            sinal = restricao['sinal']
            escopo = restricao['escopo']
            RELATORIO += '\nPerfis (' + ','.join(str(p+1) for p in perfis) \
                + f') {sinal} {percentual*100}% ({escopo})'

    # Imprime unidades
    imprimir_unidades()

    # Imprime perfis
    imprimir_perfis()

    RELATORIO += '\n------------------------------------------------------------\n'

    texto_resultado = f"O modo escolhido foi {MODO_ESCOLHIDO}"
    if MODO_ESCOLHIDO == 'todos':
        texto_resultado += "\nPrimeiramente vamos definir os parâmetros para cada critério/objetivo"

    resultado.set(texto_resultado)
    root.update()

    # Conforme o modo escolhido, faz só uma otimização ou todas
    if MODO_ESCOLHIDO != 'todos':
        texto_resultado = resultado.get()
        texto_resultado += "\nResolvendo ..."
        resultado.set(texto_resultado)
        atualiza_tela()
        resultado_final, QTDES_FINAL = otimizar(MODO_ESCOLHIDO, None, None)
    else:
        # Critérios/modos
        modos = np.array(['num', 'peq', 'tempo', 'tempo-reverso', 'ch', 'ch-reverso'])

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
        # cenário. O melhor caso deve ser determinado resolvendo o modelo usando esse critério,
        # e o pior caso fazendo uma otimização inversa.
        # Para o critério 4, o melhor caso é determinado pela solução inicial do modelo com
        # este critério e o pior caso também com uma otimização inversa.
        numero_max = int(TOTAL_AULAS / CH_MIN) if not MAX_TOTAL else N_MAX_TOTAL
        piores['num'] = numero_max
        piores['peq'] = round(numero_max*1.65, 2)

        # Percorre os critérios
        for modo_usado in modos:
            texto_resultado = resultado.get()
            texto_resultado += f"\nModo {modo_usado}: resolvendo ..."
            resultado.set(texto_resultado)
            atualiza_tela()

            # Obtém resultado e quantidades
            resultado_modo, qtdes_modo = otimizar(modo_usado, piores, melhores)

            # Registra o resultado na lista de melhores casos
            if 'reverso' not in modo_usado:
                melhores[modo_usado] = resultado_modo
            # ou na de piores casos
            else:
                piores[modo_usado.split('-')[0]] = resultado_modo

            texto_resultado = resultado.get()
            if 'reverso' not in modo_usado:
                texto_resultado += f"\nTerminado. Resultado: {melhores[modo_usado]}, "\
                    f"resolvido em {MODELOS[modo_usado].wall_time()/1000} segundos"
            else:
                texto_resultado += f"\nTerminado. Resultado: {piores[modo_usado.split('-')[0]]}, "\
                    f"resolvido em {MODELOS[modo_usado].wall_time()/1000} segundos"
            resultado.set(texto_resultado)
            atualiza_tela()

            # Imprime resultados
            imprimir_resultados(qtdes_modo)
            # Imprime parâmetros
            imprimir_parametros(qtdes_modo)

            # Para os modos de carga horária, imprime as médias e desvios
            if 'ch' in modo_usado:
                # Imprime médias e desvios
                for variable in MODELOS[modo_usado].variables():
                    nomes_busca = ['modulo', 'media']
                    if any(nome in variable.name() for nome in nomes_busca):
                        RELATORIO += f"\n{variable.name()}: {variable.solution_value():.4f}"

            RELATORIO += "\n------------------------------------------------------------\n"

        ##### -----------------Fim da primeira 'rodada'-----------------
        RELATORIO += f"\nMelhores: {melhores}\nPiores: {piores}\n"
        RELATORIO += "------------------------------------------------------------"

        ## Agora uma nova rodada do modelo usando os pesos e as listas de melhores e piores casos
        # Atualiza o texto do resultado
        texto_resultado = resultado.get()
        texto_resultado += "\n\nModo 'todos': resolvendo ..."
        resultado.set(texto_resultado)
        atualiza_tela()

        # Otimização com o modo 'todos'
        resultado_final, QTDES_FINAL = otimizar('todos', piores, melhores)

        texto_resultado = resultado.get()
        texto_resultado += f"\nTerminado. Resultado: {resultado_final}, "\
            f"resolvido em {MODELOS['todos'].wall_time()/1000} segundos"
        resultado.set(texto_resultado)
        atualiza_tela()

    # ---- Finalização ----
    # Imprime resultados
    imprimir_resultados(QTDES_FINAL)
    # Imprime parâmetros
    imprimir_parametros(QTDES_FINAL)

    if 'ch' in MODO_ESCOLHIDO:
        # Imprime médias e desvios
        for variable in MODELOS[MODO_ESCOLHIDO].variables():
            nomes_busca = ['modulo', 'media']
            if any(nome in variable.name() for nome in nomes_busca):
                RELATORIO += f"\n{variable.name()}: {variable.solution_value():.4f}"

    if MODO_ESCOLHIDO == 'todos':
        # PESOS
        RELATORIO += f"\nPESOS: {PESOS}"
        # Imprime médias e desvios
        for variable in MODELOS['todos'].variables():
            nomes_busca = ['p_', 'modulo', 'media']
            if any(nome in variable.name() for nome in nomes_busca):
                RELATORIO += f"\n{variable.name()}: {variable.solution_value():.4f}"

    # Imprime o modelo completo
    RELATORIO += "\n\n------------------Modelo:------------------\n"
    ## https://stackoverflow.com/a/1140967/3059369
    #modelo = f"\n{MODELOS[MODO_ESCOLHIDO]}"
    RELATORIO += MODELOS[MODO_ESCOLHIDO].ExportModelAsLpFormat(False).replace('\\', '').replace(',_', ',')

    texto_resultado = resultado.get()
    texto_resultado += f"\nSituação: {STATUS[MODO_ESCOLHIDO]}"
    texto_resultado += f"\nObjetivo: {resultado_final} {FORMATO_RESULTADO[MODO_ESCOLHIDO]}"
    texto_resultado += f"\nResolvido em {MODELOS[MODO_ESCOLHIDO].wall_time()/1000} segundos"

    resultado.set(texto_resultado)
    atualiza_tela()

    # Transforma em dataframe com cabeçalho e unidades
    DATA_FRAME = pd.DataFrame(QTDES_FINAL, columns=[f'x{i}' for i in range(1, N_PERFIS+1)])
    # Manter os tipos
    DATA_FRAME = DATA_FRAME.convert_dtypes()
    # Linha com totais
    DATA_FRAME.loc['Total', :] = DATA_FRAME.sum().values
    # Coluna com total
    DATA_FRAME['Total'] = DATA_FRAME.sum(axis=1, numeric_only=True)
    # Insere coluna
    DATA_FRAME.insert(0, "Unidade", np.append(MATRIZ_UNIDADES[:, 0],['Total Perfil']))

    # Mostra na tabela
    text_aba.insert(tk.END, DATA_FRAME.to_string(index=False))

    # Altura máxima
    altura_maxima = 18

    # Obtém o número total de linhas do texto
    num_linhas = int(text_aba.index(tk.END).split('.', maxsplit=1)[0])

    # Ajusta a altura do widget para mostrar no máximo altura_maxima linhas
    if num_linhas <= altura_maxima:
        text_aba.config(height=num_linhas, width=12 + N_PERFIS*5)
    else:
        text_aba.config(height=altura_maxima, width=12 + N_PERFIS*5)
        scrollbar.grid(row=2, column=1, sticky="ns")

    atualiza_tela()

    altura_janela = root.winfo_height()
    altura_tela = root.winfo_screenheight()
    # Verifica o tamanho da janela
    if altura_tela - altura_janela < 150:
        barra.grid(row=0, column=1, sticky="ns")
        root.geometry(f"{root.winfo_width()}x{altura_tela - 150}+{root.winfo_rootx()}+10")
        atualiza_tela()


def otimizar(modo, piores, melhores):
    """Função que faz a otimização conforme o modo escolhido"""
    global RELATORIO

    # -- Definir o modelo --
    MODELOS[modo] = pywraplp.Solver(f'ProfsORTools-{modo}',
        pywraplp.Solver.SAT_INTEGER_PROGRAMMING)
    infinito = pywraplp.Solver.infinity()

    #nomes
    nomes = [str(perfil) + "_"
        + MATRIZ_UNIDADES[und][0] for und in range(N_UNIDADES) for perfil in range(1, N_PERFIS+1)]

    # Variáveis de decisão
    #var_x = LpVariable.matrix("x", nomes, cat="Integer", lowBound=0)
    var_x = [MODELOS[modo].IntVar(0, infinito,
        f'x_{nome}') for nome in nomes]
    saida = np.array(var_x).reshape(N_UNIDADES, N_PERFIS)

    minima = LIMITAR_CH_MINIMA
    maxima = LIMITAR_CH_MAXIMA
    # Ativa carga horária mínima e máxima, para que todos os modos tenham as mesmas restrições
    if MODO_ESCOLHIDO == 'todos' or 'ch' in modo:
        minima = True
        maxima = True
    # No modo tempo é necessário estabelecer a carga horária mínima (ou número máximo)
    if 'tempo' in modo:
        minima = True

    # -- Restrições --
    for restricao in range(N_RESTRICOES):
        for unidade in range(N_UNIDADES):
            match CONECTORES[restricao]:
                case ">=":
                    MODELOS[modo].Add(saida[unidade].dot(MATRIZ_PERFIS[restricao]) \
                        >= MATRIZ_UNIDADES[unidade][restricao+1], \
                        NOMES_RESTRICOES[restricao] + " " + MATRIZ_UNIDADES[unidade][0])
                case "==":
                    MODELOS[modo].Add(saida[unidade].dot(MATRIZ_PERFIS[restricao]) \
                        == MATRIZ_UNIDADES[unidade][restricao+1], \
                        NOMES_RESTRICOES[restricao] + " " + MATRIZ_UNIDADES[unidade][0])
                case "<=":
                    MODELOS[modo].Add(saida[unidade].dot(MATRIZ_PERFIS[restricao]) \
                        <= MATRIZ_UNIDADES[unidade][restricao+1], \
                        NOMES_RESTRICOES[restricao] + " " + MATRIZ_UNIDADES[unidade][0])

    # Restrições de percentual por perfil
    if len(RESTRICOES_PERCENTUAIS) > 0:
        for restricao in RESTRICOES_PERCENTUAIS.values():
            # Calcula coeficientes dos perfis
            percentual = restricao['percentual']
            perfis = restricao['perfis']
            sinal = restricao['sinal']
            escopo = restricao['escopo']
            coeficientes = [1 - percentual if p in perfis else percentual*-1 for p in range(N_PERFIS)]

            # Nome da restrição
            nome_restricao = "Perfis (" + ",".join(str(p) for p in perfis) \
                + f") {sinal} {percentual*100}% "

            # Se a restrição for só nas unidades, automaticamente valerá para o total
            if escopo == 'unidades':
                for unidade in range(N_UNIDADES):
                    match sinal:
                        # Monta a restrição conforme o sinal escolhido
                        case '<=':
                            MODELOS[modo].Add(sum(saida[unidade]*coeficientes) <= 0, \
                                nome_restricao + MATRIZ_UNIDADES[unidade][0])
                        case '>=':
                            MODELOS[modo].Add(sum(saida[unidade]*coeficientes) >= 0, \
                                nome_restricao + MATRIZ_UNIDADES[unidade][0])
            else:
                # Restrição no total
                match sinal:
                    case '<=':
                        MODELOS[modo].Add(sum(saida*coeficientes) <= 0, \
                            f'{nome_restricao} geral')
                    case '>=':
                        MODELOS[modo].Add(sum(saida*coeficientes) >= 0, \
                            f'{nome_restricao} geral')

    # Restrições de carga horária média:
    # ------------------
    # Exemplo para 900 aulas, considerando carga horária média mínima de 12 aulas e máxima de 16
    # soma <= 900/12 -> soma <= 75
    # soma >= 900/16 -> soma >= 56.25
    # a soma deve estar entre 57 e 75, inclusive
    if maxima:
        # Restrição no total
        MODELOS[modo].Add(CH_MAX*sum(var_x) >= TOTAL_AULAS, \
            f"chmax: {math.ceil(TOTAL_AULAS/CH_MAX)}")
        # E por unidade
        for unidade in range(N_UNIDADES):
            n_min = math.ceil(MATRIZ_UNIDADES[unidade][1]/CH_MAX)
            MODELOS[modo].Add(CH_MAX*sum(saida[unidade]) >= MATRIZ_UNIDADES[unidade][1], \
                f"{MATRIZ_UNIDADES[unidade][0]}_chmax: {n_min}")

    if minima:
        # Restrição mínima somente no geral -> permite exceções como o IBTEC em Monte Carmelo,
        # que tem 19 aulas apenas
        MODELOS[modo].Add(CH_MIN*sum(var_x) <= TOTAL_AULAS, \
            f"chmin: {int(TOTAL_AULAS/CH_MIN)}")
        # No modo tempo-reverso e ch-reverso é preciso estabelecer um mínimo coerente por unidade
        # Foi adotado 9 aulas por professor (vide acima).
        if 'reverso' in modo:
            minimo_local = 9
            for unidade in range(N_UNIDADES):
                n_max = int(MATRIZ_UNIDADES[unidade][1]/minimo_local)
                MODELOS[modo].Add(minimo_local*sum(saida[unidade]) <= MATRIZ_UNIDADES[unidade][1], \
                    f"{MATRIZ_UNIDADES[unidade][0]}_chmin: {n_max}")

    # Restrição do número total de professores
    if MAX_TOTAL:
        MODELOS[modo].Add(sum(var_x) <= N_MAX_TOTAL, f"TotalMax: {N_MAX_TOTAL}")
    if MIN_TOTAL:
        MODELOS[modo].Add(sum(var_x) >= N_MIN_TOTAL, f"TotalMin: {N_MIN_TOTAL}")

    # Modo de equilíbrio da carga horária
    if modo in ['ch', 'ch-reverso', 'todos']:
        # variáveis auxiliares
        # Para cada unidade há um valor da média e um desvio em relação à média geral
        modulos = [MODELOS[modo].IntVar(0, infinito, f"modulo_{MATRIZ_UNIDADES[u][0]}") for u in range(N_UNIDADES)]
        #consts = LpVariable.matrix("b", MATRIZ_UNIDADES[:N_UNIDADES, 0], cat="Binary")
        consts = [MODELOS[modo].IntVar(0, 1, f"b_{MATRIZ_UNIDADES[u][0]}") for u in range(N_UNIDADES)]
        #medias = LpVariable.matrix("media", MATRIZ_UNIDADES[:N_UNIDADES, 0], \
        #                           cat="Continuous", lowBound=0)
        medias = [MODELOS[modo].IntVar(0, infinito, f"media_{MATRIZ_UNIDADES[u][0]}") for u in range(N_UNIDADES)]
        #media_geral = LpVariable("media_geral", cat="Continuous", lowBound=0)
        media_geral = MODELOS[modo].IntVar(0, infinito, "media_geral")

        # Como a média seria uma função não linear (aulas/professores), foi feita uma
        # aproximação com uma reta que passa pelos dois pontos extremos dados pelos
        # valores de carga horária mínima e máxima.
        # Coeficientes
        # coef = (y1-y0)/(x1-x0) -> y1 = media minima, y0 = media maxima,
        # x1 = numero maximo, x0 = numero minimo
        x_1 = TOTAL_AULAS / CH_MIN
        x_0 = TOTAL_AULAS / CH_MAX
        coef = round(FATOR*(CH_MIN - CH_MAX) / (x_1 - x_0))

        # O cálculo da média é inserido no modelo como uma restrição
        MODELOS[modo].Add(media_geral == coef*sum(var_x) + FATOR*(CH_MIN + CH_MAX), "Media geral")

        # O mesmo raciocínio é feito para cada unidade
        for unidade in range(N_UNIDADES):
            x1u = MATRIZ_UNIDADES[unidade][1] / CH_MIN
            x0u = MATRIZ_UNIDADES[unidade][1] / CH_MAX
            coef_u = round(FATOR*(CH_MIN - CH_MAX) / (x1u - x0u))
            # Cálculo da media
            MODELOS[modo].Add(medias[unidade] == coef_u*sum(saida[unidade]) + FATOR*(CH_MIN + CH_MAX), \
                f"{MATRIZ_UNIDADES[unidade][0]}_ch_media")
            # Cálculo do desvio
            # O desvio é dado pelo módulo da subtração, porém o PuLP não aceita a função abs()
            # Assim, são colocadas duas restrições, uma usando o valor positivo e outra o negativo
            ## https://optimization.cbe.cornell.edu/index.php?title=Optimization_with_absolute_values
            ## https://lpsolve.sourceforge.net/5.5/absolute.htm
            m_grande = 1000000
            # desvio
            # X + M * B >= x'
            MODELOS[modo].Add(medias[unidade] - media_geral\
                + m_grande*consts[unidade] >= modulos[unidade], \
                f"X + M * B >= x' {MATRIZ_UNIDADES[unidade][0]}")
            # -X + M * (1-B) >= x'
            MODELOS[modo].Add(-1*(medias[unidade] - media_geral) \
                + m_grande*(1-consts[unidade]) >= modulos[unidade], \
                f"-X + M * (1-B) >= x' {MATRIZ_UNIDADES[unidade][0]}")
            # modulo
            MODELOS[modo].Add(modulos[unidade] >= (medias[unidade] - media_geral), \
                f"{MATRIZ_UNIDADES[unidade][0]}_mod1")
            MODELOS[modo].Add(modulos[unidade] >= -1*(medias[unidade] - media_geral), \
                f"{MATRIZ_UNIDADES[unidade][0]}_mod2")

    # Modo com todos os critérios
    if modo == 'todos':
        # Variáveis com as pontuações
        pontuacoes = [MODELOS[modo].IntVar(0, infinito, f"p_{i}") for i in range(4)]
        # Restrições/cálculos
        # Caso seja dado um número exato, é necessário alterar a pontuação do critério 'num'
        # para evitar a divisão por zero

        # Calcula os coeficientes
        # Necessariamente inteiros!
        coeficientes = {}
        constantes = {}
        outro_fator = FATOR*100
        for m in ['num', 'peq', 'tempo', 'ch']:
            coeficientes[m] = round(outro_fator/(melhores[m] - piores[m]))
            constantes[m] = round(outro_fator*piores[m]/(melhores[m] - piores[m]))

        print(coeficientes)
        print(constantes)

        if MAX_TOTAL and MIN_TOTAL and N_MAX_TOTAL == N_MIN_TOTAL:
            coef = round(outro_fator/melhores['num'])
            MODELOS[modo].Add(pontuacoes[0] == coef*sum(var_x), "Pontuação número")
        else:
            #MODELOS[modo].Add(pontuacoes[0] == coeficientes['num']*(sum(var_x) - piores['num']), "Pontuação número")
            MODELOS[modo].Add(pontuacoes[0] == coeficientes['num']*sum(var_x) - constantes['num'], "Pontuação número")
        #MODELOS[modo].Add(pontuacoes[1] == coeficientes['peq']*(sum(saida.dot(MATRIZ_PEQ)) - piores['peq']), "Pontuação P-Eq")
        MODELOS[modo].Add(pontuacoes[1] == coeficientes['peq']*sum(saida.dot(MATRIZ_PEQ)) - constantes['peq'], "Pontuação P-Eq")
        #MODELOS[modo].Add(pontuacoes[2] == coeficientes['tempo']*(sum(saida.dot(MATRIZ_TEMPO)) - np.sum(MATRIZ_UNIDADES[:N_UNIDADES], axis=0)[2] - piores['tempo']), "Pontuação tempo")
        MODELOS[modo].Add(pontuacoes[2] == coeficientes['tempo']*sum(saida.dot(MATRIZ_TEMPO)) - round(coeficientes['tempo']*np.sum(MATRIZ_UNIDADES[:N_UNIDADES], axis=0)[2]) - constantes['tempo'], "Pontuação tempo")
        #MODELOS[modo].Add(pontuacoes[3] == coeficientes['ch']*(sum(modulos[:N_UNIDADES])/N_UNIDADES - piores['ch']), "Pontuação Equilíbrio")
        MODELOS[modo].Add(pontuacoes[3] == coeficientes['ch']*(sum(modulos[:N_UNIDADES])/N_UNIDADES) - constantes['ch'], "Pontuação Equilíbrio")

    # -- Função objetivo --
    if modo == 'num':
        MODELOS[modo].Minimize(sum(var_x))
    elif modo == 'peq':
        MODELOS[modo].Minimize(sum(saida.dot(MATRIZ_PEQ)))
    # No modo tempo é necessário deduzir do tempo calculado as horas
    # que serão destinadas às orientações
    elif modo == 'tempo':
        MODELOS[modo].Maximize(sum(saida.dot(MATRIZ_TEMPO)) - np.sum(MATRIZ_UNIDADES[:N_UNIDADES], axis=0)[2])
    elif modo == 'tempo-reverso':
        MODELOS[modo].Minimize(sum(saida.dot(MATRIZ_TEMPO)) - np.sum(MATRIZ_UNIDADES[:N_UNIDADES], axis=0)[2])
    # No modo de equilíbrio a medida de desempenho é o desvio médio
    elif modo == 'ch':
        MODELOS[modo].Minimize(sum(modulos[:N_UNIDADES])/N_UNIDADES)
    elif modo == 'ch-reverso':
        MODELOS[modo].Maximize(sum(modulos[:N_UNIDADES])/N_UNIDADES)
    elif modo == 'todos':
        # O objetivo é maximimizar a pontuação dada pela soma de cada
        # pontuação multiplicada por seu peso
        MODELOS[modo].Maximize(sum(pontuacoes*PESOS))

    # Imprime os parâmetros do modo
    RELATORIO += f'\nModo: {modo}'
    RELATORIO += f'\nCH Maxima: {maxima} {CH_MAX if maxima else ""}'
    RELATORIO += f'\nCH Minima: {minima} {CH_MIN if minima else ""}'

    # Ajusta limite
    # Para os modos "intermediários" não é necessário um limite maior que 30 segundos
    if MODO_ESCOLHIDO == 'todos' and modo != 'todos' and TEMPO_LIMITE > 30:
        novo_limite = 30
    else:
        novo_limite = TEMPO_LIMITE

    # Resolver o modelo
    if modo == 'todos':
        MODELOS[modo].EnableOutput()
    MODELOS[modo].set_time_limit(novo_limite*1000)
    STATUS[modo] = MODELOS[modo].Solve()

    # Resultados
    RELATORIO += f"\nSituação: {STATUS[modo]}"
    # Para cada critério o resultado é em um formato diferente
    if 'ch' in modo:
        objetivo = round(MODELOS[modo].Objective().Value(), 4)
        print(f"Media geral: {media_geral.solution_value()}")
        for u in range(N_UNIDADES):
            print(MATRIZ_UNIDADES[u][0])
            print(f"Média: {medias[u].solution_value()}")
            print(f"Desvio: {modulos[u].solution_value()}")
    elif modo == 'peq':
        objetivo = round(MODELOS[modo].Objective().Value(), 2)
    elif modo == 'num':
        objetivo = int(MODELOS[modo].Objective().Value())
    elif 'tempo' in modo:
        objetivo = MODELOS[modo].Objective().Value()
    elif modo == 'todos':
        objetivo = round(MODELOS[modo].Objective().Value(), 4)

    RELATORIO += f"\nObjetivo: {objetivo} {FORMATO_RESULTADO[modo]}"
    RELATORIO += f"\nResolvido em {MODELOS[modo].wall_time()/1000} segundos\n"

    # Extrai as quantidades
    qtdes_saida = np.full((N_UNIDADES, N_PERFIS), 0, dtype=int)
    for u in range(N_UNIDADES):
        for p in range(N_PERFIS):
            qtdes_saida[u][p] = int(saida[u][p].solution_value())

    # Retorna o valor da função objetivo e as quantidades
    return objetivo, qtdes_saida


def imprimir_resultados(qtdes):
    """Imprime resultados da quantidade de cada perfil em cada unidade"""
    global RELATORIO

    RELATORIO += "\nResultados:"
    borda_cabecalho = "\n---------+" + "-----"*N_PERFIS \
        + "-+-------+---------+----------+------------+"
    # Cabeçalho
    RELATORIO += borda_cabecalho
    RELATORIO += "\nUnidade  |  " \
        + "  ".join([f"{i: >3}" for i in [f"x{p+1}" for p in range(N_PERFIS)]]) \
        + " | Total |   P-Eq  |   Tempo  | Tempo/prof |"
    RELATORIO += borda_cabecalho
    # Uma linha por unidade
    for unidade in range(N_UNIDADES):
        total = np.sum(qtdes[unidade])
        peq = np.sum(qtdes[unidade]*MATRIZ_PEQ)
        tempo = np.sum(qtdes[unidade]*MATRIZ_TEMPO) - MATRIZ_UNIDADES[unidade][2]
        RELATORIO += f"\n{MATRIZ_UNIDADES[unidade][0]:6s}   | " \
            + " ".join([f"{qtdes[unidade][p]:4d}" for p in range(N_PERFIS)]) \
            + f" |  {total:4d} | {peq:7.2f} |  {tempo:7.2f} |    {(tempo)/total:7.3f} |"

    # Totais
    total = np.sum(qtdes)
    peq = np.sum(qtdes*MATRIZ_PEQ)
    tempo = np.sum(qtdes*MATRIZ_TEMPO) - np.sum(MATRIZ_UNIDADES[:N_UNIDADES], axis=0)[2]

    RELATORIO += borda_cabecalho
    RELATORIO += "\nTotal    | " \
        + " ".join([f"{np.sum(qtdes, axis=0)[p]:4d}" for p in range(N_PERFIS)])
    RELATORIO += f" |  {total:4d} | {peq:7.2f} |  {tempo:7.2f} |    {(tempo)/total:7.3f} |"
    RELATORIO += borda_cabecalho + "\n"


def imprimir_parametros(qtdes):
    """Imprime os dados de entrada e os resultados obtidos"""
    global RELATORIO

    RELATORIO += "\nParâmetros:"
    borda_cabecalho = "\n---------+-------------+--------------------+--------------+-----------+-----------+----------+"
    linha_cabecalho = "\nUnidade  |      aulas  |     horas_orient   |  num_orient  |   diretor |   coords. | ch media |"
    # Se houver restrições em algum perfil, acrescenta colunas
    if len(RESTRICOES_PERCENTUAIS) > 0:
        for restricao in RESTRICOES_PERCENTUAIS.values():
            # Acrescenta ao cabeçalho
            perfis = restricao['perfis']
            texto = 'P ' + ','.join(str(p+1) for p in perfis)
            texto = texto.center(11) + '|'
            borda_cabecalho += "-----------+"
            linha_cabecalho += texto

    RELATORIO += borda_cabecalho
    RELATORIO += linha_cabecalho
    RELATORIO += borda_cabecalho
    # Formatos dos números - tem que ser tudo como float, pois ao importar os valores de
    # professor-equivalente, a MATRIZ_PERFIS fica toda como float
    formatos =          ['4.0f', '7.2f', '4.0f', '3.0f', '3.0f']
    formatosdiferenca = ['3.0f', '7.2f', '4.0f', '2.0f', '2.0f']

    # Uma linha por unidade
    for unidade in range(N_UNIDADES):
        total_unidade = np.sum(qtdes[unidade])
        valores_perfis = [np.sum(qtdes[unidade]*MATRIZ_PERFIS[p]) for p in range(N_RESTRICOES)]
        diferencas = [valores_perfis[p] - MATRIZ_UNIDADES[unidade][p+1] for p in range(N_RESTRICOES)]
        strings_perfis = [f"{valores_perfis[p]:{formatos[p]}} " \
                          f"(+{diferencas[p]:{formatosdiferenca[p]}}) |" for p in range(N_RESTRICOES)]
        string_final = " ".join(strings_perfis)

        RELATORIO += f"\n{MATRIZ_UNIDADES[unidade][0]:6s}   | " + string_final \
            + f"  {MATRIZ_UNIDADES[unidade][1] / total_unidade:7.3f} |"
        # Se houver restrições em algum perfil, acrescenta colunas
        if len(RESTRICOES_PERCENTUAIS) > 0:
            for restricao in RESTRICOES_PERCENTUAIS.values():
                perfis = restricao['perfis']
                # Calcula percentuais
                quantidade = sum(qtdes[unidade][perfil] for perfil in perfis)
                perc_unidade = quantidade / total_unidade * 100
                RELATORIO += f"   {perc_unidade:6.2f}% |"

    # Totais
    total = np.sum(qtdes)
    valores_perfis = [np.sum(qtdes*MATRIZ_PERFIS[p]) for p in range(N_RESTRICOES)]
    diferencas = [valores_perfis[p] - int(np.sum(MATRIZ_UNIDADES, axis=0)[p+1]) for p in range(N_RESTRICOES)]
    strings_perfis = [f"{valores_perfis[p]:{formatos[p]}} " \
                      f"(+{diferencas[p]:{formatosdiferenca[p]}}) |" for p in range(N_RESTRICOES)]
    string_final = " ".join(strings_perfis)

    RELATORIO += borda_cabecalho
    RELATORIO += "\nTotal    | " + string_final \
        + f"  {np.sum(MATRIZ_UNIDADES, axis=0)[1]/total:7.3f} |"
    # Se houver restrições em algum perfil, acrescenta colunas
    if len(RESTRICOES_PERCENTUAIS) > 0:
        for restricao in RESTRICOES_PERCENTUAIS.values():
            perfis = restricao['perfis']
            # Quantidade total
            quantidade_total = np.sum(qtdes[:, perfis])
            perc_total = quantidade_total / total * 100
            RELATORIO += f"   {perc_total:6.2f}% |"
    RELATORIO += borda_cabecalho + "\n"


def imprimir_unidades():
    """Imprime os dados de entrada das unidades"""
    global RELATORIO

    RELATORIO += "\n\nUnidades:"
    borda_cabecalho = "\n---------+-------+--------------+------------+---------+---------+"
    linha_cabecalho = "\nUnidade  | aulas | horas_orient | num_orient | diretor | coords. |"

    RELATORIO += borda_cabecalho
    RELATORIO += linha_cabecalho
    RELATORIO += borda_cabecalho
    # Formatos dos números - tem que ser tudo como float, pois ao importar os valores de
    # professor-equivalente, a MATRIZ_PERFIS fica toda como float
    formatos =          ['5.0f', '12.2f', '10.0f', '7.0f', '7.0f']

    # Uma linha por unidade
    for unidade in range(N_UNIDADES):
        valores_restricoes = [MATRIZ_UNIDADES[unidade][p+1] for p in range(N_RESTRICOES)]
        strings_restricoes = [f"{valores_restricoes[p]:{formatos[p]}} |" for p in range(N_RESTRICOES)]
        string_final = " ".join(strings_restricoes)

        RELATORIO += f"\n{MATRIZ_UNIDADES[unidade][0]:6s}   | " + string_final

    # Totais
    valores_perfis = [np.sum(MATRIZ_UNIDADES, axis=0)[p+1] for p in range(N_RESTRICOES)]
    strings_perfis = [f"{valores_perfis[p]:{formatos[p]}} |" for p in range(N_RESTRICOES)]
    string_final = " ".join(strings_perfis)

    RELATORIO += borda_cabecalho
    RELATORIO += "\nTotal    | " + string_final
    RELATORIO += borda_cabecalho + "\n"


def imprimir_perfis():
    """Imprime os perfis utilizados"""
    global RELATORIO

    RELATORIO += '\nPerfis:'
    borda_cabecalho = '\n---------------+' + '------+'*N_PERFIS
    linha_cabecalho = '\nCaracterística |' \
        + "".join([f"  {i: >3} |" for i in [f"p{p+1}" for p in range(N_PERFIS)]])
        #+ ''.join([f'  P{p+1}  |' for p in range(N_PERFIS)])

    RELATORIO += borda_cabecalho
    RELATORIO += linha_cabecalho
    RELATORIO += borda_cabecalho

    for restricao in range(N_RESTRICOES):
        RELATORIO += f'\n{NOMES_RESTRICOES[restricao]:14s} | ' \
            + ' '.join([f'{MATRIZ_PERFIS[restricao][p]:4.0f} |' for p in range(N_PERFIS)])

    RELATORIO += '\nOutras ativ.   | ' \
        + ' '.join([f'{MATRIZ_PERFIS[N_RESTRICOES][p]:4.0f} |' for p in range(N_PERFIS)])

    RELATORIO += '\nProf-equiv.    | ' \
        + ' '.join([f'{MATRIZ_PERFIS[N_RESTRICOES+1][p]:4.2f} |' for p in range(N_PERFIS)])

    RELATORIO += borda_cabecalho + "\n"

def exportar_txt():
    """Função para exportar os resultados em formato .txt"""
    nome_arquivo = filedialog.asksaveasfilename(defaultextension=".txt",
        initialfile="Relatório.txt", filetypes=[("Arquivos de Texto", "*.txt")])
    if nome_arquivo:
        with open(nome_arquivo, "w", encoding='UTF-8') as arquivo:
            arquivo.write(RELATORIO)


def exportar_planilha():
    """Função para exportar as quantidades finais em uma planilha"""
    nome_arquivo = filedialog.asksaveasfilename(defaultextension=".xlsx",
        initialfile="Distribuição.xlsx", filetypes=[("Arquivos do Excel", "*.xlsx")])
    if nome_arquivo:
        # Salva na planilha
        DATA_FRAME.to_excel(nome_arquivo, sheet_name='Resultados', index=False, engine="openpyxl")


def excluir_restricao(nome, frame_excluir):
    """Exclui a restição percentual de perfil(s)"""
    RESTRICOES_PERCENTUAIS.pop(nome)
    frame_excluir.destroy()


def formata_erro(label):
    """Formata o label do erro"""
    label.config(bg='#f0a869', fg='#87190b')


def clique_ok(variaveis, janela, var_erro, label_erro):
    """Verifica e captura os valores selecionados"""
    escolhidos = [i for i, opcao in enumerate(variaveis['opcoes']) if opcao.get() == 1]

    # Verifica as opções
    if len(escolhidos) < 1:
        var_erro.set('Escolha pelo menos um perfil')
        formata_erro(label_erro)
        return

    sinal = variaveis['sinal'].get()
    if sinal not in ['<=', '>=']:
        var_erro.set('Sinal inválido!')
        formata_erro(label_erro)
        return

    percentual = variaveis['percentual'].get()
    if percentual <= 0 or percentual > 100:
        var_erro.set('Percentual deve ser maior que 0 e menor ou igual a 100.')
        formata_erro(label_erro)
        return

    escopo = variaveis['escopo'].get()
    if escopo not in ['total', 'unidades']:
        var_erro.set('Opção inválida')
        formata_erro(label_erro)
        return

    # Texto para o label
    texto_perfis = '. Perfis (' + ','.join(str(e+1) for e in escolhidos) \
        + f') {sinal} {percentual}% ({escopo})'
    nome_restricao = 'p' + ','.join(str(e+1) for e in escolhidos) + f'{sinal}{percentual}'

    # Cria um novo frame para a linha de labels
    frame_labels = tk.Frame(frame_perfis)
    frame_labels.pack(anchor='w')
    label_perfil = tk.Label(frame_labels, text=texto_perfis)
    label_perfil.pack(side=tk.LEFT)

    label_excluir = tk.Label(frame_labels, text='[X]', fg="#a00", cursor="hand2")
    label_excluir.pack(side=tk.LEFT)

    label_excluir.bind("<Button-1>", lambda e:excluir_restricao(nome_restricao, frame_labels))

    # Acrescenta à lista de restrições
    nova_restricao = {
        'perfis' : escolhidos,
        'sinal' : sinal,
        'percentual' : percentual/100,
        'escopo': escopo
    }
    RESTRICOES_PERCENTUAIS[nome_restricao] = nova_restricao

    # Fecha janela
    janela.destroy()


def janela_perfis():
    """Abre uma janela para o usuário selecionar os perfis que deseja limitar"""
    #global janela_nova#, opcoes, var_sinal, var_percentual
    janela_nova = tk.Toplevel(root)
    janela_nova.geometry("+320+220")

    tk.Label(janela_nova, text="Selecione os perfis:").grid(row=0, column=0)

    # Lista de perfis
    grupo_perfis = ttk.LabelFrame(janela_nova)
    grupo_perfis.grid(row=1, column=0, rowspan=2)
    lista_opcoes = []
    for perfil in range(8):
        texto = f"Perfil {perfil+1}"
        var = tk.IntVar()
        check_button = tk.Checkbutton(grupo_perfis, text=texto, variable=var)
        check_button.pack(anchor='w')
        lista_opcoes.append(var)

    # Sinal da operação
    grupo_sinal = ttk.LabelFrame(janela_nova)
    grupo_sinal.grid(row=1, column=1, columnspan=2, sticky='s')
    label_sinal = tk.Label(grupo_sinal, text="A soma das quantidades desses perfis deverá ser")
    label_sinal.grid(row=0, column=0, columnspan=3)
    var_sinal = tk.StringVar()
    combo_sinal = ttk.Combobox(grupo_sinal, textvariable=var_sinal,
                               values=['<=', '>='], state="readonly")
    combo_sinal.grid(row=1, column=0, sticky='w')

    # Valor do percentual
    var_percentual = tk.DoubleVar()
    texto_percentual = tk.Entry(grupo_sinal, textvariable=var_percentual, width=5)
    texto_percentual.grid(row=1, column=1, sticky='e')
    tk.Label(grupo_sinal, text="%").grid(row=1, column=2, sticky='w')
    ToolTip(texto_percentual, msg="Para valores não inteiros, use ponto decimal", delay=0.1)

    # Radiobutton para escolher se a restrição é só geral ou também em cada unidade
    frame_radio = tk.Frame(janela_nova)
    frame_radio.grid(row=2, column=1, padx=10, pady=10, columnspan=2)
    tk.Label(frame_radio, text="Essa restrição se aplica").grid(row=0, column=0, sticky='w')
    label_duvida = tk.Label(frame_radio, text=" (?)")
    label_duvida.grid(row=0, column=1, sticky='w')
    ToolTip(label_duvida,
            msg="Escolhendo a opção 'Somente no total' o percentual em cada unidade poderá extrapolar a restrição",
            delay=0.1)
    # Bug: https://stackoverflow.com/a/37361490/3059369
    global var_escopo
    var_escopo = tk.StringVar()
    var_escopo.set('unidades')
    opcao_unidade = tk.Radiobutton(frame_radio, text="Em cada unidade",
                                   variable=var_escopo, value='unidades')
    opcao_total = tk.Radiobutton(frame_radio, text="Somente no total", variable=var_escopo,
                                 value='total')
    opcao_unidade.grid(row=1, column=0, sticky='w')
    opcao_total.grid(row=2, column=0, sticky='w')

    lista_variaveis = {'sinal': var_sinal, 'percentual': var_percentual,
                  'escopo': var_escopo, 'opcoes': lista_opcoes}
    # Botões
    # Alterar o 'background' do bottão do ttk na verdade muda só a cor da borda
    # Por isso foram usados botões normais
    botao_salvar = tk.Button(janela_nova, text="Salvar", bg="#afed80", width=10,
        command=lambda: clique_ok(lista_variaveis, janela_nova, var_erro, texto_erro))
    botao_salvar.grid(row=3, column=1, sticky='n', pady=10)
    botao_cancelar = tk.Button(janela_nova, text="Cancelar", bg="#f2cc63", width=10,
        command=lambda: janela_nova.destroy())
    botao_cancelar.grid(row=3, column=2, sticky='n', pady=10)

    # Label para mensagem de erro
    var_erro = tk.StringVar()
    texto_erro = tk.Label(janela_nova, name="erro", textvariable=var_erro)
    texto_erro.grid(row=4, column=0, columnspan=3)


def atualiza_tela():
    """Função para atualizar o tamanho do canvas após acrescentar elementos ao frame"""
    frame.update_idletasks()
    canvas.update_idletasks()
    canvas.config(height=frame.winfo_reqheight(), width=frame.winfo_reqwidth())
    canvas.config(scrollregion=canvas.bbox("all"))
    root.update()


def rolar(event):
    """Ativa a rolagem do canvas com a roda do mouse"""
    if canvas.winfo_exists():
        canvas.yview_scroll(-event.delta//120, "units")


### Fim das funçõoes ###

# ------ Interface gráfica ------
root = tk.Tk()
root.title("Otimizador de distribuição de professores 1.0 - Pedro Santos Guimarães")
# From https://www.tutorialspoint.com/how-to-set-the-position-of-a-tkinter-window-without-setting-the-dimensions
root.geometry("+300+100")
root.minsize(700,400)

# Cria o canvas
canvas = tk.Canvas(root, borderwidth=0)
canvas.grid(row=0, column=0, sticky="nsew")

# Frame dentro do canvas
frame = tk.Frame(canvas)

# Barra de rolagem
barra = ttk.Scrollbar(root, orient="vertical", command=canvas.yview)

canvas.configure(yscrollcommand=barra.set)
canvas.create_window((4,4), window=frame, anchor='nw')
canvas.bind_all("<MouseWheel>", rolar)

frame.bind("<Configure>", lambda event,
    canvas=canvas: canvas.configure(scrollregion=canvas.bbox("all")))

root.grid_rowconfigure(0, weight=1)
root.grid_columnconfigure(0, weight=1)

# Configura a primeira linha para expandir verticalmente
frame.grid_rowconfigure(1, weight=0)
frame.grid_rowconfigure(2, weight=1)

# Tamanho da fonte para todos os objetos
fonte = font.nametofont('TkDefaultFont')
fonte.configure(size=10)

# Título
TEXTO_TITULO = """Bem vindo.
Escolha o arquivo no botão abaixo, depois ajuste as opções e clique em Executar.
Dependendo da situação o programa pode levar um certo tempo para encontrar a solução ótima.
Os passos executados serão mostrados ao lado direito e ao final a distribuição será exibida na tela.
Você poderá baixar um relatório completo ou a planilha com a distribuição clicando nos botões."""

label_titulo = tk.Label(frame, text=TEXTO_TITULO, anchor="w", justify="left")
label_titulo.grid(sticky='W', row=0, column=0, padx=10, pady=10, columnspan=2)

# Grupo arquivo
grupo_arq = ttk.LabelFrame(frame)
grupo_arq.grid(row=1, column=0, rowspan=1, padx=10, pady=10, sticky='nw')

# Texto botão arquivo
texto_botao = tk.Label(grupo_arq, text="Escolha o arquivo:")
texto_botao.grid(row=1, column=0)

# Botão para selecionar o arquivo
botao_arquivo = tk.Button(grupo_arq, text="Abrir arquivo", command=carregar_arquivo, bg="#ddd")
botao_arquivo.grid(row=1, column=1, padx=10, pady=10)

# Label com nome do arquivo
var_nome_arquivo = tk.StringVar()
label_nome_arquivo = tk.Label(grupo_arq, textvariable=var_nome_arquivo)
label_nome_arquivo.grid(row=1, column=2, padx=10, pady=10)

# Grupo opções
ttk.Style().configure('Bold.TLabelframe.Label', font=('TkDefaulFont', 10, 'bold'))
ttk.Style().configure('Bold.TLabelframe', background='#ccc')
grupo_opcoes = ttk.LabelFrame(frame, text="Opções", style='Bold.TLabelframe')
grupo_opcoes.grid(row=2, column=0, padx=10, pady=10, rowspan=1, sticky='nw')

# Checkbox para ch minima
bool_minima = tk.BooleanVar(value=True)
checkbox_minima = tk.Checkbutton(grupo_opcoes, text="CH mínima: ", variable=bool_minima,
                                 command=verifica_check_boxes)
checkbox_minima.grid(row=3, column=0, padx=10, pady=10)

# campo texto
texto_ch_min = tk.IntVar(value=12)
entrada_CH_MIN = tk.Entry(grupo_opcoes, textvariable=texto_ch_min, width=5)
entrada_CH_MIN.grid(row=3, column=1, padx=10, pady=10)

ToolTip(checkbox_minima, msg="Ativar carga horária média máxima por unidade", delay=0.1)
ToolTip(entrada_CH_MIN, msg="Valor da carga horária máxima", delay=0.1)

# Checkbox para ch maxima
bool_maxima = tk.BooleanVar(value=True)
checkbox_maxima = tk.Checkbutton(grupo_opcoes, text="CH máxima: ", variable=bool_maxima,
                                 command=verifica_check_boxes)
checkbox_maxima.grid(row=4, column=0, padx=10, pady=10)

# campo texto
texto_ch_max = tk.IntVar(value=16)
entrada_CH_MAX = tk.Entry(grupo_opcoes, textvariable=texto_ch_max, width=5)
entrada_CH_MAX.grid(row=4, column=1, padx=10, pady=10)

ToolTip(checkbox_maxima, msg="Ativar carga horária média mínima geral (para toda a Universidade)",
        delay=0.1)
ToolTip(entrada_CH_MAX, msg="Valor da carga horária mínima", delay=0.1)

# Checkbox para total minimo
bool_min_total = tk.BooleanVar(value=False)
checkbox_min_total = tk.Checkbutton(grupo_opcoes, text="Total mínimo: ",
    variable=bool_min_total, command=verifica_check_boxes)
checkbox_min_total.grid(row=3, column=3, padx=10, pady=10)

# campo texto
texto_min_total = tk.IntVar()
entrada_N_MIN_total = tk.Entry(grupo_opcoes, textvariable=texto_min_total, width=5)
entrada_N_MIN_total.grid(row=3, column=4, padx=10, pady=10)

ToolTip(checkbox_min_total, msg="Ativar número mínimo total de professores", delay=0.1)
ToolTip(entrada_N_MIN_total, msg="Valor do mínimo total", delay=0.1)

# Checkbox para total maxima
bool_max_total = tk.BooleanVar(value=False)
checkbox_max_total = tk.Checkbutton(grupo_opcoes, text="Total máximo: ",
    variable=bool_max_total, command=verifica_check_boxes)
checkbox_max_total.grid(row=4, column=3, padx=10, pady=10)

# campo texto
texto_max_total = tk.IntVar()
entrada_N_MAX_total = tk.Entry(grupo_opcoes, textvariable=texto_max_total, width=5)
entrada_N_MAX_total.grid(row=4, column=4, padx=10, pady=10)

ToolTip(checkbox_max_total, msg="Ativar número máximo total de professores", delay=0.1)
ToolTip(entrada_N_MAX_total, msg="Valor do máximo total", delay=0.1)

# Tempo limite
val_limite = tk.IntVar(value=30)
label_limite = tk.Label(grupo_opcoes, text="Tempo limite:")
label_limite.grid(row=5, column=0, padx=10, pady=10)
entrada_tempo_limite = tk.Entry(grupo_opcoes, textvariable=val_limite, width=5)
entrada_tempo_limite.grid(row=5, column=1, padx=10, pady=10)

ToolTip(label_limite, msg="Tempo máximo para procurar a solução ótima", delay=0.1)

# Limitações nos perfis
botao_perfil = tk.Button(grupo_opcoes, text="Limitar perfis", command=janela_perfis, bg="#ddd")
botao_perfil.grid(row=5, column=3, padx=10, pady=10)

# Frame para a lista de limitações
frame_perfis = tk.Frame(grupo_opcoes)
frame_perfis.grid(row=5, column=4, padx=10, pady=10)

# Combobox critério
label_criterio = ttk.Label(grupo_opcoes, text="Critério:")
label_criterio.grid(row=6, column=0, padx=10, pady=10)

combo_var = tk.StringVar()
combobox = ttk.Combobox(grupo_opcoes, textvariable=combo_var, values=list(LISTA_MODOS.keys()),
                        state="readonly")
combobox.grid(row=6, column=1, padx=10, pady=10)
combobox.bind("<<ComboboxSelected>>", lambda event: verifica_executar())

# Botão para executar
botao_executar = tk.Button(grupo_opcoes, text="Executar", state=tk.DISABLED,
                           command=executar, width=10)
botao_executar.grid(row=10, column=0, padx=10, pady=10)

# Inicialmente oculta as opções
grupo_opcoes.grid_forget()

# Grupo dos resultados
grupo_resultados = ttk.LabelFrame(frame, text="Resultados", style='Bold.TLabelframe')
grupo_resultados.grid(row=1, column=1, padx=10, pady=10, rowspan=2, sticky='nsew')

resultado = tk.StringVar()
label_resultado = tk.Label(grupo_resultados, textvariable=resultado, anchor="w", justify="left")
label_resultado.grid(row=0, column=0, padx=10, pady=10, sticky='w')

label_aba = tk.Label(grupo_resultados, text="Distribuição:")
label_aba.grid(row=1, column=0, padx=10, pady=10, sticky='w')

# Adicionando a barra de rolagem vertical
scrollbar = tk.Scrollbar(grupo_resultados, orient="vertical")

text_aba = tk.Text(grupo_resultados, height=10, width=60, yscrollcommand=scrollbar.set)
text_aba.grid(row=2, column=0, padx=10, pady=10)

# Configurando a barra de rolagem para rolar o texto no Text widget
scrollbar.config(command=text_aba.yview)

botao_relatorio = tk.Button(grupo_resultados, text="Baixar relatório", command=exportar_txt,
                            bg="#ddd")
botao_relatorio.grid(row=3, column=0, padx=10, pady=10, sticky='w')

botao_planilha = tk.Button(grupo_resultados, text="Baixar planilha", command=exportar_planilha,
                           bg="#ddd")
botao_planilha.grid(row=3, column=0, padx=10, pady=10, sticky='e')

# Inicialmente oculta os resultados
grupo_resultados.grid_forget()

root.mainloop()
