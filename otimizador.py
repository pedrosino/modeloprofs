"""Modelo de programação linear para distribuição de vagas entre unidades acadêmicas
de uma universidade federal. Desenvolvido na pesquisa do Mestrado Profissional em
Gestão Organizacional, da Faculdade de Gestão e Negócios, Universidade Federal de Uberlândia,
por Pedro Santos Guimarães, em 2023"""

from datetime import datetime
import math
import tkinter as tk
from tkinter import filedialog
import pandas as pd
import numpy as np
from tktooltip import ToolTip
from pulp import lpSum, lpDot, LpVariable, LpStatus, PULP_CBC_CMD, \
    LpProblem, LpMaximize, LpMinimize, SCIP_CMD
import customtkinter
from funcoes import imprimir_resultados, imprimir_parametros, imprimir_unidades, imprimir_perfis
import ctypes

# Customtkinter
customtkinter.set_appearance_mode("System")  # Modes: system (default), light, dark
customtkinter.set_default_color_theme("dark-blue")  # Themes: blue (default), dark-blue, green

# Variáveis globais
MATRIZ_UNIDADES = None
NOMES_UNIDADES = None
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

# Vetor com formato do resultado conforme o modo
FORMATO_RESULTADO = {}
FORMATO_RESULTADO['num'] = 'professores'
FORMATO_RESULTADO['peq'] = 'prof-equivalente'
FORMATO_RESULTADO['tempo'] = 'horas'
FORMATO_RESULTADO['tempo-reverso'] = 'horas'
FORMATO_RESULTADO['ch'] = 'aulas/prof'
FORMATO_RESULTADO['ch-reverso'] = 'aulas/prof'
FORMATO_RESULTADO['todos'] = '(na escala de 0 a 100)'

# Restrição de carga horária média máxima por unidade
LIMITAR_CH_MINIMA = False
ESCOPO_CH_MINIMA = None
CH_MIN = 16
# Restrição de carga horária média mínima por unidade
LIMITAR_CH_MAXIMA = False
ESCOPO_CH_MAXIMA = None
CH_MAX = 16
# Restrição do número máximo total
MAX_TOTAL = False
N_MAX_TOTAL = None
# Restrição do número mínimo total
MIN_TOTAL = False
N_MIN_TOTAL = None
# Total de aulas a serem ministradas
TOTAL_AULAS = None
# Tempo limite para procurar a solução ótima
TEMPO_LIMITE = 30
MODO_ESCOLHIDO = 'todos'


#------------- Funções -------------

def verifica_executar():
    """Habilita ou desabilita o botão Executar"""
    if combo_var.get() and solver_var.get():
        botao_executar.configure(state="normal")
    else:
        botao_executar.configure(state="disabled")


def verifica_check_boxes():
    """Habilita ou desabilita os campos de texto conforme o checkbox"""
    if bool_minima.get():
        for item in grupo_min:
            item.grid()
    else:
        for item in grupo_min:
            item.grid_remove()

    if bool_maxima.get():
        for item in grupo_max:
            item.grid()
    else:
        for item in grupo_max:
            item.grid_remove()

    if bool_min_total.get():
        entrada_N_MIN_total.grid()
    else:
        entrada_N_MIN_total.grid_remove()

    if bool_max_total.get():
        entrada_N_MAX_total.grid()
    else:
        entrada_N_MAX_total.grid_remove()

    atualiza_tela()


def verifica_escopo():
    """Caso o usuário tenha selecionado 'unidades', mostra um aviso"""
    if escopo_ch_min.get() == 'unidades':
        var_erro_minima.set("Verifique se essa configuração não torna o modelo sem solução")
        formata_aviso(label_erro_minima)
    else:
        var_erro_minima.set("")
        limpa_erro(label_erro_minima)

    if escopo_ch_max.get() == 'unidades':
        var_erro_maxima.set("Verifique se essa configuração não torna o modelo sem solução")
        formata_aviso(label_erro_maxima)
    else:
        var_erro_maxima.set("")
        limpa_erro(label_erro_maxima)

def carregar_arquivo():
    """Carrega os dados da planilha"""
    global MATRIZ_UNIDADES, MATRIZ_PERFIS, MATRIZ_PEQ, MATRIZ_TEMPO, N_RESTRICOES, N_UNIDADES, \
           N_PERFIS, PESOS, NOMES_RESTRICOES, CONECTORES, NOMES_UNIDADES, TOTAL_AULAS
    # Importa dados do arquivo
    arquivo = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    var_nome_arquivo.set(arquivo.split("/")[-1])

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
    NOMES_UNIDADES = MATRIZ_UNIDADES[:, 0]
    # Remove primeira coluna
    MATRIZ_UNIDADES = np.delete(MATRIZ_UNIDADES, 0, axis=1)
    # Remove colunas à direita
    MATRIZ_UNIDADES = np.delete(MATRIZ_UNIDADES, slice(N_RESTRICOES, None), axis=1)
    MATRIZ_PEQ = MATRIZ_PERFIS[N_RESTRICOES+1]
    MATRIZ_TEMPO = MATRIZ_PERFIS[N_RESTRICOES] # tempo disponível
    # Lê critérios do AHP
    df_criterios = pd.read_excel(arquivo, sheet_name='criterios', usecols="B:B").dropna()
    pesos_lidos = df_criterios.to_numpy()
    PESOS = pesos_lidos.transpose().reshape(4)

    # Total de aulas a serem ministradas semanalmente
    TOTAL_AULAS = np.sum(MATRIZ_UNIDADES[:N_UNIDADES], axis=0)[0]

    # Se a importação teve sucesso
    if(len(MATRIZ_UNIDADES) and len(MATRIZ_PERFIS)):
        # Mostra as opções
        grupo_opcoes.grid()
        verifica_check_boxes()
        atualiza_tela()
        #centralizar()


def verifica_totais():
    """Verifica se os valores digitados pelo usuário são válidos"""
    ler_valores()
    tudo_certo = True

    # Verifica números totais
    if MIN_TOTAL:
        # Não pode ser zero
        if N_MIN_TOTAL < 1:
            var_erro_total_min.set("Deve ser maior que 0")
            formata_erro(label_erro_total_min)
            tudo_certo = False
        # Não pode ser maior que aulas/ch_min
        elif N_MIN_TOTAL > int(TOTAL_AULAS / CH_MIN):
            var_erro_total_min.set("Quantidade excessiva considerando\na carga horária mínima")
            formata_erro(label_erro_total_min)
            tudo_certo = False
        # Se for um valor muito baixo, não terá efeito - mostrar aviso
        elif N_MIN_TOTAL < round(TOTAL_AULAS / CH_MAX):
            var_erro_total_min.set("Aviso: restrição não terá efeito")
            formata_aviso(label_erro_total_min)
        else:
            var_erro_total_min.set("")
            limpa_erro(label_erro_total_min)
    else:
        var_erro_total_min.set("")
        limpa_erro(label_erro_total_min)

    if MAX_TOTAL:
        # Não pode ser zero
        if N_MAX_TOTAL < 1:
            var_erro_total_max.set("Deve ser maior que 0")
            formata_erro(label_erro_total_max)
            tudo_certo = False
        # Não pode ser menor que aulas/ch_max
        elif N_MAX_TOTAL < round(TOTAL_AULAS / CH_MAX):
            var_erro_total_max.set("Quantidade insuficiente considerando\na carga horária máxima")
            formata_erro(label_erro_total_max)
            tudo_certo = False
        # Se for um valor muito alto, não terá efeito - mostrar aviso
        elif N_MAX_TOTAL > int(TOTAL_AULAS / CH_MIN):
            var_erro_total_max.set("Aviso: restrição não terá efeito")
            formata_aviso(label_erro_total_max)
        else:
            var_erro_total_max.set("")
            limpa_erro(label_erro_total_max)
    else:
        var_erro_total_max.set("")
        limpa_erro(label_erro_total_max)

    return tudo_certo


def verifica_carga_horaria():
    """Verifica se os valores digitados pelo usuário são válidos"""
    ler_valores()
    tudo_certo = True

    # 10 aulas de 50 min = 8h20 min por semana (mínimo legal é 8h - Lei 9.394, art. 57)
    if LIMITAR_CH_MINIMA and (CH_MIN < 10 or CH_MIN > 24):
        var_erro_minima.set("Deve ser entre 10 e 24")
        formata_erro(label_erro_minima)
        tudo_certo = False
    else:
        var_erro_minima.set("")
        limpa_erro(label_erro_minima)

    # Limite máximo é 24 aulas de 50 min = 20h - Decreto 9.235, art. 93, parág. único
    if LIMITAR_CH_MAXIMA and (CH_MAX < 10 or CH_MAX > 24):
        var_erro_maxima.set("Deve ser entre 10 e 24")
        formata_erro(label_erro_maxima)
        tudo_certo = False
    else:
        var_erro_maxima.set("")
        limpa_erro(label_erro_maxima)

    return tudo_certo


def ler_valores():
    """Captura os valores preenchidos pelo usuário"""
    global LIMITAR_CH_MAXIMA, LIMITAR_CH_MINIMA, CH_MAX, CH_MIN, MAX_TOTAL, MIN_TOTAL, \
        N_MIN_TOTAL, N_MAX_TOTAL, ESCOPO_CH_MINIMA, ESCOPO_CH_MAXIMA
    # Captura opções e valores escolhidos
    LIMITAR_CH_MAXIMA = bool_maxima.get()
    LIMITAR_CH_MINIMA = bool_minima.get()
    ESCOPO_CH_MINIMA = escopo_ch_min.get()
    ESCOPO_CH_MAXIMA = escopo_ch_max.get()

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


def executar():
    """Executa a otimização"""
    global QTDES_FINAL, RELATORIO, DATA_FRAME, MODO_ESCOLHIDO, TEMPO_LIMITE

    # Limpa tabela
    text_tabela.delete("1.0", tk.END)

    # Captura valores
    ler_valores()

    TEMPO_LIMITE = val_limite.get()
    MODO_ESCOLHIDO = LISTA_MODOS[combo_var.get()]

    # Verifica modo escolhido
    if MODO_ESCOLHIDO not in ['num', 'peq', 'tempo', 'tempo-reverso', 'ch', 'ch-reverso', 'todos']:
        MODO_ESCOLHIDO = 'todos'
        print("O modo escolhido era inválido. Será utilizado o modo 'todos'.")

    # Verifica valores digitados
    check_totais = verifica_totais()

    check_ch = verifica_carga_horaria()

    # Se houve algum problema, não faz o processo
    if not check_totais or not check_ch:
        return

    # Mostra resultados
    grupo_resultados.grid()
    grupo_resultados.configure(width=130 + N_PERFIS*45)
    grupo_botoes.grid()
    atualiza_tela()

    # Inicia relatório
    RELATORIO = "Relatório da execução do otimizador" \
        + f"\nData: {datetime.now().strftime('%d/%m/%Y')}\n" \
        + f"Arquivo carregado: {var_nome_arquivo.get()}\n"
    RELATORIO += f'\nModo escolhido: {[k for k, v in LISTA_MODOS.items() if v == MODO_ESCOLHIDO][0]}'
    RELATORIO += f'\nUnidades: {N_UNIDADES}'
    RELATORIO += f'\nCarga horária máxima: {CH_MAX if LIMITAR_CH_MAXIMA else "-"}'
    RELATORIO += f'\nCarga horária mínima: {CH_MIN if LIMITAR_CH_MINIMA else "-"}'
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
    RELATORIO += imprimir_unidades(N_UNIDADES, N_RESTRICOES, \
                                   MATRIZ_UNIDADES, NOMES_UNIDADES, NOMES_RESTRICOES)

    # Imprime perfis
    RELATORIO += imprimir_perfis(N_PERFIS, N_RESTRICOES, MATRIZ_PERFIS, NOMES_RESTRICOES)

    RELATORIO += '\n------------------------------------------------------------\n'

    texto_resultado = f"O modo escolhido foi {[k for k, v in LISTA_MODOS.items() if v == MODO_ESCOLHIDO][0]}"
    if MODO_ESCOLHIDO == 'todos':
        texto_resultado += "\nPrimeiro vamos definir os parâmetros para cada critério"

    resultado.set(texto_resultado)
    atualiza_tela()

    try:
        # Conforme o modo escolhido, faz só uma otimização ou todas
        if MODO_ESCOLHIDO != 'todos':
            texto_resultado = resultado.get()
            texto_resultado += "\nResolvendo ..."
            resultado.set(texto_resultado)
            atualiza_tela()
            resultado_final, QTDES_FINAL = otimizar(MODO_ESCOLHIDO, None, None)
            if resultado_final == -999999:
                raise ValueError('Erro no modelo')
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
                texto_resultado += f"\nModo {[k for k, v in LISTA_MODOS.items() if v == modo_usado][0]}: resolvendo ..."
                resultado.set(texto_resultado)
                atualiza_tela()

                # Obtém resultado e quantidades
                resultado_modo, qtdes_modo = otimizar(modo_usado, piores, melhores)

                if resultado_modo == -999999:
                    raise ValueError('Erro no modelo')

                # Registra o resultado na lista de melhores casos
                if 'reverso' not in modo_usado:
                    melhores[modo_usado] = resultado_modo
                # ou na de piores casos
                else:
                    piores[modo_usado.split('-')[0]] = resultado_modo

                texto_resultado = resultado.get()
                if 'reverso' not in modo_usado:
                    texto_resultado += f"\nTerminado. Resultado: {melhores[modo_usado]}, "\
                        f"resolvido em {MODELOS[modo_usado].solutionTime:.3f} segundos"
                else:
                    texto_resultado += f"\nTerminado. Resultado: {piores[modo_usado.split('-')[0]]}, "\
                        f"resolvido em {MODELOS[modo_usado].solutionTime:.3f} segundos"
                resultado.set(texto_resultado)
                atualiza_tela()

                # Imprime resultados
                RELATORIO += imprimir_resultados(qtdes_modo, N_PERFIS, N_UNIDADES, MATRIZ_UNIDADES, \
                                                NOMES_UNIDADES, MATRIZ_PEQ, MATRIZ_TEMPO)
                # Imprime parâmetros
                RELATORIO += imprimir_parametros(qtdes_modo, N_UNIDADES, N_RESTRICOES, \
                    MATRIZ_UNIDADES, NOMES_UNIDADES, RESTRICOES_PERCENTUAIS, MATRIZ_PERFIS, NOMES_RESTRICOES)

                # Para os modos de carga horária, imprime as médias e desvios
                if 'ch' in modo_usado:
                    for variable in MODELOS[modo_usado].variables():
                        nomes_busca = ['modulo', 'media']
                        if any(nome in variable.name for nome in nomes_busca):
                            RELATORIO += f"\n{variable.name}: {variable.value():.4f}"

                RELATORIO += "\n------------------------------------------------------------\n"

            ##### -----------------Fim da primeira 'rodada'-----------------
            RELATORIO += f"\nMelhores: {melhores}\nPiores: {piores}\n"
            RELATORIO += "------------------------------------------------------------"

            ## Uma nova rodada do modelo usando os pesos e as listas de melhores e piores casos
            # Atualiza o texto do resultado
            texto_resultado = resultado.get()
            texto_resultado += "\n\nModo Todos: resolvendo ..."
            resultado.set(texto_resultado)
            atualiza_tela()

            # Otimização com o modo 'todos'
            resultado_final, QTDES_FINAL = otimizar('todos', piores, melhores)

            texto_resultado = resultado.get()
            texto_resultado += f"\nTerminado. Resultado: {resultado_final}, "\
                f"resolvido em {MODELOS['todos'].solutionTime:.3f} segundos"
            resultado.set(texto_resultado)
            atualiza_tela()

        # ---- Finalização ----
        # Imprime resultados
        RELATORIO += imprimir_resultados(QTDES_FINAL, N_PERFIS, N_UNIDADES, MATRIZ_UNIDADES, \
                                        NOMES_UNIDADES, MATRIZ_PEQ, MATRIZ_TEMPO)
        # Imprime parâmetros
        RELATORIO += imprimir_parametros(QTDES_FINAL, N_UNIDADES, N_RESTRICOES, \
            MATRIZ_UNIDADES, NOMES_UNIDADES, RESTRICOES_PERCENTUAIS, MATRIZ_PERFIS, NOMES_RESTRICOES)

        if 'ch' in MODO_ESCOLHIDO:
            for variable in MODELOS[MODO_ESCOLHIDO].variables():
                nomes_busca = ['modulo', 'media']
                if any(nome in variable.name for nome in nomes_busca):
                    RELATORIO += f"\n{variable.name}: {variable.value():.4f}"

        if MODO_ESCOLHIDO == 'todos':
            # PESOS
            RELATORIO += f"\nPESOS: {PESOS}\n"
            # Imprime médias e desvios
            for variable in MODELOS['todos'].variables():
                nomes_busca = ['p_', 'modulo', 'media']
                if any(nome in variable.name for nome in nomes_busca):
                    RELATORIO += f"\n{variable.name}: {variable.value():.4f}"

        # Imprime o modelo completo
        RELATORIO += "\n\n------------------Modelo:------------------\n"
        ## https://stackoverflow.com/a/1140967/3059369
        modelo = f"\n{MODELOS[MODO_ESCOLHIDO]}"
        RELATORIO += "".join([s for s in modelo.splitlines(True) if s.strip("\r\n")])

        texto_resultado = resultado.get()
        texto_resultado += f"\nSituação: {MODELOS[MODO_ESCOLHIDO].status}, " \
            f"{LpStatus[MODELOS[MODO_ESCOLHIDO].status]}"
        texto_resultado += f"\nObjetivo: {resultado_final} {FORMATO_RESULTADO[MODO_ESCOLHIDO]}"
        texto_resultado += f"\nResolvido em {MODELOS[MODO_ESCOLHIDO].solutionTime:.3f} segundos"

        resultado.set(texto_resultado)
        atualiza_tela()

        # Transforma em dataframe com cabeçalho e unidades
        DATA_FRAME = pd.DataFrame(QTDES_FINAL, columns=[f'p{i}' for i in range(1, N_PERFIS+1)])
        # Manter os tipos
        DATA_FRAME = DATA_FRAME.convert_dtypes()
        # Linha com totais
        DATA_FRAME.loc['Total', :] = DATA_FRAME.sum().values
        # Coluna com total
        DATA_FRAME['Total'] = DATA_FRAME.sum(axis=1, numeric_only=True)
        # Insere coluna
        DATA_FRAME.insert(0, "Unidade", np.append(NOMES_UNIDADES,['Total Perfil']))

        # Mostra na tabela
        text_tabela.insert(tk.END, DATA_FRAME.to_string(index=False))

        # Obtém o número total de linhas do texto
        num_linhas = int(text_tabela.index(tk.END).split('.', maxsplit=1)[0])

        # Ajusta a altura do widget para mostrar no máximo altura_maxima linhas
        text_tabela.configure(height=num_linhas*18, width=120 + N_PERFIS*45)

        atualiza_tela()
        centralizar()
    except ValueError as excp:
        print(repr(excp))


def otimizar(modo, piores, melhores):
    """Função que faz a otimização conforme o modo escolhido"""
    global RELATORIO

    #nomes
    nomes = [str(perfil) + "_"
        + NOMES_UNIDADES[und] for und in range(N_UNIDADES) for perfil in range(1, N_PERFIS+1)]

    # Variáveis de decisão
    var_x = LpVariable.matrix("x", nomes, cat="Integer", lowBound=0)
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

    # -- Definir o modelo --
    if modo in ['tempo', 'ch-reverso', 'todos']:
        MODELOS[modo] = LpProblem(name=f"Professores-{modo}", sense=LpMaximize)
    else:
        MODELOS[modo] = LpProblem(name=f"Professores-{modo}", sense=LpMinimize)

    # -- Restrições --
    for restricao in range(N_RESTRICOES):
        for unidade in range(N_UNIDADES):
            match CONECTORES[restricao]:
                case ">=":
                    MODELOS[modo] += lpDot(saida[unidade], MATRIZ_PERFIS[restricao]) \
                        >= MATRIZ_UNIDADES[unidade][restricao], \
                        NOMES_RESTRICOES[restricao] + " " + NOMES_UNIDADES[unidade]
                case "==":
                    MODELOS[modo] += lpDot(saida[unidade], MATRIZ_PERFIS[restricao]) \
                        == MATRIZ_UNIDADES[unidade][restricao], \
                        NOMES_RESTRICOES[restricao] + " " + NOMES_UNIDADES[unidade]
                case "<=":
                    MODELOS[modo] += lpDot(saida[unidade], MATRIZ_PERFIS[restricao]) \
                        <= MATRIZ_UNIDADES[unidade][restricao], \
                        NOMES_RESTRICOES[restricao] + " " + NOMES_UNIDADES[unidade]

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
            #nome_restricao = "Perfis (" + ",".join(str(p) for p in perfis) \
            #    + f") {sinal} {percentual*100}% "
            # O solver SCIP não aceita sinais <= no nome da restrição
            nome_restricao = "Perfis (" + ",".join(str(p) for p in perfis) \
                + f") {'menor' if sinal == '<=' else 'maior'} {percentual*100}% "

            # Se a restrição for só nas unidades, automaticamente valerá para o total
            if escopo == 'unidades':
                for unidade in range(N_UNIDADES):
                    match sinal:
                        # Monta a restrição conforme o sinal escolhido
                        case '<=':
                            MODELOS[modo] += lpSum(saida[unidade]*coeficientes) <= 0, \
                                nome_restricao + NOMES_UNIDADES[unidade]
                        case '>=':
                            MODELOS[modo] += lpSum(saida[unidade]*coeficientes) >= 0, \
                                nome_restricao + NOMES_UNIDADES[unidade]
            else:
                # Restrição no total
                match sinal:
                    case '<=':
                        MODELOS[modo] += lpSum(saida*coeficientes) <= 0, \
                            f'{nome_restricao} geral'
                    case '>=':
                        MODELOS[modo] += lpSum(saida*coeficientes) >= 0, \
                            f'{nome_restricao} geral'

    # Restrições de carga horária média:
    # ------------------
    # Exemplo para 900 aulas, considerando carga horária média mínima de 12 aulas e máxima de 16
    # soma <= 900/12 -> soma <= 75
    # soma >= 900/16 -> soma >= 56.25
    # a soma deve estar entre 57 e 75, inclusive
    if maxima:
        # Restrição no total
        MODELOS[modo] += CH_MAX*lpSum(saida) >= TOTAL_AULAS, \
            f"chmax {math.ceil(TOTAL_AULAS/CH_MAX)}"
        # E por unidade
        for unidade in range(N_UNIDADES):
            n_min = math.ceil(MATRIZ_UNIDADES[unidade][0]/CH_MAX)
            MODELOS[modo] += CH_MAX*lpSum(saida[unidade]) >= MATRIZ_UNIDADES[unidade][0], \
                f"{NOMES_UNIDADES[unidade]}_chmax {n_min}"

    if minima:
        # Restrição mínima somente no geral -> permite exceções
        MODELOS[modo] += CH_MIN*lpSum(saida) <= TOTAL_AULAS, \
            f"chmin {int(TOTAL_AULAS/CH_MIN)}"
        # Caso a restrição se aplique também às unidades, usa o mesmo valor

        if ESCOPO_CH_MINIMA == 'unidades':
            minimo_local = CH_MIN

        # caso contrário, mínimo em cada unidade de 10 aulas/semana, equivale a 8 horas/semana,
        # para atender ao art. 57 da Lei nº 9.394, de 1996
        else:
            minimo_local = 10
        # Aplica restrição a cada unidade
        for unidade in range(N_UNIDADES):
            n_max = int(MATRIZ_UNIDADES[unidade][0]/minimo_local)
            MODELOS[modo] += minimo_local*lpSum(saida[unidade]) <= MATRIZ_UNIDADES[unidade][0], \
                f"{NOMES_UNIDADES[unidade]}_chmin {n_max}"

    # Restrição do número total de professores
    if MAX_TOTAL:
        MODELOS[modo] += lpSum(saida) <= N_MAX_TOTAL, f"TotalMax {N_MAX_TOTAL}"
    if MIN_TOTAL:
        MODELOS[modo] += lpSum(saida) >= N_MIN_TOTAL, f"TotalMin {N_MIN_TOTAL}"

    # Modo de equilíbrio da carga horária
    if modo in ['ch', 'ch-reverso', 'todos']:
        # variáveis auxiliares
        # Para cada unidade há um valor da média e um desvio em relação à média geral
        modulos = LpVariable.matrix("modulo", NOMES_UNIDADES, \
                                    cat="Continuous", lowBound=0)
        consts = LpVariable.matrix("b", NOMES_UNIDADES, cat="Binary")
        medias = LpVariable.matrix("media", NOMES_UNIDADES, \
                                   cat="Continuous", lowBound=0)
        media_geral = LpVariable("media_geral", cat="Continuous", lowBound=0)

        # Como a média seria uma função não linear (aulas/professores), foi feita uma
        # aproximação com uma reta que passa pelos dois pontos extremos dados pelos
        # valores de carga horária mínima e máxima.
        # Coeficientes
        # coef = (y1-y0)/(x1-x0) -> y1 = media minima, y0 = media maxima,
        # x1 = numero maximo, x0 = numero minimo
        x_1 = TOTAL_AULAS / CH_MIN
        x_0 = TOTAL_AULAS / CH_MAX
        coef = (CH_MIN - CH_MAX) / (x_1 - x_0)

        # O cálculo da média é inserido no modelo como uma restrição
        MODELOS[modo] += media_geral == coef*lpSum(saida) + CH_MIN + CH_MAX, "Media geral"

        # O mesmo raciocínio é feito para cada unidade
        for unidade in range(N_UNIDADES):
            x1u = MATRIZ_UNIDADES[unidade][0] / CH_MIN
            x0u = MATRIZ_UNIDADES[unidade][0] / CH_MAX
            coef_u = (CH_MIN - CH_MAX) / (x1u - x0u)
            # Cálculo da media
            MODELOS[modo] += medias[unidade] == coef_u*lpSum(saida[unidade]) + CH_MIN + CH_MAX, \
                f"{NOMES_UNIDADES[unidade]}_ch_media"
            # Cálculo do desvio
            # O desvio é dado pelo módulo da subtração, porém o PuLP não aceita a função abs()
            # Assim, são colocadas duas restrições, uma usando o valor positivo e outra o negativo
            ## https://optimization.cbe.cornell.edu/index.php?title=Optimization_with_absolute_values
            ## https://lpsolve.sourceforge.net/5.5/absolute.htm
            m_grande = 10000
            # desvio
            # X + M * B >= x'
            MODELOS[modo] += medias[unidade] - media_geral\
                + m_grande*consts[unidade] >= modulos[unidade], \
                f"XMB {NOMES_UNIDADES[unidade]}"
            # -X + M * (1-B) >= x'
            MODELOS[modo] += -1*(medias[unidade] - media_geral) \
                + m_grande*(1-consts[unidade]) >= modulos[unidade], \
                f"-XM1B {NOMES_UNIDADES[unidade]}"
            # modulo
            MODELOS[modo] += modulos[unidade] >= medias[unidade] - media_geral, \
                f"{NOMES_UNIDADES[unidade]}_mod1"
            MODELOS[modo] += modulos[unidade] >= -1*(medias[unidade] - media_geral), \
                f"{NOMES_UNIDADES[unidade]}_mod2"

    # Modo com todos os critérios
    if modo == 'todos':
        # Variáveis com as pontuações
        fator = 100
        pontuacoes = LpVariable.matrix("p", range(4), cat="Continuous", lowBound=0, upBound=fator)
        # Restrições/cálculos
        # Caso seja dado um número exato, é necessário alterar a pontuação do critério 'num'
        # para evitar a divisão por zero
        if MAX_TOTAL and MIN_TOTAL and N_MAX_TOTAL == N_MIN_TOTAL:
            MODELOS[modo] += pontuacoes[0] == fator*lpSum(saida)/melhores['num'], "Pontuação número"
        else:
            MODELOS[modo] += pontuacoes[0] == fator*(lpSum(saida) - piores['num'])/(melhores['num'] - piores['num']), "Pontuação número"
        MODELOS[modo] += pontuacoes[1] == fator*(lpSum(saida*MATRIZ_PEQ) - piores['peq'])/(melhores['peq'] - piores['peq']), "Pontuação P-Eq"
        MODELOS[modo] += pontuacoes[2] == fator*(lpSum(saida*MATRIZ_TEMPO) - np.sum(MATRIZ_UNIDADES[:N_UNIDADES], axis=0)[1] - piores['tempo'])/(melhores['tempo'] - piores['tempo']), "Pontuação tempo"
        MODELOS[modo] += pontuacoes[3] == fator*(lpSum(modulos[:N_UNIDADES])/N_UNIDADES - piores['ch'])/(melhores['ch'] - piores['ch']), "Pontuação Equilíbrio"

    # -- Função objetivo --
    if modo == 'num':
        MODELOS[modo] += lpSum(saida)
    elif modo == 'peq':
        MODELOS[modo] += lpSum(saida*MATRIZ_PEQ)
    # No modo tempo é necessário deduzir do tempo calculado as horas
    # que serão destinadas às orientações
    elif 'tempo' in modo:
        MODELOS[modo] += lpSum(saida*MATRIZ_TEMPO) - np.sum(MATRIZ_UNIDADES[:N_UNIDADES], axis=0)[1]
    # No modo de equilíbrio a medida de desempenho é o desvio médio
    elif 'ch' in modo:
        MODELOS[modo] += lpSum(modulos[:N_UNIDADES])/N_UNIDADES
    elif modo == 'todos':
        # O objetivo é maximimizar a pontuação dada pela soma de cada
        # pontuação multiplicada por seu peso
        MODELOS[modo] += lpSum(pontuacoes*PESOS)

    # Imprime os parâmetros do modo
    RELATORIO += f'\nModo: {[k for k, v in LISTA_MODOS.items() if v == modo][0]}'
    RELATORIO += f'\nCarga horária máxima: {CH_MAX if maxima else "-"}'
    RELATORIO += f'\nCarga horária mínima: {CH_MIN if minima else "-"}'

    # Ajusta limite
    # Para os modos "intermediários" não é necessário um limite maior que 30 segundos
    #if MODO_ESCOLHIDO == 'todos' and modo != 'todos' and TEMPO_LIMITE > 30:
    #    novo_limite = 30
    #else:
    #    novo_limite = TEMPO_LIMITE
    novo_limite = TEMPO_LIMITE

    # Resolver o modelo
    solver_escolhido = solver_var.get()
    if solver_escolhido == 'CBC':
        MODELOS[modo].solve(PULP_CBC_CMD(msg=1, timeLimit=novo_limite))
    elif solver_escolhido == 'SCIP':
        MODELOS[modo].solve(SCIP_CMD(msg=1, timeLimit=novo_limite,
            path="C:\\Program Files\\SCIPOptSuite 8.0.4\\bin\\scip.exe"))
    # O solver GLPK é bem mais lento
    ##MODELOS[modo].solve(GLPK_CMD(msg=1, options=["--tmlim", str(novo_limite)]))

    # Se não foi encontrada uma solução
    if MODELOS[modo].status != 1:
        # Avisa usuário e interrompe o processo
        texto_resultado = resultado.get()
        texto_resultado += "\nNão foi possível encontrar uma solução. Verifique as opções escolhidas."
        resultado.set(texto_resultado)
        atualiza_tela()
        return -999999, []

    # Resultados
    RELATORIO += f"\nSituação: {MODELOS[modo].status}, {LpStatus[MODELOS[modo].status]}"
    # Para cada critério o resultado é em um formato diferente
    if 'ch' in modo:
        objetivo = round(MODELOS[modo].objective.value(), 4)
    elif modo == 'peq':
        objetivo = round(MODELOS[modo].objective.value(), 2)
    elif modo == 'num':
        objetivo = int(MODELOS[modo].objective.value())
    elif 'tempo' in modo:
        objetivo = round(MODELOS[modo].objective.value(), 2)
    elif modo == 'todos':
        objetivo = round(MODELOS[modo].objective.value(), 4)

    RELATORIO += f"\nObjetivo: {objetivo} {FORMATO_RESULTADO[modo]}"
    RELATORIO += f"\nResolvido em {MODELOS[modo].solutionTime:.3f} segundos\n"

    # Extrai as quantidades
    qtdes_saida = np.full((N_UNIDADES, N_PERFIS), 0, dtype=int)
    for var in MODELOS[modo].variables():
        if var.name.find('x') == 0:
            _, perfil, unidade = var.name.split("_")
            perfil = int(perfil) - 1
            ind_unidade = np.where(NOMES_UNIDADES == unidade)[0][0]
            qtdes_saida[ind_unidade][perfil] = round(var.value())

    # Retorna o valor da função objetivo e as quantidades
    return objetivo, qtdes_saida


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
    label.configure(fg_color='#f0a869', text_color='#87190b')
    label.grid()
    atualiza_tela()


def limpa_erro(label):
    """Limpa o formato do label"""
    label.configure(fg_color=root.cget('fg_color'))
    label.grid_remove()
    atualiza_tela()


def formata_aviso(label):
    """Formata o label do aviso"""
    label.configure(fg_color='#f0eb69', text_color='#876e0b')
    label.grid()
    atualiza_tela()


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
    label_perfil = customtkinter.CTkLabel(frame_labels, text=texto_perfis)
    label_perfil.pack(side=tk.LEFT)

    botao_excluir = customtkinter.CTkButton(
        frame_labels, text='X', fg_color="#900", width=20,
        command=lambda: excluir_restricao(nome_restricao, frame_labels)
    )
    botao_excluir.pack(side=tk.LEFT)
    atualiza_tela()

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
    janela_nova = customtkinter.CTkToplevel(root)
    #janela_nova = tk.Toplevel(root)
    janela_nova.geometry("+320+220")

    titulo_janela = customtkinter.CTkLabel(janela_nova, text="Selecione os perfis:")
    titulo_janela.grid(row=0, column=0, padx=10, pady=10)

    # Lista de perfis
    grupo_perfis = customtkinter.CTkFrame(janela_nova)
    grupo_perfis.grid(row=1, column=0, padx=10, pady=10, rowspan=3)

    lista_opcoes = []
    for perfil in range(N_PERFIS):
        texto = f"Perfil {perfil+1}"
        var = tk.IntVar()
        check_button = customtkinter.CTkCheckBox(grupo_perfis, text=texto, variable=var)
        check_button.pack(anchor='w', padx=10, pady=(10,0))
        lista_opcoes.append(var)

    # Sinal da operação
    grupo_sinal = customtkinter.CTkFrame(janela_nova)
    grupo_sinal.grid(row=1, column=1, columnspan=2, padx=10, pady=10, sticky='n')
    label_sinal = customtkinter.CTkLabel(
        grupo_sinal, text="A soma desse(s) perfil(s) deve ser"
    )
    label_sinal.grid(row=0, column=0, padx=10, pady=10, columnspan=4)
    var_sinal = tk.StringVar()
    radio_sinal_menor = customtkinter.CTkRadioButton(
        grupo_sinal, text="<=", value="<=", variable=var_sinal, width=30
    )
    radio_sinal_maior = customtkinter.CTkRadioButton(
        grupo_sinal, text=">=", value=">=", variable=var_sinal, width=30
    )
    radio_sinal_menor.grid(row=1, column=0, sticky='e', padx=(10,0))
    radio_sinal_maior.grid(row=2, column=0, sticky='e', padx=(10,0))

    # Valor do percentual
    var_percentual = tk.DoubleVar()
    texto_percentual = customtkinter.CTkEntry(grupo_sinal, textvariable=var_percentual, width=40)
    texto_percentual.grid(row=1, column=1, rowspan=2, padx=0)
    label_perc = customtkinter.CTkLabel(grupo_sinal, text="%")
    label_perc.grid(row=1, column=2, sticky='w', rowspan=2, padx=(0,10))
    ToolTip(texto_percentual, msg="Para valores não inteiros, use ponto decimal", delay=0.1)

    # Switch para escolher se a restrição é só geral ou também em cada unidade
    frame_escopo = customtkinter.CTkFrame(janela_nova)
    frame_escopo.grid(row=2, column=1, padx=10, pady=10, columnspan=2)
    label_t = customtkinter.CTkLabel(
        frame_escopo, text="Essa restrição se aplica a"
    )
    label_t.grid(row=3, column=0, sticky='w', padx=10, pady=10, columnspan=2)
    label_duvida = customtkinter.CTkLabel(frame_escopo, text=" (?)")
    label_duvida.grid(row=3, column=2, sticky='ew')
    ToolTip(label_duvida,
        msg="Escolhendo a opção 'total' o percentual em cada unidade poderá extrapolar a restrição",
        delay=0.1)

    var_escopo = tk.StringVar()
    var_escopo.set('unidades')
    label_u = customtkinter.CTkLabel(frame_escopo, text="cada unidade")
    opcao_perfil = customtkinter.CTkSwitch(
        frame_escopo, text="total", onvalue="total", offvalue="unidades"
        , variable=var_escopo
        , fg_color=("#3B8ED0", "#4A4D50"), progress_color=("#3B8ED0", "#4A4D50")
        , button_color=("#275e8a", "#D5D9DE") , button_hover_color=("#1d4566","gray100")
    )
    label_u.grid(row=4, column=0, padx=(10,0), pady=(0,10), sticky='e')
    opcao_perfil.grid(row=4, column=1, padx=(10,0), pady=(0,10), sticky='w')

    lista_variaveis = {'sinal': var_sinal, 'percentual': var_percentual,
                  'escopo': var_escopo, 'opcoes': lista_opcoes}

    # Botões
    botao_salvar = customtkinter.CTkButton(
        janela_nova, text="Salvar", fg_color="#3e811d", hover_color="#428c20",
        command=lambda: clique_ok(lista_variaveis, janela_nova, var_erro, texto_erro)
    )
    botao_salvar.grid(row=3, column=1, sticky='s', pady=10)
    botao_cancelar = customtkinter.CTkButton(
        janela_nova, text="Cancelar", fg_color="#c16100", hover_color="#b87b09",
        command=lambda: janela_nova.destroy()
    )
    botao_cancelar.grid(row=3, column=2, sticky='s', pady=10)

    # Label para mensagem de erro
    var_erro = tk.StringVar()
    texto_erro = customtkinter.CTkLabel(janela_nova, textvariable=var_erro)
    texto_erro.grid(row=4, column=0, columnspan=3)

    return janela_nova

def abrir_janela():
    """Dá foco na janela nova"""
    janela = janela_perfis()
    janela.wm_transient(root)

def atualiza_tela():
    """Função para atualizar o tamanho do canvas após acrescentar elementos ao frame"""
    root.update_idletasks()
    root.update()


def centralizar():
    """Ajusta posição da janela na tela"""
    altura_janela = root.winfo_reqheight()
    altura_tela = root.winfo_screenheight()
    largura_janela = root.winfo_reqwidth()
    largura_tela = root.winfo_screenwidth()

    # Calcula novas coordenadas
    novo_x = round((largura_tela - largura_janela) / 2)
    novo_y = round((altura_tela - altura_janela - 100) / 2)

    # Verifica o tamanho da janela
    if altura_tela - altura_janela < 80:
        novo_y = 0
        altura_janela = altura_tela - 80

    # Bug quando a tela está com escala diferente de 100%
    # https://github.com/TomSchimansky/CustomTkinter/issues/1707
    scale_factor = ctypes.windll.shcore.GetScaleFactorForDevice(0)/100
    largura_janela_ajustada = int(largura_janela / scale_factor)

    # Atualiza tamanho do grupo_resultados
    grupo_resultados.configure(height=altura_janela - 230)
    # Atualiza tela
    root.geometry(f"{largura_janela_ajustada}x{altura_janela}+{novo_x}+{novo_y}")
    root.update()
    atualiza_tela()


def mudar_modo():
    """Muda aparência"""
    customtkinter.set_appearance_mode(opcao_modo.get())


### Fim das funçõoes ###

# ------ Interface gráfica ------
root = customtkinter.CTk()
root.title("Otimizador de distribuição de professores 1.0 - Pedro Santos Guimarães")
# From https://www.tutorialspoint.com/how-to-set-the-position-of-a-tkinter-window-without-setting-the-dimensions
root.geometry("+100+50")
#root.minsize(700,400)

#root.grid_rowconfigure(0, weight=1)
#root.grid_columnconfigure(0, weight=1)

# Título
TEXTO_TITULO = """Bem vindo.
Escolha o arquivo no botão abaixo, depois ajuste as opções e clique em Executar.
Dependendo da situação o programa pode levar um certo tempo para encontrar a solução ótima.
Os passos executados serão mostrados ao lado direito e ao final a distribuição será exibida na tela.
Você poderá baixar um relatório completo ou a planilha com a distribuição clicando nos botões."""

label_titulo = customtkinter.CTkLabel(root, text=TEXTO_TITULO, anchor="w", justify="left")
label_titulo.grid(row=0, column=0, padx=10, pady=10, sticky='W', columnspan=2)

opcao_modo = customtkinter.CTkOptionMenu(
    root, values=["Light", "Dark", "System"],
    command=lambda *args: mudar_modo()
)
#opcao_modo.grid(row=0, column=2, padx=10, sticky="w")

# -------------- Grupo arquivo --------------
grupo_arq = customtkinter.CTkFrame(root)
grupo_arq.grid(row=1, column=0, rowspan=1, padx=10, pady=10, sticky='nw')

# Texto botão arquivo
texto_botao = customtkinter.CTkLabel(grupo_arq, text="Escolha o arquivo:")
texto_botao.grid(row=0, column=0, padx=10)

# Botão para selecionar o arquivo
botao_arquivo = customtkinter.CTkButton(
    grupo_arq, text="Abrir arquivo", command=carregar_arquivo#, bg="#ddd"
)
botao_arquivo.grid(row=0, column=1, padx=10, pady=10)

# Label com nome do arquivo
var_nome_arquivo = tk.StringVar()
label_nome_arquivo = customtkinter.CTkLabel(grupo_arq, textvariable=var_nome_arquivo)
label_nome_arquivo.grid(row=0, column=2, padx=10, pady=10)

# -------------- Grupo opções --------------
#ttk.Style().configure('Bold.TLabelframe.Label', font=('TkDefaulFont', 11, 'bold'))
grupo_opcoes = customtkinter.CTkFrame(
    root
)
grupo_opcoes.grid(row=2, column=0, padx=10, pady=10, rowspan=1, sticky='nsew')

# ---Checkbox para ch minima
bool_minima = tk.BooleanVar(value=True)
checkbox_minima = customtkinter.CTkCheckBox(
    grupo_opcoes, text="CH mínima: ", variable=bool_minima, command=verifica_check_boxes
)
checkbox_minima.grid(row=0, column=0, padx=10, pady=(10,0), sticky='w')
ToolTip(checkbox_minima, msg="Ativar restrição de carga horária média mínima", delay=0.1)

# campo texto
texto_ch_min = tk.DoubleVar(value=12.0)
texto_ch_min.trace("w", lambda *args: verifica_carga_horaria())
entrada_CH_MIN = customtkinter.CTkEntry(
    grupo_opcoes, textvariable=texto_ch_min, width=40,
)
entrada_CH_MIN.grid(row=0, column=1, padx=10, pady=(10,0), sticky='w')
ToolTip(entrada_CH_MIN,
    msg="Valor da carga horária mínima. Para valores não inteiros, use ponto decimal.", delay=0.1)

# Switch para escolher se a restrição é só geral ou também em cada unidade
label_texto_min = customtkinter.CTkLabel(
    grupo_opcoes, text="Essa restrição se aplica a"
)
label_texto_min.grid(row=0, column=2, pady=(10,0), sticky='w', columnspan=2)
label_duvida_min = customtkinter.CTkLabel(grupo_opcoes, text=" (?)")
label_duvida_min.grid(row=0, column=4, pady=(10,0), sticky='w')
ToolTip(label_duvida_min,
    msg="Escolhendo a opção 'total' o valor em cada unidade poderá extrapolar a restrição. " +
    "Nesse caso o mínimo por unidade será de 10 aulas (8 horas) por semana.",
    delay=0.1)

escopo_ch_min = tk.StringVar()
escopo_ch_min.set('total')
label_unidade_min = customtkinter.CTkLabel(grupo_opcoes, text="unidades")
opcao_min = customtkinter.CTkSwitch(
    grupo_opcoes, text="total", onvalue="total", offvalue="unidades", command=verifica_escopo
    , variable=escopo_ch_min
    , fg_color=("#3B8ED0", "#275e8a"), progress_color=("#3B8ED0", "#275e8a")
    , button_color=("#275e8a", "#3B8ED0") , button_hover_color=("#1d4566","#1d4566")
)
label_unidade_min.grid(row=1, column=2, padx=0, sticky='w')
opcao_min.grid(row=1, column=3, padx=(10,0), sticky='e')
opcao_min.select()

grupo_min = [entrada_CH_MIN, label_texto_min, label_duvida_min, label_unidade_min, opcao_min]

# Label para o erro
var_erro_minima = tk.StringVar()
label_erro_minima = customtkinter.CTkLabel(grupo_opcoes, textvariable=var_erro_minima)
label_erro_minima.grid(row=1, column=0, padx=10, sticky='w', columnspan=2)
label_erro_minima.grid_remove()

# ---Checkbox para ch maxima
bool_maxima = tk.BooleanVar(value=True)
checkbox_maxima = customtkinter.CTkCheckBox(
    grupo_opcoes, text="CH máxima: ", variable=bool_maxima, command=verifica_check_boxes
)
checkbox_maxima.grid(row=2, column=0, padx=10, pady=(10,0), sticky='w')
ToolTip(checkbox_maxima, msg="Ativar restrição de carga horária média máxima", delay=0.1)

# campo texto
texto_ch_max = tk.DoubleVar(value=16.0)
texto_ch_max.trace("w", lambda *args: verifica_carga_horaria())
entrada_CH_MAX = customtkinter.CTkEntry(
    grupo_opcoes, textvariable=texto_ch_max, width=40
)
entrada_CH_MAX.grid(row=2, column=1, padx=10, pady=(10,0), sticky='w')
ToolTip(entrada_CH_MAX,
    msg="Valor da carga horária máxima. Para valores não inteiros, use ponto decimal.", delay=0.1)

# Switch para escolher se a restrição é só geral ou também em cada unidade
label_texto_max = customtkinter.CTkLabel(
    grupo_opcoes, text="Essa restrição se aplica a"
)
label_texto_max.grid(row=2, column=2, pady=(10,0), sticky='w', columnspan=2)
label_duvida_max = customtkinter.CTkLabel(grupo_opcoes, text=" (?)")
label_duvida_max.grid(row=2, column=4, pady=(10,0), sticky='w')
ToolTip(label_duvida_max,
    msg="Escolhendo a opção 'total' o valor em cada unidade poderá extrapolar a restrição. " +
    "Nesse caso o mínimo por unidade será de 10 aulas (8 horas) por semana.",
    delay=0.1)

escopo_ch_max = tk.StringVar()
escopo_ch_max.set('total')
label_unidade_max = customtkinter.CTkLabel(grupo_opcoes, text="unidades")
opcao_max = customtkinter.CTkSwitch(
    grupo_opcoes, text="total", onvalue="total", offvalue="unidades", command=verifica_escopo
    , variable=escopo_ch_max
    , fg_color=("#3B8ED0", "#275e8a"), progress_color=("#3B8ED0", "#275e8a")
    , button_color=("#275e8a", "#3B8ED0") , button_hover_color=("#1d4566","#1d4566")
)
label_unidade_max.grid(row=3, column=2, padx=0, sticky='w')
opcao_max.grid(row=3, column=3, padx=(10,0), sticky='e')
opcao_max.select()

grupo_max = [entrada_CH_MAX, label_texto_max, label_duvida_max, label_unidade_max, opcao_max]

# Label para o erro
var_erro_maxima = tk.StringVar()
label_erro_maxima = customtkinter.CTkLabel(grupo_opcoes, textvariable=var_erro_maxima)
label_erro_maxima.grid(row=3, column=0, padx=10, sticky='w', columnspan=2)
label_erro_maxima.grid_remove()

# ---Checkbox para total mínimo
bool_min_total = tk.BooleanVar(value=False)
checkbox_min_total = customtkinter.CTkCheckBox(grupo_opcoes, text="Total mínimo: ",
    variable=bool_min_total, command=verifica_check_boxes)
checkbox_min_total.grid(row=4, column=0, padx=10, pady=10, sticky='w')

# campo texto
texto_min_total = tk.IntVar()
entrada_N_MIN_total = customtkinter.CTkEntry(grupo_opcoes, textvariable=texto_min_total, width=50)
texto_min_total.trace("w", lambda *args: verifica_totais())
entrada_N_MIN_total.grid(row=4, column=1, pady=10, sticky='w')

ToolTip(checkbox_min_total, msg="Ativar número mínimo total de professores", delay=0.1)
ToolTip(entrada_N_MIN_total, msg="Valor do mínimo total", delay=0.1)

# Label para o erro
var_erro_total_min = tk.StringVar()
label_erro_total_min = customtkinter.CTkLabel(grupo_opcoes, textvariable=var_erro_total_min)
label_erro_total_min.grid(row=4, column=2, padx=10, sticky='w', columnspan=3)

# ---Checkbox para total máximo
bool_max_total = tk.BooleanVar(value=False)
checkbox_max_total = customtkinter.CTkCheckBox(grupo_opcoes, text="Total máximo: ",
    variable=bool_max_total, command=verifica_check_boxes)
checkbox_max_total.grid(row=5, column=0, padx=10, pady=10, sticky='w')

# campo texto
texto_max_total = tk.IntVar()
entrada_N_MAX_total = customtkinter.CTkEntry(grupo_opcoes, textvariable=texto_max_total, width=50)
texto_max_total.trace("w", lambda *args: verifica_totais())
entrada_N_MAX_total.grid(row=5, column=1, pady=10, sticky='w')

ToolTip(checkbox_max_total, msg="Ativar número máximo total de professores", delay=0.1)
ToolTip(entrada_N_MAX_total, msg="Valor do máximo total", delay=0.1)

# Label para o erro
var_erro_total_max = tk.StringVar()
label_erro_total_max = customtkinter.CTkLabel(grupo_opcoes, textvariable=var_erro_total_max)
label_erro_total_max.grid(row=5, column=2, padx=10, sticky='w', columnspan=3)

# ---Limitações nos perfis---
botao_perfil = customtkinter.CTkButton(
    grupo_opcoes, text="Limitar perfis", command=abrir_janela)
botao_perfil.grid(row=6, column=0, padx=10, pady=10, sticky='n')

# Frame para a lista de limitações
frame_perfis = customtkinter.CTkFrame(grupo_opcoes, height=0)
frame_perfis.grid(row=6, column=1, padx=10, pady=10, columnspan=4, sticky='w')

# Combobox critério
label_criterio = customtkinter.CTkLabel(grupo_opcoes, text="Critério:")
label_criterio.grid(row=7, column=0, padx=10, pady=10, sticky='e')

combo_var = tk.StringVar()
combobox = customtkinter.CTkOptionMenu(
    grupo_opcoes, variable=combo_var, values=list(LISTA_MODOS.keys()),
    command=lambda *args: verifica_executar()
)
combobox.grid(row=7, column=1, padx=10, pady=10, columnspan=3)

# Combo solver
label_solver = customtkinter.CTkLabel(grupo_opcoes, text="Escolha o solver:")
label_solver.grid(row=8, column=0, padx=10, pady=10, sticky='e')

solver_var = tk.StringVar()
combo_solver = customtkinter.CTkOptionMenu(
    grupo_opcoes, variable=solver_var, values=list(['CBC', 'SCIP']),
    command=lambda *args: verifica_executar()
)
combo_solver.grid(row=8, column=1, padx=10, pady=10, columnspan=3)

# Tempo limite
val_limite = tk.IntVar(value=30)
label_limite = customtkinter.CTkLabel(grupo_opcoes, text="Tempo limite:")
label_limite.grid(row=9, column=0, padx=10, pady=10)
entrada_tempo_limite = customtkinter.CTkEntry(grupo_opcoes, textvariable=val_limite, width=40)
entrada_tempo_limite.grid(row=9, column=1, padx=10, pady=10)

ToolTip(label_limite, msg="Tempo máximo para procurar a solução ótima", delay=0.1)

# Botão para executar
botao_executar = customtkinter.CTkButton(
    grupo_opcoes, text="Executar", state="disabled", command=executar)
botao_executar.grid(row=10, column=0, padx=10, pady=10)

# Inicialmente oculta as opções
grupo_opcoes.grid_remove()

# -------------- Grupo dos resultados --------------
grupo_resultados = customtkinter.CTkScrollableFrame(
    root
    , label_text="Resultados"
)
grupo_resultados.grid(row=1, column=1, padx=10, pady=10, rowspan=2, sticky='nsew')

resultado = tk.StringVar()
label_resultado = customtkinter.CTkLabel(
    grupo_resultados, textvariable=resultado, anchor="w", justify="left"
)
label_resultado.grid(row=0, column=0, padx=10, pady=10, sticky='w')

label_aba = customtkinter.CTkLabel(grupo_resultados, text="Distribuição:")
label_aba.grid(row=1, column=0, padx=10, pady=(10,0), sticky='w')

# Adicionando a barra de rolagem vertical
#scrollbar = tk.Scrollbar(grupo_resultados, orient="vertical")

fonte_tabela = customtkinter.CTkFont(family='Consolas', size=14)
text_tabela = customtkinter.CTkTextbox(
    grupo_resultados, height=10, width=60, #yscrollcommand=scrollbar.set,
    font=fonte_tabela
)
text_tabela.grid(row=2, column=0, padx=10, pady=0)

# Configurando a barra de rolagem para rolar o texto no Text widget
#scrollbar.config(command=text_tabela.yview)

# --- Frame com botões ---
grupo_botoes = customtkinter.CTkFrame(root)
grupo_botoes.grid(row=3, column=1, padx=10, pady=(0,10), sticky='sew')
botao_relatorio = customtkinter.CTkButton(
    grupo_botoes, text="Baixar relatório", command=exportar_txt
)
botao_relatorio.grid(row=0, column=0, padx=10, pady=10, sticky='w')

botao_planilha = customtkinter.CTkButton(
    grupo_botoes, text="Baixar planilha", command=exportar_planilha
)
botao_planilha.grid(row=0, column=1, padx=10, pady=10, sticky='e')

# Inicialmente oculta os resultados
grupo_resultados.grid_remove()
grupo_botoes.grid_remove()

root.mainloop()
