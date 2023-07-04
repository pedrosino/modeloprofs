# Modelo de programação linear para distribuição de vagas entre unidades acadêmicas de uma universidade federal
# Desenvolvido na pesquisa do Mestrado Profissional em Gestão Organizacional, da Faculdade de Gestão e Negócios, Universidade Federal de Uberlândia
# por Pedro Santos Guimarães, em 2023

from pulp import *
import numpy as np
from datetime import time, datetime
import sys, getopt
import pandas as pd
import math

#Configurações
# Restrição de carga horária média máxima por unidade
maxima = False 
ch_min = 12
# Restrição de carga horária média mínima por unidade
minima = False
ch_max = 16
max_total = False  # Restrição do número máximo total
min_total = False # Restrição do número mínimo total
# Restrição quanto aos perfis 5 e 6 (podem ser considerados como regime de 40h)
quarenta = False
p_quarenta = 0.1 # porcentagem máxima aceita
# Restrição quanto aos perfis 7 e 8 - regime de 20h
vinte = False
p_vinte = 0.2 # porcentagem máxima aceita
time_limit = 30 # Tempo limite para o programa procurar a solução ótima (em segundos) - esse parâmetro pode ser alterado pelo usuário ao chamar o programa
filein = 'importar.xlsx' # Arquivo com os dados para importação - esse parâmetro pode ser alterado pelo usuário ao chamar o programa

# Opções passadas na linha de comando
opts, args = getopt.getopt(sys.argv[1:],"piavqn:",["chmin=","chmax=","tmax=","tmin=","limite=","input=","pv=","pq="])
for opt, arg in opts:
    if opt == '-i': # Ativa restrição da carga horária mínima com valor padrão
        minima = True
    elif opt == '-a': # Ativa restrição da carga horária máxima com valor padrão
        maxima = True
    elif opt == '-t':
        max_total = True
        n_max_total = int(arg)
    elif opt == '-n':
        n_unidades = int(arg)
    elif opt == '-v':
        vinte = True
    elif opt == '--pv':
        vinte = True
        p_vinte = float(arg)
    elif opt == '-q':
        quarenta = True
    elif opt == 'pq':
        quarenta = True
        p_quarenta = float(arg)
    elif opt == '--chmin': # Define novo valor da carga horária mínima
        minima = True
        ch_min = int(arg)
    elif opt == '--chmax': # Define novo valor da carga horária máxima
        maxima = True
        ch_max = int(arg)
    elif opt == '--tmin': # Define quantidade total mínima
        min_total = True
        n_min_total = int(arg)
    elif opt == '--tmax': # Define quantidade total máxima
        max_total = True
        n_max_total = int(arg)
    elif opt == '--limite':
        time_limit = int(arg)
    elif opt == '--input':
        filein = arg

# Importa dados do arquivo        
df_todas = pd.read_excel(filein, sheet_name=['unidades','perfis'])
m_unidades = df_todas['unidades'].to_numpy()
m_perfis = df_todas['perfis'].to_numpy()
m_perfis = np.delete(m_perfis, 0, axis=1)

# Dados do problema
n_perfis = 8
n_restricoes = 5
n_unidades = len(m_unidades)
# Foi usado o AHP para determinar os coeficientes de cada objetivo/critério no modelo geral:
# 1) Número absoluto total:             0,1362  (aprox)   0,1336 (reais)
# 2) Professor-Equivalente total:       0,0626            0,0614
# 3) Tempo disponível:                  0,3093            0,3102
# 4) Equilíbrio na carga horária média: 0,4919            0,4948
# 
#pesos = np.array([0.1362, 0.0626, 0.3093, 0.4919])
pesos = np.array([0.1336, 0.0614, 0.3102, 0.4948])

##print("As opções escolhidas são:")
##print(f"Unidades: {n_unidades}, Max:{maxima} {ch_max if maxima else ''}, Min:{minima} {ch_min if minima else ''}, PEq:{peq}, TotalMax:{max_total} {n_max_total if max_total else ''}, TotalMin:{min_total} {n_min_total if min_total else ''}, 40h:{quarenta}, 20h{vinte}")
print("Bem vindo!")
print("Primeiramente vamos definir os parâmetros para cada critério/objetivo")

# Dá opção de alterar os parâmetros
'''tecla = input("Aperte enter para continuar ou a para alterar\n")
if(tecla=='a'):
    print("--chmin define o valor da carga horária mínima. Exemplo: --chmin 11")
    print("--chmax define o valor da carga horária máxima. Exemplo: --chmax 15")
    print("--tmax ativa a quantidade máxima total de professores. Deve ser seguido do número desejado. Exemplo: --tmax 100")
    print("--tmin ativa a quantidade mínima total de professores. Deve ser seguido do número desejado. Exemplo: --tmin 50")
    print("-n define o número de unidades. Exemplo: -n 4")
    print("-v ativa a limitação de professores em regime de 20 horas (20%)")
    print("-q ativa a limitação de professores em regime de 40 horas (10%)")
    opcoes = input("Digite as opções desejadas, separadas por espaço. Exemplo: --tmax 130 -v --chmax 15")
'''    

# Custos
##matriz_peq = np.array([1.65, 1.65, 1.65, 1.65, 1, 1, 0.6, 0.6]) #professor-equivalente - x1 a x4 = DE, x5 e x6 = 40h, x7 e x8 = 20h
matriz_peq = np.array([1.65, 1.65, 1.65, 1.65, 1.65, 1.65, 0.6, 0.6]) #professor-equivalente - x1 a x6 = DE, x7 e x8 = 20h
matriz_tempo = np.array([0, 0, 18, 12, 5, 0, 3, 0]) # tempo disponível

# Definições das restrições                              
conectores = np.array([">=", ">=", ">=", "==", "=="])
nomes_restricoes = np.array(['aulas','h_orientacoes','n_orientacoes','diretor','coords'])
 
# Variáveis de decisão
#nomes
nomes = [str(p) + "." + m_unidades[u][0] for u in range(n_unidades) for p in range(1, n_perfis+1) ]
#variaveis
nx = LpVariable.matrix("x", nomes, cat="Integer", lowBound=0)
saida = np.array(nx).reshape(n_unidades, n_perfis)

## Percorre os critérios 
modos = np.array(['num', 'peq', 'tempo', 'ch'])
modelos = {}
melhores = {}
piores = {}

# Valores conhecidos a priori
# número máximo é dado pelo total de aulas dividido pela carga horária mínima ou é definido pelo usuário
#
# Para avaliar cada cenário em relação a esses critérios, é necessário estabelecer uma forma de pontuação
# Essa pontuação varia entre 0 (pior caso) e 1 (melhor caso), e será multiplicada pelo peso de cada critério
# para obter a pontuação total daquele cenário. Para cada critério é necessário então determinar o pior e 
# o melhor caso. Nos dois primeiros o pior caso é quando o total de professores é número máximo possível,
# com base na carga horária média mínima. Por exemplo, se são 900 aulas e a carga horária mínim foi definida
# em 12 aulas por professor, o número máximo possível é 75. Para o critério 1 esse é o valor a ser considerado.
# Para o critério 2, o pior valor seria 75*1,65, que é o fator do professor 40h-DE.
# Já o melhor caso não pode ser determinado a priori, pois o número de professores deve atender às restrições
# de aulas e orientações, por exemplo. Assim, é necessário resolver o modelo com cada critério e registrar o
# valor ótimo obtido para ser a base da escala. Para esses dois critérios a escala é invertida, ou seja, quanto
# mais professores, menor a pontuação.
# No caso do critério 3 a lógica é inversa, quanto mais tempo disponível, melhor o cenário. O pior caso é quando
# não há tempo nenhum (valor 0), e o melhor caso deve ser determinado resolvendo o modelo usando esse critério.
# Para o critério 4, o melhor caso é determinado pela solução inicial do modelo com este critério. O pior caso é
# quando metade das unidades estiver na carga horária máxima e a outra metade na mímina. A média geral seria
# o valor intermediário entre as duas e o desvio total é dado por n_unidades*(ch_max - ch_min)/2
# Exemplo:
# 
# ch_max      = 16  --   d1 = d2 = (ch_max - ch_min)/2
#                   |
#                   d1   desvio total = n_unidades/2*d1 + n_unidades/2*d2
#                   |
# média geral = 14  --   desvio total = n_unidades*d1 = n_unidades*(ch_max - ch_min)/2
#                   | 
#                   d2
#                   |
# ch_min      = 12  --
numero_max = round(np.sum(m_unidades[:n_unidades], axis=0)[1] / ch_min) if not max_total else n_max_total
piores['num'] = numero_max
piores['peq'] = numero_max*1.65
piores['tempo'] = 0
piores['ch'] = (ch_max - ch_min)/2

# Imprime os resultados no arquivo de texto
original_stdout = sys.stdout
filename = f"CBC completoN={n_unidades} {datetime.now().strftime('%H-%M-%S')}.txt"
fileout =  open(filename, 'w')
sys.stdout = fileout

for modo in modos:
    sys.stdout = original_stdout
    print(f"Modo: {modo}")
    sys.stdout = fileout
    # Sempre ativa carga horária mínima e máxima, para que todos os modos tenham as mesmas restrições
    minima = True
    maxima = True
    
    # -- Definir o modelo --
    if(modo == 'tempo'):
        modelos[modo] = LpProblem(name=f"Professores-{modo}", sense=LpMaximize)
    else:
        modelos[modo] = LpProblem(name=f"Professores-{modo}", sense=LpMinimize)
        
    # -- Restrições --
    # No caso do modo de equilíbrio da carga horária não é necessário considerar essas restrições, pois apenas o número total por unidade é levado em conta
    if(modo != 'ch'):
        for r in range(n_restricoes):
            for u in range(n_unidades):
                match conectores[r]:
                    case ">=":
                        modelos[modo] += lpDot(saida[u], m_perfis[r]) >= m_unidades[u][r+1], nomes_restricoes[r] + " " + m_unidades[u][0]#"r " + str(i)
                    case "==":
                        modelos[modo] += lpDot(saida[u], m_perfis[r]) == m_unidades[u][r+1], nomes_restricoes[r] + " " + m_unidades[u][0]#"r " + str(i)
                    case "<=":
                        modelos[modo] += lpDot(saida[u], m_perfis[r]) <= m_unidades[u][r+1], nomes_restricoes[r] + " " + m_unidades[u][0]#"r " + str(i)        

        # Restrições de regimes de trabalho
        if(quarenta):
            for u in range(n_unidades):
                modelos[modo] += saida[u][4] + saida[u][5] - p_quarenta*lpSum(saida[u]) <= 0, f"40 horas {m_unidades[u][0]} <= 10%"
        if(vinte):
            for u in range(n_unidades):
                modelos[modo] += saida[u][6] + saida[u][7] - p_vinte*lpSum(saida[u]) <= 0, f"20 horas {m_unidades[u][0]} <= 20%"
        
    # Restrições de máximo e mínimo por unidade -> carga horária média
    # ------------------
    # Exemplo para 900 aulas, considerando-se carga horária média mínima de 12 aulas e máxima de 16
    # soma <= 900/12 -> soma <= 75
    # soma >= 900/16 -> soma >= 56.25
    # a soma deve estar entre 57 e 75, inclusive
    for u in range(n_unidades):
        if(maxima):# or modo == 'tempo'):
            modelos[modo] += ch_max*lpSum(saida[u]) >= m_unidades[u][1], f"{m_unidades[u][0]}_chmax: {math.ceil(m_unidades[u][1]/ch_max)}"
        # Restrição mínima somente no geral -> permite exceções como o IBTEC em Monte Carmelo, que tem 19 aulas apenas
        #if(minima):
        #    modelos[modo] += ch_min*lpSum(saida[u]) <= m_unidades[u][1], f"{m_unidades[u][0]}_chmin: {int(m_unidades[u][1]/ch_min)}"
    
    # Restrições no geral
    if(maxima):
        modelos[modo] += ch_max*lpSum(saida) >= np.sum(m_unidades[:n_unidades], axis=0)[1], f"chmax: {math.ceil(np.sum(m_unidades[:n_unidades], axis=0)[1]/ch_max)}"
    if(minima):
        modelos[modo] += ch_min*lpSum(saida) <= np.sum(m_unidades[:n_unidades], axis=0)[1], f"chmin: {int(np.sum(m_unidades[:n_unidades], axis=0)[1]/ch_min)}"

    # Restrição do número total de professores
    if(max_total):
        modelos[modo] += lpSum(saida) <= n_max_total, f"TotalMax: {n_max_total}"
    if(min_total):
        modelos[modo] += lpSum(saida) >= n_min_total, f"TotalMin: {n_min_total}"

    # Modo de equilíbrio da carga horária
    if(modo == 'ch'):            
        # variáveis auxiliares
        # Para cada unidade há um valor da média e um desvio em relação à média geral
        desvios = LpVariable.matrix("zd", m_unidades[:n_unidades, 0], cat="Continuous")
        medias = LpVariable.matrix("zm", m_unidades[:n_unidades, 0], cat="Continuous")
        media_geral = LpVariable("zmg", cat="Continuous")
        
        # Como a média seria uma função não linear (aulas/professores), foi feita uma aproximação com uma reta que passa pelos dois pontos extremos
        # dados pelos valores de carga horária mínima e máxima
        # Coeficientes
        # m = (a-b)/(c-d) -> a = media minima, b = media maxima, c = numero maximo, d = numero minimo
        c = np.sum(m_unidades[:n_unidades], axis=0)[1] / ch_min
        d = np.sum(m_unidades[:n_unidades], axis=0)[1] / ch_max
        m = (ch_min - ch_max) / (c - d)
        
        # O cálculo da média é inserido no modelo como uma restrição
        modelos[modo] += media_geral == m*lpSum(saida[:n_unidades]) + ch_min + ch_max, "Media geral"
        
        for u in range(n_unidades):        
            c = m_unidades[u][1] / ch_min
            d = m_unidades[u][1] / ch_max
            m = (ch_min - ch_max) / (c - d)
            # Cálculo da media
            modelos[modo] += medias[u] == m*lpSum(saida[u]) + ch_min + ch_max, f"{m_unidades[u][0]}_ch_media"
            # Cálculo do desvio
            # O desvio é dado pelo módulo da subtração, porém o PuLP não aceita a função abs()
            # Assim, são colocadas duas restrições, uma usando o valor positivo e outra o negativo
            modelos[modo] += desvios[u] >= medias[u] - media_geral, f"{m_unidades[u][0]}_up"
            modelos[modo] += desvios[u] >= -1*(medias[u] - media_geral), f"{m_unidades[u][0]}_low"
            
    # -- Função objetivo --
    if(modo == 'num'):
        modelos[modo] += lpSum(saida)
    elif(modo == 'peq'):
        modelos[modo] += lpSum(saida*matriz_peq)
    # No modo tempo é necessário deduzir do tempo calculo as horas que serão destinadas às orientações
    elif(modo == 'tempo'):
        modelos[modo] += lpSum(saida*matriz_tempo) - np.sum(m_unidades[:n_unidades], axis=0)[2]
    # No modo de equilíbrio a medida de desempenho é o desvio médio
    elif(modo == 'ch'):
        modelos[modo] += lpSum(desvios[:n_unidades])/n_unidades
        
    # Imprime os resultados no arquivo txt
    print(f'Modo: {modo}')
    print(f'Unidades: {n_unidades}')
    print(f'CH Maxima: {maxima} {ch_max if maxima else ""}')
    print(f'CH Minima: {minima} {ch_min if minima else ""}')
    print(f'Total: {n_min_total if min_total else "-"} a {n_max_total if max_total else "-"}')
    print(f'Vinte: {vinte}')
    print(f'Quarenta: {quarenta}')
    print()

    # Resolver o modelo
    status = modelos[modo].solve(PULP_CBC_CMD(msg=0, timeLimit=time_limit))#time_limit))
    
    # Resultados
    print(f"Situação: {modelos[modo].status}, {LpStatus[modelos[modo].status]}")
    # Para cada critério o resultado é em um formato diferente
    if(modo == 'ch'):
        objetivo = modelos[modo].objective.value()
    elif(modo == 'peq'):
        objetivo = round(modelos[modo].objective.value(), 2)
    elif(modo == 'num'):
        objetivo = int(modelos[modo].objective.value())
    elif(modo == 'tempo'):
        objetivo = modelos[modo].objective.value()
    
    print(f"Objetivo: {objetivo} {'horas' if modo == 'tempo' else 'Prof-Equivalente' if modo == 'peq' else 'Professores' if modo == 'num' else 'aulas/prof'}")
    print(f"Resolvido em {modelos[modo].solutionTime} segundos")
    print("")
    
    melhores[modo] = objetivo
    sys.stdout = original_stdout
    print(f"Ok. Resultado: {melhores[modo]}")
    sys.stdout = fileout
    
    # Formata resultados e calcula totais
    qtdes = np.full((n_unidades, n_perfis), 0, dtype=int)
    index = 0
    perfil = 0
    for var in modelos[modo].variables():
        if(var.name.find('x') == 0):
            valor = int(var.value())
            #resultados[index] += f"{valor:2d} "
            qtdes[index][perfil] = valor
            index += 1
            if(index >= n_unidades):
                index = 0
                perfil += 1
            if(perfil >= n_perfis+1):
                print(f"i: {index}, p: {perfil}")
                break
    
    print("Resultados:")
    print(f"--------+---------------------------------+-------+--------+---------+------------+")
    print(f"Unidade |  x1  x2  x3  x4  x5  x6  x7  x8 | total |   P-Eq |   Tempo | Tempo/prof |")
    print(f"--------+---------------------------------+-------+--------+---------+------------+")
    for u in range(n_unidades):
        print(f"{m_unidades[u][0]:5s}   | " + " ".join([f"{qtdes[u][p]:3d}" for p in range(n_perfis)])
        #           Total                      P-Eq                                    Tempo                 - horas de orientação                                     Tempo/prof
        + f" |  {np.sum(qtdes[u]):4d} | {np.sum(qtdes[u]*matriz_peq):6.2f} |  {np.sum(qtdes[u]*matriz_tempo) - m_unidades[u][2]:6.1f} |    {(np.sum(qtdes[u]*matriz_tempo) - m_unidades[u][2])/np.sum(qtdes[u]):7.3f} |")
    print(f"--------+---------------------------------+-------+--------+---------+------------+")
    print(f"Total   | " + " ".join([f"{np.sum(qtdes, axis=0)[p]:3d}" for p in range(n_perfis)])
    #            Total                      P-Eq                                     Tempo      -  horas de orientação                               Tempo/prof
    + f" |  {np.sum(qtdes):4d} | {np.sum(qtdes*matriz_peq):6.2f} |  {np.sum(qtdes*matriz_tempo) - np.sum(m_unidades[:n_unidades], axis=0)[2]:6.1f} |    {(np.sum(qtdes*matriz_tempo) - np.sum(m_unidades[:n_unidades], axis=0)[2])/np.sum(qtdes):7.3f} |")
    print(f"--------+---------------------------------+-------+--------+---------+------------+")
    print("")
    print("")
    print("---------------------------")
    
##### -----------------Fim da primeira 'rodada'-----------------
print("Melhores:")
print(melhores)
print("Piores:")
print(piores)
print()
print("---------------------------")

## Agora uma nova rodada do modelo usando os pesos
# Verifica modo de carga horária e ativa máximo e mínimo
#if(modo == 'ch'):
minima = True
maxima = True

# -- Definir o modelo --
modelo = LpProblem(name=f"Professores-Geral", sense=LpMaximize)

# Variáveis auxiliares
desvios = LpVariable.matrix("zd", m_unidades[:n_unidades, 0], cat="Continuous", lowBound=0)
medias = LpVariable.matrix("zm", m_unidades[:n_unidades, 0], cat="Continuous", lowBound=0)
media_geral = LpVariable("zmg", cat="Continuous", lowBound=0)
    
# Restrições
for r in range(n_restricoes):
    for u in range(n_unidades):
        match conectores[r]:
            case ">=":
                modelo += lpDot(saida[u], m_perfis[r]) >= m_unidades[u][r+1], nomes_restricoes[r] + " " + m_unidades[u][0]#"r " + str(i)
            case "==":
                modelo += lpDot(saida[u], m_perfis[r]) == m_unidades[u][r+1], nomes_restricoes[r] + " " + m_unidades[u][0]#"r " + str(i)
            case "<=":
                modelo += lpDot(saida[u], m_perfis[r]) <= m_unidades[u][r+1], nomes_restricoes[r] + " " + m_unidades[u][0]#"r " + str(i)        

# Restrições de máximo e mínimo por unidade -> carga horária média
# ------------------
# Considerou-se carga horária média mínima de 12 aulas e máxima de 16
# Exemplo para 900 aulas
# soma <= 900/12 -> soma <= 75
# soma >= 900/16 -> soma >= 56.25
# a soma deve estar entre 57 e 75, inclusive
for u in range(n_unidades):
    if(maxima):
        modelo += ch_max*lpSum(saida[u]) >= m_unidades[u][1], f"{m_unidades[u][0]}_chmax: {math.ceil(m_unidades[u][1]/ch_max)}"
    
if(maxima):
    modelo += ch_max*lpSum(saida) >= np.sum(m_unidades[:n_unidades], axis=0)[1], f"chmax: {math.ceil(np.sum(m_unidades[:n_unidades], axis=0)[1]/ch_max)}"
if(minima):
    modelo += ch_min*lpSum(saida) <= np.sum(m_unidades[:n_unidades], axis=0)[1], f"chmin: {int(np.sum(m_unidades[:n_unidades], axis=0)[1]/ch_min)}"

# Restrição do número total de professores
if(max_total):
    modelo += lpSum(saida) <= n_max_total, f"TotalMax: {n_max_total}"
if(min_total):
    modelo += lpSum(saida) >= n_min_total, f"TotalMin: {n_min_total}"

# Restrições de regimes de trabalho
if(quarenta):
    for u in range(n_unidades):
        modelo += saida[u][4] + saida[u][5] - p_quarenta*lpSum(saida[u]) <= 0, f"40 horas {m_unidades[u][0]} <= 10%"
if(vinte):
    for u in range(n_unidades):
        modelo += saida[u][6] + saida[u][7] - p_vinte*lpSum(saida[u]) <= 0, f"20 horas {m_unidades[u][0]} <= 20%"

# coeficientes
# m = (a-b)/(c-d) -> a = media minima, b = media maxima, c = numero maximo, d = numero minimo
c = np.sum(m_unidades[:n_unidades], axis=0)[1] / ch_min
d = np.sum(m_unidades[:n_unidades], axis=0)[1] / ch_max
m = (ch_min - ch_max) / (c - d)

modelo += media_geral == m*lpSum(saida[:n_unidades]) + ch_min + ch_max, "Media geral"

for u in range(n_unidades):        
    c = m_unidades[u][1] / ch_min
    d = m_unidades[u][1] / ch_max
    m = (ch_min - ch_max) / (c - d)
    # calculo da media
    modelo += medias[u] == m*lpSum(saida[u]) + ch_min + ch_max, f"{m_unidades[u][0]}_ch_media"
    # desvio
    modelo += desvios[u] >= medias[u] - media_geral, f"{m_unidades[u][0]}_up"
    modelo += desvios[u] >= -1*(medias[u] - media_geral), f"{m_unidades[u][0]}_low"

# Pontuações
pontuacoes = LpVariable.matrix("p", range(4), cat="Continuous", lowBound=0, upBound=1)
# Restrições/cálculos
# caso seja dado um número exato, é necessário alterar a pontuação do critério 'num' para evitar a divisão por zero
if(max_total and min_total and n_max_total == n_min_total):
    modelo += pontuacoes[0] == lpSum(saida)/melhores['num'], "Pontuação número"
else:
    modelo += pontuacoes[0] == (lpSum(saida) - piores['num'])/(melhores['num'] - piores['num']), "Pontuação número"
modelo += pontuacoes[1] == (lpSum(saida*matriz_peq) - piores['peq'])/(melhores['peq'] - piores['peq']), "Pontuação P-Eq"
modelo += pontuacoes[2] == (lpSum(saida*matriz_tempo) - np.sum(m_unidades[:n_unidades], axis=0)[2] - piores['tempo'])/(melhores['tempo'] - piores['tempo']), "Pontuação tempo"
modelo += pontuacoes[3] == (lpSum(desvios[:n_unidades])/n_unidades - piores['ch'])/(melhores['ch'] - piores['ch']), "Pontuação Equilíbrio"

# -- Função objetivo --
# O objetivo é maximimizar a pontuação dada pela soma de cada pontuação multiplicada por seu peso
modelo += lpSum(pontuacoes*pesos)
   
# Imprime os resultados
print(f'Modo: Geral')
print(f'Unidades: {n_unidades}')
print(f'CH Maxima: {maxima} {ch_max if maxima else ""}')
print(f'CH Minima: {minima} {ch_min if minima else ""}')
print(f'Total: {n_min_total if min_total else "-"} a {n_max_total if max_total else "-"}')
print(f'Vinte: {vinte}')
print(f'Quarenta: {quarenta}')
print()

# Resolver o modelo
status = modelo.solve(PULP_CBC_CMD(msg=0, timeLimit=time_limit))

# Resultados
print(f"Situação: {modelo.status}, {LpStatus[modelo.status]}")
objetivo = f"{modelo.objective.value():.4f}"
print(f"Objetivo: {objetivo}")
print(f"Resolvido em {modelo.solutionTime} segundos")
print("")

# Formata resultados e calcula totais
resultados = np.full(n_unidades, '', dtype=object)
#totais = np.full(n_unidades, 0, dtype=int)
qtdes = np.full((n_unidades, n_perfis), 0, dtype=int)
index = 0
perfil = 0
for var in modelo.variables():
    if(var.name.find('x') == 0):
        valor = int(var.value())
        resultados[index] += f"{valor:2d} "
        qtdes[index][perfil] = valor
        #totais[index] += valor
        index += 1
        if(index >= n_unidades):
            index = 0
            perfil += 1
        if(perfil >= n_perfis+1):
            print(f"i: {index}, p: {perfil}")
            break

print("Resultados:")
print(f"--------+---------------------------------+-------+--------+---------+------------+")
print(f"Unidade |  x1  x2  x3  x4  x5  x6  x7  x8 | total |   P-Eq |   Tempo | Tempo/prof |")
print(f"--------+---------------------------------+-------+--------+---------+------------+")
for u in range(n_unidades):
    print(f"{m_unidades[u][0]:5s}   | " + " ".join([f"{qtdes[u][p]:3d}" for p in range(n_perfis)])
    #           Total                      P-Eq                                    Tempo                     - horas de orientação                                     Tempo/prof
    + f" |  {np.sum(qtdes[u]):4d} | {np.sum(qtdes[u]*matriz_peq):6.2f} |  {np.sum(qtdes[u]*matriz_tempo) - m_unidades[u][2]:6.1f} |    {(np.sum(qtdes[u]*matriz_tempo) - m_unidades[u][2])/np.sum(qtdes[u]):7.3f} |")
print(f"--------+---------------------------------+-------+--------+---------+------------+")
print(f"Total   | " + " ".join([f"{np.sum(qtdes, axis=0)[p]:3d}" for p in range(n_perfis)])
#            Total                      P-Eq                                     Tempo          -   horas de orientação                               Tempo/prof
+ f" |  {np.sum(qtdes):4d} | {np.sum(qtdes*matriz_peq):6.2f} |  {np.sum(qtdes*matriz_tempo) - np.sum(m_unidades[:n_unidades], axis=0)[2]:6.1f} |    {(np.sum(qtdes*matriz_tempo) - np.sum(m_unidades[:n_unidades], axis=0)[2])/np.sum(qtdes):7.3f} |")
print(f"--------+---------------------------------+-------+--------+---------+------------+")
print("")

print("Parâmetros:")
print(f"--------+-------------+--------------------+--------------+-----------+-----------+---------+---------+----------+")
print(f"Unidade |      aulas  |     horas_orient   |  num_orient  |   diretor |   coords. |   40h   |   20h   | ch media |")
print(f"--------+-------------+--------------------+--------------+-----------+-----------+---------+---------+----------+")
# Formatos dos números
formatos =          ['4d', '7.2f', '4d', '3d', '3d']
formatosdiferenca = ['3d', '7.2f', '4d', '2d', '2d']
for u in range(n_unidades):
    print(f"{m_unidades[u][0]:5s}   | " 
    #+ "  ".join([f"{np.sum(qtdes[u]*m_perfis[p]):{'4d' if p != 1 else '6.1f'}} (+{(np.sum(qtdes[u]*m_perfis[p]) - m_unidades[u][p+1]):{'3d' if p != 1 else '6.1f'}}) |" for p in range(n_restricoes)])
    + " ".join([f"{np.sum(qtdes[u]*m_perfis[p]):{formatos[p]}} (+{(np.sum(qtdes[u]*m_perfis[p]) - m_unidades[u][p+1]):{formatosdiferenca[p]}}) |" for p in range(n_restricoes)])
    + f"  {((qtdes[u][4] + qtdes[u][5]) / np.sum(qtdes[u]))*100:5.2f}% |" #40h
    + f"  {((qtdes[u][6] + qtdes[u][7]) / np.sum(qtdes[u]))*100:5.2f}% |" #20h
    + f"  {m_unidades[u][1] / np.sum(qtdes[u]):7.3f} |"
    )
print(f"--------+-------------+--------------------+--------------+-----------+-----------+---------+---------+----------+")
print(f"Total   | "
#+ " ".join([f"{np.sum(qtdes*m_perfis[p]):{'4d' if p != 1 else '6.1f'}} (+{np.sum(qtdes*m_perfis[p]) - int(np.sum(m_unidades, axis=0)[p+1]):{'3d' if p != 1 else '6.1f'}}) |" for p in range(n_restricoes)])
+ " ".join([f"{np.sum(qtdes*m_perfis[p]):{formatos[p]}} (+{np.sum(qtdes*m_perfis[p]) - int(np.sum(m_unidades, axis=0)[p+1]):{formatosdiferenca[p]}}) |" for p in range(n_restricoes)])
+ f"  {(np.sum(qtdes, axis=0)[4] + np.sum(qtdes, axis=0)[5]) / np.sum(qtdes)*100:5.2f}% |"
+ f"  {(np.sum(qtdes, axis=0)[6] + np.sum(qtdes, axis=0)[7]) / np.sum(qtdes)*100:5.2f}% |"
+ f"  {np.sum(m_unidades, axis=0)[1]/np.sum(qtdes):7.3f} |"
)
print(f"--------+-------------+--------------------+--------------+-----------+-----------+---------+---------+----------+")
print()

# Imprime médias e desvios
for var in modelo.variables():
    if(var.name[0] == 'p' or var.name[0] == 'z'):
        print(f"{var.name}: {var.value()}")

# Imprime o modelo completo
print("")
print("------------------Modelo:------------------")
print(modelo)

# Fecha arquivo texto
fileout.close()

#Imprime na tela    
sys.stdout = original_stdout
print(f"Situação: {modelo.status}, {LpStatus[modelo.status]}")
objetivo = f"{modelo.objective.value():.4f}"
print(f"Objetivo: {objetivo} (na escala de 0 a 1)")
print(f"Resolvido em {modelo.solutionTime} segundos")
print("")
print(f"Verifique o arquivo {filename} para o relatório completo")
print("")

#salva em planilha
df = pd.DataFrame(qtdes, columns=['x1','x2','x3','x4','x5','x6','x7','x8'])
df.insert(0, "Unidade", m_unidades[:, 0])
df.to_excel(f'CBC_CompletoOrient.xlsx', sheet_name='Resultados', index=False)
