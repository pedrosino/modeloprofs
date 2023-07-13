# Foi usado o AHP para determinar os coeficientes de cada objetivo/critério no modelo geral:
# 1) Número absoluto total:             0,1362  (aprox)   0,1336 (reais)
# 2) Professor-Equivalente total:       0,0626            0,0614
# 3) Tempo disponível:                  0,3093            0,3102
# 4) Equilíbrio na carga horária média: 0,4919            0,4948
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

from pulp import *
import numpy as np
from datetime import time, datetime
import sys, getopt
import pandas as pd
import math

#Configurações
maxima = False #carga horária média máxima por unidade 
ch_min = 12
minima = False #carga horária média mínima por unidade
ch_max = 16
peq = False    #usar Professor-Equivalente
max_total = False  #número máximo total
min_total = False #número mínimo total
dif = False    #limitar diferença da meta
dif_tam = 50   #percentual aceito
modo = 'num'   #modo 'num' = minimizar quantidade ou P-Eq; modo 'tempo' = maximizar tempo disponível; modo 'ch' = minimizar desvio da carga horária média
quarenta = False
p_quarenta = 0.1
vinte = False
p_vinte = 0.2
time_limit = 30
filein = 'importar.xlsx'

opts, args = getopt.getopt(sys.argv[1:],"piadvqn:m:",["chmin=","chmax=","tmax=","tmin=","limite=","input="])
for opt, arg in opts:
    if opt == '-p':
        peq = True
    elif opt == '-i':
        minima = True
    elif opt == '-a':
        maxima = True
    elif opt == '-t':
        max_total = True
        n_max_total = int(arg)
    elif opt == '-n':
        n_unidades = int(arg)
    #elif opt == '-m':
        #modo = arg
    elif opt == '-d':
        dif = True
    elif opt == '-v':
        vinte = True
    elif opt == '-q':
        quarenta = True
    elif opt == '--chmin':
        minima = True
        ch_min = int(arg)
    elif opt == '--chmax':
        maxima = True
        ch_max = int(arg)
    elif opt == '--tmin':
        min_total = True
        n_min_total = int(arg)
    elif opt == '--tmax':
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
#pesos = np.array([0.1362, 0.0626, 0.3093, 0.4919])
pesos = np.array([0.1336, 0.0614, 0.3102, 0.4948])

##print("As opções escolhidas são:")
##print(f"Modo: {modo}, Unidades: {n_unidades}, Max:{maxima} {ch_max if maxima else ''}, Min:{minima} {ch_min if minima else ''}, PEq:{peq}, TotalMax:{max_total} {n_max_total if max_total else ''}, TotalMin:{min_total} {n_min_total if min_total else ''}, Dif:{dif}{dif_tam if dif else ''}, 40h:{quarenta}, 20h{vinte}")
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
matriz_prof = np.array([1, 1, 1, 1, 1, 1, 1, 1]) #só quantidade
##matriz_tempo = np.array([0, 0, 14, 8, 1, 0, 0, 0]) # tempo disponível
matriz_tempo = np.array([0, 0, 18, 12, 5, 0, 3, 0]) # tempo disponível
                              
conectores = np.array([">=", ">=", ">=", "==", "=="])
nomes_restricoes = np.array(['aulas','h_orientacoes','n_orientacoes','diretor','coords'])
 
# Variáveis de decisão
#nomes
nomes = [str(p) + "." + m_unidades[u][0] for u in range(n_unidades) for p in range(1, n_perfis+1) ]
#variaveis
nx = LpVariable.matrix("x", nomes, cat="Integer", lowBound=0)
saida = np.array(nx).reshape(n_unidades, n_perfis)

# variáveis auxiliares
#desvios = LpVariable.matrix("zd", m_unidades[:n_unidades, 0], cat="Continuous", lowBound=0)
#medias = LpVariable.matrix("zm", m_unidades[:n_unidades, 0], cat="Continuous", lowBound=0)
#media_geral = LpVariable("zmg", cat="Continuous", lowBound=0)

## Percorre os critérios 
modos = np.array(['num', 'peq', 'tempo', 'ch'])
modelos = {}
melhores = {}
piores = {}

# Valores conhecidos a priori
# número máximo é dado pelo total de aulas dividido pela carga horária mínima
# ou é definido pelo usuário
numero_max = round(np.sum(m_unidades[:n_unidades], axis=0)[1] / ch_min) if not max_total else n_max_total
piores['num'] = numero_max
piores['peq'] = numero_max*1.65
piores['tempo'] = 0
piores['ch'] = (ch_max - ch_min)/2

original_stdout = sys.stdout
#filename = f"output{datetime.now().strftime('%H-%M-%S')}.txt"
#filename = f"CBC{modo}_n{n_unidades}-Ma{maxima}-Mi{minima}-Peq{peq}-Tot{n_min_total if min_total else min_total}-{n_max_total if max_total else max_total}-Dif{dif_tam if dif else dif}-40h{quarenta}-20h{vinte}--{datetime.now().strftime('%H-%M-%S')}.txt"
filename = f"CBC completoorient {datetime.now().strftime('%H-%M-%S')}.txt"
fileout =  open(filename, 'w')
sys.stdout = fileout

for modo in modos:
    sys.stdout = original_stdout
    print(f"Modo: {modo}")
    sys.stdout = fileout
    # Verifica modo de tempo e ativa máximo se não estiver ativo
    #if(modo == 'tempo' and max_total == False and minima == False):
    #    minima = True
    
    # Verifica modo de carga horária e ativa máximo e mínimo
    #if(modo == 'ch'):
    minima = True
    maxima = True
    
    # Definir o modelo
    if(modo == 'tempo'):
        modelos[modo] = LpProblem(name=f"Professores-{modo}", sense=LpMaximize)
    else:
        modelos[modo] = LpProblem(name=f"Professores-{modo}", sense=LpMinimize)
        
    # Restrições
    if(modo != 'ch'):
        for r in range(n_restricoes):
            for u in range(n_unidades):
                match conectores[r]:
                    case ">=":
                        modelos[modo] += lpDot(saida[u], m_perfis[r]) >= m_unidades[u][r+1], nomes_restricoes[r] + " " + m_unidades[u][0]#"r " + str(i)
                        # limitar diferença da meta
                        if(dif):
                            modelos[modo] += lpDot(saida[u], m_perfis[r]) <= m_unidades[u][r+1]*(100+dif_tam)/100, f"dif {nomes_restricoes[r]} {m_unidades[u][0]}"
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
    # Considerou-se carga horária média mínima de 12 aulas e máxima de 16
    # Exemplo para 900 aulas
    # soma <= 900/12 -> soma <= 75
    # soma >= 900/16 -> soma >= 56.25
    # a soma deve estar entre 57 e 75, inclusive
    for u in range(n_unidades):
        if(maxima):# or modo == 'tempo'):
            modelos[modo] += ch_max*lpSum(saida[u]) >= m_unidades[u][1], f"{m_unidades[u][0]}_chmax: {math.ceil(m_unidades[u][1]/ch_max)}"
        #if(minima):
        #    modelos[modo] += ch_min*lpSum(saida[u]) <= m_unidades[u][1], f"{m_unidades[u][0]}_chmin: {int(m_unidades[u][1]/ch_min)}"
    
    # Restrição somente no total -> permite exceções como o IBTEC em Monte Carmelo, que tem 19 aulas apenas
    if(maxima):
        modelos[modo] += ch_max*lpSum(saida) >= np.sum(m_unidades[:n_unidades], axis=0)[1], f"chmax: {math.ceil(np.sum(m_unidades[:n_unidades], axis=0)[1]/ch_max)}"
    if(minima):
        modelos[modo] += ch_min*lpSum(saida) <= np.sum(m_unidades[:n_unidades], axis=0)[1], f"chmin: {int(np.sum(m_unidades[:n_unidades], axis=0)[1]/ch_min)}"

    # Restrição do número total de professores
    if(max_total):
        modelos[modo] += lpSum(saida*matriz_prof) <= n_max_total, f"TotalMax: {n_max_total}"
    if(min_total):
        modelos[modo] += lpSum(saida*matriz_prof) >= n_min_total, f"TotalMin: {n_min_total}"

    '''# Restrições de regimes de trabalho
    if(quarenta):
        for u in range(n_unidades):
            modelos[modo] += saida[u][4] + saida[u][5] - p_quarenta*lpSum(saida[u]) <= 0, f"40 horas {m_unidades[u][0]} <= 10%"
    if(vinte):
        for u in range(n_unidades):
            modelos[modo] += saida[u][6] + saida[u][7] - p_vinte*lpSum(saida[u]) <= 0, f"20 horas {m_unidades[u][0]} <= 20%"
    '''
    # Modo de carga horária
    if(modo == 'ch'):            
        # variáveis auxiliares
        desvios = LpVariable.matrix("zd", m_unidades[:n_unidades, 0], cat="Continuous")
        medias = LpVariable.matrix("zm", m_unidades[:n_unidades, 0], cat="Continuous")
        media_geral = LpVariable("zmg", cat="Continuous")
        
        # minimizar desvio em relação à média
        # coeficientes
        # m = (a-b)/(c-d) -> a = media minima, b = media maxima, c = numero maximo, d = numero minimo
        c = np.sum(m_unidades[:n_unidades], axis=0)[1] / ch_min
        d = np.sum(m_unidades[:n_unidades], axis=0)[1] / ch_max
        m = (ch_min - ch_max) / (c - d)
        ##print(f"Geral, c = {c}, d = {d}, m = {m}")
        
        modelos[modo] += media_geral == m*lpSum(saida[:n_unidades]) + ch_min + ch_max, "Media geral"
        
        for u in range(n_unidades):        
            c = m_unidades[u][1] / ch_min
            d = m_unidades[u][1] / ch_max
            m = (ch_min - ch_max) / (c - d)
            ##print(f"{m_unidades[u][0]}, c = {c}, d = {d}, m = {m}")
            # calculo da media
            #modelos[modo] += medias[u] == m*lpSum(saida[u][0]) + ch_min + ch_max, f"{m_unidades[u][0]}_ch_media"
            modelos[modo] += medias[u] == m*lpSum(saida[u]) + ch_min + ch_max, f"{m_unidades[u][0]}_ch_media"
            # desvio
            modelos[modo] += desvios[u] >= medias[u] - media_geral, f"{m_unidades[u][0]}_up"
            modelos[modo] += desvios[u] >= -1*(medias[u] - media_geral), f"{m_unidades[u][0]}_low"
            
    # Função objetivo
    if(modo == 'num'):
        modelos[modo] += lpSum(saida*matriz_prof)
    elif(modo == 'peq'):
        modelos[modo] += lpSum(saida*matriz_peq)
    elif(modo == 'tempo'):
        modelos[modo] += lpSum(saida*matriz_tempo) - np.sum(m_unidades[:n_unidades], axis=0)[2]
    elif(modo == 'ch'):
        modelos[modo] += lpSum(desvios[:n_unidades])/n_unidades
        
    #print(modelos[modo])
    # Imprime os resultados no arquivo txt
    
    print(f'Modo: {modo}')
    print(f'Unidades: {n_unidades}')
    print(f'CH Maxima: {maxima} {ch_max if maxima else ""}')
    print(f'CH Minima: {minima} {ch_min if minima else ""}')
    print(f'P-Equivalente: {peq}')
    print(f'Total: {n_min_total if min_total else "-"} a {n_max_total if max_total else "-"}')
    print(f'Dif: {dif} {dif_tam if dif else ""}')
    print(f'Vinte: {vinte}')
    print(f'Quarenta: {quarenta}')
    print()

    #print(modelos[modo])
    # Resolver o modelo
    #status = model.solve()
    status = modelos[modo].solve(PULP_CBC_CMD(msg=0, timeLimit=time_limit))#time_limit))
    #status = model.solve(GLPK(msg=True))

    # Resultados
    print(f"Situação: {modelos[modo].status}, {LpStatus[modelos[modo].status]}")
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
    
    #if(modo=='ch'):
    #    for var in modelos[modo].variables():
    #        if(var.name[0] == 'p' or var.name[0] == 'z'):
    #            print(f"{var.name}: {var.value()}")
    
    melhores[modo] = objetivo #if modo <> 'tempo' else objetivo*-1
    sys.stdout = original_stdout
    print(f"Ok. Resultado: {melhores[modo]}")
    # Formata resultados e calcula totais
    #resultados = np.full(n_unidades, '', dtype=object)
    #totais = np.full(n_unidades, 0, dtype=int)
    qtdes = np.full((n_unidades, n_perfis), 0, dtype=int)
    index = 0
    perfil = 0
    log = ""
    for var in modelos[modo].variables():
        #if(modo == 'ch'):
        #    print(var.name)
        if(var.name.find('x') == 0):
            log += f"({index},{perfil}): {var.value():2.0f}; "
            valor = int(var.value())
            #resultados[index] += f"{valor:2d} "
            qtdes[index][perfil] = valor
            #totais[index] += valor
            index += 1
            if(index >= n_unidades):
                index = 0
                perfil += 1
                log += "\n"
            if(perfil >= n_perfis+1):
                print(f"i: {index}, p: {perfil}")
                break
    #if(modo == 'ch'):    
    #    print(log)
    sys.stdout = fileout
    
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
    
    #salva em planilha
    ##df = pd.DataFrame(qtdes, columns=['x1','x2','x3','x4','x5','x6','x7','x8'])
    ##df.insert(0, "Unidade", m_unidades[:, 0])
    ##df.to_excel(f'CBC{modo}ori-prel{"q" if quarenta else ""}.xlsx', sheet_name='Resultados', index=False)
    
    ##print(qtdes)
    print()
    print("---------------------------")
    
##### -----------------Fim da primeira 'rodada'-----------------
#sys.stdout = original_stdout
print("Melhores:")
print(melhores)
print("Piores:")
print(piores)
print()
print("---------------------------")
#print(medias)
#print(desvios)
#print(media_geral)
sys.stdout = original_stdout
#input("Aperte enter para continuar...")
sys.stdout = fileout

## Agora uma nova rodada do modelo usando os pesos
# Verifica modo de tempo e ativa máximo se não estiver ativo
#if(modo == 'tempo' and max_total == False and minima == False):
#    minima = True

# Verifica modo de carga horária e ativa máximo e mínimo
#if(modo == 'ch'):
minima = True
maxima = True

# Definir o modelo
modelo = LpProblem(name=f"Professores-Geral", sense=LpMaximize)

#variáveis auxiliares
desvios = LpVariable.matrix("zd", m_unidades[:n_unidades, 0], cat="Continuous", lowBound=0)
medias = LpVariable.matrix("zm", m_unidades[:n_unidades, 0], cat="Continuous", lowBound=0)
media_geral = LpVariable("zmg", cat="Continuous", lowBound=0)
    
# Restrições
for r in range(n_restricoes):
    for u in range(n_unidades):
        match conectores[r]:
            case ">=":
                modelo += lpDot(saida[u], m_perfis[r]) >= m_unidades[u][r+1], nomes_restricoes[r] + " " + m_unidades[u][0]#"r " + str(i)
                # limitar diferença da meta
                if(dif):
                    modelo += lpDot(saida[u], m_perfis[r]) <= m_unidades[u][r+1]*(100+dif_tam)/100, f"dif {nomes_restricoes[r]} {m_unidades[u][0]}"
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
    if(maxima):# or modo == 'tempo'):
        modelo += ch_max*lpSum(saida[u]) >= m_unidades[u][1], f"{m_unidades[u][0]}_chmax: {math.ceil(m_unidades[u][1]/ch_max)}"
    #if(minima):
    #    modelo += ch_min*lpSum(saida[u]) <= m_unidades[u][1], f"{m_unidades[u][0]}_chmin: {int(m_unidades[u][1]/ch_min)}"

if(maxima):
    modelo += ch_max*lpSum(saida) >= np.sum(m_unidades[:n_unidades], axis=0)[1], f"chmax: {math.ceil(np.sum(m_unidades[:n_unidades], axis=0)[1]/ch_max)}"
if(minima):
    modelo += ch_min*lpSum(saida) <= np.sum(m_unidades[:n_unidades], axis=0)[1], f"chmin: {int(np.sum(m_unidades[:n_unidades], axis=0)[1]/ch_min)}"

# Restrição do número total de professores
if(max_total):
    modelo += lpSum(saida*matriz_prof) <= n_max_total, f"TotalMax: {n_max_total}"
if(min_total):
    modelo += lpSum(saida*matriz_prof) >= n_min_total, f"TotalMin: {n_min_total}"

# Restrições de regimes de trabalho
if(quarenta):
    for u in range(n_unidades):
        modelo += saida[u][4] + saida[u][5] - p_quarenta*lpSum(saida[u]) <= 0, f"40 horas {m_unidades[u][0]} <= 10%"
if(vinte):
    for u in range(n_unidades):
        modelo += saida[u][6] + saida[u][7] - p_vinte*lpSum(saida[u]) <= 0, f"20 horas {m_unidades[u][0]} <= 20%"

# minimizar desvio em relação à média
# coeficientes
# m = (a-b)/(c-d) -> a = media minima, b = media maxima, c = numero maximo, d = numero minimo
c = np.sum(m_unidades[:n_unidades], axis=0)[1] / ch_min
d = np.sum(m_unidades[:n_unidades], axis=0)[1] / ch_max
m = (ch_min - ch_max) / (c - d)
##print(f"Geral, c = {c}, d = {d}, m = {m}")

modelo += media_geral == m*lpSum(saida[:n_unidades]) + ch_min + ch_max, "Media geral"

for u in range(n_unidades):        
    c = m_unidades[u][1] / ch_min
    d = m_unidades[u][1] / ch_max
    m = (ch_min - ch_max) / (c - d)
    ##print(f"{m_unidades[u][0]}, c = {c}, d = {d}, m = {m}")
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

# Função objetivo
modelo += lpSum(pontuacoes*pesos)
    
#print(model)
# Imprime os resultados no arquivo txt

print(f'Modo: Geral')
print(f'Unidades: {n_unidades}')
print(f'CH Maxima: {maxima} {ch_max if maxima else ""}')
print(f'CH Minima: {minima} {ch_min if minima else ""}')
print(f'P-Equivalente: {peq}')
print(f'Total: {n_min_total if min_total else "-"} a {n_max_total if max_total else "-"}')
print(f'Dif: {dif} {dif_tam if dif else ""}')
print(f'Vinte: {vinte}')
print(f'Quarenta: {quarenta}')
print()

print(modelo)
# Resolver o modelo
#status = model.solve()
status = modelo.solve(PULP_CBC_CMD(msg=0, timeLimit=time_limit))
#status = model.solve(GLPK(msg=True))

# Resultados
print(f"Situação: {modelo.status}, {LpStatus[modelo.status]}")
#if(modo == 'peq' or modo == 'ch'):
objetivo = f"{modelo.objective.value():.4f}"
#else:
#    objetivo = int(modelo.objective.value())
#    if(modo == 'tempo'):
#        objetivo = objetivo*-1
#print(f"Objetivo: {objetivo} {'horas' if modo == 'tempo' else 'Prof-Equivalente' if peq else 'Professores' if modo == 'num' else 'aulas/prof'}")
print(f"Objetivo: {objetivo}")
print(f"Resolvido em {modelo.solutionTime} segundos")
print("")
#input("Aperte enter para continuar...")

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
print(f"--------+-------------+----------------+-------------+-------------+-------------+---------+---------+----------+")
print(f"Unidade |      aulas  |  horas_orient  |  num_orient |   diretor   |   coords.   |   40h   |   20h   | ch media |") #{'minim' if minimo else ''} {'maxim' if maximo or modo == 'tempo' else ''}")
print(f"--------+-------------+----------------+-------------+-------------+-------------+---------+---------+----------+")
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
print(f"--------+-------------+----------------+-------------+-------------+-------------+---------+---------+----------+")
print()
#print("------------------Modelo:------------------")
#print()
#print(model)
for var in modelo.variables():
    if(var.name[0] == 'p' or var.name[0] == 'z'):
        print(f"{var.name}: {var.value()}")
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

##input("Aperte Enter para finalizar")
