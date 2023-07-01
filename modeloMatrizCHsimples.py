from pulp import * #LpMaximize, LpMinimize, LpProblem, LpStatus, lpSum, LpVariable
import numpy as np
from datetime import time, datetime
import sys, getopt
import pandas as pd
import math

df_todas = pd.read_excel('importar.xlsx', sheet_name=['unidades','perfis'])
m_unidades = df_todas['unidades'].to_numpy()
m_perfis = df_todas['perfis'].to_numpy()
m_perfis = np.delete(m_perfis, 0, axis=1)

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

# Dados do problema
n_perfis = 8
n_restricoes = 5
n_unidades = len(m_unidades)

opts, args = getopt.getopt(sys.argv[1:],"piadvqn:m:",["chmin=","chmax=","tmax=","tmin=","limite="])
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
    elif opt == '-m':
        modo = arg
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

# Verifica modo de tempo e ativa máximo se não estiver ativo
if(modo == 'tempo' and max_total == False and minima == False):
    minima = True
    
# Verifica modo de carga horária e ativa máximo e mínimo
if(modo == 'ch'):
    minima = True
    maxima = True

print("As opções escolhidas são:")
print(f"Modo: {modo}, Unidades: {n_unidades}, Max:{maxima} {ch_max if maxima else ''}, Min:{minima} {ch_min if minima else ''}, PEq:{peq}, TotalMax:{max_total} {n_max_total if max_total else ''}, TotalMin:{min_total} {n_min_total if min_total else ''}, Dif:{dif}{dif_tam if dif else ''}, 40h:{quarenta}, 20h{vinte}")

# Dá opção de alterar os parâmetros
'''tecla = input("Aperte enter para continuar ou a para alterar\n")
if(tecla=='a'):
    print("-m escolhe o modo (num ou tempo). Exemplo: -m tempo")
    print("-i ativa a quantidade mínima por unidade e -a ativa a quantidade máxima por unidade")
    print("-p usa os fatores do professor-equivalente em vez do número de professores")
    print("-t ativa a quantidade máxima total de professores. Deve ser seguido do número desejado. Exemplo: -t 100")
    print("-n define o número de unidades. Exemplo: -n 4")
    print("-d ativa a limitação de 'sobra' nos parâmetros (20%)")
    print("-r ativa a limitação de professores em regime de 20 horas (20%)")
    print("-q ativa a limitação de professores em regime de 40 horas (10%)")
    opcoes = input("Digite as opções desejadas, separadas por espaço. Exemplo: -t 130 -d -p (define 130 como total máximo e ativa a opção de professor-equivalente e limitação da 'sobra'\n")
'''    

# Custos
matriz_peq = np.array([1.65, 1.65, 1.65, 1.65, 1, 1, 0.6, 0.6]) #professor-equivalente - x1 a x4 = DE, x5 e x6 = 40h, x7 e x8 = 20h
matriz_prof = np.array([1, 1, 1, 1, 1, 1, 1, 1]) #só quantidade
matriz_tempo = np.array([0, 0, 14, 8, 1, 0, 0, 0]) # tempo disponível
matriz_tempon = np.array([0, 0, -14, -8, -1, 0, 0, 0]) # tempo disponível
                              
conectores = np.array([">=", ">=", ">=", "==", "==", "<="])
nomes_restricoes = np.array(['aulas','orientacoes','gestao','diretor','coords'])
 
# Variáveis de decisão
#nomes
'''nomes = [str(p) + "." + m_unidades[u][0] for u in range(n_unidades) for p in range(1, n_perfis+1) ]
#variaveis
nx = LpVariable.matrix("x", nomes, cat="Integer", lowBound=0)
saida = np.array(nx).reshape(n_unidades, n_perfis)
'''

nx = LpVariable.matrix("x", m_unidades[:n_unidades, 0], cat="Integer", lowBound=0)
saida = np.array(nx)
print(nx)
print(saida)
input('...')

# variáveis auxiliares
desvios = LpVariable.matrix("zd", m_unidades[:n_unidades, 0], cat="Continuous")
medias = LpVariable.matrix("zm", m_unidades[:n_unidades, 0], cat="Continuous")
media_geral = LpVariable("zmg", cat="Continuous")

# Definir o modelo
model = LpProblem(name="Professores0.3", sense=LpMinimize)

# Função objetivo
if(modo == 'num'):
    if(peq):
        model += lpSum(saida)
    else:
        model += lpSum(saida)
elif(modo == 'tempo'):
    model += lpSum(saida)
elif(modo == 'ch'):
    model += lpSum(desvios[:n_unidades])/n_unidades
    
# Restrições
'''for r in range(n_restricoes):
    for u in range(n_unidades):
        match conectores[r]:
            case ">=":
                model += lpDot(saida[u], m_perfis[r]) >= m_unidades[u][r+1], nomes_restricoes[r] + " " + m_unidades[u][0]#"r " + str(i)
                # limitar diferença da meta
                if(dif):
                    model += lpDot(saida[u], m_perfis[r]) <= m_unidades[u][r+1]*(100+dif_tam)/100, f"dif {nomes_restricoes[r]} {m_unidades[u][0]}"
            case "==":
                model += lpDot(saida[u], m_perfis[r]) == m_unidades[u][r+1], nomes_restricoes[r] + " " + m_unidades[u][0]#"r " + str(i)
            case "<=":
                model += lpDot(saida[u], m_perfis[r]) <= m_unidades[u][r+1], nomes_restricoes[r] + " " + m_unidades[u][0]#"r " + str(i)        
'''
# Restrições de máximo e mínimo por unidade -> carga horária média
# ------------------
# Considerou-se carga horária média mínima de 12 aulas e máxima de 16
# Exemplo para 900 aulas
# soma <= 900/12 -> soma <= 75
# soma >= 900/16 -> soma >= 56.25
# a soma deve estar entre 57 e 75, inclusive
'''for u in range(n_unidades):
    if(maxima):# or modo == 'tempo'):
        model += ch_max*lpSum(saida[u]) >= m_unidades[u][1], f"{m_unidades[u][0]}_chmax: {math.ceil(m_unidades[u][1]/ch_max)}"
    if(minima):
        model += ch_min*lpSum(saida[u]) <= m_unidades[u][1], f"{m_unidades[u][0]}_chmin: {int(m_unidades[u][1]/ch_min)}"
'''
if(maxima):
    model += ch_max*lpSum(saida) >= np.sum(m_unidades[:n_unidades], axis=0)[1], f"chmax: {math.ceil(np.sum(m_unidades[:n_unidades], axis=0)[1]/ch_max)}"
if(minima):
    model += ch_min*lpSum(saida) <= np.sum(m_unidades[:n_unidades], axis=0)[1], f"chmin: {int(np.sum(m_unidades[:n_unidades], axis=0)[1]/ch_min)}"

# Restrição do número total de professores
if(max_total):
    model += lpSum(saida*matriz_prof) <= n_max_total, f"TotalMax: {n_max_total}"
if(min_total):
    model += lpSum(saida*matriz_prof) >= n_min_total, f"TotalMin: {n_min_total}"

# Restrições de regimes de trabalho
'''if(quarenta):
    for u in range(n_unidades):
        model += saida[u][4] + saida[u][5] - p_quarenta*lpSum(saida[u]) <= 0, f"40 horas {m_unidades[u][0]} <= 10%"
if(vinte):
    for u in range(n_unidades):
        model += saida[u][6] + saida[u][7] - p_vinte*lpSum(saida[u]) <= 0, f"20 horas {m_unidades[u][0]} <= 20%"
'''
# minimizar desvio em relação à média
#media_geral = np.sum(m_unidades, axis=0)[1] / ch_media
if(modo == 'ch'):
    # coeficientes
    # m = (a-b)/(c-d) -> a = media minima, b = media maxima, c = numero maximo, d = numero minimo
    c = np.sum(m_unidades[:n_unidades], axis=0)[1] / ch_min
    d = np.sum(m_unidades[:n_unidades], axis=0)[1] / ch_max
    m = (ch_min - ch_max) / (c - d)
    print(f"Geral, c = {c}, d = {d}, m = {m}")
    
    model += media_geral == m*lpSum(saida[:n_unidades]) + ch_min + ch_max, "Media geral"
    for u in range(n_unidades):
        #model += ch_media*lpSum(saida[u]) == m_unidades[u][1], f"{m_unidades[u][0]}_ch_media"
        #model += lpSum(saida[u])*desvios[u] >= (m_unidades[0][1] - media_geral*lpSum(saida[u]))
        #model += lpSum(saida[u])*desvios[u] >= (media_geral*lpSum(saida[u]) - m_unidades[0][1])
        
        ##model += ch_media*(1+desvio_max)*lpSum(saida[u]) >= m_unidades[u][1], f"{m_unidades[u][0]}_up"
        ##model += ch_media*(1-desvio_max)*lpSum(saida[u]) <= m_unidades[u][1], f"{m_unidades[u][0]}_low"
        
        c = m_unidades[u][1] / ch_min
        d = m_unidades[u][1] / ch_max
        m = (ch_min - ch_max) / (c - d)
        print(f"{m_unidades[u][0]}, c = {c}, d = {d}, m = {m}")
        # calculo da media
        model += medias[u] == m*saida[u] + ch_min + ch_max, f"{m_unidades[u][0]}_ch_media"
        # desvio
        model += desvios[u] >= medias[u] - media_geral, f"{m_unidades[u][0]}_up"
        model += desvios[u] >= -1*(medias[u] - media_geral), f"{m_unidades[u][0]}_low"

#print(model)
# Imprime os resultados no arquivo txt
original_stdout = sys.stdout
#filename = f"output{datetime.now().strftime('%H-%M-%S')}.txt"
filename = f"CBC{modo}_n{n_unidades}-Ma{maxima}-Mi{minima}-Peq{peq}-Tot{n_min_total if min_total else min_total}-{n_max_total if max_total else max_total}-Dif{dif_tam if dif else dif}-40h{quarenta}-20h{vinte}--{datetime.now().strftime('%H-%M-%S')}.txt"
fileout =  open(filename, 'w')
sys.stdout = fileout
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

# Resolver o modelo
#status = model.solve()
status = model.solve(PULP_CBC_CMD(msg=1, timeLimit=time_limit))
#status = model.solve(GLPK(msg=True))

# Resultados
print(f"Situação: {model.status}, {LpStatus[model.status]}")
if peq or modo == 'ch':
    objetivo = f"{model.objective.value():.2f}"
else:
    objetivo = int(model.objective.value())
print(f"Objetivo: {objetivo} {'horas' if modo == 'tempo' else 'Prof-Equivalente' if peq else 'Professores' if modo == 'num' else 'aulas/prof'}")
print(f"Resolvido em {model.solutionTime} segundos")
print("")

# Formata resultados e calcula totais
resultados = np.full(n_unidades, '', dtype=object)
#totais = np.full(n_unidades, 0, dtype=int)
qtdes = np.full((n_unidades, n_perfis), 0, dtype=int)
index = 0
perfil = 0
for var in model.variables():
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
print(f"--------+---------------------------------+-------+--------+-------+------------+")
print(f"Unidade |  x1  x2  x3  x4  x5  x6  x7  x8 | total |   P-Eq | Tempo | Tempo/prof |")
print(f"--------+---------------------------------+-------+--------+-------+------------+")
for u in range(n_unidades):
    print(f"{m_unidades[u][0]:5s}   | " + " ".join([f"{qtdes[u][p]:3d}" for p in range(n_perfis)]) + f" |  {np.sum(qtdes[u]):4d} | {np.sum(qtdes[u]*matriz_peq):6.2f} |  {np.sum(qtdes[u]*matriz_tempon)*-1:4d} |    {(np.sum(qtdes[u]*matriz_tempon)*-1)/np.sum(qtdes[u]):7.3f} |")
print(f"--------+---------------------------------+-------+--------+-------+------------+")
print(f"Total   | " + " ".join([f"{np.sum(qtdes, axis=0)[p]:3d}" for p in range(n_perfis)]) + f" |  {np.sum(qtdes):4d} | {np.sum(qtdes*matriz_peq):6.2f} |  {np.sum(qtdes*matriz_tempon)*-1:4d} |    {(np.sum(qtdes*matriz_tempon)*-1)/np.sum(qtdes):7.3f} |")
print(f"--------+---------------------------------+-------+--------+-------+------------+")
print("")

print("Parâmetros:")
print(f"--------+-------------+-------------+-------------+-------------+-------------+----------+----------+----------+")
print(f"Unidade |     aulas   |    orient   |    gestao   |   diretor   |   coords.   |    40h   |    20h   | ch media |") #{'minim' if minimo else ''} {'maxim' if maximo or modo == 'tempo' else ''}")
print(f"--------+-------------+-------------+-------------+-------------+-------------+----------+----------+----------+")
for u in range(n_unidades):
    print(f"{m_unidades[u][0]:5s}   | " 
    + " ".join([f"{np.sum(qtdes[u]*m_perfis[p]):4d} (+{(np.sum(qtdes[u]*m_perfis[p]) - m_unidades[u][p+1]):3d}) |" for p in range(5)])
    + f"  {((qtdes[u][4] + qtdes[u][5]) / np.sum(qtdes[u]))*100:6.2f}% |" #40h
    + f"  {((qtdes[u][6] + qtdes[u][7]) / np.sum(qtdes[u]))*100:6.2f}% |" #20h
    + f"  {m_unidades[u][1] / np.sum(qtdes[u]):7.3f} |"
    )
print(f"--------+-------------+-------------+-------------+-------------+-------------+----------+----------+----------+")
print(f"Total   | " + " ".join([f"{np.sum(qtdes*m_perfis[p]):4d} (+{np.sum(qtdes*m_perfis[p]) - int(np.sum(m_unidades, axis=0)[p+1]):3d}) |" for p in range(5)])
+ f"  {(np.sum(qtdes, axis=0)[4] + np.sum(qtdes, axis=0)[5]) / np.sum(qtdes)*100:6.2f}% |"
+ f"  {(np.sum(qtdes, axis=0)[6] + np.sum(qtdes, axis=0)[7]) / np.sum(qtdes)*100:6.2f}% |"
+ f"  {np.sum(m_unidades, axis=0)[1]/np.sum(qtdes):7.3f} |"
)
print(f"--------+-------------+-------------+-------------+-------------+-------------+----------+----------+----------+")
print()
print("------------------Modelo:------------------")
print()
print(model)
fileout.close()

#Imprime na tela    
sys.stdout = original_stdout
print(f"Situação: {model.status}, {LpStatus[model.status]}")
if peq or modo == 'ch':
    objetivo = f"{model.objective.value():.4f}"
else:
    objetivo = int(model.objective.value())
print(f"Objetivo: {objetivo} {'horas' if modo == 'tempo' else 'Prof-Equivalente' if peq else 'Professores' if modo == 'num' else 'aulas/prof'}")
print(f"Resolvido em {model.solutionTime} segundos")
print("")
print(f"Verifique o arquivo {filename} para o relatório completo")
print("")

#salva em planilha
#df = pd.DataFrame(qtdes, columns=['x1','x2','x3','x4','x5','x6','x7','x8'])
#df.insert(0, "Unidade", m_unidades[:, 0])
#df.to_excel(f'CBC{modo}-Peq{peq}-test.xlsx', sheet_name='Resultados', index=False)

for var in model.variables():
    print(f"{var.name}: {var.value()}")

##input("Aperte Enter para finalizar")
