#-#from pulp import * #LpMaximize, LpMinimize, LpProblem, LpStatus, lpSum, LpVariable
import numpy as np
from datetime import time, datetime
import sys, getopt
import pandas as pd
import math

# Import OR-Tools wrapper for linear programming
from ortools.linear_solver import pywraplp

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
modulo = False
time_limit = 15

# Dados do problema
n_perfis = 8
n_restricoes = 5
n_unidades = len(m_unidades)

opts, args = getopt.getopt(sys.argv[1:],"zpiadvqn:m:",["chmin=","chmax=","tmax=","tmin=","limite="])
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
    elif opt == '-z':
        modulo = True
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

# Definir o modelo
solver = pywraplp.Solver('Professores', pywraplp.Solver.CBC_MIXED_INTEGER_PROGRAMMING)
#solver = pywraplp.Solver('Professores', pywraplp.Solver.GLOP_LINEAR_PROGRAMMING)
#solver = pywraplp.Solver('Professores', pywraplp.Solver.SAT_INTEGER_PROGRAMMING)
solver.set_time_limit(30*1000)
solver.EnableOutput()

# Variáveis de decisão
saida = [solver.IntVar(0, solver.infinity(), f"x{p}_{m_unidades[u][0]}") for u in range(n_unidades) for p in range(1, n_perfis+1)]
outra = np.array(saida).reshape(n_unidades, n_perfis)

# auxiliares
desvios = [solver.NumVar(0, solver.infinity(), f"zd_{m_unidades[u][0]}") for u in range(n_unidades)]
medias = [solver.NumVar(0, solver.infinity(), f"zm_{m_unidades[u][0]}") for u in range(n_unidades)]
media_geral = solver.NumVar(0, solver.infinity(), "zmg")
 
# Função objetivo
if(modo == 'num'):
    if(peq):
        solver.Minimize(sum(outra.dot(matriz_peq)))
    else:
        solver.Minimize(sum(saida))
elif(modo == 'tempo'):
    solver.Minimize(sum(outra.dot(matriz_tempon)))
elif(modo == 'ch'):
    if(modulo):
        solver.Minimize(sum(desviosabs[:n_unidades]))
    else:    
        solver.Minimize(sum(desvios[:n_unidades]))

# Restrições
for r in range(n_restricoes):
    for u in range(n_unidades):
        match conectores[r]:
            case ">=":
                solver.Add(sum(outra[u][p] * m_perfis[r][p] for p in range(n_perfis)) >= m_unidades[u][r+1], nomes_restricoes[r] + " " + m_unidades[u][0])
                # limitar diferença da meta
                #-#if(dif):
                #-#    model += lpDot(saida[u], m_perfis[r]) <= m_unidades[u][r+1]*(100+dif_tam)/100, f"dif {nomes_restricoes[r]} {m_unidades[u][0]}"
            case "==":
                solver.Add(sum(outra[u][p] * m_perfis[r][p] for p in range(n_perfis)) == m_unidades[u][r+1], nomes_restricoes[r] + " " + m_unidades[u][0])
            case "<=":
                solver.Add(sum(outra[u][p] * m_perfis[r][p] for p in range(n_perfis)) <= m_unidades[u][r+1], nomes_restricoes[r] + " " + m_unidades[u][0])                

# Restrições de máximo e mínimo por unidade -> carga horária média
# ------------------
# Considerou-se carga horária média mínima de 12 aulas e máxima de 16
# Exemplo para 900 aulas
# soma <= 900/12 -> soma <= 75
# soma >= 900/16 -> soma >= 56.25
# a soma deve estar entre 57 e 75, inclusive
'''for u in range(n_unidades):
    if(maxima):# or modo == 'tempo'):
        solver.Add(ch_max*sum(outra[u]) >= m_unidades[u][1], f"{m_unidades[u][0]}_chmax: {math.ceil(m_unidades[u][1]/ch_max)}")
    if(minima):
        solver.Add(ch_min*sum(outra[u]) <= m_unidades[u][1], f"{m_unidades[u][0]}_chmin: {int(m_unidades[u][1]/ch_min)}")
'''
if(maxima):
    solver.Add(ch_max*sum(saida) >= np.sum(m_unidades[:n_unidades], axis=0)[1], f"chmax: {math.ceil(np.sum(m_unidades[:n_unidades], axis=0)[1]/ch_max)}")
if(minima):
    solver.Add(ch_min*sum(saida) <= np.sum(m_unidades[:n_unidades], axis=0)[1], f"chmin: {int(np.sum(m_unidades[:n_unidades], axis=0)[1]/ch_min)}")

# Restrição do número total de professores
if(max_total):
    solver.Add(sum(saida) <= n_max_total, f"TotalMax: {n_max_total}")
if(min_total):
    solver.Add(sum(saida) >= n_min_total, f"TotalMin: {n_min_total}")

# Restrições de regimes de trabalho
if(quarenta):
    for u in range(n_unidades):
        solver.Add(outra[u][4] + outra[u][5] - p_quarenta*sum(outra[u]) <= 0, f"40 horas {m_unidades[u][0]} <= 10%")
if(vinte):
    for u in range(n_unidades):
        solver.Add(outra[u][6] + outra[u][7] - p_vinte*sum(outra[u]) <= 0, f"20 horas {m_unidades[u][0]} <= 20%")

# minimizar desvio em relação à média
if(modo == 'ch'):
    # coeficientes
    # m = (a-b)/(c-d) -> a = media minima, b = media maxima, c = numero maximo, d = numero minimo
    c = np.sum(m_unidades[:n_unidades], axis=0)[1] / ch_min
    d = np.sum(m_unidades[:n_unidades], axis=0)[1] / ch_max
    m = (ch_min - ch_max) / (c - d)
    print(f"Geral, c = {c}, d = {d}, m = {m}")
    
    solver.Add(media_geral == m*sum(saida) + ch_min + ch_max, "Media geral")
    for u in range(n_unidades):        
        c = m_unidades[u][1] / ch_min
        d = m_unidades[u][1] / ch_max
        m = (ch_min - ch_max) / (c - d)
        print(f"{m_unidades[u][0]}, c = {c}, d = {d}, m = {m}")
        # calculo da media
        solver.Add(medias[u] == m*sum(outra[u]) + ch_min + ch_max, f"{m_unidades[u][0]}_ch_media")
        # desvio
        if(modulo):
            model += desvios[u] == medias[u] - media_geral, f"{m_unidades[u][0]}_desvio"
            model += desviosabs[u] >= desvios[u], f"{m_unidades[u][0]}_up"
            model += desviosabs[u] >= -desvios[u], f"{m_unidades[u][0]}_low"
        else:
            #solver.Add(desvios[u] >= m*sum(outra[u]) + ch_min + ch_max - media_geral, f"{m_unidades[u][0]}_up")
            #solver.Add(desvios[u] >= -1*(m*sum(outra[u]) + ch_min + ch_max - media_geral), f"{m_unidades[u][0]}_low")
            solver.Add(desvios[u] >= medias[u] - media_geral, f"{m_unidades[u][0]}_up")
            solver.Add(desvios[u] >= -1*(medias[u] - media_geral), f"{m_unidades[u][0]}_low")

# Imprime os resultados no arquivo txt
original_stdout = sys.stdout
#filename = f"output{datetime.now().strftime('%H-%M-%S')}.txt"
filename = f"CHOR-{modo}_n{n_unidades}-Ma{maxima}-Mi{minima}-Peq{peq}-Tot{n_min_total if min_total else min_total}-{n_max_total if max_total else max_total}-Dif{dif_tam if dif else dif}-40h{quarenta}-20h{vinte}--{datetime.now().strftime('%H-%M-%S')}.txt"
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
print(f'Modulo: {modulo}')
print()

# Resolver o modelo
status = solver.Solve()

# If an optimal solution has been found, print results
if status == pywraplp.Solver.OPTIMAL or pywraplp.Solver.FEASIBLE:
    print(f'Situação: {status}')
    if peq or modo == 'ch':
        objetivo = f"{solver.Objective().Value():.2f}"
    else:
        objetivo = int(solver.Objective().Value())
    print(f"Objetivo: {objetivo} {'horas' if modo == 'tempo' else 'Prof-Equivalente' if peq else 'Professores' if modo == 'num' else 'aulas/prof'}")
    print(f'Resolvido em {solver.wall_time():.2f} milissegundos em {solver.iterations()} iterações')
else:
    print('Não foi possível encontrar uma solução')
  
'''# Resultados
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
'''

# salva resultados em outra matriz
qtdes = np.full((n_unidades, n_perfis), 0, dtype=int)
for u in range(n_unidades):
    for p in range(n_perfis):
        qtdes[u][p] = int(outra[u][p].solution_value())

print("")
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
# imprime o modelo no arquivo
print(solver.ExportModelAsLpFormat(False).replace('\\', '').replace(',_', ','), sep='\n')

fileout.close()
#Imprime na tela    
sys.stdout = original_stdout

#salva em planilha
df = pd.DataFrame(qtdes, columns=['x1','x2','x3','x4','x5','x6','x7','x8'])
df.insert(0, "Unidade", m_unidades[:, 0])
df.to_excel(f'CHOR{modo}-Peq{peq}-test.xlsx', sheet_name='Resultados', index=False)

##input("Aperte Enter para finalizar")
