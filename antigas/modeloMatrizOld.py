from pulp import * #LpMaximize, LpMinimize, LpProblem, LpStatus, lpSum, LpVariable
import numpy as np
from datetime import time, datetime
import sys, getopt
import pandas as pd
import math

df = pd.read_excel('importar.xlsx')
print(df)

m_unidades = df.to_numpy()
print(m_unidades)

#Configurações
maximo = False #número máximo por unidade
ch_min = 12
minimo = False #número mínimo por unidade
ch_max = 16
peq = False    #usar Professor-Equivalente
total = False  #número máximo total
dif = False    #limitar diferença da meta
dif_tam = 50   #percentual aceito
modo = 'num'   #modo 'num' = minimizar quantidade ou P-Eq; modo 'tempo' = maximizar tempo disponível
quarenta = False
p_quarenta = 0.1
vinte = False
p_vinte = 0.2


# Dados do problema
n_perfis = 8
n_restricoes = 5
n_unidades = 5
unidades = np.array(["FACOM", "FAEFI", "FAMAT", "ICBIM", "IFILO"]) # em ordem alfabética

opts, args = getopt.getopt(sys.argv[1:],"piadvqt:n:m:")
for opt, arg in opts:
    if opt == '-p':
        peq = True
    elif opt == '-i':
        minimo = True
    elif opt == '-a':
        maximo = True
    elif opt == '-t':
        total = True
        n_total = int(arg)
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

# Verificado modo de tempo e ativa máximo se não estiver ativo
if(modo == 'tempo' and total == False and maximo == False):
    maximo = True

print("As opções escolhidas são:")
print(f"Modo: {modo}, Unidades: {n_unidades}, Max:{maximo}, Min:{minimo}, PEq:{peq}, Total:{total} {n_total if total == True else ''}, Dif:{dif}{dif_tam if dif else ''}, 40h:{quarenta}, 20h{vinte}")

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

# Quantidade de aulas e orientações por perfil
aulas_perfil =       [0, 10, 12, 16, 20, 24, 10, 12]
orientacoes_perfil = [0,  0,  4,  4,  4,  0,  3,  0]
gestao_perfil =     [40, 20,  0,  0,  0,  0,  0,  0]

matriz_perfil = np.array([
                          [0, 10, 12, 16, 20, 24, 10, 12]   #aulas
                         ,[0,  0,  4,  4,  4,  0,  3,  0]   #orientações
                        ,[40, 20,  0,  0,  0,  0,  0,  0]   #gestão
                         ,[1,  0,  0,  0,  0,  0,  0,  0]   #diretor
                         ,[0,  1,  0,  0,  0,  0,  0,  0]   #coordenadores
                         ])

##matriz_regimes = np.array([[-1, -1, -1, -1,  9,  9, -1, -1]   #auleiros (40h) <= 10%
##                          ,[-2, -2, -2, -2, -2, -2,  8,  8]]) #20h            <= 20%
# Cálculo dos coeficientes para limitação do perfil "auleiro"
# partindo de (x5+x6)/soma <= 0.1 (máximo 10% - a ser definido)
# multiplicamos por soma dos dois lados, obtendo (x5+x6) <= 0.1*soma
# colocando todas as variáveis do lado esquerdo: (x5+x6) - 0.1*soma <= 0
# multiplicamos pelo valor do percentual (10) para obter coeficientes inteiros: 10*(x5+x6) - soma <= 0
# as variáveis x5 e x6 ficam com coeficiente 9 (10-1 ou (1-0.1)*10) e as demais com -1 (-0.1*10)
# No caso do regime de 20 horas aplica-se o mesmo raciocínio, com percentual máximo de 20%.
# Assim, os coeficientes de X7 e X8 ficam 8 e dos demais -2

                             #aulas orient gestao diretor coords
matriz_restricoes = np.array([
                              [640,     70,   100,      1,    3] # FACOM
                             ,[420,     50,   100,      1,    3] # FAEFI
                            ,[1080,     30,    80,      1,    2] # FAMAT
                             ,[720,     20,    60,      1,    1] # ICBIM
                             ,[248,     35,    60,      1,    1] # IFILO
                              
                             ])

                              
                           #max min -> calculados pela carga horária média (ver abaixo)  
maximos_minimos = np.array([
                            [40, 54] #FACOM
                           ,[27, 35] #FAEFI
                           ,[68, 90] #FAMAT
                           ,[45, 60] #ICBIM
                           ,[16, 21] #IFILO
                           ])
# Restrições de carga horária média
# Considerou-se carga horária média mínima de 12 aulas e máxima de 16
# para construir a restrição do modelo, partimos de aulas/soma >= 12
# multiplicamos por soma dos dois lados: aulas >= 12*soma
# dividimos os dois lados por 12: aulas/12 >= soma
# invertendo os lados: soma <= aulas/12
# aplicando o mesmo raciocinio para a média máxima, partindo de aulas/soma <= 16
# obtemos soma >= aulas/16
# Exemplo para 900 aulas
# soma <= 900/12 -> soma <= 75
# soma >= 900/16 -> soma >= 56.25
# a soma deve estar entre 57 e 75, inclusive
                              
conectores = np.array([">=", ">=", ">=", "==", "==", "<="])
nomes_restricoes = np.array(['aulas','orientacoes','gestao','diretor','coords','auleiros'])
 
# Variáveis de decisão
##x = {i: LpVariable(name=f"x{i}", lowBound=0, cat="Integer") for i in range(1, 9)}

#nomes
##nomes = [str(i) + "." + unidades[j] for j in range(n_unidades) for i in range(1, n_perfis+1) ]
nomes = [str(i) + "." + m_unidades[j][0] for j in range(n_unidades) for i in range(1, n_perfis+1) ]
#variaveis
nx = LpVariable.matrix("x", nomes, cat="Integer", lowBound=0)
saida = np.array(nx).reshape(n_unidades, n_perfis)

#print(f"Saida: {saida}")
#print(f"Nx: {nx}")
#input('Continuar...')

# Definir o modelo
# Função objetivo
if(modo == 'num'):
    model = LpProblem(name="professores0.1", sense=LpMinimize)
    if(peq):
        model += lpSum(saida*matriz_peq)
    else:
        model += lpSum(saida*matriz_prof)
elif(modo == 'tempo'):
    model = LpProblem(name="professores0.1", sense=LpMaximize)
    model += lpSum(saida*matriz_tempo)

# Restrições
##for i in range(n_restricoes):
for i in range(1, n_restricoes+1):
    for j in range(n_unidades): #print(lpDot(nx, matriz_perfil[i]))# >= matriz_restricoes[i])
        match conectores[i]:
            case ">=":
                model += lpDot(saida[j], matriz_perfil[i-1]) >= m_unidades[j][i], nomes_restricoes[i-1] + " " + m_unidades[j][0]#"r " + str(i)
                # limitar diferença da meta
                if(dif):
                    model += lpDot(saida[j], matriz_perfil[i-1]) <= m_unidades[j][i]*(100+dif_tam)/100, f"dif {nomes_restricoes[i-1]} {m_unidades[j][0]}"
            case "==":
                model += lpDot(saida[j], matriz_perfil[i-1]) == m_unidades[j][i], nomes_restricoes[i-1] + " " + m_unidades[j][0]#"r " + str(i)
            case "<=":
                model += lpDot(saida[j], matriz_perfil[i-1]) <= m_unidades[j][i], nomes_restricoes[i-1] + " " + m_unidades[j][0]#"r " + str(i)        

#Restrições de máximo e mínimo por unidade -> carga horária média
for u in range(n_unidades):
    if(maximo):# or modo == 'tempo'):
        ##model += lpDot(saida[u], matriz_prof) <= maximos_minimos[u][1], f"{m_unidades[u][0]}_max: {maximos_minimos[u][1]}"
        #model += lpDot(saida[u], matriz_prof) <= m_unidades[u][7], f"{m_unidades[u][0]}_max: {m_unidades[u][7]}"
        model += ch_min*lpSum(saida[u]) <= m_unidades[u][1], f"{m_unidades[u][0]}_max: {int(m_unidades[u][1]/ch_min)}"
    if(minimo):
        #model += lpDot(saida[u], matriz_prof) >= m_unidades[u][6], f"{m_unidades[u][0]}_min: {m_unidades[u][6]}"
        model += ch_max*lpSum(saida[u]) >= m_unidades[u][1], f"{m_unidades[u][0]}_min: {math.ceil(m_unidades[u][1]/ch_max)}"

# Restrição do número total de professores
if(total):
    model += lpSum(saida*matriz_prof) <= n_total, f"Total: {n_total}"

# Restrições de regimes de trabalho
if(quarenta):
    for u in range(n_unidades):
        #model += lpSum(saida[u]*matriz_regimes[0]) <= 0, f"40 horas {m_unidades[u][0]} <= 10%"
        model += saida[u][4] + saida[u][5] - p_quarenta*lpSum(saida[u]) <= 0, f"40 horas {m_unidades[u][0]} <= 10%"
if(vinte):
    for u in range(n_unidades):
        #model += lpSum(saida[u]*matriz_regimes[1]) <= 0, f"20 horas {m_unidades[u][0]} <= 20%"
        model += saida[u][6] + saida[u][7] - p_vinte*lpSum(saida[u]) <= 0, f"20 horas {m_unidades[u][0]} <= 20%"

#print(model)
# Imprime os resultados no arquivo txt
original_stdout = sys.stdout
#filename = f"output{datetime.now().strftime('%H-%M-%S')}.txt"
filename = f"{modo}_n{n_unidades}-Ma{maximo}-Mi{minimo}-Peq{peq}-Tot{n_total if total else total}-Dif{dif_tam if total else dif}-40h{quarenta}-20h{vinte}--{datetime.now().strftime('%H-%M-%S')}.txt"
fileout =  open(filename, 'w')
sys.stdout = fileout
print(f'Modo: {modo}')
print(f'Unidades: {n_unidades}')
print(f'Maximo: {maximo}')
print(f'Minimo: {minimo}')
print(f'P-Equivalente: {peq}')
print(f'Total: {total} {n_total if total else ""}')
print(f'Dif: {dif} {dif_tam if dif else ""}')
print(f'Vinte: {vinte}')
print(f'Quarenta: {quarenta}')
print()
# Resolver o modelo
status = model.solve()

# Resultados
print(f"Situação: {model.status}, {LpStatus[model.status]}")
if peq:
    objetivo = f"{model.objective.value():.2f}"
else:
    objetivo = int(model.objective.value())
print(f"Objetivo: {objetivo} {'horas' if modo == 'tempo' else 'Prof-Equivalente' if peq else 'Professores'}")
print(f"Resolvido em {model.solutionTime} segundos")
print("")
# Formata resultados e calcula totais
resultados = np.full(n_unidades, '', dtype=object)
totais = np.full(n_unidades, 0, dtype=int)
qtdes = np.full((n_unidades, n_perfis), 0, dtype=int)
index = 0
perfil = 0
for var in model.variables():
    valor = int(var.value())
    resultados[index] += f"{valor:2d} "
    qtdes[index][perfil] = valor
    totais[index] += valor
    index += 1
    if(index >= n_unidades):
        index = 0
        perfil += 1
'''    
print("-------------Resultados-------------")
print("Unid.: x1 x2 x3 x4 x5 x6 x7 x8 total")
print("------------------------------------")
for r in range(n_unidades):
    print(f"{unidades[r]}: {resultados[r]} = {totais[r]:2d}")
print(f"Total geral: {np.sum(totais)}")
print("------------------------")
'''
print("Resultados")
print(f"--------+---------------------------------+-------+--------+")
#print(qtdes)
print(f"Unidade |  x1  x2  x3  x4  x5  x6  x7  x8 | total |   P-Eq |")
print(f"--------+---------------------------------+-------+--------+")
for u in range(n_unidades):
    print(f"{m_unidades[u][0]}   | " + " ".join([f"{qtdes[u][i]:3d}" for i in range(n_perfis)]) + f" |  {np.sum(qtdes[u]):4d} | {np.sum(qtdes[u]*matriz_peq):6.2f} |")
print(f"--------+---------------------------------+-------+--------+")
print(f"Total   | " + " ".join([f"{np.sum(qtdes, axis=0)[r]:3d}" for r in range(n_perfis)]) + f" |  {np.sum(qtdes):4d} | {np.sum(qtdes*matriz_peq):6.2f} |")
print(f"--------+---------------------------------+-------+--------+")
print("")

print("Parâmetros")
print(f"--------+-------------+-------------+-------------+-------------+-------------+---------+---------+----------+")
print(f"Unidade |     aulas   |    orient   |    gestao   |   diretor   |   coords.   |   40h   |   20h   | ch media |") #{'minim' if minimo else ''} {'maxim' if maximo or modo == 'tempo' else ''}")
print(f"--------+-------------+-------------+-------------+-------------+-------------+---------+---------+----------+")
for u in range(n_unidades):
    print(f"{m_unidades[u][0]}   | " 
    + " ".join([f"{np.sum(qtdes[u]*matriz_perfil[p-1]):4d} (+{(np.sum(qtdes[u]*matriz_perfil[p-1]) - m_unidades[u][p]):3d}) |" for p in range(1,6)])
    + f"  {((qtdes[u][4] + qtdes[u][5]) / np.sum(qtdes[u]))*100:5.2f}% |" #40h
    + f"  {((qtdes[u][6] + qtdes[u][7]) / np.sum(qtdes[u]))*100:5.2f}% |" #20h
    + f"  {m_unidades[u][1] / np.sum(qtdes[u]):7.3f} |"
    )
print(f"--------+-------------+-------------+-------------+-------------+-------------+---------+---------+----------+")

'''
# Formata restrições
restricoes = np.full(n_unidades, '', dtype=object)
index = 0
for name, constraint in model.constraints.items():
    restricoes[index] += f"{int(constraint.value()):5d} "
    index += 1
    if(index >= n_unidades):
        index = 0

print()
print("------------------Restrições-------------------")
print(f"Unid.: aulas orien gest. diret coord aulei {'minim' if minimo else ''} {'maxim' if maximo or modo == 'tempo' else ''}")
print("-------------------------------------------------------")
for r in range(n_unidades):
    print(f"{unidades[r]}: {restricoes[r]}")
'''
#for name, constraint in model.constraints.items():
#    print(f'{name}: valor {constraint.value():.0f}')

print()
print("----------------------------------------")

print(model)

#Imprime na tela    
sys.stdout = original_stdout
print(f"Situação: {model.status}, {LpStatus[model.status]}")
if peq:
    objetivo = f"{model.objective.value():.2f}"
else:
    objetivo = int(model.objective.value())
print(f"Objetivo: {objetivo} {'horas' if modo == 'tempo' else 'Prof-Equivalente' if peq else 'Professores'}")
print(f"Resolvido em {model.solutionTime} segundos")
print(f"Verifique o arquivo {filename} para o relatório completo")
print("")
fileout.close()
##input("Aperte Enter para finalizar")