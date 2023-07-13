from pulp import * #LpMaximize, LpMinimize, LpProblem, LpStatus, lpSum, LpVariable
import numpy as np
from datetime import time, datetime
import sys, getopt
import pandas as pd
import math

# Read information from excel file
df_todas = pd.read_excel('importar.xlsx', sheet_name=['unidades','perfis'])
dept_matrix = df_todas['unidades'].to_numpy()
profile_matrix = df_todas['perfis'].to_numpy()
profile_matrix = np.delete(profile_matrix, 0, axis=1)

# Settings
maximum = False #carga horária média máxima por unidade 
ch_min = 12
minimum = False #carga horária média mínima por unidade
ch_max = 16
max_total = False  #número máximo total
min_total = False #número mínimo total
quarenta = False
p_quarenta = 0.1
vinte = False
p_vinte = 0.2

# Dados do problema
n_profiles = 8
n_constraints = 5
n_depts = len(dept_matrix)#5

opts, args = getopt.getopt(sys.argv[1:],"iadvqn:",["chmin=","chmax=","tmax=","tmin="])
for opt, arg in opts:
    if opt == '-i':
        minimum = True
    elif opt == '-a':
        maximum = True
    elif opt == '-t':
        max_total = True
        n_max_total = int(arg)
    elif opt == '-n':
        n_depts = int(arg)
    elif opt == '-d':
        dif = True
    elif opt == '-v':
        vinte = True
    elif opt == '-q':
        quarenta = True
    elif opt == '--chmin':
        minimum = True
        ch_min = int(arg)
    elif opt == '--chmax':
        maximum = True
        ch_max = int(arg)
    elif opt == '--tmin':
        min_total = True
        n_min_total = int(arg)
    elif opt == '--tmax':
        max_total = True
        n_max_total = int(arg)


# Cost matrix
teacher_matrix = np.array([1, 1, 1, 1, 1, 1, 1, 1])
matriz_peq = np.array([1.65, 1.65, 1.65, 1.65, 1, 1, 0.6, 0.6])
                             
connectors = np.array([">=", ">=", ">=", "==", "==", "<="])
constraint_names = np.array(['aulas','orientacoes','gestao','diretor','coords'])
 
# Decision variables
# names
names = [str(p) + "." + dept_matrix[u][0] for u in range(n_depts) for p in range(1, n_profiles+1) ]
# variables
nx = LpVariable.matrix("x", names, cat="Integer", lowBound=0)
output = np.array(nx).reshape(n_depts, n_profiles)

# Model definition
# Objective function
model = LpProblem(name="TeachersDepts", sense=LpMinimize)
model += lpSum(output*teacher_matrix)

# Constraints
for r in range(n_constraints):
    for u in range(n_depts):
        match connectors[r]:
            case ">=":
                model += lpDot(output[u], profile_matrix[r]) >= dept_matrix[u][r+1], constraint_names[r] + " " + dept_matrix[u][0]#"r " + str(i)
            case "==":
                model += lpDot(output[u], profile_matrix[r]) == dept_matrix[u][r+1], constraint_names[r] + " " + dept_matrix[u][0]#"r " + str(i)
            case "<=":
                model += lpDot(output[u], profile_matrix[r]) <= dept_matrix[u][r+1], constraint_names[r] + " " + dept_matrix[u][0]#"r " + str(i)        

# Constraints on min and max average workload per teacher
for u in range(n_depts):
    if(maximum):
        model += ch_max*lpSum(output[u]) >= dept_matrix[u][1], f"{dept_matrix[u][0]}_chmax: {math.ceil(dept_matrix[u][1]/ch_max)}"
    if(minimum):
        model += ch_min*lpSum(output[u]) <= dept_matrix[u][1], f"{dept_matrix[u][0]}_chmin: {int(dept_matrix[u][1]/ch_min)}"

# Constraints on the global number of teachers
if(max_total):
    model += lpSum(output*teacher_matrix) <= n_max_total, f"TotalMax: {n_max_total}"
if(min_total):
    model += lpSum(output*teacher_matrix) >= n_min_total, f"TotalMin: {n_min_total}"

# Constraints on certain profiles
if(quarenta):
    for u in range(n_depts):
        model += output[u][4] + output[u][5] - p_quarenta*lpSum(output[u]) <= 0, f"40 horas {dept_matrix[u][0]} <= 10%"
if(vinte):
    for u in range(n_depts):
        model += output[u][6] + output[u][7] - p_vinte*lpSum(output[u]) <= 0, f"20 horas {dept_matrix[u][0]} <= 20%"

# Solve model
status = model.solve()

# Results
print(f"Situation: {model.status}, {LpStatus[model.status]}")
objective = int(model.objective.value())
print(f"Objective: {objective} teachers")
print(f"Solved in {model.solutionTime} seconds")
print("")
# Print formatted results
results = np.full(n_depts, '', dtype=object)
totals = np.full(n_depts, 0, dtype=int)
qtys = np.full((n_depts, n_profiles), 0, dtype=int)
index = 0
profile = 0
for var in model.variables():
    valor = int(var.value())
    results[index] += f"{valor:2d} "
    qtys[index][profile] = valor
    totals[index] += valor
    index += 1
    if(index >= n_depts):
        index = 0
        profile += 1

print("results")
print(f"--------+---------------------------------+-------+--------+")
#print(qtys)
print(f"Unidade |  x1  x2  x3  x4  x5  x6  x7  x8 | total |   P-Eq |")
print(f"--------+---------------------------------+-------+--------+")
for u in range(n_depts):
    print(f"{dept_matrix[u][0]}   | " + " ".join([f"{qtys[u][p]:3d}" for p in range(n_profiles)]) + f" |  {np.sum(qtys[u]):4d} | {np.sum(qtys[u]*matriz_peq):6.2f} |")
print(f"--------+---------------------------------+-------+--------+")
print(f"Total   | " + " ".join([f"{np.sum(qtys, axis=0)[p]:3d}" for p in range(n_profiles)]) + f" |  {np.sum(qtys):4d} | {np.sum(qtys*matriz_peq):6.2f} |")
print(f"--------+---------------------------------+-------+--------+")
print("")

print("Parâmetros")
print(f"--------+-------------+-------------+-------------+-------------+-------------+----------+----------+----------+")
print(f"Unidade |     aulas   |    orient   |    gestao   |   diretor   |   coords.   |    40h   |    20h   | ch media |") #{'minim' if minimo else ''} {'maxim' if maximo or modo == 'tempo' else ''}")
print(f"--------+-------------+-------------+-------------+-------------+-------------+----------+----------+----------+")
for u in range(n_depts):
    print(f"{dept_matrix[u][0]}   | " 
    + " ".join([f"{np.sum(qtys[u]*profile_matrix[p]):4d} (+{(np.sum(qtys[u]*profile_matrix[p]) - dept_matrix[u][p+1]):3d}) |" for p in range(5)])
    + f"  {((qtys[u][4] + qtys[u][5]) / np.sum(qtys[u]))*100:6.2f}% |" #40h
    + f"  {((qtys[u][6] + qtys[u][7]) / np.sum(qtys[u]))*100:6.2f}% |" #20h
    + f"  {dept_matrix[u][1] / np.sum(qtys[u]):7.3f} |"
    )
print(f"--------+-------------+-------------+-------------+-------------+-------------+----------+----------+----------+")
print(f"Total   | " + " ".join([f"{np.sum(qtys*profile_matrix[p]):4d} (+{np.sum(qtys*profile_matrix[p]) - int(np.sum(dept_matrix, axis=0)[p+1]):3d}) |" for p in range(5)])
+ f"  {(np.sum(qtys, axis=0)[4] + np.sum(qtys, axis=0)[5]) / np.sum(qtys)*100:6.2f}% |"
+ f"  {(np.sum(qtys, axis=0)[6] + np.sum(qtys, axis=0)[7]) / np.sum(qtys)*100:6.2f}% |"
+ f"  {np.sum(dept_matrix, axis=0)[1]/np.sum(qtys):7.3f} |"
)
print()
print("------------------Modelo:------------------")
print(model)