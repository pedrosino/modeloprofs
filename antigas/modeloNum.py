from pulp import LpMaximize, LpMinimize, LpProblem, LpStatus, lpSum, LpVariable

# Definir o modelo
model = LpProblem(name="professores0.1", sense=LpMinimize)
##model = LpProblem(name="professores0.1", sense=LpMaximize)

# Variáveis de decisão
x = {i: LpVariable(name=f"x{i}", lowBound=0, cat="Integer") for i in range(1, 9)}

# Função objetivo
model += lpSum(x.values())

# Restrições
model += (10*x[2] + 12*x[3] + 16*x[4] + 20*x[5] + 24*x[6] + 10*x[7] + 12*x[8] >= 1080, "aulas")
model += (4*x[3] + 4*x[4] + 4*x[5] + 3*x[7] >= 30, "orientacoes")
model += (40*x[1] + 20*x[2] == 80, "gestao")
model += (x[2] == 2, "coords")
model += (x[1] == 1, "diretor")
model += (x[6] - 0.1*lpSum(x.values()) <= 0, "auleiro")
##model += (lpSum(x.values()) >= 68, "ch_min") # 68 = 1080 / 16 aulas
##model += (lpSum(x.values()) <= 90, "ch_max") # 90 = 1080 / 12 aulas

print(model)

# Resolver o modelo
status = model.solve()

# Resultados
print(f"Situação: {model.status}, {LpStatus[model.status]}")
print(f"Objetivo: {model.objective.value()}")

for var in x.values():
    print(f"{var.name}: {var.value()}")

for name, constraint in model.constraints.items():
    print(f"{name}: {constraint.value()}")