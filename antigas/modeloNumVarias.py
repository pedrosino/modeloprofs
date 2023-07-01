from pulp import * #LpMaximize, LpMinimize, LpProblem, LpStatus, lpSum, LpVariable

# Definir o modelo
model = LpProblem(name="professores0.1", sense=LpMinimize)
##model = LpProblem(name="professores0.1", sense=LpMaximize)

# Dados do problema
#unidades = range(1, 6)
# Quantidade de aulas e orientações por perfil
aulas_perfil =       [0, 10, 12, 16, 20, 24, 10, 12]
orientacoes_perfil = [0,  0,  4,  4,  4,  0,  3,  0]
gestao_perfil =     [40, 20,  0,  0,  0,  0,  0,  0]
 
# Variáveis de decisão
x = {i: LpVariable(name=f"x{i}", lowBound=0, cat="Integer") for i in range(1, 9)}

# Função objetivo
model += lpSum(x.values())

# Restrições
model += (lpDot(aulas_perfil, list(x.values())) >= 1080, "aulas")
model += (lpDot(orientacoes_perfil, list(x.values())) >= 30, "orientacoes")
#model += (lpDot(gestao_perfil, list(x.values())) == 80, "gestao")
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