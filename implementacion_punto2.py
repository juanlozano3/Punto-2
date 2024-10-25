from gurobipy import *
import pandas as pd
import os
import numpy as np
import matplotlib.pyplot as plt
import time

# Leer datos del Excel original
current_directory = os.path.dirname(os.path.abspath(__file__))
file_name = os.path.join(current_directory, "Caviana.xlsx")
df = pd.read_excel(io=file_name, sheet_name="Caviana Cargo")

Items = df['Item'].values
pesos = df['Peso'].values

# Parámetros
aviones = 5
pallets = 16
peso_maximo = 1500
total_pallets = aviones * pallets

# Solución inicial: Matriz identidad
def generar_identidad(filas, columnas):
    identidad = np.eye(filas, columnas, dtype=int)
    return identidad

solucion_inicial = generar_identidad(120, 120)

# Inicializa una lista para almacenar las columnas generadas
columnas_generadas = []

# Problema Maestro
modelMP = Model("Problema Maestro")
modelMP.Params.OutputFlag = 0

# Variables
x = modelMP.addVars(len(solucion_inicial), vtype=GRB.CONTINUOUS, name="x")

# Restricciones
modelMP.addConstrs(
    (quicksum(solucion_inicial[i, j] * x[j] for j in range(len(solucion_inicial))) == 1
     for i in range(len(solucion_inicial))),
    name="Restricciones"
)

# Función Objetivo
modelMP.setObjective(quicksum(x[i] for i in range(len(solucion_inicial))), GRB.MINIMIZE)

# Problema Auxiliar
modelAP = Model("Problema Auxiliar")
modelAP.Params.OutputFlag = 0

# Variables auxiliares
y = modelAP.addVars(len(Items), vtype=GRB.BINARY, name="y")

# Restricción del peso máximo
modelAP.addConstr(
    quicksum(pesos[i] * y[i] for i in range(len(Items))) <= peso_maximo, name="Peso"
)

# Ciclo principal
costos_reducidos_en_iteraciones = []
iteraciones = 0
columnas = 120
start_time = time.time()

# Optimización iterativa
while True:
    iteraciones += 1
    modelMP.optimize()

    if modelMP.status != GRB.OPTIMAL:
        print("No se encontró una solución óptima.")
        break

    print(f"Pallets usados: {modelMP.ObjVal}")

    # Obtener duales
    duals = modelMP.getAttr("Pi", modelMP.getConstrs())

    # Establecer la función objetivo del problema auxiliar
    modelAP.setObjective(
        quicksum(duals[i] * y[i] for i in range(len(Items))), GRB.MAXIMIZE
    )

    modelAP.optimize()

    if modelAP.status == GRB.INFEASIBLE:
        modelAP.computeIIS()
        modelAP.write("infactible_model.ilp")
        break

    costo_reducido = 1 - modelAP.getObjective().getValue()
    costos_reducidos_en_iteraciones.append(costo_reducido)

    if costo_reducido >= -1e-10:
        print("Costo reducido no negativo. Terminando...")
        break

    # Obtener la nueva columna generada
    nueva_columna = [y[i].X for i in range(len(Items))]
    if any(nueva_columna):
        columnas_generadas.append([f"x[{columnas}]"] + nueva_columna)

    # Añadir la nueva columna al problema maestro
    new_col = Column(nueva_columna, modelMP.getConstrs())
    modelMP.addVar(vtype=GRB.CONTINUOUS, column=new_col, obj=1, name=f"x[{columnas}]")
    columnas += 1
    modelMP.update()

end_time = time.time()

# Guardar todas las columnas generadas en un Excel, con el nombre del pallet en la primera columna
df_columnas = pd.DataFrame(columnas_generadas)
df_columnas.columns = ["Pallet"] + [f"Item_{i+1}" for i in range(len(Items))]

output_excel_path = os.path.join(current_directory, "ColumnasGeneradas.xlsx")
df_columnas.to_excel(output_excel_path, sheet_name="Todas las Columnas", index=False)
print(f"Todas las columnas generadas se han guardado en {output_excel_path}")

# **Solución entera**
for v in modelMP.getVars():
    v.setAttr("Vtype", GRB.INTEGER)
print("Heuristic integer master problem")
modelMP.optimize()

if modelMP.status != GRB.OPTIMAL:
    print("No se encontró una solución óptima para el problema entero.")
else:
    print(f"Pallets usados: {modelMP.objVal}")

    # Variables con valor 1 en la solución entera
    variables_valor_1 = {
        v.VarName: v.x for v in modelMP.getVars() if abs(v.x - 1) < 1e-6
    }

    # Guardar variables con valor 1 en pallet_var.xlsx
    df_valor_1 = pd.DataFrame(list(variables_valor_1.items()), columns=["Variable", "Valor"])
    output_var_excel = os.path.join(current_directory, "pallet_var.xlsx")
    df_valor_1.to_excel(output_var_excel, sheet_name="Variables Valor 1", index=False)
    print(f"Variables con valor 1 se han guardado en {output_var_excel}")

    # Guardar configuración de pallets utilizados en PalletsUsados.xlsx
    configuracion_pallets = {
        f"Pallet_{idx}": df_columnas.iloc[idx].to_dict()
        for idx in range(len(df_columnas)) if f"x[{idx}]" in variables_valor_1
    }

    configuracion_df = pd.DataFrame.from_dict(configuracion_pallets, orient='index')
    output_config_excel = os.path.join(current_directory, "PalletsUsados.xlsx")
    configuracion_df.to_excel(output_config_excel, sheet_name="Pallets Utilizados", index=True)
    print(f"La configuración de los pallets utilizados se ha guardado en {output_config_excel}")

# Visualización del costo reducido por iteración
plt.figure(figsize=(10, 6))
plt.plot(range(1, len(costos_reducidos_en_iteraciones) + 1), costos_reducidos_en_iteraciones, marker='o')
plt.xlabel('Iteraciones')
plt.ylabel('Costo Reducido')
plt.title('Evolución del Costo Reducido por Iteración')
plt.grid(True)
plt.show()

# Reporte final
print("--------------------------------------------------------------------------------")
print(f"Total de iteraciones: {iteraciones}")
print(f"Total de columnas generadas: {len(columnas_generadas)}")
print(f"Tiempo total de optimización: {end_time - start_time:.4f} segundos")
print(f"Pallets usados: {modelMP.objVal}")
