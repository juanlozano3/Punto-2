import pandas as pd
import json
import os

# Definir el directorio actual y la ruta del archivo Excel
current_directory = os.path.dirname(os.path.abspath(__file__))
excel_path = os.path.join(current_directory, "SOLUCIÓN.xlsx")

# Verificar si el archivo Excel existe
if not os.path.exists(excel_path):
    print(f"El archivo '{excel_path}' no se encuentra.")
else:
    try:
        # Procesar la hoja 'SOLUCIÓN'
        df_solucion = pd.read_excel(excel_path, sheet_name="SOLUCIÓN")
        data_solucion = []

        for _, row in df_solucion.iterrows():
            item = row[0]  # Nombre del item
            pallets = [f"Pallet {i+1}" for i, value in enumerate(row[1:]) if value == 1]
            data_solucion.append({
                "item": item,
                "pallets": ", ".join(pallets) if pallets else "Ninguno"
            })

        # Guardar el JSON de 'SOLUCIÓN'
        with open(os.path.join(current_directory, "data_solucion.json"), "w") as f:
            json.dump(data_solucion, f, indent=4)

        # Procesar la hoja 'USO'
        df_uso = pd.read_excel(excel_path, sheet_name="USO", header=None)

        # Crear un JSON con solo los pallets y sus pesos
        data_uso = [
            {"pallet": f"Pallet {i+1}", "peso": int(peso)}
            for i, peso in enumerate(df_uso.iloc[1, 1:])
        ]

        # Guardar el JSON de 'USO' con el formato correcto
        with open(os.path.join(current_directory, "data_uso.json"), "w") as f:
            json.dump(data_uso, f, indent=4)

        print("Archivos JSON generados exitosamente.")

    except Exception as e:
        print(f"Error al procesar el archivo Excel: {e}")
