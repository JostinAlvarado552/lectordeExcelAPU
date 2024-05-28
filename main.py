import pandas as pd
import xlwings as xw

def process_id(id_value):
    detalles = []

    try:
        # Inicializar la aplicación Excel en segundo plano
        app = xw.App(visible=False)
        wb = app.books.open('apu.xlsm')
        sheet = wb.sheets['ANALISIS']
        sheet.range('C5').value = id_value
        app.calculate()

        # Proceso para el rango 18-38 (EQUIPOS)
        categoria = 1
        for i in range(18, 38):
            nombre = sheet.range(f'D{i}').value
            if nombre and nombre != 0.0:
                detalles.append(f"({id_value}, '{nombre}', {categoria})")

        # Proceso para el rango 41-61 (MANO DE OBRA)
        categoria = 2
        for i in range(41, 61):
            nombre = sheet.range(f'D{i}').value
            if nombre and nombre != 0.0:
                detalles.append(f"({id_value}, '{nombre}', {categoria})")

        # Proceso para el rango 64-84 (MATERIALES)
        categoria = 3
        for i in range(64, 84):
            nombre = sheet.range(f'D{i}').value
            if nombre and nombre != 0.0:
                detalles.append(f"({id_value}, '{nombre}', {categoria})")

        # Proceso para el rango 87-92 (TRANSPORTE)
        categoria = 4
        for i in range(87, 92):
            nombre = sheet.range(f'D{i}').value
            if nombre and nombre != 0.0:
                detalles.append(f"({id_value}, '{nombre}', {categoria})")

        wb.save()
        wb.close()

        # Cerrar la aplicación Excel en segundo plano
        app.quit()

    except Exception as e:
        return {'error': str(e)}

    return detalles


def main():
    all_details = []
    for id_value in (1,3):
        result = process_id(id_value)
        if isinstance(result, list):
            all_details.extend(result)
            print(f"Processed ID: {id_value}")  # Mostrar progreso
        else:
            print(f"Error processing ID {id_value}: {result['error']}")

    if all_details:
        insert_command = f"INSERT INTO automatizacion_apus_apu (rubro_id, contenido, categoria) VALUES {', '.join(all_details)}"

        # Guardar el comando INSERT en un archivo SQL
        with open("script_apu3.sql", "w") as file:
            file.write(insert_command)
        print("SQL script has been saved to script_apu.sql")
    else:
        print("No data to insert")


if __name__ == "__main__":
    main()

#
# def generar_script_insercion(file_path, sheet_name, output_file):
#     # Leer los datos desde las celdas A6, E6 y G6
#     df = pd.read_excel(file_path, sheet_name=sheet_name, header=None, usecols="A,E,G")
#
#     # Renombrar las columnas para facilitar el trabajo
#     df.columns = ['id', 'concepto', 'unidad']
#
#     # Filtrar las filas que comienzan en la fila 6 (índice 5 en pandas)
#     df = df.iloc[5:]
#
#     # Generar el script de inserción
#     insert_script = ""
#
#     for index, row in df.iterrows():
#         id = row['id']
#         concepto = row['concepto']
#         unidad = row['unidad']
#
#         # Escapar comillas dentro de las cadenas de texto
#         concepto = concepto.replace("'", "''") if pd.notnull(concepto) else None
#         unidad = unidad.replace("'", "''") if pd.notnull(unidad) else None
#
#         # Manejar valores nulos para evitar errores en el script SQL
#         concepto_value = f"'{concepto}'" if concepto is not None else 'NULL'
#         unidad_value = f"'{unidad}'" if unidad is not None else 'NULL'
#
#         insert_script += f"INSERT INTO automatizacion_apus_rubros (id, concepto, unidad) VALUES ({id}, {concepto_value}, {unidad_value});\n"
#
#     # Escribir el script de inserción a un archivo
#     with open(output_file, 'w') as file:
#         file.write(insert_script)
#
#     print("Script de inserción generado con éxito.")
#
# if __name__ == "__main__":
#     # Especifica la ruta al archivo Excel, el nombre de la hoja y el archivo de salida
#     file_path = 'C:/Users/josti/Documents/APU-EXCEL v7.6 Mayo 2024/APU-EXCEL v 7.6 RENDIMIENTOS GENERALES.xlsm'  # Reemplaza esto con la ruta real a tu archivo Excel
#     sheet_name = 'RUBROS'
#     output_file = 'insert_script.sql'
#
#     # Generar el script de inserciónza
#     generar_script_insercion(file_path, sheet_name, output_file)
