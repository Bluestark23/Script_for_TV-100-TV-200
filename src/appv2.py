
#usar cuando el formato del excel NO es el  utilizado para el TAS3560


import openpyxl

# Cargar el libro de trabajo de Excel
excel_dataframe = openpyxl.load_workbook("TV-455.xlsx")

# Seleccionar la hoja de trabajo
nombre_hoja = excel_dataframe.sheetnames[1]
dataframe = excel_dataframe[nombre_hoja]

valores_combinados = []

# Recorrer ambas columnas simultáneamente
for celda1, celda2 in zip(dataframe['A'], dataframe['B']):
    # Convertir el valor de la celda a cadena
    valor_columna_1 = str(celda1.value) if celda1.value is not None else ''

    # Verificar si el valor de la celda en la columna B es numérico
    if isinstance(celda2.value, (int, float)):
        # Realizar operaciones matemáticas solo si el valor es numérico
        valor_columna_2 = round(celda2.value , 2)
    else:
        # Si el valor no es numérico, asignar una cadena vacía
        valor_columna_2 = None

    # Convertir el valor de la columna 2 a cadena si es numérico
    if valor_columna_2 is not None:
        valor_columna_2 = str(valor_columna_2)

    valores_combinados.append(f"{valor_columna_1}:{valor_columna_2},")

# Imprimir los valores combinados
for valor in valores_combinados:
    print(valor)

# Abre un archivo llamado "main.txt" en modo de escritura ('w')
with open("main.txt", "w") as archivo:
    # Escribir los datos de valores_combinados divididos en grupos de 12 filas
    for i in range(0, len(valores_combinados), 12):
        fila = valores_combinados[i:i+12]  # Obtener los siguientes 12 valores
        fila_como_cadena = "".join(fila)  # Convertir la lista de valores en una cadena sin ningún separador
        archivo.write(fila_como_cadena)  # Escribir la fila en el archivo con un salto de línea al final
