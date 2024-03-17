#.\venv\Scripts\activate

#usar cuando el formato del excel es el  utilizado para el TAS3560
import openpyxl
from tabulate import tabulate

excel_dataframe = openpyxl.load_workbook("TV-456.xlsx")

nombre_hoja = excel_dataframe.sheetnames[1] 

dataframe = excel_dataframe[nombre_hoja]
print(dataframe)

valores_combinados = []


# Recorriendo ambas columnas simultáneamente
for celda1, celda2 in zip(dataframe['A'], dataframe['B']):
        valor_columna_1 = str(celda1.value * 1)  if celda1.value is not None else ''
        #Existia un fallo generar los valores, tuve que redondear con round(celda2.value * 1000, 2) y se soluciono
        valor_columna_2 = str(round(celda2.value * 1000, 2))  if celda2.value is not None else ''
        valores_combinados.append(f"{valor_columna_1}:{valor_columna_2.rstrip('.0')},")  # Formatear la cadena correctamente
    
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