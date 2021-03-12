import os
from random import choice, random, uniform, randint
import xlsxwriter
from tkinter import Tk, filedialog
from string import ascii_lowercase
from json import load
from numpy import array, transpose

titulos = [
        "Nro",
        "Coopac",
        "N° Socios",
        "Total de Activos Brutos al 31/12/2020 (S/MM)",
        "No. Agencias Total",
        "Agencias Abiertas",
        "¿Abierta Principal?",
        "Capta CTS",
        "Representatividad de MN de Fondos Disponibles",
        "Representatividad de MN de Obligaciones CP",
        "Fondos Disponibles / Total de Activos Brutos (S/ MM)",
        "Fondos Disponibles sin Restricción (MM)",
        "Depósitos de Socios  Y COOPAC CP(pasivos, solo capital) (MM)",
        "Obligaciones CP(pasivos, solo capital) (MM)",
        "Fondos Disponibles / Depósitos Socios",
        "Depósitos 10 principales depositantes(MM)",
        "% Depósitos 10 principales depositantes de Depósitos Totales (MM)",
        "Ratio de Liquidez en MN (Trimestral)",
        "Ratio de Liquidez en ME (Trimestral)"
    ]

#Matriz de valores por llenar | Especificar cuántas filas deben haber en la salida global
valores= [[] for i in range(19)]
A_files = [0,0,0,0,0]
condicion = ["Si", "No"]
#Lectura de información en base a excel llamado desde el vector
if(len(A_files) != 0):
    for i in range(99):
        #Numero de fila
        valores[0].append(int(i+1))
        #Valor de nombre de COOPAC - cell(fila,columna)
        valores[1].append("Valor de prueba"+str(i))
        #Valor de Nº Socios
        value = randint(50,501)
        valores[2].append(int(value))
        #Valor de Total de Activos Brutos al 31/12/2020
        value = uniform(1000000,5000000)
        valores[3].append(value)
        #Valor de Nº Agencias
        value = randint(1,7)
        valores[4].append(value)
        #Valor de Nº Agencias Abiertas
        value2 = randint(1, value)
        valores[5].append(value2)
        #Valor de Agencia Principal Abierta?
        value = choice(condicion)
        valores[6].append(value)
        #Valor Captan CTS?
        value = choice(condicion)
        valores[7].append(value)
        #Valor Fondos Disponibles (Cálculo)
        value = round(uniform(0.30, 0.70), 2)
        valores[8].append(round(value,2)) #Se redondea a 2 decimales hasta nuevo aviso
        #Valor Obligaciones CP (Cálculo)
        value = round(uniform(0.30, 0.90), 2)
        valores[9].append(round(value,2)) #Se redondea a 2 decimales hasta nuevo aviso
        #Valor Fondos Disponibles / Total de Activos Brutos (Cálculo)
        value = round(uniform(1000000.01, 25000000.99), 2)
        valores[10].append(value) #Se redondea a 2 decimales hasta nuevo aviso
        #Valor Fondos Disponibles Sin Restricción (Cálculo)
        value = round(uniform(2000000.01, 35000000.99), 2)
        valores[11].append(value) #Se redondea a 2 decimales hasta nuevo aviso
        #Valor Depósitos de Socios y COOPAC CP (Cálculo)
        value = round(uniform(1000000.01, 25000000.99), 2)
        valores[12].append(value) #Se redondea a 2 decimales hasta nuevo aviso
        #Valor Obligaciones CP
        value = round(uniform(1000000.01, 25000000.99), 2)
        valores[13].append(value)
        #Valor Fondos Disponibles / Depósitos Socios (Cálculo)
        value = round(uniform(0.10, 0.60), 2)
        valores[14].append(round(value,2)) #Se redondea a 2 decimales hasta nuevo aviso
        #Valor Depósitos 10 Principales depositantes
        value = round(uniform(1000000.01, 25000000.99), 2)
        valores[15].append(round(value,2)) #Se redondea a 2 decimales hasta nuevo aviso
        #Valor % Depósitos 10 Principales depositantes
        value = round(uniform(0.10, 0.60), 2)
        valores[16].append(round(value,2)) #Se redondea a 2 decimales hasta nuevo aviso
        #Valor Liquidez MN
        value = round(uniform(0.10, 0.60), 2)
        valores[17].append(round(value,2)) #Se redondea a 2 decimales hasta nuevo aviso
        #Valor Liquidez ME
        value = round(uniform(0.10, 0.60), 2)
        valores[18].append(round(value,2)) #Se redondea a 2 decimales hasta nuevo aviso

R_file = 0            

for dirName, subdirList, fileList in os.walk("./resultado"):
    for filename in fileList:
        #print(filename)                                                    
        if ".xlsx" in filename.lower() or ".xlsm" in filename.lower():
            if not filename.startswith('~$'):
                R_file = os.path.join(dirName,filename)

#print(R_file)
workbook = xlsxwriter.Workbook("./resultado/Monitor de Liquidez Prueba.xlsx", {'strings_to_numbers': True})
worksheet = workbook.add_worksheet("Liquidez")
worksheetResumen = workbook.add_worksheet("Resumen")
#Letras hasta la cantidad de columnas necesarias
columnas_titulo = []
for c in ascii_lowercase:
    columnas_titulo.append(c)

#Celdas de Título y Formato
cell_format = workbook.add_format({'bold': False, 'font_color': 'white'})
cell_format_bold = workbook.add_format({'bold': True})
cell_format.set_bg_color('#003366')
cell_format.set_align('center')
cell_format.set_text_wrap()
cell_format.set_valign('vcenter')
for i in range(len(valores)):
    worksheet.write(columnas_titulo[i+1].upper()+str(2), titulos[i], cell_format)
worksheet.write('B1',"Información expresada en S/", cell_format_bold)
row = 2
col = 2

# for col, data in enumerate(valores):
#     worksheet.write_column(row, col+1, data)

valores = array(valores)
valores = transpose(valores)
valores = valores.tolist()

worksheet.add_table('B3:T'+str(3+99), {'data': valores, 'header_row': 0})

worksheet.set_column(2, 2, 40) #Tamaño de columna nombre coopac
worksheet.set_column(3, 19, 15) #Tamaño de columna general

#Grafico de Liquidez en MN
chart = workbook.add_chart({'type': 'bar'})
chart.add_series({'values': '=Liquidez!S3:S'+str(3+99-1)})
worksheetResumen.insert_chart('C1', chart)

#Grafico de Liquidez en ME
chart = workbook.add_chart({'type': 'bar'})
chart.add_series({'values': '=Liquidez!T3:T'+str(3+99-1)})
worksheetResumen.insert_chart('J1', chart)

workbook.close()
#Matriz de resultados de análisis
