import os
#from random import choice, random, uniform, randint
import xlsxwriter
import xlrd
from tkinter import Tk, filedialog
from string import ascii_lowercase
from json import load
from numpy import array, transpose

titulos = [
        "Nro",
        "Coopac",
        "N° Socios",
        "Total de Activos Brutos al 31/12/2020",
        "No. Agencias Total",
        "Agencias Abiertas",
        "¿Abierta Principal?",
        "Capta CTS",
        "Representatividad de MN de Fondos Disponibles",
        "Representatividad de MN de Obligaciones CP",
        "Fondos Disponibles / Total de Activos Brutos",
        "Fondos Disponibles sin Restricción",
        "Depósitos de Socios  Y COOPAC CP(pasivos, solo capital)",
        "Obligaciones CP(pasivos, solo capital)",
        "Fondos Disponibles / Depósitos Socios",
        "Depósitos 10 principales depositantes",
        "% Depósitos 10 principales depositantes de Depósitos Totales",
        "Ratio de Liquidez en MN (Trimestral)",
        "Ratio de Liquidez en ME (Trimestral)"
    ]

#Solicitar carpeta al usuario
print("Seleccione la carpeta a analizar...")

#Codigo de seleccion de carpeta
root = Tk()
root.withdraw()
folder_selected = filedialog.askdirectory()

A_files = []

#Matriz de valores por llenar | Especificar cuántas filas deben haber en la salida global
valores= [[] for i in range(19)]
for dirName, subdirList, fileList in os.walk(folder_selected):                        
    for filename in fileList:   
        if ".xlsx" in filename.lower() or ".xlsm" in filename.lower(): 
            if not filename.startswith('~$'):
                A_files.append(os.path.join(dirName,filename)) 


#Arreglos de obligaciones a CP de <10%, >=10% y <=20% ... >50%
oblig1 = 0
oblig2 = 0
oblig3 = 0
oblig4 = 0
oblig5 = 0
oblig6 = 0
#Arreglos de Top 10 Depositantes <10%, >=10% y <=20% ... >50%
deposit1 = 0
deposit2 = 0
deposit3 = 0
deposit4 = 0
deposit5 = 0
deposit6 = 0
#Arreglos de liquidez con rangos de >=8%, <8% y >=20%,  <20%
liq_critico_mn = 0
liq_bajo_mn = 0               
liq_normal_mn = 0     
liq_critico_me = 0
liq_bajo_me = 0               
liq_normal_me = 0    
fondos_disp = 0
condicion = ["Si", "No"]
#Lectura de información en base a excel llamado desde el vector
if(len(A_files) != 0):
    for i in range(len(A_files)):
        #Numero de fila
        valores[0].append(int(i+1))
        #Lectura de excel
        workbook = xlrd.open_workbook(A_files[i])
        worksheet = workbook.sheet_by_name('Requerimiento') #Nombre de hoja a leer del Excel
        #Valor de nombre de COOPAC - cell(fila,columna)
        value = worksheet.cell(4, 4).value
        valores[1].append(value)
        #Valor de Nº Socios
        value = worksheet.cell(10, 4).value
        valores[2].append(int(value))
        #Valor de Total de Activos Brutos al 31/12/2020
        value = worksheet.cell(11, 4).value
        valores[3].append(value)
        #Valor de Nº Agencias
        value = worksheet.cell(12, 4).value
        valores[4].append(value)
        #Valor de Nº Agencias Abiertas
        value = worksheet.cell(13, 4).value
        valores[5].append(value)
        #Valor de Agencia Principal Abierta?
        value = worksheet.cell(14, 4).value
        valores[6].append(value)
        #Valor Captan CTS?
        value = worksheet.cell(15, 4).value
        valores[7].append(value)
        #Valor Fondos Disponibles (Cálculo)
        cal1 = worksheet.cell(25, 5).value
        cal2 = worksheet.cell(25, 7).value
        #fondos_disp.append(round(cal2,2))
        value = cal1/cal2 #Fondos disponibles -> Tabla 1 total en MN / total
        valores[8].append(round(value,2)) #Se redondea a 2 decimales hasta nuevo aviso
        #Valor Obligaciones CP (Cálculo)
        cal1 = worksheet.cell(63, 5).value
        cal2 = worksheet.cell(63, 7).value
        value = cal1/cal2 #
        if value < 0.10:
            oblig1 += 1
        elif value >=0.10 and value <0.20:
            oblig2 += 1
        elif value >= 0.20 and value <0.30:
            oblig3 += 1
        elif value >= 0.30 and value <0.40:
            oblig4 += 1
        elif value >= 0.40 and value <0.50:
            oblig5 += 1
        elif value >= 0.50:
            oblig6 += 1
        valores[9].append(round(value,2)) #Se redondea a 2 decimales hasta nuevo aviso
        #Valor Fondos Disponibles / Total de Activos Brutos (Cálculo)
        cal1 = worksheet.cell(25, 7).value
        cal2 = worksheet.cell(11, 4).value #Por verificar
        value = cal1/cal2 #
        valores[10].append(value) #Se redondea a 2 decimales hasta nuevo aviso
        #Valor Fondos Disponibles Sin Restricción (Cálculo)
        cal1 = worksheet.cell(25, 7).value
        cal2 = worksheet.cell(53, 7).value 
        value = cal1 - cal2 #
        valores[11].append(value) #Se redondea a 2 decimales hasta nuevo aviso
        #Valor Depósitos de Socios y COOPAC CP (Cálculo)
        cal1 = worksheet.cell(59, 7).value
        cal2 = worksheet.cell(60, 7).value 
        value = cal1 + cal2 #
        valores[12].append(value) #Se redondea a 2 decimales hasta nuevo aviso
        #Valor Obligaciones CP
        value = worksheet.cell(63, 7).value
        valores[13].append(value)
        #Valor Fondos Disponibles / Depósitos Socios (Cálculo)
        cal1 = worksheet.cell(25, 7).value
        cal2 = worksheet.cell(59, 7).value
        cal3 = worksheet.cell(67,7).value
        value = cal1 / (cal2+cal3) #
        valores[14].append(round(value,2)) #Se redondea a 2 decimales hasta nuevo aviso
        #Valor Depósitos 10 Principales depositantes
        value = worksheet.cell(75, 7).value
        valores[15].append(round(value,2)) #Se redondea a 2 decimales hasta nuevo aviso
        #Valor % Depósitos 10 Principales depositantes
        cal1 = worksheet.cell(75, 7).value
        cal2 = worksheet.cell(59, 7).value
        cal3 = worksheet.cell(60,7).value
        cal4 = worksheet.cell(67,7).value
        cal5 = worksheet.cell(68,7).value
        value = cal1 / (cal2+cal3+cal4+cal5) #
        if value < 0.50:
            deposit1 += 1
        elif value >=0.50 and value <0.60:
            deposit2 += 1
        elif value >= 0.60 and value <0.70:
            deposit3 += 1
        elif value >= 0.70 and value <0.80:
            deposit4 += 1
        elif value >= 0.80 and value <0.90:
            deposit5 += 1
        elif value >= 0.90:
            deposit6 += 1
        valores[16].append(round(value,2)) #Se redondea a 2 decimales hasta nuevo aviso
        #Valor Liquidez MN
        cal1 = worksheet.cell(25, 5).value
        cal2 = worksheet.cell(63, 5).value
        value = cal1 / cal2 #
        if (value < 0.08):
            liq_critico_mn += 1
        elif (value >= 0.08 and value <= 0.2):
            liq_bajo_mn += 1
        elif (value > 0.2):
            liq_normal_mn += 1
        valores[17].append(round(value,2)) #Se redondea a 2 decimales hasta nuevo aviso
        #Valor Liquidez ME
        cal1 = worksheet.cell(25, 6).value
        cal2 = worksheet.cell(63, 6).value
        value = cal1 / cal2 #
        if (value < 0.2):
            liq_critico_me += 1
        elif (value >= 0.2 ):
            liq_bajo_me += 1
        
        valores[18].append(round(value,2)) #Se redondea a 2 decimales hasta nuevo aviso

R_file = 0            
# Arreglos de rangos de Liquidez MN y ME
liquidez_rangos = ['Menor a 8%', 'Entre 8% y 20%', 'Mayor a 20%','Menor o igual a 20%' ,'Mayor a 20%']
liquidez_mn = [liq_critico_mn, liq_bajo_mn, liq_normal_mn]
liquidez_me = [liq_critico_me, liq_bajo_me, liq_normal_me]
#Arreglos de Obligaciones a CP
oblig_rango = ['Menor a 10%', 'Entre 10% y 20%', 'Entre 20% y 30%', 'Entre 30% y 40%', 'Entre 40% y 50%', 'Mayor a 50%']
obligaciones_cp = [oblig1, oblig2, oblig3, oblig4, oblig5, oblig6]
#Arreglos de % Top 10 Depositantes
depos_rango = ['Menor a 50%', 'Entre 50% y 60%', 'Entre 60% y 70%', 'Entre 70% y 80%', 'Entre 80% y 90%', 'Mayor a 90%']
depositantes_pctj = [deposit1, deposit2, deposit3, deposit4, deposit5, deposit6]

#Solicitar carpeta al usuario
print("Seleccione la carpeta de depósito de información...")

#Codigo de seleccion de carpeta
root = Tk()
root.withdraw()
folder_selected_r = filedialog.askdirectory()

#print(R_file)
workbook = xlsxwriter.Workbook(folder_selected_r+"/Monitor de Liquidez.xlsx", {'strings_to_numbers': True})
worksheetResumen = workbook.add_worksheet("Resumen")
worksheet = workbook.add_worksheet("Liquidez")
worksheetCalculos = workbook.add_worksheet("Calculos")
worksheetCalculos.hide()
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
cell_format_titulo_resumen = workbook.add_format({'bold': True})
cell_format_titulo_resumen.set_color('#003366')
cell_format_titulo_resumen.set_font_size(24)


for i in range(len(valores)):
    worksheet.write(columnas_titulo[i+1].upper()+str(2), titulos[i], cell_format)
worksheetResumen.write('A1',"Graficos de Liquidez", cell_format_titulo_resumen)    
worksheet.write('B1',"Información expresada en S/", cell_format_bold)

row = 2
col = 2

valores = array(valores)
valores = transpose(valores)
valores = valores.tolist()

worksheet.add_table('B3:T'+str(3+len(A_files)-1), {'data': valores, 'header_row': 0})

worksheet.set_column(2, 2, 40) #Tamaño de columna nombre coopac
worksheet.set_column(3, 19, 15) #Tamaño de columna general
worksheetCalculos.write_row(0,0, liquidez_rangos)
worksheetCalculos.write_row(1,0, liquidez_mn)
worksheetCalculos.write_row(2,0, liquidez_me)
worksheetCalculos.write_row(4,0, oblig_rango)
worksheetCalculos.write_row(5,0, obligaciones_cp)
worksheetCalculos.write_row(7,0, depos_rango)
worksheetCalculos.write_row(8,0, depositantes_pctj)
#Grafico de Liquidez en MN
chart = workbook.add_chart({'type': 'column'})
chart.set_y_axis({'name': 'Cantidad de COOPAC'})
chart.set_legend({'position': 'none'})
chart.add_series({
    'name':       'Estado de Liquidez en MN',
    'categories': 'Calculos!A1:C1',
    'values': '=Calculos!A2:C2',
    'data_labels': {'value': True},
    'legend_key': {'value': True} ,
    })
worksheetResumen.insert_chart('C3', chart)

#Grafico de Liquidez en ME
chart = workbook.add_chart({'type': 'column'})
chart.set_y_axis({'name': 'Cantidad de COOPAC'})
chart.set_legend({'position': 'none'})
chart.add_series({
    'name':       'Estado de Liquidez en ME',
    'categories': 'Calculos!D1:E1',
    'values': '=Calculos!A3:B3',
    'data_labels': {'value': True},
    'legend_key':  {'value': True},
    })
worksheetResumen.insert_chart('L3', chart)
#Grafico de Obligaciones a CP
chart = workbook.add_chart({'type': 'column'})
chart.set_y_axis({'name': 'Cantidad de COOPAC'})
chart.set_legend({'position': 'none'})
chart.add_series({
    'name':       'Concentración de MN de Obligaciones a CP',
    'categories': 'Calculos!A5:F5',
    'values': '=Calculos!A6:F6',
    'data_labels': {'value': True},
    })
worksheetResumen.insert_chart('C18', chart)

#Grafico de Depositantes %
chart = workbook.add_chart({'type': 'column'})
chart.set_y_axis({'name': 'Cantidad de COOPAC'})
chart.set_legend({'position': 'none'})
chart.add_series({
    'name':       'Concentración de los 10 Principales Socios con respecto al Total de Depósitos de Socios',
    'categories': 'Calculos!A5:F5',
    'values': '=Calculos!A6:F6',
    'data_labels': {'value': True},
    #'fill':   {'color': 'red'},
    })
worksheetResumen.insert_chart('L18', chart)

workbook.close()
#Matriz de resultados de análisis
