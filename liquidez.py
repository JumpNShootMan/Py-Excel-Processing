import os
import xlrd
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

#Solicitar carpeta al usuario
print("Seleccione la carpeta a analizar...")

#Codigo de seleccion de carpeta
root = Tk()
root.withdraw()
folder_selected = filedialog.askdirectory()

#Validacion de ruta seleccionada
#print(folder_selected)
Path = folder_selected

#Listado de archivos del directorio y asignación a un vector
A_files = []  
#Arreglos de liquidez con rangos de >=8%, <8% y >=20%,  <20%
liq_critico_mn = 0
liq_bajo_mn = 0               
liq_normal_mn = 0     
liq_critico_me = 0
liq_bajo_me = 0               
liq_normal_me = 0    
#print("Listado de archivos en ruta:")                                            
for dirName, subdirList, fileList in os.walk(Path):                        
    for filename in fileList:   
        if ".xlsx" in filename.lower() or ".xlsm" in filename.lower(): 
            if not filename.startswith('~$'):
                A_files.append(os.path.join(dirName,filename)) 

#print(A_files)

#Matriz de valores por llenar | Especificar cuántas filas deben haber en la salida global
valores= [[] for i in range(19)]

#Lectura de información en base a excel llamado desde el vector
if(len(A_files) > 1):
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
        value = cal1/cal2 #Fondos disponibles -> Tabla 1 total en MN / total
        valores[8].append(round(value,2)) #Se redondea a 2 decimales hasta nuevo aviso
        #Valor Obligaciones CP (Cálculo)
        cal1 = worksheet.cell(63, 5).value
        cal2 = worksheet.cell(63, 7).value
        value = cal1/cal2 #
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
        #Valor Fondos Disponibles / Depósitos Socios (Cálculo)
        cal1 = worksheet.cell(75, 7).value
        cal2 = worksheet.cell(59, 7).value
        cal3 = worksheet.cell(60,7).value
        cal4 = worksheet.cell(67,7).value
        cal5 = worksheet.cell(68,7).value
        value = cal1 / (cal2+cal3+cal4+cal5) #
        valores[16].append(round(value,2)) #Se redondea a 2 decimales hasta nuevo aviso
        #Valor Liquidez MN
        cal1 = worksheet.cell(25, 5).value
        cal2 = worksheet.cell(63, 5).value
        value = cal1 / cal2 #
        if (value <= 0.08):
            liq_critico_mn += 1
        elif (value > 0.08 and value <= 0.2):
            liq_bajo_mn += 1
        elif (value > 0.2):
            liq_normal_mn += 1
            
        valores[17].append(round(value,2)) #Se redondea a 2 decimales hasta nuevo aviso
        #Valor Liquidez ME
        cal1 = worksheet.cell(25, 6).value
        cal2 = worksheet.cell(63, 6).value
        value = cal1 / cal2 #
        if (value <= 0.08):
            liq_critico_me += 1
        elif (value > 0.08 and value <= 0.2):
            liq_bajo_me += 1
        elif (value > 0.2):
            liq_normal_me += 1
            
        valores[18].append(round(value,2)) #Se redondea a 2 decimales hasta nuevo aviso
print("Seleccione la carpeta de depósito de información...")
folder_selected_r = filedialog.askdirectory()
# print("Listado de archivos en ruta:")
R_file = 0            
# Arreglos de rangos de Liquidez MN y ME
liquidez_mn = [liq_critico_mn, liq_bajo_mn, liq_normal_mn]
liquidez_me = [liq_critico_me, liq_bajo_me, liq_normal_me]
for dirName, subdirList, fileList in os.walk(folder_selected_r):
    for filename in fileList:
        #print(filename)                                                    
        if ".xlsx" in filename.lower() or ".xlsm" in filename.lower():
            if not filename.startswith('~$'):
                R_file = os.path.join(dirName,filename)

#print(R_file)
workbook = xlsxwriter.Workbook(R_file, {'strings_to_numbers': True})
worksheet = workbook.add_worksheet("Liquidez")
worksheetResumen = workbook.add_worksheet("Resumen")
#Letras hasta la cantidad de columnas necesarias
columnas_titulo = []
for c in ascii_lowercase:
    columnas_titulo.append(c)

#-----------------------------------------------
#print("Seleccione folder de títulos...")
#folder_selected_t = filedialog.askdirectory()
#with open(titul, encoding='utf-8') as f:
#    d = load(f)
#    titulos = d["titulos"]
#print(titulos)
#-----------------------------------------------

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

worksheet.add_table('B3:T'+str(3+len(A_files)-1), {'data': valores, 'header_row': 0})

worksheet.set_column(2, 2, 40) #Tamaño de columna nombre coopac
worksheet.set_column(3, 19, 15) #Tamaño de columna general
print(liquidez_me)
#Grafico de Liquidez en MN
chart = workbook.add_chart({'type': 'column'})
chart.add_series({'values': '=Liquidez!S3:S'+str(3+len(A_files)-1)})
worksheetResumen.insert_chart('C1', chart)



workbook.close()
#Matriz de resultados de análisis
k=input("Los resultados de encuentran en la carpeta. Presionar intro para cerrar")