#Manejo de fechas y operaciones con fecha
import openpyxl
import datetime
from datetime import date

clase= openpyxl.load_workbook('Deimer.xlsx')

vamos = clase.active

#Vamos a cambiar el nombre de una celda
vamos['C1'] = "Día de reporte 1"

#Vamos a seleccionar la columna de fechas que es la vamos a utilizar
celdas=vamos['C2':'C6']

formato1 = "%a %b %d %H:%M:%S de %Y"

#for fechas in celdas:
  #fechas=datetime.date(fechas)

listafecha=[]

for fecha in celdas:
  formato=[celda.value for celda in fecha]
  listafecha.append(formato)

print(formato)    
#print(listafecha)
#C2=date(listafecha[0])
#C3=date(listafecha[1])
#C4=date(listafecha[2])
#C5=date(listafecha[3])
#C6=date(listafecha[4])
c2= formato[0]


vamos['C2']="{:%a %b %d %H:%M:%S de %Y}".format(c2)
#vamos['C3']="{:%a %b %d %H:%M:%S de %Y}".format(c3)
#vamos['C4']="{:%a %b %d %H:%M:%S de %Y}".format(vamos['C4'])
#vamos['C5']="{:%a %b %d %H:%M:%S de %Y}".format(vamos['C5'])
#vamos['C6']="{:%a %b %d %H:%M:%S de %Y}".format(vamos['C6'])




clase.save("Respuestas1.xlsx")



formato1 = "%a %b %d %H:%M:%S %Y"
hoy = date.today()
cadena1 = hoy.strftime(formato1) 
print("Formato1:", cadena1)

ahora = date.today()  # Obtiene fecha y hora actual
print("Fecha y Hora:", ahora)  # Muestra fecha y hora
print("Día:",ahora.day)  # Muestra día
print("Mes:",ahora.month)  # Muestra mes
print("Año:",ahora.year)  # Muestra año

#ejemplo 2
#En este segundo ejemplo voy a crear una nueva colomuna en la cuál voy a introducir fechas en diferentes formatos

import openpyxl
import datetime
from datetime import date

clase= openpyxl.load_workbook('Deimer.xlsx')

vamos = clase.active

#Vamos a seleccionar la columna de fechas que es la vamos a utilizar

fecha= datetime.date(2021,3, 6,)
celdas=vamos['D2':'D6']

dia=fecha.day
mes=fecha.month
año=fecha.year

print("{:%A, %d de %b de %Y}".format(fecha))

vamos['D1'] = "Día de reporte 2"
vamos['D2']="{:%A, %d de %b de %Y}".format(fecha)
vamos['D3']="{:%A, %d de %b de %Y}".format(fecha)
vamos['D4']="{:%A, %d de %b de %Y}".format(fecha)
vamos['D5']="{:%A, %d de %b de %Y}".format(fecha)
vamos['D6']="{:%A, %d de %b de %Y}".format(fecha)


formato1 = "%a %b %d %H:%M:%S %Y"

clase.save("Respuestas2.xlsx")