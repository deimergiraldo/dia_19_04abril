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

formato1 = "%a %b %d %H:%M:%S %Y"

for fila in celdas:
  for celda in fila:
    print(celda.value)

celdas=celda  


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