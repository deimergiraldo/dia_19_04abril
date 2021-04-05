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

fecha= datetime.date(2021,3, 10)
fecha2= datetime.date(2021,3, 11)
fecha3= datetime.date(2021,3, 12)
fecha4= datetime.date(2021,3, 13)
fecha5= datetime.date(2021,3, 14)


vamos['C2']="{:%a %b %d %H:%M:%S de %Y}".format(fecha)
vamos['C3']="{:%a %b %d %H:%M:%S de %Y}".format(fecha2)
vamos['C4']="{:%a %b %d %H:%M:%S de %Y}".format(fecha3)
vamos['C5']="{:%a %b %d %H:%M:%S de %Y}".format(fecha4)
vamos['C6']="{:%a %b %d %H:%M:%S de %Y}".format(fecha5)




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