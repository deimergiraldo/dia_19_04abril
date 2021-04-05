#Manejo de fechas y operaciones con fecha
import openpyxl
import datetime
from datetime import date

clase= openpyxl.load_workbook('Deimer.xlsx')

vamos = clase.active

#Vamos a cambiar el nombre de una celda
vamos['C1'] = "Día de entrega de calificaciones"

#Vamos a seleccionar la columna de fechas que es la vamos a utilizar
celdas=vamos['C2':'C6']

formato1 = "%a %b %d %H:%M:%S %Y"

for fila in celdas:
  for celda in fila:
    print(celda.value)

celdas=celda  


clase.save("Nuevodocumento.xlsx")



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

