#Manejo de fechas y operaciones con fecha
import openpyxl
import datetime
from datetime import date

clase= openpyxl.load_workbook('Deimer.xlsx')

vamos = clase.active

#Aquí elegí cada una de las fechas que aparecían en el documento de excel
fecha= datetime.date(2021,3, 10)
fecha2= datetime.date(2021,3, 11)
fecha3= datetime.date(2021,3, 12)
fecha4= datetime.date(2021,3, 13)
fecha5= datetime.date(2021,3, 14)

#Aquí estoy a cambiando el nombre de una celda e insertando la fecha en un formato en especifico

vamos['D1'] = "Día de reporte 1"
vamos['D2']="{:%a %b %d %H:%M:%S de %Y}".format(fecha)
vamos['D3']="{:%a %b %d %H:%M:%S de %Y}".format(fecha2)
vamos['D4']="{:%a %b %d %H:%M:%S de %Y}".format(fecha3)
vamos['D5']="{:%a %b %d %H:%M:%S de %Y}".format(fecha4)
vamos['D6']="{:%a %b %d %H:%M:%S de %Y}".format(fecha5)

clase.save("Respuestas1.xlsx")


#ejemplo 2
#En este segundo ejemplo voy a crear una nueva colomuna en la cuál voy a introducir fechas en diferentes formatos

import openpyxl
import datetime
from datetime import date

clase= openpyxl.load_workbook('Deimer.xlsx')

vamos = clase.active

#Vamos a introducir la fecha que vamos a utlizar

fecha= datetime.date(2021,3, 6,)

print("{:%A, %d de %b de %Y}".format(fecha))

#Ahora vamos a introducir esas fechas a cada casilla con un formato en específico

vamos['D1'] = "Día de reporte 2"
vamos['D2']="{:%A, %d de %b de %Y}".format(fecha)
vamos['D3']="{:%A, %d de %b de %Y}".format(fecha)
vamos['D4']="{:%A, %d de %b de %Y}".format(fecha)
vamos['D5']="{:%A, %d de %b de %Y}".format(fecha)
vamos['D6']="{:%A, %d de %b de %Y}".format(fecha)


clase.save("Respuestas2.xlsx")



#Estás son algunas de las operaciones que realice para entender mejor el manejo de fechas y horas y que no incluí en los documentos

formato1 = "%a %b %d %H:%M:%S %Y"
hoy = date.today()
cadena1 = hoy.strftime(formato1) 
print("Formato1:", cadena1)

ahora = date.today()  # Obtiene fecha y hora actual
print("Fecha y Hora:", ahora)  # Muestra fecha y hora
print("Día:",ahora.day)  # Muestra día
print("Mes:",ahora.month)  # Muestra mes
print("Año:",ahora.year)  # Muestra año