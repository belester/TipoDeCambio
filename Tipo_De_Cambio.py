import pyodbc
from datetime import datetime, timedelta
from conChrome import conCromeDriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import calendar
from udis import UDIS_P
from sqlalchemy import create_engine
from dolar import dolar
from ConexionBD.Con82 import conServidor

conCromeDriver()

cursor = conServidor()

cursor.execute('SELECT MAX(Fecha) FROM C0001 WHERE id_moneda=2;')
ult_fechaBD = cursor.fetchone()[0]

#Prueba de impresión, recordar comentar
#print(f'Ultima Fecha obtenida por la BD: {ult_fechaBD} \n')

#Cerrar la consulta
cursor.close()

# Obtener la fecha actual
fecha_actual = datetime.now().date()

# Función para obtener el número de días de un mes
def obtener_dias_mes(mes, anio):
    if mes == 2:  # Febrero
        if calendar.isleap(anio):  # Verificar si el año es bisiesto
            return 29
        else:
            return 28
    else:
        return calendar.monthrange(anio, mes)[1]

# Convertir la fecha a un objeto datetime, si no es None
if ult_fechaBD:
        if isinstance(ult_fechaBD, str):
            ult_fechaBD = datetime.strptime(ult_fechaBD, "%Y-%m-%d").date()
        elif isinstance(ult_fechaBD, datetime.datetime):
            ult_fechaBD = ult_fechaBD.date()

#Recordar comentar la siguiente linea de código
#print(f'Fecha después de la conversión: {ult_fechaBD}')

# Verificar si se encontró una fecha en la base de datos
if ult_fechaBD:
    # Obtener el siguiente día de la última fecha en la base de datos
    fecha_inicio = ult_fechaBD + timedelta(days=1)
else:
    # No se encontró una fecha en la base de datos, comenzar desde el primer día del mes actual
    fecha_inicio = fecha_actual.replace(day=1).date()

#Recuerda comentar la siguiente linea de código
#print(f'Fecha de Inicio: {fecha_inicio}')


print(f'Fecha de Inicio-----: {fecha_inicio}')
print(f'Fecha de Final-----: {fecha_actual}')

#UDIS_P()

dolar(fecha_Inicio=fecha_inicio, fecha_Actual=fecha_actual)

# Cerrar la conexión a SQL Server
conn.close()