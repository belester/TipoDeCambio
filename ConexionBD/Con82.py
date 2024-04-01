import pyodbc

def conServidor():

    server = '15.10.155.82\DWHDES001'
    database = 'Rentabilidad'

    #Conexion a la bd
    try:
        # Conexión a SQL Server
        conn = pyodbc.connect(
            'Driver={SQL Server};Server=' + server + ';Database=' + database + ';Trusted_Connection=yes;')
        print('Conexión Exitosa')

    except Exception as e:
        print(f'Error al intentar conectarse debido a: {e}')

    # Consulta a la base de datos
    cursor = conn.cursor()

    #Devolver cursor

    return cursor