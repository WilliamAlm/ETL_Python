import pymssql
import pandas as pd
from sqlalchemy import create_engine,text
import os
from pathlib import Path
import time

#Conexion con la DB
server = 'localhost'
database = 'Base de datos'

### Conexion para procesos ###
conn = pymssql.connect(server = server, database = database)
cursor = conn.cursor()

#Cadena de conexion con drivers
cs = f"mssql+pyodbc://{server}/{database}?driver=ODBC+Driver+17+for+SQL+Server&trusted_connection=yes"
engine = create_engine(cs)

#Inicio
inicio = time.time()
print(f'⏱️ Iniciado a las: {inicio}')

#Ubicacion de archivo
ruta = r'\\ruta compartida del documento excel\excel.xlsx'
df = pd.read_excel(ruta)

#limpio la tabla en sql server
truncate = text(""" 
TRUNCATE TABLE [base_de_datos].[dbo].[tabla]
""")

#Ejecutar el query anterior
with engine.begin() as connection:
    connection.execute(truncate)

#Insertando la data del excel a sql sever
print('#####--------> 1. Ejecutando la Carga de la data <--------#####')
df.to_sql('tabla', con=engine, if_exists='append', index=False)
print('Carga Completada.....')

#Ejecutar el StoredProceduce que esta en SQL Server
print('#####--------> 2. Ejecutando el proceso de SQL Server <--------#####')
cursor.callproc('SP_INSERT_DATA')
conn.commit()

#Fin del proceso, marca el tiempo final de la ejecucion
fin = time.time()
duracion = (fin - inicio)/60

print(f'Proceso Completado en: {duracion}')