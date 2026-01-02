import pandas as pd
from sqlalchemy import create_engine
from datetime import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from email.utils import formataddr
import os

### Variables de Conexion ###
server = 'Servidor SQL'
database = 'BASE DE DATOS'

### Cadena de Conexion ###
connection_string = f"mssql+pyodbc://{server}/{database}?driver=ODBC+Driver+17+for+SQL+Server&trusted_connection=yes"
engine = create_engine(connection_string)


query = """
SELECT * 
FROM
(
select top 1 'Analisis_Creditos' AS TABLA,COUNT(*) AS CANTIDAD, CAST(FECHA AS DATE) AS [FECHA DE CARGA]  from openquery(BSC,'select * from Analisis_Creditos') GROUP BY FECHA
UNION 
select top 1 'TC' AS TABLA,COUNT(*) AS CANTIDAD, CAST(FECHACARGA AS DATE) AS [FECHA DE CARGA]  from openquery(BSC,'SELECT * FROM TC') GROUP BY FECHACARGA
UNION
select top 1 'PR' AS TABLA,COUNT(*) AS CANTIDAD, CAST(FECHACARGA AS DATE)  AS [FECHA DE CARGA]  from openquery(BSC,'select * from PR') GROUP BY FECHACARGA
UNION
select top 1 'CLIENTES' AS TABLA,COUNT(*) AS CANTIDAD, CAST(FECHACARGA AS DATE) AS [FECHA DE CARGA]  from openquery(BSC,'select * from Clientes') GROUP BY FECHACARGA
) A
"""

df = pd.read_sql_query(query, engine)

#Validaciones
#df['ACTUALIZACION'] = df['FECHA_CARGA'].dt.date == fecha_actual

# Convertir la columna de fechas a tipo datetime
df['FECHA DE CARGA'] = pd.to_datetime(df['FECHA DE CARGA'])
df['CANTIDAD'] = df['CANTIDAD'].apply(lambda x: f"{x:,}")

# Fecha Actual
fecha_actual = datetime.now().date()

#Cuerpo del correo con formato HTML
df_html = f"""
<!DOCTYPE html>
<html lang="es">
  <body bgcolor="#1e1e1e" style="margin:0;padding:0;background:#1e1e1e;">
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" bgcolor="#1e1e1e" style="background:#1e1e1e;">
      <tr>
        <td align="center" style="padding:24px 12px;">
          <table role="presentation" width="600" cellpadding="0" cellspacing="0" border="0" style="background:#1e1e1e;border-radius:12px;">

            <!-- Encabezado -->
            <tr>
              <td bgcolor="#3f5995" style="background:#3f5995;padding:20px;border-radius:12px 12px 0 0;text-align:center;">
                <div style="color:#FFFFFF;font-family:Calibri,Arial,sans-serif;font-size:20px;font-weight:800;">
                  📊 Reporte de Carga de Tablas
                </div>
                <div style="color:#E2E8F0;font-family:Calibri,Arial,sans-serif;font-size:12px;margin-top:6px;">
                  Fecha de actualización al {fecha_actual}
                </div>
              </td>
            </tr>

            <!-- Contenido -->
            <tr>
              <td style="background:#2b2b2b;border:1px solid #444;border-top:0;border-radius:0 0 12px 12px;padding:18px;text-align:center;">
                <div style="color:#E2E8F0;font-family:Calibri,Arial,sans-serif;font-size:14px;margin-bottom:14px;">
                  Tablas de origen (MySql) procesadas.
                </div>

                <!-- Tabla -->
                <table width="100%" cellpadding="10" cellspacing="2" border="0" style="border-collapse:separate;border-spacing:0;width:100%;border:1px solid #444;border-radius:8px;overflow:hidden;text-align:center;">
                  <tbody>
                    {df.to_html(index=False, border=0, justify='center', classes='email-table')}
                  </tbody>
                </table>

                <!-- Pie -->
                <div style="color:#888;font-family:Calibri,Arial,sans-serif;font-size:12px;margin-top:16px;border-top:1px solid #444;padding-top:10px;">
                  Este correo fue generado por el equipo de <strong style="color:#ccc;">Estructura de datos Riesgos</strong>.
                </div>
              </td>
            </tr>

          </table>
        </td>
      </tr>
    </table>
  </body>
</html>
"""

df_html = df_html.replace('<td>', '<td style="text-align:center">')
df_html = df_html.replace('<th>', '<th style="text-align:center">')

#######################################################################ENVIO DE CORREO#############################################################################
 
# Configura los detalles del correo electrónico
nombre_remitente = 'Fecha de tablas origen'
direccion_remitente = 'walmanzar@correo.com.do'
destinatarios = ['nmendoza@correo.com.do','nmejia@correo.com.do','n','n','n']
#destinatarios = ['wialmanzar@bsc.com.do']
asunto = 'Informacion de fechas de las tablas Origen'
cuerpo = df_html #.format(df.to_html(index=False))
 
# Configura el servidor SMTP
servidor_smtp = 'server'  # Nombre del servidor SMTP
puerto = 25  # Puerto para el servidor SMTP
 
# Crear un objeto MIMEText con el cuerpo del correo electrónico
mensaje = MIMEMultipart()
mensaje['From'] = formataddr((nombre_remitente, direccion_remitente))
mensaje['To'] = ', '.join(destinatarios)
mensaje['Subject'] = asunto
mensaje.attach(MIMEText(cuerpo, 'html'))

# Iniciar una conexión SMTP
with smtplib.SMTP(servidor_smtp, puerto) as servidor:
    servidor.send_message(mensaje)  # Enviar el correo electrónico
 
# Imprimir mensaje de confirmación
print(f"El correo electrónico ha sido enviado.")