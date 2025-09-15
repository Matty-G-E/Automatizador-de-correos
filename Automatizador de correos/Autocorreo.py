
#pandas: Es una biblioteca de Py que proporciona estructuras de datos y herramientas de analisis de datos faciles de usar y flexibles para el lenguaje de programacion
import pandas as pd


#smtplib: es una biblioteca estandar de Python que define una interfaz para enviar correos electronicos utilizando el Protocolo Simple de Transferencia de Correo (SMTP)
import smtplib

#tkinter: es una biblioteca estandar de Python para crear interfaces graficas de usuario (GUI)
from tkinter import Tk

#Askopenfilename: es una funcion del modulo filedialog que abre un cuadro de dialogo de "abrir archivo", para que el usuario pueda seleccionar un archivo
from tkinter.filedialog import askopenfilename

#MIMEMultipart: es una clase del modulo email.mime.multipart que permite crear mensajes de correo electronico con multiples partes (texto, imagenes, archivos adjuntos, etc.)
from email.mime.multipart import MIMEMultipart

#MIMEText: es una clase del modulo email.mime.text que permite crear partes de texto para los mensajes de correo electronico
from email.mime.text import MIMEText

#Ocultar ventana principal de Tkinter
Tk().withdraw()

#Abrir el explorador de archivos
archivo_excel = askopenfilename(title="Selecciona el archivo Excel", filetypes=[("Excel files", "*.xlsx")])


#Leer el archivo Excel
df = pd.read_excel(archivo_excel)

#Seleccionar las columnas deseadas para enviar en el correo
columnas_deseadas = [
    'RUT Proveedor.1',
    'RUT Proveedor.2',
    'Razon Social',
    'Folio',
    'Fecha Docto',
    'Fecha Recepcion',
    'Monto Exento',
    'Monto Neto',
    'Monto IVA Recuperable',
    'Monto Total'
    
]

#Filtrar solo esas columnas
tabla = df[columnas_deseadas]

#Aqui convertimos la tabla en formato HTML, con bordes
tabla_html = tabla.to_html(index=False, border=2)

#Datos del correo
remitente = "micorreo@outlook.cl"
clave = "******"  # Reemplaza con la clave real o utiliza un metodo seguro para manejar contraseñas
destinatario = "correodestinatario@.cl"

#Crear el mensaje
msg = MIMEMultipart()
msg['From'] = remitente
msg['To'] = destinatario
msg['Subject'] = "Reporte de FCC"

#Cuerpo del correo

cuerpo_html = f"""
<html>
  <body>
    <p>Estimado ,</p>
    <p> Espero que se encuentre bien,</p>
    <p>Adjunto el reporte de FCC.</p>,
    {tabla_html}
    <p>Saludos cordiales.</p>
  </body>
</html>
"""

msg.attach(MIMEText(cuerpo_html, 'html'))

#enviar el correo

try:
    servidor = smtplib.SMTP('smtp.office365.com', 587)
    servidor.starttls()
    servidor.login(remitente, clave)
    servidor.sendmail(remitente, destinatario, msg.as_string())
    servidor.quit()
    print("Correo enviado exitosamente.")
except Exception as e:
    print(f"Error al enviar el correo: {e}")