#smtplib es la libreria que te conecta con el servidor SMTP, que es el que se encarga de enviar correos
import smtplib

#Estas dos te ayudan a construir el cuerpo del correo MIME significa "Multipurpose Internet Mail Extensions", que sirve para que el correo tenga texto, HTML, archivos adj, etc
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

#Datos del correo
remitente = "jperez@novaseguridad.cl"
destinatario = "destino@gmail.com"
asunto = "Correo Automatico"
mensaje = "Prueba de correo automatizado"

#Aqui construiremos el mensaje
#MIMEMultipart() crea un contenedor para el correo
msg = MIMEMultipart

#Estos tres son los metadatos del correo
msg['From'] = remitente
msg['To'] = destinatario
msg['Subject'] = asunto

#attach(MIMEText(...)), aqui es donde metemos el cuerpo del mensaje, 'plain' significa que el texto es plano y si quisiera cambiarse a otro tipo, ejemplo HTML, se cambia de 'plain' a 'html'
msg.attach(MIMEText(mensaje, 'plain'))

#Conexion con el servidor SMTP
#'smtp.gmail.com', es el servidor de gmail. El puerto 587 es el que se usa para conexiones seguras con TLS
servidor = smtplib.SMTP('smtp.office365.com', 587)

#starttls Inicia la conexion segura
servidor.starttls()

#login(...) es donde autenticas. 
servidor.login(remitente, 'TU_CLAVE_DE_APLICACION')

#enviar y cerrar
#sendmail(...), es aqui donde se manda el coreo, msg.as_string() convierte todo el mensaje en texto plano para que el servidor pueda entenderlo
servidor.sendmail(remitente, destinatario, msg.as_string())

#quit() cierra la conexion con el servidor
servidor.quit()
print("Coreo enviado con exito desde Outlook =)")

#esto es por si algo falla, el programa no se caiga, si no que te diga que fue lo que fallo
try:
  #codigo
except Exception as e:
  print(f"Error al enviar el correo: {e}")

