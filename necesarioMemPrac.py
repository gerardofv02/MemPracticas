import requests
import smtplib
import pandas as pd
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import time

def descargar_archivo(url, nombre_archivo):
    try:
        # Realizar la solicitud HTTP GET para descargar el archivo
        respuesta = requests.get(url)
        
        # Verificar si la solicitud fue exitosa (código de estado 200)
        if respuesta.status_code == 200:
            # Abrir el archivo en modo binario y escribir los datos descargados
            with open(nombre_archivo, 'wb') as archivo:
                archivo.write(respuesta.content)
            print(f'Descarga exitosa: {nombre_archivo}')
        else:
            print(f'Error al descargar el archivo. Código de estado: {respuesta.status_code}')
    except Exception as e:
        print(f'Error: {e}')

def enviar_correo(destinatario, asunto, cuerpo):
    # Configuración del servidor SMTP
    smtp_host = 'server'
    smtp_port = 0 #port
    smtp_usuario = 'user'
    smtp_contrasena = 'pass'

    # Configuración del mensaje
    mensaje = MIMEMultipart()
    mensaje['From'] = smtp_usuario
    mensaje['To'] = destinatario
    mensaje['Subject'] = asunto
    mensaje.attach(MIMEText(cuerpo, 'plain'))

    # Iniciar conexión con el servidor SMTP
    servidor_smtp = smtplib.SMTP(smtp_host, smtp_port)
    servidor_smtp.starttls()
    servidor_smtp.login(smtp_usuario, smtp_contrasena)

    # Enviar el correo
    servidor_smtp.sendmail(smtp_usuario, destinatario, mensaje.as_string())

    # Cerrar la conexión con el servidor SMTP
    servidor_smtp.quit()


def archivo_excel(archivo):
    cont = 0
    df = pd.read_excel(archivo,"Hoja1")
    i = 0
    print(df)
    for i in range(len(df)):
        Nota = df.iloc[i]["Nota"]
        destinatario = df.iloc[i]["Correo"]

        if Nota < 5:
            df.loc[i, "Status"] = "Suspenso"
            try:
                enviar_correo(destinatario,asunto="Status examen", cuerpo="Ha suspendido")
            except: 
                continue
        else:
            df.loc[i, "Status"] = "Aprobado"
            try:
                enviar_correo(destinatario,asunto="Status examen", cuerpo="Ha aprobado")
            except:
                continue


    
    df.to_excel(archivo)


inicio = time.time()
url = "url"
nombre_archivo = "NecesarioMemPrac.xlsx"
descargar_archivo(url,nombre_archivo)
archivo_excel(archivo="./NecesarioMemPrac.xlsx")
fin = time.time()
print(fin-inicio) 