import os
import pandas as pd
import json
import time
from emailSender import ReportEmailSender

def mover_archivo_enviado(archivo, enviados_dir, index):
    """
    Mueve un archivo a la carpeta 'enviados' con un índice al principio del nombre.
    """
    if not os.path.exists(enviados_dir):
        os.makedirs(enviados_dir)
    
    nombre_base = os.path.basename(archivo)
    nombre_con_indice = f"{index:03d}_{nombre_base}"
    ruta_destino = os.path.join(enviados_dir, nombre_con_indice)
    
    try:
        os.rename(archivo, ruta_destino)
        print(f"✅ Archivo movido a: {ruta_destino}")
        return True
    except Exception as e:
        print(f"❌ Error al mover el archivo {archivo}: {e}")
        return False

# Cargar configuración
with open("config.json", "r", encoding="utf-8") as f:
    config = json.load(f)

# Carpeta donde están los PDFs
base_dir = os.path.dirname(os.path.abspath(__file__))
directorio_email = os.path.join(base_dir, "dist", "output", "prueba 2")

# Crear la carpeta de enviados
enviados_dir = os.path.join(directorio_email, "enviados")
os.makedirs(enviados_dir, exist_ok=True)

email_archivos = []

# Leer archivos y extraer teléfonos
for archivo in os.listdir(directorio_email):
    if archivo.lower().endswith(".pdf") and "!" in archivo:
        ruta_completa = os.path.join(directorio_email, archivo)
        base = archivo.removesuffix(".pdf")
        partes = base.split("!")
        if len(partes) == 3:
            _, _, email = partes
            email_archivos.append([
                email,
                ruta_completa
            ])
            print(f"📂 Preparado: {email} | {ruta_completa}")

print(f"\nLista final de contactos: {email_archivos}")

# Pasar contactos a la clase y ejecutar
if email_archivos:
    email_sender = ReportEmailSender("cm090457@gmail.com", "xrjw xnuz jtox qrmv", "asunto", "cuerpo")
    exitosos = 0
    fallidos = 0
    conteo_enviados = 1

    for n in email_archivos:
        destinatario = n[0]
        ruta_archivo = n[1]
        
        print(f"\n✉️ Enviando a: {destinatario} el archivo: {os.path.basename(ruta_archivo)}")
        
        # Asumiendo que send_mail retorna True si el envío fue exitoso, False en caso contrario.
        if email_sender.send_mail(destinatario, ruta_archivo):
            print(f"✅ Correo enviado con éxito a {destinatario}.")
            if mover_archivo_enviado(ruta_archivo, enviados_dir, conteo_enviados):
                exitosos += 1
                conteo_enviados += 1
            else:
                fallidos += 1
        else:
            print(f"❌ Fallo al enviar el correo a {destinatario}.")
            fallidos += 1
            
    print(f"\nResumen de envío:")
    print(f"✅ Exitosos: {exitosos}")
    print(f"❌ Fallidos: {fallidos}")
else:
    print("⚠️ No se encontraron archivos PDF válidos para enviar.")