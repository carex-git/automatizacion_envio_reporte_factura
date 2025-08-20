import os
import pandas as pd
import json
import time
from WhatsAppSender import WhatsAppSafeSender
from Reporte_Proveedor import ReporteProveedor

# Cargar configuración
with open("config.json", "r", encoding="utf-8") as f:
    config = json.load(f)

def procesar_contactos():
    base_dir = config['base_dir']
    print(base_dir)

    # Carpeta donde se guardará el Excel de verificación
    output_dir = os.path.join(base_dir, "output", "tel_verif")
    os.makedirs(output_dir, exist_ok=True)

    # Carpeta donde buscará los PDFs
    directorio_tel = os.path.join(base_dir,"dist", "output", "prueba")

    contactos_archivos = []
    contactos = []

    for archivo in os.listdir(directorio_tel):
        if archivo.lower().endswith(".pdf") and "!" in archivo:
            ruta_completa = os.path.join(directorio_tel, archivo)
            base = archivo.removesuffix(".pdf")
            partes = base.split("!")
            if len(partes) == 3:
                nombre, cc, telefono = partes
                contactos.append({
                    "C.c": cc,
                    "nombre": nombre,
                    "celular": telefono,
                    "verificado": None
                })
                numero_formateado = "+57" + telefono
                contactos_archivos.append({
                    "numero": numero_formateado,
                    "archivo": ruta_completa,
                    "nombre": nombre
                })

    # Guardar en Excel
    output_path = os.path.join(output_dir, "tel_verificacion.xlsx")
    pd.DataFrame(contactos).to_excel(output_path, index=False)

    return contactos_archivos

if __name__ == "__main__":
    contactos_archivos = procesar_contactos()

    if contactos_archivos:
        # Se crean las rutas para la carpeta de enviados
        directorio_tel = os.path.join(config['base_dir'], "dist", "output", "prueba")
        enviados_dir = os.path.join(directorio_tel, "enviados")
        
        # Se pasa la ruta de la carpeta de enviados al constructor
        sender = WhatsAppSafeSender(contacts=contactos_archivos,
                                send_buttons=config['send_buttons'],
                                mensaje=config["menssage_whatsApp"],
                                profile_path=config['profile_path'],
                                attach_buttons=config["attach_buttons"],
                                document_buttons=config["document_buttons"],
                                no_contact_buttons=config["no_contact_buttons"],
                                enviados_dir=enviados_dir)
        
        sender.main()
    else:
        print("⚠️ No se encontraron archivos PDF válidos para enviar.")