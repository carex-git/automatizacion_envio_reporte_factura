import os
from emailSender import ReportEmailSender

# Carpeta donde están los PDFs
base_dir = os.path.dirname(os.path.abspath(__file__))
proyecto_dir = os.path.dirname(base_dir)
directorio_tel = os.path.join(proyecto_dir, "output", "prueba")

email_archivos = []

# Leer archivos y extraer teléfonos
for archivo in os.listdir(directorio_tel):
    if archivo.lower().endswith(".pdf") and "!" in archivo:
        ruta_completa = os.path.join(directorio_tel, archivo)
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
    for n in email_archivos:
        print("enviado: a", n[0] , "el archivo: ", n[1])
        print()
        email_sender.send_mail(n[0], n[1])
else:
    print("⚠️ No se encontraron archivos PDF válidos para enviar.")
