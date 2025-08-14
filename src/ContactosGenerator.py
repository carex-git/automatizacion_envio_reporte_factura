import os
from WhatsAppSender import WhatsAppSender

# Carpeta donde están los PDFs
base_dir = os.path.dirname(os.path.abspath(__file__))
proyecto_dir = os.path.dirname(base_dir)
directorio_tel = os.path.join(proyecto_dir, "output", "prueba")

contactos_archivos = []

# Leer archivos y extraer teléfonos
for archivo in os.listdir(directorio_tel):
    if archivo.lower().endswith(".pdf") and "!" in archivo:
        ruta_completa = os.path.join(directorio_tel, archivo)
        base = archivo.removesuffix(".pdf")
        partes = base.split("!")
        if len(partes) == 3:
            _, _, telefono = partes
            numero_formateado = "+57" + telefono
            contactos_archivos.append({
                "numero": numero_formateado,
                "archivo": ruta_completa
            })
            print(f"📂 Preparado: {numero_formateado} | {ruta_completa}")

print(f"\nLista final de contactos: {contactos_archivos}")

# Pasar contactos a la clase y ejecutar
if contactos_archivos:
    sender = WhatsAppSender(contacts=contactos_archivos)
    sender.main()
else:
    print("⚠️ No se encontraron archivos PDF válidos para enviar.")
