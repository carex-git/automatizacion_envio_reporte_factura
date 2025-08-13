import fitz  # PyMuPDF
import pywhatkit
import os

# Ruta del PDF
pdf_path = r"C:\Users\aprsistemas\Downloads\IMPRESION 2 JULIO.pdf"

# Convertir PDF a imagen (primera página)
pdf = fitz.open(pdf_path)
pagina = pdf[0]
pix = pagina.get_pixmap(dpi=200)

# Guardar imagen temporal
imagen_path = pdf_path.replace(".pdf", ".jpg")
pix.save(imagen_path)

# Número de WhatsApp
numero = "+573178935198"

# Enviar imagen por WhatsApp
pywhatkit.sendwhats_image(
    receiver=numero,
    img_path=imagen_path,
    caption="Hola, te envío tu documento 📄",
    wait_time=30,
    tab_close=True
)

print("Imagen enviada ✅")
