import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.text import MIMEText
from email import encoders

class ReportEmailSender:
    def __init__(self, remitente, password, asunto, cuerpo):
        self.remitente = remitente
        self.password = password
        self.asunto = asunto
        self.cuerpo = cuerpo

    def send_mail(self, destinatario, archivo):
        mensaje = MIMEMultipart()
        mensaje["From"] = self.remitente
        mensaje["To"] = destinatario
        mensaje["Subject"] = self.asunto
        mensaje.attach(MIMEText(self.cuerpo, "plain"))

        # Adjuntar archivo
        if os.path.exists(archivo):
            with open(archivo, "rb") as adj:
                if archivo.lower().endswith((".png", ".jpg", ".jpeg", ".gif")):
                    imagen = MIMEImage(adj.read(), name=os.path.basename(archivo))
                    mensaje.attach(imagen)
                else:
                    parte = MIMEBase("application", "octet-stream")
                    parte.set_payload(adj.read())
                    encoders.encode_base64(parte)
                    parte.add_header("Content-Disposition", f"attachment; filename={os.path.basename(archivo)}")
                    mensaje.attach(parte)
            print(f"📎 Archivo adjuntado: {archivo}")
        else:
            print(f"⚠ Archivo no encontrado: {archivo}")
            return False

        # Enviar correo
        try:
            servidor = smtplib.SMTP("smtp.gmail.com", 587)
            servidor.starttls()
            servidor.login(self.remitente, self.password)
            servidor.send_message(mensaje)
            servidor.quit()
            print(f"✅ Correo enviado a {destinatario}")
            return True
        except Exception as e:
            print(f"❌ Error al enviar el correo a {destinatario}: {e}")
            return False
