import pyautogui
import cv2
import numpy as np
import time
import os
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
from pyautogui import ImageNotFoundException


class WhatsAppSender:
    def __init__(self,contacts):
        self.CONTACTOS = contacts
        self.MENSAJE = "Te envío el documento que me pediste"

        self.ATTACH_BUTTON_TEMPLATE = r"C:\Users\aprsistemas\Desktop\trabajo\automatizacion_envio_reporte_factura\src\attach_button_template.png"
        self.DOCUMENT_BUTTON_TEMPLATE = r"C:\Users\aprsistemas\Desktop\trabajo\automatizacion_envio_reporte_factura\src\document_button_template.png"
        self.NO_CONTACT_TEMPLATE = r"C:\Users\aprsistemas\Desktop\trabajo\automatizacion_envio_reporte_factura\src\no_contact_template.png"

        self.CAPTION_BOX_XPATH = '//div[@contenteditable="true"][@data-tab="10"]'
        self.SEARCH_BOX_XPATH = '//div[@contenteditable="true"][@data-tab="3"]'

    def click_image(self, template_path, confidence=0.8, timeout=8):
        start_time = time.time()
        while time.time() - start_time < timeout:
            try:
                if not os.path.exists(template_path):
                    print(f"❌ Error: El archivo de plantilla no existe: {template_path}")
                    return False
                location = pyautogui.locateCenterOnScreen(template_path, confidence=confidence, grayscale=True)
                if location:
                    pyautogui.click(location)
                    print(f"✅ Clic en la imagen: {template_path}")
                    return True
            except pyautogui.PyAutoGUIException as e:
                print(f"ℹ️ Intentando buscar imagen: {template_path}. Error: {e}")
            time.sleep(2)
        print(f"❌ No se encontró la imagen: {template_path}")
        return False

    def iniciar_driver(self):
        try:
            service = Service(ChromeDriverManager().install())
            options = Options()
            options.add_argument("--start-maximized")
            options.add_argument("--disable-blink-features=AutomationControlled")
            profile_path = r"C:\Users\aprsistemas\AppData\Local\Google\Chrome\User Data\WhatsAppSession"
            options.add_argument(f"--user-data-dir={profile_path}")
            driver = webdriver.Chrome(service=service, options=options)
            wait = WebDriverWait(driver, 60)
            return driver, wait
        except Exception as e:
            print(f"❌ Error al iniciar el driver: {e}")
            return None, None

    def esperar_whatsapp_cargado(self, wait):
        try:
            wait.until(EC.presence_of_element_located((By.XPATH, self.SEARCH_BOX_XPATH)))
            print("✅ WhatsApp Web cargado")
            return True
        except Exception:
            print("❌ No se pudo cargar WhatsApp Web")
            return False

    def abrir_chat_con_contacto(self, driver, wait, numero):
        chat_abierto = False
        try:
            search_box = wait.until(EC.element_to_be_clickable((By.XPATH, self.SEARCH_BOX_XPATH)))
            search_box.clear()
            search_box.send_keys(numero)
            time.sleep(3)
            if not pyautogui.locateOnScreen(self.NO_CONTACT_TEMPLATE, confidence=0.8, grayscale=True):
                search_box.send_keys(Keys.ENTER)
                wait.until(EC.presence_of_element_located((By.XPATH, self.CAPTION_BOX_XPATH)))
                chat_abierto = True
        except ImageNotFoundException:
            search_box.send_keys(Keys.ENTER)
            wait.until(EC.presence_of_element_located((By.XPATH, self.CAPTION_BOX_XPATH)))
            chat_abierto = True
        except Exception as e:
            print(f"❌ Error búsqueda: {e}")
        if not chat_abierto:
            try:
                driver.get(f"https://web.whatsapp.com/send?phone={numero}")
                wait.until(EC.presence_of_element_located((By.XPATH, self.CAPTION_BOX_XPATH)))
                chat_abierto = True
            except Exception:
                driver.get("https://web.whatsapp.com")
                self.esperar_whatsapp_cargado(wait)
        return chat_abierto

    def enviar_documento_autogui(self, wait, numero, archivo):
        if not self.click_image(self.ATTACH_BUTTON_TEMPLATE):
            return False
        if not self.click_image(self.DOCUMENT_BUTTON_TEMPLATE):
            return False
        time.sleep(3)
        try:
            pyautogui.write(os.path.abspath(archivo))
            pyautogui.press('enter')
            print(f"✅ Archivo seleccionado: {archivo}")
        except Exception as e:
            print(f"❌ Error archivo: {e}")
            return False
        try:
            caption_box = wait.until(EC.presence_of_element_located((By.XPATH, self.CAPTION_BOX_XPATH)))
            caption_box.send_keys(self.MENSAJE)
            caption_box.send_keys(Keys.ENTER)
            print(f"✅ Mensaje enviado a {numero}")
        except Exception as e:
            print(f"❌ Error mensaje: {e}")
            return False
        return True

    def main(self):
        if not all(os.path.exists(p) for p in [self.NO_CONTACT_TEMPLATE, self.ATTACH_BUTTON_TEMPLATE, self.DOCUMENT_BUTTON_TEMPLATE]):
            print("❌ Faltan plantillas")
            return
        driver, wait = self.iniciar_driver()
        if not driver:
            return
        driver.get("https://web.whatsapp.com")
        if not self.esperar_whatsapp_cargado(wait):
            driver.quit()
            return
        exitosos, fallidos = 0, 0
        for contacto in self.CONTACTOS:
            numero = contacto["numero"]
            archivo = contacto["archivo"]
            print(f"\n📱 Procesando: {numero}")
            if not os.path.exists(archivo):
                print(f"❌ Archivo no encontrado para {numero}: {archivo}")
                fallidos += 1
                continue
            if self.abrir_chat_con_contacto(driver, wait, numero):
                if self.enviar_documento_autogui(wait, numero, archivo):
                    exitosos += 1
                else:
                    fallidos += 1
            else:
                fallidos += 1
            time.sleep(5)
        print(f"✅ Exitosos: {exitosos} / ❌ Fallidos: {fallidos}")
        time.sleep(6)
        
        
