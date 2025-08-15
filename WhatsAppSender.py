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
    """
    Clase para automatizar el env\u00edo de documentos a contactos de WhatsApp Web.
    La l\u00f3gica ha sido optimizada para usar la b\u00fasqueda como m\u00e9todo principal
    y la URL directa como fallback \u00fanicamente si la b\u00fasqueda falla.
    """
    def __init__(self, contacts, mensaje,profile_path,attach_buuton, document_button,no_contact_button):
        self.CONTACTOS = contacts
        self.MENSAJE = mensaje
        self.profile_path = profile_path

        # Rutas a las plantillas de imagen
        self.ATTACH_BUTTON_TEMPLATE = attach_buuton
        self.DOCUMENT_BUTTON_TEMPLATE = document_button
        self.NO_CONTACT_TEMPLATE = no_contact_button
        # XPaths para elementos de WhatsApp Web
        self.CAPTION_BOX_XPATH = '//div[@contenteditable="true"][@data-tab="10"]'
        self.SEARCH_BOX_XPATH = '//div[@contenteditable="true"][@data-tab="3"]'

    def click_image(self, template_path, confidence=0.8, timeout=15):
        """
        Busca y hace clic en una imagen en la pantalla usando PyAutoGUI.
        Retorna True si se encontr\u00f3 y se hizo clic, False si no.
        El timeout ha sido aumentado para manejar la carga de p\u00e1gina m\u00e1s lenta.
        """
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
            except pyautogui.PyAutoGUIException:
                pass  # La excepci\u00f3n es normal si la imagen no se encuentra, seguimos intentando.
            time.sleep(1)
        print(f"❌ No se encontr\u00f3 la imagen: {template_path}")
        return False

    def iniciar_driver(self):
        """
        Configura y devuelve el driver de Selenium para Chrome.
        """
        try:
            service = Service(ChromeDriverManager().install())
            options = Options()
            options.add_argument("--start-maximized")
            options.add_argument("--disable-blink-features=AutomationControlled")
            profile_path = self.profile_path
            options.add_argument(f"--user-data-dir={profile_path}")
            driver = webdriver.Chrome(service=service, options=options)
            wait = WebDriverWait(driver, 60)
            return driver, wait
        except Exception as e:
            print(f"❌ Error al iniciar el driver: {e}")
            return None, None

    def esperar_whatsapp_cargado(self, wait):
        """
        Espera hasta que WhatsApp Web est\u00e9 completamente cargado y listo.
        """
        try:
            wait.until(EC.presence_of_element_located((By.XPATH, self.SEARCH_BOX_XPATH)))
            print("✅ WhatsApp Web cargado")
            return True
        except Exception:
            print("❌ No se pudo cargar WhatsApp Web")
            return False

    def abrir_chat_con_contacto(self, driver, wait, numero):
        """
        Intenta abrir un chat con un contacto. Prioriza la b\u00fasqueda.
        Solo usa la URL directa si la b\u00fasqueda falla completamente.
        """
        chat_abierto = False
        try:
            search_box = wait.until(EC.element_to_be_clickable((By.XPATH, self.SEARCH_BOX_XPATH)))
            search_box.clear()
            search_box.send_keys(numero)
            time.sleep(3)
            
            try:
                pyautogui.locateOnScreen(self.NO_CONTACT_TEMPLATE, confidence=0.8, grayscale=True)
                print(f"❌ Contacto {numero} no encontrado en la b\u00fasqueda.")
            except ImageNotFoundException:
                print(f"✅ Contacto {numero} encontrado en la b\u00fasqueda. Abriendo chat...")
                search_box.send_keys(Keys.ENTER)
                wait.until(EC.presence_of_element_located((By.XPATH, self.CAPTION_BOX_XPATH)))
                chat_abierto = True
        
        except Exception as e:
            print(f"⚠️ Error al buscar contacto con Selenium y PyAutoGUI: {e}.")
        
        # Si la b\u00fasqueda no tuvo \u00e9xito, usar la URL directa como fallback.
        if not chat_abierto:
            try:
                print(f"🔄 Intentando abrir el chat para {numero} usando URL directa...")
                driver.get(f"https://web.whatsapp.com/send?phone={numero}")
                wait.until(EC.presence_of_element_located((By.XPATH, self.CAPTION_BOX_XPATH)))
                # A\u00f1adimos una pausa para que la UI cargue los botones
                time.sleep(5)
                print(f"✅ Chat abierto para {numero} usando URL directa.")
                chat_abierto = True
            except Exception as e:
                print(f"❌ Fall\u00f3 el m\u00e9todo de URL directa para {numero}. Error: {e}")
                driver.get("https://web.whatsapp.com")
                self.esperar_whatsapp_cargado(wait)
        
        return chat_abierto

    def enviar_documento_autogui(self, wait, numero, archivo):
        """
        Usa PyAutoGUI para adjuntar y enviar un archivo.
        """
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
        """
        Funci\u00f3n principal que orquesta todo el proceso de env\u00edo.
        """
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
