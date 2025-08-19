import pyautogui
import cv2
import numpy as np
import time
import os
import random 
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
from pyautogui import ImageNotFoundException
from selenium.common.exceptions import TimeoutException, NoSuchElementException, NoSuchWindowException
import pandas as pd
import json

class WhatsAppSender:
    """
    Clase para automatizar el envío de documentos a contactos de WhatsApp Web.
    La lógica ha sido optimizada para usar la búsqueda como método principal
    y la URL directa como fallback únicamente si la búsqueda falla.
    """
    def __init__(self, contacts, mensaje, profile_path, attach_button, document_button, no_contact_button, enviados_dir):
        self.CONTACTOS = contacts
        self.MENSAJE = mensaje
        self.profile_path = profile_path
        self.enviados_dir = enviados_dir

        # Rutas a las plantillas de imagen
        self.ATTACH_BUTTON_TEMPLATE = attach_button
        self.DOCUMENT_BUTTON_TEMPLATE = document_button
        self.NO_CONTACT_TEMPLATE = no_contact_button


        # XPaths para elementos de WhatsApp Web
        self.MESSAGE_BOX_XPATH = '//div[@contenteditable="true"][@data-tab="10"]'
        self.SEARCH_BOX_XPATH = '//div[@contenteditable="true"][@data-tab="3"]'
        self.INVALID_PHONE_XPATH = '//div[contains(text(), "El número de teléfono no es un usuario válido de WhatsApp.")]'
        
        # Banco de mensajes
        self.MENSAJES_BANCO = [
            "Hola {nombre}, te adjunto el documento que me pediste. Saludos cordiales.",
            "Qué tal {nombre}, espero que estés bien. Aquí está el archivo que necesitas. ¡Gracias!",
            "Hola {nombre}, adjunto el documento solicitado. ¡Feliz día!",
            "Buenos días {nombre}, aquí está tu archivo. ¡Cualquier cosa me avisas!",
            "Hola {nombre}, envío el documento. Un saludo."
        ]
    
    def escribir_como_humano(self, element, text, min_delay=0.01, max_delay=0.1):
        """
        Escribe el texto en un elemento de forma pausada para simular un humano.
        """
        for char in text:
            element.send_keys(char)
            time.sleep(random.uniform(min_delay, max_delay))
        print(f"✅ Texto escrito de forma humana: '{text}'")

    def click_image(self, template_path, confidence=0.8, timeout=15):
        """
        Busca y hace clic en una imagen en la pantalla usando PyAutoGUI.
        Retorna True si se encontró y se hizo clic, False si no.
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
                pass
            time.sleep(1)
        print(f"❌ No se encontró la imagen: {template_path}")
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
        Espera hasta que WhatsApp Web esté completamente cargado y listo.
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
        Intenta abrir un chat con un contacto. Prioriza la búsqueda.
        Solo usa la URL directa si la búsqueda falla completamente.
        """
        chat_abierto = False
        try:
            search_box = wait.until(EC.element_to_be_clickable((By.XPATH, self.SEARCH_BOX_XPATH)))
            search_box.clear()
            self.escribir_como_humano(search_box, numero)
            
            time.sleep(3) 
            
            try:
                if pyautogui.locateOnScreen(self.NO_CONTACT_TEMPLATE, confidence=0.8, grayscale=True):
                    print(f"❌ Contacto {numero} no encontrado en la búsqueda.")
                    search_box.clear() 
                    search_box.send_keys(Keys.ESCAPE)
                else:
                    print(f"✅ Contacto {numero} encontrado en la búsqueda. Abriendo chat...")
                    search_box.send_keys(Keys.ENTER)
                    wait.until(EC.presence_of_element_located((By.XPATH, self.MESSAGE_BOX_XPATH)))
                    chat_abierto = True
            except ImageNotFoundException:
                print(f"✅ Contacto {numero} encontrado en la búsqueda. Abriendo chat...")
                search_box.send_keys(Keys.ENTER)
                wait.until(EC.presence_of_element_located((By.XPATH, self.MESSAGE_BOX_XPATH)))
                chat_abierto = True
        
        except Exception as e:
            print(f"⚠️ Error al buscar contacto con Selenium y PyAutoGUI: {e}.")
        
        if not chat_abierto:
            try:
                print(f"🔄 Intentando abrir el chat para {numero} usando URL directa...")
                driver.get(f"https://web.whatsapp.com/send?phone={numero}")
                wait.until(EC.any_of(
                    EC.presence_of_element_located((By.XPATH, self.MESSAGE_BOX_XPATH)),
                    EC.presence_of_element_located((By.XPATH, self.INVALID_PHONE_XPATH))
                ))
                
                if len(driver.find_elements(By.XPATH, self.INVALID_PHONE_XPATH)) > 0:
                    print(f"❌ El número {numero} no es un usuario válido de WhatsApp.")
                    driver.get("https://web.whatsapp.com")
                    self.esperar_whatsapp_cargado(wait)
                    return False
                else:
                    print(f"✅ Chat abierto para {numero} usando URL directa.")
                    chat_abierto = True
            except TimeoutException as e:
                print(f"❌ Falló el método de URL directa para {numero}. Error de tiempo de espera: {e}")
                driver.get("https://web.whatsapp.com")
                self.esperar_whatsapp_cargado(wait)
            except NoSuchWindowException:
                print("❌ Se cerró la ventana del navegador. Terminando el script.")
                return False
        
        return chat_abierto
    
    def enviar_documento_autogui(self, wait, numero, archivo, nombre_contacto):
        """
        Usa PyAutoGUI para adjuntar un archivo y Selenium para enviarlo.
        Ahora recibe el nombre del contacto para personalizar el mensaje.
        """
        try:
            # Obtener las partes del nombre
            partes_nombre = nombre_contacto.split()
            
            # Validar si el nombre tiene suficientes partes (al menos 3)
            if len(partes_nombre) >= 3:
                # Obtener el primer y tercer nombre
                nombre_personalizado = f"{partes_nombre[0]} {partes_nombre[2]}"
            else:
                # Si el nombre no tiene al menos tres partes, usar solo el primero.
                nombre_personalizado = partes_nombre[0] if partes_nombre else "Estimado"

            # Seleccionar un mensaje aleatorio del banco y formatearlo
            mensaje_elegido = random.choice(self.MENSAJES_BANCO)
            mensaje_personalizado = mensaje_elegido.format(nombre=nombre_personalizado)
            
            # Encontrar el cuadro de mensaje, borrar cualquier texto y escribir el mensaje humanizado
            message_box = wait.until(EC.presence_of_element_located((By.XPATH, self.MESSAGE_BOX_XPATH)))
            message_box.send_keys(Keys.CONTROL + 'a')
            message_box.send_keys(Keys.DELETE)
            time.sleep(0.5) 
            self.escribir_como_humano(message_box, mensaje_personalizado)
            time.sleep(1) 
            
            # Clic en el botón de adjuntar y seleccionar documento
            if not self.click_image(self.ATTACH_BUTTON_TEMPLATE):
                print(f"❌ Fallo al hacer clic en el botón de adjuntar para {numero}.")
                return False
            
            if not self.click_image(self.DOCUMENT_BUTTON_TEMPLATE):
                print(f"❌ Fallo al hacer clic en el botón de documento para {numero}.")
                return False
            
            time.sleep(3)
            
            # Escribir ruta del archivo de forma más humana
            ruta_archivo = os.path.abspath(archivo)
            for char in ruta_archivo:
                pyautogui.write(char)
                time.sleep(random.uniform(0.01, 0.03))
            
            time.sleep(random.uniform(0.5, 1.0))
            pyautogui.press('enter')
            print(f"✅ Archivo seleccionado: {archivo}")
            # Esperar un tiempo prudente antes de hacer clic en enviar
            time.sleep(3)
            
            print(f"✅ Documento enviado a {numero}")
            
        except Exception as e:
            print(f"❌ Error al enviar el documento para {numero}: {e}")
            return False
        
        return True

    def mover_archivo_enviado(self, archivo, index):
        """
        Mueve un archivo a la carpeta 'enviados' con un índice al principio del nombre.
        """
        if not os.path.exists(self.enviados_dir):
            os.makedirs(self.enviados_dir)
        
        nombre_base = os.path.basename(archivo)
        nombre_con_indice = f"{index:03d}_{nombre_base}"
        ruta_destino = os.path.join(self.enviados_dir, nombre_con_indice)
        
        try:
            os.rename(archivo, ruta_destino)
            print(f"✅ Archivo movido a: {ruta_destino}")
            return True
        except Exception as e:
            print(f"❌ Error al mover el archivo {archivo}: {e}")
            return False

    def main(self):
        """
        Función principal que orquesta todo el proceso de envío.
        """
        if not all(os.path.exists(p) for p in [self.NO_CONTACT_TEMPLATE, self.ATTACH_BUTTON_TEMPLATE, self.DOCUMENT_BUTTON_TEMPLATE]):
            print("❌ Error: Faltan plantillas de imagen. Asegúrate de que todas existen en la ruta especificada.")
            return
        
        driver, wait = self.iniciar_driver()
        if not driver:
            return
        
        driver.get("https://web.whatsapp.com")
        if not self.esperar_whatsapp_cargado(wait):
            driver.quit()
            return
        
        exitosos, fallidos = 0, 0
        conteo_enviados = 1 
        for contacto in self.CONTACTOS:
            numero = contacto["numero"]
            archivo = contacto["archivo"]
            nombre = contacto["nombre"]
            
            print(f"\n📱 Procesando: {numero} ({nombre})")
            
            if not os.path.exists(archivo):
                print(f"❌ Archivo no encontrado para {numero}: {archivo}")
                fallidos += 1
                continue
            
            if self.abrir_chat_con_contacto(driver, wait, numero):
                if self.enviar_documento_autogui(wait, numero, archivo, nombre):
                    if self.mover_archivo_enviado(archivo, conteo_enviados):
                        exitosos += 1
                        conteo_enviados += 1 
                    else:
                        fallidos += 1
                else:
                    fallidos += 1
            else:
                fallidos += 1
            
            pausa = random.uniform(10, 20)
            print(f"⏳ Pausando por {pausa:.2f} segundos para evitar bloqueos...")
            time.sleep(pausa)
            
        driver.quit()
        print(f"\nResumen de envío:")
        print(f"✅ Exitosos: {exitosos}")
        print(f"❌ Fallidos: {fallidos}")