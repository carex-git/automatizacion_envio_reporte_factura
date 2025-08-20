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
import datetime
import pickle
import pyautogui


class QuotaManager:
    def __init__(self, limite_diario=40, limite_horario=8):
        self.limite_diario = limite_diario
        self.limite_horario = limite_horario
        self.archivo_datos = "quota_data.pkl"
        self.cargar_datos()

    def cargar_datos(self):
        try:
            with open(self.archivo_datos, 'rb') as f:
                datos = pickle.load(f)
                self.mensajes_hoy = datos.get('mensajes_hoy', 0)
                self.fecha_actual = datos.get('fecha_actual', datetime.date.today())
                self.historial_horas = datos.get('historial_horas', {})
        except (FileNotFoundError, EOFError):
            self.mensajes_hoy = 0
            self.fecha_actual = datetime.date.today()
            self.historial_horas = {}

    def guardar_datos(self):
        datos = {
            'mensajes_hoy': self.mensajes_hoy,
            'fecha_actual': self.fecha_actual,
            'historial_horas': self.historial_horas
        }
        with open(self.archivo_datos, 'wb') as f:
            pickle.dump(datos, f)

    def limpiar_historial_antiguo(self):
        hora_actual = datetime.datetime.now()
        horas_a_eliminar = []
        for hora_str in self.historial_horas:
            hora = datetime.datetime.fromisoformat(hora_str)
            if (hora_actual - hora).total_seconds() > 3600:
                horas_a_eliminar.append(hora_str)
        for hora in horas_a_eliminar:
            del self.historial_horas[hora]

    def puede_enviar(self):
        if datetime.date.today() != self.fecha_actual:
            self.mensajes_hoy = 0
            self.fecha_actual = datetime.date.today()
            self.historial_horas = {}

        self.limpiar_historial_antiguo()

        if self.mensajes_hoy >= self.limite_diario:
            print(f"🚫 Límite diario alcanzado ({self.mensajes_hoy}/{self.limite_diario})")
            return False

        mensajes_ultima_hora = len(self.historial_horas)
        if mensajes_ultima_hora >= self.limite_horario:
            print(f"🚫 Límite por hora alcanzado ({mensajes_ultima_hora}/{self.limite_horario})")
            return False

        if not self.es_horario_permitido():
            return False

        return True

    def es_horario_permitido(self):
        hora_actual = datetime.datetime.now().hour
        if not (7 <= hora_actual <= 21):
            print(f"🚫 Fuera del horario permitido (7AM-9PM). Hora actual: {hora_actual}:00")
            return False
        return True

    def registrar_envio(self):
        self.mensajes_hoy += 1
        hora_actual = datetime.datetime.now().isoformat()
        self.historial_horas[hora_actual] = True
        self.guardar_datos()
        print(f"📊 Mensajes hoy: {self.mensajes_hoy}/{self.limite_diario} | Última hora: {len(self.historial_horas)}/{self.limite_horario}")

    def obtener_tiempo_espera_recomendado(self):
        if len(self.historial_horas) >= self.limite_horario - 2:
            return random.uniform(60, 90)
        elif self.mensajes_hoy >= self.limite_diario - 5:
            return random.uniform(40, 60)
        else:
            return random.uniform(20, 40)


class WhatsAppSafeSender:
    def __init__(self, contacts, mensaje, profile_path, attach_buttons, document_buttons, no_contact_buttons, send_buttons, enviados_dir):
        self.CONTACTOS = contacts
        self.MENSAJE = mensaje
        self.profile_path = profile_path
        self.enviados_dir = enviados_dir
        self.quota_manager = QuotaManager()

        self.ATTACH_BUTTON_TEMPLATE = attach_buttons
        self.DOCUMENT_BUTTON_TEMPLATE = document_buttons
        self.NO_CONTACT_TEMPLATE = no_contact_buttons
        self.SEND_BUTTON_TEMPLATE = send_buttons

        self.CAPTION_BOX_XPATH = '//div[@contenteditable="true"][@data-tab="10"]'
        self.SEARCH_BOX_XPATH = '//div[@contenteditable="true"][@data-tab="3"]'
        self.MESSAGE_BOX_XPATH = '//div[@contenteditable="true"][@data-tab="10"]'

        self.ATTACH_BUTTON_XPATH = '//div[@title="Adjuntar"]'
        self.DOCUMENT_BUTTON_XPATH = '//div[@title="Documento"]'
        self.FILE_INPUT_XPATH = '//input[@type="file"]'
        self.SEND_BUTTON_XPATH = '//span[@data-icon="send"]'
        self.NO_CONTACT_XPATH = '//div[text()="No se encontró ningún chat, contacto ni mensaje."]'
        self.INVALID_PHONE_XPATH = '//div[contains(text(), "El número de teléfono no es un usuario válido de WhatsApp.")]'

        self.mensajes_variados = [
            "Hola {nombre}, {mensaje}",
            "Buenos días {nombre}, {mensaje}",
            "Estimado/a {nombre}, {mensaje}",
            "{nombre}, {mensaje}",
            "Saludos {nombre}, {mensaje}"
        ]

    def iniciar_driver(self):
        try:
            service = Service(ChromeDriverManager().install())
            options = Options()
            options.add_argument("--start-maximized")
            options.add_argument("--disable-blink-features=AutomationControlled")
            options.add_argument("--disable-dev-shm-usage")
            options.add_argument("--no-sandbox")
            # Hacer que el navegador parezca más normal
            options.add_argument("--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")

            profile_path = self.profile_path
            options.add_argument(f"--user-data-dir={profile_path}")

            driver = webdriver.Chrome(service=service, options=options)
            wait = WebDriverWait(driver, 60)
            return driver, wait
        except Exception as e:
            print(f"❌ Error al iniciar el driver: {e}")
            return None, None

    def esperar_whatsapp_cargado(self, wait):
        max_intentos = 3
        for intento in range(max_intentos):
            try:
                wait.until(EC.presence_of_element_located((By.XPATH, self.SEARCH_BOX_XPATH)))
                print("✅ WhatsApp Web cargado")
                return True
            except Exception as e:
                print(f"⚠️ Intento {intento + 1}/{max_intentos} fallido: {e}")
                if intento < max_intentos - 1:
                    print("🔄 Reintentando en 10 segundos…")
                    time.sleep(10)
        print("❌ No se pudo cargar WhatsApp Web después de varios intentos")
        return False
    
    def detectar_bloqueo_o_problema(self, driver):
        """Detecta si WhatsApp está bloqueado o hay problemas"""

        elementos_problema = [
            # Bloqueos y restricciones
            "//div[contains(text(), 'bloqueado')]",
            "//div[contains(text(), 'blocked')]",
            "//div[contains(text(), 'restricted')]",
            "//div[contains(text(), 'temporarily banned')]",
            "//div[contains(text(), 'cuenta suspendida')]",
            "//div[contains(text(), 'violado')]",

            # Errores de red y conexión
            "//div[contains(text(), 'Error de conexión')]",
            "//div[contains(text(), 'Connection failed')]",
            "//div[contains(text(), 'Sin conexión')]",
            "//div[contains(text(), 'Reconectando')]",
            "//div[contains(text(), 'Reconnecting')]",

            # Límites de velocidad
            "//div[contains(text(), 'demasiados intentos')]",
            "//div[contains(text(), 'too many attempts')]",
            "//div[contains(text(), 'Espera un momento')]",
            "//div[contains(text(), 'Wait a moment')]",

            # Captcha o verificaciones
            "//div[contains(text(), 'verificar')]",
            "//div[contains(text(), 'verify')]",
            "//input[@placeholder='Código de verificación']",

            # Problemas de carga
            "//div[contains(text(), 'Cargando')]",
            "//div[contains(text(), 'Loading')]"
        ]

        for xpath in elementos_problema:
            try:
                elementos = driver.find_elements(By.XPATH, xpath)
                if elementos and len(elementos) > 0:
                    texto = elementos[0].text
                    print(f"🚨 PROBLEMA DETECTADO: {texto}")
                    return True, texto
            except Exception:
                continue

        return False, None
    

    def escribir_como_humano(self, element, text, min_delay=0.02, max_delay=0.15):
        """Escribe texto simulando comportamiento humano más realista"""
        for i, char in enumerate(text):
            element.send_keys(char)

            if char in [' ', '.', ',', '!', '?']:
                delay = random.uniform(max_delay * 1.5, max_delay * 3)
            elif char.isalpha():
                delay = random.uniform(min_delay, max_delay)
            else:
                delay = random.uniform(min_delay * 1.2, max_delay * 1.2)

            time.sleep(delay)

            if i > 0 and i % random.randint(15, 25) == 0:
                time.sleep(random.uniform(0.3, 0.8))

        print(f"✅ Texto escrito de forma humana: '{text[:50]}{'...' if len(text) > 50 else ''}'")



    def abrir_chat_con_contacto(self, driver, wait, numero):
        problema, texto = self.detectar_bloqueo_o_problema(driver)
        if problema:
            print(f"🚨 DETENIENDO: {texto}")
            return False

        chat_abierto = False
        try:
            search_box = wait.until(EC.element_to_be_clickable((By.XPATH, self.SEARCH_BOX_XPATH)))
            search_box.clear()
            time.sleep(random.uniform(1, 2))
            self.escribir_como_humano(search_box, numero)
            time.sleep(random.uniform(3, 5))
            try:
                if pyautogui.locateOnScreen(self.NO_CONTACT_TEMPLATE, confidence=0.8, grayscale=True):
                    print(f"❌ Contacto {numero} no encontrado en la búsqueda.")
                    search_box.clear()
                    search_box.send_keys(Keys.ESCAPE)
                else:
                    print(f"✅ Contacto {numero} encontrado en la búsqueda. Abriendo chat…")
                    search_box.send_keys(Keys.ENTER)
                    time.sleep(random.uniform(2, 4))
                    wait.until(EC.presence_of_element_located((By.XPATH, self.MESSAGE_BOX_XPATH)))
                    chat_abierto = True
            except ImageNotFoundException:
                print(f"✅ Contacto {numero} encontrado en la búsqueda. Abriendo chat…")
                search_box.send_keys(Keys.ENTER)
                time.sleep(random.uniform(2, 4))
                wait.until(EC.presence_of_element_located((By.XPATH, self.MESSAGE_BOX_XPATH)))
                chat_abierto = True

        except Exception as e:
            print(f"⚠️ Error al buscar contacto con Selenium y PyAutoGUI: {e}.")

        if not chat_abierto:
            try:
                print(f"🔄 Intentando abrir el chat para {numero} usando URL directa…")
                driver.get(f"https://web.whatsapp.com/send?phone={numero}")
                time.sleep(random.uniform(3, 6))
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
    def generar_mensaje_personalizado(self, nombre):

        """Genera un mensaje personalizado y variado"""

        plantilla = random.choice(self.mensajes_variados)

        # Usar solo el primer nombre para ser más natural
        primer_nombre = nombre.split()[0] if nombre else "Estimado/a"

        return plantilla.format(nombre=primer_nombre, mensaje=self.MENSAJE)
    
    def click_image(self, template_paths, confidence=0.8, timeout=15):
        start_time = time.time()
        if isinstance(template_paths, str):
            template_paths = [template_paths]

        while time.time() - start_time < timeout:
            try:
                for template_path in template_paths:
                    if not os.path.exists(template_path):
                        print(f"⚠️ Archivo no encontrado: {template_path}")
                        continue

                    location = pyautogui.locateCenterOnScreen(template_path, confidence=confidence, grayscale=True)
                    if location:
                        offset_x = random.randint(-2, 2)
                        offset_y = random.randint(-2, 2)
                        pyautogui.click(location.x + offset_x, location.y + offset_y)
                        print(f"✅ Clic en la imagen: {template_path}")
                        return True
            except pyautogui.PyAutoGUIException:
                pass
            time.sleep(1)

        print(f"❌ No se encontró ninguna de las imágenes: {template_paths}")
        return False


    def verificar_estado_chat(self, driver):
        """Verifica si el chat se cargó correctamente"""
        try:
            # Verificar si hay elementos que indican problemas
            problema, texto = self.detectar_bloqueo_o_problema(driver)
            if problema:
                return False, f"Problema detectado: {texto}"

            # Verificar si el mensaje anterior se envió correctamente
            mensajes_enviados = driver.find_elements(By.XPATH, "//span[@data-icon='msg-check' or @data-icon='msg-dblcheck']")
            if not mensajes_enviados:
                print("⚠️ No se detectaron confirmaciones de entrega recientes")

            return True, "OK"

        except Exception as e:
            return False, f"Error al verificar estado: {str(e)}"



    def enviar_documento_autogui(self, wait, numero, archivo, nombre_contacto):
        try:
            problema, texto = self.detectar_bloqueo_o_problema(wait._driver)
            if problema:
                print(f"🚨 PROBLEMA DETECTADO ANTES DEL ENVÍO: {texto}")
                return False

            mensaje_personalizado = self.generar_mensaje_personalizado(nombre_contacto)

            message_box = wait.until(EC.presence_of_element_located((By.XPATH, self.MESSAGE_BOX_XPATH)))
            message_box.send_keys(Keys.CONTROL + 'a')
            message_box.send_keys(Keys.DELETE)
            time.sleep(random.uniform(0.5, 1.0))
            self.escribir_como_humano(message_box, mensaje_personalizado)
            time.sleep(random.uniform(1, 2))

            if not self.click_image(self.ATTACH_BUTTON_TEMPLATE):
                print(f"❌ Fallo al hacer clic en el botón de adjuntar para {numero}.")
                return False
            time.sleep(random.uniform(1, 2))

            if not self.click_image(self.DOCUMENT_BUTTON_TEMPLATE):
                print(f"❌ Fallo al hacer clic en el botón de documento para {numero}.")
                return False
            time.sleep(random.uniform(3, 5))

            ruta_archivo = os.path.abspath(archivo)
            for char in ruta_archivo:
                pyautogui.write(char)
                time.sleep(random.uniform(0.001, 0.003))
            time.sleep(random.uniform(0.5, 1.0))
            pyautogui.press('enter')
            print(f"✅ Archivo seleccionado: {archivo}")
            time.sleep(random.uniform(5, 8))

            problema, texto = self.detectar_bloqueo_o_problema(wait._driver)
            if problema:
                print(f"🚨 PROBLEMA DETECTADO ANTES DEL ENVÍO FINAL: {texto}")
                return False

            if not self.click_image(self.SEND_BUTTON_TEMPLATE):
                print(f"❌ No se pudo encontrar el botón de enviar para {numero}.")
                return False
            time.sleep(random.uniform(3, 5))

            estado_ok, mensaje_estado = self.verificar_estado_chat(wait._driver)
            if not estado_ok:
                print(f"⚠️ Posible problema después del envío: {mensaje_estado}")

            print(f"✅ Documento enviado a {numero}")

        except Exception as e:
            print(f"❌ Error al enviar el documento para {numero}: {e}")
            return False

        return True

    def mover_archivo_enviado(self, archivo, index):
        if not os.path.exists(self.enviados_dir):
            os.makedirs(self.enviados_dir)

        nombre_base = os.path.basename(archivo)
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        nombre_con_indice = f"{index:03d}_{timestamp}_{nombre_base}"
        ruta_destino = os.path.join(self.enviados_dir, nombre_con_indice)

        try:
            os.rename(archivo, ruta_destino)
            print(f"✅ Archivo movido a: {ruta_destino}")
            return True
        except Exception as e:
            print(f"❌ Error al mover el archivo {archivo}: {e}")
            return False
        
    def pausa_inteligente(self):
        """Implementa pausas inteligentes basadas en el estado actual"""
        tiempo_espera = self.quota_manager.obtener_tiempo_espera_recomendado()
        
        # Añadir variabilidad adicional
        variacion = random.uniform(0.5, 0.8)
        tiempo_final = tiempo_espera * variacion
        
        minutos = tiempo_final / 60*3
        print(f"⏳ Pausando por {minutos:.1f} minutos ({tiempo_final:.0f}s) para evitar detección...")
        
        # Pausa en intervalos para permitir interrupciones
        intervalos = int(tiempo_final / 30)  # Intervalos de 30 segundos
        for i in range(intervalos):
            time.sleep(20)
            # Verificar si es un buen momento para continuar
            if not self.quota_manager.es_horario_permitido():
                print("🚫 Fuera del horario permitido. Pausando hasta mañana...")
                return False
        
        # Tiempo restante
        tiempo_restante = tiempo_final - (intervalos * 10)
        if tiempo_restante > 0:
            time.sleep(tiempo_restante)
        
        return True


    def main(self):
        print("🚀 Iniciando WhatsApp Sender Seguro")
        print(f"📊 Estado inicial - Límite diario: {self.quota_manager.limite_diario}, Límite horario: {self.quota_manager.limite_horario}")

        # 📌 Lista central con todas las plantillas necesarias
        self.TEMPLATES = [
            self.NO_CONTACT_TEMPLATE,
            self.ATTACH_BUTTON_TEMPLATE,
            self.DOCUMENT_BUTTON_TEMPLATE,
            self.SEND_BUTTON_TEMPLATE
        ]

        # 📌 Verificación de plantillas
        if not all(os.path.exists(p) for sublist in self.TEMPLATES for p in (sublist if isinstance(sublist, list) else [sublist])):
            print("⚠️ Alguna plantilla no existe")



        if not self.quota_manager.puede_enviar():
            print("🚫 No se puede enviar en este momento debido a limitaciones de cuota u horario.")
            return

        driver, wait = self.iniciar_driver()
        if not driver:
            return

        try:
            driver.get("https://web.whatsapp.com")
            if not self.esperar_whatsapp_cargado(wait):
                driver.quit()
                return

            exitosos, fallidos, detenido_por_seguridad = 0, 0, False
            conteo_enviados = 1

            for i, contacto in enumerate(self.CONTACTOS):
                if not self.quota_manager.puede_enviar():
                    print("🚫 Límite de cuotas alcanzado. Deteniendo envíos.")
                    detenido_por_seguridad = True
                    break

                numero = contacto["numero"]
                archivo = contacto["archivo"]
                nombre = contacto["nombre"]

                print(f"\n📱 Procesando {i+1}/{len(self.CONTACTOS)}: {numero} ({nombre})")

                problema, texto = self.detectar_bloqueo_o_problema(driver)
                if problema:
                    print(f"🚨 DETENIENDO POR SEGURIDAD: {texto}")
                    detenido_por_seguridad = True
                    break

                if not os.path.exists(archivo):
                    print(f"❌ Archivo no encontrado para {numero}: {archivo}")
                    fallidos += 1
                    continue

                if self.abrir_chat_con_contacto(driver, wait, numero):
                    if self.enviar_documento_autogui(wait, numero, archivo, nombre):
                        if self.mover_archivo_enviado(archivo, conteo_enviados):
                            exitosos += 1
                            conteo_enviados += 1
                            self.quota_manager.registrar_envio()
                        else:
                            fallidos += 1
                    else:
                        fallidos += 1
                        problema, texto = self.detectar_bloqueo_o_problema(driver)
                        if problema:
                            print(f"🚨 DETENIENDO POR PROBLEMA DETECTADO: {texto}")
                            detenido_por_seguridad = True
                            break
                else:
                    fallidos += 1

                if i < len(self.CONTACTOS) - 1:
                    if not self.pausa_inteligente():
                        print("🚫 Deteniendo por horario no permitido.")
                        detenido_por_seguridad = True
                        break

        except KeyboardInterrupt:
            print("\n⏹️ Proceso interrumpido por el usuario.")
        except Exception as e:
            print(f"\n❌ Error inesperado: {e}")
        finally:
            driver.quit()

        print(f"\n📊 === RESUMEN FINAL ===")
        print(f"✅ Enviados exitosamente: {exitosos}")
        print(f"❌ Fallidos: {fallidos}")
        print(f"🛡️ Detenido por seguridad: {'Sí' if detenido_por_seguridad else 'No'}")
        print(f"📈 Total de mensajes hoy: {self.quota_manager.mensajes_hoy}/{self.quota_manager.limite_diario}")
        if detenido_por_seguridad:
            print("\n⚠️ IMPORTANTE: El proceso se detuvo por medidas de seguridad.")
            print("   Esto ayuda a proteger tu cuenta de posibles bloqueos.")
            print("   Puedes intentar nuevamente más tarde respetando los límites.")
