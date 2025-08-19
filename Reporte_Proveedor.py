"""
Generador de Reportes de Facturación PDF
========================================

Este módulo contiene todas las clases y funciones necesarias para cargar datos
desde un archivo Excel y generar reportes de facturación individuales en PDF.

Autor: Sistema de Facturación
Fecha: 2025
"""

import pandas as pd
from fpdf import FPDF, XPos, YPos
import os
import math
from typing import Dict, List, Any, Optional

class ConfiguracionReporte:
    """Configuración centralizada para el reporte PDF"""
    
    # Configuración de página
    ORIENTACION = 'L'  # Landscape
    UNIDAD = 'mm'
    FORMATO = 'A4'
    
    # Márgenes
    MARGEN_AUTO = 15
    MARGEN_LATERAL = 10
    MARGEN_VERTICAL = 10
    
    # Colores (RGB)
    COLOR_BORDE = (200, 200, 200)
    COLOR_ENCABEZADO = (230, 230, 230)
    COLOR_FILA_PAR = (245, 245, 245)
    COLOR_FILA_IMPAR = (255, 255, 255)
    
    # Medidas
    ALTURA_LINEA = 4
    
    # Rutas
    DIRECTORIO_SALIDA_TEL = 'output/tel'
    DIRECTORIO_SALIDA_EMAIL = 'output/email'
    DIRECTORIO_SALIDA = 'output'
    RUTA_LOGO = './logo.png'
    
    # Información de la empresa (valores por defecto - ahora se pueden sobrescribir)
    EMPRESA_NIT = '800.176.428-6'
    EMPRESA_NOMBRE = 'COMERCIALIZADORA INTERNACIONAL CARIBBEAN EXOTICS S. A.'
    EMPRESA_DIRECCION = 'Vda cimarronas Km 1 Rionegro-Marinilla'
    
    # Configuración de tabla
    ENCABEZADOS_TABLA = [
        "FRUTA", "FECHA\nINGRESO", "KILOS\nRECIBIDOS", "% EXP", "EXP", "NAL", "AVE",
        "PRECIO EXP", "PRECIO NAL", "PRECIO AVE", "VALOR\nBRUTO", 
        "RETENCIÓN\nFUENTE", "RETENCIÓN\nFONDO", "VALOR\nTOTAL"
    ]
    
    PROPORCIONES_COLUMNAS = [
        0.06, 0.06, 0.07, 0.06, 0.07, 0.07, 0.05, 
        0.08, 0.08, 0.08, 0.08, 0.07, 0.07, 0.09
    ]
    
    ALINEACIONES_COLUMNAS = [
        'L', 'C', 'R', 'R', 'R', 'R', 'R', 
        'R', 'R', 'R', 'R', 'R', 'R', 'R'
    ]


class UtilFormato:
    """Utilidades para formatear datos en el reporte"""
    
    @staticmethod
    def formatear_fecha(fecha) -> str:
        """Formatea una fecha para mostrar en el reporte"""
        if pd.notna(fecha) and isinstance(fecha, pd.Timestamp):
            return fecha.strftime('%Y-%m-%d')
        return str(fecha) if fecha else ''
    
    @staticmethod
    def formatear_numero(valor: float, decimales: int = 2) -> str:
        """Formatea un número con separadores de miles"""
        try:
            return f"{float(valor):,.{decimales}f}"
        except (ValueError, TypeError):
            return "0.00"
    
    @staticmethod
    def formatear_moneda(valor: float) -> str:
        """Formatea un valor como moneda"""
        return f"${UtilFormato.formatear_numero(valor)}"
    
    @staticmethod
    def formatear_porcentaje(numerador: float, denominador: float) -> str:
        """Calcula y formatea un porcentaje"""
        if denominador > 0:
            porcentaje = (numerador / denominador) * 100
            return f"{porcentaje:.2f}%"
        return "0.00%"


class GestorDatos:
    """Clase para gestionar la carga y procesamiento de datos desde Excel"""
    
    def __init__(self, archivo_excel: str):
        self.archivo_excel = archivo_excel
        self.df_liquidacion = pd.DataFrame()
        self.df_cer_fl_gl = pd.DataFrame()
        self.df_liquidacion = None
        self.df_bd_pro = None
        self.df_cer_fl_gl = None
        self.df_tel_email = None  # NUEVA HOJA PARA TELÉFONOS
        
    def cargar_datos(self):
        """Carga todos los datos necesarios desde el archivo Excel"""
        try:
            # Cargar hoja principal BD LIQUIDACION
            self.df_liquidacion = pd.read_excel(
                self.archivo_excel, 
                sheet_name='BD LIQUIDACION'
            )
            
            self.df_liquidacion.columns = self.df_liquidacion.columns.str.strip()
            
            # Cargar hoja BD PRO para direcciones
            try:
                self.df_bd_pro = pd.read_excel(
                    self.archivo_excel, 
                    sheet_name='BD PRO'
                )
                self.df_bd_pro.columns = self.df_bd_pro.columns.str.strip()
                print("Columnas de 'BD PRO':", self.df_bd_pro.columns)
            except ValueError:
                print("⚠️ Advertencia: No se encontró la hoja 'BD PRO'. Se usará dirección por defecto.")
                self.df_bd_pro = pd.DataFrame()
            
            # Cargar hoja CER FL GL para certificaciones
            try:
                self.df_cer_fl_gl = pd.read_excel(
                    self.archivo_excel, 
                    sheet_name='CER FL GL'
                )
                
                self.df_cer_fl_gl.columns = self.df_cer_fl_gl.columns.str.strip()
                
                print("Columnas de 'CER FL GL':", self.df_cer_fl_gl.columns)
            except ValueError:
                print("⚠️ Advertencia: No se encontró la hoja 'CER FL GL'. Se usarán certificaciones por defecto.")
                self.df_cer_fl_gl = pd.DataFrame()
                
            # Cargar hoja INFO PRO para TELEFONOS Y CORREOS
            try:
                self.df_tel_email = pd.read_excel(
                    self.archivo_excel, 
                    sheet_name='INFO PRO'
                )
                self.df_tel_email.columns = self.df_tel_email.columns.str.strip()
                print("Columnas de 'INFO PRO':", self.df_tel_email.columns)
            except ValueError:
                print("⚠️ Advertencia: No se encontró la hoja 'INFO PRO'. Se usará dirección por defecto.")
                self.df_bd_pro = pd.DataFrame()
            
            return self._limpiar_datos()
            
        except FileNotFoundError:
            raise FileNotFoundError(f"El archivo '{self.archivo_excel}' no se encontró.")
        except ValueError as e:
            raise ValueError(f"Error al cargar datos: {str(e)}")
    
    def _limpiar_datos(self):
        """Limpia y prepara los datos para el procesamiento"""
        # Limpiar datos principales
        self.df_liquidacion['NOMBRE'] = self.df_liquidacion['NOMBRE'].fillna('Sin Nombre')
        self.df_liquidacion['CEDULA'] = self.df_liquidacion['CEDULA'].fillna('000000')
        
        # Asegurar que las columnas numéricas estén correctas
        columnas_numericas = [
            'KILOS RECIBIDOS', 'KG. EXP', 'KG. NAL', 'KG. AVE',
            'PRECIO EXP', 'PRECIO NAL', 'PRECIO AVE',
            'TOTAL BRUTO', 'RETE FUENTE', 'FONDO HORTIFRU', 
            'D 2500', 'DES ANALISIS', 'VALOR TOTAL'
        ]
        
        for col in columnas_numericas:
            if col in self.df_liquidacion.columns:
                self.df_liquidacion[col] = pd.to_numeric(
                    self.df_liquidacion[col], errors='coerce'
                ).fillna(0)
        
        return self.df_liquidacion
    
    def obtener_info_cliente(self, cedula: str) -> Dict[str, Any]:
        """Obtiene información completa del cliente desde BD PRO"""
        info = {
            'direccion': '',  # Por defecto
            'municipio': '',
            'telefono': '',
            'email': ''
        }
        
        if not self.df_bd_pro.empty and cedula:
            cliente_info = self.df_bd_pro[
                self.df_bd_pro['Código'].astype(str).values == str(cedula)
            ]
            
            if not cliente_info.empty:
                fila = cliente_info.iloc[0]
                
                if 'Dirección 1' in cliente_info.columns and pd.notna(fila['Dirección 1']):
                    info['direccion'] = str(fila['Dirección 1'])
                
                if 'Ciudad' in cliente_info.columns and pd.notna(fila['Ciudad']):
                    info['municipio'] = str(fila['Ciudad'])
                    
                    
        if not self.df_tel_email.empty and cedula:
            telefono_info = self.df_tel_email[
                self.df_tel_email['CEDULA'].astype(str) == str(cedula)
            ]
            
            if not telefono_info.empty:
                fila_tel = telefono_info.iloc[0]
                if 'WHATSAPP' in telefono_info.columns and pd.notna(fila_tel['WHATSAPP'] and fila_tel['WHATSAPP'] != '0' and fila_tel['WHATSAPP'] != 0):
                    info['telefono'] = str(fila_tel['WHATSAPP'])
                    
        if not self.df_tel_email.empty and cedula:
            telefono_info = self.df_tel_email[
                self.df_tel_email['CEDULA'].astype(str) == str(cedula)
            ]
            
            if not telefono_info.empty:
                fila_tel = telefono_info.iloc[0]
                if 'EMAIL 1' in telefono_info.columns and pd.notna(fila_tel['EMAIL 1'] and fila_tel['EMAIL 1'] != '0' and fila_tel['EMAIL 1'] != 0):
                    info['email'] = str(fila_tel['EMAIL 1'])
        
        return info
    
    def obtener_certificacion(self, cedula: str, cert_tipo: str) -> Dict[str, str]:
        """
        Determina las certificaciones de un cliente buscando su cédula en
        la hoja 'CER FL GL' y devuelve los datos en un diccionario.
        """
        certificaciones = {'flo': '', 'gap': ''}
        
        if pd.isna(cedula) or not str(cedula).strip():
            return certificaciones

        if not self.df_cer_fl_gl.empty:
            # Buscar por cédula en la hoja de certificaciones
            
            cert_info = self.df_cer_fl_gl[
                self.df_cer_fl_gl['CEDULA'].astype(str) == str(cedula)
            ]

            if not cert_info.empty:
                fila = cert_info.iloc[0]
                cert_tipo = cert_tipo.upper().strip()

                if 'FL' in cert_tipo:
                    id_asoc = str(fila.get('CERTIFICADO ASOC', '')).strip()
                    id_carex = str(fila.get('CERTIFICADO CAREX', '')).strip()
                    
                    if id_asoc or id_carex:
                        certificaciones['flo'] = (
                            "CERTIFICADO FLO\n"
                            f"ID. ASOC. {id_asoc}\n"
                            f"ID. CAREX {id_carex}"
                        )

                if 'GL' in cert_tipo:
                    ggn = str(fila.get('CERTIFICADO GLO', '')).strip()
                    codigo = str(fila.get('CODIGO', '')).strip()
                    if ggn:
                        certificaciones['gap'] = f"GLOBALG.A.P.\n{codigo}"
        
        return certificaciones


class ReporteProveedor(FPDF):
    """
    Clase principal para generar reportes de facturación en PDF
    
    Esta clase hereda de FPDF y añade funcionalidades específicas
    para generar reportes de compra con formato personalizado.
    """
    
    def __init__(self, nit_empresa: Optional[str] = None, nombre_empresa: Optional[str] = None, 
                 direccion_empresa: Optional[str] = None, subtitle: Optional[str] = None):
        """
        Inicializa el generador de reportes con configuración dinámica de empresa
        
        Args:
            nit_empresa: NIT de la empresa (opcional)
            nombre_empresa: Nombre de la empresa (opcional)
            direccion_empresa: Dirección de la empresa (opcional)
        """
        super().__init__(
            ConfiguracionReporte.ORIENTACION, 
            ConfiguracionReporte.UNIDAD, 
            ConfiguracionReporte.FORMATO
        )
        
        # Configurar información de la empresa (usar valores dinámicos o por defecto)
        self.empresa_nit = nit_empresa or ConfiguracionReporte.EMPRESA_NIT
        self.empresa_nombre = nombre_empresa or ConfiguracionReporte.EMPRESA_NOMBRE
        self.empresa_direccion = direccion_empresa or ConfiguracionReporte.EMPRESA_DIRECCION
        self.subtitle = subtitle
        
        self._configurar_pdf()
        self._crear_directorio_salida()
        self.certificacion_flo = ""
        self.certificacion_gap = ""
        self.DIRECTORIO_SALIDA = ConfiguracionReporte.DIRECTORIO_SALIDA
        self.DIRECTORIO_SALIDA_EMAIL = ConfiguracionReporte.DIRECTORIO_SALIDA_EMAIL
        self.DIRECTORIO_SALIDA_TEL = ConfiguracionReporte.DIRECTORIO_SALIDA_TEL
    
    def _configurar_pdf(self):
        """Configura los parámetros básicos del PDF"""
        self.set_auto_page_break(
            auto=True, 
            margin=ConfiguracionReporte.MARGEN_AUTO
        )
        self.set_margins(
            ConfiguracionReporte.MARGEN_LATERAL,
            ConfiguracionReporte.MARGEN_VERTICAL,
            ConfiguracionReporte.MARGEN_LATERAL
        )
    
    def _crear_directorio_salida(self):
        """Crea el directorio de salida si no existe"""
        os.makedirs(ConfiguracionReporte.DIRECTORIO_SALIDA, exist_ok=True)
        os.makedirs(ConfiguracionReporte.DIRECTORIO_SALIDA_TEL, exist_ok=True)
        os.makedirs(ConfiguracionReporte.DIRECTORIO_SALIDA_EMAIL, exist_ok=True)
    
    def header(self):
        """Define el encabezado de cada página"""
        self._agregar_logo()
        self._agregar_titulo_principal()
        self._agregar_subtitulo()
        self.ln(8)
    
    def _agregar_logo(self):
        """Agrega el logo de la empresa al encabezado"""
        if os.path.exists(ConfiguracionReporte.RUTA_LOGO):
            self.image(
                ConfiguracionReporte.RUTA_LOGO, 
                x=10, y=0, w=30
            )
    
    def _agregar_titulo_principal(self):
        """Agrega el título principal del documento"""
        self.set_font('Helvetica', 'B', 16)
        titulo = 'DOCUMENTO FACTURA DE COMPRA'
        x_posicion = self.w / 2 - self.get_string_width(titulo) / 2
        self.set_x(x_posicion)
        self.cell(
            90, 10, titulo, 
            align='C', 
            new_x=XPos.CENTER, 
            new_y=YPos.NEXT
        )
    
    def _agregar_subtitulo(self):
        """Agrega el subtítulo del documento"""
        subtitle = self.subtitle
        x_posicion = self.w / 2 - self.get_string_width(subtitle) / 2
        self.set_x(x_posicion)
        self.set_font('Helvetica', '', 12)
        self.cell(
            90, 10, subtitle,
            align='C', 
            new_x=XPos.LMARGIN, 
            new_y=YPos.NEXT
        )

    def footer(self):
        """Define el pie de página"""
        self.set_y(-15)
        self.set_font('Helvetica', 'I', 8)
        self.cell(
            0, 10, f'Página {self.page_no()}',
            align='C', 
            new_x=XPos.LMARGIN, 
            new_y=YPos.NEXT
        )

    def _agregar_certificacion_final(self):
        """
        Agrega los recuadros de certificación. 
        Este método se debe llamar explícitamente al final del documento.
        """
        self.set_y(-25) 
        self.set_font('Helvetica', 'B', 10)
        
        ancho_cert = 50
        altura_cert = 15
        
        # Dibujar el recuadro de la certificación FLO si existe
        if self.certificacion_flo:
            self.set_x(10)
            self.set_fill_color(230, 230, 230)
            self.rect(self.get_x(), self.get_y(), ancho_cert, altura_cert, 'DF')
            self.multi_cell(
                ancho_cert, 4, self.certificacion_flo,
                border=1, align='C',
                new_x=XPos.RIGHT, new_y=YPos.TOP
            )
        
        # Dibujar el recuadro de la certificación GLOBALG.A.P. si existe
        if self.certificacion_gap:
            self.set_x(10 + ancho_cert + 5 if self.certificacion_flo else 10)
            self.set_y(-25)
            self.set_fill_color(230, 230, 230)
            self.rect(self.get_x(), self.get_y(), ancho_cert, altura_cert, 'DF')
            self.multi_cell(
                ancho_cert, 4, self.certificacion_gap,
                border=1, align='C',
                new_x=XPos.RIGHT, new_y=YPos.TOP
            )
        
        self.set_y(-15)
    
    def establecer_certificacion(self, certificaciones: Dict[str, str]):
        """Establece los textos de certificación para mostrar en el footer"""
        self.certificacion_flo = certificaciones.get('flo', '')
        self.certificacion_gap = certificaciones.get('gap', '')
    
    def agregar_informacion_cliente(self, datos_cliente: pd.DataFrame, info_adicional: Dict[str, Any]):
        """
        Agrega la información del cliente al reporte con formato mejorado
        """
        self._configurar_colores_info()
        
        ancho_total = self.w - self.l_margin - self.r_margin
        ancho_izquierda = ancho_total * 0.4
        ancho_derecha = ancho_total * 0.4
        
        self._agregar_fila_nombre_comprador_mejorada(datos_cliente, ancho_izquierda, ancho_derecha)
        self._agregar_fila_cedulas_mejorada(datos_cliente, ancho_izquierda, ancho_derecha)
        self._agregar_fila_direcciones_mejorada(ancho_izquierda, ancho_derecha, info_adicional)
        
        if info_adicional.get('municipio') or info_adicional.get('telefono'):
            self._agregar_fila_municipio_telefono(ancho_izquierda, ancho_derecha, info_adicional)
        
        self._agregar_titulo_detalle()
    
    def _configurar_colores_info(self):
        """Configura los colores para la sección de información"""
        self.set_fill_color(*ConfiguracionReporte.COLOR_ENCABEZADO)
        self.set_text_color(0)
    
    def _agregar_fila_nombre_comprador_mejorada(self, datos_cliente: pd.DataFrame, ancho_izq: float, ancho_der: float):
        """Agrega la fila con nombre del cliente y comprador con formato mejorado"""
        y_inicio = self.get_y()
        x_inicio = self.get_x()
        
        # Usar la información dinámica de la empresa
        comprador_text = self.empresa_nombre
        nombre_cliente = str(datos_cliente.iloc[0]['NOMBRE'])
        
        ancho_etiqueta_izq = ancho_izq * 0.35
        ancho_valor_izq = ancho_izq * 0.65
        ancho_etiqueta_der = ancho_der * 0.35
        ancho_valor_der = ancho_der * 0.65
        
        self.set_font('Helvetica', '', 10)
        altura_cliente = self._calcular_altura_texto(nombre_cliente, ancho_valor_izq)
        altura_comprador = self._calcular_altura_texto(comprador_text, ancho_valor_der)
        altura_total = max(altura_cliente, altura_comprador, 8)
        
        self.set_xy(x_inicio, y_inicio)
        self._crear_celda_con_altura_fija('NOMBRE / RAZÓN SOCIAL', ancho_etiqueta_izq, altura_total, es_etiqueta=True)
        self._crear_celda_con_altura_fija(nombre_cliente, ancho_valor_izq, altura_total, es_etiqueta=False)
        
        x_derecha = x_inicio + ancho_izq + 5
        self.set_xy(x_derecha, y_inicio)
        self._crear_celda_con_altura_fija('COMPRADOR', ancho_etiqueta_der, altura_total, es_etiqueta=True)
        self._crear_celda_con_altura_fija(comprador_text, ancho_valor_der, altura_total, es_etiqueta=False)
        
        self.set_xy(x_inicio, y_inicio + altura_total + 2)
    
    def _agregar_fila_cedulas_mejorada(self, datos_cliente: pd.DataFrame, ancho_izq: float, ancho_der: float):
        """Agrega la fila con cédulas/NIT con formato mejorado"""
        y_inicio = self.get_y()
        x_inicio = self.get_x()
        
        cedula_cliente = str(datos_cliente.iloc[0]['CEDULA'])
        # Usar el NIT dinámico de la empresa
        nit_empresa = self.empresa_nit
        
        ancho_etiqueta_izq = ancho_izq * 0.35
        ancho_valor_izq = ancho_izq * 0.65
        ancho_etiqueta_der = ancho_der * 0.35
        ancho_valor_der = ancho_der * 0.65
        
        self.set_font('Helvetica', '', 10)
        altura_cedula = self._calcular_altura_texto(cedula_cliente, ancho_valor_izq)
        altura_nit = self._calcular_altura_texto(nit_empresa, ancho_valor_der)
        altura_total = max(altura_cedula, altura_nit, 8)
        
        self.set_xy(x_inicio, y_inicio)
        self._crear_celda_con_altura_fija('CÉDULA / NIT', ancho_etiqueta_izq, altura_total, es_etiqueta=True)
        self._crear_celda_con_altura_fija(cedula_cliente, ancho_valor_izq, altura_total, es_etiqueta=False)
        
        x_derecha = x_inicio + ancho_izq + 5
        self.set_xy(x_derecha, y_inicio)
        self._crear_celda_con_altura_fija('NIT EMPRESA', ancho_etiqueta_der, altura_total, es_etiqueta=True)
        self._crear_celda_con_altura_fija(nit_empresa, ancho_valor_der, altura_total, es_etiqueta=False)
        
        self.set_xy(x_inicio, y_inicio + altura_total + 2)
    
    def _agregar_fila_direcciones_mejorada(self, ancho_izq: float, ancho_der: float, info_adicional: Dict[str, Any]):
        """Agrega la fila con direcciones con formato mejorado"""
        y_inicio = self.get_y()
        x_inicio = self.get_x()
        
        direccion_cliente = info_adicional.get('direccion', 'Vda cimarronas Km 1 Rionegro-Marinilla')
        # Usar la dirección dinámica de la empresa
        direccion_empresa = self.empresa_direccion
        
        ancho_etiqueta_izq = ancho_izq * 0.35
        ancho_valor_izq = ancho_izq * 0.65
        ancho_etiqueta_der = ancho_der * 0.35
        ancho_valor_der = ancho_der * 0.65
        
        self.set_font('Helvetica', '', 10)
        altura_dir_cliente = self._calcular_altura_texto(direccion_cliente, ancho_valor_izq)
        altura_dir_empresa = self._calcular_altura_texto(direccion_empresa, ancho_valor_der)
        altura_total = max(altura_dir_cliente, altura_dir_empresa, 8)
        
        self.set_xy(x_inicio, y_inicio)
        self._crear_celda_con_altura_fija('DIRECCIÓN', ancho_etiqueta_izq, altura_total, es_etiqueta=True)
        self._crear_celda_con_altura_fija(direccion_cliente, ancho_valor_izq, altura_total, es_etiqueta=False)
        
        x_derecha = x_inicio + ancho_izq + 5
        self.set_xy(x_derecha, y_inicio)
        self._crear_celda_con_altura_fija('DIRECCIÓN EMPRESA', ancho_etiqueta_der, altura_total, es_etiqueta=True)
        self._crear_celda_con_altura_fija(direccion_empresa, ancho_valor_der, altura_total, es_etiqueta=False)
        
        self.set_xy(x_inicio, y_inicio + altura_total + 2)
    
    def _agregar_fila_municipio_telefono(self, ancho_izq: float, ancho_der: float, info_adicional: Dict[str, Any]):
        """Agrega la fila con municipio y Celular si están disponibles"""
        y_inicio = self.get_y()
        x_inicio = self.get_x()
        
        municipio = info_adicional.get('municipio', '')
        telefono = info_adicional.get('telefono', '')
        
        if not municipio and not telefono:
            return
        
        ancho_etiqueta_izq = ancho_izq * 0.35
        ancho_valor_izq = ancho_izq * 0.65
        ancho_etiqueta_der = ancho_der * 0.35
        ancho_valor_der = ancho_der * 0.65
        
        self.set_font('Helvetica', '', 10)
        altura_municipio = self._calcular_altura_texto(municipio, ancho_valor_izq) if municipio else 8
        altura_telefono = self._calcular_altura_texto(telefono, ancho_valor_der) if telefono else 8
        altura_total = max(altura_municipio, altura_telefono, 8)
        
        if municipio:
            self.set_xy(x_inicio, y_inicio)
            self._crear_celda_con_altura_fija('MUNICIPIO', ancho_etiqueta_izq, altura_total, es_etiqueta=True)
            self._crear_celda_con_altura_fija(municipio, ancho_valor_izq, altura_total, es_etiqueta=False)
        
        if telefono:
            x_derecha = x_inicio + ancho_izq + 5
            self.set_xy(x_derecha, y_inicio)
            self._crear_celda_con_altura_fija('CELULAR', ancho_etiqueta_der, altura_total, es_etiqueta=True)
            self._crear_celda_con_altura_fija(telefono, ancho_valor_der, altura_total, es_etiqueta=False)
        
        self.set_xy(x_inicio, y_inicio + altura_total + 6)
    
    def _calcular_altura_texto(self, texto: str, ancho_disponible: float) -> float:
        """Calcula la altura necesaria para un texto en el ancho disponible"""
        if not texto:
            return ConfiguracionReporte.ALTURA_LINEA
            
        ancho_texto = self.get_string_width(texto)
        if ancho_texto <= ancho_disponible:
            return ConfiguracionReporte.ALTURA_LINEA
        else:
            num_lineas = math.ceil(ancho_texto / ancho_disponible)
            return num_lineas * ConfiguracionReporte.ALTURA_LINEA
    
    def _crear_celda_con_altura_fija(self, texto: str, ancho: float, altura_fija: float, es_etiqueta: bool = False):
        """Crea una celda con altura fija específica, centrado verticalmente el contenido"""
        x_actual = self.get_x()
        y_actual = self.get_y()
        
        if es_etiqueta:
            self.set_font('Helvetica', 'B', 10)
            self.set_fill_color(*ConfiguracionReporte.COLOR_ENCABEZADO)
            alineacion = 'C'
        else:
            self.set_font('Helvetica', '', 10)
            self.set_fill_color(255, 255, 255)
            alineacion = 'L'
        
        self.rect(x_actual, y_actual, ancho, altura_fija, 'DF')
        self.rect(x_actual, y_actual, ancho, altura_fija, 'D')
        
        ancho_texto = self.get_string_width(texto)
        if ancho_texto <= ancho:
            y_texto = y_actual + (altura_fija - ConfiguracionReporte.ALTURA_LINEA) / 2
            self.set_xy(x_actual, y_texto)
            self.cell(ancho, ConfiguracionReporte.ALTURA_LINEA, texto, border=0, align=alineacion)
        else:
            num_lineas_necesarias = math.ceil(ancho_texto / ancho)
            altura_texto_total = num_lineas_necesarias * ConfiguracionReporte.ALTURA_LINEA
            
            y_texto = y_actual + (altura_fija - altura_texto_total) / 2
            self.set_xy(x_actual, y_texto)
            
            self.multi_cell(
                ancho, ConfiguracionReporte.ALTURA_LINEA, texto,
                border=0, align=alineacion,
                new_x=XPos.RIGHT, new_y=YPos.TOP
            )
        
        self.set_xy(x_actual + ancho, y_actual)
    
    def _agregar_titulo_detalle(self):
        """Agrega el título de la sección de detalle"""
        self.set_font('Helvetica', 'B', 14)
        self.cell(
            0, 10, "",
            align="C", 
            new_x=XPos.LMARGIN, 
            new_y=YPos.NEXT
        )
        self.ln(2)
    
    # 1. MODIFICAR la función obtener_textos_fila
    def obtener_textos_fila(self, fila: pd.Series) -> List[str]:
        """Convierte una fila de datos en textos formateados para la tabla"""
        fecha_formateada = UtilFormato.formatear_fecha(fila.get('FCHA INGRESO', ''))
        porcentaje_exp = UtilFormato.formatear_porcentaje(
            fila.get('KG. EXP', 0),
            fila.get('KILOS RECIBIDOS', 1)
        )
        
        # CALCULAR EL VALOR TOTAL: TOTAL BRUTO - RETE FUENTE - FONDO HORTIFRU
        total_bruto = fila.get('TOTAL BRUTO', 0)
        rete_fuente = fila.get('RETE FUENTE', 0)
        fondo_hortifru = fila.get('FONDO HORTIFRU', 0)
        valor_total_calculado = total_bruto - rete_fuente - fondo_hortifru
    
        return [
            str(fila.get('FRUTA', '')),
            fecha_formateada,
            UtilFormato.formatear_numero(fila.get('KILOS RECIBIDOS', 0)),
            porcentaje_exp,
            UtilFormato.formatear_numero(fila.get('KG. EXP', 0)),
            UtilFormato.formatear_numero(fila.get('KG. NAL', 0)),
            UtilFormato.formatear_numero(fila.get('KG. AVE', 0)),
            UtilFormato.formatear_moneda(fila.get('PRECIO EXP', 0)),
            UtilFormato.formatear_moneda(fila.get('PRECIO NAL', 0)),
            UtilFormato.formatear_moneda(fila.get('PRECIO AVE', 0)),
            UtilFormato.formatear_moneda(fila.get('TOTAL BRUTO', 0)),
            UtilFormato.formatear_moneda(fila.get('RETE FUENTE', 0)),
            UtilFormato.formatear_moneda(fila.get('FONDO HORTIFRU', 0)),
            UtilFormato.formatear_moneda(valor_total_calculado),  # VALOR CALCULADO
        ]
    
    def agregar_tabla_detalle(self, datos: pd.DataFrame):
        """Agrega la tabla con el detalle de compras"""
        self._configurar_tabla()
        anchos_columnas = self._calcular_anchos_columnas()
        
        self._agregar_encabezados_tabla(anchos_columnas)
        self._agregar_filas_tabla(datos, anchos_columnas)
        
    
    def _configurar_tabla(self):
        """Configura los colores y fuente para la tabla"""
        self.set_font('Helvetica', 'B', 8)
        self.set_fill_color(*ConfiguracionReporte.COLOR_ENCABEZADO)
        self.set_text_color(0)
        self.set_draw_color(*ConfiguracionReporte.COLOR_BORDE)
    
    def _calcular_anchos_columnas(self) -> List[float]:
        """Calcula los anchos de las columnas basado en proporciones"""
        ancho_pagina = self.w - self.l_margin - self.r_margin
        return [prop * ancho_pagina for prop in ConfiguracionReporte.PROPORCIONES_COLUMNAS]
    
    def _agregar_encabezados_tabla(self, anchos_columnas: List[float]):
        """Agrega los encabezados de la tabla"""
        y_inicio = self.get_y()
        altura_maxima = self._calcular_altura_encabezados()
        
        for i, encabezado in enumerate(ConfiguracionReporte.ENCABEZADOS_TABLA):
            x_actual = self.get_x()
            altura_celda = altura_maxima / len(encabezado.split('\n'))
            
            self.multi_cell(
                anchos_columnas[i], altura_celda, encabezado,
                border=1, align='C', fill=True,
                new_x=XPos.RIGHT, new_y=YPos.TOP
            )
            self.set_xy(x_actual + anchos_columnas[i], y_inicio)
        
        self.ln(altura_maxima)
    
    def _calcular_altura_encabezados(self) -> float:
        """Calcula la altura máxima necesaria para los encabezados"""
        altura_maxima = 0
        for encabezado in ConfiguracionReporte.ENCABEZADOS_TABLA:
            num_lineas = len(encabezado.split('\n'))
            altura = num_lineas * ConfiguracionReporte.ALTURA_LINEA
            if altura > altura_maxima:
                altura_maxima = altura
        return altura_maxima
    
    def _agregar_filas_tabla(self, datos: pd.DataFrame, anchos_columnas: List[float]):
        """Agrega las filas de datos a la tabla"""
        self.set_font('Helvetica', '', 7)
        alternar_color = False
        
        for _, fila in datos.iterrows():
            self._agregar_fila_individual(fila, anchos_columnas, alternar_color)
            alternar_color = not alternar_color
         # 💡 MODIFICACIÓN: Agregar la fila de totales aquí, al final de la tabla de detalles.
        self._agregar_fila_total(datos, anchos_columnas)
        
    
    # 2. MODIFICAR la función _agregar_fila_total
    def _agregar_fila_total(self, datos: pd.DataFrame, anchos_columnas: List[float]):
        """
        Agrega una fila con los totales de las columnas calculando el VALOR TOTAL.
        """
        # Sumar las columnas necesarias para obtener los totales.
        total_bruto = datos['TOTAL BRUTO'].sum()
        total_rete_fuente = datos['RETE FUENTE'].sum()
        total_fondo_hortifru = datos['FONDO HORTIFRU'].sum()
        
        # CALCULAR EL TOTAL DEL VALOR TOTAL: TOTAL BRUTO - RETENCIONES
        total_valor_total = total_bruto - total_rete_fuente - total_fondo_hortifru

        # Preparar el contenido de la fila de totales
        textos_celda = [''] * len(ConfiguracionReporte.ENCABEZADOS_TABLA)
    
        # Encontramos las posiciones de las columnas necesarias
        try:
            indice_total_bruto = ConfiguracionReporte.ENCABEZADOS_TABLA.index("VALOR\nBRUTO")
            indice_rete_fuente = ConfiguracionReporte.ENCABEZADOS_TABLA.index("RETENCIÓN\nFUENTE")
            indice_fondo_hortifru = ConfiguracionReporte.ENCABEZADOS_TABLA.index("RETENCIÓN\nFONDO")
            indice_valor_total = ConfiguracionReporte.ENCABEZADOS_TABLA.index("VALOR\nTOTAL")
        except ValueError as e:
            print(f"Error: Una de las columnas no se encontró en la configuración: {e}")
            return

        # Insertamos los textos formateados en las posiciones correctas
        textos_celda[indice_total_bruto] = UtilFormato.formatear_moneda(total_bruto)
        textos_celda[indice_rete_fuente] = UtilFormato.formatear_moneda(total_rete_fuente)
        textos_celda[indice_fondo_hortifru] = UtilFormato.formatear_moneda(total_fondo_hortifru)
        textos_celda[indice_valor_total] = UtilFormato.formatear_moneda(total_valor_total)  # VALOR CALCULADO

        # Establecer la posición y el estilo para la fila de totales
        self.set_font('Helvetica', 'B', 8)
        self.set_fill_color(230, 230, 230)
    
        # Dibujar la fila completa
        y_inicio = self.get_y()
        x_inicio = self.get_x()
    
        # Encontrar la primera columna que tiene total (la más a la izquierda)
        columnas_con_total = [indice_total_bruto, indice_rete_fuente, indice_fondo_hortifru, indice_valor_total]
        primera_columna_con_total = min(columnas_con_total)
    
        # Dibujar la celda de la etiqueta "TOTAL"
        ancho_etiqueta = sum(anchos_columnas[:primera_columna_con_total])
        self.set_xy(x_inicio, y_inicio)
        self.cell(ancho_etiqueta, ConfiguracionReporte.ALTURA_LINEA * 2, "TOTAL", border=1, align='L', fill=True)

        # Dibujar cada celda con su total correspondiente
        for i in range(len(anchos_columnas)):
            if i >= primera_columna_con_total:
                x_celda = x_inicio + sum(anchos_columnas[:i])
                self.set_xy(x_celda, y_inicio)
                texto_celda = textos_celda[i] if textos_celda[i] else ''
                self.cell(anchos_columnas[i], ConfiguracionReporte.ALTURA_LINEA * 2, texto_celda, border=1, align='R', fill=True)

        self.ln(ConfiguracionReporte.ALTURA_LINEA * 2)
    
    def _agregar_fila_individual(self, fila: pd.Series, anchos_columnas: List[float], usar_color_alternativo: bool):
        """Agrega una fila individual a la tabla"""
        y_inicio = self.get_y()
        x_inicio = self.get_x()
        
        textos_celda = self.obtener_textos_fila(fila)
        altura_fila = self._calcular_altura_fila(textos_celda, anchos_columnas)
        
        if self.get_y() + altura_fila > self.page_break_trigger:
            self.add_page()
            self._agregar_encabezados_tabla(anchos_columnas)
            y_inicio = self.get_y()
            x_inicio = self.get_x()
        
        color_fondo = (
            ConfiguracionReporte.COLOR_FILA_PAR if usar_color_alternativo 
            else ConfiguracionReporte.COLOR_FILA_IMPAR
        )
        self.set_fill_color(*color_fondo)
        
        ancho_total = sum(anchos_columnas)
        self.rect(x_inicio, y_inicio, ancho_total, altura_fila, 'DF')
        
        self._agregar_contenido_celdas(textos_celda, anchos_columnas, x_inicio, y_inicio, altura_fila)
        
        self.set_xy(x_inicio, y_inicio + altura_fila)
    
    def _calcular_altura_fila(self, textos: List[str], anchos: List[float]) -> float:
        """Calcula la altura necesaria para una fila"""
        altura_maxima = ConfiguracionReporte.ALTURA_LINEA
        
        for i, texto in enumerate(textos):
            ancho_texto = self.get_string_width(texto)
            if ancho_texto > anchos[i]:
                lineas_necesarias = math.ceil(ancho_texto / anchos[i])
                altura_celda = lineas_necesarias * ConfiguracionReporte.ALTURA_LINEA
                if altura_celda > altura_maxima:
                    altura_maxima = altura_celda
        
        return altura_maxima
    
    def _agregar_contenido_celdas(self, textos: List[str], anchos: List[float], 
                                 x_inicio: float, y_inicio: float, altura_fila: float):
        """Agrega el contenido de texto a las celdas de la fila"""
        for i, texto in enumerate(textos):
            x_celda = x_inicio + sum(anchos[:i])
            
            ancho_texto = self.get_string_width(texto)
            altura_texto = ConfiguracionReporte.ALTURA_LINEA
            if ancho_texto > anchos[i]:
                altura_texto *= math.ceil(ancho_texto / anchos[i])
            
            y_texto = y_inicio + (altura_fila - altura_texto) / 2
            
            self.set_xy(x_celda, y_texto)
            self.multi_cell(
                anchos[i], ConfiguracionReporte.ALTURA_LINEA, texto,
                border=0, align=ConfiguracionReporte.ALINEACIONES_COLUMNAS[i],
                new_x=XPos.RIGHT, new_y=YPos.TOP
            )
    
    def agregar_tabla_resumen_y_cert(self, datos: pd.DataFrame):
        """
        Agrega la tabla de resumen con totales y las certificaciones al lado,
        asegurando que el margen solo afecte a las certificaciones.
        """
        
        self.ln(6)
        self.set_font('Helvetica', 'B', 9)
        
        valores_resumen = self._calcular_valores_resumen(datos)
        
        ancho_celda = 45
        altura_celda = 8
        x_posicion_tabla = self.w - self.r_margin - ancho_celda * 2
        
        y_inicio = self.get_y() # <-- Obtener la posición 'y' inicial
        
        # 1. Dibujar las certificaciones con un margen superior
        self._dibujar_certificacion(y_inicio, x_posicion_tabla)
        
        # 2. Dibujar la tabla de resumen, restaurando la posición 'y' original
        self.set_xy(x_posicion_tabla, y_inicio) # <-- Posicionar el cursor en la 'y' original, pero en la nueva 'x'
        self._agregar_filas_resumen(valores_resumen, x_posicion_tabla, ancho_celda, altura_celda)

    def _dibujar_certificacion(self, y_inicio: float, x_limite_derecho: float):
        """
        Agrega los recuadros de certificación al lado izquierdo de la tabla de resumen.
        """
        if self.certificacion_flo or self.certificacion_gap:
            ancho_disponible = x_limite_derecho - self.l_margin - 5
            ancho_cert = 50
            altura_cert = 10
            
            self.set_font('Helvetica', 'B', 10)

            # 💡 MODIFICACIÓN: Definir el margen aquí, no es necesario pasarlo
            margin_top = 8 
            
            y_posicion_cert = y_inicio + margin_top # <-- Aplicar el margen solo a la posición 'y' de las certificaciones
            
            # Dibujar el recuadro de la certificación FLO si existe
            if self.certificacion_flo and ancho_disponible >= ancho_cert:
                self.set_xy(self.l_margin, y_posicion_cert)
                self.set_fill_color(255, 255, 255)
                self.multi_cell(
                    ancho_cert, 5, self.certificacion_flo,
                    border=0, align='C'
                )

            # Dibujar el recuadro de la certificación GLOBALG.A.P. si existe
            if self.certificacion_gap and ancho_disponible >= ancho_cert:
                x_pos_gap = self.l_margin + (ancho_cert + 5 if self.certificacion_flo else 0)
                if x_pos_gap + ancho_cert <= ancho_disponible + self.l_margin:
                    self.set_xy(x_pos_gap, y_posicion_cert) # <-- Usar la nueva posición 'y'
                    self.set_fill_color(255, 255, 255)
                    self.multi_cell(
                        ancho_cert, 5, self.certificacion_gap,
                        border=0, align='C'
                    )
    def _calcular_valores_resumen(self, datos: pd.DataFrame) -> Dict[str, float]:
        """Calcula los valores para el resumen financiero"""
        # CALCULAR EL VALOR TOTAL: TOTAL BRUTO - RETENCIONES
        total_bruto = datos['TOTAL BRUTO'].sum()
        total_rete_fuente = datos['RETE FUENTE'].sum()
        total_fondo_hortifru = datos['FONDO HORTIFRU'].sum()
        
        # EL SUBTOTAL ES LA SUMA DE TODOS LOS VALORES TOTALES CALCULADOS
        subtotal_calculado = total_bruto - total_rete_fuente - total_fondo_hortifru
        
        return {
            'subtotal': subtotal_calculado,
            'descuento_2500': datos.get('D 2500', pd.Series([0])).sum(),
            'descuento_plantas': 0,
            'otros_descuentos': datos.get('DES ANALISIS', pd.Series([0])).sum(),
            'total_documento': subtotal_calculado -
            datos.get('D 2500', pd.Series([0])).sum()- datos.get('DES ANALISIS', pd.Series([0])).sum() #0 es organizar descuento plantas
        }
        
    def _agregar_filas_resumen(self, valores: Dict[str, float], x_pos: float, 
                               ancho_celda: float, altura_celda: float):
        """Agrega una fila individual a la tabla de resumen"""
        self.set_font('Helvetica', '', 9)
        self.set_fill_color(255, 255, 255)
    
        filas_resumen = [
            ("SUBTOTAL", valores['subtotal']),
            ("DESCUENTO 2500", valores['descuento_2500']),
            ("DESCUENTO PLANTAS",valores['descuento_plantas']),
            ("OTROS DESCUENTOS",valores['otros_descuentos']) 
        ]
        
        for etiqueta, valor in filas_resumen:
            self.set_xy(x_pos, self.get_y())
            self.cell(ancho_celda, altura_celda, etiqueta, border=1, align='C', fill=True)
            self.cell(ancho_celda, altura_celda, UtilFormato.formatear_moneda(valor), border=1, align='R', new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        
        self.set_x(x_pos)
        self.set_font('Helvetica', 'B', 12)
        self.cell(ancho_celda, altura_celda, "TOTAL A PAGAR", border=1, align='C', fill=True)
        self.set_font('Helvetica', 'B', 12)
        self.cell(ancho_celda, altura_celda, UtilFormato.formatear_moneda(valores['total_documento']), border=1, align='R', fill=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT)


    def generar_reportes(archivo_excel: str, nit_empresa: Optional[str] = None, 
                        nombre_empresa: Optional[str] = None, direccion_empresa: Optional[str] = None, subtitle : Optional[str] = None):
        """
        Función principal para cargar datos y generar reportes PDF con configuración dinámica de empresa.
        
        Args:
            archivo_excel: Ruta del archivo Excel con los datos
            nit_empresa: NIT de la empresa (opcional)
            nombre_empresa: Nombre de la empresa (opcional) 
            direccion_empresa: Dirección de la empresa (opcional)
        """
        try:
            gestor_datos = GestorDatos(archivo_excel)
            df_liquidacion = gestor_datos.cargar_datos()

            if df_liquidacion.empty:
                print("❌ El DataFrame de liquidación está vacío. No se pueden generar reportes.")
                return

            clientes = df_liquidacion.groupby('CEDULA')
            total_clientes = len(clientes)
            
            for i, (cedula, datos_cliente) in enumerate(clientes, 1):
                try:
                    nombre = datos_cliente['NOMBRE'].iloc[0]
                    
                    # Obtener el tipo de certificación de la columna 'FL-GL' del DataFrame principal
                    cert_tipo_liquidacion = str(datos_cliente['FL-GL'].iloc[0]).upper().strip()
                    
                    info_adicional = gestor_datos.obtener_info_cliente(cedula)
                    
                    telefono = info_adicional.get('telefono','')
                    
                    email = info_adicional.get('email','')
                    
                    # Pasar el tipo de certificación a la función obtener_certificacion
                    certificaciones = gestor_datos.obtener_certificacion(cedula, cert_tipo_liquidacion)
                    
                    # 🔥 MODIFICACIÓN PRINCIPAL: Pasar los parámetros dinámicos al constructor
                    reporte = ReporteProveedor(
                        nit_empresa=nit_empresa,
                        nombre_empresa=nombre_empresa, 
                        direccion_empresa=direccion_empresa,
                        subtitle=subtitle
                    )
                    reporte.add_page()
                    reporte.establecer_certificacion(certificaciones)
                    reporte.agregar_informacion_cliente(datos_cliente, info_adicional)
                    reporte.agregar_tabla_detalle(datos_cliente)
                    reporte.agregar_tabla_resumen_y_cert(datos_cliente)
                            
                    if telefono: 
                        telefono_limpio = "".join(filter(str.isdigit, telefono))
                        nombre_pdf = f"{nombre}!{cedula}!{telefono_limpio}.pdf" 
                        ruta_salida = os.path.join(reporte.DIRECTORIO_SALIDA_TEL, nombre_pdf)
                        reporte.output(ruta_salida)
                        
                    elif email and email not in ['no@no.com','2@2.com', '2@2.COM']: 
                        nombre_pdf = f"{nombre}!{cedula}!{email}.pdf" 
                        ruta_salida = os.path.join(reporte.DIRECTORIO_SALIDA_EMAIL, nombre_pdf)
                        reporte.output(ruta_salida)
                        
                    else:
                        # Si no hay teléfono, usa solo el nombre y la cédula
                        nombre_pdf = f"{nombre}!{cedula}.pdf"
                        ruta_salida = os.path.join(reporte.DIRECTORIO_SALIDA, nombre_pdf)
                        reporte.output(ruta_salida)
                    
                    print(f"✅ [{i}/{total_clientes}] Reporte generado para {nombre} (Cédula: {cedula}) - Certificación: {cert_tipo_liquidacion}")

                except Exception as e:
                    print(f"❌ [{i}/{total_clientes}] Error generando reporte para {nombre} (Cédula: {cedula}): {repr(e)}")
                    
            print("\n✅ Todos los reportes han sido generados exitosamente.")

        except (FileNotFoundError, ValueError) as e:
            print(f"❌ Error crítico: {e}")

