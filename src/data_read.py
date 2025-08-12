"""
Generador de Reportes de Facturación PDF
========================================

Este módulo genera reportes de facturación individuales en formato PDF
a partir de datos de un archivo Excel.

Autor: Sistema de Facturación
Fecha: 2025
"""

import pandas as pd
from fpdf import FPDF, XPos, YPos
import os
import math
from typing import Dict, List, Any


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
    DIRECTORIO_SALIDA = 'output'
    RUTA_LOGO = './logo.png'
    
    # Información de la empresa
    EMPRESA_NIT = '800.176.428-6'
    EMPRESA_NOMBRE = 'COMERCIALIZADORA INTERNACIONAL CARIBBEAN EXOTICS S. A.'
    EMPRESA_DIRECCION = 'DIRECCIÓN NO DISPONIBLE'
    
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


class ReporteCliente(FPDF):
    """
    Clase principal para generar reportes de facturación en PDF
    
    Esta clase hereda de FPDF y añade funcionalidades específicas
    para generar reportes de compra con formato personalizado.
    """
    
    def __init__(self):
        """Inicializa el generador de reportes"""
        super().__init__(
            ConfiguracionReporte.ORIENTACION, 
            ConfiguracionReporte.UNIDAD, 
            ConfiguracionReporte.FORMATO
        )
        self._configurar_pdf()
        self._crear_directorio_salida()
    
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
    
    def header(self):
        """Define el encabezado de cada página"""
        self._agregar_logo()
        self._agregar_titulo_principal()
        self._agregar_subtitulo()
        self.ln(12)
    
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
        self.set_font('Helvetica', '', 12)
        self.cell(
            -20, 10, 'SEGUNDA QUINCENA DE JULIO',
            align='C', 
            new_x=XPos.LEFT, 
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
    
    def agregar_informacion_cliente(self, datos_cliente: pd.DataFrame):
        """
        Agrega la información del cliente al reporte con formato mejorado
        
        Args:
            datos_cliente: DataFrame con los datos del cliente
        """
        self._configurar_colores_info()
        
        # Configurar dimensiones de la tabla de información
        ancho_total = self.w - self.l_margin - self.r_margin
        ancho_izquierda = ancho_total * 0.5  # 50% para cliente
        ancho_derecha = ancho_total * 0.5    # 50% para comprador
        
        # Fila 1: Nombre/Razón Social y Comprador
        self._agregar_fila_nombre_comprador_mejorada(datos_cliente, ancho_izquierda, ancho_derecha)
        
        # Fila 2: Cédula/NIT y NIT empresa
        self._agregar_fila_cedulas_mejorada(datos_cliente, ancho_izquierda, ancho_derecha)
        
        # Fila 3: Direcciones
        self._agregar_fila_direcciones_mejorada(ancho_izquierda, ancho_derecha)
        
        # Título de detalle
        self._agregar_titulo_detalle()
    
    def _configurar_colores_info(self):
        """Configura los colores para la sección de información"""
        self.set_fill_color(*ConfiguracionReporte.COLOR_ENCABEZADO)
        self.set_text_color(0)
    
    def _agregar_fila_nombre_comprador_mejorada(self, datos_cliente: pd.DataFrame, ancho_izq: float, ancho_der: float):
        """Agrega la fila con nombre del cliente y comprador con formato mejorado"""
        y_inicio = self.get_y()
        x_inicio = self.get_x()
        
        # Datos a mostrar
        comprador_text = ConfiguracionReporte.EMPRESA_NOMBRE
        nombre_cliente = str(datos_cliente.iloc[0]['NOMBRE'])
        
        # Calcular anchos de valores (mismo para ambos lados para simetría)
        ancho_etiqueta_izq = ancho_izq * 0.35
        ancho_valor_izq = ancho_izq * 0.65
        ancho_etiqueta_der = ancho_der * 0.35
        ancho_valor_der = ancho_der * 0.65
        
        # Calcular altura necesaria para ambos lados con la misma fuente
        self.set_font('Helvetica', '', 10)
        altura_cliente = self._calcular_altura_texto(nombre_cliente, ancho_valor_izq)
        altura_comprador = self._calcular_altura_texto(comprador_text, ancho_valor_der)
        altura_total = max(altura_cliente, altura_comprador, 8)  # Ambas tendrán la misma altura
        
        # LADO IZQUIERDO - Cliente (usar altura total calculada)
        self.set_xy(x_inicio, y_inicio)
        self._crear_celda_con_altura_fija('NOMBRE / RAZÓN SOCIAL', ancho_etiqueta_izq, altura_total, es_etiqueta=True)
        self._crear_celda_con_altura_fija(nombre_cliente, ancho_valor_izq, altura_total, es_etiqueta=False)
        
        # LADO DERECHO - Comprador (usar la MISMA altura total)
        x_derecha = x_inicio + ancho_izq + 5
        self.set_xy(x_derecha, y_inicio)
        self._crear_celda_con_altura_fija('COMPRADOR', ancho_etiqueta_der, altura_total, es_etiqueta=True)
        self._crear_celda_con_altura_fija(comprador_text, ancho_valor_der, altura_total, es_etiqueta=False)
        
        # Posicionar para la siguiente fila
        self.set_xy(x_inicio, y_inicio + altura_total + 2)
    
    def _agregar_fila_cedulas_mejorada(self, datos_cliente: pd.DataFrame, ancho_izq: float, ancho_der: float):
        """Agrega la fila con cédulas/NIT con formato mejorado"""
        y_inicio = self.get_y()
        x_inicio = self.get_x()
        
        # Datos a mostrar
        cedula_cliente = str(datos_cliente.iloc[0]['CEDULA'])
        nit_empresa = ConfiguracionReporte.EMPRESA_NIT
        
        # Calcular anchos (simétricos)
        ancho_etiqueta_izq = ancho_izq * 0.35
        ancho_valor_izq = ancho_izq * 0.65
        ancho_etiqueta_der = ancho_der * 0.35
        ancho_valor_der = ancho_der * 0.65
        
        # Calcular altura necesaria para ambos lados
        self.set_font('Helvetica', '', 10)
        altura_cedula = self._calcular_altura_texto(cedula_cliente, ancho_valor_izq)
        altura_nit = self._calcular_altura_texto(nit_empresa, ancho_valor_der)
        altura_total = max(altura_cedula, altura_nit, 8)  # Misma altura para ambos
        
        # LADO IZQUIERDO - Cédula del cliente
        self.set_xy(x_inicio, y_inicio)
        self._crear_celda_con_altura_fija('CÉDULA / NIT', ancho_etiqueta_izq, altura_total, es_etiqueta=True)
        self._crear_celda_con_altura_fija(cedula_cliente, ancho_valor_izq, altura_total, es_etiqueta=False)
        
        # LADO DERECHO - NIT de la empresa
        x_derecha = x_inicio + ancho_izq + 5
        self.set_xy(x_derecha, y_inicio)
        self._crear_celda_con_altura_fija('NIT EMPRESA', ancho_etiqueta_der, altura_total, es_etiqueta=True)
        self._crear_celda_con_altura_fija(nit_empresa, ancho_valor_der, altura_total, es_etiqueta=False)
        
        # Posicionar para la siguiente fila
        self.set_xy(x_inicio, y_inicio + altura_total + 2)
    
    def _agregar_fila_direcciones_mejorada(self, ancho_izq: float, ancho_der: float):
        """Agrega la fila con direcciones con formato mejorado"""
        y_inicio = self.get_y()
        x_inicio = self.get_x()
        
        # Direcciones
        direccion_cliente = 'Vda cimarronas Km 1 Rionegro-Marinilla'
        direccion_empresa = ConfiguracionReporte.EMPRESA_DIRECCION
        
        # Calcular anchos (simétricos)
        ancho_etiqueta_izq = ancho_izq * 0.35
        ancho_valor_izq = ancho_izq * 0.65
        ancho_etiqueta_der = ancho_der * 0.35
        ancho_valor_der = ancho_der * 0.65
        
        # Calcular altura necesaria para ambos lados
        self.set_font('Helvetica', '', 10)
        altura_dir_cliente = self._calcular_altura_texto(direccion_cliente, ancho_valor_izq)
        altura_dir_empresa = self._calcular_altura_texto(direccion_empresa, ancho_valor_der)
        altura_total = max(altura_dir_cliente, altura_dir_empresa, 8)  # Misma altura para ambos
        
        # LADO IZQUIERDO - Dirección del cliente
        self.set_xy(x_inicio, y_inicio)
        self._crear_celda_con_altura_fija('DIRECCIÓN', ancho_etiqueta_izq, altura_total, es_etiqueta=True)
        self._crear_celda_con_altura_fija(direccion_cliente, ancho_valor_izq, altura_total, es_etiqueta=False)
        
        # LADO DERECHO - Dirección de la empresa
        x_derecha = x_inicio + ancho_izq + 5
        self.set_xy(x_derecha, y_inicio)
        self._crear_celda_con_altura_fija('DIRECCIÓN EMPRESA', ancho_etiqueta_der, altura_total, es_etiqueta=True)
        self._crear_celda_con_altura_fija(direccion_empresa, ancho_valor_der, altura_total, es_etiqueta=False)
        
        # Espaciado final
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
        """
        Crea una celda con altura fija específica, centrado verticalmente el contenido
        
        Args:
            texto: Texto a mostrar
            ancho: Ancho de la celda
            altura_fija: Altura fija que debe tener la celda (ya calculada)
            es_etiqueta: Si es True, aplica formato de etiqueta, sino de valor
        """
        x_actual = self.get_x()
        y_actual = self.get_y()
        
        # Configurar formato según el tipo
        if es_etiqueta:
            self.set_font('Helvetica', 'B', 10)
            self.set_fill_color(*ConfiguracionReporte.COLOR_ENCABEZADO)
            alineacion = 'C'
        else:
            self.set_font('Helvetica', '', 10)
            self.set_fill_color(255, 255, 255)  # Fondo blanco
            alineacion = 'L'
        
        # Dibujar el rectángulo de fondo con la altura fija
        self.rect(x_actual, y_actual, ancho, altura_fija, 'DF')
        
        # Dibujar el borde
        self.rect(x_actual, y_actual, ancho, altura_fija, 'D')
        
        # Calcular cuántas líneas necesita realmente el texto
        ancho_texto = self.get_string_width(texto)
        if ancho_texto <= ancho:
            # Texto cabe en una línea - centrar verticalmente
            y_texto = y_actual + (altura_fija - ConfiguracionReporte.ALTURA_LINEA) / 2
            self.set_xy(x_actual, y_texto)
            self.cell(ancho, ConfiguracionReporte.ALTURA_LINEA, texto, border=0, align=alineacion)
        else:
            # Texto necesita múltiples líneas
            num_lineas_necesarias = math.ceil(ancho_texto / ancho)
            altura_texto_total = num_lineas_necesarias * ConfiguracionReporte.ALTURA_LINEA
            
            # Centrar verticalmente el bloque de texto
            y_texto = y_actual + (altura_fija - altura_texto_total) / 2
            self.set_xy(x_actual, y_texto)
            
            # Usar multi_cell para el texto
            self.multi_cell(
                ancho, ConfiguracionReporte.ALTURA_LINEA, texto,
                border=0, align=alineacion,
                new_x=XPos.RIGHT, new_y=YPos.TOP
            )
        
        # Posicionar para la siguiente celda (a la derecha)
        self.set_xy(x_actual + ancho, y_actual)
    
    def _crear_celda_etiqueta(self, texto: str, ancho: float, altura: float):
        """Crea una celda con formato de etiqueta"""
        self.set_font('Helvetica', 'B', 11)
        self.cell(
            ancho, altura, texto,
            border=1, 
            new_x=XPos.RIGHT, 
            new_y=YPos.TOP, 
            fill=True
        )
    
    def _crear_celda_valor(self, texto: str, ancho: float, altura: float, nueva_linea: bool = False):
        """Crea una celda con formato de valor"""
        self.set_font('Helvetica', '', 11)
        new_x = XPos.LMARGIN if nueva_linea else XPos.RIGHT
        new_y = YPos.NEXT if nueva_linea else YPos.TOP
        self.cell(
            ancho, altura, texto,
            border=1, 
            new_x=new_x, 
            new_y=new_y
        )
    
    def _agregar_titulo_detalle(self):
        """Agrega el título de la sección de detalle"""
        self.set_font('Helvetica', 'B', 14)
        self.cell(
            0, 10, "DETALLE DE COMPRA",
            align="C", 
            new_x=XPos.LMARGIN, 
            new_y=YPos.NEXT
        )
        self.ln(2)
    
    def obtener_textos_fila(self, fila: pd.Series) -> List[str]:
        """
        Convierte una fila de datos en textos formateados para la tabla
        
        Args:
            fila: Serie de pandas con los datos de la fila
            
        Returns:
            Lista de strings formateados para cada columna
        """
        fecha_formateada = UtilFormato.formatear_fecha(fila.get('FCHA INGRESO', ''))
        porcentaje_exp = UtilFormato.formatear_porcentaje(
            fila.get('KG. EXP', 0), 
            fila.get('KILOS RECIBIDOS', 1)
        )
        
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
            UtilFormato.formatear_moneda(fila.get('VALOR BRUTO', 0)),
            UtilFormato.formatear_moneda(fila.get('RETENCIÓN FUENTE', 0)),
            UtilFormato.formatear_moneda(fila.get('RETENCIÓN FONDO', 0)),
            UtilFormato.formatear_moneda(fila.get('VALOR TOTAL', 0)),
        ]
    
    def agregar_tabla_detalle(self, datos: pd.DataFrame):
        """
        Agrega la tabla con el detalle de compras
        
        Args:
            datos: DataFrame con los datos a mostrar en la tabla
        """
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
    
    def _agregar_fila_individual(self, fila: pd.Series, anchos_columnas: List[float], usar_color_alternativo: bool):
        """Agrega una fila individual a la tabla"""
        y_inicio = self.get_y()
        x_inicio = self.get_x()
        
        # Calcular altura necesaria para la fila
        textos_celda = self.obtener_textos_fila(fila)
        altura_fila = self._calcular_altura_fila(textos_celda, anchos_columnas)
        
        # Verificar salto de página
        if self.get_y() + altura_fila > self.page_break_trigger:
            self.add_page()
            self._agregar_encabezados_tabla(anchos_columnas)
            y_inicio = self.get_y()
            x_inicio = self.get_x()
        
        # Configurar color de fondo
        color_fondo = (
            ConfiguracionReporte.COLOR_FILA_PAR if usar_color_alternativo 
            else ConfiguracionReporte.COLOR_FILA_IMPAR
        )
        self.set_fill_color(*color_fondo)
        
        # Dibujar fondo de la fila
        ancho_total = sum(anchos_columnas)
        self.rect(x_inicio, y_inicio, ancho_total, altura_fila, 'DF')
        
        # Agregar contenido de las celdas
        self._agregar_contenido_celdas(textos_celda, anchos_columnas, x_inicio, y_inicio, altura_fila)
        
        # Posicionar para la siguiente fila
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
            
            # Calcular posición vertical centrada
            ancho_texto = self.get_string_width(texto)
            altura_texto = ConfiguracionReporte.ALTURA_LINEA
            if ancho_texto > anchos[i]:
                altura_texto *= math.ceil(ancho_texto / anchos[i])
            
            y_texto = y_inicio + (altura_fila - altura_texto) / 2
            
            # Posicionar y agregar texto
            self.set_xy(x_celda, y_texto)
            self.multi_cell(
                anchos[i], ConfiguracionReporte.ALTURA_LINEA, texto,
                border=0, align=ConfiguracionReporte.ALINEACIONES_COLUMNAS[i],
                new_x=XPos.RIGHT, new_y=YPos.TOP
            )
    
    def agregar_tabla_resumen(self, datos: pd.DataFrame):
        """
        Agrega la tabla de resumen con totales
        
        Args:
            datos: DataFrame con los datos para calcular el resumen
        """
        self.ln(4)
        self.set_font('Helvetica', 'B', 9)
        
        # Calcular valores del resumen
        valores_resumen = self._calcular_valores_resumen(datos)
        
        # Configurar posición de la tabla
        ancho_celda = 45
        altura_celda = 8
        x_posicion = self.w - self.r_margin - ancho_celda * 2
        
        # Agregar filas del resumen
        self._agregar_filas_resumen(valores_resumen, x_posicion, ancho_celda, altura_celda)
    
    def _calcular_valores_resumen(self, datos: pd.DataFrame) -> Dict[str, float]:
        """Calcula los valores para el resumen financiero"""
        return {
            'subtotal': datos['TOTAL'].sum(),
            'total_documento': datos['TOTAL T'].sum(),
            'descuento_2500': datos.get('D 2500', pd.Series([0])).sum(),
            'descuento_plantas': datos.get('OTROS DESCUENTOS', pd.Series([0])).sum(),
            'otros_descuentos': datos.get('DES ANALISIS', pd.Series([0])).sum()
        }
    
    def _agregar_filas_resumen(self, valores: Dict[str, float], x_pos: float, 
                              ancho_celda: float, altura_celda: float):
        """Agrega las filas de la tabla de resumen"""
        filas_resumen = [
            ("SUBTOTAL", valores['subtotal']),
            ("DESCUENTO 2500", valores['descuento_2500']),
            ("DESCUENTO PLANTAS", valores['descuento_plantas']),
            ("OTROS DESCUENTOS", valores['otros_descuentos']),
        ]
        
        # Filas regulares
        for etiqueta, valor in filas_resumen:
            self._agregar_fila_resumen(etiqueta, valor, x_pos, ancho_celda, altura_celda)
        
        # Fila de total (con formato especial)
        self.ln(2)
        self.set_x(x_pos)
        self.set_font('Helvetica', 'B', 10)
        self._agregar_fila_resumen("TOTAL DOCUMENTO", valores['total_documento'], 
                                  x_pos, ancho_celda, altura_celda, es_total=True)
        self.ln(5)
    
    def _agregar_fila_resumen(self, etiqueta: str, valor: float, x_pos: float,
                             ancho_celda: float, altura_celda: float, es_total: bool = False):
        """Agrega una fila individual al resumen"""
        if not es_total:
            self.set_x(x_pos)
        
        self.cell(
            ancho_celda, altura_celda, etiqueta,
            border=1, align='L',
            new_x=XPos.RIGHT, new_y=YPos.TOP
        )
        self.cell(
            ancho_celda, altura_celda, UtilFormato.formatear_moneda(valor),
            border=1, align='R',
            new_x=XPos.LMARGIN, new_y=YPos.NEXT
        )
    
    def generar_reportes(self, archivo_excel: str):
        """
        Genera reportes individuales para cada cliente desde un archivo Excel
        
        Args:
            archivo_excel: Ruta al archivo Excel con los datos
        """
        try:
            # Cargar datos desde Excel
            df = self._cargar_datos_excel(archivo_excel)
            
            # Limpiar y procesar datos
            df = self._limpiar_datos(df)
            
            # Generar reportes por cliente
            self._generar_reportes_por_cliente(df)
            
            print("✅ Todos los reportes han sido generados exitosamente")
            
        except Exception as e:
            print(f"❌ Error al generar reportes: {str(e)}")
            raise
    
    def _cargar_datos_excel(self, archivo_excel: str) -> pd.DataFrame:
        """Carga los datos desde el archivo Excel"""
        try:
            return pd.read_excel(archivo_excel, sheet_name='BD LIQUIDACION')
        except FileNotFoundError:
            raise FileNotFoundError(f"El archivo '{archivo_excel}' no se encontró.")
        except ValueError:
            raise ValueError(f"La hoja 'BD LIQUIDACION' no existe en '{archivo_excel}'.")
    
    def _limpiar_datos(self, df: pd.DataFrame) -> pd.DataFrame:
        """Limpia y prepara los datos para el procesamiento"""
        df['NOMBRE'] = df['NOMBRE'].fillna('Sin Nombre')
        df['CEDULA'] = df['CEDULA'].fillna('000000')
        return df
    
    def _generar_reportes_por_cliente(self, df: pd.DataFrame):
        """Genera un reporte individual para cada cliente"""
        clientes = df.groupby(['NOMBRE', 'CEDULA'])
        
        total_clientes = len(clientes)
        print(f"📊 Generando reportes para {total_clientes} clientes...")
        
        for i, ((nombre_cliente, cedula_cliente), datos_cliente) in enumerate(clientes, 1):
            try:
                # Generar nombre de archivo seguro
                nombre_archivo = self._generar_nombre_archivo(nombre_cliente, cedula_cliente)
                ruta_archivo = os.path.join(ConfiguracionReporte.DIRECTORIO_SALIDA, nombre_archivo)
                
                # Crear y configurar PDF
                pdf = ReporteCliente()
                pdf.add_page()
                
                # Agregar contenido
                pdf.agregar_informacion_cliente(datos_cliente)
                pdf.agregar_tabla_detalle(datos_cliente)
                pdf.agregar_tabla_resumen(datos_cliente)
                
                # Guardar archivo
                pdf.output(ruta_archivo)
                
                print(f"✓ [{i}/{total_clientes}] {nombre_cliente} - {nombre_archivo}")
                
            except Exception as e:
                print(f"❌ Error generando reporte para {nombre_cliente}: {str(e)}")
    
    def _generar_nombre_archivo(self, nombre: str, cedula: str) -> str:
        """Genera un nombre de archivo seguro para el PDF"""
        # Limpiar caracteres especiales del nombre
        nombre_limpio = "".join(c for c in nombre if c.isalnum() or c in (' ', '-', '_')).strip()
        nombre_limpio = nombre_limpio.replace(' ', '_')
        return f"{nombre_limpio}_{cedula}.pdf"


def main():
    """Función principal para ejecutar el generador de reportes"""
    try:
        # Configuración
        excel_path = './data/IMPRESION 2 JULIO.xlsx'
        
        print("🚀 Iniciando generación de reportes...")
        print(f"📁 Archivo de entrada: {excel_path}")
        print(f"📁 Directorio de salida: {ConfiguracionReporte.DIRECTORIO_SALIDA}")
        
        # Generar reportes
        generador = ReporteCliente()
        generador.generar_reportes(excel_path)
        
        print(f"\n✅ Proceso completado. Revise la carpeta '{ConfiguracionReporte.DIRECTORIO_SALIDA}' para ver los reportes generados.")
        
    except Exception as e:
        print(f"\n❌ Error en la ejecución: {str(e)}")
        return 1
    
    return 0


if __name__ == '__main__':
    exit(main())