from Reporte_Proveedor import ReporteProveedor
import json


# Cargar configuración
with open("config.json", "r", encoding="utf-8") as f:
    config = json.load(f)

if __name__ == '__main__':
    
    
    base_dir = config['base_dir']
    # Ruta de tu archivo Excel
    ruta_archivo_excel = base_dir + config['ruta_file']
    
    ReporteProveedor.generar_reportes(
        archivo_excel=ruta_archivo_excel,
        nit_empresa=config['nit_empresa'], 
        nombre_empresa=config['nombre_empresa'],
        direccion_empresa=config['direccion_empresa'],
        subtitle = config['nombre_documento']
    )