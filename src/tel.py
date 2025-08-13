import os

# Ruta base: carpeta donde está el script
base_dir = os.path.dirname(os.path.abspath(__file__))

# Subir un nivel (de src a la carpeta raíz del proyecto)
proyecto_dir = os.path.dirname(base_dir)

# Ruta de la carpeta de teléfonos
directorio_tel = os.path.join(proyecto_dir, "output", "tel")
print("Leyendo desde:", directorio_tel)

# Leer PDFs
print(len(os.listdir(directorio_tel)))
for archivo in os.listdir(directorio_tel):
    if archivo.lower().endswith(".pdf") and "!" in archivo:
        base = archivo.removesuffix(".pdf")
        partes = base.split("!")

        if len(partes) == 3:
            nombre, cedula, telefono = partes
            print(f"Nombre: {nombre}")
            print(f"Cédula: {cedula}")
            print(f"Teléfono: {telefono}")
            print("-" * 30)
