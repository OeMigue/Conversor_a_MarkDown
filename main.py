EXTENSIONES_NO_SOPORTADAS = [".doc", ".ppt", ".rtf", ".odt", ".ods", ".odp"]
NOMBRE_APP = "ConversorMarkDown"


def obtener_ruta_config():
    """Devuelve la ruta del archivo de configuración, dentro de AppData."""
    import os

    carpeta_appdata = os.path.join(os.getenv("APPDATA"), NOMBRE_APP)
    os.makedirs(carpeta_appdata, exist_ok=True)
    return os.path.join(carpeta_appdata, "config.json")


def solicitar_carpeta_base():
    """Abre una ventana para que el usuario elija dónde crear las carpetas del bot."""
    import tkinter as tk
    from tkinter import filedialog, messagebox

    ventana = tk.Tk()
    ventana.withdraw()  # oculta la ventana principal, solo queremos el diálogo

    messagebox.showinfo(
        "Configuración inicial",
        "Hola!"
        "\n\nEste es tu conversor MarkDown de confianza"
        "\n\nComo es la primera vez que vas a ejecutar el programa, debes elige la carpeta de tu quipo donde se van a crear las carpetas de trabajo."
    )

    carpeta_elegida = filedialog.askdirectory(title="Elige la carpeta base para el Conversor MarkDown")

    ventana.destroy()

    if not carpeta_elegida:
        raise Exception("No se eligió ninguna carpeta. El programa no puede continuar.")

    return carpeta_elegida


def mostrar_aviso_carpetas_creadas(rutas):
    """Muestra una ventana que se queda abierta hasta que el usuario toque Aceptar."""
    import tkinter as tk

    ventana = tk.Tk()
    ventana.title("Configuración completada")
    ventana.geometry("420x380")
    ventana.resizable(False, False)

    mensaje = (
        "Las carpetas de trabajo se crearon correctamente:\n\n"
        f"• {rutas['originales']}\n"
        f"• {rutas['markdown']}\n"
        f"• {rutas['temporal']}\n\n"
        "¡AVISO!:\n\n"
        "Para empezar a utilizar el conversor haga lo siguiente:\n\n"
        '1. Introdúzca los archivos que quiere convertir a la carpeta llamada "Archivos Originales"\n'
        '2. Ejecute el programa ejecutable "Conversor de archivos a MarkDown"\n' 
        '3. Sus archivos MarkDown están listos para utilizarse en la carpeta "Archivos en Formato MarkDown"\n\n'
        "¡MUCHAS GRACIAS POR UTILIZAR EL PROGRAMA!"
    )

    etiqueta = tk.Label(ventana, text=mensaje, justify="left", wraplength=380, padx=10, pady=10)
    etiqueta.pack()

    boton = tk.Button(ventana, text="Aceptar", width=15, command=ventana.destroy)
    boton.pack(pady=10)

    ventana.mainloop()  # ← pausa el programa hasta que el usuario cierre con el botón


def crear_configuracion_inicial():
    """Crea las 3 carpetas de trabajo y guarda la configuración."""
    import os
    import json

    carpeta_base = solicitar_carpeta_base()

    rutas = {
        "originales": os.path.join(carpeta_base, "Archivos Originales"),
        "markdown": os.path.join(carpeta_base, "Archivos en Formato MarkDown"),
        "temporal": os.path.join(carpeta_base, "Temp_Convertidos")
    }

    for ruta in rutas.values():
        os.makedirs(ruta, exist_ok=True)

    ruta_config = obtener_ruta_config()
    with open(ruta_config, "w", encoding="utf-8") as archivo_config:
        json.dump(rutas, archivo_config, indent=4)

    print(f"Configuración creada correctamente en: {ruta_config}")

    mostrar_aviso_carpetas_creadas(rutas)

    return rutas


def cargar_configuracion():
    """Carga la configuración si ya existe, o la crea si es la primera vez."""
    import os
    import json

    ruta_config = obtener_ruta_config()

    if os.path.exists(ruta_config):
        with open(ruta_config, "r", encoding="utf-8") as archivo_config:
            rutas = json.load(archivo_config)
        return rutas, False  # False = no es primera vez

    rutas = crear_configuracion_inicial()
    return rutas, True  # True = fue la primera vez


def listar_archivos(carpeta):
    import os

    archivos = []
    try:
        for nombre in os.listdir(carpeta):
            ruta = os.path.join(carpeta, nombre)
            if os.path.isfile(ruta):
                archivos.append(ruta)

        if not archivos:
            print(f"NO HAY ARCHIVOS EN LA CARPETA {carpeta}")
            return None

        return archivos

    except Exception as e:
        print(f"ERROR al listar archivos: {e}")
        return None


def convertir_xlsb_a_xlsx(ruta_xlsb, carpeta_temporal):
    import os
    import pandas as pd

    os.makedirs(carpeta_temporal, exist_ok=True)

    nombre_original = os.path.basename(ruta_xlsb)
    nombre_sin_extension = os.path.splitext(nombre_original)[0]
    ruta_xlsx = os.path.join(carpeta_temporal, nombre_sin_extension + ".xlsx")

    hojas = pd.read_excel(ruta_xlsb, sheet_name=None, engine="pyxlsb")

    with pd.ExcelWriter(ruta_xlsx, engine="openpyxl") as writer:
        for nombre_hoja, datos in hojas.items():
            datos.to_excel(writer, sheet_name=nombre_hoja, index=False)

    return ruta_xlsx


def convertir_archivos(archivos, carpeta_destino, carpeta_temporal):
    import os
    from markitdown import MarkItDown

    md = MarkItDown()
    os.makedirs(carpeta_destino, exist_ok=True)

    for archivo in archivos:
        nombre_original = os.path.basename(archivo)
        extension = os.path.splitext(archivo)[1].lower()

        if extension in EXTENSIONES_NO_SOPORTADAS:
            print(f"AVISO: '{nombre_original}' tiene un formato ({extension}) que no se puede procesar. Se omite.")
            continue

        try:
            archivo_a_procesar = archivo

            if extension == ".xlsb":
                print(f"Convirtiendo formato: {nombre_original} (xlsb -> xlsx)")
                archivo_a_procesar = convertir_xlsb_a_xlsx(archivo, carpeta_temporal)

            nombre_sin_extension = os.path.splitext(nombre_original)[0]
            nombre_markdown = nombre_sin_extension + ".md"
            ruta_destino = os.path.join(carpeta_destino, nombre_markdown)

            resultado = md.convert(archivo_a_procesar)

            with open(ruta_destino, "w", encoding="utf-8") as archivo_a_guardar:
                archivo_a_guardar.write(resultado.text_content)

            print(f"Convertido: {nombre_original} -> {nombre_markdown}")

        except Exception as e:
            print(f"ERROR al convertir '{nombre_original}': {e}")
            continue


def main():
    rutas, es_primera_vez = cargar_configuracion()

    if es_primera_vez:
        # Ya se mostró el aviso y el usuario cerró con Aceptar. Corta acá.
        return

    archivos = listar_archivos(rutas["originales"])

    if archivos is None:
        return

    convertir_archivos(archivos, rutas["markdown"], rutas["temporal"])


if __name__ == '__main__':
    main()