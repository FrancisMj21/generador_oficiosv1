import json
import os
import re
import tkinter as tk

from pathlib import Path
from datetime import datetime

from tkinter import ttk, filedialog, messagebox

import pandas as pd
from docxtpl import DocxTemplate
from docx2pdf import convert


# ============================================================
# CONFIGURACIÓN GENERAL
# ============================================================

BASE_DIR = Path(__file__).resolve().parent

CONFIG_DIR = BASE_DIR / "config"
SALIDA_DIR = BASE_DIR / "salida"
PLANTILLAS_DIR = BASE_DIR / "plantillas"

CONTADOR_FILE = CONFIG_DIR / "contador.json"


# Crear carpetas automáticamente
CONFIG_DIR.mkdir(
    parents=True,
    exist_ok=True
)

SALIDA_DIR.mkdir(
    parents=True,
    exist_ok=True
)

PLANTILLAS_DIR.mkdir(
    parents=True,
    exist_ok=True
)


# ============================================================
# CONFIGURACIÓN DEL EXCEL
# ============================================================

# El Excel tiene:
#
# Fila 1 -> título
# Fila 2 -> información
# Fila 3 -> actualización
# Fila 4 -> encabezados
# Fila 5 -> primer registro
#
# Por lo tanto:
#
# header=3
#
# porque pandas empieza a contar desde 0.

FILA_ENCABEZADO = 3


# ============================================================
# INFORMACIÓN GENERAL
# ============================================================

ANIO_OFICIO = 2026

CODIGO_OFICIO = "GG/GOVERTECH"


# ============================================================
# FUNCIONES AUXILIARES
# ============================================================

def limpiar_valor(valor):
    """
    Convierte un valor del Excel en texto limpio.
    """

    if pd.isna(valor):
        return ""

    texto = str(valor).strip()

    # Eliminar .0 cuando Excel convierte
    # números en flotantes.
    if texto.endswith(".0"):
        texto = texto[:-2]

    return texto


def limpiar_nombre_archivo(nombre):
    """
    Elimina caracteres no permitidos
    en nombres de archivos de Windows.
    """

    return re.sub(
        r'[<>:"/\\|?*]',
        "_",
        str(nombre)
    )


def fecha_en_espanol(fecha):
    """
    Convierte:

        05/10/2026

    en:

        5 de octubre de 2026
    """

    meses = [
        "enero",
        "febrero",
        "marzo",
        "abril",
        "mayo",
        "junio",
        "julio",
        "agosto",
        "septiembre",
        "octubre",
        "noviembre",
        "diciembre"
    ]

    try:

        fecha_obj = datetime.strptime(
            fecha,
            "%d/%m/%Y"
        )

        return (
            f"{fecha_obj.day} de "
            f"{meses[fecha_obj.month - 1]} de "
            f"{fecha_obj.year}"
        )

    except ValueError:

        return fecha


# ============================================================
# CONTADOR DE OFICIOS
# ============================================================

class Contador:

    def __init__(self):

        self.archivo = CONTADOR_FILE

        self.crear_archivo()


    def crear_archivo(self):

        if not self.archivo.exists():

            self.guardar({})


    def cargar(self):

        try:

            with open(
                self.archivo,
                "r",
                encoding="utf-8"
            ) as archivo:

                return json.load(archivo)

        except (
            FileNotFoundError,
            json.JSONDecodeError
        ):

            return {}


    def guardar(self, datos):

        with open(
            self.archivo,
            "w",
            encoding="utf-8"
        ) as archivo:

            json.dump(
                datos,
                archivo,
                indent=4,
                ensure_ascii=False
            )


    def siguiente(self, tipo):

        datos = self.cargar()

        ultimo = int(
            datos.get(tipo, 0)
        )

        return ultimo + 1


    def confirmar(self, tipo, numero):

        datos = self.cargar()

        datos[tipo] = numero

        self.guardar(datos)


# ============================================================
# LEER EXCEL
# ============================================================

def obtener_hojas_excel(archivo):
    """
    Obtiene las hojas disponibles
    del archivo Excel.
    """

    extension = Path(
        archivo
    ).suffix.lower()


    if extension == ".xlsx":

        excel = pd.ExcelFile(
            archivo,
            engine="openpyxl"
        )

    elif extension == ".xls":

        excel = pd.ExcelFile(
            archivo,
            engine="xlrd"
        )

    else:

        raise ValueError(
            "Formato no compatible.\n\n"
            "Selecciona un archivo .xlsx o .xls."
        )


    return excel.sheet_names


# ============================================================
# LEER EXCEL DE MUNICIPALIDADES
# ============================================================

def leer_excel_municipalidad(
    archivo,
    hoja
):
    """
    Lee el Excel de municipalidades.

    Estructura:

        B -> Departamento
        C -> Provincia
        D -> Distrito
        E -> Cargo
        F -> Nombre
        G -> Sexo
        H -> Dirección

    Encabezados:
        fila 4

    Datos:
        fila 5 en adelante.
    """

    extension = Path(
        archivo
    ).suffix.lower()


    # --------------------------------------------------------
    # XLSX
    # --------------------------------------------------------

    if extension == ".xlsx":

        dataframe = pd.read_excel(

            archivo,

            sheet_name=hoja,

            header=FILA_ENCABEZADO,

            engine="openpyxl"
        )


    # --------------------------------------------------------
    # XLS
    # --------------------------------------------------------

    elif extension == ".xls":

        dataframe = pd.read_excel(

            archivo,

            sheet_name=hoja,

            header=FILA_ENCABEZADO,

            engine="xlrd"
        )


    else:

        raise ValueError(
            "El archivo seleccionado "
            "no es un Excel válido."
        )


    # --------------------------------------------------------
    # VALIDAR COLUMNAS
    # --------------------------------------------------------

    if len(dataframe.columns) < 8:

        raise ValueError(

            "El Excel no tiene suficientes columnas.\n\n"

            "Se esperan datos hasta la columna H."
        )


    # --------------------------------------------------------
    # TOMAR B:H
    # --------------------------------------------------------

    dataframe = dataframe.iloc[
        :,
        1:8
    ].copy()


    # --------------------------------------------------------
    # NOMBRES INTERNOS
    # --------------------------------------------------------

    dataframe.columns = [

        "departamento",

        "provincia",

        "distrito",

        "cargo",

        "nombre_alcalde",

        "sexo",

        "direccion"
    ]


    return dataframe


# ============================================================
# CREAR DATOS DE MUNICIPALIDAD
# ============================================================

def crear_datos_municipalidad(
    fila
):
    """
    Convierte una fila del Excel
    en los datos necesarios para
    generar el oficio.
    """

    departamento = limpiar_valor(
        fila["departamento"]
    )

    provincia = limpiar_valor(
        fila["provincia"]
    )

    distrito = limpiar_valor(
        fila["distrito"]
    )

    cargo_original = limpiar_valor(
        fila["cargo"]
    )

    nombre_alcalde = limpiar_valor(
        fila["nombre_alcalde"]
    )

    direccion = limpiar_valor(
        fila["direccion"]
    )


    # ========================================================
    # DETERMINAR CARGO
    # ========================================================

    cargo_minuscula = (
        cargo_original.lower()
    )


    if "provincial" in cargo_minuscula:

        alcalde = "Alcalde Provincial"

        municipalidad = (
            f"Municipalidad Provincial de "
            f"{provincia}"
        )


    elif "distrital" in cargo_minuscula:

        alcalde = "Alcalde Distrital"

        municipalidad = (
            f"Municipalidad Distrital de "
            f"{distrito}"
        )


    else:

        # Si aparece otro valor,
        # usamos exactamente lo que
        # viene del Excel.

        alcalde = cargo_original

        if provincia:

            municipalidad = (
                f"Municipalidad Provincial de "
                f"{provincia}"
            )

        else:

            municipalidad = (
                f"Municipalidad Distrital de "
                f"{distrito}"
            )


    # ========================================================
    # REGIÓN
    # ========================================================

    # Según tu estructura:
    #
    # Departamento = Región

    region = departamento


    # ========================================================
    # UBICACIÓN
    # ========================================================

    ubicacion = (
        f"{region}/"
        f"{provincia}/"
        f"{distrito}"
    )


    return {

        "alcalde":
            alcalde,

        "nombre_alcalde":
            nombre_alcalde,

        "municipalidad":
            municipalidad,

        "direccion":
            direccion,

        "region":
            region,

        "departamento":
            departamento,

        "provincia":
            provincia,

        "distrito":
            distrito,

        "ubicacion":
            ubicacion
    }


# ============================================================
# CONVERTIR WORD A PDF
# ============================================================

def convertir_a_pdf(
    archivo_word,
    carpeta_salida
):
    """
    Convierte DOCX a PDF usando
    Microsoft Word.

    Requiere Microsoft Word
    instalado en Windows.
    """

    archivo_word = Path(
        archivo_word
    )

    carpeta_salida = Path(
        carpeta_salida
    )


    carpeta_salida.mkdir(
        parents=True,
        exist_ok=True
    )


    try:

        convert(

            str(archivo_word),

            str(carpeta_salida)
        )


    except Exception as error:

        raise RuntimeError(

            "No se pudo convertir "
            "el documento a PDF.\n\n"

            "Verifica que Microsoft Word "
            "esté instalado.\n\n"

            f"Detalle:\n{error}"
        )


    archivo_pdf = (
        carpeta_salida /
        f"{archivo_word.stem}.pdf"
    )


    if not archivo_pdf.exists():

        raise RuntimeError(

            "Microsoft Word no generó "
            "el archivo PDF."
        )


    return archivo_pdf


# ============================================================
# GENERAR OFICIO DE MUNICIPALIDAD
# ============================================================

def generar_oficio_municipalidad(

    fila,

    numero,

    fecha,

    plantilla,

    tipo_salida

):

    datos = crear_datos_municipalidad(
        fila
    )


    # ========================================================
    # CONTEXTO DE LA PLANTILLA
    # ========================================================

    contexto = {

        # Número
        "numero":
            f"{numero:03d}",

        # Fecha
        "fecha":
            fecha,

        # Cargo
        "alcalde":
            datos["alcalde"],

        # Nombre
        "nombre_alcalde":
            datos["nombre_alcalde"],

        # Municipalidad
        "municipalidad":
            datos["municipalidad"],

        # Dirección
        "direccion":
            datos["direccion"],

        # Región
        "region":
            datos["region"],

        # Departamento
        "departamento":
            datos["departamento"],

        # Provincia
        "provincia":
            datos["provincia"],

        # Distrito
        "distrito":
            datos["distrito"],

        # Ubicación completa
        "ubicacion":
            datos["ubicacion"],

        # Código
        "codigo_oficio":
            CODIGO_OFICIO
    }


    # ========================================================
    # CARGAR PLANTILLA
    # ========================================================

    documento = DocxTemplate(
        plantilla
    )


    # ========================================================
    # REEMPLAZAR VARIABLES
    # ========================================================

    documento.render(
        contexto
    )


    # ========================================================
    # IDENTIFICADOR
    # ========================================================

    identificador = (

        datos["municipalidad"]

        or datos["distrito"]

        or datos["provincia"]

        or f"OFICIO_{numero}"
    )


    identificador = limpiar_nombre_archivo(
        identificador
    )


    # ========================================================
    # NOMBRE BASE
    # ========================================================

    nombre_base = (

        f"OFICIO_N_"

        f"{numero:03d}"

        f"-{ANIO_OFICIO}"

        f"-GG-GOVERTECH-"

        f"{identificador}"
    )


    # ========================================================
    # WORD
    # ========================================================

    if tipo_salida == "Word (.docx)":

        ruta_word = (

            SALIDA_DIR /

            f"{nombre_base}.docx"
        )


        documento.save(
            ruta_word
        )


        return ruta_word


    # ========================================================
    # PDF
    # ========================================================

    if tipo_salida == "PDF (.pdf)":

        # -----------------------------------------------
        # Carpeta temporal
        # -----------------------------------------------

        carpeta_temporal = (
            SALIDA_DIR /
            "_temp"
        )


        carpeta_temporal.mkdir(
            parents=True,
            exist_ok=True
        )


        # -----------------------------------------------
        # Word temporal
        # -----------------------------------------------

        ruta_word_temporal = (

            carpeta_temporal /

            f"{nombre_base}.docx"
        )


        documento.save(
            ruta_word_temporal
        )


        # -----------------------------------------------
        # Convertir
        # -----------------------------------------------

        ruta_pdf = convertir_a_pdf(

            ruta_word_temporal,

            SALIDA_DIR
        )


        # -----------------------------------------------
        # Eliminar Word temporal
        # -----------------------------------------------

        try:

            ruta_word_temporal.unlink()

        except OSError:

            pass


        return ruta_pdf


    # ========================================================
    # ERROR
    # ========================================================

    raise ValueError(
        "Tipo de salida no válido."
    )


# ============================================================
# PROCESAR MUNICIPALIDADES
# ============================================================

def procesar_municipalidades(

    archivo,

    hoja,

    plantilla,

    fecha,

    tipo_salida,

    callback_progreso=None

):

    # ========================================================
    # LEER EXCEL
    # ========================================================

    dataframe = leer_excel_municipalidad(

        archivo,

        hoja
    )


    # ========================================================
    # CONTADOR
    # ========================================================

    contador = Contador()


    # ========================================================
    # TOTAL DE FILAS
    # ========================================================

    total_filas = len(
        dataframe
    )


    procesados = 0

    documentos = []


    # ========================================================
    # RECORRER EXCEL
    # ========================================================

    for _, fila in dataframe.iterrows():

        datos = crear_datos_municipalidad(
            fila
        )


        # ====================================================
        # IGNORAR FILAS SIN NOMBRE
        # ====================================================

        if not datos["nombre_alcalde"]:

            continue


        # ====================================================
        # OBTENER SIGUIENTE NÚMERO
        # ====================================================

        numero = contador.siguiente(
            "municipalidad"
        )


        # ====================================================
        # GENERAR
        # ====================================================

        ruta = generar_oficio_municipalidad(

            fila=fila,

            numero=numero,

            fecha=fecha,

            plantilla=plantilla,

            tipo_salida=tipo_salida
        )


        # ====================================================
        # CONFIRMAR NÚMERO
        # ====================================================

        contador.confirmar(

            "municipalidad",

            numero
        )


        documentos.append(
            ruta
        )


        # ====================================================
        # ACTUALIZAR PROGRESO
        # ====================================================

        procesados += 1


        if callback_progreso:

            porcentaje = (

                procesados /

                total_filas *

                100

                if total_filas > 0

                else 100
            )


            callback_progreso(

                porcentaje,

                procesados,

                total_filas,

                datos["municipalidad"]
            )


    return documentos


# ============================================================
# INTERFAZ
# ============================================================

class Aplicacion:

    def __init__(self, root):

        self.root = root


        # ====================================================
        # VENTANA
        # ====================================================

        self.root.title(
            "Generador de Oficios - GOVERTECH"
        )


        self.root.geometry(
            "820x650"
        )


        self.root.resizable(
            False,
            False
        )


        # ====================================================
        # VARIABLES
        # ====================================================

        self.tipo_oficio = tk.StringVar(

            value="Municipalidad"
        )


        self.tipo_salida = tk.StringVar(

            value="Word (.docx)"
        )


        self.archivo_excel = tk.StringVar()


        self.hoja = tk.StringVar()


        self.plantilla = tk.StringVar()


        self.fecha = tk.StringVar(

            value=datetime.now().strftime(
                "%d/%m/%Y"
            )
        )


        self.estado = tk.StringVar(

            value="Listo para generar"
        )


        self.porcentaje = tk.DoubleVar(

            value=0
        )


        # ====================================================
        # CREAR INTERFAZ
        # ====================================================

        self.crear_interfaz()


    # ========================================================
    # CREAR INTERFAZ
    # ========================================================

    def crear_interfaz(self):

        # ====================================================
        # TÍTULO
        # ====================================================

        titulo = ttk.Label(

            self.root,

            text="GENERADOR DE OFICIOS",

            font=(
                "Segoe UI",
                22,
                "bold"
            )
        )


        titulo.pack(
            pady=(30, 5)
        )


        subtitulo = ttk.Label(

            self.root,

            text="GOVERTECH",

            font=(
                "Segoe UI",
                11
            )
        )


        subtitulo.pack(
            pady=(0, 20)
        )


        # ====================================================
        # CONTENEDOR
        # ====================================================

        frame = ttk.Frame(

            self.root,

            padding=20
        )


        frame.pack(

            fill="x",

            padx=35
        )


        # ====================================================
        # TIPO DE OFICIO
        # ====================================================

        ttk.Label(

            frame,

            text="Tipo de oficio:"
        ).grid(

            row=0,

            column=0,

            sticky="w",

            pady=10
        )


        self.combo_tipo_oficio = ttk.Combobox(

            frame,

            textvariable=self.tipo_oficio,

            values=[

                "Municipalidad",

                "Gobierno Regional"
            ],

            state="readonly",

            width=48
        )


        self.combo_tipo_oficio.grid(

            row=0,

            column=1,

            padx=10,

            pady=10
        )


        # ====================================================
        # TIPO DE SALIDA
        # ====================================================

        ttk.Label(

            frame,

            text="Tipo de salida:"
        ).grid(

            row=1,

            column=0,

            sticky="w",

            pady=10
        )


        ttk.Combobox(

            frame,

            textvariable=self.tipo_salida,

            values=[

                "Word (.docx)",

                "PDF (.pdf)"
            ],

            state="readonly",

            width=48
        ).grid(

            row=1,

            column=1,

            padx=10,

            pady=10
        )


        # ====================================================
        # EXCEL
        # ====================================================

        ttk.Label(

            frame,

            text="Archivo Excel:"
        ).grid(

            row=2,

            column=0,

            sticky="w",

            pady=10
        )


        ttk.Entry(

            frame,

            textvariable=self.archivo_excel,

            width=50
        ).grid(

            row=2,

            column=1,

            padx=10
        )


        ttk.Button(

            frame,

            text="Seleccionar",

            command=self.seleccionar_excel

        ).grid(

            row=2,

            column=2
        )


        # ====================================================
        # HOJA
        # ====================================================

        ttk.Label(

            frame,

            text="Hoja:"
        ).grid(

            row=3,

            column=0,

            sticky="w",

            pady=10
        )


        self.combo_hoja = ttk.Combobox(

            frame,

            textvariable=self.hoja,

            state="readonly",

            width=48
        )


        self.combo_hoja.grid(

            row=3,

            column=1,

            padx=10
        )


        # ====================================================
        # PLANTILLA
        # ====================================================

        ttk.Label(

            frame,

            text="Plantilla Word:"
        ).grid(

            row=4,

            column=0,

            sticky="w",

            pady=10
        )


        ttk.Entry(

            frame,

            textvariable=self.plantilla,

            width=50
        ).grid(

            row=4,

            column=1,

            padx=10
        )


        ttk.Button(

            frame,

            text="Seleccionar",

            command=self.seleccionar_plantilla

        ).grid(

            row=4,

            column=2
        )


        # ====================================================
        # FECHA
        # ====================================================

        ttk.Label(

            frame,

            text="Fecha:"
        ).grid(

            row=5,

            column=0,

            sticky="w",

            pady=10
        )


        ttk.Entry(

            frame,

            textvariable=self.fecha,

            width=20
        ).grid(

            row=5,

            column=1,

            sticky="w",

            padx=10
        )


        # ====================================================
        # SEPARADOR
        # ====================================================

        ttk.Separator(

            self.root,

            orient="horizontal"

        ).pack(

            fill="x",

            padx=45,

            pady=15
        )


        # ====================================================
        # ESTADO
        # ====================================================

        ttk.Label(

            self.root,

            textvariable=self.estado,

            font=(
                "Segoe UI",
                10
            )

        ).pack(

            pady=5
        )


        # ====================================================
        # BARRA DE PROGRESO
        # ====================================================

        self.barra_progreso = ttk.Progressbar(

            self.root,

            orient="horizontal",

            length=650,

            maximum=100,

            variable=self.porcentaje
        )


        self.barra_progreso.pack(

            pady=8
        )


        # ====================================================
        # PORCENTAJE
        # ====================================================

        self.label_porcentaje = ttk.Label(

            self.root,

            text="0%",

            font=(
                "Segoe UI",
                18,
                "bold"
            )
        )


        self.label_porcentaje.pack(

            pady=5
        )


        # ====================================================
        # BOTÓN GENERAR
        # ====================================================

        self.boton_generar = ttk.Button(

            self.root,

            text="GENERAR OFICIOS",

            command=self.generar
        )


        self.boton_generar.pack(

            pady=15,

            ipadx=35,

            ipady=8
        )


        # ====================================================
        # ABRIR SALIDA
        # ====================================================

        ttk.Button(

            self.root,

            text="Abrir carpeta de salida",

            command=self.abrir_salida

        ).pack(

            pady=5
        )


    # ========================================================
    # SELECCIONAR EXCEL
    # ========================================================

    def seleccionar_excel(self):

        archivo = filedialog.askopenfilename(

            title="Seleccionar archivo Excel",

            filetypes=[

                (
                    "Archivos Excel",
                    "*.xlsx *.xls"
                ),

                (
                    "Todos los archivos",
                    "*.*"
                )
            ]
        )


        if not archivo:

            return


        self.archivo_excel.set(
            archivo
        )


        # ====================================================
        # OBTENER HOJAS
        # ====================================================

        try:

            hojas = obtener_hojas_excel(
                archivo
            )


            self.combo_hoja["values"] = hojas


            if hojas:

                self.hoja.set(
                    hojas[0]
                )


        except Exception as error:

            messagebox.showerror(

                "Error",

                (
                    "No se pudieron "
                    "leer las hojas del Excel.\n\n"

                    f"{error}"
                )
            )


    # ========================================================
    # SELECCIONAR PLANTILLA
    # ========================================================

    def seleccionar_plantilla(self):

        archivo = filedialog.askopenfilename(

            title="Seleccionar plantilla Word",

            initialdir=PLANTILLAS_DIR,

            filetypes=[

                (
                    "Documento Word",
                    "*.docx"
                )
            ]
        )


        if archivo:

            self.plantilla.set(
                archivo
            )


    # ========================================================
    # ACTUALIZAR PROGRESO
    # ========================================================

    def actualizar_progreso(

        self,

        porcentaje,

        procesados,

        total,

        nombre

    ):

        # ----------------------------------------------------
        # BARRA
        # ----------------------------------------------------

        self.porcentaje.set(
            porcentaje
        )


        # ----------------------------------------------------
        # PORCENTAJE
        # ----------------------------------------------------

        self.label_porcentaje.config(

            text=f"{porcentaje:.1f}%"
        )


        # ----------------------------------------------------
        # ESTADO
        # ----------------------------------------------------

        self.estado.set(

            f"Generando {procesados} "
            f"de {total} — "
            f"{nombre}"
        )


        # ----------------------------------------------------
        # ACTUALIZAR TKINTER
        # ----------------------------------------------------

        self.root.update_idletasks()


    # ========================================================
    # GENERAR
    # ========================================================

    def generar(self):

        # ====================================================
        # VALIDAR EXCEL
        # ====================================================

        if not self.archivo_excel.get():

            messagebox.showwarning(

                "Falta Excel",

                "Selecciona el archivo Excel."
            )

            return


        # ====================================================
        # VALIDAR HOJA
        # ====================================================

        if not self.hoja.get():

            messagebox.showwarning(

                "Falta hoja",

                "Selecciona la hoja del Excel."
            )

            return


        # ====================================================
        # VALIDAR PLANTILLA
        # ====================================================

        if not self.plantilla.get():

            messagebox.showwarning(

                "Falta plantilla",

                "Selecciona la plantilla Word."
            )

            return


        # ====================================================
        # VALIDAR FECHA
        # ====================================================

        try:

            datetime.strptime(

                self.fecha.get(),

                "%d/%m/%Y"
            )

        except ValueError:

            messagebox.showerror(

                "Fecha incorrecta",

                (
                    "La fecha debe tener "
                    "el formato:\n\n"

                    "DD/MM/YYYY\n\n"

                    "Ejemplo:\n"

                    "05/10/2026"
                )
            )

            return


        # ====================================================
        # VALIDAR TIPO DE OFICIO
        # ====================================================

        if self.tipo_oficio.get() == "Gobierno Regional":

            messagebox.showinfo(

                "Gobierno Regional",

                (
                    "El tipo Gobierno Regional "
                    "ya está disponible en la interfaz.\n\n"

                    "Falta definir la estructura "
                    "de las columnas de su Excel "
                    "para implementar su generación."
                )
            )

            return


        # ====================================================
        # PREPARAR FECHA
        # ====================================================

        fecha = fecha_en_espanol(

            self.fecha.get()
        )


        # ====================================================
        # REINICIAR PROGRESO
        # ====================================================

        self.porcentaje.set(
            0
        )


        self.label_porcentaje.config(

            text="0%"
        )


        self.estado.set(

            "Preparando generación..."
        )


        self.boton_generar.config(

            state="disabled"
        )


        self.root.update_idletasks()


        # ====================================================
        # GENERAR
        # ====================================================

        try:

            documentos = procesar_municipalidades(

                archivo=self.archivo_excel.get(),

                hoja=self.hoja.get(),

                plantilla=self.plantilla.get(),

                fecha=fecha,

                tipo_salida=self.tipo_salida.get(),

                callback_progreso=
                    self.actualizar_progreso
            )


            # =================================================
            # FINALIZAR
            # =================================================

            self.porcentaje.set(
                100
            )


            self.label_porcentaje.config(

                text="100%"
            )


            self.estado.set(

                f"Proceso terminado — "
                f"{len(documentos)} documentos"
            )


            # =================================================
            # MENSAJE
            # =================================================

            messagebox.showinfo(

                "Proceso terminado",

                (
                    f"Se generaron "
                    f"{len(documentos)} documentos.\n\n"

                    f"Formato: "
                    f"{self.tipo_salida.get()}\n\n"

                    f"Carpeta de salida:\n"
                    f"{SALIDA_DIR}"
                )
            )


        except Exception as error:

            self.estado.set(

                "Error durante la generación"
            )


            messagebox.showerror(

                "Error",

                (
                    "Ocurrió un error "
                    "durante la generación.\n\n"

                    f"Detalle:\n{error}"
                )
            )


        finally:

            self.boton_generar.config(

                state="normal"
            )


    # ========================================================
    # ABRIR CARPETA DE SALIDA
    # ========================================================

    def abrir_salida(self):

        SALIDA_DIR.mkdir(

            parents=True,

            exist_ok=True
        )


        os.startfile(
            SALIDA_DIR
        )


# ============================================================
# EJECUTAR APLICACIÓN
# ============================================================

if __name__ == "__main__":

    root = tk.Tk()

    app = Aplicacion(
        root
    )

    root.mainloop()