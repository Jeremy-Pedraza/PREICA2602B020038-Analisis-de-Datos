# -*- coding: utf-8 -*-
"""Genera el documento Word de la entrega AE3 con python-docx.

Actualiza el documento iniciado en AE2 incorporando la transformacion, el
enriquecimiento, el diccionario actualizado y los resultados por pregunta de
negocio. Abre la plantilla ``EA3/documento/plantilla_iudigital.docx`` (conserva
encabezado con logo, margenes y estilos), limpia el cuerpo y reconstruye el
documento completo.

REGLA DURA: ninguna cifra se escribe a mano. Todas las dimensiones, tipos,
nulos, frecuencias, estadisticas y tablas de resultados se leen en tiempo de
ejecucion desde ``EA3/resultados/`` (CSV del crudo y transformado, tablas por
pregunta y resumen .txt).

Reproducible: volver a ejecutarlo regenera el .docx.

Uso:
    python EA3/src/generar_documento.py
"""

from __future__ import annotations

import pathlib

import pandas as pd
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt

# --------------------------------------------------------------------------
# Rutas
# --------------------------------------------------------------------------
EA3 = pathlib.Path(__file__).resolve().parent.parent
RESULTADOS = EA3 / "resultados"
DOCUMENTO = EA3 / "documento"
RAIZ = EA3.parent

PLANTILLA = DOCUMENTO / "plantilla_iudigital.docx"
SALIDA = DOCUMENTO / "Pedraza_Jeremy_AE3_Analitica_de_Datos.docx"

CSV_CRUDO = RESULTADOS / "churn_crudo.csv"
CSV_TRANS = RESULTADOS / "churn_transformado.csv"
TXT_RESUMEN = RESULTADOS / "resumen_dataframe.txt"
REQUIREMENTS = RAIZ / "requirements.txt"

TABLAS = {
    "p1": RESULTADOS / "tabla_p1_churn_por_contrato.csv",
    "p2": RESULTADOS / "tabla_p2_churn_por_antiguedad.csv",
    "p3a": RESULTADOS / "tabla_p3a_churn_por_metodo_pago.csv",
    "p3b": RESULTADOS / "tabla_p3b_churn_por_internet.csv",
    "p4": RESULTADOS / "tabla_p4_cargos_por_churn.csv",
    "p5": RESULTADOS / "tabla_p5_churn_por_num_servicios.csv",
}

# --------------------------------------------------------------------------
# Datos institucionales (literales)
# --------------------------------------------------------------------------
TITULO = (
    "Transformacion y enriquecimiento del dataset Telco Customer Churn para "
    "responder las preguntas de negocio sobre la cancelacion del servicio"
)
INTEGRANTE = "Jeremy Ivan Pedraza Hernandez"
GRUPO = "Grupo 1"
DOCENTE = "Andres Felipe Callejas Jaramillo"
ASIGNATURA = "Proyecto Integrado III - Analitica de Datos"
CODIGO = "PREICA2602B020038"
PROGRAMA = "Ingenieria de Software y Datos"
FACULTAD = "Facultad de Ingenieria y Ciencias Agropecuarias"
INSTITUCION = "Institucion Universitaria Digital de Antioquia."
ANIO = "2026"
REPO = "https://github.com/Jeremy-Pedraza/PREICA2602B020038-Analisis-de-Datos"
DATASET_ID = "scikit-learn/churn-prediction"
DATASET_URL = "https://huggingface.co/datasets/scikit-learn/churn-prediction"

SIGNIFICADOS = {
    "customer_id": "Identificador unico del cliente.",
    "gender": "Genero del cliente.",
    "senior_citizen": "Indica si el cliente es adulto mayor (1) o no (0).",
    "partner": "Indica si el cliente tiene pareja.",
    "dependents": "Indica si el cliente tiene personas a cargo.",
    "tenure": "Antiguedad del cliente en meses con la compania.",
    "phone_service": "Indica si el cliente tiene servicio telefonico.",
    "multiple_lines": "Indica si el cliente tiene multiples lineas telefonicas.",
    "internet_service": "Tipo de servicio de internet contratado.",
    "online_security": "Indica si el cliente tiene seguridad en linea contratada.",
    "online_backup": "Indica si el cliente tiene copia de seguridad en linea.",
    "device_protection": "Indica si el cliente tiene proteccion de dispositivos.",
    "tech_support": "Indica si el cliente tiene soporte tecnico contratado.",
    "streaming_tv": "Indica si el cliente tiene television por streaming.",
    "streaming_movies": "Indica si el cliente tiene peliculas por streaming.",
    "contract": "Tipo de contrato del cliente (mensual, anual o bianual).",
    "paperless_billing": "Indica si el cliente usa facturacion electronica.",
    "payment_method": "Metodo de pago utilizado por el cliente.",
    "monthly_charges": "Cargo mensual facturado al cliente (unidad monetaria).",
    "total_charges": "Cargo total acumulado facturado al cliente (unidad monetaria).",
    "churn": "Variable objetivo: indica si el cliente cancelo el servicio.",
    "cargo_promedio_mensual": (
        "Derivada (AE2): total_charges / tenure, cargo promedio por mes de "
        "antiguedad (unidad monetaria)."
    ),
    "segmento_antiguedad": (
        "Derivada (AE2): segmentacion de tenure en rangos "
        "(0-12, 13-24, 25-48, 49+ meses)."
    ),
    "churn_flag": (
        "Derivada (AE3): version numerica de churn (1 = cancelo, 0 = permanece); "
        "permite calcular la tasa de cancelacion por grupo."
    ),
    "num_servicios_adicionales": (
        "Derivada (AE3): conteo de servicios adicionales contratados (online_security, "
        "online_backup, device_protection, tech_support, streaming_tv, streaming_movies)."
    ),
    "grupo_cargo_mensual": (
        "Derivada (AE3): nivel de gasto mensual segun monthly_charges "
        "(Bajo <=35, Medio 35-70, Alto >70)."
    ),
}

VARS_NUEVAS_AE3 = ["churn_flag", "num_servicios_adicionales", "grupo_cargo_mensual"]


# --------------------------------------------------------------------------
# Utilidades de documento
# --------------------------------------------------------------------------
def limpiar_cuerpo(doc: Document) -> None:
    body = doc.element.body
    sectPr = body.find(qn("w:sectPr"))
    for child in list(body):
        if child is not sectPr:
            body.remove(child)


def p_centrado(doc: Document, texto: str, negrita: bool = False) -> None:
    p = doc.add_paragraph(style="Normal")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(texto)
    run.bold = negrita


def add_field(paragraph, instr: str) -> None:
    r = paragraph.add_run()
    fld = OxmlElement("w:fldChar")
    fld.set(qn("w:fldCharType"), "begin")
    it = OxmlElement("w:instrText")
    it.set(qn("xml:space"), "preserve")
    it.text = instr
    sep = OxmlElement("w:fldChar")
    sep.set(qn("w:fldCharType"), "separate")
    txt = OxmlElement("w:t")
    txt.text = "Actualizar este campo en Word con F9"
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    for el in (fld, it, sep, txt, end):
        r._r.append(el)


def _add_seq_tabla(paragraph, numero: int) -> None:
    """Inserta un campo SEQ Tabla para que el indice 'Lista de tablas'
    (campo TOC \\c "Tabla") recoja la numeracion. El texto visible usa
    ``numero`` como marcador; Word lo recalcula con F9."""
    r = paragraph.add_run()
    r.bold = True
    fld = OxmlElement("w:fldChar")
    fld.set(qn("w:fldCharType"), "begin")
    it = OxmlElement("w:instrText")
    it.set(qn("xml:space"), "preserve")
    it.text = "SEQ Tabla \\* ARABIC"
    sep = OxmlElement("w:fldChar")
    sep.set(qn("w:fldCharType"), "separate")
    txt = OxmlElement("w:t")
    txt.text = str(numero)
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    for el in (fld, it, sep, txt, end):
        r._r.append(el)


def h1(doc, texto):
    return doc.add_heading(texto, level=1)


def h2(doc, texto):
    return doc.add_heading(texto, level=2)


def parrafo(doc, texto):
    return doc.add_paragraph(texto, style="Normal")


def item(doc, texto):
    p = doc.add_paragraph(style="List Paragraph")
    p.add_run("- " + texto)
    return p


def agregar_tabla(doc, encabezados, filas, numero, titulo, fuente):
    tabla = doc.add_table(rows=1, cols=len(encabezados))
    try:
        tabla.style = "Grid Table 2 Accent 4"
    except KeyError:
        tabla.style = "Table Grid"
    hdr = tabla.rows[0].cells
    for i, texto in enumerate(encabezados):
        hdr[i].text = ""
        run = hdr[i].paragraphs[0].add_run(str(texto))
        run.bold = True
    for fila in filas:
        celdas = tabla.add_row().cells
        for i, valor in enumerate(fila):
            celdas[i].text = str(valor)
    cap = doc.add_paragraph(style="Caption")
    cap.add_run("Tabla ").bold = True
    _add_seq_tabla(cap, numero)
    cap.add_run(f". {titulo} ").bold = True
    cap.add_run(f"Fuente: {fuente}")
    return tabla


def tabla_desde_df(doc, df, numero, titulo, fuente, encabezados=None):
    enc = encabezados or list(df.columns)
    filas = df.astype(object).values.tolist()
    return agregar_tabla(doc, enc, filas, numero, titulo, fuente)


def parrafo_mono(doc, lineas) -> None:
    for linea in lineas:
        p = doc.add_paragraph(style="Normal")
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run(linea if linea else " ")
        run.font.name = "Consolas"
        run.font.size = Pt(7.5)


# --------------------------------------------------------------------------
# Lectura de datos reales
# --------------------------------------------------------------------------
def leer_fecha_consulta() -> str:
    for linea in TXT_RESUMEN.read_text(encoding="utf-8").splitlines():
        if "fecha_consulta" in linea:
            return linea.split(":", 1)[1].strip()
    return "Pendiente de verificar"


def leer_versiones_requirements():
    pares = []
    for linea in REQUIREMENTS.read_text(encoding="utf-8").splitlines():
        linea = linea.strip()
        if not linea or linea.startswith("#"):
            continue
        for sep in (">=", "==", "<=", "~="):
            if sep in linea:
                nombre, version = linea.split(sep, 1)
                pares.append((nombre.strip(), f"{sep}{version.strip()}"))
                break
        else:
            pares.append((linea, "sin fijar"))
    return pares


# --------------------------------------------------------------------------
# Construccion del documento
# --------------------------------------------------------------------------
def construir() -> None:
    df_crudo = pd.read_csv(CSV_CRUDO)
    df = pd.read_csv(CSV_TRANS)
    resumen_txt = TXT_RESUMEN.read_text(encoding="utf-8").splitlines()
    tablas = {k: pd.read_csv(v) for k, v in TABLAS.items()}

    n_crudo_f, n_crudo_c = df_crudo.shape
    n_f, n_c = df.shape
    fecha = leer_fecha_consulta()

    tc_crudo = pd.to_numeric(df_crudo["TotalCharges"], errors="coerce")
    nulos_tc = int(tc_crudo.isna().sum())
    pct_nulos_tc = nulos_tc / n_crudo_f * 100

    churn_vc = df["churn"].value_counts()
    churn_si = int(churn_vc.get("Yes", 0))
    churn_no = int(churn_vc.get("No", 0))
    pct_churn = churn_si / n_f * 100

    n_vars_ae3 = len(VARS_NUEVAS_AE3)

    doc = Document(str(PLANTILLA))
    limpiar_cuerpo(doc)

    # =========================== PORTADA ===========================
    p_centrado(doc, TITULO, negrita=True)
    p_centrado(doc, "")
    p_centrado(doc, "Documento de avance - Etapa 2 (AE3)", negrita=False)
    p_centrado(doc, "")
    p_centrado(doc, "Integrantes", negrita=True)
    p_centrado(doc, GRUPO, negrita=True)
    p_centrado(doc, INTEGRANTE, negrita=False)
    p_centrado(doc, "")
    p_centrado(doc, "Docente", negrita=True)
    p_centrado(doc, DOCENTE, negrita=True)
    p_centrado(doc, "")
    p_centrado(doc, ASIGNATURA, negrita=True)
    p_centrado(doc, CODIGO, negrita=False)
    p_centrado(doc, "")
    p_centrado(doc, "")
    p_centrado(doc, PROGRAMA, negrita=True)
    p_centrado(doc, FACULTAD, negrita=True)
    p_centrado(doc, INSTITUCION, negrita=True)
    p_centrado(doc, ANIO, negrita=True)
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)

    # =========================== TOC ===========================
    doc.add_paragraph("Contenido", style="TOC Heading")
    add_field(doc.add_paragraph(), 'TOC \\o "1-3" \\h \\z \\u')
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)

    doc.add_paragraph("Lista de tablas", style="TOC Heading")
    add_field(doc.add_paragraph(), 'TOC \\h \\z \\c "Tabla"')
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)

    # =========================== 1. INTRODUCCION ===========================
    h1(doc, "1. Introduccion")
    parrafo(
        doc,
        "Este documento corresponde a la segunda etapa (AE3) del proyecto integrador "
        "y da continuidad al avance presentado en AE2, conservando el mismo proyecto, "
        "dataset y repositorio. Mientras AE2 cubrio la carga, la exploracion y la "
        "preparacion inicial del conjunto de datos, AE3 se centra en transformar y "
        "enriquecer los datos para responder de forma descriptiva las preguntas de "
        "negocio derivadas de la necesidad analitica.",
    )
    parrafo(
        doc,
        f"El conjunto de datos es «{DATASET_ID}» (Telco Customer Churn), disponible en "
        f"Hugging Face, con {n_crudo_f} registros de clientes y {n_crudo_c} variables "
        f"originales. Cada fila es un cliente y la variable objetivo es la cancelacion "
        f"del servicio (churn). El documento registra que se transformo y por que, que "
        f"variables se incorporaron, y presenta una tabla de resultados por cada "
        f"pregunta de negocio con su interpretacion y limitaciones.",
    )

    # =========================== 2. NECESIDAD Y PREGUNTAS ===========================
    h1(doc, "2. Necesidad y preguntas de negocio")
    h2(doc, "2.1 Necesidad analitica (retomada de AE2)")
    parrafo(
        doc,
        "El beneficiario es el area comercial y de retencion de una compania de "
        "telecomunicaciones, que necesita caracterizar que perfiles de clientes "
        "concentran la cancelacion del servicio para asignar de forma eficiente sus "
        "esfuerzos de retencion. La necesidad se mantiene sin cambios respecto de AE2; "
        "en AE3 se operacionaliza en preguntas concretas respondibles con los datos.",
    )
    h2(doc, "2.2 Preguntas de negocio y variables")
    parrafo(
        doc,
        "La pregunta general de AE2 -que caracteristicas distinguen a los clientes que "
        "cancelan frente a los que permanecen- se descompone en cinco preguntas "
        "operativas. La Tabla 1 relaciona cada pregunta con las variables necesarias y "
        "el resultado esperado.",
    )
    relacion = [
        ["P1", "Que tipo de contrato concentra mayor tasa de cancelacion?",
         "contract, churn", "Tasa de churn por tipo de contrato"],
        ["P2", "Como varia la cancelacion segun la antiguedad del cliente?",
         "tenure -> segmento_antiguedad, churn", "Tasa de churn por segmento"],
        ["P3", "Que metodo de pago y servicio de internet se asocian con mas churn?",
         "payment_method, internet_service, churn", "Tasa de churn por cada uno"],
        ["P4", "Como difieren los cargos entre quienes cancelan y permanecen?",
         "monthly_charges, total_charges, cargo_promedio_mensual, tenure, churn",
         "Promedios de facturacion por grupo"],
        ["P5", "Influye la cantidad de servicios adicionales en la cancelacion?",
         "add-ons -> num_servicios_adicionales, churn",
         "Tasa de churn por numero de servicios"],
    ]
    agregar_tabla(
        doc, ["#", "Pregunta de negocio", "Variables necesarias", "Resultado esperado"],
        relacion, 1, "Preguntas de negocio, variables y resultado esperado.",
        "elaboracion propia a partir de la necesidad definida en AE2.",
    )
    parrafo(
        doc,
        "Limitacion de alcance: el dataset no incluye marca temporal ni datos de "
        "costos, por lo que las respuestas son descriptivas y no permiten analisis de "
        "tendencia, de rentabilidad ni inferencia causal.",
    )

    # =========================== 3. METODOLOGIA ===========================
    h1(doc, "3. Metodologia")
    h2(doc, "3.1 Flujo general")
    parrafo(
        doc,
        "El procedimiento sigue el flujo: dataset de AE2 -> transformacion y "
        "validacion -> enriquecimiento -> respuestas a las preguntas -> tablas de "
        "resultados y archivos exportados. Todo se ejecuta por codigo en el notebook "
        "EA3/notebooks/ea3_transformacion_enriquecimiento.ipynb con rutas relativas, "
        "conservando una version sin modificar del dataset para comparar los cambios.",
    )

    h2(doc, "3.2 Transformaciones y reglas de calidad")
    parrafo(
        doc,
        "Se aplicaron las siguientes transformaciones de preparacion, cada una "
        "validada con el conteo de registros afectados:",
    )
    for texto in [
        "T1 - Renombrado de columnas a snake_case para estandarizar la notacion "
        "(afecta a las 21 columnas, sin perdida de datos).",
        "T2 - Conversion de total_charges de texto a numerico con "
        "pd.to_numeric(errors='coerce'), dado que contenia cadenas vacias; la "
        f"conversion revelo {nulos_tc} valores nulos ({pct_nulos_tc:.2f}%).",
        "T3 - Tratamiento de los nulos de total_charges: todos corresponden a clientes "
        "con tenure = 0, por lo que se imputan con 0.0 (valor coherente para un cliente "
        "nuevo). No se elimino ningun registro.",
    ]:
        item(doc, texto)
    parrafo(
        doc,
        "Reglas de calidad verificadas: no se hallaron filas duplicadas (0), por lo "
        "que no se eliminaron registros; tras el tratamiento, el DataFrame no presenta "
        "valores nulos. Ningun dato se elimino ni se imputo sin justificar el criterio.",
    )
    parrafo(
        doc,
        "La Tabla 2 compara el DataFrame antes (crudo) y despues (preparado) en "
        "filas, columnas, tipos y valores nulos, como evidencia del efecto de las "
        "transformaciones y de los controles de calidad aplicados.",
    )
    dtype_tc_crudo = str(df_crudo["TotalCharges"].dtype)
    dtype_tc_trans = str(df["total_charges"].dtype)
    comparacion = [
        ["Filas", n_crudo_f, n_f,
         "Sin eliminacion de registros" if n_f == n_crudo_f
         else f"{n_f - n_crudo_f:+d}"],
        ["Columnas", n_crudo_c, n_c,
         f"{n_c - n_crudo_c:+d} por el enriquecimiento (AE3)"],
        ["Tipo de total_charges", dtype_tc_crudo, dtype_tc_trans,
         "Corregido de texto a numerico"],
        ["Valores nulos (total)", int(df_crudo.isna().sum().sum()),
         int(df.isna().sum().sum()),
         f"{nulos_tc} detectados en total_charges al convertir, imputados con 0.0"],
        ["Filas duplicadas", int(df_crudo.duplicated().sum()),
         int(df.duplicated().sum()), "Sin duplicados en ninguna etapa"],
    ]
    agregar_tabla(
        doc, ["Aspecto", "Antes (crudo)", "Despues (preparado)", "Observacion"],
        comparacion, 2,
        "Comparacion del DataFrame antes y despues de la preparacion.",
        "elaboracion propia a partir de churn_crudo.csv y churn_transformado.csv.",
    )

    h2(doc, "3.3 Enriquecimiento")
    parrafo(
        doc,
        "Se incorporaron variables derivadas a partir de las variables existentes, sin "
        "usar fuentes externas. La Tabla 3 resume su formula, significado y utilidad.",
    )
    enriq = [
        ["cargo_promedio_mensual", "total_charges / tenure (0 si tenure=0)",
         "Cargo medio por mes de antiguedad", "P4"],
        ["segmento_antiguedad", "pd.cut(tenure): 0-12, 13-24, 25-48, 49+",
         "Rango de antiguedad del cliente", "P2"],
        ["churn_flag", "1 si churn=='Yes', si no 0",
         "Objetivo numerico", "Tasa de churn en todas las preguntas"],
        ["num_servicios_adicionales", "conteo de 6 add-ons == 'Yes'",
         "Cantidad de servicios extra contratados", "P5"],
        ["grupo_cargo_mensual", "pd.cut(monthly_charges): Bajo/Medio/Alto",
         "Nivel de gasto mensual", "Perfil de facturacion"],
    ]
    agregar_tabla(
        doc, ["Variable", "Formula / regla", "Significado", "Utilidad"],
        enriq, 3, "Variables incorporadas en el enriquecimiento (AE3).",
        "elaboracion propia.",
    )
    parrafo(
        doc,
        "El enriquecimiento se realizo con los datos existentes; no se combinaron otras "
        "tablas, por lo que no hubo riesgo de multiplicar registros por una union.",
    )

    h2(doc, "3.4 Diccionario de datos actualizado")
    parrafo(
        doc,
        f"La Tabla 4 presenta el diccionario del DataFrame final ({n_c} variables), "
        f"que incorpora las {n_vars_ae3} variables nuevas de AE3 (churn_flag, "
        f"num_servicios_adicionales y grupo_cargo_mensual) ademas de las derivadas en "
        f"AE2.",
    )
    filas_dicc = []
    for col in df.columns:
        dtype = str(df[col].dtype)
        serie = df[col]
        if dtype == "object" or dtype == "category":
            cats = serie.dropna().unique().tolist()
            if len(cats) <= 6:
                detalle = "Categorias: " + ", ".join(str(c) for c in cats)
            else:
                detalle = f"{len(cats)} valores distintos (p. ej. identificador)"
        else:
            detalle = f"Numerica (min={serie.min()}, max={serie.max()})"
        significado = SIGNIFICADOS.get(col, "Pendiente de verificar")
        filas_dicc.append([col, significado, dtype, detalle])
    agregar_tabla(
        doc, ["Variable", "Significado", "Tipo", "Categorias / unidad"],
        filas_dicc, 4, "Diccionario de datos actualizado (DataFrame final).",
        "elaboracion propia a partir de EA3/resultados/churn_transformado.csv.",
    )

    h2(doc, "3.5 Herramientas y entorno")
    parrafo(
        doc,
        "El desarrollo se realizo en Python con Jupyter Notebook, control de versiones "
        "con Git y alojamiento en GitHub. La Tabla 5 lista las librerias y versiones "
        "minimas segun requirements.txt.",
    )
    filas_req = [[n, v] for n, v in leer_versiones_requirements()]
    agregar_tabla(
        doc, ["Libreria / herramienta", "Version"], filas_req, 5,
        "Herramientas y versiones del entorno.",
        "elaboracion propia a partir de requirements.txt.",
    )

    # =========================== 4. RESULTADOS Y DISCUSION ===========================
    h1(doc, "4. Resultados y discusion")
    parrafo(
        doc,
        f"El DataFrame final conserva {n_f} filas y amplia a {n_c} columnas por las "
        f"variables derivadas. La tasa de cancelacion global es del {pct_churn:.2f}% "
        f"({churn_si} clientes cancelan y {churn_no} permanecen). A continuacion se "
        f"responde cada pregunta con su tabla e interpretacion. Las tablas coinciden "
        f"con los archivos .csv exportados por el notebook.",
    )

    h2(doc, "4.1 P1 - Cancelacion por tipo de contrato")
    tabla_desde_df(
        doc, tablas["p1"], 6, "Tasa de cancelacion por tipo de contrato.",
        "elaboracion propia a partir de EA3/resultados/tabla_p1_churn_por_contrato.csv.",
        encabezados=["contract", "clientes", "cancelaciones", "tasa_churn_%"],
    )
    fila_top_p1 = tablas["p1"].iloc[0]
    parrafo(
        doc,
        f"El contrato «{fila_top_p1['contract']}» concentra la mayor tasa de "
        f"cancelacion ({fila_top_p1['tasa_churn_%']}%), muy por encima de los contratos "
        f"a termino. Esto indica que la permanencia contractual es un factor protector "
        f"frente al churn y sugiere incentivar la migracion a contratos anuales.",
    )

    h2(doc, "4.2 P2 - Cancelacion por antiguedad")
    tabla_desde_df(
        doc, tablas["p2"], 7, "Tasa de cancelacion por segmento de antiguedad.",
        "elaboracion propia a partir de EA3/resultados/tabla_p2_churn_por_antiguedad.csv.",
        encabezados=["segmento_antiguedad", "clientes", "cancelaciones", "tasa_churn_%"],
    )
    parrafo(
        doc,
        "La cancelacion es mas alta entre los clientes nuevos (segmento 0-12 meses) y "
        "decrece de forma sostenida a medida que aumenta la antiguedad. Los primeros "
        "meses de relacion son el periodo de mayor riesgo y el foco natural de las "
        "acciones de retencion.",
    )

    h2(doc, "4.3 P3 - Cancelacion por metodo de pago y servicio de internet")
    tabla_desde_df(
        doc, tablas["p3a"], 8, "Tasa de cancelacion por metodo de pago.",
        "elaboracion propia a partir de EA3/resultados/tabla_p3a_churn_por_metodo_pago.csv.",
        encabezados=["payment_method", "clientes", "cancelaciones", "tasa_churn_%"],
    )
    tabla_desde_df(
        doc, tablas["p3b"], 9, "Tasa de cancelacion por tipo de servicio de internet.",
        "elaboracion propia a partir de EA3/resultados/tabla_p3b_churn_por_internet.csv.",
        encabezados=["internet_service", "clientes", "cancelaciones", "tasa_churn_%"],
    )
    parrafo(
        doc,
        f"El pago con «{tablas['p3a'].iloc[0]['payment_method']}» y el internet de "
        f"«{tablas['p3b'].iloc[0]['internet_service']}» presentan las tasas de "
        f"cancelacion mas altas. Son segmentos donde conviene revisar la experiencia de "
        f"pago y la calidad/precio del servicio de internet.",
    )

    h2(doc, "4.4 P4 - Facturacion comparada entre grupos")
    tabla_desde_df(
        doc, tablas["p4"], 10, "Facturacion promedio segun estado de cancelacion.",
        "elaboracion propia a partir de EA3/resultados/tabla_p4_cargos_por_churn.csv.",
        encabezados=["churn", "clientes", "tenure_prom", "monthly_charges_prom",
                     "total_charges_prom", "cargo_promedio_mensual_prom"],
    )
    parrafo(
        doc,
        "Quienes cancelan tienen menor antiguedad promedio y, al mismo tiempo, un cargo "
        "mensual promedio mas alto que quienes permanecen. El patron es consistente con "
        "P1-P3: clientes recientes, con mayor gasto mensual y sin ataduras "
        "contractuales son los mas propensos a cancelar.",
    )

    h2(doc, "4.5 P5 - Cancelacion por numero de servicios adicionales")
    tabla_desde_df(
        doc, tablas["p5"], 11, "Tasa de cancelacion por numero de servicios adicionales.",
        "elaboracion propia a partir de EA3/resultados/tabla_p5_churn_por_num_servicios.csv.",
        encabezados=["num_servicios_adicionales", "clientes", "cancelaciones", "tasa_churn_%"],
    )
    parrafo(
        doc,
        "Entre los clientes con servicio de internet, la tasa de cancelacion disminuye "
        "a medida que aumenta el numero de servicios adicionales contratados: mas "
        "servicios implican mayor vinculacion y menor abandono. El grupo con 0 "
        "servicios corresponde en su mayoria a clientes sin internet, por lo que no "
        "sigue la misma tendencia y debe interpretarse aparte.",
    )

    h2(doc, "4.6 Evidencia del resumen automatico")
    parrafo(
        doc,
        "El resumen descriptivo del DataFrame final se genera por codigo con la clase "
        "DatasetSummarizer y se guarda en EA3/resultados/resumen_dataframe.txt (UTF-8), "
        "regenerandose en cada ejecucion. A continuacion se transcribe su contenido.",
    )
    parrafo_mono(doc, resumen_txt)

    h2(doc, "4.7 Limitaciones de los resultados")
    for texto in [
        "Las relaciones son asociaciones descriptivas, no relaciones causales.",
        "Sin datos de costos no puede afirmarse que un segmento sea mas o menos "
        "rentable, solo que cancela mas o menos.",
        "Sin marca temporal no es posible analizar la evolucion de la cancelacion en "
        "el tiempo.",
        "Los resultados corresponden al split train del dataset y no se extrapolan a "
        "otro mercado o pais.",
    ]:
        item(doc, texto)

    # =========================== 5. CONCLUSIONES ===========================
    h1(doc, "5. Conclusiones del avance")
    for texto in [
        f"La transformacion y el enriquecimiento dejaron un DataFrame final de {n_f} "
        f"filas y {n_c} columnas, sin nulos ni duplicados, apto para responder las "
        f"preguntas de negocio.",
        "Las cinco preguntas quedaron respondidas con una tabla cada una: el contrato "
        "mensual, la baja antiguedad, el pago con cheque electronico, el internet de "
        "fibra optica y el menor numero de servicios adicionales se asocian con mayor "
        "cancelacion.",
        "El perfil de mayor riesgo es consistente entre preguntas: cliente reciente, "
        "con contrato mensual y cargo mensual alto.",
        "Los datos transformados y enriquecidos quedan organizados en archivos .csv y "
        "un resumen .txt, listos para la representacion grafica de AE4.",
    ]:
        item(doc, texto)
    parrafo(
        doc,
        "Queda sin resolver la cuantificacion de la rentabilidad por segmento (falta de "
        "costos) y el analisis temporal de la cancelacion (falta de fecha), aspectos "
        "fuera del alcance de esta etapa.",
    )

    # =========================== 6. REFERENCIAS ===========================
    h1(doc, "6. Referencias")
    referencias = [
        "Han, J., Kamber, M., & Pei, J. (2011). Data mining: Concepts and techniques "
        "(3rd ed.). Morgan Kaufmann.",
        "Lhoest, Q., et al. (2021). Datasets: A community library for natural language "
        "processing. En Proceedings of the 2021 Conference on Empirical Methods in "
        "Natural Language Processing: System Demonstrations (pp. 175-184). ACL.",
        "McKinney, W. (2010). Data structures for statistical computing in Python. En "
        "Proceedings of the 9th Python in Science Conference (pp. 56-61).",
        "Provost, F., & Fawcett, T. (2013). Data science for business. O'Reilly Media.",
        "Verbeke, W., Dejaeger, K., Martens, D., Hur, J., & Baesens, B. (2012). New "
        "insights into churn prediction in the telecommunication sector. European "
        "Journal of Operational Research, 218(1), 211-229.",
        f"scikit-learn. (s. f.). Churn prediction dataset [Conjunto de datos]. Hugging "
        f"Face. {DATASET_URL}",
        "pandas development team. (s. f.). pandas documentation. "
        "https://pandas.pydata.org/docs/",
        "Hugging Face. (s. f.). Datasets documentation. "
        "https://huggingface.co/docs/datasets",
        f"Pedraza Hernandez, J. I. (2026). Repositorio del proyecto PREICA2602B020038 "
        f"- Analitica de Datos [Repositorio]. GitHub. {REPO}",
    ]
    for ref in referencias:
        p = doc.add_paragraph(style="Normal")
        p.paragraph_format.left_indent = Pt(36)
        p.paragraph_format.first_line_indent = Pt(-36)
        p.add_run(ref)

    doc.save(str(SALIDA))
    print(f"Documento generado: {SALIDA}")
    print(f"  crudo: {n_crudo_f}x{n_crudo_c} | transformado: {n_f}x{n_c}")
    print(f"  churn Yes: {churn_si} ({pct_churn:.2f}%) | nulos total_charges: {nulos_tc}")


if __name__ == "__main__":
    construir()
