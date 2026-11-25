# -*- coding: utf-8 -*-
"""Genera el DOCUMENTO COMPLETO del proyecto integrador (entrega final EA4).

Integra el trabajo desarrollado durante el periodo (EA1, AE2, EA3 y EA4) en un
unico informe: portada, indices, introduccion, necesidad, justificacion,
objetivos, marco teorico, metodologia, transformacion y enriquecimiento de
datos, metricas y visualizaciones, resultados y discusion, conclusiones y
referencias. Incorpora las figuras de la aplicacion de Streamlit.

REGLA DURA: ninguna cifra se escribe a mano. Todas las dimensiones, tasas,
estadisticas y tablas se leen en tiempo de ejecucion desde ``EA4/resultados/``
(datos preparados, tablas por pregunta y resumen .txt) usando el modulo
``metricas`` -- la misma fuente que alimenta la aplicacion de Streamlit, de modo
que el documento y la app muestran cifras identicas.

Requiere las figuras PNG (``python EA4/src/generar_figuras.py``).

Uso:
    python EA4/src/generar_figuras.py
    python EA4/src/generar_documento.py
"""

from __future__ import annotations

import pathlib

import pandas as pd
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, Inches

try:
    from . import metricas as M
    from . import generar_figuras
except ImportError:  # ejecucion como script
    import sys
    sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent))
    import metricas as M
    import generar_figuras

# --------------------------------------------------------------------------
# Rutas
# --------------------------------------------------------------------------
EA4 = pathlib.Path(__file__).resolve().parent.parent
RESULTADOS = EA4 / "resultados"
FIGURAS = RESULTADOS / "figuras"
DOCUMENTO = EA4 / "documento"
RAIZ = EA4.parent

PLANTILLA = DOCUMENTO / "plantilla_iudigital.docx"
SALIDA = DOCUMENTO / "Pedraza_Jeremy_AE4_Analitica_de_Datos.docx"
REQUIREMENTS = RAIZ / "requirements.txt"

P = M.PROYECTO
STREAMLIT_URL = "https://bmt9gfifmznktubbdfv9k3.streamlit.app"  # app publicada (Streamlit Community Cloud)

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
    "cargo_promedio_mensual": "Derivada (AE2): total_charges / tenure (unidad monetaria).",
    "segmento_antiguedad": "Derivada (AE2): segmentacion de tenure (0-12, 13-24, 25-48, 49+).",
    "churn_flag": "Derivada (EA3): churn numerico (1 cancelo / 0 permanece).",
    "num_servicios_adicionales": "Derivada (EA3): conteo de 6 servicios adicionales.",
    "grupo_cargo_mensual": "Derivada (EA3): nivel de gasto (Bajo/Medio/Alto).",
}


# --------------------------------------------------------------------------
# Utilidades de documento
# --------------------------------------------------------------------------
def limpiar_cuerpo(doc: Document) -> None:
    body = doc.element.body
    sectPr = body.find(qn("w:sectPr"))
    for child in list(body):
        if child is not sectPr:
            body.remove(child)


def p_centrado(doc, texto, negrita=False):
    p = doc.add_paragraph(style="Normal")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(texto)
    run.bold = negrita


def add_field(paragraph, instr: str) -> None:
    r = paragraph.add_run()
    fld = OxmlElement("w:fldChar"); fld.set(qn("w:fldCharType"), "begin")
    it = OxmlElement("w:instrText"); it.set(qn("xml:space"), "preserve"); it.text = instr
    sep = OxmlElement("w:fldChar"); sep.set(qn("w:fldCharType"), "separate")
    txt = OxmlElement("w:t"); txt.text = "Actualizar este campo en Word con F9"
    end = OxmlElement("w:fldChar"); end.set(qn("w:fldCharType"), "end")
    for el in (fld, it, sep, txt, end):
        r._r.append(el)


def _add_seq(paragraph, etiqueta: str, numero: int) -> None:
    """Inserta un campo SEQ (Tabla o Figura) para los indices automaticos."""
    r = paragraph.add_run(); r.bold = True
    fld = OxmlElement("w:fldChar"); fld.set(qn("w:fldCharType"), "begin")
    it = OxmlElement("w:instrText"); it.set(qn("xml:space"), "preserve")
    it.text = f"SEQ {etiqueta} \\* ARABIC"
    sep = OxmlElement("w:fldChar"); sep.set(qn("w:fldCharType"), "separate")
    txt = OxmlElement("w:t"); txt.text = str(numero)
    end = OxmlElement("w:fldChar"); end.set(qn("w:fldCharType"), "end")
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
        hdr[i].paragraphs[0].add_run(str(texto)).bold = True
    for fila in filas:
        celdas = tabla.add_row().cells
        for i, valor in enumerate(fila):
            celdas[i].text = str(valor)
    cap = doc.add_paragraph(style="Caption")
    cap.add_run("Tabla ").bold = True
    _add_seq(cap, "Tabla", numero)
    cap.add_run(f". {titulo} ").bold = True
    cap.add_run(f"Fuente: {fuente}")
    return tabla


def tabla_desde_df(doc, df, numero, titulo, fuente, encabezados=None):
    enc = encabezados or list(df.columns)
    filas = df.astype(object).values.tolist()
    return agregar_tabla(doc, enc, filas, numero, titulo, fuente)


def agregar_figura(doc, ruta_png: pathlib.Path, numero: int, titulo: str, fuente: str):
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run().add_picture(str(ruta_png), width=Inches(5.6))
    cap = doc.add_paragraph(style="Caption")
    cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cap.add_run("Figura ").bold = True
    _add_seq(cap, "Figura", numero)
    cap.add_run(f". {titulo} ").bold = True
    cap.add_run(f"Fuente: {fuente}")


def parrafo_mono(doc, lineas) -> None:
    for linea in lineas:
        p = doc.add_paragraph(style="Normal")
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run(linea if linea else " ")
        run.font.name = "Consolas"
        run.font.size = Pt(7.5)


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
    # Asegurar figuras
    if not all((FIGURAS / p.figura).exists() for p in M.PREGUNTAS):
        generar_figuras.construir()

    df = M.cargar_transformado()
    df_crudo = pd.read_csv(M.CSV_CRUDO)
    tablas = M.cargar_tablas()
    kpis = M.kpis_globales(df)
    resumen_txt = M.leer_resumen().splitlines()
    fecha = M.leer_fecha_consulta()

    n_crudo_f, n_crudo_c = df_crudo.shape
    n_f, n_c = df.shape
    tc_crudo = pd.to_numeric(df_crudo["TotalCharges"], errors="coerce")
    nulos_tc = int(tc_crudo.isna().sum())
    pct_nulos_tc = nulos_tc / n_crudo_f * 100

    doc = Document(str(PLANTILLA))
    limpiar_cuerpo(doc)

    # ============================ PORTADA ============================
    p_centrado(doc, P["titulo"], negrita=True)
    p_centrado(doc, "Informe final del proyecto integrador", negrita=False)
    p_centrado(doc, "Metricas y visualizacion de datos (EA4)", negrita=False)
    p_centrado(doc, "")
    p_centrado(doc, "Integrantes", negrita=True)
    p_centrado(doc, P["grupo"], negrita=True)
    p_centrado(doc, P["integrante"], negrita=False)
    p_centrado(doc, "")
    p_centrado(doc, "Docente", negrita=True)
    p_centrado(doc, P["docente"], negrita=True)
    p_centrado(doc, "")
    p_centrado(doc, P["asignatura"], negrita=True)
    p_centrado(doc, P["codigo"], negrita=False)
    p_centrado(doc, "")
    p_centrado(doc, P["programa"], negrita=True)
    p_centrado(doc, P["facultad"], negrita=True)
    p_centrado(doc, P["institucion"] + ".", negrita=True)
    p_centrado(doc, P["anio"], negrita=True)
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)

    # ============================ INDICES ============================
    doc.add_paragraph("Contenido", style="TOC Heading")
    add_field(doc.add_paragraph(), 'TOC \\o "1-3" \\h \\z \\u')
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)

    doc.add_paragraph("Lista de tablas", style="TOC Heading")
    add_field(doc.add_paragraph(), 'TOC \\h \\z \\c "Tabla"')
    doc.add_paragraph()
    doc.add_paragraph("Lista de figuras", style="TOC Heading")
    add_field(doc.add_paragraph(), 'TOC \\h \\z \\c "Figura"')
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)

    # ============================ 1. INTRODUCCION ============================
    h1(doc, "1. Introduccion")
    parrafo(
        doc,
        "Este documento constituye el informe final del proyecto integrador de la "
        "asignatura Proyecto Integrado III - Analitica de Datos e integra el trabajo "
        "desarrollado durante el periodo. El proyecto aborda el abandono de clientes "
        "(churn) en una compania de telecomunicaciones y recorre todas las etapas del "
        "analisis: la carga y preparacion del dataset (EA1 y AE2), su transformacion y "
        "enriquecimiento para responder preguntas de negocio (EA3) y, en esta entrega "
        "(EA4), la construccion de metricas y visualizaciones presentadas en una "
        "aplicacion interactiva de Streamlit.",
    )
    parrafo(
        doc,
        f"El conjunto de datos es «{P['dataset_id']}» (Telco Customer Churn), disponible "
        f"en Hugging Face, con {n_crudo_f} registros de clientes. Tras la preparacion y el "
        f"enriquecimiento, el DataFrame final tiene {n_f} filas y {n_c} variables. La "
        f"tasa de cancelacion global es del {kpis['tasa_churn']:.2f}%. El documento "
        f"presenta, para cada pregunta de negocio, su metrica, su grafico y su "
        f"interpretacion, de forma coherente con la aplicacion publicada en el "
        f"repositorio.",
    )
    parrafo(doc, f"Aplicacion de Streamlit: {STREAMLIT_URL}")
    parrafo(doc, f"Repositorio del proyecto: {P['repo']}")

    # ============================ 2. NECESIDAD ============================
    h1(doc, "2. Necesidad analitica")
    parrafo(doc, M.NECESIDAD)
    parrafo(
        doc,
        "El beneficiario es el area comercial y de retencion de la compania, que necesita "
        "caracterizar que perfiles de clientes concentran la cancelacion para asignar de "
        "forma eficiente sus esfuerzos de retencion. La necesidad se mantiene estable a lo "
        "largo de todas las etapas del proyecto.",
    )

    # ============================ 3. JUSTIFICACION ============================
    h1(doc, "3. Justificacion")
    parrafo(
        doc,
        f"La retencion de clientes es mas economica que la adquisicion de nuevos, por lo "
        f"que anticipar el abandono tiene un impacto directo en los ingresos. En esta base "
        f"de clientes, los {kpis['churn_si']} clientes que cancelan representan el "
        f"{kpis['pct_ingreso_riesgo']:.1f}% del ingreso mensual recurrente "
        f"({kpis['ingreso_en_riesgo']:,.0f} de {kpis['ingreso_mensual_total']:,.0f} "
        f"unidades monetarias). Identificar los perfiles de mayor riesgo permite focalizar "
        f"las acciones de retencion donde generan mayor retorno y sustentar decisiones con "
        f"evidencia en lugar de intuicion.",
    )

    # ============================ 4. OBJETIVOS ============================
    h1(doc, "4. Objetivos")
    h2(doc, "4.1 Objetivo general")
    parrafo(
        doc,
        "Analizar, mediante metricas y visualizaciones, los factores asociados al abandono "
        "de clientes (churn) en una compania de telecomunicaciones, y presentarlos en una "
        "aplicacion interactiva que apoye las decisiones de retencion.",
    )
    h2(doc, "4.2 Objetivos especificos")
    for texto in [
        "Cargar, preparar y validar el dataset de churn de forma reproducible.",
        "Enriquecer el conjunto de datos con variables derivadas utiles para el analisis.",
        "Definir una metrica por cada pregunta de negocio con su calculo y unidad.",
        "Construir visualizaciones claras que respondan cada pregunta.",
        "Publicar los resultados en una aplicacion de Streamlit vinculada al repositorio.",
    ]:
        item(doc, texto)

    # ============================ 5. MARCO TEORICO ============================
    h1(doc, "5. Marco teorico")
    parrafo(
        doc,
        "El abandono de clientes (customer churn) es la perdida de clientes en un periodo y "
        "es un indicador central en industrias de servicios por suscripcion como las "
        "telecomunicaciones (Verbeke et al., 2012). Su analisis se enmarca en la analitica "
        "descriptiva, que busca describir que ha ocurrido mediante metricas y "
        "visualizaciones, a diferencia de la analitica predictiva o causal (Provost & "
        "Fawcett, 2013).",
    )
    parrafo(
        doc,
        "Una metrica es una medida cuantitativa con un calculo y una unidad definidos; en "
        "este proyecto la principal es la tasa de churn (porcentaje de clientes que "
        "cancelan sobre el total de un grupo). La visualizacion de datos traduce esas "
        "metricas en graficos que facilitan la comparacion entre grupos (Han, Kamber & "
        "Pei, 2011). El flujo de preparacion -limpieza, transformacion y enriquecimiento- "
        "sigue las practicas habituales del tratamiento de datos tabulares con pandas "
        "(McKinney, 2010), y la presentacion se realiza con Streamlit, un marco de Python "
        "para construir aplicaciones de datos interactivas.",
    )

    # ============================ 6. METODOLOGIA ============================
    h1(doc, "6. Metodologia")
    h2(doc, "6.1 Flujo general")
    parrafo(
        doc,
        "El proyecto sigue el flujo: carga del dataset (EA1/AE2) -> preparacion y "
        "validacion -> enriquecimiento con variables derivadas (EA3) -> definicion de "
        "metricas y construccion de graficos por pregunta -> publicacion en Streamlit "
        "(EA4). Todo se ejecuta por codigo con rutas relativas, conservando una version sin "
        "modificar del dataset para comparar los cambios. Las mismas tablas de resultados "
        "(.csv) alimentan el notebook, el documento y la aplicacion.",
    )

    h2(doc, "6.2 Transformaciones y reglas de calidad")
    parrafo(doc, "Se aplicaron las siguientes transformaciones, cada una validada con el "
                 "conteo de registros afectados:")
    for texto in [
        "T1 - Renombrado de columnas a snake_case para estandarizar la notacion.",
        f"T2 - Conversion de total_charges de texto a numerico con "
        f"pd.to_numeric(errors='coerce'); la conversion revelo {nulos_tc} valores nulos "
        f"({pct_nulos_tc:.2f}%).",
        "T3 - Imputacion de esos nulos (clientes con tenure = 0) con 0.0, sin eliminar "
        "registros.",
    ]:
        item(doc, texto)
    dtype_tc_crudo = str(df_crudo["TotalCharges"].dtype)
    dtype_tc_trans = str(df["total_charges"].dtype)
    comparacion = [
        ["Filas", n_crudo_f, n_f, "Sin eliminacion de registros"],
        ["Columnas", n_crudo_c, n_c, f"{n_c - n_crudo_c:+d} por el enriquecimiento"],
        ["Tipo de total_charges", dtype_tc_crudo, dtype_tc_trans, "Corregido a numerico"],
        ["Valores nulos (total)", int(df_crudo.isna().sum().sum()),
         int(df.isna().sum().sum()), f"{nulos_tc} imputados con 0.0"],
        ["Filas duplicadas", int(df_crudo.duplicated().sum()),
         int(df.duplicated().sum()), "Sin duplicados"],
    ]
    agregar_tabla(
        doc, ["Aspecto", "Antes (crudo)", "Despues (preparado)", "Observacion"],
        comparacion, 1, "Comparacion del DataFrame antes y despues de la preparacion.",
        "elaboracion propia a partir de churn_crudo.csv y churn_transformado.csv.",
    )

    h2(doc, "6.3 Enriquecimiento")
    parrafo(doc, "Se incorporaron variables derivadas a partir de las existentes, sin "
                 "fuentes externas. La Tabla 2 resume su formula y utilidad.")
    enriq = [
        ["cargo_promedio_mensual", "total_charges / tenure (0 si tenure=0)",
         "Cargo medio por mes de antiguedad", "P4"],
        ["segmento_antiguedad", "pd.cut(tenure): 0-12, 13-24, 25-48, 49+",
         "Rango de antiguedad", "P2"],
        ["churn_flag", "1 si churn=='Yes', si no 0", "Objetivo numerico", "Todas"],
        ["num_servicios_adicionales", "conteo de 6 add-ons == 'Yes'",
         "Servicios extra contratados", "P5"],
        ["grupo_cargo_mensual", "pd.cut(monthly_charges): Bajo/Medio/Alto",
         "Nivel de gasto mensual", "Perfil"],
    ]
    agregar_tabla(
        doc, ["Variable", "Formula / regla", "Significado", "Utilidad"],
        enriq, 2, "Variables incorporadas en el enriquecimiento.",
        "elaboracion propia.",
    )

    h2(doc, "6.4 Diccionario de datos")
    parrafo(doc, f"La Tabla 3 presenta el diccionario del DataFrame final ({n_c} "
                 f"variables).")
    filas_dicc = []
    for col in df.columns:
        dtype = str(df[col].dtype)
        serie = df[col]
        if dtype in ("object", "category"):
            cats = serie.dropna().unique().tolist()
            detalle = ("Categorias: " + ", ".join(str(c) for c in cats)) if len(cats) <= 6 \
                else f"{len(cats)} valores distintos"
        else:
            detalle = f"Numerica (min={serie.min()}, max={serie.max()})"
        filas_dicc.append([col, SIGNIFICADOS.get(col, "n/d"), dtype, detalle])
    agregar_tabla(
        doc, ["Variable", "Significado", "Tipo", "Categorias / unidad"],
        filas_dicc, 3, "Diccionario de datos del DataFrame final.",
        "elaboracion propia a partir de churn_transformado.csv.",
    )

    h2(doc, "6.5 Herramientas y entorno")
    parrafo(doc, "La Tabla 4 lista las librerias y versiones minimas segun "
                 "requirements.txt.")
    filas_req = [[n, v] for n, v in leer_versiones_requirements()]
    agregar_tabla(
        doc, ["Libreria / herramienta", "Version"], filas_req, 4,
        "Herramientas y versiones del entorno.",
        "elaboracion propia a partir de requirements.txt.",
    )

    # ============================ 7. METRICAS Y VISUALIZACIONES ============================
    h1(doc, "7. Metricas y visualizaciones")
    parrafo(
        doc,
        "Para cada pregunta de negocio se define una metrica con su calculo y unidad, y se "
        "selecciona un grafico que permite comparar los resultados. La Tabla 5 resume las "
        "metricas; las figuras que siguen son las mismas que presenta la aplicacion de "
        "Streamlit (generadas a partir de las tablas .csv del proyecto).",
    )
    filas_met = [[p.id, p.metrica, p.calculo, p.unidad] for p in M.PREGUNTAS]
    agregar_tabla(
        doc, ["#", "Metrica", "Calculo", "Unidad"], filas_met, 5,
        "Metricas definidas por pregunta de negocio.",
        "elaboracion propia.",
    )

    # ============================ 8. RESULTADOS Y DISCUSION ============================
    h1(doc, "8. Resultados y discusion")
    parrafo(
        doc,
        f"La tasa de cancelacion global es del {kpis['tasa_churn']:.2f}% "
        f"({kpis['churn_si']} clientes cancelan y {kpis['churn_no']} permanecen). A "
        f"continuacion se responde cada pregunta con su tabla, su figura y su "
        f"interpretacion. Las cifras coinciden con los archivos .csv y con la aplicacion.",
    )

    # Orden de presentacion y numeracion de tablas/figuras
    sec = [
        ("8.1 P1 - Cancelacion por tipo de contrato", "P1", "p1",
         ["contract", "clientes", "cancelaciones", "tasa_churn_%"]),
        ("8.2 P2 - Cancelacion por antiguedad", "P2", "p2",
         ["segmento_antiguedad", "clientes", "cancelaciones", "tasa_churn_%"]),
        ("8.3 P3a - Cancelacion por metodo de pago", "P3a", "p3a",
         ["payment_method", "clientes", "cancelaciones", "tasa_churn_%"]),
        ("8.4 P3b - Cancelacion por servicio de internet", "P3b", "p3b",
         ["internet_service", "clientes", "cancelaciones", "tasa_churn_%"]),
        ("8.5 P4 - Facturacion comparada entre grupos", "P4", "p4",
         ["churn", "clientes", "tenure_prom", "monthly_charges_prom",
          "total_charges_prom", "cargo_promedio_mensual_prom"]),
        ("8.6 P5 - Cancelacion por numero de servicios adicionales", "P5", "p5",
         ["num_servicios_adicionales", "clientes", "cancelaciones", "tasa_churn_%"]),
    ]
    num_tabla = 6   # ya se usaron 1..5
    num_fig = 1
    for titulo_sec, pid, tkey, enc in sec:
        preg = M.pregunta_por_id(pid)
        h2(doc, titulo_sec)
        tabla_desde_df(
            doc, tablas[tkey], num_tabla, f"{preg.metrica}.",
            f"elaboracion propia a partir de EA4/resultados/{M.TABLAS[tkey].name}.",
            encabezados=enc,
        )
        agregar_figura(
            doc, FIGURAS / preg.figura, num_fig,
            f"{preg.id} - {preg.titulo}.",
            "elaboracion propia (grafico identico al de la aplicacion de Streamlit).",
        )
        parrafo(doc, f"Que muestra: {preg.interpretacion}")
        parrafo(doc, f"Decision que apoya: {preg.decision}")
        parrafo(doc, f"Limitacion: {preg.limitacion}")
        num_tabla += 1
        num_fig += 1

    h2(doc, "8.7 Evidencia del resumen automatico")
    parrafo(
        doc,
        "El resumen descriptivo del DataFrame final se genera por codigo con la clase "
        "DatasetSummarizer y se guarda en EA4/resultados/resumen_dataframe.txt, "
        "regenerandose en cada ejecucion. A continuacion se transcribe.",
    )
    parrafo_mono(doc, resumen_txt)

    # ============================ 9. CONCLUSIONES ============================
    h1(doc, "9. Conclusiones y recomendaciones")
    h2(doc, "9.1 Conclusiones")
    for c in M.CONCLUSIONES:
        item(doc, c)
    h2(doc, "9.2 Recomendaciones")
    for r in M.RECOMENDACIONES:
        item(doc, r)
    h2(doc, "9.3 Limitaciones")
    for l in M.LIMITACIONES:
        item(doc, l)

    # ============================ 10. REFERENCIAS ============================
    h1(doc, "10. Referencias")
    referencias = [
        "Han, J., Kamber, M., & Pei, J. (2011). Data mining: Concepts and techniques "
        "(3rd ed.). Morgan Kaufmann.",
        "Lhoest, Q., et al. (2021). Datasets: A community library for natural language "
        "processing. En Proceedings of EMNLP 2021: System Demonstrations (pp. 175-184). ACL.",
        "McKinney, W. (2010). Data structures for statistical computing in Python. En "
        "Proceedings of the 9th Python in Science Conference (pp. 56-61).",
        "Provost, F., & Fawcett, T. (2013). Data science for business. O'Reilly Media.",
        "Streamlit Inc. (s. f.). Streamlit documentation. https://docs.streamlit.io",
        "Verbeke, W., Dejaeger, K., Martens, D., Hur, J., & Baesens, B. (2012). New "
        "insights into churn prediction in the telecommunication sector. European Journal "
        "of Operational Research, 218(1), 211-229.",
        f"scikit-learn. (s. f.). Churn prediction dataset [Conjunto de datos]. Hugging "
        f"Face. {P['dataset_url']}",
        "pandas development team. (s. f.). pandas documentation. "
        "https://pandas.pydata.org/docs/",
        f"Pedraza Hernandez, J. I. ({P['anio']}). Repositorio del proyecto {P['codigo']} "
        f"- Analitica de Datos [Repositorio]. GitHub. {P['repo']}",
    ]
    for ref in referencias:
        p = doc.add_paragraph(style="Normal")
        p.paragraph_format.left_indent = Pt(36)
        p.paragraph_format.first_line_indent = Pt(-36)
        p.add_run(ref)

    doc.save(str(SALIDA))
    print(f"Documento generado: {SALIDA}")
    print(f"  crudo {n_crudo_f}x{n_crudo_c} | final {n_f}x{n_c} | "
          f"churn {kpis['tasa_churn']:.2f}% | figuras {num_fig-1} | tablas {num_tabla-1}")


if __name__ == "__main__":
    construir()
