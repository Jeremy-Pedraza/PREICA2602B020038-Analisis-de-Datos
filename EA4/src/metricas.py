# -*- coding: utf-8 -*-
"""Fuente unica de metricas y metadatos de las preguntas de negocio (EA4).

Centraliza la carga de los datos preparados y las tablas de resultados del
proyecto, el calculo de los indicadores clave (KPI) y la descripcion de cada
pregunta de negocio (metrica, calculo, unidad, grafico, interpretacion,
decision y limitacion).

REGLA DURA: ninguna cifra se escribe a mano. Todos los indicadores se calculan
en tiempo de ejecucion desde ``EA4/resultados/`` para que la aplicacion de
Streamlit, las figuras del documento y el PDF muestren exactamente las mismas
cifras que el notebook del proyecto.

Lo usan: la app de Streamlit (``EA4/app/streamlit_app.py``), el generador de
figuras (``EA4/src/generar_figuras.py``) y el generador del documento
(``EA4/src/generar_documento.py``).
"""

from __future__ import annotations

import pathlib
from dataclasses import dataclass, field

import pandas as pd

# --------------------------------------------------------------------------
# Rutas (relativas al paquete, portables a Streamlit Cloud)
# --------------------------------------------------------------------------
EA4 = pathlib.Path(__file__).resolve().parent.parent
RESULTADOS = EA4 / "resultados"
FIGURAS = RESULTADOS / "figuras"

CSV_TRANS = RESULTADOS / "churn_transformado.csv"
CSV_CRUDO = RESULTADOS / "churn_crudo.csv"
TXT_RESUMEN = RESULTADOS / "resumen_dataframe.txt"

TABLAS = {
    "p1": RESULTADOS / "tabla_p1_churn_por_contrato.csv",
    "p2": RESULTADOS / "tabla_p2_churn_por_antiguedad.csv",
    "p3a": RESULTADOS / "tabla_p3a_churn_por_metodo_pago.csv",
    "p3b": RESULTADOS / "tabla_p3b_churn_por_internet.csv",
    "p4": RESULTADOS / "tabla_p4_cargos_por_churn.csv",
    "p5": RESULTADOS / "tabla_p5_churn_por_num_servicios.csv",
}

# Paleta coherente (accesible en claro). Rojo = riesgo, azul = proteccion.
COLOR_RIESGO = "#D7263D"
COLOR_OK = "#1B5299"
COLOR_NEUTRO = "#7A7A7A"
SECUENCIA = ["#1B5299", "#2E86AB", "#6FB1D6", "#D7263D", "#F46036", "#E4A010", "#4C9F70"]

# --------------------------------------------------------------------------
# Contexto institucional / del proyecto
# --------------------------------------------------------------------------
PROYECTO = {
    "titulo": "Analitica del abandono de clientes (churn) en telecomunicaciones",
    "integrante": "Jeremy Ivan Pedraza Hernandez",
    "grupo": "AE_Grupo 10",
    "docente": "Andres Felipe Callejas Jaramillo",
    "asignatura": "Proyecto Integrado III - Analitica de Datos",
    "codigo": "PREICA2602B020038",
    "programa": "Ingenieria de Software y Datos",
    "facultad": "Facultad de Ingenieria y Ciencias Agropecuarias",
    "institucion": "Institucion Universitaria Digital de Antioquia",
    "anio": "2026",
    "repo": "https://github.com/Jeremy-Pedraza/PREICA2602B020038-Analisis-de-Datos",
    "dataset_id": "scikit-learn/churn-prediction",
    "dataset_url": "https://huggingface.co/datasets/scikit-learn/churn-prediction",
}

NECESIDAD = (
    "Una compania de telecomunicaciones necesita anticipar el abandono de clientes "
    "(churn). Cada registro es un cliente; a partir de su antiguedad, tipo de contrato, "
    "servicios contratados y facturacion se busca identificar los perfiles asociados al "
    "abandono para dirigir, de forma eficiente, las acciones de retencion."
)


# --------------------------------------------------------------------------
# Carga de datos
# --------------------------------------------------------------------------
def cargar_transformado() -> pd.DataFrame:
    return pd.read_csv(CSV_TRANS)


def cargar_tablas() -> dict[str, pd.DataFrame]:
    return {k: pd.read_csv(v) for k, v in TABLAS.items()}


def leer_resumen() -> str:
    return TXT_RESUMEN.read_text(encoding="utf-8")


def leer_fecha_consulta() -> str:
    for linea in leer_resumen().splitlines():
        if "fecha_consulta" in linea:
            return linea.split(":", 1)[1].strip()
    return "n/d"


# --------------------------------------------------------------------------
# Indicadores globales (KPI)
# --------------------------------------------------------------------------
def kpis_globales(df: pd.DataFrame | None = None) -> dict:
    """Indicadores que describen la necesidad: tamano de la base, tasa de
    churn global e ingreso mensual recurrente en riesgo."""
    if df is None:
        df = cargar_transformado()
    n = len(df)
    churn_si = int((df["churn"] == "Yes").sum())
    churn_no = n - churn_si
    tasa = churn_si / n * 100 if n else 0.0
    ingreso_mensual_total = float(df["monthly_charges"].sum())
    ingreso_en_riesgo = float(df.loc[df["churn"] == "Yes", "monthly_charges"].sum())
    pct_ingreso_riesgo = ingreso_en_riesgo / ingreso_mensual_total * 100 if ingreso_mensual_total else 0.0
    return {
        "clientes": n,
        "variables": df.shape[1],
        "churn_si": churn_si,
        "churn_no": churn_no,
        "tasa_churn": tasa,
        "monthly_prom": float(df["monthly_charges"].mean()),
        "tenure_prom": float(df["tenure"].mean()),
        "ingreso_mensual_total": ingreso_mensual_total,
        "ingreso_en_riesgo": ingreso_en_riesgo,
        "pct_ingreso_riesgo": pct_ingreso_riesgo,
    }


# --------------------------------------------------------------------------
# Metadatos de las preguntas de negocio
# --------------------------------------------------------------------------
@dataclass
class Pregunta:
    id: str
    titulo: str
    pregunta: str
    metrica: str
    calculo: str
    unidad: str
    tabla: str            # clave en TABLAS / cargar_tablas()
    cat_col: str          # columna categorica (eje X)
    val_col: str          # columna de valor (eje Y)
    tipo: str             # "barras_tasa" | "comparacion_cargos"
    interpretacion: str
    decision: str
    limitacion: str
    figura: str = ""      # nombre del PNG en resultados/figuras
    etiqueta_cat: str = ""
    etiqueta_val: str = ""

    def __post_init__(self):
        if not self.figura:
            self.figura = f"fig_{self.id.lower()}.png"


PREGUNTAS: list[Pregunta] = [
    Pregunta(
        id="P1",
        titulo="Cancelacion por tipo de contrato",
        pregunta="Que tipo de contrato concentra mayor tasa de cancelacion?",
        metrica="Tasa de churn por tipo de contrato",
        calculo="cancelaciones / clientes * 100, agrupando por contract",
        unidad="% de clientes",
        tabla="p1",
        cat_col="contract",
        val_col="tasa_churn_%",
        tipo="barras_tasa",
        etiqueta_cat="Tipo de contrato",
        etiqueta_val="Tasa de churn (%)",
        interpretacion=(
            "El contrato mes a mes concentra la mayor tasa de cancelacion, muy por "
            "encima de los contratos a uno y dos anos. La permanencia contractual actua "
            "como factor protector frente al abandono."
        ),
        decision=(
            "Incentivar la migracion de clientes mes a mes hacia contratos anuales o "
            "bianuales (descuentos o beneficios por permanencia)."
        ),
        limitacion=(
            "Es una asociacion descriptiva: el tipo de contrato se correlaciona con el "
            "churn, pero no prueba por si solo una relacion causal."
        ),
    ),
    Pregunta(
        id="P2",
        titulo="Cancelacion por antiguedad",
        pregunta="Como varia la cancelacion segun la antiguedad del cliente?",
        metrica="Tasa de churn por segmento de antiguedad",
        calculo="cancelaciones / clientes * 100, agrupando por segmento_antiguedad (meses)",
        unidad="% de clientes",
        tabla="p2",
        cat_col="segmento_antiguedad",
        val_col="tasa_churn_%",
        tipo="barras_tasa",
        etiqueta_cat="Segmento de antiguedad (meses)",
        etiqueta_val="Tasa de churn (%)",
        interpretacion=(
            "El abandono es mas alto entre los clientes nuevos (0-12 meses) y decrece de "
            "forma sostenida a medida que aumenta la antiguedad. Los primeros meses son "
            "el periodo de mayor riesgo."
        ),
        decision=(
            "Reforzar el acompanamiento (onboarding) y las ofertas de fidelizacion "
            "durante el primer ano de relacion con el cliente."
        ),
        limitacion=(
            "Sin marca temporal no se puede analizar la evolucion del churn en el tiempo; "
            "la antiguedad es una foto, no una serie historica."
        ),
    ),
    Pregunta(
        id="P3a",
        titulo="Cancelacion por metodo de pago",
        pregunta="Que metodo de pago se asocia con mas churn?",
        metrica="Tasa de churn por metodo de pago",
        calculo="cancelaciones / clientes * 100, agrupando por payment_method",
        unidad="% de clientes",
        tabla="p3a",
        cat_col="payment_method",
        val_col="tasa_churn_%",
        tipo="barras_tasa",
        etiqueta_cat="Metodo de pago",
        etiqueta_val="Tasa de churn (%)",
        interpretacion=(
            "El pago con cheque electronico (electronic check) presenta la tasa de "
            "cancelacion mas alta, muy por encima de los pagos automaticos con tarjeta o "
            "transferencia."
        ),
        decision=(
            "Promover la domiciliacion del pago (debito automatico o tarjeta) en el "
            "segmento de cheque electronico y revisar la friccion de ese medio de pago."
        ),
        limitacion=(
            "El metodo de pago puede ser un indicador (proxy) del perfil del cliente, no "
            "necesariamente la causa directa del abandono."
        ),
    ),
    Pregunta(
        id="P3b",
        titulo="Cancelacion por servicio de internet",
        pregunta="Que servicio de internet se asocia con mas churn?",
        metrica="Tasa de churn por tipo de servicio de internet",
        calculo="cancelaciones / clientes * 100, agrupando por internet_service",
        unidad="% de clientes",
        tabla="p3b",
        cat_col="internet_service",
        val_col="tasa_churn_%",
        tipo="barras_tasa",
        etiqueta_cat="Servicio de internet",
        etiqueta_val="Tasa de churn (%)",
        interpretacion=(
            "El internet de fibra optica (fiber optic) muestra la mayor tasa de "
            "cancelacion, por encima de DSL y de quienes no tienen internet. Sugiere "
            "revisar la relacion calidad/precio de ese servicio."
        ),
        decision=(
            "Auditar la experiencia del servicio de fibra (calidad, precio, soporte) en "
            "el segmento de mayor abandono."
        ),
        limitacion=(
            "El dataset no incluye indicadores de calidad del servicio (cortes, "
            "velocidad real), por lo que no puede atribuirse la causa exacta."
        ),
    ),
    Pregunta(
        id="P4",
        titulo="Facturacion comparada entre grupos",
        pregunta="Como difieren los cargos entre quienes cancelan y permanecen?",
        metrica="Cargo mensual y antiguedad promedio por estado de cancelacion",
        calculo="promedio de monthly_charges y tenure, agrupando por churn",
        unidad="unidad monetaria (cargo) y meses (antiguedad)",
        tabla="p4",
        cat_col="churn",
        val_col="monthly_charges_prom",
        tipo="comparacion_cargos",
        etiqueta_cat="Estado del cliente",
        etiqueta_val="Cargo mensual promedio",
        interpretacion=(
            "Quienes cancelan tienen menor antiguedad promedio y, a la vez, un cargo "
            "mensual promedio mas alto que quienes permanecen. El cliente que abandona "
            "paga mas por mes pero lleva menos tiempo."
        ),
        decision=(
            "Vigilar a los clientes nuevos con cargo mensual alto (segmento de doble "
            "riesgo) con ofertas de valor o ajustes de plan."
        ),
        limitacion=(
            "Son ingresos facturados, no rentabilidad: sin datos de costos no puede "
            "afirmarse que un grupo sea mas o menos rentable."
        ),
    ),
    Pregunta(
        id="P5",
        titulo="Cancelacion por numero de servicios adicionales",
        pregunta="Influye la cantidad de servicios adicionales en la cancelacion?",
        metrica="Tasa de churn por numero de servicios adicionales",
        calculo="cancelaciones / clientes * 100, agrupando por num_servicios_adicionales",
        unidad="% de clientes",
        tabla="p5",
        cat_col="num_servicios_adicionales",
        val_col="tasa_churn_%",
        tipo="barras_tasa",
        etiqueta_cat="Numero de servicios adicionales",
        etiqueta_val="Tasa de churn (%)",
        interpretacion=(
            "Entre los clientes con internet, a mayor numero de servicios adicionales "
            "contratados menor es la tasa de cancelacion: mas servicios implican mayor "
            "vinculacion. El grupo con 0 servicios son en su mayoria clientes sin "
            "internet y debe leerse aparte."
        ),
        decision=(
            "Impulsar la adopcion de servicios adicionales (seguridad, soporte, backup) "
            "como estrategia de fidelizacion."
        ),
        limitacion=(
            "El grupo de 0 servicios mezcla clientes sin internet con clientes de solo "
            "conexion, por lo que no sigue la tendencia monotona."
        ),
    ),
]


def pregunta_por_id(pid: str) -> Pregunta:
    for p in PREGUNTAS:
        if p.id == pid:
            return p
    raise KeyError(pid)


# --------------------------------------------------------------------------
# Conclusiones, recomendaciones y limitaciones (texto compartido)
# --------------------------------------------------------------------------
CONCLUSIONES = [
    "El perfil de mayor riesgo es consistente entre las cinco preguntas: cliente "
    "reciente, con contrato mes a mes, pago por cheque electronico, internet de fibra, "
    "cargo mensual alto y pocos servicios adicionales.",
    "La permanencia contractual y la antiguedad son los factores protectores mas "
    "claros frente al abandono.",
    "Las metricas y los graficos de la aplicacion coinciden con las tablas .csv "
    "exportadas por el notebook, garantizando la trazabilidad de las cifras.",
]

RECOMENDACIONES = [
    "Priorizar la retencion en el primer ano de relacion con el cliente.",
    "Migrar clientes mes a mes a contratos de mayor permanencia.",
    "Promover pagos automaticos y la adopcion de servicios adicionales.",
    "Auditar la relacion calidad/precio del servicio de fibra optica.",
]

LIMITACIONES = [
    "Las relaciones son asociaciones descriptivas, no causales.",
    "Se analizan ingresos facturados, no rentabilidad (no hay datos de costos).",
    "Sin marca temporal no es posible el analisis de tendencia en el tiempo.",
    "Los resultados corresponden al split train del dataset y no se extrapolan a otro "
    "mercado.",
]
