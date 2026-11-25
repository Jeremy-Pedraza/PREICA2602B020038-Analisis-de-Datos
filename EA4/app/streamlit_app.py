# -*- coding: utf-8 -*-
"""Aplicacion Streamlit - EA4 (Metricas y Visualizacion de Datos).

Recorre los resultados del proyecto integrador sobre el abandono de clientes
(churn) en telecomunicaciones: presenta la necesidad, las preguntas de negocio
y la identificacion del dataset; una pagina por pregunta con su metrica, su
grafico interactivo y su interpretacion; y las conclusiones, recomendaciones y
limitaciones del analisis.

Los datos se cargan desde ``EA4/resultados/`` (datos preparados y tablas de
resultados del notebook). Las cifras y los graficos se generan a partir de esos
archivos, de modo que coinciden con el notebook y con el documento PDF.

Ejecutar en local (desde la raiz del repositorio):
    streamlit run EA4/app/streamlit_app.py
"""

from __future__ import annotations

import sys
import pathlib

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# Importar el modulo de metricas directamente (sin cargar el paquete src, que
# arrastra dependencias pesadas como `datasets`), portable a Streamlit Cloud.
RAIZ_EA4 = pathlib.Path(__file__).resolve().parent.parent
sys.path.insert(0, str(RAIZ_EA4 / "src"))
import metricas as M  # noqa: E402

st.set_page_config(
    page_title="Churn Telco - EA4",
    page_icon="📉",
    layout="wide",
    initial_sidebar_state="expanded",
)


# --------------------------------------------------------------------------
# Carga cacheada
# --------------------------------------------------------------------------
@st.cache_data
def _cargar():
    df = M.cargar_transformado()
    tablas = M.cargar_tablas()
    kpis = M.kpis_globales(df)
    resumen = M.leer_resumen()
    fecha = M.leer_fecha_consulta()
    return df, tablas, kpis, resumen, fecha


df, TABLAS, KPIS, RESUMEN, FECHA = _cargar()
P = M.PROYECTO


# --------------------------------------------------------------------------
# Graficos (Plotly) - misma logica que las figuras del PDF
# --------------------------------------------------------------------------
def fig_barras_tasa(preg: M.Pregunta) -> go.Figure:
    datos = TABLAS[preg.tabla].sort_values(preg.val_col, ascending=False).copy()
    datos[preg.cat_col] = datos[preg.cat_col].astype(str)
    maximo = datos[preg.val_col].max()
    colores = [M.COLOR_RIESGO if i == 0 else M.COLOR_OK
               for i in range(len(datos))]
    fig = go.Figure(go.Bar(
        x=datos[preg.cat_col],
        y=datos[preg.val_col],
        marker_color=colores,
        text=[f"{v:.1f}%" for v in datos[preg.val_col]],
        textposition="outside",
        customdata=datos[["clientes", "cancelaciones"]].values,
        hovertemplate=(
            "<b>%{x}</b><br>Tasa de churn: %{y:.2f}%<br>"
            "Clientes: %{customdata[0]:,}<br>"
            "Cancelaciones: %{customdata[1]:,}<extra></extra>"
        ),
    ))
    fig.update_layout(
        title=f"{preg.id} - {preg.titulo}",
        xaxis_title=preg.etiqueta_cat,
        yaxis_title=preg.etiqueta_val,
        yaxis_range=[0, maximo * 1.18],
        margin=dict(t=60, b=40, l=10, r=10),
        height=430,
    )
    return fig


def fig_comparacion_cargos(preg: M.Pregunta) -> go.Figure:
    datos = TABLAS[preg.tabla].copy()
    orden = {"No": 0, "Yes": 1}
    datos = datos.sort_values("churn", key=lambda s: s.map(orden))
    estados = ["Permanece (No)" if c == "No" else "Cancela (Yes)"
               for c in datos["churn"]]
    fig = go.Figure()
    fig.add_bar(x=estados, y=datos["monthly_charges_prom"],
                name="Cargo mensual promedio", marker_color=M.COLOR_RIESGO,
                text=[f"{v:.1f}" for v in datos["monthly_charges_prom"]],
                textposition="outside", yaxis="y1")
    fig.add_bar(x=estados, y=datos["tenure_prom"],
                name="Antiguedad promedio (meses)", marker_color=M.COLOR_OK,
                text=[f"{v:.1f}" for v in datos["tenure_prom"]],
                textposition="outside", yaxis="y2")
    fig.update_layout(
        title=f"{preg.id} - {preg.titulo}",
        barmode="group",
        yaxis=dict(title="Cargo mensual promedio", side="left"),
        yaxis2=dict(title="Antiguedad promedio (meses)", overlaying="y",
                    side="right", showgrid=False),
        legend=dict(orientation="h", yanchor="bottom", y=-0.25,
                    xanchor="center", x=0.5),
        margin=dict(t=60, b=40, l=10, r=10),
        height=430,
    )
    return fig


def figura_de(preg: M.Pregunta) -> go.Figure:
    if preg.tipo == "comparacion_cargos":
        return fig_comparacion_cargos(preg)
    return fig_barras_tasa(preg)


def bloque_pregunta(preg: M.Pregunta) -> None:
    st.subheader(f"{preg.id}. {preg.pregunta}")
    izq, der = st.columns([3, 2], gap="large")
    with izq:
        st.plotly_chart(figura_de(preg), use_container_width=True)
    with der:
        st.markdown(f"**Metrica:** {preg.metrica}")
        st.markdown(f"**Calculo:** {preg.calculo}")
        st.markdown(f"**Unidad:** {preg.unidad}")
        st.markdown(f"**Que muestra:** {preg.interpretacion}")
        st.success(f"**Decision que apoya:** {preg.decision}")
        st.caption(f"⚠️ Limitacion: {preg.limitacion}")
    with st.expander("Ver tabla de datos (fuente del grafico)"):
        st.dataframe(TABLAS[preg.tabla], use_container_width=True, hide_index=True)


# --------------------------------------------------------------------------
# Barra lateral / navegacion
# --------------------------------------------------------------------------
st.sidebar.title("📉 Churn Telco - EA4")
st.sidebar.caption("Proyecto Integrado III - Analitica de Datos")
seccion = st.sidebar.radio(
    "Secciones",
    ["Inicio y necesidad",
     "P1 - Contrato",
     "P2 - Antiguedad",
     "P3 - Pago e internet",
     "P4 - Facturacion",
     "P5 - Servicios adicionales",
     "Conclusiones",
     "Datos y resumen"],
)
st.sidebar.divider()
st.sidebar.markdown(f"**Integrante:** {P['integrante']}")
st.sidebar.markdown(f"**{P['grupo']}** · Docente: {P['docente']}")
st.sidebar.markdown(f"[Repositorio del proyecto]({P['repo']})")
st.sidebar.markdown(f"[Dataset en Hugging Face]({P['dataset_url']})")


# --------------------------------------------------------------------------
# Secciones
# --------------------------------------------------------------------------
if seccion == "Inicio y necesidad":
    st.title("Analitica del abandono de clientes (churn)")
    st.markdown(f"**{P['asignatura']}** · {P['codigo']}")
    st.markdown("### Necesidad del proyecto")
    st.write(M.NECESIDAD)

    st.markdown("### Indicadores de la base de clientes")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Clientes", f"{KPIS['clientes']:,}")
    c2.metric("Tasa de churn global", f"{KPIS['tasa_churn']:.2f}%",
              help="Clientes que cancelaron sobre el total")
    c3.metric("Clientes que cancelan", f"{KPIS['churn_si']:,}")
    c4.metric("Cargo mensual promedio", f"{KPIS['monthly_prom']:.1f}")

    c5, c6, c7 = st.columns(3)
    c5.metric("Ingreso mensual total", f"{KPIS['ingreso_mensual_total']:,.0f}")
    c6.metric("Ingreso mensual en riesgo", f"{KPIS['ingreso_en_riesgo']:,.0f}",
              delta=f"-{KPIS['pct_ingreso_riesgo']:.1f}% del total",
              delta_color="inverse",
              help="Facturacion mensual de los clientes que cancelan")
    c7.metric("Antiguedad promedio", f"{KPIS['tenure_prom']:.1f} meses")

    st.markdown("### Preguntas de negocio")
    tabla_preg = pd.DataFrame(
        [{"#": p.id, "Pregunta": p.pregunta, "Metrica": p.metrica}
         for p in M.PREGUNTAS]
    )
    st.dataframe(tabla_preg, use_container_width=True, hide_index=True)

    st.markdown("### Dataset")
    st.markdown(
        f"- **Fuente:** `{P['dataset_id']}` (Telco Customer Churn) — "
        f"[Hugging Face]({P['dataset_url']})\n"
        f"- **Registros:** {KPIS['clientes']:,} clientes · "
        f"**Variables:** {KPIS['variables']} (tras enriquecimiento)\n"
        f"- **Consulta:** {FECHA}\n"
        f"- **Variable objetivo:** `churn` (cancelo el servicio: Si / No)"
    )
    st.info(
        "Use el menu lateral para recorrer cada pregunta de negocio con su "
        "metrica, grafico interactivo e interpretacion."
    )

elif seccion == "P1 - Contrato":
    bloque_pregunta(M.pregunta_por_id("P1"))

elif seccion == "P2 - Antiguedad":
    bloque_pregunta(M.pregunta_por_id("P2"))

elif seccion == "P3 - Pago e internet":
    st.title("P3 - Metodo de pago y servicio de internet")
    bloque_pregunta(M.pregunta_por_id("P3a"))
    st.divider()
    bloque_pregunta(M.pregunta_por_id("P3b"))

elif seccion == "P4 - Facturacion":
    bloque_pregunta(M.pregunta_por_id("P4"))

elif seccion == "P5 - Servicios adicionales":
    bloque_pregunta(M.pregunta_por_id("P5"))

elif seccion == "Conclusiones":
    st.title("Conclusiones, recomendaciones y limitaciones")
    st.markdown("### Conclusiones")
    for c in M.CONCLUSIONES:
        st.markdown(f"- {c}")
    st.markdown("### Recomendaciones")
    for r in M.RECOMENDACIONES:
        st.markdown(f"- {r}")
    st.markdown("### Limitaciones del analisis")
    for l in M.LIMITACIONES:
        st.markdown(f"- {l}")
    st.info(
        "Distincion clave: el analisis describe **ingresos** facturados y "
        "**asociaciones** descriptivas, no **rentabilidad** ni relaciones causales."
    )

elif seccion == "Datos y resumen":
    st.title("Datos preparados y resumen automatico")
    st.markdown("### DataFrame preparado y enriquecido")
    st.caption(f"{df.shape[0]:,} filas × {df.shape[1]} columnas "
               "(EA4/resultados/churn_transformado.csv)")
    st.dataframe(df.head(50), use_container_width=True)
    st.markdown("### Resumen descriptivo automatico (DatasetSummarizer)")
    st.code(RESUMEN, language="text")
