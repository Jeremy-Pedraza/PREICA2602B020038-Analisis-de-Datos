# -*- coding: utf-8 -*-
"""Genera las figuras (PNG) de EA4 a partir de las tablas de resultados.

Las mismas cifras alimentan la aplicacion de Streamlit (graficos interactivos
con Plotly) y el documento PDF (estas imagenes estaticas con Matplotlib), de
modo que ambos medios muestran resultados identicos.

Salida: ``EA4/resultados/figuras/fig_p1.png`` ... ``fig_p5.png``.

Uso:
    python EA4/src/generar_figuras.py
"""

from __future__ import annotations

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt

try:  # permitir ejecucion como script o como modulo
    from . import metricas as M
except ImportError:  # pragma: no cover
    import metricas as M

plt.rcParams.update({
    "font.size": 11,
    "axes.titlesize": 13,
    "axes.titleweight": "bold",
    "axes.spines.top": False,
    "axes.spines.right": False,
    "figure.dpi": 150,
})


def _barras_tasa(p: M.Pregunta, df):
    datos = df.sort_values(p.val_col, ascending=False)
    cats = datos[p.cat_col].astype(str).tolist()
    vals = datos[p.val_col].tolist()
    colores = [M.COLOR_RIESGO if i == 0 else M.COLOR_OK for i in range(len(vals))]

    fig, ax = plt.subplots(figsize=(7.2, 4.2))
    barras = ax.bar(cats, vals, color=colores)
    ax.set_title(f"{p.id} - {p.titulo}")
    ax.set_xlabel(p.etiqueta_cat)
    ax.set_ylabel(p.etiqueta_val)
    ax.set_ylim(0, max(vals) * 1.18)
    for b, v in zip(barras, vals):
        ax.text(b.get_x() + b.get_width() / 2, v + max(vals) * 0.02,
                f"{v:.1f}%", ha="center", va="bottom", fontweight="bold")
    if max(len(c) for c in cats) > 10:
        plt.setp(ax.get_xticklabels(), rotation=20, ha="right")
    fig.tight_layout()
    return fig


def _comparacion_cargos(p: M.Pregunta, df):
    """P4: barras agrupadas de cargo mensual promedio y antiguedad promedio
    por estado de cancelacion, con doble eje."""
    orden = {"No": 0, "Yes": 1}
    datos = df.sort_values("churn", key=lambda s: s.map(orden))
    estados = ["Permanece (No)" if c == "No" else "Cancela (Yes)"
               for c in datos["churn"]]
    cargo = datos["monthly_charges_prom"].tolist()
    tenure = datos["tenure_prom"].tolist()

    fig, ax1 = plt.subplots(figsize=(7.2, 4.2))
    x = range(len(estados))
    w = 0.38
    b1 = ax1.bar([i - w / 2 for i in x], cargo, width=w,
                 color=M.COLOR_RIESGO, label="Cargo mensual promedio")
    ax1.set_ylabel("Cargo mensual promedio", color=M.COLOR_RIESGO)
    ax1.tick_params(axis="y", labelcolor=M.COLOR_RIESGO)
    ax1.set_xticks(list(x))
    ax1.set_xticklabels(estados)
    ax1.set_ylim(0, max(cargo) * 1.25)

    ax2 = ax1.twinx()
    ax2.spines["top"].set_visible(False)
    b2 = ax2.bar([i + w / 2 for i in x], tenure, width=w,
                 color=M.COLOR_OK, label="Antiguedad promedio (meses)")
    ax2.set_ylabel("Antiguedad promedio (meses)", color=M.COLOR_OK)
    ax2.tick_params(axis="y", labelcolor=M.COLOR_OK)
    ax2.set_ylim(0, max(tenure) * 1.25)

    for b, v in zip(b1, cargo):
        ax1.text(b.get_x() + b.get_width() / 2, v + max(cargo) * 0.02,
                 f"{v:.1f}", ha="center", va="bottom", fontweight="bold",
                 color=M.COLOR_RIESGO)
    for b, v in zip(b2, tenure):
        ax2.text(b.get_x() + b.get_width() / 2, v + max(tenure) * 0.02,
                 f"{v:.1f}", ha="center", va="bottom", fontweight="bold",
                 color=M.COLOR_OK)

    ax1.set_title(f"{p.id} - {p.titulo}")
    fig.legend(loc="upper center", ncol=2, frameon=False,
               bbox_to_anchor=(0.5, 0.0))
    fig.tight_layout(rect=(0, 0.06, 1, 1))
    return fig


def construir() -> list[str]:
    M.FIGURAS.mkdir(parents=True, exist_ok=True)
    tablas = M.cargar_tablas()
    generadas = []
    for p in M.PREGUNTAS:
        df = tablas[p.tabla]
        if p.tipo == "comparacion_cargos":
            fig = _comparacion_cargos(p, df)
        else:
            fig = _barras_tasa(p, df)
        destino = M.FIGURAS / p.figura
        fig.savefig(destino, bbox_inches="tight")
        plt.close(fig)
        generadas.append(str(destino))
        print(f"  figura: {destino.name}")
    return generadas


if __name__ == "__main__":
    print("Generando figuras EA4...")
    construir()
    print("Listo.")
