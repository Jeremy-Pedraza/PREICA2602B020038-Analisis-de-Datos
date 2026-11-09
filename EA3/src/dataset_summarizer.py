"""Resumen descriptivo de un ``DataFrame`` ya cargado.

Expone la clase :class:`DatasetSummarizer`, que recibe un ``DataFrame`` y
produce un resumen de texto. No se conecta a Hugging Face ni modifica los datos.
"""

from __future__ import annotations

import pathlib

import pandas as pd


class DatasetSummarizer:
    """Construye un resumen textual de un ``DataFrame``.

    Parameters
    ----------
    df:
        DataFrame a resumir.
    metadata:
        Metadatos opcionales (por ejemplo los de ``HFDatasetLoader.get_metadata``)
        usados en la sección de identificación.
    """

    def __init__(self, df: pd.DataFrame, metadata: dict | None = None) -> None:
        self.df = df
        self.metadata = metadata or {}

    def _seccion_identificacion(self) -> list[str]:
        m = self.metadata
        return [
            "1. IDENTIFICACION",
            "-" * 60,
            f"   dataset_id      : {m.get('dataset_id', 'N/D')}",
            f"   split           : {m.get('split', 'N/D')}",
            f"   configuracion   : {m.get('config', 'N/D')}",
            f"   revision        : {m.get('revision', 'N/D')}",
            f"   fecha_consulta  : {m.get('fecha_consulta', 'N/D')}",
            f"   es_muestra      : {m.get('es_muestra', 'N/D')}",
        ]

    def _seccion_dimensiones(self) -> list[str]:
        filas, columnas = self.df.shape
        return [
            "",
            "2. DIMENSIONES",
            "-" * 60,
            f"   filas    : {filas}",
            f"   columnas : {columnas}",
        ]

    def _seccion_columnas(self) -> list[str]:
        lineas = ["", "3. COLUMNAS Y TIPOS", "-" * 60]
        for nombre, tipo in self.df.dtypes.items():
            lineas.append(f"   {nombre:<30} {tipo}")
        return lineas

    def _seccion_nulos(self) -> list[str]:
        lineas = ["", "4. VALORES NULOS POR COLUMNA", "-" * 60]
        total = len(self.df)
        nulos = self.df.isna().sum()
        for nombre, cantidad in nulos.items():
            pct = (cantidad / total * 100) if total else 0.0
            lineas.append(f"   {nombre:<30} {int(cantidad):>8}  ({pct:5.2f}%)")
        return lineas

    def _seccion_duplicados(self) -> list[str]:
        dup = int(self.df.duplicated().sum())
        return ["", "5. FILAS DUPLICADAS", "-" * 60, f"   duplicados : {dup}"]

    def _seccion_numericas(self) -> list[str]:
        lineas = ["", "6. ESTADISTICAS DESCRIPTIVAS (NUMERICAS)", "-" * 60]
        numericas = self.df.select_dtypes(include="number")
        if numericas.empty:
            lineas.append("   (sin columnas numericas)")
        else:
            lineas.append(numericas.describe().to_string())
        return lineas

    def _seccion_categoricas(self) -> list[str]:
        lineas = ["", "7. FRECUENCIAS (CATEGORICAS, TOP 5)", "-" * 60]
        categoricas = self.df.select_dtypes(include=["object", "category", "bool"])
        if categoricas.empty:
            lineas.append("   (sin columnas categoricas)")
            return lineas
        for columna in categoricas.columns:
            lineas.append(f"   [{columna}]")
            top = self.df[columna].value_counts(dropna=False).head(5)
            for valor, cantidad in top.items():
                lineas.append(f"      {str(valor):<30} {int(cantidad)}")
        return lineas

    def build_summary(self) -> str:
        """Construye el texto completo del resumen.

        Returns
        -------
        str
            Resumen con las secciones de identificación, dimensiones, columnas,
            nulos, duplicados, estadísticas numéricas y frecuencias categóricas.
        """
        bloques: list[str] = ["=" * 60, "RESUMEN DEL DATASET", "=" * 60]
        bloques += self._seccion_identificacion()
        bloques += self._seccion_dimensiones()
        bloques += self._seccion_columnas()
        bloques += self._seccion_nulos()
        bloques += self._seccion_duplicados()
        bloques += self._seccion_numericas()
        bloques += self._seccion_categoricas()
        bloques.append("")
        return "\n".join(bloques)

    def save(self, path: str | pathlib.Path) -> pathlib.Path:
        """Guarda el resumen en ``path`` en UTF-8, sobrescribiendo.

        Crea la carpeta contenedora si no existe.

        Parameters
        ----------
        path:
            Ruta del archivo de salida.

        Returns
        -------
        pathlib.Path
            Ruta donde se escribió el resumen.
        """
        destino = pathlib.Path(path)
        destino.parent.mkdir(parents=True, exist_ok=True)
        destino.write_text(self.build_summary(), encoding="utf-8")
        return destino

    def __str__(self) -> str:
        return self.build_summary()
