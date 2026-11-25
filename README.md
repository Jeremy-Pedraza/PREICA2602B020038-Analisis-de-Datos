# PREICA2602B020038 - Analítica de Datos

Repositorio de las entregas **EA1** (carga y creación de dataset), **AE2** (primera etapa del proyecto integrador), **EA3** (transformación y enriquecimiento) y **EA4** (métricas, visualización y aplicación de Streamlit — entrega final) de la asignatura Proyecto Integrado III – Analítica de Datos.

> **Aplicación Streamlit (EA4):** 🔗 **[https://bmt9gfifmznktubbdfv9k3.streamlit.app](https://bmt9gfifmznktubbdfv9k3.streamlit.app)** · ejecución local: `streamlit run EA4/app/streamlit_app.py` (ver sección [Métricas y visualización (EA4)](#métricas-y-visualización-ea4)).

## Propósito

Cargar datasets públicos desde Hugging Face Hub, aplicarles transformaciones de preparación y generar un resumen descriptivo automático en texto, de forma reproducible y autocontenida por entrega.

## Asignatura, docente e integrante

- **Asignatura:** Proyecto Integrado III – Analítica de Datos
- **Código:** PREICA2602B020038
- **Docente:** Andrés Felipe Callejas Jaramillo
- **Integrante:** Jeremy Iván Pedraza Hernández (Grupo 1)

## Necesidad analítica (AE2)

Una compañía de telecomunicaciones necesita anticipar el **abandono de clientes** (_churn_). Cada registro del dataset es un cliente; a partir de su antigüedad, tipo de contrato, servicios contratados y facturación se busca identificar los factores asociados al abandono para diseñar acciones de retención. La primera etapa (AE2) cubre la carga, exploración descriptiva y preparación del dataset.

## Transformación y enriquecimiento (EA3)

La segunda etapa (EA3) continúa con el mismo dataset y repositorio. Transforma y enriquece los datos para responder, de forma descriptiva, cinco **preguntas de negocio** derivadas de la necesidad:

| #  | Pregunta de negocio                                                        | Tabla de resultados                        |
| -- | ------------------------------------------------------------------------- | ------------------------------------------ |
| P1 | ¿Qué tipo de contrato concentra mayor tasa de cancelación?                | `tabla_p1_churn_por_contrato.csv`          |
| P2 | ¿Cómo varía la cancelación según la antigüedad del cliente?               | `tabla_p2_churn_por_antiguedad.csv`        |
| P3 | ¿Qué método de pago y servicio de internet se asocian con más churn?      | `tabla_p3a_*.csv`, `tabla_p3b_*.csv`       |
| P4 | ¿Cómo difieren los cargos entre quienes cancelan y permanecen?            | `tabla_p4_cargos_por_churn.csv`            |
| P5 | ¿Influye la cantidad de servicios adicionales en la cancelación?          | `tabla_p5_churn_por_num_servicios.csv`     |

**Transformaciones (validadas con conteo de registros afectados):** T1 renombrado a `snake_case`; T2 conversión de `total_charges` a numérico (`pd.to_numeric(errors='coerce')`); T3 imputación de los 11 nulos de `total_charges` (clientes con `tenure = 0`) con `0.0`. Sin duplicados; sin nulos tras el tratamiento. El notebook y el documento incluyen una **comparación antes/después** (filas, columnas, tipos, nulos y duplicados del DataFrame crudo frente al preparado).

**Enriquecimiento (variables derivadas, sin fuentes externas):** `cargo_promedio_mensual` y `segmento_antiguedad` (de AE2); `churn_flag` (objetivo numérico), `num_servicios_adicionales` (conteo de 6 add-ons) y `grupo_cargo_mensual` (nivel de gasto). DataFrame final: 7043 filas × 26 columnas.

## Métricas y visualización (EA4)

La etapa final (EA4) responde las mismas cinco preguntas de negocio mediante **métricas y gráficos** y los presenta en una **aplicación de Streamlit** vinculada a este repositorio. Para cada pregunta se define la métrica (cálculo y unidad), se elige un gráfico de comparación y se acompaña de una interpretación, la decisión que apoya y sus limitaciones.

| #   | Métrica                                                   | Gráfico                       |
| --- | -------------------------------------------------------- | ----------------------------- |
| P1  | Tasa de churn por tipo de contrato (%)                   | Barras por contrato           |
| P2  | Tasa de churn por segmento de antigüedad (%)             | Barras por segmento           |
| P3a | Tasa de churn por método de pago (%)                     | Barras por método de pago     |
| P3b | Tasa de churn por servicio de internet (%)               | Barras por servicio           |
| P4  | Cargo mensual y antigüedad promedio por estado de churn  | Barras agrupadas (doble eje)  |
| P5  | Tasa de churn por número de servicios adicionales (%)    | Barras por nº de servicios    |

**Indicadores globales (necesidad):** 7043 clientes · tasa de churn global **26,54 %** · ~**30,5 %** del ingreso mensual recurrente concentrado en clientes que cancelan.

**Fuente única de cifras:** el módulo `EA4/src/metricas.py` carga los datos y las tablas de resultados y calcula todos los indicadores en tiempo de ejecución. La app (Plotly, interactivo) y el documento PDF (figuras Matplotlib en `EA4/resultados/figuras/`) usan ese módulo, por lo que **muestran cifras idénticas** al notebook.

**Documento completo (PDF):** `EA4/documento/Pedraza_Jeremy_AE4_Analitica_de_Datos.pdf` integra todas las etapas del periodo (portada, índices de contenido/tablas/figuras, introducción, necesidad, justificación, objetivos, marco teórico, metodología, transformación y enriquecimiento, métricas y visualizaciones, resultados y discusión, conclusiones y referencias).

**Despliegue en Streamlit Community Cloud:** la app está publicada en **[https://bmt9gfifmznktubbdfv9k3.streamlit.app](https://bmt9gfifmznktubbdfv9k3.streamlit.app)** (repositorio, rama `main`, archivo principal `EA4/app/streamlit_app.py`; dependencias desde `requirements.txt` de la raíz). Es pública: cualquier persona con el enlace puede verla, sin iniciar sesión.

## Datasets usados

| Entrega | Dataset                         | Split   | Enlace                                                        |
| ------- | ------------------------------- | ------- | ------------------------------------------------------------- |
| EA1     | `scikit-learn/iris`             | `train` | https://huggingface.co/datasets/scikit-learn/iris             |
| AE2     | `scikit-learn/churn-prediction` | `train` | https://huggingface.co/datasets/scikit-learn/churn-prediction |
| EA3     | `scikit-learn/churn-prediction` | `train` | https://huggingface.co/datasets/scikit-learn/churn-prediction |

## Estructura de carpetas

```
.
├── README.md
├── requirements.txt
├── .gitignore
├── EA1/
│   ├── src/                 # HFDatasetLoader + DatasetSummarizer
│   ├── notebooks/           # ea1_carga_y_transformaciones.ipynb
│   └── resultados/          # iris_crudo.csv, iris_transformado.csv, resumen_dataset.txt
├── AE2/
│   ├── documento/           # plantilla, generar vía script → .docx + .pdf
│   ├── src/                 # copias idénticas de EA1/src + generar_documento.py
│   ├── notebooks/           # ae2_exploracion_descriptiva.ipynb
│   └── resultados/          # churn_crudo.csv, churn_transformado.csv, resumen_dataframe.txt
├── EA3/
│   ├── documento/           # plantilla + Pedraza_Jeremy_AE3_*.docx / .pdf
│   ├── src/                 # copias de las clases + generar_documento.py (AE3)
│   ├── notebooks/           # ea3_transformacion_enriquecimiento.ipynb
│   └── resultados/          # churn_crudo.csv, churn_transformado.csv (enriquecido),
│                            # tabla_p1..p5 *.csv, resumen_dataframe.txt
└── EA4/
    ├── app/                 # streamlit_app.py (aplicación de métricas y gráficos)
    ├── src/                 # clases + metricas.py (fuente única) +
    │                        # generar_figuras.py + generar_documento.py (doc completo)
    ├── documento/           # plantilla + Pedraza_Jeremy_AE4_*.docx / .pdf (informe final)
    └── resultados/          # datos preparados, tabla_p1..p5 *.csv, resumen_dataframe.txt
                             # y figuras/ (PNG de los gráficos del documento)
```

## Pasos de ejecución

1. **Crear entorno virtual:**
   ```bash
   python -m venv .venv
   # Windows
   .venv\Scripts\activate
   # Linux/Mac
   source .venv/bin/activate
   ```
2. **Instalar dependencias:**
   ```bash
   pip install -r requirements.txt
   ```
3. **Ejecutar los notebooks** (desde la raíz del repositorio):
   ```bash
   jupyter nbconvert --to notebook --execute --inplace EA1/notebooks/ea1_carga_y_transformaciones.ipynb
   jupyter nbconvert --to notebook --execute --inplace AE2/notebooks/ae2_exploracion_descriptiva.ipynb
   jupyter nbconvert --to notebook --execute --inplace EA3/notebooks/ea3_transformacion_enriquecimiento.ipynb
   ```
   O abrirlos con `jupyter notebook` y ejecutar todas las celdas.
4. **Regenerar el documento Word de EA3** (lee las cifras de `EA3/resultados/`):
   ```bash
   python EA3/src/generar_documento.py
   ```
   El `.docx` usa campos TOC para el índice de contenido y la **lista de tablas** (campos `SEQ Tabla`); al abrirlo en Word, actualizar los índices con `Ctrl+A` y `F9`.

   > **Rutas relativas:** el notebook de EA3 se ejecuta desde la raíz del repositorio y sus salidas muestran rutas relativas (p. ej. `EA3/resultados/churn_transformado.csv`), no rutas absolutas del equipo.

   **Google Colab:** los notebooks de AE2 y EA3 también corren en Colab sin configuración manual. La primera celda detecta el entorno: si no encuentra las clases en local, clona este repositorio e instala `datasets` automáticamente. Basta con `Runtime → Run all`. Requiere que el repositorio sea público (el clon usa HTTPS).
5. **Ejecutar la aplicación de Streamlit (EA4)** desde la raíz del repositorio:
   ```bash
   streamlit run EA4/app/streamlit_app.py
   ```
   La app lee los datos de `EA4/resultados/` y muestra la necesidad, las cinco preguntas con sus métricas y gráficos interactivos, y las conclusiones. No requiere conexión a internet ni descargar el dataset.
6. **Regenerar las figuras y el documento completo (EA4)** (leen las cifras de `EA4/resultados/`):
   ```bash
   python EA4/src/generar_figuras.py     # PNG de los gráficos (para el PDF)
   python EA4/src/generar_documento.py   # informe final .docx
   ```
   El `.docx` usa campos TOC para los índices de contenido, tablas y figuras, y campos `SEQ`. Para resolverlos y exportar el PDF sin abrir Word, se usa LibreOffice en modo headless (UNO): refresca los campos, actualiza los índices y exporta `writer_pdf_Export`. El PDF entregado ya muestra los índices resueltos.

## Archivos generados

- **EA1/resultados/**
  - `iris_crudo.csv` — dataset Iris tal como se descargó.
  - `iris_transformado.csv` — tras las cinco transformaciones.
  - `resumen_dataset.txt` — resumen descriptivo del DataFrame transformado.
- **AE2/resultados/**
  - `churn_crudo.csv` — dataset Churn tal como se descargó.
  - `churn_transformado.csv` — tras las transformaciones de preparación.
  - `resumen_dataframe.txt` — resumen descriptivo del DataFrame transformado.
- **EA3/resultados/**
  - `churn_crudo.csv` — dataset Churn sin modificar (versión de referencia).
  - `churn_transformado.csv` — DataFrame transformado y enriquecido (26 columnas).
  - `tabla_p1_churn_por_contrato.csv` … `tabla_p5_churn_por_num_servicios.csv` — una tabla por pregunta de negocio.
  - `resumen_dataframe.txt` — resumen descriptivo del DataFrame final, generado automáticamente por código en cada ejecución.
- **EA3/documento/**
  - `Pedraza_Jeremy_AE3_Analitica_de_Datos.docx` / `.pdf` — documento de avance actualizado, con comparación antes/después (Tabla 2), diccionario, resultados por pregunta e índice de tablas.
- **EA4/resultados/**
  - `churn_transformado.csv`, `tabla_p1..p5 *.csv`, `resumen_dataframe.txt` — copias de los datos preparados y las tablas de resultados que consume la aplicación.
  - `figuras/fig_p1..p5.png` — gráficos de cada pregunta (Matplotlib), embebidos en el PDF e idénticos a los de la app.
- **EA4/documento/**
  - `Pedraza_Jeremy_AE4_Analitica_de_Datos.docx` / `.pdf` — **informe final completo** del proyecto (todas las etapas), con índices de contenido, tablas y figuras resueltos.
