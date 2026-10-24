# PREICA2602B020038 - Analítica de Datos

Repositorio de las entregas **EA1** (carga y creación de dataset) y **AE2** (primera etapa del proyecto integrador) de la asignatura Proyecto Integrado III – Analítica de Datos.

## Propósito

Cargar datasets públicos desde Hugging Face Hub, aplicarles transformaciones de preparación y generar un resumen descriptivo automático en texto, de forma reproducible y autocontenida por entrega.

## Asignatura, docente e integrante

- **Asignatura:** Proyecto Integrado III – Analítica de Datos
- **Código:** PREICA2602B020038
- **Docente:** Jeremy Pedraza
- **Integrante:** Jeremy Pedraza

## Necesidad analítica (AE2)

Una compañía de telecomunicaciones necesita anticipar el **abandono de clientes** (*churn*). Cada registro del dataset es un cliente; a partir de su antigüedad, tipo de contrato, servicios contratados y facturación se busca identificar los factores asociados al abandono para diseñar acciones de retención. Esta primera etapa cubre la carga, exploración descriptiva y preparación del dataset.

## Datasets usados

| Entrega | Dataset | Split | Enlace |
|---------|---------|-------|--------|
| EA1 | `scikit-learn/iris` | `train` | https://huggingface.co/datasets/scikit-learn/iris |
| AE2 | `scikit-learn/churn-prediction` | `train` | https://huggingface.co/datasets/scikit-learn/churn-prediction |

## Estructura de carpetas

```
.
├── README.md
├── requirements.txt
├── .gitignore
├── docs/
│   ├── master.md
│   ├── plans/plan-ea1-ae2.md
│   └── checkpoints/2026-10-24-sesion-01.md
├── EA1/
│   ├── src/                 # HFDatasetLoader + DatasetSummarizer
│   ├── notebooks/           # ea1_carga_y_transformaciones.ipynb
│   └── resultados/          # iris_crudo.csv, iris_transformado.csv, resumen_dataset.txt
└── AE2/
    ├── documento/           # plantilla, generar vía script → .docx + .pdf
    ├── src/                 # copias idénticas de EA1/src + generar_documento.py
    ├── notebooks/           # ae2_exploracion_descriptiva.ipynb
    └── resultados/          # churn_crudo.csv, churn_transformado.csv, resumen_dataframe.txt
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
   ```
   O abrirlos con `jupyter notebook` y ejecutar todas las celdas.

## Archivos generados

- **EA1/resultados/**
  - `iris_crudo.csv` — dataset Iris tal como se descargó.
  - `iris_transformado.csv` — tras las cinco transformaciones.
  - `resumen_dataset.txt` — resumen descriptivo del DataFrame transformado.
- **AE2/resultados/**
  - `churn_crudo.csv` — dataset Churn tal como se descargó.
  - `churn_transformado.csv` — tras las transformaciones de preparación.
  - `resumen_dataframe.txt` — resumen descriptivo del DataFrame transformado.

## Documento de la AE2

El documento Word se genera de forma reproducible:

```bash
python AE2/src/generar_documento.py
```

Lee la plantilla `AE2/documento/plantilla_iudigital.docx` y las salidas en `AE2/resultados/` (todas las cifras provienen de ahí) y produce `AE2/documento/Pedraza_Jeremy_AE2_Analitica_de_Datos.docx`.

> **Nota:** al abrir el `.docx` en Word, seleccionar la tabla de contenido ("Contenido" y "Lista de tablas") y presionar **F9** para numerar las páginas e índices.

## Seguridad

No se publican tokens, credenciales ni archivos `.env`. El acceso a Hugging Face Hub se hace de forma anónima (datasets públicos).
