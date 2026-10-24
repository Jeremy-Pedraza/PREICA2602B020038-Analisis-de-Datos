# Master - Fuente de verdad

Documento maestro de las entregas EA1 y AE2 de Analítica de Datos (PREICA2602B020038).

## Objetivo de cada entrega

- **EA1 — Carga y creación de dataset:** demostrar la carga de un dataset desde Hugging Face, su exportación en crudo, la aplicación de al menos cuatro transformaciones distintas (se implementan cinco) y la generación de un resumen descriptivo en `.txt`.
- **AE2 — Primera etapa del proyecto integrador:** cargar el dataset de churn, realizar la exploración descriptiva, justificar la necesidad analítica, aplicar transformaciones de preparación y generar el resumen del DataFrame.

## Decisiones tomadas

- **Datasets (no se cambian):**
  - EA1 → `scikit-learn/iris`, split `train`. Dataset clásico, limpio, ideal para ilustrar transformaciones.
  - AE2 → `scikit-learn/churn-prediction`, split `train` (Telco Customer Churn, ~7.04k filas). Caso de negocio real de abandono de clientes.
- **Arquitectura de clases:** responsabilidad única.
  - `HFDatasetLoader`: solo carga desde el Hub y entrega un DataFrame.
  - `DatasetSummarizer`: solo resume un DataFrame recibido.
- **Autocontención:** `AE2/src/` son copias idénticas de `EA1/src/` para que cada entrega sea independiente.
- **Versionado de resultados:** los `.csv` y `.txt` de `resultados/` SÍ se versionan (son evidencia para el docente).

## Índice de docs/

- `docs/master.md` — este documento (fuente de verdad).
- `docs/plans/plan-ea1-ae2.md` — plan único con las fases y su estado.
- `docs/checkpoints/2026-10-24-sesion-01.md` — progreso de la sesión.

## Mapeo rúbrica → evidencia

| Criterio | Evidencia |
|----------|-----------|
| Carga de dataset desde Hugging Face | `EA1/src/hf_dataset_loader.py`, celda 2 de cada notebook |
| Exportación del dataset crudo | `EA1/resultados/iris_crudo.csv`, `AE2/resultados/churn_crudo.csv` |
| Mínimo 4 transformaciones distintas | T1–T5 en ambos notebooks |
| Dataset transformado exportado | `*_transformado.csv` en cada `resultados/` |
| Resumen descriptivo del DataFrame | `DatasetSummarizer`, `resumen_dataset.txt` y `resumen_dataframe.txt` |
| Exploración descriptiva (AE2) | Celdas 4 del notebook AE2 (`head`, `describe`, nulos, duplicados, `value_counts`) |
| Justificación de la necesidad analítica | Celda markdown 5 del notebook AE2 y README |
| Documentación del repositorio | `README.md`, `docs/` |
| Sin credenciales expuestas | `.gitignore`, sección Seguridad del README |
