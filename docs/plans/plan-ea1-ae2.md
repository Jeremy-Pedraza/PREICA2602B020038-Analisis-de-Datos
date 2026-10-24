# Plan EA1 + AE2

Plan único y actualizable de las entregas. Estados: ✅ hecho · 🔄 en curso · ⬜ pendiente.

## Fase 1 — Clases de Python ✅
- `HFDatasetLoader` (carga desde Hugging Face → DataFrame).
- `DatasetSummarizer` (resumen de DataFrame → `.txt`).
- Copias idénticas en `EA1/src/` y `AE2/src/`.

## Fase 2 — Notebook EA1 ✅
- `ea1_carga_y_transformaciones.ipynb`: carga iris, crudo, T1–T5, transformado, resumen, verificación.

## Fase 3 — Notebook AE2 ✅
- `ae2_exploracion_descriptiva.ipynb`: carga churn, crudo, exploración, necesidad, T1–T5, transformado, resumen, verificación.

## Fase 4 — Archivos de soporte ✅
- `requirements.txt`, `.gitignore`, `README.md`.
- `docs/master.md`, este plan, checkpoint de sesión.

## Fase 5 — Ejecución y verificación ⬜
- Instalar dependencias.
- Ejecutar ambos notebooks con `nbconvert --execute --inplace`.
- Verificar los 2 `.txt` y los 4 `.csv`.
- Imprimir ambos resúmenes.

## Fase 6 — Git ⬜
- Commits descriptivos por bloque.
- Push a `main`.
