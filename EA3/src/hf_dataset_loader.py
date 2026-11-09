"""Carga de datasets desde Hugging Face Hub.

Expone la clase :class:`HFDatasetLoader`, cuya única responsabilidad es
conectarse a Hugging Face y devolver un ``pandas.DataFrame``. No resume ni
transforma datos: ese trabajo corresponde a ``DatasetSummarizer`` y a los
notebooks.
"""

from __future__ import annotations

from datetime import datetime, timezone

import pandas as pd


class HFDatasetLoader:
    """Conecta con Hugging Face Hub y entrega un ``DataFrame`` de pandas.

    Parameters
    ----------
    dataset_id:
        Identificador del dataset en el Hub (por ejemplo ``"scikit-learn/iris"``).
    split:
        Partición a cargar. Por defecto ``"train"``.
    config:
        Nombre de la configuración/subconjunto, si el dataset tiene varias.
    revision:
        Revisión (rama, tag o commit) del dataset a fijar para reproducibilidad.
    """

    def __init__(
        self,
        dataset_id: str,
        split: str = "train",
        config: str | None = None,
        revision: str | None = None,
    ) -> None:
        self.dataset_id = dataset_id
        self.split = split
        self.config = config
        self.revision = revision

        # Estado interno poblado tras llamar a ``load()``.
        self._dataset = None
        self._dataframe: pd.DataFrame | None = None
        self._es_muestra: bool = False
        self._fecha_consulta: str | None = None

    def available_splits(self) -> list[str]:
        """Devuelve la lista de splits disponibles para el dataset.

        Returns
        -------
        list[str]
            Nombres de los splits publicados en el Hub.

        Raises
        ------
        RuntimeError
            Si falla la consulta de red al Hub.
        """
        try:
            from datasets import get_dataset_split_names

            return list(
                get_dataset_split_names(
                    self.dataset_id,
                    config_name=self.config,
                    revision=self.revision,
                )
            )
        except Exception as exc:  # pragma: no cover - depende de la red
            raise RuntimeError(
                f"No se pudieron consultar los splits de '{self.dataset_id}': {exc}"
            ) from exc

    def load(self):
        """Ejecuta ``datasets.load_dataset`` y guarda el resultado internamente.

        Valida que el split solicitado exista antes de cargar; si no existe,
        lanza ``ValueError`` con un mensaje claro.

        Returns
        -------
        datasets.Dataset
            El objeto ``Dataset`` cargado.

        Raises
        ------
        ValueError
            Si el split solicitado no está disponible.
        RuntimeError
            Si ocurre un error de red o de carga.
        """
        try:
            from datasets import load_dataset

            splits = self.available_splits()
            if self.split not in splits:
                raise ValueError(
                    f"El split '{self.split}' no existe en '{self.dataset_id}'. "
                    f"Splits disponibles: {splits}"
                )

            self._dataset = load_dataset(
                self.dataset_id,
                name=self.config,
                split=self.split,
                revision=self.revision,
            )
            self._fecha_consulta = datetime.now(timezone.utc).isoformat()
            return self._dataset
        except ValueError:
            raise
        except Exception as exc:  # pragma: no cover - depende de la red
            raise RuntimeError(
                f"Error al cargar '{self.dataset_id}' (split='{self.split}'): {exc}"
            ) from exc

    def to_dataframe(self) -> pd.DataFrame:
        """Devuelve el dataset como ``DataFrame``; carga si aún no se ha hecho.

        Returns
        -------
        pandas.DataFrame
            Contenido completo del split cargado.
        """
        if self._dataset is None:
            self.load()
        self._dataframe = self._dataset.to_pandas()
        self._es_muestra = False
        return self._dataframe

    def sample(self, n: int, random_state: int = 42) -> pd.DataFrame:
        """Devuelve una muestra reproducible de ``n`` filas.

        Parameters
        ----------
        n:
            Cantidad de filas a muestrear.
        random_state:
            Semilla para reproducibilidad.

        Returns
        -------
        pandas.DataFrame
            Muestra aleatoria del dataset.
        """
        df = self.to_dataframe()
        n = min(n, len(df))
        muestra = df.sample(n=n, random_state=random_state).reset_index(drop=True)
        self._es_muestra = True
        return muestra

    def get_metadata(self) -> dict:
        """Devuelve metadatos de la consulta realizada.

        Returns
        -------
        dict
            Diccionario con ``dataset_id``, ``split``, ``config``, ``revision``,
            ``fecha_consulta`` (ISO), ``n_filas``, ``n_columnas`` y ``es_muestra``.
        """
        df = self._dataframe if self._dataframe is not None else self.to_dataframe()
        return {
            "dataset_id": self.dataset_id,
            "split": self.split,
            "config": self.config,
            "revision": self.revision,
            "fecha_consulta": self._fecha_consulta,
            "n_filas": int(df.shape[0]),
            "n_columnas": int(df.shape[1]),
            "es_muestra": self._es_muestra,
        }
