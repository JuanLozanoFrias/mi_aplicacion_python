# logic/data_loader.py
from pathlib import Path
import pandas as pd

def load_equipo_options(path: str | None = None) -> list[str]:
    """
    Lee la columna A de la hoja “MueblesFríos” dentro de data/basedatos.xlsx
    (o la ruta suministrada) y devuelve una lista de cadenas únicas.
    """
    xls = Path(path) if path else Path(__file__).resolve().parents[1] / "data" / "basedatos.xlsx"
    df = pd.read_excel(xls, sheet_name="MUEBLESFRIOS", usecols="A", header=None)
    valores = (df.iloc[:, 0]
                 .dropna()
                 .astype(str)
                 .str.strip()
                 .unique()
                 .tolist())
    return valores
