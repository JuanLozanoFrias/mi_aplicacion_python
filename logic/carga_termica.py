# logic/carga_termica.py
import pandas as pd

class CargaTermicaCalculator:
    def from_excel(self, path: str) -> pd.DataFrame:
        return pd.read_excel(path, engine="openpyxl")

    def calcular(self, df: pd.DataFrame) -> dict:
        registros = len(df)
        total_kw = df["kW"].sum()
        promedio = df["kW"].mean()
        return {
            "Registros": registros,
            "Total kW": round(total_kw,2),
            "Promedio": round(promedio,2)
        }
