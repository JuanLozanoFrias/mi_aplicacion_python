# logic/calculo_cargas.py

def calcular_tabla_cargas(n_baja: int, n_media: int) -> list[tuple[int,str]]:
    """
    Devuelve una lista de (ramal_id, grupo) para poblar la tabla.
    """
    total = n_baja + n_media
    filas = []
    for i in range(total):
        grupo = "B" if i < n_baja else "M"
        filas.append((i+1, grupo))
    return filas
