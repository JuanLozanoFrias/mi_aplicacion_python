# Calvo - Guia de desarrollo (Windows 11)

## Requisitos
- Python 3.13 (recomendado)
- Git (si vas a clonar/actualizar)

## Instalacion rapida
1) Abre PowerShell en la carpeta del proyecto.
2) Ejecuta:
   - `start_dev.bat`

Eso crea el entorno virtual, instala dependencias y abre la app.

## Instalacion manual (si no usas el .bat)
```powershell
cd C:\ruta\al\proyecto
py -3.13 -m venv .venv
.\.venv\Scripts\activate
python -m pip install --upgrade pip
pip install -r requirements.txt
python main.py
```

## Notas importantes
- Ejecuta siempre desde la carpeta raiz (la que tiene `main.py` y `data/`).
- Si actualizas el codigo, vuelve a ejecutar `pip install -r requirements.txt`.

## SIESA (opcional)
Si vas a usar el modulo SIESA:
1) Instala ODBC Driver 18 for SQL Server.
2) Crea credencial en Windows (Administrador de credenciales):
   - Target: CalvoSiesaUNOEE
   - Usuario: sa
   - Clave: (tu clave)

Comando alterno:
```
cmdkey /generic:CalvoSiesaUNOEE /user:sa /pass:TU_CLAVE
```

## Soporte rapido
Si la app no abre:
1) Verifica que `python --version` sea 3.13 (o 3.12).
2) Activa el entorno y reinstala dependencias.
3) Ejecuta `python main.py` desde la raiz del repo.
