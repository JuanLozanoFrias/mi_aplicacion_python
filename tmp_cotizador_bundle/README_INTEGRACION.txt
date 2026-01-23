INTEGRACION COTIZADOR (CALVOBOT)

1) Copia estas carpetas al root del proyecto destino:
   - gui/pages/cotizador_page.py
   - logic/cotizador/__init__.py
   - logic/cotizador/cotizador_engine.py
   - data/cotizador/*
   - data/CUADRO DE CARGAS -INDIVIDUAL JUAN LOZANO.xlsx
   - (opcional) README_COTIZADOR.md

2) Integra el menu lateral en gui/main_window.py:
   - Agrega el import:
       from gui.pages.cotizador_page import CotizadorPage
   - Agrega el item de menu donde corresponda:
       add_nav_item("COTIZADOR", CotizadorPage())
   - Si ya tienes otras paginas, no elimines nada. Solo suma el item.

3) Dependencias:
   - pandas
   - openpyxl

4) Rutas:
   - El modulo usa rutas relativas al root del proyecto.
   - No uses rutas absolutas.

5) Datos demo:
   - JSONs en data/cotizador
   - El catalogo de equipos se lee del Excel (data/CUADRO DE CARGAS -INDIVIDUAL JUAN LOZANO.xlsx)

6) Prueba rapida:
   - Ejecuta CalvoBot
   - Abre COTIZADOR en el sidebar
   - Agrega un equipo y exporta PDF

Notas:
- Si el proyecto destino ya tiene una version de cotizador, revisa colisiones de nombres.
- Este bundle no modifica otras areas; solo agrega el modulo y el item de menu.
