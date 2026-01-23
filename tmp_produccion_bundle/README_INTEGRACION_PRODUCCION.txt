INTEGRACION PRODUCCION (CALVOBOT)

1) Copia estas carpetas/archivos al root del proyecto destino:
   - gui/pages/produccion_page.py
   - logic/produccion/*.py
   - data/produccion/* (incluye subcarpeta MATERIALES)
   - (opcional) data/siesa/sql/ordenes_produccion.sql

2) Integra el menu lateral en gui/main_window.py:
   - Agrega el import:
       from gui.pages.produccion_page import ProduccionPage
   - Agrega el item de menu donde corresponda:
       add_nav_item("PRODUCCION", ProduccionPage())
   - No elimines paginas existentes. Solo suma el item.

3) Dependencias:
   - pandas
   - openpyxl
   - PySide6 (ya existente)
   - PySide6-QtCharts (opcional; si no esta instalado, la pagina funciona sin graficas)

4) Datos:
   - Los Excel se buscan en data/produccion con nombres que incluyan:
       * ORDENES_PRODUCCION
       * CONTROL PERSONAL
   - Los listados de materiales se leen desde data/produccion/MATERIALES

5) Prueba rapida:
   - Ejecuta CalvoBot
   - Abre PRODUCCION en el sidebar
   - Verifica que cargue OPs y personal

Notas:
- Este bundle no modifica otras areas; solo agrega la pagina y el item de menu.
- Si ya existe un modulo de Produccion, revisa colisiones de nombres.
