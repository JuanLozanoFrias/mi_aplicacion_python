README — Starter “Activos + Inventario en JSON” para Calvo
========================================================

Qué incluye:
- company_data/ : paquete JSON de ejemplo con manifest + hashes
- datahub.py    : loader/validador (hashes) + getters para Calvo
- siesa_uno_client.py : cliente ODBC (diagnóstico + consultas)
- export_siesa_to_package.py : exportador Siesa -> snapshots + manifest

Uso rápido:
1) Copia la carpeta company_data/ a tu proyecto (por ejemplo: data/company_data/)
2) En Calvo:
   from datahub import CompanyDataHub
   hub = CompanyDataHub("data/company_data")
   hub.load(verify_hashes=True)
   inv = hub.get_inventory("Weston")
   assets = hub.get_assets()

3) Para exportar desde Siesa (en el PC con acceso):
   python export_siesa_to_package.py

Notas:
- Para Calvo vía ODBC: instala pyodbc y el ODBC Driver 17/18 para SQL Server.
