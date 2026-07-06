"""
locales.py
----------
Definición única de los locales de Europa. La importan report.py, generar_acum2025.py
y fetch_shares.py, así no repetimos la lista en cada archivo.

Grupos: Propias vs Franquicias (mismo criterio que USA).
Locales sin histórico 2025 (Madrid, Alicante) abrieron en 2026: no se comparan.
NO hace falta listarlos acá como caso especial: el motor decide "sin comparación"
cuando el acumulado 2025 de ese local da cero (regla robusta que también cubre a
Granada en los meses en que todavía no operaba).
"""

PROPIAS = ["Roma", "Barcelona 1", "Barcelona 2", "Madrid", "Málaga 1"]
FRANQUICIAS = ["Málaga 3", "Granada", "Alicante", "Valencia"]

# Orden de aparición en el reporte (propias primero, luego franquicias)
BRANCH_ORDER = PROPIAS + FRANQUICIAS
