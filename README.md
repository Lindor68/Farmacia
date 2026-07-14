# Dashboard Farmacia — instalación y ejecución

## 1. Instalar dependencias

Requiere Python 3.10+ ya instalado. Desde la carpeta del proyecto:

```
pip install -r requirements.txt
```

Instala: `pandas`, `numpy`, `plotly`, `openpyxl`, `statsmodels`, `scipy` (estas dos
últimas son nuevas: las usa el módulo de pronóstico de demanda para
Holt-Winters y SARIMA).

## 2. Archivos de entrada (reemplazar cada mes)

- `Consumos_Historicos.xlsx` — histórico de consumo mensual por establecimiento.
- `Sal_Art_*.xlsx` — un archivo por bodega, export de stock/saldos.

## 3. Ejecutar

```
python generar_dashboard_completo.py
```

Genera:
- `consolidado.xlsx` — stock crudo consolidado (para auditar).
- `resumen_stock_BFC.xlsx` — saldos, valorización, proyecciones y alertas.
- `dashboard_farmacia.html` — dashboard interactivo, 100% autocontenido (se
  abre con doble clic, no requiere internet ni servidor).

La primera vez que corre (o cada vez que cambian los archivos de entrada),
el módulo de pronóstico de demanda evalúa miles de series código+
establecimiento con validación temporal y puede tardar entre 10 y 30
minutos. Si los archivos de entrada no cambiaron desde la última corrida, se
reusa un caché (`.pronostico_cache.pkl`) y esa etapa se salta.

## 4. Otros scripts del proyecto

`generar_pedido.py`, `generar_resumen_stock.py`, `generar_resumen_stock_compra.py`
y `generar_dashboard.py` son pasos individuales del mismo flujo (previos a
que existiera `generar_dashboard_completo.py`, que los reemplaza corriendo
todo en una sola pasada). Se mantienen sin cambios; no son necesarios para
generar el dashboard si usas `generar_dashboard_completo.py`.
