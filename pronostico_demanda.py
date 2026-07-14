"""
pronostico_demanda.py

Motor de pronóstico de demanda y abastecimiento farmacéutico. Módulo puro:
no lee ni escribe archivos — recibe estructuras ya cargadas por
generar_dashboard_completo.py y devuelve resultados listos para embeber en
el dashboard. Así el pipeline existente no se toca, solo se le agrega un
paso nuevo que llama a este módulo.

Resumen del enfoque (ver también la sección "Metodología" del dashboard):
  1. Se reconstruye cada serie código+establecimiento (y su agregado en toda
     la red, "RED") en una grilla mensual continua, distinguiendo mes sin
     registro (NaN en la grilla original) de consumo cero explícito.
  2. Se evalúan varios métodos de pronóstico por serie mediante validación
     temporal walk-forward (nunca se usa dato futuro para entrenar).
  3. Se selecciona el modelo con menor WAPE (desempate por MASE) respetando
     las reglas de historia mínima pedidas por el negocio.
  4. Se generan pronósticos a 12 meses (el dashboard recorta a 3/6/12) con
     intervalos de predicción 80% y 95%.

Todas las cifras resultantes son ESTIMACIONES ESTADÍSTICAS sujetas a
incertidumbre, no valores exactos.
"""
import warnings
import numpy as np
import pandas as pd

warnings.filterwarnings("ignore")

try:
    from statsmodels.tsa.holtwinters import SimpleExpSmoothing, Holt, ExponentialSmoothing
    from statsmodels.tsa.statespace.sarimax import SARIMAX
    STATSMODELS_DISPONIBLE = True
except ImportError:
    STATSMODELS_DISPONIBLE = False


# ── Parámetros del motor ──────────────────────────────────────────────────────
Z80 = 1.2815515655446004   # z de la normal estándar para 80% de cobertura
Z95 = 1.9599639845400545   # z de la normal estándar para 95% de cobertura

MIN_MESES_NO_ESTACIONAL = 12   # bajo esto: "historia insuficiente"
MIN_MESES_ESTACIONAL    = 24   # bajo esto: sin modelos estacionales
UMBRAL_INTERMITENTE     = 0.30 # % de meses en cero para tratar como demanda intermitente
FOLDS_CV                = 2    # folds walk-forward para modelos "livianos"
FOLDS_CV_COSTOSO        = 1    # folds walk-forward para SARIMA/Holt-Winters (más caro de ajustar)
HORIZONTE_MAX           = 12
PERIODO_ESTACIONAL      = 12

# SARIMA solo se evalúa a nivel agregado de red (no por establecimiento):
# es el modelo más costoso de ajustar y el que menos aporta sobre series
# delgadas/ruidosas de un solo centro; se documenta como limitación conocida.
SARIMA_SOLO_AGREGADO_RED = True
SARIMA_ORDEN = (1, 1, 1)
SARIMA_ORDEN_ESTACIONAL = (0, 1, 1, PERIODO_ESTACIONAL)

DESCRIPCION_MODELOS = {
    "media_simple":     "Promedio simple (historia insuficiente)",
    "naive_simple":     "Naive (último valor)",
    "promedio_movil":   "Promedio móvil (3 meses)",
    "ses":              "Suavizamiento exponencial simple",
    "holt":             "Holt (tendencia amortiguada)",
    "holt_winters":     "Holt-Winters (estacional aditivo)",
    "sarima":           "SARIMA",
    "naive_estacional": "Naive estacional (referencia)",
    "croston":          "Croston (demanda intermitente)",
    "tsb":              "TSB (demanda intermitente)",
}


# ════════════════════════════════════════════════════════════════════════════
# 1. CALIDAD DE DATOS Y CONSTRUCCIÓN DE LA GRILLA MENSUAL
# ════════════════════════════════════════════════════════════════════════════
def construir_grillas(ch, cols_consumo):
    """A partir del histórico ancho (una fila por código+fecha, una columna
    por establecimiento), arma por cada código una grilla mensual continua
    (desde su primer hasta su último mes con registro) con una columna extra
    "RED" = suma de todos los establecimientos de consumo real.

    Un mes sin fila original queda NaN ("dato ausente"); un mes con fila pero
    valor 0 queda en 0 ("consumo cero explícito"). Se detectan y corrigen
    duplicados (se suman) y negativos (se llevan a 0) antes de construir la
    grilla, y se cuentan para el reporte de calidad.
    """
    ch = ch.copy()
    ch["Fecha"] = pd.to_datetime(ch["Fecha"])
    n_duplicados = int(ch.duplicated(subset=["Codigo", "Fecha"]).sum())

    agg = ch.groupby(["Codigo", "Fecha"], as_index=False)[cols_consumo].sum()
    n_negativos = int((agg[cols_consumo].to_numpy() < 0).sum())
    agg[cols_consumo] = agg[cols_consumo].clip(lower=0)
    agg["RED"] = agg[cols_consumo].sum(axis=1)

    columnas_valor = cols_consumo + ["RED"]
    grillas = {}
    calidad_codigo = {}

    for codigo, grupo in agg.groupby("Codigo"):
        grupo = grupo.sort_values("Fecha")
        periodos = grupo["Fecha"].dt.to_period("M")
        idx_completo = pd.period_range(periodos.min(), periodos.max(), freq="M")

        serie_df = grupo.set_index(periodos)[columnas_valor]
        serie_df = serie_df[~serie_df.index.duplicated(keep="first")]
        serie_df = serie_df.reindex(idx_completo)  # NaN = mes ausente

        meses_esperados = len(idx_completo)
        meses_ausentes = int(serie_df["RED"].isna().sum())

        grillas[codigo] = serie_df
        calidad_codigo[codigo] = {
            "meses_esperados": meses_esperados,
            "meses_con_registro": meses_esperados - meses_ausentes,
            "meses_ausentes": meses_ausentes,
            "primer_mes": str(idx_completo[0]),
            "ultimo_mes": str(idx_completo[-1]),
            "historia_insuficiente": meses_esperados < MIN_MESES_NO_ESTACIONAL,
        }

    return grillas, calidad_codigo, n_duplicados, n_negativos


def detectar_outliers(valores):
    """Outliers vía z-score robusto (MAD). Solo informativo: no se eliminan
    de la serie usada para entrenar los modelos (una alza real de consumo,
    p.ej. una campaña, no debe borrarse silenciosamente)."""
    v = valores[~np.isnan(valores)]
    if len(v) < 4:
        return np.zeros_like(valores, dtype=bool)
    mediana = np.median(v)
    mad = np.median(np.abs(v - mediana))
    if mad == 0:
        return np.zeros_like(valores, dtype=bool)
    z = np.abs(valores - mediana) / (1.4826 * mad)
    return np.nan_to_num(z, nan=0.0) > 3.5


# ════════════════════════════════════════════════════════════════════════════
# 2. MODELOS DE PRONÓSTICO
# ════════════════════════════════════════════════════════════════════════════
def _resultado(forecast, sigma, ci80=None, ci95=None):
    forecast = np.clip(np.asarray(forecast, dtype=float), 0, None)
    return {"forecast": forecast, "sigma": float(max(sigma, 0.0)), "ci80": ci80, "ci95": ci95}


def _media_simple(y, h):
    nivel = float(np.mean(y)) if len(y) else 0.0
    resid = y - nivel if len(y) else np.array([0.0])
    return _resultado(np.full(h, nivel), np.std(resid))


def _naive_simple(y, h):
    nivel = float(y[-1]) if len(y) else 0.0
    resid = np.diff(y) if len(y) > 1 else np.array([0.0])
    return _resultado(np.full(h, nivel), np.std(resid))


def _promedio_movil(y, h, ventana=3):
    ventana = min(ventana, len(y)) or 1
    s = pd.Series(y)
    pred_in_sample = s.rolling(ventana, min_periods=1).mean().shift(1)
    resid = (s - pred_in_sample).dropna().to_numpy()
    nivel = float(s.tail(ventana).mean())
    sigma = np.std(resid) if len(resid) else np.std(y)
    return _resultado(np.full(h, nivel), sigma)


def _naive_estacional(y, h, periodo=PERIODO_ESTACIONAL):
    if len(y) < periodo:
        return None
    ultimo_ciclo = y[-periodo:]
    reps = int(np.ceil(h / periodo))
    forecast = np.tile(ultimo_ciclo, reps)[:h]
    resid = y[periodo:] - y[:-periodo]
    sigma = np.std(resid) if len(resid) else np.std(y)
    return _resultado(forecast, sigma)


def _ses(y, h):
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        modelo = SimpleExpSmoothing(y, initialization_method="estimated").fit(optimized=True)
    return _resultado(modelo.forecast(h), np.std(modelo.resid))


def _holt(y, h):
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        modelo = Holt(y, damped_trend=True, initialization_method="estimated").fit(optimized=True)
    return _resultado(modelo.forecast(h), np.std(modelo.resid))


def _holt_winters(y, h, periodo=PERIODO_ESTACIONAL):
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        modelo = ExponentialSmoothing(
            y, trend="add", damped_trend=True, seasonal="add",
            seasonal_periods=periodo, initialization_method="estimated",
        ).fit(optimized=True)
    return _resultado(modelo.forecast(h), np.std(modelo.resid))


def _sarima(y, h):
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        modelo = SARIMAX(
            y, order=SARIMA_ORDEN, seasonal_order=SARIMA_ORDEN_ESTACIONAL,
            enforce_stationarity=False, enforce_invertibility=False,
        ).fit(disp=False, maxiter=60)
    pred = modelo.get_forecast(h)
    forecast = np.clip(np.asarray(pred.predicted_mean), 0, None)
    ci80 = np.asarray(pred.conf_int(alpha=0.20))
    ci95 = np.asarray(pred.conf_int(alpha=0.05))
    ci80 = (np.clip(ci80[:, 0], 0, None), ci80[:, 1])
    ci95 = (np.clip(ci95[:, 0], 0, None), ci95[:, 1])
    sigma = np.std(modelo.resid)
    return _resultado(forecast, sigma, ci80, ci95)


def _croston(y, h, alpha=0.1):
    if not np.any(y > 0):
        return _resultado(np.zeros(h), 0.0)
    z = p = None
    q = 1
    demandas = []
    for val in y:
        if val > 0:
            if z is None:
                z, p = float(val), float(q)
            else:
                z = alpha * val + (1 - alpha) * z
                p = alpha * q + (1 - alpha) * p
            demandas.append(val)
            q = 1
        else:
            q += 1
    if not p:
        return _resultado(np.zeros(h), 0.0)
    nivel = z / p
    sigma = (np.std(demandas) / p) if len(demandas) > 1 else nivel * 0.5
    return _resultado(np.full(h, nivel), sigma)


def _tsb(y, h, alpha=0.1, beta=0.1):
    if len(y) == 0:
        return _resultado(np.zeros(h), 0.0)
    ocurre = (y > 0).astype(float)
    prob = ocurre[0]
    nivel_demanda = float(y[0]) if y[0] > 0 else (float(np.mean(y[y > 0])) if np.any(y > 0) else 0.0)
    niveles = []
    for t in range(1, len(y)):
        prob = alpha * ocurre[t] + (1 - alpha) * prob
        if y[t] > 0:
            nivel_demanda = beta * y[t] + (1 - beta) * nivel_demanda
        niveles.append(prob * nivel_demanda)
    nivel_final = prob * nivel_demanda
    resid = y[1:] - np.array(niveles) if niveles else np.array([0.0])
    return _resultado(np.full(h, nivel_final), np.std(resid))


def ajustar_modelo(nombre, y, h):
    """Ajusta el modelo `nombre` sobre la serie `y` y pronostica `h` meses.
    Nunca lanza: cualquier falla de convergencia se traduce en None y el
    llamador sigue con el siguiente candidato (una serie problemática no
    puede tumbar la corrida completa)."""
    try:
        if nombre == "media_simple":
            return _media_simple(y, h)
        if nombre == "naive_simple":
            return _naive_simple(y, h)
        if nombre == "promedio_movil":
            return _promedio_movil(y, h)
        if nombre == "naive_estacional":
            return _naive_estacional(y, h)
        if nombre == "croston":
            return _croston(y, h)
        if nombre == "tsb":
            return _tsb(y, h)
        if not STATSMODELS_DISPONIBLE:
            return None
        if nombre == "ses":
            return _ses(y, h)
        if nombre == "holt":
            return _holt(y, h)
        if nombre == "holt_winters":
            return _holt_winters(y, h)
        if nombre == "sarima":
            return _sarima(y, h)
    except Exception:
        return None
    return None


# ════════════════════════════════════════════════════════════════════════════
# 3. VALIDACIÓN TEMPORAL (WALK-FORWARD) Y MÉTRICAS
# ════════════════════════════════════════════════════════════════════════════
def _walk_forward(nombre, y, folds, min_train):
    n = len(y)
    folds_reales = min(folds, max(0, n - min_train))
    if folds_reales <= 0:
        return None
    reales, predichos = [], []
    for k in range(folds_reales, 0, -1):
        corte = n - k
        if corte < min_train:
            continue
        resultado = ajustar_modelo(nombre, y[:corte], 1)
        if resultado is None:
            continue
        predichos.append(resultado["forecast"][0])
        reales.append(y[corte])
    if not predichos:
        return None
    return np.array(reales), np.array(predichos)


def calcular_metricas(reales, predichos, y_hist):
    """MAE, RMSE, WAPE, MASE y sesgo. Nunca se calcula MAPE (inestable con
    consumos en cero, frecuentes en este rubro)."""
    errores = predichos - reales
    mae = float(np.mean(np.abs(errores)))
    rmse = float(np.sqrt(np.mean(errores ** 2)))
    suma_real = float(np.sum(np.abs(reales)))
    wape = float(np.sum(np.abs(errores)) / suma_real) if suma_real > 0 else None

    if len(y_hist) > PERIODO_ESTACIONAL:
        diffs = np.abs(y_hist[PERIODO_ESTACIONAL:] - y_hist[:-PERIODO_ESTACIONAL])
    else:
        diffs = np.abs(np.diff(y_hist))
    denom = float(np.mean(diffs)) if len(diffs) and np.mean(diffs) > 0 else None
    mase = float(mae / denom) if denom else None

    return {"mae": mae, "rmse": rmse, "wape": wape, "mase": mase, "sesgo": float(np.mean(errores))}


def _clasificar_confianza(n, metricas):
    if n < MIN_MESES_NO_ESTACIONAL:
        return "Historia insuficiente"
    if metricas is None or metricas.get("wape") is None:
        return "Baja"
    wape = metricas["wape"]
    if n >= MIN_MESES_ESTACIONAL and wape <= 0.20:
        return "Alta"
    if wape <= 0.35:
        return "Media"
    return "Baja"


def _seleccionar_modelo(y, candidatos):
    resultados = {}
    for nombre in candidatos:
        folds = FOLDS_CV_COSTOSO if nombre in ("sarima", "holt_winters") else FOLDS_CV
        ev = _walk_forward(nombre, y, folds=folds, min_train=max(6, len(y) - FOLDS_CV - 1))
        if ev is None:
            continue
        reales, predichos = ev
        met = calcular_metricas(reales, predichos, y)
        if met["wape"] is None:
            continue
        resultados[nombre] = met
    if not resultados:
        return None, {}
    mejor = min(
        resultados.items(),
        key=lambda kv: (round(kv[1]["wape"], 4), kv[1]["mase"] if kv[1]["mase"] is not None else 1e9),
    )
    return mejor[0], resultados


def _bandas(resultado, h):
    forecast, sigma = resultado["forecast"], resultado["sigma"]
    if resultado.get("ci80") is not None and resultado.get("ci95") is not None:
        li80, ls80 = resultado["ci80"]
        li95, ls95 = resultado["ci95"]
    else:
        pasos = np.sqrt(np.arange(1, h + 1))
        li80, ls80 = forecast - Z80 * sigma * pasos, forecast + Z80 * sigma * pasos
        li95, ls95 = forecast - Z95 * sigma * pasos, forecast + Z95 * sigma * pasos
    return (np.clip(li80, 0, None), np.clip(ls80, 0, None),
            np.clip(li95, 0, None), np.clip(ls95, 0, None))


def analizar_serie(valores, permitir_sarima, h_max=HORIZONTE_MAX):
    """Evalúa candidatos elegibles según la historia disponible, selecciona
    el mejor por validación temporal y arma el pronóstico final + bandas.
    `valores` ya debe venir con los meses ausentes rellenados en 0 (se maneja
    fuera de esta función para poder contar la métrica de calidad aparte)."""
    n = len(valores)
    ratio_cero = float(np.mean(valores == 0)) if n else 1.0
    intermitente = ratio_cero > UMBRAL_INTERMITENTE

    if n < MIN_MESES_NO_ESTACIONAL:
        resultado = ajustar_modelo("media_simple", valores, h_max)
        modelo, metricas = "media_simple", None
    else:
        candidatos = ["naive_simple", "promedio_movil", "ses", "holt"]
        if intermitente:
            candidatos += ["croston", "tsb"]
        if n >= MIN_MESES_ESTACIONAL:
            candidatos += ["holt_winters", "naive_estacional"]
            if permitir_sarima and not intermitente:
                candidatos.append("sarima")

        modelo, evaluados = _seleccionar_modelo(valores, candidatos)
        if modelo is None:
            modelo, metricas = "naive_simple", None
            resultado = ajustar_modelo(modelo, valores, h_max)
        else:
            metricas = evaluados.get(modelo)
            resultado = ajustar_modelo(modelo, valores, h_max)
            if resultado is None:
                modelo, metricas = "naive_simple", None
                resultado = ajustar_modelo(modelo, valores, h_max)

    if resultado is None:  # último resguardo: nunca dejar una serie sin pronóstico
        modelo, metricas = "naive_simple", None
        resultado = ajustar_modelo(modelo, valores, h_max)

    li80, ls80, li95, ls95 = _bandas(resultado, h_max)
    return {
        "modelo": modelo,
        "modelo_desc": DESCRIPCION_MODELOS.get(modelo, modelo),
        "metricas": metricas,
        "nivel_confianza": _clasificar_confianza(n, metricas),
        "intermitente": intermitente,
        "sigma": resultado["sigma"],
        "forecast": resultado["forecast"],
        "li80": li80, "ls80": ls80, "li95": li95, "ls95": ls95,
    }


# ════════════════════════════════════════════════════════════════════════════
# 4. CLASIFICACIÓN ABC / XYZ
# ════════════════════════════════════════════════════════════════════════════
def clasificar_abc_xyz(valor_anual, coef_variacion):
    """valor_anual: dict código -> valor de consumo anual (unidades*precio).
    coef_variacion: dict código -> CV de la demanda mensual, o None si no hay
    historia suficiente para calcularlo. Devuelve dict código -> (abc, xyz)."""
    codigos = list(valor_anual.keys())
    if not codigos:
        return {}
    valores = np.array([valor_anual[c] for c in codigos], dtype=float)
    orden = np.argsort(-valores)
    total = valores.sum()
    abc = {}
    acumulado = 0.0
    for idx in orden:
        c = codigos[idx]
        acumulado += valores[idx]
        pct = (acumulado / total) if total > 0 else 1.0
        abc[c] = "A" if pct <= 0.80 else ("B" if pct <= 0.95 else "C")

    resultado = {}
    for c in codigos:
        cv = coef_variacion.get(c)
        if cv is None:
            xyz = "N/D"
        elif cv < 0.5:
            xyz = "X"
        elif cv <= 1.0:
            xyz = "Y"
        else:
            xyz = "Z"
        resultado[c] = {"abc": abc.get(c, "C"), "xyz": xyz}
    return resultado


# ════════════════════════════════════════════════════════════════════════════
# 5. ORQUESTADOR: analiza todas las series código+establecimiento y código+RED
# ════════════════════════════════════════════════════════════════════════════
def generar_pronosticos(ch, cols_consumo, precios_por_codigo, progreso_cada=500,
                         log=print):
    """Punto de entrada único usado por generar_dashboard_completo.py.

    ch: DataFrame de consumo histórico (Codigo, Fecha, columnas de
        establecimiento ya numéricas).
    cols_consumo: columnas de establecimiento que cuentan como consumo real
        (sin la bodega de tránsito).
    precios_por_codigo: dict código -> último precio ingresado (para ABC).

    Devuelve dict con "series" (lista de resultados por código+establecimiento
    y código+RED), "calidad" (resumen agregado de calidad de datos) y
    "clasificacion" (ABC/XYZ por código).
    """
    grillas, calidad_codigo, n_duplicados, n_negativos = construir_grillas(ch, cols_consumo)

    series_salida = []
    valor_anual, coef_variacion = {}, {}
    n_outliers_total = 0
    total_series = len(grillas) * (len(cols_consumo) + 1)
    procesadas = 0

    for codigo, grid in grillas.items():
        cal = calidad_codigo[codigo]
        for columna in cols_consumo + ["RED"]:
            procesadas += 1
            if progreso_cada and procesadas % progreso_cada == 0:
                log(f"  Pronósticos: {procesadas}/{total_series} series procesadas...")

            crudos = grid[columna].to_numpy(dtype=float)
            if np.all(np.isnan(crudos)) or np.nansum(crudos) <= 0:
                continue  # ese establecimiento nunca tuvo consumo de este producto

            n_outliers_total += int(detectar_outliers(crudos).sum())
            valores = np.nan_to_num(crudos, nan=0.0)

            try:
                resultado = analizar_serie(
                    valores,
                    permitir_sarima=(not SARIMA_SOLO_AGREGADO_RED) or (columna == "RED"),
                )
            except Exception as exc:  # nunca debe tumbar la corrida completa
                log(f"  Aviso: código {codigo} / {columna} falló ({exc}); se omite del pronóstico")
                continue

            if columna == "RED":
                ultimos12 = valores[-12:] if len(valores) >= 1 else valores
                valor_anual[codigo] = float(np.sum(ultimos12)) * float(precios_por_codigo.get(codigo, 0.0))
                base_cv = valores[-24:] if len(valores) >= MIN_MESES_ESTACIONAL else valores
                media_cv = float(np.mean(base_cv)) if len(base_cv) else 0.0
                coef_variacion[codigo] = (
                    float(np.std(base_cv) / media_cv)
                    if len(base_cv) >= MIN_MESES_NO_ESTACIONAL and media_cv > 0 else None
                )

            series_salida.append({
                "codigo": int(codigo),
                "establecimiento": columna,
                "n_meses": int(len(valores)),
                "meses_ausentes": cal["meses_ausentes"],
                **resultado,
            })

    clasificacion = clasificar_abc_xyz(valor_anual, coef_variacion)

    n_historia_insuficiente = sum(1 for c in calidad_codigo.values() if c["historia_insuficiente"])
    calidad = {
        "n_codigos": len(calidad_codigo),
        "n_series_con_pronostico": len(series_salida),
        "duplicados_codigo_fecha": n_duplicados,
        "valores_negativos_corregidos": n_negativos,
        "outliers_detectados": n_outliers_total,
        "codigos_historia_insuficiente": n_historia_insuficiente,
        "total_meses_ausentes": sum(c["meses_ausentes"] for c in calidad_codigo.values()),
    }

    log(f"Pronósticos generados para {len(series_salida)} series "
        f"({calidad['codigos_historia_insuficiente']} códigos con historia insuficiente)")

    return {"series": series_salida, "calidad": calidad, "clasificacion": clasificacion}
