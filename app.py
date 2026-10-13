# =============================================================================
#  SEGUIMIENTO DE ÍNDICES DE DESASTRES - PERÚ
#  Portada (/) + Índices climáticos (/indices-mensuales) + IRD (/ird) + América Latina (/latam)
#  Archivos junto a app.py: INDICE_IRD_FINAL_ROBUSTO.xlsx, ICEM.xlsx, EMDAT_LATAM.xlsx y portada.jpg
#  Pega todo en UNA celda de JupyterLab y ejecútala. Se abre en: http://127.0.0.1:8050
#
#  Librerías (una sola vez, en otra celda):
#      !pip install -U dash plotly pandas openpyxl
#
#  Secciones:
#    1. CONFIGURACIÓN   -> ruta del Excel, colores, KPIs, textos (cambia aquí)
#    2. FORMATO         -> números y nombres
#    3. DATOS           -> carga del Excel y cálculos
#    4. GRÁFICOS        -> una función por gráfico
#    5. DISEÑO          -> íconos, portada, tarjetas, paneles y estilos CSS
#    6. INTERACTIVIDAD  -> navegación entre páginas y callbacks
#    7. EJECUTAR
# =============================================================================

import json
import math
import os
import urllib.parse
import urllib.request

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import plotly.io as pio
from dash import Dash, Input, Output, State, ctx, dcc, html, no_update
from flask import send_from_directory


# =============================================================================
# 1. CONFIGURACIÓN
# =============================================================================
# El Excel está en la misma carpeta que app.py (raíz del repositorio de GitHub)
CARPETA = os.path.dirname(os.path.abspath(__file__)) if "__file__" in globals() else os.getcwd()
ARCHIVO_EXCEL = os.path.join(CARPETA, "INDICE_IRD_FINAL_ROBUSTO.xlsx")
HOJA_BASE = "BASE_CONSOLIDADA"

# Mapa: si no está en la carpeta del Excel, se descarga y se guarda ahí
URL_GEOJSON = ("https://raw.githubusercontent.com/juaneladio/peru-geojson/"
               "master/peru_departamental_simple.geojson")
RUTA_GEOJSON = os.path.join(os.path.dirname(ARCHIVO_EXCEL), "peru_departamental.geojson")
CLAVE_GEOJSON = "properties.NOMBDEP"

# Columnas del Excel
COL_ANIO = "AÑO"
COL_DPTO = "DPTO."
COL_IRD = "IRD_0_100"            # en tu Excel va de 0 a 10 (= IRD_0_1 x 10)
COL_FRECUENCIA = "I_Frecuencia"
COL_SEVERIDAD = "I_Severidad"
COL_INCIDENCIAS = "NUMERO_DE_INCIDENCIAS"
COL_DAMNIFICADOS = "DAMNIFICADOS"
ESCALA_IRD = 10

# Indicadores que se pueden elegir (nombre visible -> columna).
# Al cambiarlo, cambian el mapa, el ranking, el velocímetro y todos los gráficos de cuartiles.
INDICADORES = {
    "IRD": COL_IRD,
    "Frecuencia": COL_FRECUENCIA,
    "Severidad": COL_SEVERIDAD,
}

# Dimensiones de la severidad (nombre visible -> columna). El orden = orden de apilado.
DIMENSIONES = {
    "Humana": "Sev_HUMANA",
    "Social": "Sev_SOCIAL",
    "Infraestructura": "Sev_INFRAESTRUCTURA",
    "Productiva": "Sev_PRODUCTIVA",
}

NACIONAL = "PERÚ (nacional)"
NOMBRES_VISIBLES = {"ANCASH": "Áncash", "APURIMAC": "Apurímac", "HUANUCO": "Huánuco",
                    "JUNIN": "Junín", "SAN MARTIN": "San Martín", "MADRE DE DIOS": "Madre de Dios"}
MS_POR_ANIO = 1400  # velocidad del botón "Reproducir"

# KPIs de arriba: puedes poner cualquier columna del Excel
# icono: nombre de un ícono de la lista ICONOS (sección 5)
KPIS = [
    {"id": "incidencias", "columna": "NUMERO_DE_INCIDENCIAS", "titulo": "Emergencias registradas",
     "decimales": 0, "icono": "alerta"},
    {"id": "damnificados", "columna": "DAMNIFICADOS", "titulo": "Personas damnificadas",
     "decimales": 0, "icono": "personas"},
    {"id": "fallecidos", "columna": "FALLECIDOS", "titulo": "Personas fallecidas",
     "decimales": 0, "icono": "persona"},
    {"id": "viviendas", "columna": "VIVIENDAS DESTRUIDAS", "titulo": "Viviendas destruidas",
     "decimales": 0, "icono": "casa"},
    {"id": "cultivos", "columna": "HAS CULTIVO DESTRUIDO", "titulo": "Hectáreas de cultivo perdidas",
     "decimales": 0, "icono": "brote"},
]

# Colores. Si cambias uno, cámbialo también en CSS (sección 5).
C = {
    "pacifico": "#0b3a66", "pacifico_medio": "#1f6fb2", "niebla": "#edf3f9",
    "superficie": "#ffffff", "borde": "#d5e2ef", "tinta": "#0f2f50",
    "tinta_2": "#3f6488", "tinta_3": "#6f8cab", "rejilla": "#e6eef6",
    "contexto": "#a7bcd2", "mejora": "#2a78d6", "empeora": "#e34948", "azul_suave": "#e3eefa",
}
COLORES_DIMENSION = {"Humana": "#e87ba4", "Social": "#2a78d6",
                     "Infraestructura": "#1baf7a", "Productiva": "#4a3aa7"}
# Cuartiles (colores de tu mapa): 1 = 25 % más bajo ... 4 = 25 % más alto
COLORES_CUARTIL = {1: "#eae885", 2: "#fd8d3c", 3: "#e31a1c", 4: "#800026"}
NOMBRES_CUARTIL = {1: "Cuartil 1", 2: "Cuartil 2", 3: "Cuartil 3", 4: "Cuartil 4"}
BORDE_MAPA = "#696969"

# Textos y formato
TITULO = "Índice de Riesgo de Desastres"
SUBTITULO = "Emergencias por peligros naturales en el Perú, 2003–2025"
FUENTE_DATOS = "Fuente: INDICE_IRD_FINAL_ROBUSTO.xlsx (elaboración propia). Cuartil 4 = 25 % de departamentos con valor más alto en el año."
FUENTE = "Archivo"
URL_FUENTES = "https://fonts.googleapis.com/css2?family=Archivo:wdth,wght@62..125,300..800&display=swap"
SEPARADOR_DECIMAL = ","
SEPARADOR_MILES = "."
EVENTOS = {2017: "Niño Costero", 2023: "Ciclón Yaku"}   # anotaciones en la línea
ETIQUETAR_TODOS_LOS_PUNTOS = False   # True = valor en cada punto de la línea

# Portada: los 3 botones. "lista": True = el tablero ya funciona; False = muestra "En preparación".
TITULO_PORTADA = "Seguimiento de Índices de Desastres"
ANTETITULO_PORTADA = "Perú · Gestión del riesgo de desastres"
SUBTITULO_PORTADA = ("Tres tableros para seguir las amenazas climáticas, el riesgo en los departamentos "
                     "y la situación de América Latina.")
RUTA_IRD = "/ird"
SECCIONES = [
    {"ruta": "/indices-mensuales", "icono": "ola", "lista": True,
     "titulo": "Índices climáticos mensuales",
     "descripcion": "Seguimiento mes a mes de los índices que miden si el mar está en condición "
                    "cálida, neutra o fría.",
     "etiquetas": ["LABCOS", "ITCP", "ICEN", "EHNOS"]},
    {"ruta": RUTA_IRD, "icono": "medidor", "lista": True,
     "titulo": "Índice de Riesgo de Desastres",
     "descripcion": "Frecuencia y severidad de las emergencias en los 25 departamentos, "
                    "con mapa, ranking y evolución anual.",
     "etiquetas": ["IRD", "Frecuencia", "Severidad"]},
    {"ruta": "/latam", "icono": "globo", "lista": True,
     "titulo": "Estadísticas de América Latina",
     "descripcion": "Frecuencia, costo económico y tipo de los desastres naturales en 26 países "
                    "de la región, para comparar al Perú con sus vecinos.",
     "etiquetas": ["EM-DAT", "Frecuencia", "Costo % PBI"]},
]
ARCHIVO_PORTADA = "portada.jpg"      # imagen de fondo de la portada (junto a app.py)

# ---- Índices climáticos mensuales (archivo ICEM.xlsx) -------------------------
ARCHIVO_CLIMA = os.path.join(CARPETA, "ICEM.xlsx")
HOJA_CLIMA = 0                       # 0 = primera hoja del Excel
COL_CLIMA_ANIO, COL_CLIMA_MES = "yyyy", "mm"
RUTA_CLIMA = "/indices-mensuales"
MESES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto",
         "Septiembre", "Octubre", "Noviembre", "Diciembre"]
COLORES_CALIDOS = ["#f03b20", "#fd8d3c", "#fecc5c", "#ffffb2"]   # de más fuerte a más débil
COLORES_FRIOS = ["#bdd7e7", "#6baed6", "#3182bd"]                # de más débil a más fuerte
COLOR_NEUTRO = "#ffffff"
# Cada índice: columna del Excel, nombre, decimales y criterio (tus tablas de "CRITERIO").
#   calidas: (magnitud, umbral) de la más fuerte a la más débil -> cálida si valor > umbral
#   frias:   (magnitud, umbral) de la más débil a la más fuerte -> fría si valor < umbral
#   incluye_limite=True cuando el criterio usa ≥ y ≤ (como el MEI v2)
INDICES_CLIMA = {
    "LABCOS": {"columna": "LABCOS", "nombre": "Índice LABCOS", "decimales": 1, "decimales_criterio": 2,
               "calidas": [("Muy fuerte", 2.77), ("Fuerte", 1.37), ("Moderada", 0.78), ("Débil", 0.27)],
               "frias": [("Débil", -0.78), ("Moderada", -1.10), ("Fuerte", -1.31)]},
    "ITCP": {"columna": "ITCP", "nombre": "Índice Térmico Costero Peruano", "decimales": 1,
             "calidas": [("Cálida", 0.4)], "frias": [("Fría", -0.6)]},
    "ICEN": {"columna": "ICE", "nombre": "Índice Costero El Niño", "decimales": 2,
             "calidas": [("Extraordinaria", 3.5), ("Fuerte", 2.1), ("Moderada", 1.3), ("Débil", 0.5)],
             "frias": [("Débil", -0.7), ("Moderada", -1.1), ("Fuerte", -1.3)]},
    # EHNOS: se usa el criterio del MEI v2 (±0,5). Si su criterio es otro, cámbialo aquí.
    "EHNOS": {"columna": "EHNOS", "nombre": "Índice EHNOS", "decimales": 2,
              "calidas": [("Cálida", 0.5)], "frias": [("Fría", -0.5)], "incluye_limite": True},
}
COLOR_CALIDO, COLOR_FRIO, COLOR_NEUTRO_GRIS = "#f03b20", "#3182bd", "#c9d6e3"
MESES_VENTANA = 36       # meses que muestran "Variación a 12 meses" y "¿Coinciden los índices?"

# ---- Estadísticas de América Latina (archivo EMDAT_LATAM.xlsx, ya procesado) ---
ARCHIVO_LATAM = os.path.join(CARPETA, "EMDAT_LATAM.xlsx")
RUTA_LATAM = "/latam"
PAIS_INICIAL = "Perú"
ULTIMO_ANIO_COSTO = 2025     # EM-DAT todavía no tiene costos de 2026
CORTE_PERIODOS = 2002        # Gráfico de dos periodos: 1980-2002 y 2003-2026
# Colores de la ayuda memoria
COLORES_SUBGRUPO = {"Biológico": "#156082", "Climatológico": "#7F7F7F", "Geofísico": "#F2AA84",
                    "Hidrológico": "#D9D9D9", "Meteorológico": "#94DCF8"}
ORDEN_SUBGRUPOS_PAIS = ["Hidrológico", "Geofísico", "Meteorológico", "Biológico", "Climatológico"]
COLORES_CUARTIL_LATAM = {1: "#F2F089", 2: "#FD8D3C", 3: "#E31A1C", 4: "#800026"}
COLOR_PERIODO_1, COLOR_PERIODO_2 = "#7F7F7F", "#C00000"
COLOR_DISPERSION = "#156082"
COLOR_TIPO_PAIS = "#F2AA84"
INDICADORES_LATAM = {"eventos": "Número de desastres", "costo_pbi": "Costo económico (% del PBI)",
                     "fallecidos": "Personas fallecidas"}


# =============================================================================
# 2. FORMATO
# =============================================================================
def es_vacio(valor):
    if valor is None:
        return True
    try:
        return math.isnan(float(valor))
    except (TypeError, ValueError):
        return False


def numero(valor, decimales=0):
    """1234567.8 -> '1.234.567,8'"""
    if es_vacio(valor):
        return "—"
    texto = f"{float(valor):,.{decimales}f}"
    return texto.replace(",", "§").replace(".", SEPARADOR_DECIMAL).replace("§", SEPARADOR_MILES)


def variacion_pct(actual, anterior):
    if es_vacio(actual) or es_vacio(anterior) or float(anterior) == 0:
        return None
    return (float(actual) - float(anterior)) / float(anterior) * 100


def con_signo(valor, decimales=0):
    if es_vacio(valor):
        return "—"
    signo = "+" if valor > 0 else ("−" if valor < 0 else "")
    return f"{signo}{numero(abs(valor), decimales)}"


def nombre_visible(dpto):
    if dpto == NACIONAL:
        return "Perú"
    return NOMBRES_VISIBLES.get(dpto, dpto.title().replace(" De ", " de "))


def nombre_indicador(col):
    return next((k for k, v in INDICADORES.items() if v == col), col)


def decimales_de(col):
    """Los tres indicadores van de 0 a 10: 1 decimal en las etiquetas."""
    return 1


def hex_a_rgba(color_hex, alfa):
    color_hex = color_hex.lstrip("#")
    r, g, b = (int(color_hex[i:i + 2], 16) for i in (0, 2, 4))
    return f"rgba({r},{g},{b},{alfa})"


# =============================================================================
# 3. DATOS
# =============================================================================
def col_cuartil(col):
    return f"CUARTIL_{col}"


def col_puesto(col):
    return f"PUESTO_{col}"


# Partes de cada indicador para el gráfico "¿Qué explica…?" (nombre -> columna)
APORTES = {
    COL_IRD: {n: f"APORTE_IRD_{n}" for n in DIMENSIONES},       # Frecuencia x dimensión / 4 x 10
    COL_SEVERIDAD: {n: f"APORTE_SEV_{n}" for n in DIMENSIONES},  # dimensión / 4 x 10
    COL_FRECUENCIA: {"Frecuencia": COL_FRECUENCIA},              # no se divide
}
COLUMNAS_INDICE = (list(INDICADORES.values()) + list(DIMENSIONES.values())
                   + list(APORTES[COL_IRD].values()) + list(APORTES[COL_SEVERIDAD].values()))
COLUMNAS_CONTEO = list(dict.fromkeys([COL_INCIDENCIAS, COL_DAMNIFICADOS] + [k["columna"] for k in KPIS]))


def cargar_base():
    df = pd.read_excel(ARCHIVO_EXCEL, sheet_name=HOJA_BASE)
    df.columns = df.columns.astype(str).str.strip()
    df[COL_DPTO] = df[COL_DPTO].astype(str).str.strip().str.upper()
    df[COL_ANIO] = pd.to_numeric(df[COL_ANIO], errors="coerce")
    df = df.dropna(subset=[COL_ANIO]).copy()
    df[COL_ANIO] = df[COL_ANIO].astype(int)
    for col in COLUMNAS_CONTEO:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    for col in list(INDICADORES.values()) + list(DIMENSIONES.values()):
        df[col] = pd.to_numeric(df[col], errors="coerce")

    # Aporte exacto de cada dimensión (con frecuencia y severidad originales de 0 a 1)
    for nombre, col in DIMENSIONES.items():
        df[APORTES[COL_SEVERIDAD][nombre]] = df[col] / len(DIMENSIONES) * ESCALA_IRD
        df[APORTES[COL_IRD][nombre]] = df[COL_FRECUENCIA] * df[col] / len(DIMENSIONES) * ESCALA_IRD

    # Frecuencia y severidad x 10 -> misma escala que el IRD (0 a 10).
    # El IRD NO se recalcula: se usa tal cual viene del Excel.
    df[COL_FRECUENCIA] = df[COL_FRECUENCIA] * ESCALA_IRD
    df[COL_SEVERIDAD] = df[COL_SEVERIDAD] * ESCALA_IRD

    # Cuartil (1 a 4) y puesto (1 = el más alto) de cada indicador, año por año
    for col in INDICADORES.values():
        por_anio = df.groupby(COL_ANIO)[col]
        df[col_cuartil(col)] = por_anio.transform(lambda s: pd.qcut(s.rank(method="first"), 4, labels=False) + 1)
        df[col_puesto(col)] = por_anio.rank(ascending=False, method="min")

    df["NOMBRE"] = df[COL_DPTO].map(nombre_visible)
    return df.sort_values([COL_ANIO, COL_DPTO]).reset_index(drop=True)


def cargar_geojson():
    if os.path.exists(RUTA_GEOJSON):
        with open(RUTA_GEOJSON, encoding="utf-8") as archivo:
            return json.load(archivo)
    with urllib.request.urlopen(URL_GEOJSON, timeout=30) as respuesta:
        texto = respuesta.read().decode("utf-8")
    try:  # se guarda para no descargarlo otra vez
        with open(RUTA_GEOJSON, "w", encoding="utf-8") as archivo:
            archivo.write(texto)
    except OSError:
        pass
    return json.loads(texto)


def lista_anios(df):
    return sorted(int(a) for a in df[COL_ANIO].unique())


def serie_anual(df, dpto):
    """Una fila por año. Nacional: suma conteos y promedia índices."""
    if dpto == NACIONAL:
        reglas = {c: "sum" for c in COLUMNAS_CONTEO}
        reglas.update({c: "mean" for c in COLUMNAS_INDICE})
        serie = df.groupby(COL_ANIO).agg(reglas)
    else:
        serie = df.loc[df[COL_DPTO] == dpto].set_index(COL_ANIO)
    return serie.reindex(lista_anios(df))


def corte_anual(df, anio, col):
    """Todos los departamentos en un año, del valor más alto al más bajo del indicador."""
    corte = df.loc[df[COL_ANIO] == anio].copy()
    return corte.sort_values(col, ascending=False).reset_index(drop=True)


def rango_departamentos(df, col):
    agrupado = df.groupby(COL_ANIO)[col]
    return pd.DataFrame({"p25": agrupado.quantile(0.25), "p75": agrupado.quantile(0.75)})


def variacion_anual(df, anio, col):
    actual = df.loc[df[COL_ANIO] == anio].set_index(COL_DPTO)[col]
    previo = df.loc[df[COL_ANIO] == anio - 1].set_index(COL_DPTO)[col]
    tabla = pd.DataFrame({"previo": previo, "actual": actual}).dropna()
    tabla["cambio"] = tabla["actual"] - tabla["previo"]
    return tabla.sort_values("cambio")


def umbrales_cuartil(df, anio, col):
    """[(cuartil, desde, hasta), ...] del indicador en un año (bandas del velocímetro)."""
    corte = corte_anual(df, anio, col)
    limites = corte.groupby(col_cuartil(col))[col].min().sort_index()
    maximo = float(corte[col].max())
    tramos, desde, cuartiles = [], 0.0, list(limites.index)
    for i, q in enumerate(cuartiles):
        hasta = float(limites.iloc[i + 1]) if i + 1 < len(cuartiles) else maximo
        tramos.append((int(q), desde, hasta))
        desde = hasta
    return tramos


def resumen(df, dpto, anio, col):
    """Valor del indicador, valor del año anterior, puesto y cuartil."""
    serie = serie_anual(df, dpto)
    corte = corte_anual(df, anio, col)
    r = {"anio": anio, "es_nacional": dpto == NACIONAL, "n_dptos": len(corte),
         "valor": np.nan, "previo": np.nan, "puesto": None, "cuartil": None,
         "promedio_nacional": corte[col].mean()}
    if anio in serie.index and not es_vacio(serie.loc[anio, col]):
        r["valor"] = float(serie.loc[anio, col])
        if (anio - 1) in serie.index:
            r["previo"] = serie.loc[anio - 1, col]
        if not r["es_nacional"]:
            r["puesto"] = int(serie.loc[anio, col_puesto(col)])
            r["cuartil"] = int(serie.loc[anio, col_cuartil(col)])
    return r


def centroides(geojson, clave="NOMBDEP"):
    """Centro aproximado de cada departamento (para las etiquetas del mapa)."""
    resultado = {}
    for feature in geojson["features"]:
        geom = feature["geometry"]
        poligonos = geom["coordinates"] if geom["type"] == "MultiPolygon" else [geom["coordinates"]]
        mejor, mejor_area = None, -1.0
        for poligono in poligonos:
            anillo = np.asarray(poligono[0], dtype=float)
            x, y = anillo[:, 0], anillo[:, 1]
            cruz = x * np.roll(y, -1) - np.roll(x, -1) * y
            area = cruz.sum() / 2
            if area != 0 and abs(area) > mejor_area:
                cx = ((x + np.roll(x, -1)) * cruz).sum() / (6 * area)
                cy = ((y + np.roll(y, -1)) * cruz).sum() / (6 * area)
                mejor, mejor_area = (float(cx), float(cy)), abs(area)
        resultado[feature["properties"][clave]] = mejor
    return resultado


# ---- Datos de los índices climáticos ------------------------------------------
def a_numero(serie):
    """'4.7*' -> 4.7 ; '0,35' -> 0.35 ; vacío -> NaN"""
    texto = serie.astype(str).str.strip().str.replace("*", "", regex=False).str.replace(",", ".", regex=False)
    return pd.to_numeric(texto, errors="coerce") + 0.0   # + 0.0 convierte -0.0 en 0.0


def cargar_clima():
    """Una fila por mes. Si el Excel no está, devuelve None (la página queda «En preparación»)."""
    if not os.path.exists(ARCHIVO_CLIMA):
        return None
    df = pd.read_excel(ARCHIVO_CLIMA, sheet_name=HOJA_CLIMA)
    df.columns = df.columns.astype(str).str.strip()
    df = df.rename(columns={COL_CLIMA_ANIO: "anio", COL_CLIMA_MES: "mes"})
    df["anio"] = pd.to_numeric(df["anio"], errors="coerce")
    df["mes"] = pd.to_numeric(df["mes"], errors="coerce")
    df = df.dropna(subset=["anio", "mes"]).astype({"anio": int, "mes": int})
    for clave, cfg in INDICES_CLIMA.items():
        original = df[cfg["columna"]]
        df[clave] = a_numero(original)
        df[clave + "_prelim"] = original.astype(str).str.contains(r"\*", regex=True)
    df["fecha"] = pd.to_datetime(dict(year=df["anio"], month=df["mes"], day=1))
    df = df.sort_values("fecha").reset_index(drop=True)
    df["etiqueta"] = [f"{MESES[m - 1][:3]} {a}" for a, m in zip(df["anio"], df["mes"])]
    return df


def segmentos_clima(cfg):
    """Categorías de izquierda (más cálida) a derecha (más fría), con su color y su rango."""
    calidas, frias = cfg["calidas"], cfg["frias"]
    varias = len(calidas) > 1
    neutro = dict(grupo="Neutro", magnitud="Neutro", color=COLOR_NEUTRO,
                  desde=frias[0][1], hasta=calidas[-1][1])
    extra_c = calidas[0][1] - calidas[1][1] if varias else neutro["hasta"] - neutro["desde"]
    extra_f = frias[-2][1] - frias[-1][1] if len(frias) > 1 else neutro["hasta"] - neutro["desde"]
    segmentos = []
    for i, (magnitud, umbral) in enumerate(calidas):
        hasta = calidas[i - 1][1] if i else umbral + extra_c
        segmentos.append(dict(grupo="Condición cálida", magnitud=magnitud, color=COLORES_CALIDOS[i],
                              desde=umbral, hasta=hasta, abierto=(i == 0)))
    segmentos.append(dict(neutro, abierto=False))
    colores_frios = COLORES_FRIOS[-len(frias):]
    for j, (magnitud, umbral) in enumerate(frias):
        desde = frias[j + 1][1] if j + 1 < len(frias) else umbral - extra_f
        segmentos.append(dict(grupo="Condición fría", magnitud=magnitud, color=colores_frios[j],
                              desde=desde, hasta=umbral, abierto=(j == len(frias) - 1)))
    for seg in segmentos:
        seg["texto"] = seg["grupo"] if not varias or seg["grupo"] == "Neutro" else f"{seg['grupo']} {seg['magnitud'].lower()}"
    return segmentos


def clasificar_clima(valor, cfg):
    """Posición (0 = más cálida) de la categoría del valor, o None si no hay dato."""
    if es_vacio(valor):
        return None
    incluye = cfg.get("incluye_limite", False)
    for i, (_, umbral) in enumerate(cfg["calidas"]):
        if valor > umbral or (incluye and valor == umbral):
            return i
    n_calidas = len(cfg["calidas"])
    for j in range(len(cfg["frias"]) - 1, -1, -1):  # de la más fuerte a la más débil
        umbral = cfg["frias"][j][1]
        if valor < umbral or (incluye and valor == umbral):
            return n_calidas + 1 + j
    return n_calidas  # neutro


def valor_clima(valor, decimales):
    return "—" if es_vacio(valor) else numero(round(float(valor), decimales) + 0.0, decimales)


def umbral_texto(valor, decimales=1):
    texto = f"{valor:+.{decimales}f}".replace(".", SEPARADOR_DECIMAL)
    return texto.replace("-", "−")


def intervalo_texto(seg, cfg):
    incluye, dec = cfg.get("incluye_limite", False), cfg.get("decimales_criterio", 1)
    if seg["grupo"] == "Condición cálida" and seg["abierto"]:
        return f"{'≥' if incluye else '>'} {umbral_texto(seg['desde'], dec)}"
    if seg["grupo"] == "Condición fría" and seg["abierto"]:
        return f"{'≤' if incluye else '<'} {umbral_texto(seg['hasta'], dec)}"
    return f"{umbral_texto(seg['desde'], dec)} a {umbral_texto(seg['hasta'], dec)}"


# ---- Datos de América Latina (EM-DAT) -----------------------------------------
def cargar_latam():
    """Lee EMDAT_LATAM.xlsx (hojas «eventos» y «pbi»). Si no está, devuelve None."""
    if not os.path.exists(ARCHIVO_LATAM):
        return None
    eventos = pd.read_excel(ARCHIVO_LATAM, sheet_name="eventos")
    pbi = pd.read_excel(ARCHIVO_LATAM, sheet_name="pbi").set_index("pais")
    return {"eventos": eventos, "pbi": pbi}


def eventos_periodo(latam, a0, a1):
    e = latam["eventos"]
    return e[e["anio"].between(a0, a1)]


def resumen_paises(latam, a0, a1):
    """Una fila por país: desastres, fallecidos, afectados, daño (miles US$) y costo (% del PBI)."""
    paises = latam["eventos"].drop_duplicates("pais").set_index("pais")[["iso3", "zona"]]
    tabla = eventos_periodo(latam, a0, a1).groupby("pais").agg(
        eventos=("id", "size"), fallecidos=("fallecidos", "sum"), afectados=("afectados", "sum"),
        dano_miles=("dano_miles_usd", "sum"))
    tabla = paises.join(tabla).fillna({"eventos": 0, "fallecidos": 0, "afectados": 0, "dano_miles": 0})
    tabla["costo_pbi"] = tabla["dano_miles"] / (latam["pbi"]["pbi_millones_usd"] * 1000) * 100
    return tabla


def cuartiles_latam(serie):
    return pd.qcut(serie.rank(method="first"), 4, labels=False) + 1


def valor_latam(valor, indicador):
    return f"{numero(valor, 1)} %" if indicador == "costo_pbi" else numero(valor, 0)


def texto_periodo(a0, a1):
    return f"{a0}–{a1}" if a0 != a1 else str(a0)


# =============================================================================
# 4. GRÁFICOS  (una función por gráfico; todas devuelven una figura de Plotly)
#    "col" es la columna del indicador elegido (IRD, Frecuencia o Severidad)
# =============================================================================
FUENTE_PLOTLY = f"{FUENTE}, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif"

# Estilo común de todos los gráficos
pio.templates["ird"] = go.layout.Template(layout=dict(
    font=dict(family=FUENTE_PLOTLY, size=12, color=C["tinta_2"]),
    paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
    separators=SEPARADOR_DECIMAL + SEPARADOR_MILES,
    margin=dict(l=8, r=8, t=8, b=8),
    hoverlabel=dict(bgcolor=C["superficie"], bordercolor=C["borde"],
                    font=dict(family=FUENTE_PLOTLY, size=12, color=C["tinta"])),
    xaxis=dict(automargin=True, showgrid=False, zeroline=False, ticks="", showline=True,
               linecolor=C["borde"], tickfont=dict(color=C["tinta_3"], size=11),
               title=dict(font=dict(color=C["tinta_3"], size=11))),
    yaxis=dict(automargin=True, gridcolor=C["rejilla"], zeroline=False, ticks="", showline=False,
               tickfont=dict(color=C["tinta_3"], size=11),
               title=dict(font=dict(color=C["tinta_3"], size=11))),
    legend=dict(orientation="h", x=0, xanchor="left", y=1.02, yanchor="bottom",
                font=dict(color=C["tinta_2"], size=11), bgcolor="rgba(0,0,0,0)"),
    colorway=list(COLORES_DIMENSION.values()), bargap=0.22, barcornerradius=3,
))
pio.templates.default = "ird"

CONFIG_GRAFICO = {"displaylogo": False, "responsive": True,
                  "modeBarButtonsToRemove": ["select2d", "lasso2d", "autoScale2d", "toggleSpikelines"],
                  "toImageButtonOptions": {"format": "png", "scale": 2}}
CONFIG_MINI = {"displayModeBar": False}


def figura_vacia(mensaje):
    fig = go.Figure()
    fig.add_annotation(text=mensaje, x=0.5, y=0.5, xref="paper", yref="paper",
                       showarrow=False, font=dict(size=13, color=C["tinta_3"]))
    fig.update_xaxes(visible=False)
    fig.update_yaxes(visible=False)
    return fig


def marcar_anio(fig, anio):
    fig.add_vline(x=anio, line=dict(color=C["pacifico_medio"], width=1, dash="dot"), layer="below")


def opacidades(etiquetas, seleccion, atenuado=0.35):
    return [1.0 if e == seleccion else atenuado for e in etiquetas]


def colores_cuartil(cuartiles):
    return [COLORES_CUARTIL.get(int(q), C["contexto"]) if not es_vacio(q) else C["contexto"] for q in cuartiles]


# ---- 4.1 Mini-línea de cada KPI ---------------------------------------------
def grafico_sparkline(serie, anio, decimales=0):
    fig = go.Figure()
    fig.add_scatter(x=serie.index, y=serie.values, mode="lines",
                    line=dict(color=C["pacifico_medio"], width=1.6, shape="spline", smoothing=0.6),
                    fill="tozeroy", fillcolor=hex_a_rgba(C["pacifico_medio"], 0.10),
                    hovertemplate=f"%{{x}}: %{{y:,.{decimales}f}}<extra></extra>")
    if anio in serie.index and not es_vacio(serie.loc[anio]):
        fig.add_scatter(x=[anio], y=[serie.loc[anio]], mode="markers", hoverinfo="skip",
                        marker=dict(size=8, color=C["pacifico"], line=dict(color="white", width=1.5)))
    fig.update_layout(height=52, margin=dict(l=0, r=0, t=4, b=0), showlegend=False, hovermode="x",
                      xaxis=dict(visible=False), yaxis=dict(visible=False, rangemode="tozero"))
    return fig


# ---- 4.2 ¿Qué explica el indicador cada año? (barras apiladas) --------------
def grafico_estratos_anuales(serie, anio, col):
    anios = list(serie.index)
    opac = opacidades(anios, anio, atenuado=0.4)
    fig = go.Figure()
    for nombre, columna in APORTES[col].items():
        color = COLORES_DIMENSION.get(nombre, C["pacifico_medio"])
        fig.add_bar(x=anios, y=serie[columna], name=nombre,
                    marker=dict(color=color, opacity=opac, line=dict(color=C["superficie"], width=1)),
                    hovertemplate=f"{nombre}: %{{y:.2f}}<extra></extra>")
    total = serie[col]
    dec = decimales_de(col)
    anio_max = total.idxmax() if total.notna().any() else None
    textos = [numero(v, dec) if a in (anio, anio_max) and not es_vacio(v) else ""
              for a, v in zip(anios, total)]
    fig.add_scatter(x=anios, y=total, mode="text", text=textos, textposition="top center",
                    textfont=dict(size=11, color=C["tinta"]), showlegend=False,
                    hovertemplate=f"<b>{nombre_indicador(col)} total: %{{y:.2f}}</b><extra></extra>")
    tope = total.max() if total.notna().any() else 1
    fig.update_layout(barmode="stack", hovermode="x unified", showlegend=len(APORTES[col]) > 1,
                      margin=dict(l=4, r=4, t=30, b=4), legend=dict(y=1.03, traceorder="normal"),
                      yaxis=dict(title=f"Aporte a {nombre_indicador(col)}", range=[0, tope * 1.15]),
                      xaxis=dict(dtick=2, tick0=anios[0], tickangle=0, tickfont=dict(size=10)))
    return fig


# ---- 4.3 Ranking de departamentos (barras por cuartil) ----------------------
def grafico_ranking(corte, dpto_sel, col):
    datos = corte.sort_values(col)  # el mayor queda arriba
    hay_seleccion = dpto_sel != NACIONAL
    etiquetas = [f"<b>{n}</b>" if d == dpto_sel else n for d, n in zip(datos[COL_DPTO], datos["NOMBRE"])]
    custom = pd.DataFrame({"n": datos["NOMBRE"], "p": datos[col_puesto(col)].astype("Int64").astype(str),
                           "q": datos[col_cuartil(col)].astype("Int64").astype(str)}).to_numpy(dtype=object)
    fig = go.Figure(go.Bar(
        y=datos[COL_DPTO], x=datos[col], orientation="h",
        marker=dict(color=colores_cuartil(datos[col_cuartil(col)]),
                    opacity=opacidades(datos[COL_DPTO], dpto_sel) if hay_seleccion else 1.0),
        text=[numero(v, decimales_de(col)) for v in datos[col]], textposition="outside",
        textfont=dict(size=10, color=C["tinta"]), cliponaxis=False, customdata=custom,
        hovertemplate=("<b>%{customdata[0]}</b><br>" + nombre_indicador(col) + ": %{x:.2f}"
                       "<br>Puesto: %{customdata[1]}<br>Cuartil: %{customdata[2]}<extra></extra>")))
    fig.update_layout(
        bargap=0.18, showlegend=False, margin=dict(l=4, r=4, t=4, b=4),
        xaxis=dict(visible=False, range=[0, corte[col].max() * 1.18]),
        yaxis=dict(tickvals=datos[COL_DPTO], ticktext=etiquetas, showline=True, linecolor=C["borde"],
                   gridcolor="rgba(0,0,0,0)", tickfont=dict(size=10, color=C["tinta"])))
    return fig


# ---- 4.4 Mapa del Perú (siempre por cuartil del indicador) ------------------
def grafico_mapa(corte, geojson, centros, dpto_sel, col):
    fig = go.Figure()
    hover = ("<b>%{customdata[0]}</b><br>" + nombre_indicador(col) + ": %{customdata[1]:.2f}"
             "<br>Puesto: %{customdata[2]} de " + str(len(corte)) + "<extra></extra>")
    for q in sorted(COLORES_CUARTIL, reverse=True):  # Cuartil 4 primero en la leyenda
        sub = corte[corte[col_cuartil(col)] == q]
        color = COLORES_CUARTIL[q]
        fig.add_trace(go.Choropleth(
            geojson=geojson, featureidkey=CLAVE_GEOJSON, locations=sub[COL_DPTO], z=[q] * len(sub),
            colorscale=[[0, color], [1, color]], showscale=False, name=NOMBRES_CUARTIL[q], showlegend=True,
            marker=dict(line=dict(color=BORDE_MAPA, width=0.6)),
            customdata=pd.DataFrame({"n": sub["NOMBRE"], "v": sub[col],
                                     "p": sub[col_puesto(col)].astype(int)}).to_numpy(dtype=object),
            hovertemplate=hover))
    fig.update_layout(legend=dict(orientation="v", x=0.02, y=0.04, yanchor="bottom",
                                  title=dict(text="Cuartil", font=dict(size=11))))

    if dpto_sel != NACIONAL:  # contorno del seleccionado
        fig.add_trace(go.Choropleth(
            geojson=geojson, featureidkey=CLAVE_GEOJSON, locations=[dpto_sel], z=[1],
            colorscale=[[0, "rgba(0,0,0,0)"], [1, "rgba(0,0,0,0)"]], showscale=False,
            showlegend=False, hoverinfo="skip", marker=dict(line=dict(color=C["tinta"], width=2.6))))

    # Insignias 1, 2, 3 en los departamentos con el valor más alto
    filas = corte.set_index(COL_DPTO)
    top3 = [d for d in corte[COL_DPTO].head(3) if d in centros]
    fig.add_trace(go.Scattergeo(
        lon=[centros[d][0] for d in top3], lat=[centros[d][1] for d in top3],
        text=[str(int(filas.loc[d, col_puesto(col)])) for d in top3],
        mode="markers+text", textposition="middle center", showlegend=False, hoverinfo="skip",
        marker=dict(size=19, color=C["superficie"], line=dict(color=C["tinta"], width=1.5)),
        textfont=dict(size=11, color=C["tinta"], family=FUENTE_PLOTLY)))

    if dpto_sel in centros and dpto_sel in filas.index:  # nombre del seleccionado
        lon, lat = centros[dpto_sel]
        fig.add_trace(go.Scattergeo(
            lon=[lon], lat=[lat], text=[f"<b>{filas.loc[dpto_sel, 'NOMBRE']}</b>"], mode="text",
            textposition="bottom center" if dpto_sel in top3 else "middle center",
            showlegend=False, hoverinfo="skip",
            textfont=dict(size=12, color=C["tinta"], family=FUENTE_PLOTLY,
                          shadow="0 0 3px #fff, 0 0 3px #fff, 0 0 2px #fff")))

    fig.update_geos(fitbounds="locations", visible=False, bgcolor="rgba(0,0,0,0)",
                    projection_type="mercator")
    fig.update_layout(margin=dict(l=0, r=0, t=0, b=0), dragmode=False)
    return fig


# ---- 4.5 Evolución del indicador con tendencia ------------------------------
def grafico_tendencia(serie_nac, banda, anio, col, serie_dpto=None, nombre_dpto=""):
    fig = go.Figure()
    fig.add_scatter(x=banda.index, y=banda["p75"], mode="lines", line=dict(width=0),
                    hoverinfo="skip", showlegend=False)
    fig.add_scatter(x=banda.index, y=banda["p25"], mode="lines", line=dict(width=0), fill="tonexty",
                    fillcolor=hex_a_rgba(C["contexto"], 0.22), hoverinfo="skip",
                    name="Rango típico de dptos. (P25–P75)")
    if serie_dpto is not None:
        principal, nombre = serie_dpto[col], nombre_dpto
        fig.add_scatter(x=serie_nac.index, y=serie_nac[col], name="Promedio nacional",
                        mode="lines", line=dict(color=C["contexto"], width=1.6),
                        hovertemplate="Promedio nacional: %{y:.2f}<extra></extra>")
    else:
        principal, nombre = serie_nac[col], "Promedio nacional"

    fig.add_scatter(x=principal.index, y=principal, name=nombre, mode="lines+markers",
                    line=dict(color=C["pacifico"], width=2.4),
                    marker=dict(size=6, color=C["pacifico"], line=dict(color="white", width=1)),
                    hovertemplate=f"{nombre}: %{{y:.2f}}<extra></extra>")

    validos = principal.dropna()
    if len(validos) > 2:  # tendencia lineal
        x = validos.index.astype(float).to_numpy()
        pendiente, intercepto = np.polyfit(x, validos.to_numpy(dtype=float), 1)
        fig.add_scatter(x=validos.index, y=intercepto + pendiente * x, mode="lines",
                        name=f"Tendencia ({con_signo(pendiente, 3)} por año)",
                        line=dict(color=C["tinta_2"], width=1.4, dash="dot"), hoverinfo="skip")

    if ETIQUETAR_TODOS_LOS_PUNTOS:
        claves = set(validos.index)
    else:
        claves = {validos.index.min(), validos.index.max(), validos.idxmax(), anio} if len(validos) else set()
    for a, v in principal.items():
        if a in claves and not es_vacio(v):
            fig.add_annotation(x=a, y=v, text=numero(v, decimales_de(col)), showarrow=False, yshift=15,
                               font=dict(size=11, color=C["tinta"]))

    if anio in principal.index and not es_vacio(principal.loc[anio]):
        fig.add_scatter(x=[anio], y=[principal.loc[anio]], mode="markers", showlegend=False,
                        hoverinfo="skip",
                        marker=dict(size=13, color=C["superficie"], line=dict(color=C["pacifico"], width=3)))
    marcar_anio(fig, anio)

    for anio_evento, texto in EVENTOS.items():
        if anio_evento in principal.index and not es_vacio(principal.loc[anio_evento]):
            fig.add_annotation(x=anio_evento, y=principal.loc[anio_evento], text=texto,
                               ax=-46, ay=-34, showarrow=True, arrowhead=0, arrowwidth=1,
                               arrowcolor=C["tinta_3"], font=dict(size=10, color=C["tinta_2"]),
                               bgcolor=hex_a_rgba(C["superficie"], 0.85))

    tope = np.nanmax([principal.max(), banda["p75"].max()])
    fig.update_layout(hovermode="x unified", margin=dict(l=4, r=4, t=40, b=4),
                      legend=dict(y=1.06, traceorder="normal"),
                      xaxis=dict(dtick=2, tick0=int(principal.index.min()), tickfont=dict(size=10)),
                      yaxis=dict(title=nombre_indicador(col), range=[0, tope * 1.22]))
    return fig


# ---- 4.6 Velocímetro (KPI principal) ----------------------------------------
def grafico_velocimetro(r, umbrales, tope):
    if es_vacio(r["valor"]):
        return figura_vacia(f"Sin datos para {r['anio']}")
    tramos = [dict(range=[desde, hasta], color=COLORES_CUARTIL[q]) for q, desde, hasta in umbrales]
    if tramos:
        tramos[-1]["range"][1] = tope
    gauge = dict(shape="angular",
                 axis=dict(range=[0, tope], tickwidth=1, tickcolor=C["borde"],
                           tickfont=dict(size=10, color=C["tinta_3"]), nticks=6),
                 bar=dict(color=C["tinta"], thickness=0.22),
                 bgcolor=C["rejilla"], borderwidth=0, steps=tramos)
    if not r["es_nacional"]:  # marca azul = promedio nacional
        gauge["threshold"] = dict(line=dict(color=C["pacifico_medio"], width=3),
                                  thickness=0.9, value=r["promedio_nacional"])
    hay_previo = not es_vacio(r["previo"])
    indicador = go.Indicator(mode="gauge+number+delta" if hay_previo else "gauge+number",
                             value=r["valor"], gauge=gauge, domain=dict(x=[0, 1], y=[0, 1]),
                             number=dict(valueformat=".2f", font=dict(size=34, color=C["tinta"])))
    if hay_previo:  # ▲ rojo si sube, ▼ azul si baja
        indicador.delta = dict(reference=float(r["previo"]), valueformat=".2f", position="bottom",
                               increasing=dict(color=C["empeora"], symbol="▲ "),
                               decreasing=dict(color=C["mejora"], symbol="▼ "), font=dict(size=13))
    fig = go.Figure(indicador)
    fig.update_layout(margin=dict(l=22, r=22, t=18, b=0))
    return fig


# ---- 4.7 Mapa de calor: departamento x año, en colores de cuartil -----------
def grafico_mapa_calor(base, dpto_sel, anio, col, ultimos_anios=5):
    cuartiles = base.pivot_table(index=COL_DPTO, columns=COL_ANIO, values=col_cuartil(col))
    valores = base.pivot_table(index=COL_DPTO, columns=COL_ANIO, values=col)
    orden = valores.iloc[:, -ultimos_anios:].mean(axis=1).sort_values(ascending=False).index
    cuartiles, valores = cuartiles.loc[orden], valores.loc[orden]
    anios = list(cuartiles.columns)
    escala = []
    for i, q in enumerate(sorted(COLORES_CUARTIL)):  # escala de 4 bloques de color
        escala += [[i / 4, COLORES_CUARTIL[q]], [(i + 1) / 4, COLORES_CUARTIL[q]]]
    custom = [[[d, None if es_vacio(v) else float(v)] for v in fila]
              for d, fila in zip(orden, valores.to_numpy())]
    fig = go.Figure(go.Heatmap(
        z=cuartiles.to_numpy(), x=anios, y=[nombre_visible(d) for d in orden], customdata=custom,
        zmin=0.5, zmax=4.5, colorscale=escala, xgap=2, ygap=2, hoverongaps=False,
        colorbar=dict(title=dict(text="Cuartil", font=dict(size=11)), thickness=10, len=0.5,
                      tickvals=[1, 2, 3, 4], ticktext=["1", "2", "3", "4"], outlinewidth=0,
                      tickfont=dict(size=10), y=1, yanchor="top"),
        hovertemplate=("<b>%{y}</b>, %{x}<br>" + nombre_indicador(col)
                       + ": %{customdata[1]:.2f}<br>Cuartil: %{z}<extra></extra>")))
    marco = dict(type="rect", line=dict(color=C["tinta"], width=2), fillcolor="rgba(0,0,0,0)")
    if dpto_sel in list(orden):
        i = list(orden).index(dpto_sel)
        fig.add_shape(**marco, x0=anios[0] - 0.5, x1=anios[-1] + 0.5, y0=i - 0.5, y1=i + 0.5)
    if anio in anios:
        fig.add_shape(**marco, x0=anio - 0.5, x1=anio + 0.5, y0=-0.5, y1=len(orden) - 0.5)
    fig.update_layout(margin=dict(l=4, r=4, t=4, b=4),
                      xaxis=dict(dtick=1, tickangle=-90, tickfont=dict(size=10), showline=False, side="top"),
                      yaxis=dict(autorange="reversed", tickfont=dict(size=11, color=C["tinta_2"]),
                                 showgrid=False))
    return fig


# ---- 4.8 Matriz de riesgo animada: frecuencia vs severidad ------------------
def grafico_matriz_riesgo(base, dpto_sel, anio, col):
    """Burbuja = departamento; tamaño = damnificados; color = cuartil del indicador.
    Devuelve un dict (Dash lo acepta) porque así los 23 cuadros se arman rápido."""
    anios = lista_anios(base)
    max_x = base[COL_FRECUENCIA].max() * 1.06
    max_y = base[COL_SEVERIDAD].max() * 1.12
    diametro_max, tam_min = 58, 6
    tam_ref = base[COL_DAMNIFICADOS].max() / 2 / diametro_max ** 2
    cq = col_cuartil(col)

    def trazas_del_anio(a):
        d = base[base[COL_ANIO] == a].sort_values(COL_DPTO)
        etiquetados = set(d.nlargest(3, col)[COL_DPTO]) if dpto_sel == NACIONAL else {dpto_sel}
        principal = dict(
            type="scatter", x=d[COL_FRECUENCIA].tolist(), y=d[COL_SEVERIDAD].tolist(),
            ids=d[COL_DPTO].tolist(), mode="markers+text", textposition="top center",
            text=[n if k in etiquetados else "" for k, n in zip(d[COL_DPTO], d["NOMBRE"])],
            textfont=dict(size=10, color=C["tinta"]), showlegend=False,
            marker=dict(size=d[COL_DAMNIFICADOS].tolist(), sizemode="area", sizeref=tam_ref,
                        sizemin=tam_min, opacity=0.9, line=dict(color=BORDE_MAPA, width=0.8),
                        color=colores_cuartil(d[cq])),
            customdata=pd.DataFrame({"k": d[COL_DPTO], "n": d["NOMBRE"], "v": d[col],
                                     "dm": d[COL_DAMNIFICADOS], "q": d[cq].astype(int)})
            .to_numpy(dtype=object).tolist(),
            hovertemplate=("<b>%{customdata[1]}</b> (" + str(a) + ")<br>Frecuencia: %{x:.2f}"
                           "<br>Severidad: %{y:.2f}<br>" + nombre_indicador(col) + ": %{customdata[2]:.2f}"
                           "<br>Damnificados: %{customdata[3]:,.0f}<br>Cuartil: %{customdata[4]}<extra></extra>"))
        sel = d[d[COL_DPTO] == dpto_sel]
        anillo = dict(
            type="scatter", x=sel[COL_FRECUENCIA].tolist(), y=sel[COL_SEVERIDAD].tolist(),
            mode="markers", hoverinfo="skip", showlegend=False,
            marker=dict(size=[max(np.sqrt(v / 2 / tam_ref), tam_min) + 9 for v in sel[COL_DAMNIFICADOS]],
                        symbol="circle-open", color=C["tinta"], line=dict(width=2.5)))
        return [principal, anillo]

    leyenda = [go.Scatter(x=[None], y=[None], mode="markers", name=NOMBRES_CUARTIL[q],
                          marker=dict(size=11, color=COLORES_CUARTIL[q], line=dict(color=BORDE_MAPA, width=0.8)))
               for q in sorted(COLORES_CUARTIL, reverse=True)]
    fig = go.Figure(data=trazas_del_anio(anio) + leyenda)

    linea = dict(color=C["tinta_3"], width=1, dash="dash")  # cuadrantes = medianas
    fig.add_vline(x=base[COL_FRECUENCIA].median(), line=linea, layer="below")
    fig.add_hline(y=base[COL_SEVERIDAD].median(), line=linea, layer="below")
    for x, y, texto, xa, ya in [(max_x, max_y, "Frecuente y severo", "right", "top"),
                                (0, max_y, "Poco frecuente pero severo", "left", "top"),
                                (max_x, 0, "Frecuente pero leve", "right", "bottom"),
                                (0, 0, "Poco frecuente y leve", "left", "bottom")]:
        fig.add_annotation(x=x, y=y, text=texto, showarrow=False, xanchor=xa, yanchor=ya,
                           font=dict(size=11, color=C["tinta_3"]))

    animar = dict(frame=dict(duration=900, redraw=False), fromcurrent=True,
                  transition=dict(duration=600, easing="cubic-in-out"))
    pausar = dict(frame=dict(duration=0, redraw=False), mode="immediate", transition=dict(duration=0))
    fig.update_layout(
        margin=dict(l=4, r=4, t=36, b=86),
        legend=dict(y=1.04, title=dict(text="Cuartil ", font=dict(size=11))),
        xaxis=dict(title="Índice de frecuencia", range=[0, max_x], showgrid=True, gridcolor=C["rejilla"]),
        yaxis=dict(title="Índice de severidad", range=[0, max_y]),
        updatemenus=[dict(type="buttons", direction="left", x=0, xanchor="left", y=-0.14, yanchor="top",
                          showactive=False, pad=dict(t=0, r=8), bgcolor=C["superficie"],
                          bordercolor=C["borde"], font=dict(size=11, color=C["tinta"]),
                          buttons=[dict(label="▶ Reproducir", method="animate", args=[None, animar]),
                                   dict(label="❚❚ Pausa", method="animate", args=[[None], pausar])])],
        sliders=[dict(active=anios.index(anio) if anio in anios else 0,
                      x=0.28, len=0.72, y=-0.1, yanchor="top", pad=dict(t=0, b=0),
                      currentvalue=dict(prefix="Año ", font=dict(size=12, color=C["tinta"]), xanchor="right"),
                      tickcolor=C["borde"], font=dict(size=9, color=C["tinta_3"]),
                      bgcolor=C["rejilla"], activebgcolor=C["pacifico_medio"], bordercolor=C["borde"],
                      steps=[dict(label=str(a) if (a - anios[0]) % 4 == 0 or a == anios[-1] else "",
                                  method="animate", args=[[str(a)], pausar]) for a in anios])])
    figura = fig.to_plotly_json()
    figura["frames"] = [dict(name=str(a), traces=[0, 1], data=trazas_del_anio(a)) for a in anios]
    return figura


# ---- 4.9 ¿Quién sube y quién baja? (barras divergentes) --------------------
def grafico_variacion(tabla, dpto_sel, anio, col):
    if tabla.empty:
        return figura_vacia(f"No hay un año anterior a {anio} para comparar.")
    nombres = [nombre_visible(d) for d in tabla.index]
    extremos = set(tabla.index[:3]) | set(tabla.index[-3:]) | {dpto_sel}
    fig = go.Figure(go.Bar(
        x=tabla["cambio"], y=tabla.index, orientation="h", showlegend=False,
        marker=dict(color=[C["empeora"] if v > 0 else C["mejora"] for v in tabla["cambio"]],
                    opacity=opacidades(tabla.index, dpto_sel) if dpto_sel != NACIONAL else 1.0),
        text=[con_signo(v, 2) if d in extremos else "" for d, v in tabla["cambio"].items()],
        textposition="outside", textfont=dict(size=10, color=C["tinta"]), cliponaxis=False,
        customdata=pd.DataFrame({"n": nombres, "p": tabla["previo"].to_numpy(),
                                 "a": tabla["actual"].to_numpy()}).to_numpy(dtype=object),
        hovertemplate=("<b>%{customdata[0]}</b><br>" + f"{anio - 1}: " + "%{customdata[1]:.2f}"
                       + f"  →  {anio}: " + "%{customdata[2]:.2f}<br>Cambio: %{x:+.2f}<extra></extra>")))
    for nombre, color in [("Sube", C["empeora"]), ("Baja", C["mejora"])]:
        fig.add_bar(x=[None], y=[None], name=nombre, marker_color=color)
    etiquetas = [f"<b>{n}</b>" if d == dpto_sel else n for d, n in zip(tabla.index, nombres)]
    rango = tabla["cambio"].abs().max() * 1.35
    fig.update_layout(
        margin=dict(l=4, r=4, t=36, b=4), legend=dict(y=1.04), bargap=0.25,
        xaxis=dict(title=f"Cambio de {nombre_indicador(col)}, {anio - 1} → {anio}", zeroline=True,
                   zerolinecolor=C["tinta_3"], range=[-rango, rango], showgrid=True,
                   gridcolor=C["rejilla"], showline=False),
        yaxis=dict(tickvals=list(tabla.index), ticktext=etiquetas, tickfont=dict(size=10)))
    return fig


# ---- 4.10 ¿Cambió el tipo de daño? (área 100 %) -----------------------------
def grafico_composicion(serie, anio):
    partes = list(APORTES[COL_SEVERIDAD].values())
    total = serie[partes].sum(axis=1).replace(0, np.nan)
    fig = go.Figure()
    for nombre, columna in APORTES[COL_SEVERIDAD].items():
        fig.add_scatter(x=serie.index, y=serie[columna] / total * 100, name=nombre, stackgroup="uno",
                        mode="lines", line=dict(width=1, color=C["superficie"]),
                        fillcolor=hex_a_rgba(COLORES_DIMENSION[nombre], 0.88),
                        hovertemplate=f"{nombre}: %{{y:.0f}} %<extra></extra>")
    marcar_anio(fig, anio)
    fig.update_layout(hovermode="x unified", margin=dict(l=4, r=4, t=36, b=4),
                      legend=dict(y=1.04, traceorder="normal"),
                      xaxis=dict(dtick=2, tick0=int(serie.index.min()), tickfont=dict(size=10),
                                 showline=False, range=[serie.index.min() - 0.4, serie.index.max() + 0.4]),
                      yaxis=dict(range=[0, 100], ticksuffix=" %", title="Parte del daño"))
    return fig


# ---- 4.11 Velocímetro de un índice climático (como el de tus reportes) -------
def grafico_velocimetro_clima(cfg, valor):
    segs = segmentos_clima(cfg)
    n, r_in = len(segs), 0.36
    paso = 180 / n
    fig = go.Figure()

    def arco(a1, a2, radio, puntos=24):
        ang = np.radians(np.linspace(a1, a2, puntos))
        return list(radio * np.cos(ang)), list(radio * np.sin(ang))

    for k, seg in enumerate(segs):  # sectores de igual tamaño, cálido a la izquierda
        a1, a2 = 180 - k * paso, 180 - (k + 1) * paso
        xo, yo = arco(a1, a2, 1)
        xi, yi = arco(a2, a1, r_in)
        fig.add_scatter(x=xo + xi + [xo[0]], y=yo + yi + [yo[0]], fill="toself", mode="lines",
                        fillcolor=seg["color"], line=dict(color="#4d4d4d", width=0.8), hoverinfo="skip")
        medio = np.radians((a1 + a2) / 2)
        etiqueta = "NEUTRO" if seg["grupo"] == "Neutro" else seg["magnitud"].upper()
        etiqueta = {"MUY FUERTE": "MUY<br>FUERTE", "EXTRAORDINARIA": "EXTRA-<br>ORDINARIA"}.get(etiqueta, etiqueta)
        fig.add_annotation(x=1.12 * np.cos(medio), y=1.12 * np.sin(medio), text=etiqueta, showarrow=False,
                           textangle=90 - (a1 + a2) / 2, font=dict(size=8, color=C["tinta"]))
    xh, yh = arco(180, 0, r_in)  # centro blanco
    fig.add_scatter(x=xh, y=yh, fill="toself", mode="lines", fillcolor="#ffffff",
                    line=dict(color="#4d4d4d", width=0.8), hoverinfo="skip")
    fig.add_shape(type="line", x0=-1, x1=1, y0=0, y1=0, line=dict(color="#4d4d4d", width=1))

    k = clasificar_clima(valor, cfg)
    if k is not None:  # aguja: dentro del sector de su categoría
        seg = segs[k]
        fraccion = (seg["hasta"] - float(valor)) / (seg["hasta"] - seg["desde"])
        fraccion = min(max(fraccion, 0.06), 0.94)
        angulo = np.radians(180 - (k + fraccion) * paso)
        largo = r_in * 0.86
        fig.add_annotation(x=largo * np.cos(angulo), y=largo * np.sin(angulo), ax=0, ay=0,
                           xref="x", yref="y", axref="x", ayref="y", text="", showarrow=True,
                           arrowhead=2, arrowsize=1, arrowwidth=2.6, arrowcolor="#111111")
        fig.add_scatter(x=[0], y=[0], mode="markers", hoverinfo="skip",
                        marker=dict(size=9, color="#111111"))
    fig.add_annotation(x=-1, y=-0.02, text="<b>CÁLIDA</b>", showarrow=False, xanchor="left", yanchor="top",
                       font=dict(size=9.5, color=C["tinta"]))
    fig.add_annotation(x=1, y=-0.02, text="<b>FRÍA</b>", showarrow=False, xanchor="right", yanchor="top",
                       font=dict(size=9.5, color=C["tinta"]))
    fig.update_layout(showlegend=False, margin=dict(l=2, r=2, t=2, b=2),
                      xaxis=dict(visible=False, range=[-1.24, 1.24], fixedrange=True),
                      yaxis=dict(visible=False, range=[-0.17, 1.27], fixedrange=True,
                                 scaleanchor="x", scaleratio=1))
    return fig


# ---- Ayudas para los gráficos climáticos --------------------------------------
def limites_neutro(cfg):
    return cfg["frias"][0][1], cfg["calidas"][-1][1]


def grupo_clima(valor, cfg):
    k = clasificar_clima(valor, cfg)
    return None if k is None else segmentos_clima(cfg)[k]["grupo"]


def ticks_semestrales(etiquetas, meses):
    """Etiquetas del eje solo en enero y julio (para ejes por categoría)."""
    return [e for e, m in zip(etiquetas, meses) if m in (1, 7)]


def picos(serie, ventana, umbral, maximos=True, cuantos=3):
    """Meses que son el pico de su episodio (máximo o mínimo en ±ventana meses)."""
    s = serie if maximos else -serie
    techo = s.rolling(2 * ventana + 1, center=True, min_periods=1).max()
    candidatos = s[(s == techo) & (s > (umbral if maximos else -umbral))].sort_values(ascending=False)
    elegidos = []
    for i in candidatos.index:  # un solo pico por episodio (evita empates vecinos)
        if all(abs(i - j) > ventana for j in elegidos):
            elegidos.append(i)
    return elegidos[:cuantos]


# ---- 4.12 Evolución: relleno rojo arriba de cero y azul abajo -----------------
def grafico_evolucion_clima(clima, clave, idx):
    cfg = INDICES_CLIMA[clave]
    dec = cfg["decimales"]
    lim_frio, lim_calido = limites_neutro(cfg)
    datos = clima[["fecha", "etiqueta", clave, clave + "_prelim"]].dropna(subset=[clave])
    t = datos["fecha"].to_numpy(dtype="datetime64[ns]").astype("int64")
    v = datos[clave].to_numpy(dtype=float)
    xs, ys = [], []
    for i in range(len(v)):  # se agrega el punto donde la curva cruza el cero
        if i and v[i - 1] * v[i] < 0:
            f = v[i - 1] / (v[i - 1] - v[i])
            xs.append(t[i - 1] + f * (t[i] - t[i - 1]))
            ys.append(0.0)
        xs.append(t[i])
        ys.append(v[i])
    xs = pd.to_datetime(np.array(xs).astype("int64"))
    ys = np.array(ys)
    fig = go.Figure()
    fig.add_scatter(x=xs, y=np.maximum(ys, 0), mode="lines", fill="tozeroy", line=dict(width=0),
                    fillcolor=hex_a_rgba(COLOR_CALIDO, 0.62), hoverinfo="skip", showlegend=False)
    fig.add_scatter(x=xs, y=np.minimum(ys, 0), mode="lines", fill="tozeroy", line=dict(width=0),
                    fillcolor=hex_a_rgba(COLOR_FRIO, 0.62), hoverinfo="skip", showlegend=False)
    segs = segmentos_clima(cfg)
    custom = [[e, valor_clima(x, dec), segs[clasificar_clima(x, cfg)]["texto"], " (preliminar)" if p else ""]
              for e, x, p in zip(datos["etiqueta"], v, datos[clave + "_prelim"])]
    fig.add_scatter(x=datos["fecha"], y=v, mode="lines", line=dict(color=C["tinta"], width=1.4),
                    customdata=custom, showlegend=False,
                    hovertemplate="<b>%{customdata[0]}</b>: %{customdata[1]}%{customdata[3]}"
                                  "<br>%{customdata[2]}<extra></extra>")
    for y, texto, color, ancla in [(lim_calido, f"Umbral cálido {umbral_texto(lim_calido, cfg.get('decimales_criterio', 1))}",
                                    COLOR_CALIDO, "bottom"),
                                   (lim_frio, f"Umbral frío {umbral_texto(lim_frio, cfg.get('decimales_criterio', 1))}",
                                    COLOR_FRIO, "top")]:
        fig.add_hline(y=y, line=dict(color=color, width=1.2, dash="dash"), layer="below")
        fig.add_annotation(x=0, xref="paper", y=y, text=texto, showarrow=False, xanchor="left", yanchor=ancla,
                           font=dict(size=10.5, color=color), bgcolor=hex_a_rgba(C["superficie"], 0.8))
    fig.add_hline(y=0, line=dict(color=C["tinta_3"], width=1), layer="below")

    serie = datos[clave]
    for i in picos(serie, 9, lim_calido, True, 3) + picos(serie, 9, lim_frio, False, 2):
        if i == idx:
            continue
        arriba = serie[i] > 0
        fig.add_annotation(x=datos.loc[i, "fecha"], y=serie[i], ax=0, ay=-26 if arriba else 26,
                           text=f"{valor_clima(serie[i], dec)} · {datos.loc[i, 'etiqueta']}", showarrow=True,
                           arrowhead=0, arrowwidth=1, arrowcolor=C["tinta_3"],
                           font=dict(size=10, color=C["tinta_2"]), bgcolor=hex_a_rgba(C["superficie"], 0.8))
    fila = clima.loc[idx]
    if not es_vacio(fila[clave]):
        fig.add_scatter(x=[fila["fecha"]], y=[fila[clave]], mode="markers", hoverinfo="skip", showlegend=False,
                        marker=dict(size=14, color=C["superficie"], line=dict(color=C["tinta"], width=3)))
        fig.add_annotation(x=fila["fecha"], y=fila[clave], ax=0, ay=-34 if fila[clave] >= 0 else 34,
                           text=f"<b>{valor_clima(fila[clave], dec)} · {fila['etiqueta']}</b>", showarrow=True,
                           arrowhead=0, arrowwidth=1.4, arrowcolor=C["tinta"],
                           font=dict(size=12, color=C["superficie"]), bgcolor=C["tinta"], borderpad=4)
    fig.add_vline(x=fila["fecha"], line=dict(color=C["pacifico_medio"], width=1, dash="dot"), layer="below")
    bajo, alto = min(np.nanmin(v), lim_frio), max(np.nanmax(v), lim_calido)
    pad = (alto - bajo) * 0.1
    fig.update_layout(
        hovermode="closest", margin=dict(l=4, r=4, t=36, b=4),
        yaxis=dict(title=clave, range=[bajo - pad, alto + pad * 1.8], zeroline=False),
        xaxis=dict(showgrid=False, tickformat="%Y", dtick="M24", rangeselector=dict(
            x=0, y=1.02, yanchor="bottom", bgcolor=C["niebla"], activecolor=C["azul_suave"],
            font=dict(size=11, color=C["tinta"]),
            buttons=[dict(count=2, label="2 años", step="year", stepmode="backward"),
                     dict(count=5, label="5 años", step="year", stepmode="backward"),
                     dict(step="all", label="Todo")])))
    return fig


# ---- 4.13 Variación a 12 meses (barras divergentes) --------------------------
def variacion_12(clima, clave):
    return clima[clave] - clima[clave].shift(12)


def grafico_variacion12_clima(clima, clave, idx):
    if idx < 12:
        return figura_vacia("Se necesitan 12 meses previos para calcular la variación.")
    dec = INDICES_CLIMA[clave]["decimales"]
    cambio = variacion_12(clima, clave)
    tramo = clima.loc[max(12, idx - MESES_VENTANA + 1):idx]
    y = cambio.loc[tramo.index]
    custom = [[e, valor_clima(clima.loc[i, clave], dec), clima.loc[i - 12, "etiqueta"],
               valor_clima(clima.loc[i - 12, clave], dec), con_signo(round(c, dec) + 0.0, dec) if not es_vacio(c) else "—"]
              for i, e, c in zip(tramo.index, tramo["etiqueta"], y)]
    colores = [C["contexto"] if es_vacio(c) else (COLOR_CALIDO if c > 0 else COLOR_FRIO) for c in y]
    fig = go.Figure(go.Bar(
        x=tramo["etiqueta"], y=y, marker=dict(color=colores, opacity=[1 if i == idx else 0.5 for i in tramo.index]),
        text=[f"<b>{con_signo(round(c, dec) + 0.0, dec)}</b>" if i == idx and not es_vacio(c) else ""
              for i, c in zip(tramo.index, y)],
        textposition="outside", cliponaxis=False, constraintext="none", textfont=dict(size=12, color=C["tinta"]),
        customdata=custom,
        hovertemplate="<b>%{customdata[0]}</b>: %{customdata[1]}<br>%{customdata[2]}: %{customdata[3]}"
                      "<br>Cambio: <b>%{customdata[4]}</b><extra></extra>"))
    fig.add_hline(y=0, line=dict(color=C["tinta_3"], width=1))
    tope = np.nanmax(np.abs(y.to_numpy(dtype=float))) if y.notna().any() else 1
    fig.update_layout(margin=dict(l=4, r=26, t=22, b=4), bargap=0.25, showlegend=False,
                      xaxis=dict(type="category", tickvals=ticks_semestrales(tramo["etiqueta"], tramo["mes"]),
                                 showline=False),
                      yaxis=dict(title="Cambio vs. mismo mes del año anterior", range=[-tope * 1.25, tope * 1.25],
                                 zeroline=False))
    return fig


# ---- 4.14 Hace un año vs. hoy (los cuatro índices) ---------------------------
def grafico_anio_vs_hoy(clima, idx):
    if idx < 12:
        return figura_vacia("Se necesitan 12 meses previos para comparar.")
    hoy, antes = clima.loc[idx], clima.loc[idx - 12]
    claves = list(INDICES_CLIMA)[::-1]  # el primero queda arriba
    fig = go.Figure()
    for k in claves:
        cfg = INDICES_CLIMA[k]
        x0, x1 = antes[k], hoy[k]
        if es_vacio(x0) or es_vacio(x1):
            continue
        sube = x1 > x0
        fig.add_annotation(x=x1, y=k, ax=x0, ay=k, xref="x", yref="y", axref="x", ayref="y", text="",
                           showarrow=True, arrowhead=2, arrowsize=1.1, arrowwidth=3, standoff=11,
                           arrowcolor=hex_a_rgba(COLOR_CALIDO if sube else COLOR_FRIO, 0.75))
        seg = segmentos_clima(cfg)[clasificar_clima(x1, cfg)]
        fig.add_scatter(x=[x0], y=[k], mode="markers+text", text=[valor_clima(x0, cfg["decimales"])],
                        textposition="bottom center", textfont=dict(size=11, color=C["tinta_3"]),
                        marker=dict(size=14, color=C["superficie"], line=dict(color=C["tinta_3"], width=2)),
                        showlegend=False, hovertemplate=f"<b>{k}</b> {antes['etiqueta']}: %{{x:.2f}}<extra></extra>")
        fig.add_scatter(x=[x1], y=[k], mode="markers+text", text=[f"<b>{valor_clima(x1, cfg['decimales'])}</b>"],
                        textposition="top center", textfont=dict(size=13, color=C["tinta"]),
                        marker=dict(size=20, color=seg["color"], line=dict(color=C["tinta"], width=1.5)),
                        showlegend=False,
                        hovertemplate=f"<b>{k}</b> {hoy['etiqueta']}: %{{x:.2f}}<br>{seg['texto']}<extra></extra>")
    fig.add_vline(x=0, line=dict(color=C["tinta_3"], width=1, dash="dot"), layer="below")
    fig.add_scatter(x=[None], y=[None], mode="markers", name=antes["etiqueta"],
                    marker=dict(size=12, color=C["superficie"], line=dict(color=C["tinta_3"], width=2)))
    fig.add_scatter(x=[None], y=[None], mode="markers", name=f"{hoy['etiqueta']} (color de su categoría)",
                    marker=dict(size=12, color=COLOR_CALIDO, line=dict(color=C["tinta"], width=1.5)))
    fig.update_layout(margin=dict(l=4, r=12, t=36, b=4), legend=dict(y=1.04),
                      xaxis=dict(title="Valor del índice", showgrid=True, gridcolor=C["rejilla"], zeroline=False),
                      yaxis=dict(type="category", categoryorder="array", categoryarray=claves, showgrid=False,
                                 tickfont=dict(size=13, color=C["tinta"])))
    return fig


# ---- 4.15 Cada año, mes a mes (estacionalidad) -------------------------------
def grafico_estacional_clima(clima, clave, idx):
    cfg = INDICES_CLIMA[clave]
    dec = cfg["decimales"]
    lim_frio, lim_calido = limites_neutro(cfg)
    anio_sel = int(clima.loc[idx, "anio"])
    tabla = clima.pivot_table(index="anio", columns="mes", values=clave, aggfunc="first").reindex(columns=range(1, 13))
    anio_max = int(tabla.max(axis=1).idxmax())
    fig = go.Figure()
    fig.add_hrect(y0=lim_frio, y1=lim_calido, fillcolor=C["niebla"], opacity=1, line_width=0, layer="below")
    fig.add_annotation(x=12.4, y=(lim_frio + lim_calido) / 2, text="Neutro", showarrow=False, xanchor="left",
                       font=dict(size=10, color=C["tinta_3"]))

    def linea(anio, color, ancho, marcadores=False, opacidad=1.0):
        fila = tabla.loc[anio]
        custom = [[anio, MESES[m - 1], valor_clima(x, dec)] for m, x in fila.items()]
        fig.add_scatter(x=list(range(1, 13)), y=fila.values, mode="lines+markers" if marcadores else "lines",
                        line=dict(color=color, width=ancho, shape="spline", smoothing=0.5), opacity=opacidad,
                        marker=dict(size=7, color=color), customdata=custom, showlegend=False,
                        hovertemplate="<b>%{customdata[1]} %{customdata[0]}</b>: %{customdata[2]}<extra></extra>")
        validos = fila.dropna()
        if marcadores or anio == anio_max:
            fig.add_annotation(x=validos.index[-1], y=validos.iloc[-1], text=f"<b>{anio}</b>", showarrow=False,
                               xanchor="left", xshift=8, font=dict(size=12, color=color))

    for anio in tabla.index:
        if anio not in (anio_sel, anio_max):
            linea(anio, "#b9c9d9", 1.2, opacidad=0.8)
    if anio_max != anio_sel:
        linea(anio_max, COLOR_CALIDO, 2.4)
    linea(anio_sel, C["tinta"], 3.4, marcadores=True)
    valor = clima.loc[idx, clave]
    if not es_vacio(valor):
        fig.add_scatter(x=[clima.loc[idx, "mes"]], y=[valor], mode="markers", hoverinfo="skip", showlegend=False,
                        marker=dict(size=16, color="rgba(0,0,0,0)", line=dict(color=C["pacifico_medio"], width=3)))
    fig.update_layout(hovermode="closest", margin=dict(l=4, r=40, t=8, b=4),
                      xaxis=dict(tickvals=list(range(1, 13)), ticktext=[m[:3] for m in MESES], range=[0.7, 13.2],
                                 showgrid=False),
                      yaxis=dict(title=clave, zeroline=True, zerolinecolor=C["contexto"]))
    return fig


# ---- 4.16 Los 10 meses más cálidos -------------------------------------------
def grafico_top_clima(clima, clave, idx):
    cfg = INDICES_CLIMA[clave]
    segs, dec = segmentos_clima(cfg), cfg["decimales"]
    top = clima[["etiqueta", clave]].dropna().nlargest(10, clave).iloc[::-1]
    sel = clima.loc[idx, "etiqueta"]
    colores = [segs[clasificar_clima(x, cfg)]["color"] for x in top[clave]]
    fig = go.Figure(go.Bar(
        y=top["etiqueta"], x=top[clave], orientation="h",
        marker=dict(color=colores, line=dict(color=[C["tinta"] if e == sel else BORDE_MAPA for e in top["etiqueta"]],
                                             width=[3 if e == sel else 0.6 for e in top["etiqueta"]])),
        text=[valor_clima(x, dec) for x in top[clave]], textposition="outside", cliponaxis=False,
        textfont=dict(size=11, color=C["tinta"]),
        customdata=[segs[clasificar_clima(x, cfg)]["texto"] for x in top[clave]],
        hovertemplate="<b>%{y}</b>: %{x:.2f}<br>%{customdata}<extra></extra>"))
    fig.update_layout(margin=dict(l=4, r=4, t=8, b=4), bargap=0.2, showlegend=False,
                      xaxis=dict(visible=False, range=[0, top[clave].max() * 1.15]),
                      yaxis=dict(type="category", tickvals=list(top["etiqueta"]),
                                 ticktext=[f"<b>{e}</b>" if e == sel else e for e in top["etiqueta"]],
                                 tickfont=dict(size=11, color=C["tinta"]), showline=True, linecolor=C["borde"]))
    return fig


# ---- 4.17 ¿Coinciden los cuatro índices? (barras apiladas) -------------------
def grafico_consenso_clima(clima, idx):
    tramo = clima.loc[max(0, idx - MESES_VENTANA + 1):idx]
    conteo = {"Condición cálida": [], "Neutro": [], "Condición fría": []}
    for _, fila in tramo.iterrows():
        grupos = [grupo_clima(fila[k], cfg) for k, cfg in INDICES_CLIMA.items()]
        for g in conteo:
            conteo[g].append(grupos.count(g))
    fig = go.Figure()
    for g, color, nombre in [("Condición cálida", COLOR_CALIDO, "Cálida"), ("Neutro", COLOR_NEUTRO_GRIS, "Neutro"),
                             ("Condición fría", COLOR_FRIO, "Fría")]:
        fig.add_bar(x=tramo["etiqueta"], y=conteo[g], name=nombre, marker=dict(color=color, line=dict(width=0)),
                    hovertemplate=f"%{{x}}: %{{y}} índice(s) en condición {nombre.lower()}<extra></extra>")
    pos = list(tramo.index).index(idx)
    fig.add_shape(type="rect", x0=pos - 0.5, x1=pos + 0.5, y0=0, y1=len(INDICES_CLIMA), xref="x", yref="y",
                  line=dict(color=C["tinta"], width=2.5), fillcolor="rgba(0,0,0,0)")
    fig.update_layout(barmode="stack", bargap=0.12, margin=dict(l=4, r=4, t=36, b=4),
                      legend=dict(y=1.04, traceorder="normal"),
                      xaxis=dict(type="category", tickvals=ticks_semestrales(tramo["etiqueta"], tramo["mes"]),
                                 showline=False),
                      yaxis=dict(title="Número de índices", range=[0, len(INDICES_CLIMA)], dtick=1))
    return fig


# ---- 4.18 Calendario: año x mes en colores del criterio ----------------------
def grafico_calendario_clima(clima, clave, idx):
    cfg = INDICES_CLIMA[clave]
    segs, dec = segmentos_clima(cfg), cfg["decimales"]
    anios = sorted(clima["anio"].unique(), reverse=True)
    tabla = clima.pivot_table(index="anio", columns="mes", values=clave, aggfunc="first")
    tabla = tabla.reindex(index=anios, columns=range(1, 13))
    z = [[clasificar_clima(v, cfg) for v in fila] for fila in tabla.to_numpy()]
    texto = [[valor_clima(v, 1) if not es_vacio(v) else "" for v in fila] for fila in tabla.to_numpy()]
    custom = [[[f"{MESES[m - 1]} {a}", valor_clima(v, dec),
                segs[c]["texto"] if c is not None else "Sin dato"]
               for m, v, c in zip(range(1, 13), fila, zfila)]
              for a, fila, zfila in zip(anios, tabla.to_numpy(), z)]
    n = len(segs)
    escala = []
    for i, seg in enumerate(segs):
        escala += [[i / n, seg["color"]], [(i + 1) / n, seg["color"]]]
    fig = go.Figure(go.Heatmap(
        z=z, x=list(range(1, 13)), y=anios, text=texto, texttemplate="%{text}",
        textfont=dict(size=10), customdata=custom, colorscale=escala, zmin=-0.5, zmax=n - 0.5,
        showscale=False, xgap=2, ygap=2, hoverongaps=False,
        hovertemplate="<b>%{customdata[0]}</b>: %{customdata[1]}<br>%{customdata[2]}<extra></extra>"))
    fila = clima.loc[idx]
    fig.add_shape(type="rect", x0=fila["mes"] - 0.5, x1=fila["mes"] + 0.5, y0=fila["anio"] - 0.5,
                  y1=fila["anio"] + 0.5, line=dict(color=C["tinta"], width=2.5), fillcolor="rgba(0,0,0,0)")
    fig.update_layout(margin=dict(l=4, r=4, t=4, b=4), plot_bgcolor=C["niebla"],
                      xaxis=dict(tickvals=list(range(1, 13)), ticktext=[m[:3] for m in MESES], side="top",
                                 showline=False, tickfont=dict(size=11)),
                      yaxis=dict(tickvals=anios, autorange="reversed", showgrid=False, tickfont=dict(size=11)))
    return fig


# ---- 4.19 Ranking de países (barras en colores de cuartil) --------------------
def grafico_ranking_latam(tabla, pais_sel, ind):
    datos = tabla.assign(q=cuartiles_latam(tabla[ind])).sort_values(ind)
    sel = datos.index == pais_sel
    fig = go.Figure(go.Bar(
        y=datos.index, x=datos[ind], orientation="h",
        marker=dict(color=[COLORES_CUARTIL_LATAM[q] for q in datos["q"]],
                    line=dict(color=[C["tinta"] if s else "rgba(0,0,0,0)" for s in sel],
                              width=[2.5 if s else 0 for s in sel])),
        text=[valor_latam(v, ind) for v in datos[ind]], textposition="outside", cliponaxis=False,
        textfont=dict(size=10.5, color=C["tinta"]),
        hovertemplate="<b>%{y}</b>: %{text}<extra></extra>"))
    fig.update_layout(margin=dict(l=4, r=8, t=4, b=4), bargap=0.22, showlegend=False,
                      xaxis=dict(visible=False, range=[0, datos[ind].max() * 1.16]),
                      yaxis=dict(tickvals=list(datos.index),
                                 ticktext=[f"<b>{p}</b>" if p == pais_sel else p for p in datos.index],
                                 tickfont=dict(size=11, color=C["tinta"]), showline=True, linecolor=C["borde"]))
    return fig


# ---- 4.20 Mapa de América Latina (cuartiles) ---------------------------------
def grafico_mapa_latam(tabla, pais_sel, ind):
    datos = tabla.assign(q=cuartiles_latam(tabla[ind]))
    minimos = datos.groupby("q")[ind].min()
    fig = go.Figure()
    for q in sorted(COLORES_CUARTIL_LATAM, reverse=True):
        sub = datos[datos["q"] == q]
        hasta = valor_latam(minimos[q + 1], ind) if q < 4 else "más"
        color = COLORES_CUARTIL_LATAM[q]
        fig.add_trace(go.Choropleth(
            locations=sub["iso3"], z=[q] * len(sub), locationmode="ISO-3", showscale=False,
            colorscale=[[0, color], [1, color]], name=f"{valor_latam(minimos[q], ind)} a {hasta}",
            showlegend=True, marker=dict(line=dict(color="#ffffff", width=0.7)),
            customdata=[[p, valor_latam(v, ind)] for p, v in zip(sub.index, sub[ind])],
            hovertemplate="<b>%{customdata[0]}</b>: %{customdata[1]}<extra></extra>"))
    if pais_sel in datos.index:
        fig.add_trace(go.Choropleth(
            locations=[datos.loc[pais_sel, "iso3"]], z=[1], locationmode="ISO-3", showscale=False,
            colorscale=[[0, "rgba(0,0,0,0)"], [1, "rgba(0,0,0,0)"]], showlegend=False, hoverinfo="skip",
            marker=dict(line=dict(color=C["tinta"], width=2.4))))
    fig.update_geos(projection_type="mercator", resolution=50, showframe=False, showcoastlines=False,
                    showcountries=True, countrycolor="#ffffff", showland=True, landcolor="#e4ebf2",
                    showocean=False, showlakes=False, bgcolor="rgba(0,0,0,0)",
                    lonaxis=dict(range=[-118, -33]), lataxis=dict(range=[-56, 33]))
    fig.update_layout(margin=dict(l=0, r=0, t=0, b=0), dragmode=False,
                      legend=dict(orientation="v", x=0.02, y=0.05, yanchor="bottom",
                                  title=dict(text=INDICADORES_LATAM[ind], font=dict(size=11)),
                                  bgcolor=hex_a_rgba(C["superficie"], 0.85)))
    return fig


# ---- 4.21 Dos periodos: 1980-2002 vs 2003-2026 --------------------------------
def grafico_periodos_latam(latam, pais_sel):
    e = latam["eventos"]
    antes = e[e["anio"] <= CORTE_PERIODOS].groupby("pais").size()
    despues = e[e["anio"] > CORTE_PERIODOS].groupby("pais").size()
    tabla = pd.DataFrame({"antes": antes, "despues": despues}).fillna(0).sort_values("despues", ascending=False)
    a0, a1 = int(e["anio"].min()), int(e["anio"].max())
    fig = go.Figure()
    for col, nombre, color in [("antes", f"{a0}–{CORTE_PERIODOS}", COLOR_PERIODO_1),
                               ("despues", f"{CORTE_PERIODOS + 1}–{a1}", COLOR_PERIODO_2)]:
        fig.add_bar(x=tabla.index, y=tabla[col], name=nombre, marker_color=color,
                    text=[numero(v, 0) for v in tabla[col]], textposition="outside", cliponaxis=False,
                    textfont=dict(size=10, color=C["tinta_2"]),
                    hovertemplate="<b>%{x}</b> " + nombre + ": %{y}<extra></extra>")
    fig.update_layout(barmode="group", bargap=0.25, bargroupgap=0.05, margin=dict(l=4, r=4, t=36, b=4),
                      legend=dict(y=1.04),
                      xaxis=dict(tickangle=-45, tickvals=list(tabla.index),
                                 ticktext=[f"<b>{p}</b>" if p == pais_sel else p for p in tabla.index],
                                 tickfont=dict(size=10.5, color=C["tinta"])),
                      yaxis=dict(title="Número de desastres", range=[0, tabla.values.max() * 1.15]))
    return fig


# ---- 4.22 Desastres y costo por año (dispersión con tendencia) ---------------
def grafico_dispersion_latam(latam, a0, a1):
    e = eventos_periodo(latam, a0, min(a1, ULTIMO_ANIO_COSTO))
    anual = e.groupby("anio").agg(n=("id", "size"), costo=("dano_miles_usd", "sum"))
    anual["costo"] = anual["costo"] / 1e6  # miles de millones de US$
    if len(anual) < 3:
        return figura_vacia("Elige un periodo más largo para ver la relación.")
    fig = go.Figure()
    fig.add_scatter(x=anual["n"], y=anual["costo"], mode="markers", showlegend=False,
                    marker=dict(size=11, color=COLOR_DISPERSION, opacity=0.85, line=dict(color="#ffffff", width=1)),
                    customdata=[[a, numero(c, 1)] for a, c in zip(anual.index, anual["costo"])],
                    hovertemplate="<b>%{customdata[0]}</b><br>%{x} desastres<br>"
                                  "US$ %{customdata[1]} mil millones<extra></extra>")
    pendiente, intercepto = np.polyfit(anual["n"], anual["costo"], 1)
    xs = np.array([anual["n"].min(), anual["n"].max()])
    fig.add_scatter(x=xs, y=intercepto + pendiente * xs, mode="lines", hoverinfo="skip", showlegend=False,
                    line=dict(color=COLOR_DISPERSION, width=1.6, dash="dot"))
    for anio in anual["costo"].nlargest(3).index:
        fig.add_annotation(x=anual.loc[anio, "n"], y=anual.loc[anio, "costo"], text=str(anio), showarrow=False,
                           yshift=14, font=dict(size=11, color=C["tinta"]))
    r = np.corrcoef(anual["n"], anual["costo"])[0, 1]
    fig.add_annotation(x=0.01, y=0.98, xref="paper", yref="paper", xanchor="left", yanchor="top", showarrow=False,
                       text=f"Correlación: <b>{numero(r, 2)}</b>", font=dict(size=12, color=C["tinta"]),
                       bgcolor=C["niebla"], borderpad=6)
    fig.update_layout(margin=dict(l=4, r=4, t=8, b=4), hovermode="closest",
                      xaxis=dict(title="Número de desastres en el año", showgrid=True, gridcolor=C["rejilla"]),
                      yaxis=dict(title="Costo anual (miles de millones de US$)", rangemode="tozero"))
    return fig


# ---- 4.23 Distribución por tipo de desastre en los países de una zona --------
def grafico_tipos_zona(latam, a0, a1, zona, pais_sel, medida):
    e = eventos_periodo(latam, a0, a1)
    e = e[e["zona"] == zona]
    if medida == "eventos":
        tabla = e.pivot_table(index="pais", columns="subgrupo", values="id", aggfunc="count", fill_value=0)
    else:
        tabla = e.pivot_table(index="pais", columns="subgrupo", values="dano_miles_usd", aggfunc="sum", fill_value=0)
    paises = sorted(latam["eventos"].loc[latam["eventos"]["zona"] == zona, "pais"].unique())
    tabla = tabla.reindex(index=paises, columns=list(COLORES_SUBGRUPO), fill_value=0)
    total = tabla.sum(axis=1).replace(0, np.nan)
    partes = tabla.div(total, axis=0) * 100
    fig = go.Figure()
    for sub, color in COLORES_SUBGRUPO.items():
        oscuro = sub in ("Biológico", "Climatológico")
        fig.add_bar(x=partes.index, y=partes[sub], name=sub, marker=dict(color=color, line=dict(color="#ffffff", width=0.6)),
                    text=[f"{v:.0f}%" if v >= 12 else "" for v in partes[sub].fillna(0)], textposition="inside",
                    insidetextanchor="middle", textfont=dict(size=10.5, color="#ffffff" if oscuro else C["tinta"]),
                    hovertemplate="<b>%{x}</b> · " + sub + ": %{y:.1f} %<extra></extra>")
    fig.update_layout(barmode="stack", bargap=0.28, margin=dict(l=4, r=4, t=36, b=4),
                      legend=dict(y=1.04, traceorder="normal"),
                      xaxis=dict(tickangle=-35 if len(paises) > 8 else 0, tickvals=list(partes.index),
                                 ticktext=[f"<b>{p}</b>" if p == pais_sel else p for p in partes.index],
                                 tickfont=dict(size=10.5, color=C["tinta"])),
                      yaxis=dict(range=[0, 100], ticksuffix=" %"))
    return fig


# ---- 4.24 Un país: desastres o costo por tipo y subtipo ----------------------
def grafico_tipos_pais(latam, pais, a0, a1, medida):
    e = eventos_periodo(latam, a0, a1)
    e = e[e["pais"] == pais]
    pares = latam["eventos"][["subgrupo", "tipo"]].drop_duplicates()
    if medida == "eventos":
        valores = e.groupby("tipo").size()
    else:
        pbi = latam["pbi"].loc[pais, "pbi_millones_usd"] * 1000
        valores = e.groupby("tipo")["dano_miles_usd"].sum() / pbi * 100
    pares = pares.assign(valor=pares["tipo"].map(valores).fillna(0))
    pares["orden"] = pares["subgrupo"].map({s: i for i, s in enumerate(ORDEN_SUBGRUPOS_PAIS)})
    pares = pares.sort_values(["orden", "valor"], ascending=[True, False])
    textos = [(f"{numero(v, 1)} %" if medida == "costo" else numero(v, 0)) for v in pares["valor"]]
    fig = go.Figure(go.Bar(
        x=[list(pares["subgrupo"]), list(pares["tipo"])], y=pares["valor"], marker_color=COLOR_TIPO_PAIS,
        text=textos, textposition="outside", cliponaxis=False, textfont=dict(size=11, color=C["tinta"]),
        customdata=textos, hovertemplate="<b>%{x}</b>: %{customdata}<extra></extra>"))
    tope = pares["valor"].max() or 1
    fig.update_layout(margin=dict(l=4, r=4, t=16, b=4), bargap=0.35, showlegend=False,
                      xaxis=dict(tickangle=-90, tickfont=dict(size=10, color=C["tinta"])),
                      yaxis=dict(visible=False, range=[0, tope * 1.18]))
    return fig


# ---- 4.25 Evolución anual en la región (por subgrupo) + el país elegido ------
def grafico_evolucion_latam(latam, a0, a1, pais_sel):
    e = eventos_periodo(latam, a0, a1)
    tabla = e.pivot_table(index="anio", columns="subgrupo", values="id", aggfunc="count", fill_value=0)
    tabla = tabla.reindex(index=range(a0, a1 + 1), columns=list(COLORES_SUBGRUPO), fill_value=0)
    fig = go.Figure()
    for sub, color in COLORES_SUBGRUPO.items():
        fig.add_bar(x=tabla.index, y=tabla[sub], name=sub, marker=dict(color=color, line=dict(width=0)),
                    hovertemplate=sub + ": %{y}<extra></extra>")
    pais = e[e["pais"] == pais_sel].groupby("anio").size().reindex(tabla.index, fill_value=0)
    fig.add_scatter(x=pais.index, y=pais.values, name=pais_sel, mode="lines+markers",
                    line=dict(color=C["tinta"], width=2.4), marker=dict(size=6, color=C["tinta"]),
                    hovertemplate=f"<b>{pais_sel}</b>: %{{y}}<extra></extra>")
    fig.update_layout(barmode="stack", bargap=0.15, hovermode="x unified", margin=dict(l=4, r=4, t=36, b=4),
                      legend=dict(y=1.04, traceorder="normal"),
                      xaxis=dict(dtick=2 if a1 - a0 > 12 else 1, tickfont=dict(size=10.5)),
                      yaxis=dict(title="Número de desastres"))
    return fig


# =============================================================================
# 5. DISEÑO  (íconos, piezas de la página y estilos CSS)
# =============================================================================
# Íconos dibujados en SVG (24 x 24). Para agregar uno, pon aquí su nombre y su dibujo.
ICONOS = {
    "alerta": "<path d='M12 3 2 21h20L12 3z'/><path d='M12 10v5'/><path d='M12 18h.01'/>",
    "personas": "<circle cx='9' cy='8' r='3.5'/><path d='M2.5 20c0-3.6 2.9-6 6.5-6s6.5 2.4 6.5 6'/>"
                "<circle cx='17' cy='9' r='2.8'/><path d='M16.5 14.2c3 .3 5 2.6 5 5.8'/>",
    "persona": "<circle cx='12' cy='8' r='4'/><path d='M4 21c0-4.4 3.6-7 8-7s8 2.6 8 7'/>",
    "casa": "<path d='M3 11 12 3l9 8'/><path d='M5 10v10h14V10'/><path d='M10 20v-6h4v6'/>",
    "brote": "<path d='M12 21v-9'/><path d='M12 12C12 7 8 4 3 4c0 5 4 8 9 8z'/>"
             "<path d='M12 14c0-4 3-7 8-7 0 4-3 7-8 7z'/>",
    "medidor": "<path d='M3.5 17a8.5 8.5 0 1 1 17 0'/><path d='M12 17l4.5-5.5'/><circle cx='12' cy='17' r='1.4'/>",
    "mapa": "<path d='M12 21s-7-6.2-7-11.5a7 7 0 0 1 14 0C19 14.8 12 21 12 21z'/><circle cx='12' cy='9.5' r='2.5'/>",
    "ranking": "<path d='M4 5h16'/><path d='M4 10h12'/><path d='M4 15h8'/><path d='M4 20h4'/>",
    "subebaja": "<path d='M7 20V5'/><path d='M3 9l4-4 4 4'/><path d='M17 4v15'/><path d='M13 15l4 4 4-4'/>",
    "barras": "<path d='M5 20v-7'/><path d='M12 20V5'/><path d='M19 20v-10'/><path d='M3 20h18'/>",
    "dona": "<path d='M12 3a9 9 0 1 0 9 9h-9z'/><path d='M15 3.5A9 9 0 0 1 20.5 9H15z'/>",
    "tendencia": "<path d='M3 17l6-6 4 4 8-8'/><path d='M15 7h6v6'/>",
    "burbujas": "<circle cx='8' cy='15' r='4.5'/><circle cx='16.5' cy='7.5' r='3'/><circle cx='18' cy='17.5' r='2'/>",
    "cuadricula": "<rect x='3' y='3' width='7' height='7' rx='1.5'/><rect x='14' y='3' width='7' height='7' rx='1.5'/>"
                  "<rect x='3' y='14' width='7' height='7' rx='1.5'/><rect x='14' y='14' width='7' height='7' rx='1.5'/>",
    "calendario": "<rect x='3' y='5' width='18' height='16' rx='2'/><path d='M3 10h18'/><path d='M8 3v4'/><path d='M16 3v4'/>",
    "indicador": "<path d='M4 20V10'/><path d='M10 20V4'/><path d='M16 20v-7'/><path d='M22 20H2'/>",
    "ola": "<path d='M2 7c2.5-2.7 5-2.7 7.5 0s5 2.7 7.5 0c1.7-1.8 3.3-2.3 5-1.5'/>"
           "<path d='M2 12.5c2.5-2.7 5-2.7 7.5 0s5 2.7 7.5 0c1.7-1.8 3.3-2.3 5-1.5'/>"
           "<path d='M2 18c2.5-2.7 5-2.7 7.5 0s5 2.7 7.5 0c1.7-1.8 3.3-2.3 5-1.5'/>",
    "globo": "<circle cx='12' cy='12' r='9'/><path d='M3 12h18'/>"
             "<path d='M12 3c2.6 2.6 3.8 5.6 3.8 9s-1.2 6.4-3.8 9c-2.6-2.6-3.8-5.6-3.8-9s1.2-6.4 3.8-9z'/>",
    "flecha": "<path d='M5 12h14'/><path d='M13 6l6 6-6 6'/>",
    "volver": "<path d='M19 12H5'/><path d='M11 18l-6-6 6-6'/>",
    "reloj": "<circle cx='12' cy='12' r='9'/><path d='M12 7v5l3 2'/>",
}


def css_iconos():
    """Convierte ICONOS en clases CSS (.icono-alerta, .icono-casa, ...)."""
    reglas = []
    for nombre, dibujo in ICONOS.items():
        svg = ("<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='none' stroke='black' "
               "stroke-width='2' stroke-linecap='round' stroke-linejoin='round'>" + dibujo + "</svg>")
        reglas.append(f'.icono-{nombre}{{--i:url("data:image/svg+xml,{urllib.parse.quote(svg)}")}}')
    return "\n".join(reglas)


def icono(nombre):
    return html.Span(className=f"icono icono-{nombre}")


def barra_lateral(departamentos, anios):
    opciones = [{"label": "Todo el Perú", "value": NACIONAL}] + [
        {"label": nombre_visible(d), "value": d} for d in departamentos]
    marcas = {a: str(a) for a in anios if (a - anios[0]) % 5 == 0 and anios[-1] - a >= 3}
    marcas[anios[-1]] = str(anios[-1])
    return html.Aside(className="lateral", children=[
        dcc.Link([icono("volver"), "Portada"], href="/", className="lateral-volver"),
        html.Div(className="marca", children=[html.Span("IRD", className="marca-sigla"),
                                              html.Span(TITULO, className="marca-nombre")]),
        html.Div(className="seleccion", children=[html.Div(id="lateral-anio", className="seleccion-anio"),
                                                  html.Div(id="lateral-dpto", className="seleccion-dpto")]),
        html.Div(className="controles", children=[
            html.Label([icono("mapa"), "Departamento"], className="control-etiqueta"),
            dcc.Dropdown(id="filtro-dpto", options=opciones, value=NACIONAL, clearable=False),
            html.Div(className="control-fila", children=[
                html.Label([icono("calendario"), "Año"], className="control-etiqueta"),
                html.Button("▶ Reproducir", id="btn-play", n_clicks=0, className="boton-play")]),
            dcc.Slider(id="filtro-anio", min=anios[0], max=anios[-1], step=1, value=anios[-1],
                       marks=marcas, included=False,
                       tooltip={"placement": "bottom", "always_visible": False}),
            html.Label([icono("indicador"), "Indicador"], className="control-etiqueta"),
            dcc.RadioItems(id="filtro-indicador", value=COL_IRD, className="control-radio",
                           options=[{"label": k, "value": v} for k, v in INDICADORES.items()]),
            html.Button(f"Volver a Perú, {anios[-1]}", id="btn-reiniciar", n_clicks=0,
                        className="boton-secundario")]),
        html.P("Consejo: haz clic en el mapa, en las barras o en el mapa de calor para cambiar "
               "el departamento o el año.", className="lateral-nota"),
        dcc.Interval(id="intervalo", interval=MS_POR_ANIO, disabled=True, n_intervals=0),
    ])


def fila_kpis():
    """Fila 1: el velocímetro del indicador (KPI principal) y luego las 5 tarjetas."""
    principal = panel("graf-velocimetro", "IRD", None, area="principal", alto=130, icono_panel="medidor",
                      debajo=html.P(id="texto-velocimetro", className="velocimetro-texto"))
    tarjetas = [
        html.Div(className="kpi", children=[
            html.Div(className="kpi-cabeza", children=[
                html.Span(icono(k.get("icono", "alerta")), className="kpi-icono"),
                html.Span(k["titulo"], className="kpi-titulo")]),
            html.Div(id=f"kpi-{k['id']}-valor", className="kpi-valor"),
            html.Div(id=f"kpi-{k['id']}-delta", className="kpi-delta"),
            dcc.Graph(id=f"kpi-{k['id']}-spark", config=CONFIG_MINI, style={"height": "52px"})])
        for k in KPIS]
    return html.Section(className="fila-kpis", children=[principal] + tarjetas)


def delta_kpi(actual, previo, anio):
    if es_vacio(actual):
        return html.Span("Sin datos este año", className="delta delta-neutro")
    if es_vacio(previo):
        return html.Span(f"Sin dato de {anio - 1}", className="delta delta-neutro")
    if float(previo) == 0 and float(actual) > 0:
        return html.Span([html.Span("▲", className="delta-icono"), f" desde 0 en {anio - 1}"],
                         className="delta delta-sube")
    cambio = variacion_pct(actual, previo) or 0.0
    if abs(cambio) < 0.5:
        return html.Span(f"= igual que {anio - 1}", className="delta delta-neutro")
    sube = cambio > 0
    return html.Span([html.Span("▲" if sube else "▼", className="delta-icono"),
                      f" {numero(abs(cambio), 0)} % vs {anio - 1}"],
                     className="delta " + ("delta-sube" if sube else "delta-baja"))


def leyenda_cuartiles():
    return html.Ul(className="leyenda", children=[
        html.Li([html.Span(className="chip-color", style={"background": COLORES_CUARTIL[q]}),
                 NOMBRES_CUARTIL[q]])
        for q in sorted(COLORES_CUARTIL, reverse=True)])


def panel(id_grafico, titulo, descripcion, area, alto=380, debajo=None, leyenda=None, icono_panel=None):
    titulo_html = [html.H3(titulo, id=f"{id_grafico}-titulo")]
    if icono_panel:
        titulo_html.insert(0, html.Span(icono(icono_panel), className="panel-icono"))
    cabecera = [html.Div(titulo_html, className="panel-titulo")]
    if descripcion:
        cabecera.append(html.P(descripcion, className="panel-descripcion"))
    hijos = [html.Header(className="panel-cabecera", children=cabecera)]
    if leyenda is not None:
        hijos.append(leyenda)
    hijos.append(html.Div(className="grafico", style={"height": f"{alto}px"}, children=dcc.Graph(
        id=id_grafico, config=CONFIG_GRAFICO, responsive=True, style={"height": "100%"})))
    if debajo is not None:
        hijos.append(debajo)
    return html.Section(className=f"panel area-{area}", children=hijos)


def crear_layout(base):
    anios = lista_anios(base)
    departamentos = sorted(base[COL_DPTO].unique())

    cabecera = html.Header(className="cabecera", children=[
        html.H1(id="titulo-panorama"),
        html.P(SUBTITULO, className="cabecera-sub")])

    # Fila 2: mapa, ranking y quién sube / quién baja
    fila_2 = html.Div(className="fila fila-3", children=[
        panel("graf-mapa", "Mapa por cuartil",
              "Colores según el cuartil del indicador elegido. Clic en un departamento para elegirlo.",
              area="mapa", alto=480, icono_panel="mapa"),
        panel("graf-ranking", "Ranking de departamentos",
              "Valor del indicador en el año, con el color de su cuartil. Clic para elegir un departamento.",
              area="ranking", alto=450, icono_panel="ranking", leyenda=leyenda_cuartiles()),
        panel("graf-variacion", "¿Quién sube y quién baja?",
              "Cambio del indicador de cada departamento respecto al año anterior.",
              area="variacion", alto=480, icono_panel="subebaja"),
    ])
    # Fila 3: qué explica el indicador y cómo cambió el tipo de daño
    fila_3 = html.Div(className="fila fila-2", children=[
        panel("graf-anual", "¿Qué explica el indicador cada año?",
              "Aporte de cada dimensión del daño. Haz clic en una barra para elegir ese año.",
              area="anual", alto=440, icono_panel="barras"),
        panel("graf-composicion", "¿Cambió el tipo de daño?",
              "Parte de la severidad que explica cada dimensión, año por año.",
              area="composicion", alto=440, icono_panel="dona"),
    ])
    # Fila 4: línea de tiempo y matriz de riesgo (frecuencia contra severidad)
    fila_4 = html.Div(className="fila fila-2", children=[
        panel("graf-tendencia", "Evolución del IRD",
              "La banda gris es el rango típico de los departamentos; la línea punteada, la tendencia.",
              area="tendencia", alto=500, icono_panel="tendencia"),
        panel("graf-matriz", "Matriz de riesgo: frecuencia contra severidad",
              "Cada burbuja es un departamento; su tamaño son los damnificados y su color, el cuartil. "
              "Pulsa «Reproducir» para ver cómo se movieron año a año.",
              area="matriz", alto=500, icono_panel="burbujas"),
    ])
    # Al final: mapa de calor
    fila_5 = html.Div(className="fila", children=[
        panel("graf-calor", "Mapa de calor: 23 años de un vistazo",
              "Cada celda es el cuartil de un departamento en un año. Ordenado por el promedio de los "
              "últimos 5 años. Clic en una celda para elegir departamento y año.",
              area="calor", alto=640, icono_panel="cuadricula", leyenda=leyenda_cuartiles()),
    ])
    contenido = html.Main(className="contenido", children=[
        cabecera, fila_kpis(), fila_2, fila_3, fila_4, fila_5,
        html.Footer(className="pie", children=[html.P(FUENTE_DATOS)]),
    ])
    return html.Div(className="app", children=[barra_lateral(departamentos, anios), contenido])


# ---- Página de índices climáticos ---------------------------------------------
def barra_lateral_clima(clima):
    n = len(clima)
    marcas = {i: str(a) for i, (a, m) in enumerate(zip(clima["anio"], clima["mes"]))
              if m == 1 and (a - clima["anio"].iloc[0]) % 4 == 0 and i < n - 30}
    marcas[n - 1] = clima["etiqueta"].iloc[-1]
    return html.Aside(className="lateral", children=[
        dcc.Link([icono("volver"), "Portada"], href="/", className="lateral-volver"),
        html.Div(className="marca", children=[html.Span("ICM", className="marca-sigla"),
                                              html.Span("Índices climáticos mensuales", className="marca-nombre")]),
        html.Div(className="seleccion", children=[html.Div(id="clima-lateral-anio", className="seleccion-anio"),
                                                  html.Div(id="clima-lateral-mes", className="seleccion-dpto")]),
        html.Div(className="controles", children=[
            html.Label([icono("indicador"), "Índice"], className="control-etiqueta"),
            dcc.RadioItems(id="clima-indice", value=list(INDICES_CLIMA)[0], className="control-radio",
                           options=[{"label": k, "value": k} for k in INDICES_CLIMA]),
            html.Div(className="control-fila", children=[
                html.Label([icono("calendario"), "Mes"], className="control-etiqueta"),
                html.Button("▶ Reproducir", id="clima-btn-play", n_clicks=0, className="boton-play")]),
            dcc.Slider(id="clima-mes", min=0, max=n - 1, step=1, value=n - 1, marks=marcas, included=False),
            html.Button(f"Ir al último mes ({clima['etiqueta'].iloc[-1]})", id="clima-btn-ultimo",
                        n_clicks=0, className="boton-secundario")]),
        html.P("Consejo: haz clic en una tarjeta para ver ese índice, o en el calendario para elegir un mes.",
               className="lateral-nota"),
        dcc.Interval(id="clima-intervalo", interval=700, disabled=True, n_intervals=0),
    ])


def tarjeta_clima(clave):
    cfg = INDICES_CLIMA[clave]
    return html.Div(id=f"clima-tarjeta-{clave}", n_clicks=0, className="clima-tarjeta", children=[
        html.Div(className="clima-tarjeta-cabeza", children=[
            html.Span(icono("ola"), className="kpi-icono"),
            html.Div([html.Span(clave, className="clima-sigla"),
                      html.Span(cfg["nombre"], className="clima-nombre")])]),
        dcc.Graph(id=f"clima-graf-{clave}", config={"displayModeBar": False, "staticPlot": True},
                  style={"height": "170px"}),
        html.P(id=f"clima-valor-{clave}", className="clima-valor"),
        html.Div(id=f"clima-estado-{clave}", className="clima-estado"),
        html.Div(id=f"clima-delta-{clave}", className="kpi-delta clima-delta"),
    ])


def tabla_criterio(clave):
    cfg = INDICES_CLIMA[clave]
    segs = segmentos_clima(cfg)
    varias = len(cfg["calidas"]) > 1
    cabecera = html.Thead(html.Tr([html.Th("Categoría"), html.Th("Magnitud" if varias else ""),
                                   html.Th("Intervalo")]))
    filas, anterior = [], None
    for seg in segs:
        grupo = seg["grupo"]
        filas.append(html.Tr([
            html.Td(grupo if grupo != anterior else "", className="criterio-grupo"),
            html.Td([html.Span(className="criterio-color", style={"background": seg["color"]}),
                     "" if grupo == "Neutro" or not varias else seg["magnitud"]]),
            html.Td(intervalo_texto(seg, cfg), className="num")]))
        anterior = grupo
    return html.Table(className="criterio", children=[cabecera, html.Tbody(filas)])


def crear_layout_clima(clima):
    primero, ultimo = clima["etiqueta"].iloc[0], clima["etiqueta"].iloc[-1]
    cabecera = html.Header(className="cabecera", children=[
        html.H1(id="clima-titulo"),
        html.P(f"Condición del mar según cuatro índices mensuales, {primero} – {ultimo}",
               className="cabecera-sub")])
    criterio = html.Section(className="panel area-criterio", children=[
        html.Header(className="panel-cabecera", children=[
            html.Div([html.Span(icono("ranking"), className="panel-icono"),
                      html.H3(id="clima-criterio-titulo")], className="panel-titulo"),
            html.P("Rangos con los que se clasifica cada mes.", className="panel-descripcion")]),
        html.Div(id="clima-criterio", className="criterio-contenedor")])
    fila_evolucion = html.Div(className="fila fila-historia", children=[
        panel("clima-graf-evolucion", "Evolución del índice",
              "Rojo: por encima de cero (más cálido de lo normal). Azul: por debajo. Las líneas punteadas son "
              "los umbrales del criterio. Clic en la línea para elegir un mes.",
              area="evolucion", alto=440, icono_panel="tendencia"),
        criterio])
    fila_variacion = html.Div(className="fila fila-2", children=[
        panel("clima-graf-variacion", "Variación a 12 meses",
              "Diferencia entre cada mes y el mismo mes del año anterior. Rojo: se calentó; azul: se enfrió.",
              area="variacion12", alto=400, icono_panel="subebaja"),
        panel("clima-graf-anio", "Hace un año vs. hoy",
              "Valor de cada índice hace 12 meses (gris) y ahora (color de su categoría). "
              "Clic en un índice para elegirlo.",
              area="anio", alto=400, icono_panel="flecha")])
    fila_estacional = html.Div(className="fila fila-2", children=[
        panel("clima-graf-estacional", "Cada año, mes a mes",
              "Cada línea es un año, de enero a diciembre. Resaltados: el año elegido (azul) y el año con el "
              "valor más alto (rojo). La franja gris es el rango neutro.",
              area="estacional", alto=440, icono_panel="ola"),
        panel("clima-graf-top", "Los 10 meses más cálidos",
              "Meses con el valor más alto desde 2010, con el color de su categoría. "
              "El mes elegido va con borde oscuro.",
              area="top", alto=440, icono_panel="ranking")])
    fila_consenso = html.Div(className="fila fila-2", children=[
        panel("clima-graf-consenso", "¿Coinciden los cuatro índices?",
              f"Cuántos índices marcan condición cálida, neutra o fría en cada uno de los últimos "
              f"{MESES_VENTANA} meses. Clic en una barra para elegir ese mes.",
              area="consenso", alto=470, icono_panel="barras"),
        panel("clima-graf-calendario", "Calendario del índice",
              "Cada celda es un mes, con el color de su categoría. Clic en una celda para elegir ese mes.",
              area="calendario", alto=470, icono_panel="cuadricula")])
    contenido = html.Main(className="contenido", children=[
        cabecera,
        html.Section(className="clima-tarjetas", children=[tarjeta_clima(k) for k in INDICES_CLIMA]),
        fila_evolucion, fila_variacion, fila_estacional, fila_consenso,
        html.Footer(className="pie", children=[html.P(
            "Fuente: ICEM.xlsx. Los valores marcados con * en el Excel se muestran como preliminares.")]),
    ])
    return html.Div(className="app", children=[barra_lateral_clima(clima), contenido])


# ---- Página de América Latina ---------------------------------------------------
KPIS_LATAM = [  # (id, título, ícono)
    ("eventos", "Desastres registrados", "alerta"),
    ("fallecidos", "Personas fallecidas", "persona"),
    ("afectados", "Personas afectadas", "personas"),
    ("dano", "Costo económico (US$ millones)", "barras"),
    ("costo_pbi", "Costo económico (% del PBI)", "medidor"),
]


def barra_lateral_latam(latam):
    anios = latam["eventos"]["anio"]
    a0, a1 = int(anios.min()), int(anios.max())
    marcas = {a: str(a) for a in range(a0, a1 + 1) if (a - a0) % 10 == 0 and a1 - a >= 8}
    marcas[a1] = str(a1)
    paises = sorted(latam["eventos"]["pais"].unique())
    return html.Aside(className="lateral", children=[
        dcc.Link([icono("volver"), "Portada"], href="/", className="lateral-volver"),
        html.Div(className="marca", children=[html.Span("LAT", className="marca-sigla"),
                                              html.Span("Estadísticas de América Latina", className="marca-nombre")]),
        html.Div(className="seleccion", children=[html.Div(id="latam-lateral-pais", className="seleccion-pais"),
                                                  html.Div(id="latam-lateral-periodo", className="seleccion-dpto")]),
        html.Div(className="controles", children=[
            html.Label([icono("mapa"), "País"], className="control-etiqueta"),
            dcc.Dropdown(id="latam-pais", options=[{"label": p, "value": p} for p in paises],
                         value=PAIS_INICIAL, clearable=False),
            html.Label([icono("calendario"), "Periodo"], className="control-etiqueta"),
            dcc.RangeSlider(id="latam-periodo", min=a0, max=a1, step=1, value=[a0, a1], marks=marcas,
                            allowCross=False, tooltip={"placement": "bottom", "always_visible": False}),
            html.Label([icono("indicador"), "Indicador del ranking y el mapa"], className="control-etiqueta"),
            dcc.RadioItems(id="latam-indicador", value="eventos", className="control-radio",
                           options=[{"label": v, "value": k} for k, v in INDICADORES_LATAM.items()]),
            html.Button(f"Volver a {PAIS_INICIAL}, {a0}–{a1}", id="latam-btn-reiniciar", n_clicks=0,
                        className="boton-secundario")]),
        html.P("Consejo: haz clic en una barra del ranking o en el mapa para elegir un país.",
               className="lateral-nota"),
    ])


def tarjetas_latam():
    tarjetas = []
    for i, (clave, titulo, ic) in enumerate(KPIS_LATAM):
        tarjetas.append(html.Div(className="kpi" + (" kpi-destacado" if i == 0 else ""), children=[
            html.Div(className="kpi-cabeza", children=[html.Span(icono(ic), className="kpi-icono"),
                                                       html.Span(titulo, className="kpi-titulo")]),
            html.Div(id=f"latam-kpi-{clave}-valor", className="kpi-valor"),
            html.Div(id=f"latam-kpi-{clave}-nota", className="kpi-delta")]))
    return html.Section(className="latam-kpis", children=tarjetas)


def crear_layout_latam(latam):
    cabecera = html.Header(className="cabecera", children=[
        html.H1(id="latam-titulo"),
        html.P("Desastres naturales en 26 países de América Latina y el Caribe. Fuente: EM-DAT.",
               className="cabecera-sub")])
    fila_1 = html.Div(className="fila fila-2", children=[
        panel("latam-graf-ranking", "Ranking de países",
              "Colores por cuartil: el más oscuro es el 25 % de países con el valor más alto. "
              "Clic en una barra para elegir un país.", area="ranking-latam", alto=640, icono_panel="ranking"),
        panel("latam-graf-mapa", "Mapa de América Latina",
              "Mismos colores que el ranking. Clic en un país para elegirlo.",
              area="mapa-latam", alto=640, icono_panel="globo")])
    fila_2 = html.Div(className="fila fila-2", children=[
        panel("latam-graf-periodos", "Desastres por país en dos periodos",
              "Número de desastres en 1980–2002 (gris) y 2003–2026 (rojo). No depende del periodo elegido.",
              area="periodos", alto=440, icono_panel="barras"),
        panel("latam-graf-dispersion", "¿Más desastres, más costo?",
              "Cada punto es un año de la región: número de desastres contra su costo económico. "
              "La línea punteada es la tendencia.", area="dispersion", alto=440, icono_panel="burbujas")])
    fila_3 = html.Div(className="fila fila-2", children=[
        panel("latam-graf-zona-eventos", "Desastres según tipo de evento",
              "Parte de los desastres de cada país según su origen.", area="zona-eventos", alto=440,
              icono_panel="dona"),
        panel("latam-graf-zona-costo", "Costo económico según tipo de evento",
              "Parte del costo económico de cada país según el origen del desastre.", area="zona-costo",
              alto=440, icono_panel="dona")])
    fila_4 = html.Div(className="fila fila-2", children=[
        panel("latam-graf-pais-eventos", "Desastres por tipo y subtipo",
              "Número de desastres del país elegido, agrupados por origen.", area="pais-eventos", alto=520,
              icono_panel="alerta"),
        panel("latam-graf-pais-costo", "Costo por tipo y subtipo (% del PBI)",
              "Daños acumulados de cada tipo como porcentaje del PBI del país.", area="pais-costo", alto=520,
              icono_panel="medidor")])
    fila_5 = html.Div(className="fila", children=[
        panel("latam-graf-evolucion", "Evolución anual de los desastres en la región",
              "Barras: desastres de los 26 países por origen. Línea: el país elegido.", area="evolucion-latam",
              alto=420, icono_panel="tendencia")])
    contenido = html.Main(className="contenido", children=[
        cabecera, tarjetas_latam(), fila_1, fila_2, fila_3, fila_4, fila_5,
        html.Footer(className="pie", children=[html.P(
            "Fuente: EM-DAT (archivo EMDAT_LATAM.xlsx). Costos ajustados por inflación; el % del PBI compara los "
            "daños acumulados del periodo con el PBI del último año disponible. Hay costos hasta 2025.")]),
    ])
    return html.Div(className="app", children=[barra_lateral_latam(latam), contenido])


# ---- Portada y páginas en preparación ---------------------------------------
def tarjeta_seccion(numero_tarjeta, seccion, dato):
    """Un botón grande de la portada (toda la tarjeta es clicable)."""
    lista = seccion["lista"]
    return dcc.Link(href=seccion["ruta"], className="tarjeta" + ("" if lista else " tarjeta-pendiente"),
                    children=[
        html.Div(className="tarjeta-cabeza", children=[
            html.Span(icono(seccion["icono"]), className="tarjeta-icono"),
            html.Span(f"{numero_tarjeta:02d}", className="tarjeta-numero")]),
        html.Span("Disponible" if lista else "En preparación",
                  className="estado " + ("estado-listo" if lista else "estado-pendiente")),
        html.H2(seccion["titulo"]),
        html.P(seccion["descripcion"], className="tarjeta-texto"),
        html.Ul([html.Li(e) for e in seccion["etiquetas"]], className="etiquetas"),
        html.P(dato, className="tarjeta-dato"),
        html.Span([("Entrar al tablero" if lista else "Ver sección"), icono("flecha")],
                  className="tarjeta-boton"),
    ])


def crear_portada(base, clima=None, latam=None):
    anios = lista_anios(base)
    ultimo = anios[-1]
    ird_nacional = base.loc[base[COL_ANIO] == ultimo, COL_IRD].mean()
    partes = [f"{anios[0]}–{ultimo}", f"{base[COL_DPTO].nunique()} departamentos",
              f"IRD nacional {ultimo}: {numero(ird_nacional, 1)}"]
    dato_ird = [html.Span(t + (" · " if i < len(partes) - 1 else "")) for i, t in enumerate(partes)]
    datos = {RUTA_IRD: dato_ird}
    if clima is not None:
        datos[RUTA_CLIMA] = [html.Span(f"{clima['etiqueta'].iloc[0]} – {clima['etiqueta'].iloc[-1]} · "),
                             html.Span(f"{len(INDICES_CLIMA)} índices mensuales")]
    if latam is not None:
        e = latam["eventos"]
        datos[RUTA_LATAM] = [html.Span(f"{e['anio'].min()}–{e['anio'].max()} · "),
                             html.Span(f"{e['pais'].nunique()} países · "),
                             html.Span(f"{numero(len(e), 0)} desastres")]
    tarjetas = [tarjeta_seccion(i, s, datos.get(s["ruta"], "Base de datos por cargar"))
                for i, s in enumerate(SECCIONES, start=1)]
    return html.Div(className="portada", children=[
        html.Header(className="portada-hero", children=html.Div(className="portada-interior", children=[
            html.P(ANTETITULO_PORTADA, className="portada-antetitulo"),
            html.H1(TITULO_PORTADA),
            html.P(SUBTITULO_PORTADA, className="portada-sub")])),
        html.Nav(tarjetas, className="portada-tarjetas", **{"aria-label": "Tableros"}),
        html.Footer(className="portada-pie", children=[
            html.P("Elige un tablero para empezar. Puedes volver a esta portada desde cualquier página.")]),
    ])


def pagina_en_preparacion(seccion):
    return html.Div(className="pagina-simple", children=[
        dcc.Link([icono("volver"), "Volver a la portada"], href="/", className="enlace-volver"),
        html.Section(className="aviso", children=[
            html.Div(className="tarjeta-cabeza", children=[
                html.Span(icono(seccion["icono"]), className="tarjeta-icono"),
                html.Span("En preparación", className="estado estado-pendiente")]),
            html.H1(seccion["titulo"]),
            html.P(seccion["descripcion"], className="tarjeta-texto"),
            html.Ul([html.Li(e) for e in seccion["etiquetas"]], className="etiquetas"),
            html.Div(className="aviso-nota", children=[
                html.Span(icono("reloj"), className="aviso-nota-icono"),
                html.P("Este tablero se activará cuando se cargue su base de datos.")]),
        ]),
    ])


def crear_marco():
    """Lo que siempre está en la página: la dirección (URL) y el espacio donde se dibuja cada página."""
    return html.Div([dcc.Location(id="url", refresh=False), dcc.Store(id="arriba"),
                     html.Div(id="pagina")])


CSS = """
:root{--pacifico:#0b3a66;--pacifico-medio:#1f6fb2;--niebla:#edf3f9;--azul-suave:#e3eefa;
--superficie:#fff;--borde:#d5e2ef;--tinta:#0f2f50;--tinta-2:#3f6488;--tinta-3:#6f8cab;
--sube:#c93a39;--baja:#2367b8;--fuente:"Archivo","Segoe UI",Roboto,Helvetica,Arial,sans-serif;
--radio:12px;--espacio:18px;
--Dash-Fill-Interactive-Strong:#1f6fb2;--Dash-Text-Strong:#0f2f50}
*{box-sizing:border-box}
html,body{max-width:100%;overflow-x:hidden}
body{margin:0;background:var(--niebla);color:var(--tinta);font-family:var(--fuente);font-size:15px;line-height:1.5}
h1,h2,h3,h4,p{margin:0}
.app{display:grid;grid-template-columns:288px minmax(0,1fr);min-height:100vh}
.contenido{padding:26px 30px 40px;display:flex;flex-direction:column;gap:var(--espacio);min-width:0}

/* Barra lateral */
.lateral{position:sticky;top:0;height:100vh;overflow-y:auto;color:#fff;
background:linear-gradient(180deg,#0b3a66 0%,#0a2f54 100%);
padding:26px 22px 22px;display:flex;flex-direction:column;gap:18px}
.marca{display:flex;align-items:baseline;gap:10px}
.marca-sigla{font-weight:800;font-stretch:125%;font-size:15px;padding:2px 7px;border:1.5px solid rgba(255,255,255,.55);border-radius:6px}
.marca-nombre{font-size:13px;color:rgba(255,255,255,.78)}
.seleccion-anio{font-size:80px;line-height:.9;font-weight:800;font-stretch:125%;letter-spacing:-.035em;font-variant-numeric:tabular-nums}
.seleccion-dpto{margin-top:8px;font-size:24px;font-weight:500;color:#9fcbf2}
.controles{background:var(--superficie);color:var(--tinta);border-radius:var(--radio);padding:16px 16px 18px;display:flex;flex-direction:column;gap:8px}
.control-etiqueta{display:flex;align-items:center;gap:7px;font-size:13px;font-weight:650;color:var(--pacifico);margin-top:6px}
.control-etiqueta .icono{width:16px;height:16px;color:var(--pacifico-medio)}
.control-fila{display:flex;align-items:center;justify-content:space-between;margin-top:6px}
.control-fila .control-etiqueta{margin-top:0}
.control-radio{display:flex;flex-direction:column;gap:4px;font-size:14px}
.control-radio label{display:flex;align-items:center;gap:8px;cursor:pointer;color:var(--tinta)!important}
.control-radio input{accent-color:var(--pacifico-medio)}
.boton-play,.boton-secundario{font-family:var(--fuente);font-size:13px;font-weight:600;border-radius:999px;cursor:pointer}
.boton-play{border:0;background:var(--pacifico-medio);color:#fff;padding:5px 12px}
.boton-play:hover{background:var(--pacifico)}
.boton-secundario{margin-top:8px;border:1px solid var(--borde);background:var(--superficie);color:var(--tinta);padding:7px 14px}
.boton-secundario:hover{background:var(--niebla)}
.lateral-nota{margin-top:auto;font-size:12.5px;color:rgba(255,255,255,.64)}

/* Cabecera */
.cabecera h1{font-size:34px;line-height:1.12;font-weight:700;font-stretch:112%;letter-spacing:-.015em;color:var(--pacifico)}
.cabecera-sub{font-size:15px;color:var(--pacifico-medio);font-weight:500;margin-top:4px}

/* Fila 1: KPI principal + 5 KPIs */
.fila-kpis{display:grid;gap:14px;grid-template-columns:minmax(0,1.3fr) repeat(5,minmax(0,1fr))}
.panel.area-principal{border:2px solid var(--pacifico-medio);padding:12px 12px 6px;
background:linear-gradient(180deg,#f3f8fd 0%,#fff 60%)}
.area-principal .panel-titulo{justify-content:center}
.area-principal .panel-cabecera h3{font-size:15px}
.velocimetro-texto{font-size:12px;line-height:1.35;color:var(--tinta-2);text-align:center;padding:0 2px 4px}
.kpi{background:var(--superficie);border:1px solid var(--borde);border-radius:var(--radio);padding:14px 16px 8px;min-width:0;
display:flex;flex-direction:column}
.kpi .dash-graph{margin-top:auto}
.kpi-cabeza{display:flex;align-items:center;gap:10px;min-height:2.9em}
.kpi-icono{width:38px;height:38px;border-radius:11px;background:var(--azul-suave);color:var(--pacifico-medio);
display:grid;place-items:center;flex:none}
.kpi-icono .icono{width:21px;height:21px}
.kpi-titulo{font-size:13px;line-height:1.3;color:var(--tinta-2);font-weight:500}
.kpi-valor{margin-top:6px;font-size:28px;font-weight:700;font-stretch:110%;line-height:1.2;font-variant-numeric:tabular-nums;white-space:nowrap;color:var(--pacifico)}
.kpi-delta{font-size:12.5px;min-height:19px}
.delta-sube{color:var(--sube)}.delta-baja{color:var(--baja)}.delta-neutro{color:var(--tinta-3)}
.delta-icono{font-size:10px}

/* Paneles y filas */
.panel{background:var(--superficie);border:1px solid var(--borde);border-radius:var(--radio);padding:16px 16px 10px;min-width:0;
transition:box-shadow .2s ease}
.panel:hover{box-shadow:0 6px 18px rgba(11,58,102,.08)}
.panel-cabecera{margin-bottom:4px}
.panel-titulo{display:flex;align-items:center;gap:10px}
.panel-icono{width:30px;height:30px;border-radius:9px;background:var(--azul-suave);color:var(--pacifico-medio);
display:grid;place-items:center;flex:none}
.panel-icono .icono{width:17px;height:17px}
.panel-cabecera h3{font-size:16.5px;font-weight:650;line-height:1.3;color:var(--pacifico)}
.panel-descripcion{margin-top:4px;font-size:13px;color:var(--tinta-2);max-width:70ch}
.grafico{width:100%}
.leyenda{list-style:none;margin:8px 0 0;padding:0;display:flex;flex-wrap:wrap;gap:4px 16px;font-size:12.5px;color:var(--tinta-2)}
.leyenda li{display:inline-flex;align-items:center;gap:6px}
.chip-color{width:10px;height:10px;border-radius:3px;display:inline-block;flex:none}
.icono{display:inline-block;width:18px;height:18px;flex:none;background-color:currentColor;
-webkit-mask:var(--i) center/contain no-repeat;mask:var(--i) center/contain no-repeat}
.fila{display:grid;gap:var(--espacio);grid-template-columns:minmax(0,1fr)}
.fila-3{grid-template-columns:repeat(3,minmax(0,1fr))}
.fila-2{grid-template-columns:repeat(2,minmax(0,1fr))}
.pie{color:var(--tinta-3);font-size:12.5px}
.rc-slider-mark-text{font-size:11px;color:var(--tinta-3)!important}
.rc-slider-handle{border-color:var(--pacifico-medio)!important}

/* Enlace a la portada (barra lateral del IRD) */
.lateral-volver{display:inline-flex;align-items:center;gap:7px;align-self:flex-start;color:#9fcbf2;
font-size:13px;font-weight:600;text-decoration:none;margin-bottom:-6px}
.lateral-volver:hover{color:#fff}
.lateral-volver .icono{width:16px;height:16px}

/* Portada */
.portada{min-height:100vh;display:flex;flex-direction:column}
.portada-hero{position:relative;overflow:hidden;color:#fff;padding:96px 0 300px;min-height:640px;
background:#0b3a66 url("/portada.jpg") center top/cover no-repeat}
.portada-hero::before{content:"";position:absolute;inset:0;pointer-events:none;
background:linear-gradient(180deg,rgba(237,243,249,0) 55%,#edf3f9 100%),
linear-gradient(100deg,rgba(6,30,54,.88) 0%,rgba(6,30,54,.55) 42%,rgba(6,30,54,0) 72%)}
.portada-hero h1,.portada-sub{text-shadow:0 2px 14px rgba(0,0,0,.35)}
.portada-interior,.portada-tarjetas,.portada-pie{width:100%;max-width:1200px;margin-left:auto;margin-right:auto}
.portada-interior{position:relative;z-index:1;padding:0 48px}
.portada-antetitulo{font-size:13px;font-weight:650;letter-spacing:.09em;text-transform:uppercase;color:#9fcbf2}
.portada h1{margin-top:14px;font-size:54px;line-height:1.04;font-weight:800;font-stretch:118%;letter-spacing:-.025em;max-width:20ch}
.portada-sub{margin-top:18px;font-size:18px;line-height:1.5;color:rgba(255,255,255,.82);max-width:58ch}
.portada-tarjetas{position:relative;z-index:2;margin-top:-210px;padding:0 48px;display:grid;gap:22px;
grid-template-columns:repeat(3,minmax(0,1fr))}
.tarjeta{display:flex;flex-direction:column;gap:12px;background:rgba(255,255,255,.96);border:1px solid var(--borde);
border-radius:16px;padding:24px 24px 22px;color:var(--tinta);text-decoration:none;
box-shadow:0 10px 30px rgba(11,58,102,.10);transition:transform .2s ease,box-shadow .2s ease,border-color .2s ease}
.tarjeta:hover{transform:translateY(-4px);border-color:var(--pacifico-medio);box-shadow:0 18px 40px rgba(11,58,102,.16)}
.tarjeta:focus-visible{outline:3px solid #9fcbf2;outline-offset:3px}
.tarjeta-cabeza{display:flex;align-items:center;justify-content:space-between;gap:12px}
.tarjeta-icono{width:54px;height:54px;border-radius:15px;background:var(--azul-suave);color:var(--pacifico-medio);
display:grid;place-items:center;flex:none}
.tarjeta-icono .icono{width:29px;height:29px}
.tarjeta-numero{font-size:28px;font-weight:800;font-stretch:118%;color:var(--borde);font-variant-numeric:tabular-nums}
.estado{align-self:flex-start;display:inline-flex;align-items:center;font-size:12px;font-weight:650;padding:3px 10px;border-radius:999px}
.estado-listo{background:#e2f3ea;color:#0f6b48}
.estado-pendiente{background:#fff2dc;color:#8a5200}
.tarjeta h2,.aviso h1{font-size:22px;line-height:1.2;font-weight:700;font-stretch:108%;color:var(--pacifico)}
.tarjeta-texto{font-size:14.5px;color:var(--tinta-2)}
.etiquetas{list-style:none;margin:0;padding:0;display:flex;flex-wrap:wrap;gap:6px}
.etiquetas li{font-size:12px;font-weight:600;color:var(--pacifico-medio);background:var(--niebla);
border:1px solid var(--borde);border-radius:6px;padding:2px 8px}
.tarjeta-dato{margin-top:auto;padding-top:12px;border-top:1px solid var(--borde);font-size:13px;color:var(--tinta-3);
font-variant-numeric:tabular-nums}
.tarjeta-dato span{white-space:nowrap}
.tarjeta-boton{display:inline-flex;align-items:center;justify-content:center;gap:8px;border-radius:999px;
padding:11px 18px;font-size:14.5px;font-weight:650;background:var(--pacifico-medio);color:#fff;
border:1.5px solid var(--pacifico-medio);transition:background .2s ease,color .2s ease}
.tarjeta:hover .tarjeta-boton{background:var(--pacifico);border-color:var(--pacifico)}
.tarjeta-pendiente .tarjeta-boton{background:var(--superficie);color:var(--pacifico-medio)}
.tarjeta-pendiente:hover .tarjeta-boton{background:var(--azul-suave);color:var(--pacifico);border-color:var(--pacifico-medio)}
.tarjeta-boton .icono{width:18px;height:18px;transition:transform .2s ease}
.tarjeta:hover .tarjeta-boton .icono{transform:translateX(3px)}
.portada-pie{padding:28px 48px 36px;color:var(--tinta-3);font-size:13px}

/* Índices climáticos */
.clima-tarjetas{display:grid;gap:14px;grid-template-columns:repeat(4,minmax(0,1fr))}
.clima-tarjeta{background:var(--superficie);border:2px solid var(--borde);border-radius:var(--radio);
padding:14px 14px 12px;min-width:0;cursor:pointer;display:flex;flex-direction:column;gap:4px;
transition:border-color .2s ease,box-shadow .2s ease}
.clima-tarjeta:hover{box-shadow:0 6px 18px rgba(11,58,102,.10)}
.clima-tarjeta.activa{border-color:var(--pacifico-medio);background:linear-gradient(180deg,#f3f8fd 0%,#fff 55%)}
.clima-tarjeta-cabeza{display:flex;align-items:center;gap:10px}
.clima-sigla{display:block;font-size:17px;font-weight:750;font-stretch:112%;color:var(--pacifico);line-height:1.15}
.clima-nombre{display:block;font-size:12px;color:var(--tinta-2);line-height:1.3}
.clima-valor{text-align:center;font-size:15px;color:var(--tinta)}
.clima-valor strong{font-size:24px;font-weight:750;font-stretch:110%;color:var(--pacifico);font-variant-numeric:tabular-nums}
.clima-valor sub{font-size:11px;color:var(--tinta-2)}
.clima-estado{align-self:center;display:inline-flex;align-items:center;gap:7px;font-size:13px;font-weight:650;
color:var(--tinta);background:var(--niebla);border:1px solid var(--borde);border-radius:999px;padding:3px 12px}
.clima-estado .criterio-color{margin-right:0}
.clima-delta{text-align:center}
.fila-historia{grid-template-columns:minmax(0,2fr) minmax(0,1fr)}
.criterio-contenedor{margin-top:10px;overflow-x:auto}
.criterio{width:100%;border-collapse:collapse;font-size:13.5px}
.criterio th{text-align:left;font-size:12px;font-weight:700;color:var(--pacifico);text-transform:uppercase;
letter-spacing:.04em;padding:8px 8px;border-bottom:2px solid var(--borde)}
.criterio td{padding:7px 8px;border-bottom:1px solid var(--borde);color:var(--tinta);vertical-align:middle}
.criterio td.num{text-align:right;font-variant-numeric:tabular-nums;white-space:nowrap}
.criterio-grupo{font-weight:600;color:var(--tinta-2)}
.criterio-color{display:inline-block;width:22px;height:16px;border-radius:3px;border:1px solid #9aa9b8;
vertical-align:middle;margin-right:8px}

/* América Latina */
.seleccion-pais{font-size:44px;line-height:1;font-weight:800;font-stretch:118%;letter-spacing:-.02em}
.latam-kpis{display:grid;gap:14px;grid-template-columns:repeat(5,minmax(0,1fr))}
.kpi.kpi-destacado{border:2px solid var(--pacifico-medio);background:linear-gradient(180deg,#f3f8fd 0%,#fff 60%)}
.latam-kpis .kpi{padding-bottom:12px}

/* Páginas en preparación */
.pagina-simple{max-width:760px;margin:0 auto;padding:40px 24px 56px}
.enlace-volver{display:inline-flex;align-items:center;gap:8px;color:var(--pacifico-medio);font-size:14px;
font-weight:600;text-decoration:none}
.enlace-volver:hover{color:var(--pacifico)}
.enlace-volver .icono{width:17px;height:17px}
.aviso{margin-top:18px;background:var(--superficie);border:1px solid var(--borde);border-radius:16px;
padding:32px;display:flex;flex-direction:column;gap:14px}
.aviso h1{font-size:30px}
.aviso-nota{display:flex;align-items:center;gap:12px;margin-top:8px;padding:14px 16px;border-radius:12px;
background:var(--niebla);color:var(--tinta-2);font-size:14px}
.aviso-nota-icono{color:var(--pacifico-medio);display:grid;place-items:center}

/* Pantallas medianas */
@media (max-width:1280px){
.clima-tarjetas{grid-template-columns:repeat(2,minmax(0,1fr))}
.fila-historia{grid-template-columns:minmax(0,1fr)}
.latam-kpis{grid-template-columns:repeat(3,minmax(0,1fr))}
.fila-kpis{grid-template-columns:repeat(3,minmax(0,1fr))}
.fila-3{grid-template-columns:repeat(2,minmax(0,1fr))}
.fila-3 .area-variacion{grid-column:span 2}}

/* Celulares */
@media (max-width:900px){
.app{grid-template-columns:minmax(0,1fr)}
.lateral{position:static;height:auto;padding:18px 16px;gap:12px}
.seleccion{display:flex;align-items:baseline;gap:12px;flex-wrap:wrap}
.seleccion-anio{font-size:54px}
.seleccion-dpto{margin-top:0;font-size:22px}
.lateral-nota{display:none}
.contenido{padding:18px 12px 28px;gap:14px}
.cabecera h1{font-size:24px}
.cabecera-sub{font-size:13.5px}
.fila-kpis{grid-template-columns:repeat(2,minmax(0,1fr));gap:10px}
.fila-kpis .area-principal{grid-column:span 2}
.kpi{padding:12px 12px 6px}
.kpi-icono{width:32px;height:32px}
.kpi-icono .icono{width:18px;height:18px}
.kpi-titulo{font-size:12px}
.kpi-valor{font-size:22px}
.fila-3,.fila-2{grid-template-columns:minmax(0,1fr)}
.fila-3 .area-variacion{grid-column:auto}
.panel{padding:12px 10px 8px}
.panel-cabecera h3{font-size:15px}
.area-mapa .grafico{height:400px!important}
.area-calor .grafico{height:560px!important}
.portada-hero{padding:40px 0 200px;min-height:0}
.portada-interior{padding:0 16px}
.portada h1{font-size:32px}
.portada-sub{font-size:15.5px}
.portada-tarjetas{grid-template-columns:minmax(0,1fr);padding:0 12px;margin-top:-130px;gap:14px}
.tarjeta{padding:20px 18px 18px}
.portada-pie{padding:22px 16px 28px}
.pagina-simple{padding:24px 12px 40px}
.aviso{padding:22px 18px}
.aviso h1{font-size:24px}
.clima-tarjetas{grid-template-columns:minmax(0,1fr);gap:10px}
.latam-kpis{grid-template-columns:repeat(2,minmax(0,1fr));gap:10px}
.latam-kpis .kpi-destacado{grid-column:span 2}
.seleccion-pais{font-size:34px}}
"""

CSS = CSS + css_iconos()

PAGINA = ("<!DOCTYPE html><html lang='es'><head>{%metas%}<title>{%title%}</title>{%favicon%}{%css%}"
          "<style>" + CSS + "</style></head><body>{%app_entry%}<footer>{%config%}{%scripts%}"
          "{%renderer%}</footer></body></html>")


# =============================================================================
# 6. INTERACTIVIDAD (callbacks)
# =============================================================================
def registrar_navegacion(app, base, clima, latam):
    """Muestra la portada o el tablero según la dirección: /, /ird, /indices-mensuales, /latam."""
    secciones = {s["ruta"]: s for s in SECCIONES}

    @app.callback(Output("pagina", "children"), Input("url", "pathname"))
    def mostrar_pagina(ruta):
        ruta = (ruta or "/").rstrip("/") or "/"
        if ruta == RUTA_IRD:
            return crear_layout(base)
        if ruta == RUTA_CLIMA and clima is not None:
            return crear_layout_clima(clima)
        if ruta == RUTA_LATAM and latam is not None:
            return crear_layout_latam(latam)
        if ruta in secciones:
            return pagina_en_preparacion(secciones[ruta])
        return crear_portada(base, clima, latam)

    # Al cambiar de página, empezar desde arriba
    app.clientside_callback("function(ruta){window.scrollTo(0, 0); return ruta;}",
                            Output("arriba", "data"), Input("url", "pathname"))


def registrar_callbacks_clima(app, clima):
    claves = list(INDICES_CLIMA)
    ultimo = len(clima) - 1
    posicion = {(a, m): i for i, (a, m) in enumerate(zip(clima["anio"], clima["mes"]))}
    por_fecha = {f.strftime("%Y-%m"): i for i, f in enumerate(clima["fecha"])}
    por_etiqueta = {e: i for i, e in enumerate(clima["etiqueta"])}
    graficos_clic = ["clima-graf-calendario", "clima-graf-evolucion", "clima-graf-variacion", "clima-graf-anio",
                     "clima-graf-estacional", "clima-graf-top", "clima-graf-consenso"]

    # Clics en tarjetas y gráficos, reproducir y último mes -> índice y mes
    @app.callback(Output("clima-indice", "value"), Output("clima-mes", "value"),
                  [Input(f"clima-tarjeta-{k}", "n_clicks") for k in claves]
                  + [Input(g, "clickData") for g in graficos_clic]
                  + [Input("clima-intervalo", "n_intervals"), Input("clima-btn-ultimo", "n_clicks")],
                  State("clima-mes", "value"), prevent_initial_call=True)
    def sincronizar_clima(*args):
        mes_actual = args[-1]
        origen = ctx.triggered_id
        clic = ctx.triggered[0]["value"] if ctx.triggered else None
        puntos = (clic or {}).get("points") if isinstance(clic, dict) else None
        punto = puntos[0] if puntos else {}

        def mes(i):
            return i if i is not None else no_update

        if isinstance(origen, str) and origen.startswith("clima-tarjeta-"):
            return origen.replace("clima-tarjeta-", ""), no_update
        if origen == "clima-btn-ultimo":
            return no_update, ultimo
        if origen == "clima-intervalo":
            return no_update, (mes_actual + 1 if mes_actual < ultimo else 0)
        if origen == "clima-graf-calendario":
            return no_update, mes(posicion.get((punto.get("y"), punto.get("x"))))
        if origen == "clima-graf-evolucion":
            return no_update, mes(por_fecha.get(str(punto.get("x", ""))[:7]))
        if origen in ("clima-graf-variacion", "clima-graf-consenso"):
            return no_update, mes(por_etiqueta.get(punto.get("x")))
        if origen == "clima-graf-top":
            return no_update, mes(por_etiqueta.get(punto.get("y")))
        if origen == "clima-graf-estacional":
            custom = punto.get("customdata") or [None]
            return no_update, mes(posicion.get((custom[0], punto.get("x"))))
        if origen == "clima-graf-anio":
            clave = punto.get("y")
            return (clave if clave in INDICES_CLIMA else no_update), no_update
        return no_update, no_update

    @app.callback(Output("clima-intervalo", "disabled"), Output("clima-btn-play", "children"),
                  Input("clima-btn-play", "n_clicks"), State("clima-intervalo", "disabled"),
                  prevent_initial_call=True)
    def alternar_reproduccion_clima(_n, desactivado):
        reproducir = bool(desactivado)
        return (not reproducir), ("❚❚ Pausar" if reproducir else "▶ Reproducir")

    # Barra lateral, título y las cuatro tarjetas
    salidas = [Output("clima-lateral-anio", "children"), Output("clima-lateral-mes", "children"),
               Output("clima-titulo", "children")]
    for k in claves:
        salidas += [Output(f"clima-graf-{k}", "figure"), Output(f"clima-valor-{k}", "children"),
                    Output(f"clima-estado-{k}", "children"), Output(f"clima-delta-{k}", "children"),
                    Output(f"clima-tarjeta-{k}", "className")]

    @app.callback(salidas, Input("clima-mes", "value"), Input("clima-indice", "value"))
    def actualizar_tarjetas_clima(idx, clave_sel):
        fila = clima.loc[idx]
        previa = clima.loc[idx - 1] if idx > 0 else None
        nombre_mes = MESES[fila["mes"] - 1]
        salida = [str(fila["anio"]), nombre_mes, f"Índices climáticos, {nombre_mes.lower()} de {fila['anio']}"]
        for k in claves:
            cfg = INDICES_CLIMA[k]
            valor, dec = fila[k], cfg["decimales"]
            cat = clasificar_clima(valor, cfg)
            if cat is None:
                valor_html = html.Span(f"Sin dato en {fila['etiqueta']}")
                estado = html.Span("Sin dato")
            else:
                seg = segmentos_clima(cfg)[cat]
                valor_html = ["Valor preliminar" if fila[k + "_prelim"] else "Valor",
                              html.Sub(f" {fila['etiqueta']}"), " = ", html.Strong(valor_clima(valor, dec))]
                estado = [html.Span(className="criterio-color", style={"background": seg["color"]}), seg["texto"]]
            anterior = previa[k] if previa is not None else np.nan
            if es_vacio(valor) or es_vacio(anterior):
                delta = html.Span("Sin mes anterior para comparar", className="delta delta-neutro")
            else:
                cambio = round(float(valor) - float(anterior), dec) + 0.0
                if cambio == 0:
                    delta = html.Span(f"= igual que {previa['etiqueta']}", className="delta delta-neutro")
                else:
                    sube = cambio > 0
                    delta = html.Span([html.Span("▲" if sube else "▼", className="delta-icono"),
                                       f" {con_signo(cambio, dec)} vs {previa['etiqueta']}"],
                                      className="delta " + ("delta-sube" if sube else "delta-baja"))
            salida += [grafico_velocimetro_clima(cfg, valor), valor_html, estado, delta,
                       "clima-tarjeta activa" if k == clave_sel else "clima-tarjeta"]
        return salida

    # Evolución, criterio, variación a 12 meses y hace un año vs. hoy
    @app.callback(Output("clima-graf-evolucion", "figure"), Output("clima-graf-evolucion-titulo", "children"),
                  Output("clima-criterio", "children"), Output("clima-criterio-titulo", "children"),
                  Output("clima-graf-variacion", "figure"), Output("clima-graf-variacion-titulo", "children"),
                  Output("clima-graf-anio", "figure"), Output("clima-graf-anio-titulo", "children"),
                  Input("clima-mes", "value"), Input("clima-indice", "value"))
    def actualizar_evolucion_clima(idx, clave):
        cfg = INDICES_CLIMA[clave]
        fila = clima.loc[idx]
        titulo_var = f"Variación a 12 meses del {clave}"
        if idx >= 12:
            cambio = variacion_12(clima, clave).loc[idx]
            if not es_vacio(cambio):
                titulo_var += (f": {con_signo(round(cambio, cfg['decimales']) + 0.0, cfg['decimales'])} "
                               f"({clima.loc[idx - 12, 'etiqueta']} → {fila['etiqueta']})")
        titulo_anio = (f"Hace un año vs. hoy: {clima.loc[idx - 12, 'etiqueta']} → {fila['etiqueta']}"
                       if idx >= 12 else "Hace un año vs. hoy")
        return (grafico_evolucion_clima(clima, clave, idx), f"Evolución del {clave} ({cfg['nombre']})",
                tabla_criterio(clave), f"Criterio del {clave}",
                grafico_variacion12_clima(clima, clave, idx), titulo_var,
                grafico_anio_vs_hoy(clima, idx), titulo_anio)

    # Estacionalidad, top 10, coincidencia de índices y calendario
    @app.callback(Output("clima-graf-estacional", "figure"), Output("clima-graf-estacional-titulo", "children"),
                  Output("clima-graf-top", "figure"), Output("clima-graf-top-titulo", "children"),
                  Output("clima-graf-consenso", "figure"),
                  Output("clima-graf-calendario", "figure"), Output("clima-graf-calendario-titulo", "children"),
                  Input("clima-mes", "value"), Input("clima-indice", "value"))
    def actualizar_patrones_clima(idx, clave):
        fila = clima.loc[idx]
        serie = clima[clave].dropna()
        titulo_top = f"Los 10 meses más cálidos del {clave}"
        if not es_vacio(fila[clave]):
            puesto = int((serie > fila[clave]).sum()) + 1
            titulo_top += f" · {fila['etiqueta']}: puesto {puesto} de {len(serie)}"
        return (grafico_estacional_clima(clima, clave, idx), f"Cada año, mes a mes: {clave}",
                grafico_top_clima(clima, clave, idx), titulo_top,
                grafico_consenso_clima(clima, idx),
                grafico_calendario_clima(clima, clave, idx), f"Calendario del {clave}")


def registrar_callbacks_latam(app, latam):
    e = latam["eventos"]
    a0, a1 = int(e["anio"].min()), int(e["anio"].max())
    por_iso = dict(zip(e["iso3"], e["pais"]))
    paises = set(e["pais"])
    zonas = dict(zip(e["pais"], e["zona"]))

    # Clic en el ranking o en el mapa, y botón de reinicio -> país y periodo
    @app.callback(Output("latam-pais", "value"), Output("latam-periodo", "value"),
                  Input("latam-graf-ranking", "clickData"), Input("latam-graf-mapa", "clickData"),
                  Input("latam-btn-reiniciar", "n_clicks"), prevent_initial_call=True)
    def sincronizar_latam(clic_ranking, clic_mapa, _n):
        origen = ctx.triggered_id
        if origen == "latam-btn-reiniciar":
            return PAIS_INICIAL, [a0, a1]
        clic = clic_ranking if origen == "latam-graf-ranking" else clic_mapa
        punto = ((clic or {}).get("points") or [{}])[0]
        pais = punto.get("y") if origen == "latam-graf-ranking" else por_iso.get(punto.get("location"))
        return (pais if pais in paises else no_update), no_update

    # Barra lateral, título y tarjetas
    @app.callback([Output("latam-lateral-pais", "children"), Output("latam-lateral-periodo", "children"),
                   Output("latam-titulo", "children")]
                  + [Output(f"latam-kpi-{k}-{parte}", "children") for k, _, _ in KPIS_LATAM for parte in ("valor", "nota")],
                  Input("latam-pais", "value"), Input("latam-periodo", "value"))
    def actualizar_resumen_latam(pais, periodo):
        p0, p1 = periodo
        tabla = resumen_paises(latam, p0, p1)
        fila = tabla.loc[pais]
        n = len(tabla)

        def puesto(col):
            return f"Puesto {int((tabla[col] > fila[col]).sum()) + 1} de {n} países"

        antes = len(e[(e["pais"] == pais) & (e["anio"] <= CORTE_PERIODOS)])
        despues = len(e[(e["pais"] == pais) & (e["anio"] > CORTE_PERIODOS)])
        cambio = variacion_pct(despues, antes)
        nota_eventos = [puesto("eventos")]
        if cambio is not None:
            nota_eventos = [html.Span(puesto("eventos")), html.Br(),
                            html.Span(f"{'▲' if cambio > 0 else '▼'} {numero(abs(cambio), 0)} % en {CORTE_PERIODOS + 1}–{a1} "
                                      f"vs {a0}–{CORTE_PERIODOS}", className="delta-sube" if cambio > 0 else "delta-baja")]
        anio_pbi = int(latam["pbi"].loc[pais, "anio_pbi"])
        salida = [pais, texto_periodo(p0, p1), f"Desastres naturales en {pais}, {texto_periodo(p0, p1)}",
                  numero(fila["eventos"], 0), nota_eventos,
                  numero(fila["fallecidos"], 0), puesto("fallecidos"),
                  numero(fila["afectados"], 0), puesto("afectados"),
                  numero(fila["dano_miles"] / 1000, 0), "Ajustado por inflación · " + puesto("dano_miles"),
                  f"{numero(fila['costo_pbi'], 1)} %", f"Del PBI {anio_pbi} · " + puesto("costo_pbi")]
        return salida

    # Ranking, mapa, dos periodos, dispersión y evolución
    @app.callback(Output("latam-graf-ranking", "figure"), Output("latam-graf-ranking-titulo", "children"),
                  Output("latam-graf-mapa", "figure"), Output("latam-graf-mapa-titulo", "children"),
                  Output("latam-graf-periodos", "figure"), Output("latam-graf-dispersion", "figure"),
                  Output("latam-graf-dispersion-titulo", "children"), Output("latam-graf-evolucion", "figure"),
                  Input("latam-pais", "value"), Input("latam-periodo", "value"), Input("latam-indicador", "value"))
    def actualizar_regional_latam(pais, periodo, ind):
        p0, p1 = periodo
        tabla = resumen_paises(latam, p0, p1)
        titulo = f"{INDICADORES_LATAM[ind]} por país, {texto_periodo(p0, p1)}"
        return (grafico_ranking_latam(tabla, pais, ind), titulo,
                grafico_mapa_latam(tabla, pais, ind), f"Mapa: {INDICADORES_LATAM[ind].lower()}, {texto_periodo(p0, p1)}",
                grafico_periodos_latam(latam, pais),
                grafico_dispersion_latam(latam, p0, p1),
                f"¿Más desastres, más costo? Región, {texto_periodo(p0, min(p1, ULTIMO_ANIO_COSTO))}",
                grafico_evolucion_latam(latam, p0, p1, pais))

    # Tipos de desastre: zona del país y el país elegido
    @app.callback(Output("latam-graf-zona-eventos", "figure"), Output("latam-graf-zona-eventos-titulo", "children"),
                  Output("latam-graf-zona-costo", "figure"), Output("latam-graf-zona-costo-titulo", "children"),
                  Output("latam-graf-pais-eventos", "figure"), Output("latam-graf-pais-eventos-titulo", "children"),
                  Output("latam-graf-pais-costo", "figure"), Output("latam-graf-pais-costo-titulo", "children"),
                  Input("latam-pais", "value"), Input("latam-periodo", "value"))
    def actualizar_tipos_latam(pais, periodo):
        p0, p1 = periodo
        zona = zonas[pais]
        en_zona = {"Sudamérica": "en Sudamérica", "Caribe": "en el Caribe"}.get(zona, "en México y Centroamérica")
        texto = texto_periodo(p0, p1)
        return (grafico_tipos_zona(latam, p0, p1, zona, pais, "eventos"),
                f"Desastres según tipo de evento {en_zona}, {texto} (%)",
                grafico_tipos_zona(latam, p0, p1, zona, pais, "costo"),
                f"Costo económico según tipo de evento {en_zona}, {texto} (%)",
                grafico_tipos_pais(latam, pais, p0, p1, "eventos"),
                f"Desastres en {pais} por tipo y subtipo, {texto}",
                grafico_tipos_pais(latam, pais, p0, p1, "costo"),
                f"Costo de los desastres en {pais} por tipo y subtipo, {texto} (% del PBI)")


def registrar_callbacks(app, base, geojson):
    anios = lista_anios(base)
    centros = centroides(geojson)
    serie_nacional = serie_anual(base, NACIONAL)
    bandas = {col: rango_departamentos(base, col) for col in INDICADORES.values()}
    topes = {col: math.ceil(base[col].max()) for col in INDICADORES.values()}
    departamentos = set(base[COL_DPTO])

    # Clics en gráficos / reproducir / reiniciar -> filtros
    @app.callback(
        Output("filtro-dpto", "value"), Output("filtro-anio", "value"),
        Input("graf-mapa", "clickData"), Input("graf-ranking", "clickData"),
        Input("graf-variacion", "clickData"), Input("graf-matriz", "clickData"),
        Input("graf-anual", "clickData"), Input("graf-tendencia", "clickData"),
        Input("graf-composicion", "clickData"), Input("graf-calor", "clickData"),
        Input("intervalo", "n_intervals"), Input("btn-reiniciar", "n_clicks"),
        State("filtro-dpto", "value"), State("filtro-anio", "value"),
        prevent_initial_call=True)
    def sincronizar_filtros(*args):
        dpto_actual, anio_actual = args[-2], args[-1]
        origen = ctx.triggered_id
        clic = ctx.triggered[0]["value"] if ctx.triggered else None
        puntos = (clic or {}).get("points") if isinstance(clic, dict) else None
        punto = puntos[0] if puntos else {}

        def alternar(dpto):  # segundo clic en el mismo dpto = volver a Perú
            if dpto not in departamentos:
                return no_update
            return NACIONAL if dpto == dpto_actual else dpto

        def anio_valido(valor):
            return int(valor) if valor in anios else no_update

        if origen == "btn-reiniciar":
            return NACIONAL, anios[-1]
        if origen == "intervalo":
            return no_update, (anio_actual + 1 if anio_actual < anios[-1] else anios[0])
        if origen == "graf-mapa":
            return alternar(punto.get("location")), no_update
        if origen in ("graf-ranking", "graf-variacion"):
            return alternar(punto.get("y")), no_update
        if origen in ("graf-matriz", "graf-calor"):
            custom = punto.get("customdata")
            dpto = alternar(custom[0]) if custom else no_update
            anio = anio_valido(punto.get("x")) if origen == "graf-calor" else no_update
            return dpto, anio
        if origen in ("graf-anual", "graf-tendencia", "graf-composicion"):
            return no_update, anio_valido(punto.get("x"))
        return no_update, no_update

    # Botón reproducir / pausar
    @app.callback(Output("intervalo", "disabled"), Output("btn-play", "children"),
                  Input("btn-play", "n_clicks"), State("intervalo", "disabled"),
                  prevent_initial_call=True)
    def alternar_reproduccion(_n, desactivado):
        reproducir = bool(desactivado)
        return (not reproducir), ("❚❚ Pausar" if reproducir else "▶ Reproducir")

    # Barra lateral, título y KPIs
    salidas_kpi = []
    for k in KPIS:
        salidas_kpi += [Output(f"kpi-{k['id']}-valor", "children"),
                        Output(f"kpi-{k['id']}-delta", "children"),
                        Output(f"kpi-{k['id']}-spark", "figure")]

    @app.callback([Output("lateral-anio", "children"), Output("lateral-dpto", "children"),
                   Output("titulo-panorama", "children")] + salidas_kpi,
                  [Input("filtro-dpto", "value"), Input("filtro-anio", "value")])
    def actualizar_resumen(dpto, anio):
        serie = serie_anual(base, dpto)
        lugar = "todo el Perú" if dpto == NACIONAL else nombre_visible(dpto)
        salida = [str(anio), nombre_visible(dpto), f"Riesgo de desastres en {lugar}, {anio}"]
        for k in KPIS:
            col = k["columna"]
            actual, previo = serie[col].get(anio), serie[col].get(anio - 1)
            salida += [numero(actual, k["decimales"]), delta_kpi(actual, previo, anio),
                       grafico_sparkline(serie[col], anio, k["decimales"])]
        return salida

    # Velocímetro, mapa, ranking, barras y línea de tiempo
    @app.callback(Output("graf-velocimetro", "figure"), Output("texto-velocimetro", "children"),
                  Output("graf-velocimetro-titulo", "children"), Output("graf-tendencia-titulo", "children"),
                  Output("graf-anual", "figure"), Output("graf-ranking", "figure"),
                  Output("graf-mapa", "figure"), Output("graf-tendencia", "figure"),
                  Input("filtro-dpto", "value"), Input("filtro-anio", "value"),
                  Input("filtro-indicador", "value"))
    def actualizar_panorama(dpto, anio, col):
        serie = serie_anual(base, dpto)
        corte = corte_anual(base, anio, col)
        r = resumen(base, dpto, anio, col)
        nombre = nombre_indicador(col)
        if es_vacio(r["valor"]):
            texto = ""
        elif r["es_nacional"]:
            texto = f"Promedio de los {r['n_dptos']} departamentos"
        else:
            texto = (f"Puesto {r['puesto']} de {r['n_dptos']} · {NOMBRES_CUARTIL[r['cuartil']]}. "
                     f"Marca azul: promedio nacional ({numero(r['promedio_nacional'], 2)})")
        return (grafico_velocimetro(r, umbrales_cuartil(base, anio, col), topes[col]),
                texto,
                "Índice de Riesgo de Desastres (IRD)" if col == COL_IRD else f"Índice de {nombre}",
                "Evolución del IRD" if col == COL_IRD else f"Evolución de la {nombre}",
                grafico_estratos_anuales(serie, anio, col),
                grafico_ranking(corte, dpto, col),
                grafico_mapa(corte, geojson, centros, dpto, col),
                grafico_tendencia(serie_nacional, bandas[col], anio, col,
                                  None if r["es_nacional"] else serie, nombre_visible(dpto)))

    # Variación, composición, matriz y mapa de calor
    @app.callback(Output("graf-calor", "figure"), Output("graf-matriz", "figure"),
                  Output("graf-variacion", "figure"), Output("graf-composicion", "figure"),
                  Input("filtro-dpto", "value"), Input("filtro-anio", "value"),
                  Input("filtro-indicador", "value"))
    def actualizar_analisis(dpto, anio, col):
        return (grafico_mapa_calor(base, dpto, anio, col),
                grafico_matriz_riesgo(base, dpto, anio, col),
                grafico_variacion(variacion_anual(base, anio, col), dpto, anio, col),
                grafico_composicion(serie_anual(base, dpto), anio))


# =============================================================================
# 7. EJECUTAR
# =============================================================================
BASE = cargar_base()
GEOJSON = cargar_geojson()
CLIMA = cargar_clima()
LATAM = cargar_latam()

app = Dash(__name__, title=f"{TITULO_PORTADA} | Perú", external_stylesheets=[URL_FUENTES],
           suppress_callback_exceptions=True,  # las páginas se dibujan según la dirección
           meta_tags=[{"name": "viewport", "content": "width=device-width, initial-scale=1"}])
app.index_string = PAGINA
server = app.server


@server.route("/" + ARCHIVO_PORTADA)
def imagen_portada():  # la imagen de la portada se sirve desde la carpeta de app.py
    return send_from_directory(CARPETA, ARCHIVO_PORTADA)


app.layout = crear_marco()
paginas = [crear_marco(), crear_portada(BASE, CLIMA, LATAM), crear_layout(BASE)]
if CLIMA is not None:
    paginas.append(crear_layout_clima(CLIMA))
if LATAM is not None:
    paginas.append(crear_layout_latam(LATAM))
app.validation_layout = html.Div(paginas)
registrar_navegacion(app, BASE, CLIMA, LATAM)
registrar_callbacks(app, BASE, GEOJSON)
if CLIMA is not None:
    registrar_callbacks_clima(app, CLIMA)
if LATAM is not None:
    registrar_callbacks_latam(app, LATAM)

if __name__ == "__main__":
    # En JupyterLab: "external" muestra el enlace http://127.0.0.1:8050 (ábrelo en el navegador).
    # Para verlo en tu celular (misma red WiFi): cambia a host="0.0.0.0" y abre en el celular
    # http://IP-DE-TU-PC:8050  (la IP la ves con el comando ipconfig en Windows).
    app.run(jupyter_mode="external", host="127.0.0.1", port=8050, debug=True)
