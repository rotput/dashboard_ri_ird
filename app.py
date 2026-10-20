# =============================================================================
#  SEGUIMIENTO DE ÍNDICES DE DESASTRES - PERÚ
#
#  Páginas:
#    /                     Portada (4 botones)
#    /indices-mensuales    Índices climáticos mensuales      <- ICEM.xlsx
#    /indice-desastres     Índice de Desastres (ID)          <- IRD_RESULTADO.xlsx
#    /inform-risk          Índice INFORM Risk                <- inform.xlsx
#    /latam                Estadísticas de América Latina    <- EMDAT_LATAM.xlsx
#    /estimacion-costos    Estimación de costos (en preparación)
#  Archivos junto a app.py: ICEM.xlsx, IRD_RESULTADO.xlsx, inform.xlsx, EMDAT_LATAM.xlsx y portada.jpg
#
#  Pega todo en UNA celda de JupyterLab y ejecútala. Se abre en: http://127.0.0.1:8050
#  Librerías (una sola vez, en otra celda):
#      !pip install -U dash plotly pandas openpyxl
#
#  Secciones:
#    1. CONFIGURACIÓN   -> archivos, colores, textos y definiciones (cambia aquí)
#         1.1 General   1.2 Portada   1.3 Índices climáticos   1.4 Índice de Desastres (ID)
#         1.5 INFORM Risk   1.6 América Latina
#    2. FORMATO         -> números y textos
#    3. DATOS           -> lectura de los Excel y cálculos (mismo orden que la sección 1)
#    4. GRÁFICOS        -> una función por gráfico (mismo orden)
#    5. DISEÑO          -> íconos, piezas comunes, portada, una función por página y estilos CSS
#    6. INTERACTIVIDAD  -> navegación y callbacks de cada página
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
# ---- 1.1 General ----------------------------------------------------------------
# Los Excel están en la misma carpeta que app.py (raíz del repositorio de GitHub)
CARPETA = os.path.dirname(os.path.abspath(__file__)) if "__file__" in globals() else os.getcwd()

# Colores de la página. Si cambias uno, cámbialo también en CSS (sección 5).
C = {
    "pacifico": "#0b3a66", "pacifico_medio": "#1f6fb2", "niebla": "#edf3f9",
    "superficie": "#ffffff", "borde": "#d5e2ef", "tinta": "#0f2f50",
    "tinta_2": "#3f6488", "tinta_3": "#6f8cab", "rejilla": "#e6eef6",
    "contexto": "#a7bcd2", "mejora": "#2a78d6", "empeora": "#e34948", "azul_suave": "#e3eefa",
}
COLOR_ROJO_PPT, COLOR_AZUL_PPT = "#d62f2f", "#0b2a5b"     # rojo y azul de la presentación
# Cuartiles (colores de la presentación): 1 = valores más bajos ... 4 = valores más altos
COLORES_CUARTIL = {1: "#F2F089", 2: "#FD8D3C", 3: "#E31A1C", 4: "#800026"}
BORDE_MAPA = "#696969"

# Cómo se arman los cuartiles de los mapas y rankings (para que coincidan con la presentación):
#   "excel_inc" = CUARTIL.INC de Excel   "excel_exc" = CUARTIL.EXC de Excel
#   "grupos"    = grupos de igual tamaño contados desde el valor más alto
FUENTE = "Archivo"
URL_FUENTES = "https://fonts.googleapis.com/css2?family=Archivo:wdth,wght@62..125,300..800&display=swap"
SEPARADOR_DECIMAL = ","
SEPARADOR_MILES = "."

# Mapa del Perú: si no está en la carpeta, se descarga y se guarda ahí
URL_GEOJSON = ("https://raw.githubusercontent.com/juaneladio/peru-geojson/"
               "master/peru_departamental_simple.geojson")
RUTA_GEOJSON = os.path.join(CARPETA, "peru_departamental.geojson")
CLAVE_GEOJSON = "properties.NOMBDEP"

# ---- 1.2 Portada ------------------------------------------------------------------
# "lista": True = el tablero ya funciona; False = muestra "En preparación".
TITULO_PORTADA = "Seguimiento de Índices de Desastres"
ANTETITULO_PORTADA = "Perú · Gestión del riesgo de desastres"
SUBTITULO_PORTADA = ("Tableros para seguir las amenazas climáticas, los índices de desastres, "
                     "la situación de América Latina y el costo de los desastres.")
ARCHIVO_PORTADA = "portada.jpg"      # imagen de fondo de la portada (junto a app.py)
RUTA_CLIMA = "/indices-mensuales"
RUTA_ID = "/indice-desastres"
RUTA_INFORM = "/inform-risk"
RUTA_LATAM = "/latam"
RUTA_COSTOS = "/estimacion-costos"
SECCIONES = [
    {"ruta": RUTA_CLIMA, "icono": "ola", "lista": True,
     "titulo": "Índices climáticos mensuales",
     "descripcion": "Seguimiento mes a mes de los índices que miden si el mar está en condición "
                    "cálida, neutra o fría.",
     "etiquetas": ["ICEN", "ITCP", "LABCOS", "EHNOS"]},
    {"ruta": RUTA_ID, "icono": "medidor", "lista": True,
     "titulo": "Índices de desastres",
     "descripcion": "Índice de Desastres (ID) de los 25 departamentos del Perú e índice INFORM Risk "
                    "de América Latina y el mundo.",
     "etiquetas": ["ID", "INFORM Risk"]},
    {"ruta": RUTA_LATAM, "icono": "globo", "lista": True,
     "titulo": "Estadísticas de América Latina",
     "descripcion": "Frecuencia, costo económico y tipo de los desastres naturales en 26 países "
                    "de la región.",
     "etiquetas": ["EM-DAT", "Frecuencia", "Costo % PBI"]},
    {"ruta": RUTA_COSTOS, "icono": "moneda", "lista": False,
     "titulo": "Estimación de costos",
     "descripcion": "Estimación del costo fiscal de los desastres naturales.",
     "etiquetas": ["Costos", "Pasivos contingentes"]},
]

# ---- 1.3 Índices climáticos mensuales (archivo ICEM.xlsx) -----------------------
ARCHIVO_CLIMA = os.path.join(CARPETA, "ICEM.xlsx")
HOJA_CLIMA = 0                       # 0 = primera hoja del Excel
COL_CLIMA_ANIO, COL_CLIMA_MES = "yyyy", "mm"
MESES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto",
         "Septiembre", "Octubre", "Noviembre", "Diciembre"]
COLORES_CALIDOS = ["#f03b20", "#fd8d3c", "#fecc5c", "#ffffb2"]   # de más fuerte a más débil
COLORES_FRIOS = ["#bdd7e7", "#6baed6", "#3182bd"]                # de más débil a más fuerte
COLOR_NEUTRO = "#ffffff"
COLOR_CALIDO, COLOR_FRIO = "#f03b20", "#3182bd"
# Cada índice: columna del Excel, nombre, decimales, unidad, definición y criterio (tus tablas de "CRITERIO").
#   El ORDEN de esta lista es el orden de los filtros y de las tarjetas.
#   unidad: se muestra en los ejes, las tarjetas y el criterio ("" = sin unidad)
#   calidas: (magnitud, umbral) de la más fuerte a la más débil -> cálida si valor > umbral
#   frias:   (magnitud, umbral) de la más débil a la más fuerte -> fría si valor < umbral
#   incluye_limite=True cuando el criterio usa ≥ y ≤ (como el MEI v2)
INDICES_CLIMA = {
    "ICEN": {"columna": "ICE", "nombre": "Índice Costero El Niño", "decimales": 2, "unidad": "°C",
             "definicion": ("Mide las anomalías de temperatura superficial del mar en la región Niño 1+2, situada "
                            "frente a Ecuador y el norte del Perú. Mediante un promedio de tres meses, permite al "
                            "ENFEN diagnosticar la presencia y magnitud de las condiciones de El Niño y La Niña "
                            "costeras."),
             "calidas": [("Extraordinaria", 3.5), ("Fuerte", 2.1), ("Moderada", 1.3), ("Débil", 0.5)],
             "frias": [("Débil", -0.7), ("Moderada", -1.1), ("Fuerte", -1.3)]},
    "ITCP": {"columna": "ITCP", "nombre": "Índice Térmico Costero Peruano", "decimales": 1, "unidad": "°C",
             "definicion": ("Evalúa las variaciones de temperatura del mar en el ecosistema de afloramiento costero "
                            "peruano, una zona marina donde ascienden aguas profundas, frías y ricas en nutrientes. "
                            "Permite caracterizar el estado térmico de este ecosistema, fundamental para la "
                            "productividad marina y pesquera, pero no mide directamente la abundancia de peces ni la "
                            "biodiversidad."),
             "calidas": [("Cálida", 0.4)], "frias": [("Fría", -0.6)]},
    "LABCOS": {"columna": "LABCOS", "nombre": "Índice de Laboratorios Costeros", "decimales": 1,
               "decimales_criterio": 2, "unidad": "°C",
               "definicion": ("Mide las variaciones de la temperatura superficial del mar en la franja costera del "
                              "Perú, mediante registros de estaciones distribuidas a lo largo del litoral, desde Paita "
                              "hasta Ilo. Integra estas mediciones para identificar si las aguas próximas a la costa "
                              "peruana presentan condiciones más cálidas o más frías de lo habitual."),
               "calidas": [("Muy fuerte", 2.77), ("Fuerte", 1.37), ("Moderada", 0.78), ("Débil", 0.27)],
               "frias": [("Débil", -0.78), ("Moderada", -1.10), ("Fuerte", -1.31)]},
    # EHNOS: se usa el criterio del MEI v2 (±0,5). Si su criterio es otro, cámbialo aquí.
    "EHNOS": {"columna": "EHNOS", "nombre": "Índice EHNOS", "decimales": 2, "unidad": "°C",
              "definicion": ("Índice asociado al EHNOS (El Niño-Oscilación del Sur), un fenómeno climático que combina "
                             "variaciones del océano y la atmósfera en el Pacífico tropical y cuyos efectos pueden "
                             "extenderse a distintas regiones del mundo."),
              "calidas": [("Cálida", 0.5)], "frias": [("Fría", -0.5)], "incluye_limite": True},
}
FUENTE_CLIMA = "Fuente: IMARPE (SIOFEN). Los valores marcados con * en el Excel se muestran como preliminares."
FUENTE_CLIMA_CORTA = "Fuente: IMARPE (SIOFEN)."   # la que va debajo de cada gráfico

# ---- 1.4 Índice de Desastres (archivo IRD_RESULTADO.xlsx, hoja IRD_ANUAL) --------
ARCHIVO_ID = os.path.join(CARPETA, "IRD_RESULTADO.xlsx")
HOJA_ID = "IRD_ANUAL"
COL_ANIO, COL_DPTO = "ANIO", "DPTO"
COL_ID, COL_FREC, COL_SEV, COL_INCIDENCIAS = "IRD", "I_FREC", "I_DANO", "NUM_INCIDENCIAS"
# Indicadores del filtro (columna -> textos). Al cambiarlo cambian ranking, mapa, evolución y mapa de calor.
INDICADORES_ID = {
    COL_ID: {"nombre": "Índice de Desastres (ID)", "corto": "ID", "tope": 10, "decimales": 1,
             "unidad": "(Puntaje de 0 a 10)"},
    COL_FREC: {"nombre": "Índice de frecuencia", "corto": "Frecuencia", "tope": 1, "decimales": 2,
               "unidad": "(Índice de 0 a 1)"},
    COL_SEV: {"nombre": "Severidad global", "corto": "Severidad global", "tope": 1, "decimales": 2,
              "unidad": "(Índice de 0 a 1)"},
}
# Los 4 pilares de la severidad (25 % cada uno), como en la ficha metodológica
PILARES = {
    "Humana": {"columna": "S_HUMANA", "tarjeta": "Pilar humano", "color": "#D7262B", "icono": "persona",
               "detalle": "Fallecidos, desaparecidos y lesionados"},
    "Social": {"columna": "S_SOCIAL", "tarjeta": "Pilar social", "color": "#B5541C", "icono": "casa",
               "detalle": "Damnificados, afectados y viviendas"},
    "Infraestructura": {"columna": "S_INFRAESTRUCTURA", "tarjeta": "Pilar infraestructura", "color": "#2160C4", "icono": "puente",
                        "detalle": "Educación, salud, vías y canales de riego"},
    "Productiva": {"columna": "S_PRODUCTIVA", "tarjeta": "Pilar productivo", "color": "#2E7D32", "icono": "brote",
                   "detalle": "Cultivos y ganado"},
}
DEFINICION_ID = ("Mide la severidad de los daños ocasionados por desastres naturales en un departamento durante "
                 "un periodo determinado. ID = 10 × Frecuencia × Severidad global (pilares humano, social, "
                 "infraestructura y productivo, 25 % cada uno). Escala de 0 (afectación mínima) a 10 "
                 "(afectación muy alta).")
FUENTE_ID = "Fuente: INDECI. Elaboración: DRF-VMH."
METODO_CUARTIL_ID = "excel_inc"
NACIONAL = "PERÚ (promedio)"
NOMBRES_VISIBLES = {"ANCASH": "Áncash", "APURIMAC": "Apurímac", "HUANUCO": "Huánuco",
                    "JUNIN": "Junín", "SAN MARTIN": "San Martín", "MADRE DE DIOS": "Madre de Dios"}
EVENTOS = {2017: "El Niño Costero", 2023: "Ciclón Yaku"}   # anotaciones en la evolución
MS_POR_ANIO = 1400  # velocidad del botón "Reproducir"

# ---- 1.5 INFORM Risk (archivo inform.xlsx) -----------------------------------------
ARCHIVO_INFORM = os.path.join(CARPETA, "inform.xlsx")
HOJA_INFORM = 0
ANIO_INFORM = 2026
PAIS_INICIAL_INFORM = "PER"
DEFINICION_INFORM = ("Mide, en una escala de 0 a 10, el riesgo de que ocurran crisis humanitarias y desastres que "
                     "superen la capacidad de respuesta nacional. INFORM Risk = (Amenaza y exposición)^1/3 × "
                     "(Vulnerabilidad)^1/3 × (Falta de capacidad de respuesta)^1/3. A mayor valor, mayor riesgo.")
FUENTE_INFORM = "Fuente: Comisión Europea (2026), INFORM Risk. Elaboración: DRF-VMH."
# Indicadores del filtro (columna del Excel -> nombre). "Amenazas naturales" = riesgo de desastres naturales.
INDICADORES_INFORM = {
    "INFORM RISK": "INFORM Risk",
    "HAZARD & EXPOSURE": "Amenaza y exposición",
    "Natural": "Amenazas naturales",
    "VULNERABILITY": "Vulnerabilidad",
    "LACK OF COPING CAPACITY": "Falta de capacidad de respuesta",
}
DIMENSIONES_INFORM = {   # las 3 dimensiones de la fórmula (Gráfico 5)
    "HAZARD & EXPOSURE": ("Amenaza y exposición", "#D7262B"),
    "VULNERABILITY": ("Vulnerabilidad", "#1F4E9A"),
    "LACK OF COPING CAPACITY": ("Falta de capacidad de respuesta", "#C0661C"),
}
AMENAZAS_NATURALES = {"Earthquake": "Terremoto", "River Flood": "Inundación fluvial", "Tsunami": "Tsunami",
                      "Tropical Cyclone": "Ciclón tropical", "Coastal flood": "Inundación costera",
                      "Drought": "Sequía", "Epidemic": "Epidemia"}
SUBCOMPONENTES_INFORM = {   # Gráfico 6: (dimensión, nombre)
    "Development & Deprivation": ("Vulnerabilidad", "Desarrollo y privaciones"),
    "Inequality": ("Vulnerabilidad", "Desigualdad"),
    "Economic Dependency": ("Vulnerabilidad", "Dependencia económica"),
    "Uprooted people": ("Vulnerabilidad", "Población desplazada"),
    "Health Conditions": ("Vulnerabilidad", "Condiciones de salud"),
    "Food Security": ("Vulnerabilidad", "Inseguridad alimentaria"),
    "DRR": ("Falta de capacidad", "Gestión del riesgo de desastres"),
    "Governance": ("Falta de capacidad", "Gobernanza"),
    "Communication": ("Falta de capacidad", "Comunicación"),
    "Physical infrastructure": ("Falta de capacidad", "Infraestructura física"),
    "Access to health care": ("Falta de capacidad", "Acceso a la salud"),
}
# Clases del INFORM (umbrales de la clase de riesgo del Excel) y sus colores
CLASES_INFORM = [  # (nombre, desde, color)
    ("Muy bajo", 0.0, "#FBF7CF"), ("Bajo", 2.35, "#F2F089"), ("Medio", 3.55, "#FD8D3C"),
    ("Alto", 5.35, "#E31A1C"), ("Muy alto", 6.75, "#800026")]
# Los 21 países de América Latina y el Caribe del Gráfico N.º 10 de la presentación
PAISES_INFORM = {
    "HTI": "Haití", "VEN": "Venezuela", "GTM": "Guatemala", "ECU": "Ecuador", "HND": "Honduras",
    "COL": "Colombia", "PER": "Perú", "MEX": "México", "BOL": "Bolivia", "BRA": "Brasil",
    "SLV": "El Salvador", "DOM": "Rep. Dominicana", "NIC": "Nicaragua", "PAN": "Panamá",
    "CRI": "Costa Rica", "GUY": "Guyana", "ARG": "Argentina", "CUB": "Cuba", "PRY": "Paraguay",
    "CHL": "Chile", "URY": "Uruguay"}
# Solo en el mapa (casi todos los países de la región)
PAISES_INFORM_MAPA = {"JAM": "Jamaica", "BLZ": "Belice", "SUR": "Surinam", "TTO": "Trinidad y Tobago",
                      "BHS": "Bahamas", "BRB": "Barbados", "ATG": "Antigua y Barbuda", "DMA": "Dominica",
                      "GRD": "Granada", "KNA": "San Cristóbal y Nieves", "LCA": "Santa Lucía",
                      "VCT": "San Vicente y las Granadinas"}
GRUPOS_INFORM = {  # Gráfico 3 (como el Gráfico N.º 11 de la presentación)
    "Economías avanzadas": {"JPN": "Japón", "GBR": "Reino Unido", "DEU": "Alemania", "CAN": "Canadá",
                            "USA": "EE. UU."},
    "Economías emergentes": {"CHN": "China", "IDN": "Indonesia", "IND": "India", "TUR": "Turquía",
                             "ZAF": "Sudáfrica"},
    "América Latina": {"CHL": "Chile", "BRA": "Brasil", "MEX": "México", "PER": "Perú", "COL": "Colombia"},
}
COLORES_GRUPO_INFORM = {"Economías avanzadas": "#A9B6CC", "Economías emergentes": "#7193C9",
                        "América Latina": "#D71E25"}

# ---- 1.6 Estadísticas de América Latina (archivo EMDAT_LATAM.xlsx, ya procesado) ---
ARCHIVO_LATAM = os.path.join(CARPETA, "EMDAT_LATAM.xlsx")
PAIS_INICIAL = "Perú"
ULTIMO_ANIO_COSTO = 2025     # EM-DAT todavía no tiene costos de 2026
CORTE_PERIODOS = 2002        # Gráfico de dos periodos: 1980-2002 y 2003-2026
FUENTE_LATAM = "Fuente: EM-DAT. Elaboración: DRF-VMH."
ARTICULOS = {"Perú": "el Perú"}  # para los títulos: «en el Perú»
# Los 2 filtros del tablero. Al cambiarlo cambian tarjetas, títulos y los 6 gráficos.
INDICADORES_LATAM = {"eventos": "Número de desastres", "costo_pbi": "Costo económico (% del PBI)"}
# Cuartiles como en los mapas de la presentación (Mapa N.º 1: grupos; Mapa N.º 2: CUARTIL.EXC)
METODO_CUARTIL_LATAM = {"eventos": "grupos", "costo_pbi": "excel_exc"}
TEXTOS_LATAM = {   # términos de la presentación
    "eventos": {"indicador": "Indicador 1: Frecuencia de los desastres",
                "que_mide": "Número de desastres naturales registrados en cada país o territorio."},
    "costo_pbi": {"indicador": "Indicador 2: Costo económico de los desastres",
                  "que_mide": "Monto de los daños económicos ocasionados por los desastres, ajustado por "
                              "inflación, como porcentaje del PBI."},
}
COLORES_SUBGRUPO = {"Biológico": "#156082", "Climatológico": "#7F7F7F", "Geofísico": "#F2AA84",   # Gráfico 4
                    "Hidrológico": "#D9D9D9", "Meteorológico": "#94DCF8"}
COLORES_TIPO = {"Hidrológico": "#2160C4", "Geofísico": "#D7262B", "Meteorológico": "#3BA9B9",     # Tarjetas y
                "Biológico": "#8C93A3", "Climatológico": "#F0A04B"}                                # Gráficos 5 y 6
ORDEN_SUBGRUPOS_PAIS = ["Hidrológico", "Geofísico", "Meteorológico", "Biológico", "Climatológico"]
ICONOS_TIPO = {"Hidrológico": "ola", "Geofísico": "sismo", "Meteorológico": "nube",
               "Biológico": "virus", "Climatológico": "sol"}
NOMBRES_CORTOS = {"Movimiento de masa (húmedo)": "Mov. de masa (húmedo)",
                  "Movimiento de masa (seco)": "Mov. de masa (seco)"}
COLOR_PERIODO_1, COLOR_PERIODO_2 = "#7F7F7F", "#C00000"
COLOR_DISPERSION, COLOR_TENDENCIA = "#0B2A5B", "#C00000"


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


def hex_a_rgba(color_hex, alfa):
    color_hex = color_hex.lstrip("#")
    r, g, b = (int(color_hex[i:i + 2], 16) for i in (0, 2, 4))
    return f"rgba({r},{g},{b},{alfa})"


def nombre_visible(dpto):
    if dpto == NACIONAL:
        return "Perú"
    return NOMBRES_VISIBLES.get(dpto, dpto.title().replace(" De ", " de "))


def clases_cuartil(serie, metodo="excel_inc"):
    """Clase de 1 (valores más bajos) a 4 (más altos) de cada valor y los 3 cortes.
    metodo: "excel_inc" (CUARTIL.INC), "excel_exc" (CUARTIL.EXC) o "grupos" (grupos iguales desde arriba)."""
    s = pd.Series(serie, dtype=float)
    v = np.sort(s.dropna().to_numpy())
    n = len(v)
    if n == 0:
        return pd.Series(1, index=s.index), [np.nan] * 3
    if metodo == "grupos":
        tamanio = math.ceil(n / 4)
        orden = s.rank(ascending=False, method="first")
        clase = (4 - (orden - 1) // tamanio).clip(1, 4).astype(int)
        cortes = [s[clase >= q].min() for q in (2, 3, 4)]
        return clase, cortes
    if metodo == "excel_exc":
        posiciones = [p * (n + 1) for p in (0.25, 0.5, 0.75)]
    else:
        posiciones = [1 + p * (n - 1) for p in (0.25, 0.5, 0.75)]
    cortes = []
    for pos in posiciones:
        pos = min(max(pos, 1), n)
        i = int(math.floor(pos))
        cortes.append(v[i - 1] + ((pos - i) * (v[i] - v[i - 1]) if i < n else 0.0))
    clase = 1 + sum((s >= c - 1e-12).astype(int) for c in cortes)
    return clase.astype(int), cortes


def textos_clases(valores, clases, formato, tope=None):
    """Leyenda de cada clase: «4,0 – 10,0», «2,9 – 4,0» ... (mínimo y máximo de los valores de la clase).
    formato: función que convierte un número en texto. tope=None -> la clase 4 dice «… a más»."""
    textos = {}
    for q in COLORES_CUARTIL:
        sub = pd.Series(valores)[pd.Series(clases).to_numpy() == q]
        if sub.empty:
            textos[q] = "—"
        elif q == 4:
            textos[q] = f"{formato(sub.min())} a más" if tope is None else f"{formato(sub.min())} – {formato(tope)}"
        else:
            textos[q] = f"{formato(sub.min())} – {formato(sub.max())}"
    return textos


# =============================================================================
# 3. DATOS
# =============================================================================
# ---- 3.1 Índices climáticos ---------------------------------------------------
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


def valor_con_unidad(valor, cfg):
    """0.35 -> '0,35 °C' (la unidad sale de INDICES_CLIMA)."""
    texto = valor_clima(valor, cfg["decimales"])
    return texto if texto == "—" or not cfg.get("unidad") else f"{texto} {cfg['unidad']}"


def intervalo_texto(seg, cfg):
    incluye, dec = cfg.get("incluye_limite", False), cfg.get("decimales_criterio", 1)
    if seg["grupo"] == "Condición cálida" and seg["abierto"]:
        return f"{'≥' if incluye else '>'} {umbral_texto(seg['desde'], dec)}"
    if seg["grupo"] == "Condición fría" and seg["abierto"]:
        return f"{'≤' if incluye else '<'} {umbral_texto(seg['hasta'], dec)}"
    return f"{umbral_texto(seg['desde'], dec)} a {umbral_texto(seg['hasta'], dec)}"


# ---- 3.2 Índice de Desastres (ID) ----------------------------------------------
def col_cuartil(col):
    return f"CUARTIL_{col}"


def col_puesto(col):
    return f"PUESTO_{col}"


def cargar_id():
    """Una fila por departamento y año, con el cuartil y el puesto de cada indicador en su año. Si no está, None."""
    if not os.path.exists(ARCHIVO_ID):
        return None
    df = pd.read_excel(ARCHIVO_ID, sheet_name=HOJA_ID)
    df.columns = df.columns.astype(str).str.strip()
    df[COL_DPTO] = df[COL_DPTO].astype(str).str.strip().str.upper()
    df[COL_ANIO] = pd.to_numeric(df[COL_ANIO], errors="coerce")
    df = df.dropna(subset=[COL_ANIO]).astype({COL_ANIO: int})
    for col in [COL_ID, COL_FREC, COL_SEV, COL_INCIDENCIAS] + [p["columna"] for p in PILARES.values()]:
        df[col] = pd.to_numeric(df[col], errors="coerce")
    for col in INDICADORES_ID:
        df[col_cuartil(col)] = 1
        for _, filas in df.groupby(COL_ANIO).groups.items():
            df.loc[filas, col_cuartil(col)] = clases_cuartil(df.loc[filas, col], METODO_CUARTIL_ID)[0]
        df[col_puesto(col)] = df.groupby(COL_ANIO)[col].rank(ascending=False, method="min")
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


def lista_anios(df):
    return sorted(int(a) for a in df[COL_ANIO].unique())


def serie_id(df, dpto):
    """Una fila por año. Perú: promedio de los departamentos (incidencias: suma)."""
    if dpto == NACIONAL:
        indices = [COL_ID, COL_FREC, COL_SEV] + [p["columna"] for p in PILARES.values()]
        serie = df.groupby(COL_ANIO).agg({**{c: "mean" for c in indices}, COL_INCIDENCIAS: "sum"})
    else:
        serie = df.loc[df[COL_DPTO] == dpto].set_index(COL_ANIO)
    return serie.reindex(lista_anios(df))


def corte_id(df, anio, col):
    """Los departamentos en un año, del valor más alto al más bajo."""
    return df.loc[df[COL_ANIO] == anio].sort_values(col, ascending=False).reset_index(drop=True)


def formato_id(col):
    dec = INDICADORES_ID[col]["decimales"]
    return lambda v: numero(v, dec)


# ---- 3.3 INFORM Risk -------------------------------------------------------------
def cargar_inform():
    """Una fila por país (191). La segunda fila del Excel (las escalas) se salta. Si no está, None."""
    if not os.path.exists(ARCHIVO_INFORM):
        return None
    df = pd.read_excel(ARCHIVO_INFORM, sheet_name=HOJA_INFORM, skiprows=[1])
    df.columns = df.columns.astype(str).str.strip()
    columnas = (list(INDICADORES_INFORM) + list(AMENAZAS_NATURALES) + list(SUBCOMPONENTES_INFORM)
                + ["Rank", "Human"])
    for col in columnas:
        df[col] = pd.to_numeric(df[col], errors="coerce")
    df["ISO3"] = df["ISO3"].astype(str).str.strip()
    nombres = {**PAISES_INFORM, **PAISES_INFORM_MAPA,
               **{iso: n for grupo in GRUPOS_INFORM.values() for iso, n in grupo.items()}}
    df["NOMBRE"] = df["ISO3"].map(nombres).fillna(df["COUNTRY"])
    return df.set_index("ISO3")


def clase_inform(valor):
    """(nombre, color) de la clase de riesgo de un puntaje de 0 a 10."""
    if es_vacio(valor):
        return "Sin dato", C["contexto"]
    nombre, color = CLASES_INFORM[0][0], CLASES_INFORM[0][2]
    for n, desde, c in CLASES_INFORM:
        if valor >= desde:
            nombre, color = n, c
    return nombre, color


def tabla_alc(inform):
    """Los 21 países de América Latina y el Caribe del ranking."""
    return inform.loc[[iso for iso in PAISES_INFORM if iso in inform.index]]


# ---- 3.4 América Latina (EM-DAT) ---------------------------------------------------
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


def valor_latam(valor, indicador):
    return f"{numero(valor, 1)} %" if indicador == "costo_pbi" else numero(valor, 0)


def texto_periodo(a0, a1):
    return f"{a0}–{a1}" if a0 != a1 else str(a0)


def nombre_corto(tipo):
    return NOMBRES_CORTOS.get(tipo, tipo)


def porcentaje_corto(v):
    """Etiquetas de % del PBI: 0 -> '' ; 0,03 -> '<0,1 %'."""
    if v <= 0:
        return ""
    return "<0,1 %" if v < 0.05 else f"{numero(v, 1)} %"




# =============================================================================
# 4. GRÁFICOS  (una función por gráfico; todas devuelven una figura de Plotly)
# =============================================================================
# ---- 4.0 Estilo común de todos los gráficos ----------------------------------------
FUENTE_PLOTLY = f"{FUENTE}, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif"

pio.templates["tablero"] = go.layout.Template(layout=dict(
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
    colorway=[p["color"] for p in PILARES.values()], bargap=0.22, barcornerradius=3,
))
pio.templates.default = "tablero"

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


def colores_clase(clases):
    return [COLORES_CUARTIL.get(int(q), C["contexto"]) if not es_vacio(q) else C["contexto"] for q in clases]


def contorno(seleccion, ancho=2.5):
    """Borde azul oscuro solo en la barra elegida."""
    return dict(color=[COLOR_AZUL_PPT if s else "rgba(0,0,0,0)" for s in seleccion],
                width=[ancho if s else 0 for s in seleccion])


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


# ---- 4.1 Índices climáticos ---------------------------------------------------------
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


def limites_neutro(cfg):
    return cfg["frias"][0][1], cfg["calidas"][-1][1]


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


def grafico_evolucion_clima(clima, clave, idx):
    cfg = INDICES_CLIMA[clave]
    dec = cfg["decimales"]
    lim_frio, lim_calido = limites_neutro(cfg)
    unidad = cfg.get("unidad", "")
    sufijo = f" {unidad}" if unidad else ""
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
    custom = [[e, valor_con_unidad(x, cfg), segs[clasificar_clima(x, cfg)]["texto"], " (preliminar)" if p else ""]
              for e, x, p in zip(datos["etiqueta"], v, datos[clave + "_prelim"])]
    fig.add_scatter(x=datos["fecha"], y=v, mode="lines", line=dict(color=C["tinta"], width=1.4),
                    customdata=custom, showlegend=False,
                    hovertemplate="<b>%{customdata[0]}</b>: %{customdata[1]}%{customdata[3]}"
                                  "<br>%{customdata[2]}<extra></extra>")
    for y, texto, color, ancla in [(lim_calido, f"Umbral cálido {umbral_texto(lim_calido, cfg.get('decimales_criterio', 1))} {unidad}",
                                    COLOR_CALIDO, "bottom"),
                                   (lim_frio, f"Umbral frío {umbral_texto(lim_frio, cfg.get('decimales_criterio', 1))} {unidad}",
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
                           text=f"{valor_con_unidad(serie[i], cfg)} · {datos.loc[i, 'etiqueta']}", showarrow=True,
                           arrowhead=0, arrowwidth=1, arrowcolor=C["tinta_3"],
                           font=dict(size=10, color=C["tinta_2"]), bgcolor=hex_a_rgba(C["superficie"], 0.8))
    fila = clima.loc[idx]
    if not es_vacio(fila[clave]):
        fig.add_scatter(x=[fila["fecha"]], y=[fila[clave]], mode="markers", hoverinfo="skip", showlegend=False,
                        marker=dict(size=14, color=C["superficie"], line=dict(color=C["tinta"], width=3)))
        cerca_del_final = fila["fecha"] > datos["fecha"].max() - pd.DateOffset(months=18)
        fig.add_annotation(x=fila["fecha"], y=fila[clave], ax=-60 if cerca_del_final else 0,
                           ay=-34 if fila[clave] >= 0 else 34,
                           text=f"<b>{valor_con_unidad(fila[clave], cfg)} · {fila['etiqueta']}</b>", showarrow=True,
                           arrowhead=0, arrowwidth=1.4, arrowcolor=C["tinta"],
                           font=dict(size=12, color=C["superficie"]), bgcolor=C["tinta"], borderpad=4)
    fig.add_vline(x=fila["fecha"], line=dict(color=C["pacifico_medio"], width=1, dash="dot"), layer="below")
    bajo, alto = min(np.nanmin(v), lim_frio), max(np.nanmax(v), lim_calido)
    pad = (alto - bajo) * 0.1
    fig.update_layout(
        hovermode="closest", margin=dict(l=4, r=4, t=36, b=4),
        yaxis=dict(title=clave, range=[bajo - pad, alto + pad * 1.8], zeroline=False, ticksuffix=sufijo),
        xaxis=dict(showgrid=False, tickformat="%Y", dtick="M24",
                   range=[datos["fecha"].min() - pd.DateOffset(months=2), datos["fecha"].max() + pd.DateOffset(months=4)],
                   rangeselector=dict(
            x=0, y=1.02, yanchor="bottom", bgcolor=C["niebla"], activecolor=C["azul_suave"],
            font=dict(size=11, color=C["tinta"]),
            buttons=[dict(count=2, label="2 años", step="year", stepmode="backward"),
                     dict(count=5, label="5 años", step="year", stepmode="backward"),
                     dict(step="all", label="Todo")])))
    return fig


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
        custom = [[anio, MESES[m - 1], valor_con_unidad(x, cfg)] for m, x in fila.items()]
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
                      yaxis=dict(title=clave, zeroline=True, zerolinecolor=C["contexto"],
                                 ticksuffix=f" {cfg['unidad']}" if cfg.get("unidad") else ""))
    return fig


def grafico_calendario_clima(clima, clave, idx):
    cfg = INDICES_CLIMA[clave]
    segs, dec = segmentos_clima(cfg), cfg["decimales"]
    anios = sorted(clima["anio"].unique(), reverse=True)
    tabla = clima.pivot_table(index="anio", columns="mes", values=clave, aggfunc="first")
    tabla = tabla.reindex(index=anios, columns=range(1, 13))
    z = [[clasificar_clima(v, cfg) for v in fila] for fila in tabla.to_numpy()]
    texto = [[valor_clima(v, 1) if not es_vacio(v) else "" for v in fila] for fila in tabla.to_numpy()]
    custom = [[[f"{MESES[m - 1]} {a}", valor_con_unidad(v, cfg),
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




# ---- 4.2 Índice de Desastres (ID) ---------------------------------------------------
def grafico_velocimetro_id(valor, previo, cortes, referencia=None):
    """Tarjeta principal: ID de 0 a 10, con los colores de los cuartiles del año."""
    if es_vacio(valor):
        return figura_vacia("Sin dato")
    tope = INDICADORES_ID[COL_ID]["tope"]
    limites = [0] + [float(c) for c in cortes] + [tope]
    gauge = dict(shape="angular", bgcolor=C["rejilla"], borderwidth=0,
                 axis=dict(range=[0, tope], tickvals=list(range(0, tope + 1, 2)), tickwidth=1,
                           tickcolor=C["borde"], tickfont=dict(size=10, color=C["tinta_3"])),
                 bar=dict(color=C["tinta"], thickness=0.22),
                 steps=[dict(range=[limites[i], limites[i + 1]], color=COLORES_CUARTIL[i + 1]) for i in range(4)])
    if referencia is not None:  # marca azul = promedio del Perú
        gauge["threshold"] = dict(line=dict(color=C["pacifico_medio"], width=3), thickness=0.9, value=referencia)
    hay_previo = not es_vacio(previo)
    indicador = go.Indicator(mode="gauge+number+delta" if hay_previo else "gauge+number", value=valor,
                             gauge=gauge, domain=dict(x=[0, 1], y=[0, 1]),
                             number=dict(valueformat=".1f", font=dict(size=34, color=COLOR_AZUL_PPT)))
    if hay_previo:  # ▲ rojo si sube, ▼ azul si baja
        indicador.delta = dict(reference=float(previo), valueformat=".1f", position="bottom",
                               increasing=dict(color=C["empeora"], symbol="▲ "),
                               decreasing=dict(color=C["mejora"], symbol="▼ "), font=dict(size=13))
    fig = go.Figure(indicador)
    fig.update_layout(margin=dict(l=22, r=22, t=14, b=0))
    return fig


def grafico_ranking_id(corte, dpto_sel, col):
    """Gráfico 1: departamentos de mayor a menor, en colores de cuartil."""
    datos = corte.sort_values(col)  # el mayor queda arriba
    dec = INDICADORES_ID[col]["decimales"]
    sel = (datos[COL_DPTO] == dpto_sel).to_numpy()
    fig = go.Figure(go.Bar(
        y=datos[COL_DPTO], x=datos[col], orientation="h",
        marker=dict(color=colores_clase(datos[col_cuartil(col)]), line=contorno(sel)),
        text=[numero(v, dec) for v in datos[col]], textposition="outside", cliponaxis=False,
        textfont=dict(size=10.5, color=COLOR_AZUL_PPT), customdata=datos[["NOMBRE"]].to_numpy(),
        hovertemplate="<b>%{customdata[0]}</b>: %{text}<extra></extra>"))
    fig.update_layout(bargap=0.2, showlegend=False, margin=dict(l=4, r=8, t=4, b=4),
                      xaxis=dict(visible=False, range=[0, corte[col].max() * 1.15]),
                      yaxis=dict(tickvals=list(datos[COL_DPTO]), showgrid=False, showline=True, linecolor=C["borde"],
                                 ticktext=[f"<b>{n}</b>" if s else n for n, s in zip(datos["NOMBRE"], sel)],
                                 tickfont=dict(size=10.5, color=C["tinta"])))
    return fig


def grafico_mapa_id(corte, geojson, centros, dpto_sel, col):
    """Gráfico 2: mapa departamental en colores de cuartil."""
    info = INDICADORES_ID[col]
    textos = textos_clases(corte[col], corte[col_cuartil(col)], formato_id(col), tope=info["tope"])
    fig = go.Figure()
    hover = ("<b>%{customdata[0]}</b><br>" + info["corto"] + ": %{customdata[1]}"
             "<br>Puesto: %{customdata[2]} de " + str(len(corte)) + "<extra></extra>")
    for q in sorted(COLORES_CUARTIL, reverse=True):
        sub = corte[corte[col_cuartil(col)] == q]
        color = COLORES_CUARTIL[q]
        fig.add_trace(go.Choropleth(
            geojson=geojson, featureidkey=CLAVE_GEOJSON, locations=sub[COL_DPTO], z=[q] * len(sub),
            colorscale=[[0, color], [1, color]], showscale=False, name=textos[q], showlegend=True,
            marker=dict(line=dict(color=BORDE_MAPA, width=0.6)),
            customdata=[[n, numero(v, info["decimales"]), int(p)]
                        for n, v, p in zip(sub["NOMBRE"], sub[col], sub[col_puesto(col)])],
            hovertemplate=hover))
    if dpto_sel != NACIONAL:  # contorno del seleccionado
        fig.add_trace(go.Choropleth(
            geojson=geojson, featureidkey=CLAVE_GEOJSON, locations=[dpto_sel], z=[1],
            colorscale=[[0, "rgba(0,0,0,0)"], [1, "rgba(0,0,0,0)"]], showscale=False,
            showlegend=False, hoverinfo="skip", marker=dict(line=dict(color=COLOR_AZUL_PPT, width=2.8))))
    filas = corte.set_index(COL_DPTO)
    top3 = [d for d in corte[COL_DPTO].head(3) if d in centros]  # insignias 1, 2 y 3
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
    fig.update_geos(fitbounds="locations", visible=False, bgcolor="rgba(0,0,0,0)", projection_type="mercator")
    fig.update_layout(margin=dict(l=0, r=0, t=0, b=0), dragmode=False,
                      legend=dict(orientation="v", x=0.02, y=0.04, yanchor="bottom",
                                  font=dict(size=12, color=C["tinta"])))
    return fig


def grafico_evolucion_id(serie_nac, anio, col, serie_dpto=None, nombre_dpto=""):
    """Gráfico 3: evolución del promedio del Perú (y del departamento elegido), con tendencia y eventos."""
    info = INDICADORES_ID[col]
    dec = info["decimales"]
    fig = go.Figure()
    if serie_dpto is not None:
        principal, nombre = serie_dpto[col], nombre_dpto
        fig.add_scatter(x=serie_nac.index, y=serie_nac[col], name="Promedio del Perú", mode="lines",
                        line=dict(color=C["contexto"], width=2),
                        hovertemplate="Promedio del Perú: %{y:.2f}<extra></extra>")
    else:
        principal, nombre = serie_nac[col], "Promedio del Perú"
    fig.add_scatter(x=principal.index, y=principal, name=nombre, mode="lines+markers+text",
                    line=dict(color=COLOR_AZUL_PPT, width=2.6), marker=dict(size=6, color=COLOR_AZUL_PPT),
                    text=[numero(v, dec) for v in principal], textposition="top center",
                    textfont=dict(size=10.5, color=COLOR_AZUL_PPT),
                    hovertemplate=f"{nombre}: %{{y:.2f}}<extra></extra>")
    validos = principal.dropna()
    if len(validos) > 2:  # tendencia lineal
        x = validos.index.astype(float).to_numpy()
        pendiente, intercepto = np.polyfit(x, validos.to_numpy(dtype=float), 1)
        fig.add_scatter(x=validos.index, y=intercepto + pendiente * x, mode="lines", hoverinfo="skip",
                        name=f"Tendencia ({con_signo(pendiente, 3)} por año)",
                        line=dict(color=C["pacifico_medio"], width=1.6, dash="dot"))
    eventos = [a for a in EVENTOS if a in principal.index and not es_vacio(principal.loc[a])]
    fig.add_scatter(x=eventos, y=[principal.loc[a] for a in eventos], mode="markers", showlegend=False,
                    hoverinfo="skip", marker=dict(size=10, color=COLOR_ROJO_PPT, line=dict(color="white", width=1.5)))
    for a in eventos:
        fig.add_annotation(x=a, y=principal.loc[a], text=EVENTOS[a], ax=-44, ay=-46, showarrow=True, arrowhead=0,
                           arrowwidth=1, arrowcolor=COLOR_ROJO_PPT, font=dict(size=10.5, color=COLOR_ROJO_PPT),
                           bgcolor=hex_a_rgba(C["superficie"], 0.9))
    if anio in principal.index and not es_vacio(principal.loc[anio]):
        fig.add_scatter(x=[anio], y=[principal.loc[anio]], mode="markers", showlegend=False, hoverinfo="skip",
                        marker=dict(size=14, color="rgba(0,0,0,0)", line=dict(color=C["pacifico_medio"], width=3)))
    marcar_anio(fig, anio)
    tope = np.nanmax([principal.max(), serie_nac[col].max()])
    fig.update_layout(hovermode="x unified", margin=dict(l=4, r=4, t=36, b=4),
                      legend=dict(y=1.04, traceorder="normal"),
                      xaxis=dict(dtick=2, tick0=int(principal.index.min()), tickfont=dict(size=10.5)),
                      yaxis=dict(range=[0, tope * 1.3], tickformat=f".{dec}f"))
    return fig


def grafico_relacion_id(base, dpto_sel, anio, col):
    """Gráfico 4: frecuencia contra severidad global, año a año (con «Reproducir»).
    Burbuja = departamento; tamaño = ID; color = cuartil del indicador; curvas = mismo ID.
    Devuelve un dict (Dash lo acepta) porque así los cuadros de la animación se arman rápido."""
    anios = lista_anios(base)
    max_y = base[COL_SEV].max() * 1.12
    tam_ref = base[COL_ID].max() / 34 ** 2
    cq = col_cuartil(col)

    def trazas_del_anio(a):
        d = base[base[COL_ANIO] == a].sort_values(COL_DPTO)
        etiquetados = set(d.nlargest(3, col)[COL_DPTO]) if dpto_sel == NACIONAL else {dpto_sel}
        principal = dict(
            type="scatter", x=d[COL_FREC].tolist(), y=d[COL_SEV].tolist(), ids=d[COL_DPTO].tolist(),
            mode="markers+text", textposition="top center", showlegend=False,
            text=[n if k in etiquetados else "" for k, n in zip(d[COL_DPTO], d["NOMBRE"])],
            textfont=dict(size=10.5, color=C["tinta"]),
            marker=dict(size=d[COL_ID].clip(lower=0.05).tolist(), sizemode="area", sizeref=tam_ref, sizemin=5,
                        opacity=0.9, line=dict(color=BORDE_MAPA, width=0.8), color=colores_clase(d[cq])),
            customdata=[[k, n, numero(i, 1)] for k, n, i in zip(d[COL_DPTO], d["NOMBRE"], d[COL_ID])],
            hovertemplate=("<b>%{customdata[1]}</b> (" + str(a) + ")<br>Frecuencia: %{x:.2f}"
                           "<br>Severidad global: %{y:.2f}<br>ID: %{customdata[2]}<extra></extra>"))
        sel = d[d[COL_DPTO] == dpto_sel]
        anillo = dict(type="scatter", x=sel[COL_FREC].tolist(), y=sel[COL_SEV].tolist(), mode="markers",
                      hoverinfo="skip", showlegend=False,
                      marker=dict(size=[max(np.sqrt(v / tam_ref), 5) + 9 for v in sel[COL_ID]],
                                  symbol="circle-open", color=COLOR_AZUL_PPT, line=dict(width=2.5)))
        return [principal, anillo]

    leyenda = [go.Scatter(x=[None], y=[None], mode="markers", name=f"Cuartil {q}",
                          marker=dict(size=11, color=COLORES_CUARTIL[q], line=dict(color=BORDE_MAPA, width=0.8)))
               for q in sorted(COLORES_CUARTIL, reverse=True)]
    fig = go.Figure(data=trazas_del_anio(anio) + leyenda)
    for valor_id in (2, 4, 6):  # curvas de igual ID: severidad = ID / (10 x frecuencia)
        xs = np.linspace(max(valor_id / 10 / max_y, 0.02), 1.02, 60)
        fig.add_scatter(x=xs, y=valor_id / (10 * xs), mode="lines", hoverinfo="skip", showlegend=False,
                        line=dict(color=C["contexto"], width=1, dash="dash"))
        fig.add_annotation(x=1.02, y=valor_id / 10.2, text=f"ID = {valor_id}", showarrow=False, xanchor="left",
                           font=dict(size=10, color=C["tinta_3"]))
    for x, y, texto, xa, ya in [(1.0, max_y, "Frecuente y severo", "right", "top"),
                                (0.0, 0.0, "Poco frecuente y leve", "left", "bottom")]:
        fig.add_annotation(x=x, y=y, text=texto, showarrow=False, xanchor=xa, yanchor=ya,
                           font=dict(size=11, color=C["tinta_3"]))
    animar = dict(frame=dict(duration=900, redraw=False), fromcurrent=True,
                  transition=dict(duration=600, easing="cubic-in-out"))
    pausar = dict(frame=dict(duration=0, redraw=False), mode="immediate", transition=dict(duration=0))
    fig.update_layout(
        margin=dict(l=4, r=40, t=36, b=80), legend=dict(y=1.04),
        xaxis=dict(title="Índice de frecuencia", range=[0, 1.02], showgrid=True, gridcolor=C["rejilla"]),
        yaxis=dict(title="Severidad global", range=[0, max_y]),
        updatemenus=[dict(type="buttons", direction="left", x=0, xanchor="left", y=-0.16, yanchor="top",
                          showactive=False, pad=dict(t=0, r=8), bgcolor=C["superficie"],
                          bordercolor=C["borde"], font=dict(size=11, color=C["tinta"]),
                          buttons=[dict(label="▶ Reproducir", method="animate", args=[None, animar]),
                                   dict(label="❚❚ Pausa", method="animate", args=[[None], pausar])])],
        sliders=[dict(active=anios.index(anio) if anio in anios else 0,
                      x=0.28, len=0.72, y=-0.12, yanchor="top", pad=dict(t=0, b=0),
                      currentvalue=dict(prefix="Año ", font=dict(size=12, color=C["tinta"]), xanchor="right"),
                      tickcolor=C["borde"], font=dict(size=9, color=C["tinta_3"]),
                      bgcolor=C["rejilla"], activebgcolor=C["pacifico_medio"], bordercolor=C["borde"],
                      steps=[dict(label=str(a) if (a - anios[0]) % 4 == 0 or a == anios[-1] else "",
                                  method="animate", args=[[str(a)], pausar]) for a in anios])])
    figura = fig.to_plotly_json()
    figura["frames"] = [dict(name=str(a), traces=[0, 1], data=trazas_del_anio(a)) for a in anios]
    return figura


def grafico_calor_id(base, dpto_sel, anio, col):
    """Gráfico 5: departamento x año en colores de cuartil. Arriba los más oscuros (cuartil promedio más alto)."""
    cuartiles = base.pivot_table(index=COL_DPTO, columns=COL_ANIO, values=col_cuartil(col))
    valores = base.pivot_table(index=COL_DPTO, columns=COL_ANIO, values=col)
    orden = (pd.DataFrame({"q": cuartiles.mean(axis=1), "v": valores.mean(axis=1)})
             .sort_values(["q", "v"], ascending=False).index)
    cuartiles, valores = cuartiles.loc[orden], valores.loc[orden]
    anios = list(cuartiles.columns)
    escala = []
    for i, q in enumerate(sorted(COLORES_CUARTIL)):  # escala de 4 bloques de color
        escala += [[i / 4, COLORES_CUARTIL[q]], [(i + 1) / 4, COLORES_CUARTIL[q]]]
    dec = INDICADORES_ID[col]["decimales"]
    custom = [[[nombre_visible(d), numero(v, dec)] for v in fila] for d, fila in zip(orden, valores.to_numpy())]
    fig = go.Figure(go.Heatmap(
        z=cuartiles.to_numpy(), x=anios, y=[nombre_visible(d) for d in orden], customdata=custom,
        zmin=0.5, zmax=4.5, colorscale=escala, xgap=2, ygap=2, hoverongaps=False, showscale=False,
        hovertemplate="<b>%{customdata[0]}</b>, %{x}<br>" + INDICADORES_ID[col]["corto"]
                      + ": %{customdata[1]}<br>Cuartil: %{z}<extra></extra>"))
    for q in sorted(COLORES_CUARTIL, reverse=True):  # leyenda
        fig.add_scatter(x=[None], y=[None], mode="markers", name=f"Cuartil {q}",
                        marker=dict(symbol="square", size=12, color=COLORES_CUARTIL[q]))
    marco = dict(type="rect", line=dict(color=COLOR_AZUL_PPT, width=2.2), fillcolor="rgba(0,0,0,0)")
    if dpto_sel in list(orden):
        i = list(orden).index(dpto_sel)
        fig.add_shape(**marco, x0=anios[0] - 0.5, x1=anios[-1] + 0.5, y0=i - 0.5, y1=i + 0.5)
    if anio in anios:
        fig.add_shape(**marco, x0=anio - 0.5, x1=anio + 0.5, y0=-0.5, y1=len(orden) - 0.5)
    fig.update_layout(margin=dict(l=4, r=4, t=30, b=4), legend=dict(y=1.0, x=1, xanchor="right"),
                      xaxis=dict(dtick=1, tickfont=dict(size=10.5), showline=False, side="top"),
                      yaxis=dict(autorange="reversed", tickfont=dict(size=11, color=C["tinta"]), showgrid=False))
    return fig


# ---- 4.3 INFORM Risk ------------------------------------------------------------------
def grafico_escala_inform(valor):
    """Tarjeta: escala de 0 a 10 con los colores de las clases y una marca en el valor (como la presentación)."""
    fig = go.Figure()
    limites = [c[1] for c in CLASES_INFORM] + [10]
    for (nombre, desde, color), hasta in zip(CLASES_INFORM, limites[1:]):
        fig.add_shape(type="rect", x0=desde, x1=hasta, y0=0, y1=1, fillcolor=color, line=dict(width=0))
    if not es_vacio(valor):
        fig.add_scatter(x=[valor], y=[1.45], mode="markers", hoverinfo="skip",
                        marker=dict(symbol="triangle-down", size=13, color=COLOR_AZUL_PPT))
    fig.update_layout(margin=dict(l=4, r=4, t=2, b=2), showlegend=False,
                      xaxis=dict(range=[0, 10], tickvals=[0, 10], showline=False, tickfont=dict(size=10),
                                 fixedrange=True),
                      yaxis=dict(visible=False, range=[-0.1, 1.9], fixedrange=True))
    return fig


def leyenda_clases_inform(fig, presentes=None):
    """Agrega a la figura la leyenda de las clases del INFORM (de mayor a menor riesgo)."""
    for nombre, _, color in reversed(CLASES_INFORM):
        if presentes is None or nombre in presentes:
            fig.add_bar(x=[None], y=[None], name=nombre, marker_color=color, orientation="h")


def grafico_ranking_inform(alc, iso_sel, col):
    """Gráfico 1: los 21 países de mayor a menor, con el color de su clase de riesgo."""
    datos = alc.sort_values(col)
    clases = [clase_inform(v) for v in datos[col]]
    sel = (datos.index == iso_sel)
    fig = go.Figure(go.Bar(
        y=datos["NOMBRE"], x=datos[col], orientation="h", showlegend=False,
        marker=dict(color=[c[1] for c in clases], line=contorno(sel)),
        text=[numero(v, 1) for v in datos[col]], textposition="outside", cliponaxis=False,
        textfont=dict(size=10.5, color=COLOR_AZUL_PPT), customdata=[[c[0]] for c in clases],
        hovertemplate="<b>%{y}</b>: %{text} · %{customdata[0]}<extra></extra>"))
    leyenda_clases_inform(fig, {c[0] for c in clases})
    fig.update_layout(bargap=0.2, barmode="overlay", margin=dict(l=4, r=8, t=30, b=4), legend=dict(y=1.02),
                      xaxis=dict(visible=False, range=[0, 10]),
                      yaxis=dict(showgrid=False, showline=True, linecolor=C["borde"], tickvals=list(datos["NOMBRE"]),
                                 ticktext=[f"<b>{n}</b>" if s else n for n, s in zip(datos["NOMBRE"], sel)],
                                 tickfont=dict(size=10.5, color=C["tinta"])))
    return fig


def grafico_mapa_inform(inform, iso_sel, col):
    """Gráfico 2: mapa de la región con el nivel de riesgo (clases del INFORM)."""
    region = inform.loc[[i for i in list(PAISES_INFORM) + list(PAISES_INFORM_MAPA) if i in inform.index]]
    clases = pd.Series([clase_inform(v)[0] for v in region[col]], index=region.index)
    fig = go.Figure()
    for nombre, _, color in reversed(CLASES_INFORM):
        sub = region[clases == nombre]
        if sub.empty:
            continue
        fig.add_trace(go.Choropleth(
            locations=sub.index, z=[1] * len(sub), locationmode="ISO-3", showscale=False, name=nombre,
            showlegend=True,
            colorscale=[[0, color], [1, color]], marker=dict(line=dict(color="#ffffff", width=0.7)),
            customdata=[[n, numero(v, 1)] for n, v in zip(sub["NOMBRE"], sub[col])],
            hovertemplate="<b>%{customdata[0]}</b>: %{customdata[1]}<extra>" + nombre + "</extra>"))
    if iso_sel in region.index:
        fig.add_trace(go.Choropleth(
            locations=[iso_sel], z=[1], locationmode="ISO-3", showscale=False, showlegend=False, hoverinfo="skip",
            colorscale=[[0, "rgba(0,0,0,0)"], [1, "rgba(0,0,0,0)"]],
            marker=dict(line=dict(color=COLOR_AZUL_PPT, width=2.4))))
    fig.update_geos(projection_type="mercator", resolution=50, showframe=False, showcoastlines=False,
                    showcountries=True, countrycolor="#ffffff", showland=True, landcolor="#e4ebf2",
                    showocean=False, showlakes=False, bgcolor="rgba(0,0,0,0)",
                    lonaxis=dict(range=[-118, -33]), lataxis=dict(range=[-56, 33]))
    fig.update_layout(margin=dict(l=0, r=0, t=0, b=0), dragmode=False,
                      legend=dict(orientation="v", x=0.02, y=0.05, yanchor="bottom", traceorder="normal",
                                  font=dict(size=12, color=C["tinta"]), bgcolor=hex_a_rgba(C["superficie"], 0.85)))
    return fig


def grafico_grupos_inform(inform, iso_sel, col):
    """Gráfico 3: América Latina frente a economías avanzadas y emergentes (como el Gráfico N.º 11)."""
    grupos, nombres, valores, colores, sel = [], [], [], [], []
    for grupo, paises in GRUPOS_INFORM.items():
        paises = dict(paises)
        if grupo == "América Latina" and iso_sel not in paises and iso_sel in PAISES_INFORM:
            paises[iso_sel] = PAISES_INFORM[iso_sel]
        orden = sorted((iso for iso in paises if iso in inform.index), key=lambda i: inform.loc[i, col])
        for iso in orden:
            grupos.append(grupo)
            nombres.append(paises[iso])
            valores.append(inform.loc[iso, col])
            colores.append(COLORES_GRUPO_INFORM[grupo])
            sel.append(iso == iso_sel)
    fig = go.Figure(go.Bar(
        x=[grupos, nombres], y=valores, marker=dict(color=colores, line=contorno(sel)),
        text=[numero(v, 1) for v in valores], textposition="outside", cliponaxis=False,
        textfont=dict(size=11, color=COLOR_AZUL_PPT), hovertemplate="<b>%{x}</b>: %{text}<extra></extra>"))
    fig.update_layout(bargap=0.3, showlegend=False, margin=dict(l=4, r=4, t=10, b=4),
                      xaxis=dict(tickfont=dict(size=10.5, color=C["tinta"]), tickangle=-45),
                      yaxis=dict(range=[0, max(valores) * 1.2], title="Puntaje de 0 a 10"))
    return fig


def grafico_amenazas_inform(inform, iso_sel):
    """Gráfico 4: amenazas naturales del país frente al promedio de América Latina (21 países)."""
    fila, alc = inform.loc[iso_sel], tabla_alc(inform)
    datos = pd.DataFrame({"pais": [fila[c] for c in AMENAZAS_NATURALES],
                          "alc": [alc[c].mean() for c in AMENAZAS_NATURALES]},
                         index=list(AMENAZAS_NATURALES.values())).sort_values("pais")
    nombre = fila["NOMBRE"]
    fig = go.Figure()
    fig.add_bar(y=datos.index, x=datos["pais"], orientation="h", name=nombre, marker_color=COLOR_ROJO_PPT,
                text=[numero(v, 1) for v in datos["pais"]], textposition="outside", cliponaxis=False,
                textfont=dict(size=11, color=COLOR_AZUL_PPT),
                hovertemplate=f"<b>%{{y}}</b> · {nombre}: %{{x:.1f}}<extra></extra>")
    fig.add_scatter(y=datos.index, x=datos["alc"], mode="markers", name="Promedio de América Latina",
                    marker=dict(symbol="diamond", size=11, color=COLOR_AZUL_PPT, line=dict(color="white", width=1)),
                    hovertemplate="<b>%{y}</b> · promedio ALC: %{x:.1f}<extra></extra>")
    fig.update_layout(bargap=0.35, margin=dict(l=4, r=8, t=30, b=4), legend=dict(y=1.02),
                      xaxis=dict(range=[0, 10.8], showgrid=True, gridcolor=C["rejilla"], showline=False),
                      yaxis=dict(tickfont=dict(size=11.5, color=C["tinta"])))
    return fig


def grafico_dimensiones_inform(alc, iso_sel):
    """Gráfico 5: las 3 dimensiones de la fórmula en cada país (puntos) y el INFORM Risk (raya)."""
    datos = alc.sort_values("INFORM RISK")
    nombres = list(datos["NOMBRE"])
    fig = go.Figure()
    for nombre, fila in zip(nombres, datos.itertuples()):  # línea gris entre el mínimo y el máximo
        vals = [datos.loc[fila.Index, c] for c in DIMENSIONES_INFORM]
        fig.add_scatter(x=[min(vals), max(vals)], y=[nombre, nombre], mode="lines", showlegend=False,
                        hoverinfo="skip", line=dict(color=C["borde"], width=3))
    for col, (texto, color) in DIMENSIONES_INFORM.items():
        fig.add_scatter(x=datos[col], y=nombres, mode="markers", name=texto,
                        marker=dict(size=11, color=color, line=dict(color="white", width=1)),
                        hovertemplate="<b>%{y}</b> · " + texto + ": %{x:.1f}<extra></extra>")
    fig.add_scatter(x=datos["INFORM RISK"], y=nombres, mode="markers", name="INFORM Risk",
                    marker=dict(symbol="line-ns", size=16, line=dict(color=C["tinta"], width=2.5)),
                    hovertemplate="<b>%{y}</b> · INFORM Risk: %{x:.1f}<extra></extra>")
    fig.update_layout(margin=dict(l=4, r=8, t=30, b=4), legend=dict(y=1.02),
                      xaxis=dict(range=[0, 10], showgrid=True, gridcolor=C["rejilla"], showline=False,
                                 title="Puntaje de 0 a 10"),
                      yaxis=dict(categoryorder="array", categoryarray=nombres, tickvals=nombres,
                                 ticktext=[f"<b>{n}</b>" if i == iso_sel else n for n, i in zip(nombres, datos.index)],
                                 tickfont=dict(size=10.5, color=C["tinta"]), showgrid=False))
    return fig


def grafico_subcomponentes_inform(inform, iso_sel):
    """Gráfico 6: qué explica la vulnerabilidad y la falta de capacidad de respuesta del país."""
    fila, alc = inform.loc[iso_sel], tabla_alc(inform)
    dims = [d for d, _ in SUBCOMPONENTES_INFORM.values()]
    nombres = [n for _, n in SUBCOMPONENTES_INFORM.values()]
    colores = {"Vulnerabilidad": DIMENSIONES_INFORM["VULNERABILITY"][1],
               "Falta de capacidad": DIMENSIONES_INFORM["LACK OF COPING CAPACITY"][1]}
    valores = [fila[c] for c in SUBCOMPONENTES_INFORM]
    fig = go.Figure()
    fig.add_bar(y=[dims, nombres], x=valores, orientation="h", name=fila["NOMBRE"], showlegend=False,
                marker_color=[colores[d] for d in dims], text=[numero(v, 1) for v in valores],
                textposition="outside", cliponaxis=False, textfont=dict(size=11, color=COLOR_AZUL_PPT),
                hovertemplate="<b>%{y}</b>: %{x:.1f}<extra></extra>")
    fig.add_scatter(y=[dims, nombres], x=[alc[c].mean() for c in SUBCOMPONENTES_INFORM], mode="markers",
                    name="Promedio de América Latina",
                    marker=dict(symbol="diamond", size=10, color=C["tinta"], line=dict(color="white", width=1)),
                    hovertemplate="<b>%{y}</b> · promedio ALC: %{x:.1f}<extra></extra>")
    fig.update_layout(bargap=0.3, margin=dict(l=4, r=8, t=30, b=4), legend=dict(y=1.02),
                      xaxis=dict(range=[0, 10.8], showgrid=True, gridcolor=C["rejilla"], showline=False),
                      yaxis=dict(autorange="reversed", tickfont=dict(size=11, color=C["tinta"])))
    return fig


# ---- 4.4 América Latina ------------------------------------------------------------------
def clases_latam(tabla, ind):
    return clases_cuartil(tabla[ind], METODO_CUARTIL_LATAM[ind])[0]


def grafico_ranking_latam(tabla, pais_sel, ind):
    """Gráfico 1: ranking de países en colores de cuartil."""
    datos = tabla.assign(q=clases_latam(tabla, ind)).sort_values(ind)
    sel = datos.index == pais_sel
    fig = go.Figure(go.Bar(
        y=datos.index, x=datos[ind], orientation="h",
        marker=dict(color=[COLORES_CUARTIL[q] for q in datos["q"]], line=contorno(sel)),
        text=[valor_latam(v, ind) for v in datos[ind]], textposition="outside", cliponaxis=False,
        textfont=dict(size=10.5, color=COLOR_AZUL_PPT),
        hovertemplate="<b>%{y}</b>: %{text}<extra></extra>"))
    fig.update_layout(margin=dict(l=4, r=8, t=4, b=4), bargap=0.22, showlegend=False,
                      xaxis=dict(visible=False, range=[0, max(datos[ind].max(), 1e-9) * 1.16]),
                      yaxis=dict(tickvals=list(datos.index),
                                 ticktext=[f"<b>{p}</b>" if p == pais_sel else p for p in datos.index],
                                 tickfont=dict(size=10.5, color=C["tinta"]), showline=True, linecolor=C["borde"]))
    return fig


def grafico_mapa_latam(tabla, pais_sel, ind):
    """Gráfico 2: mapa de América Latina (mismos colores que el ranking)."""
    datos = tabla.assign(q=clases_latam(tabla, ind))
    textos = textos_clases(datos[ind], datos["q"], lambda v: valor_latam(v, ind))
    fig = go.Figure()
    for q in sorted(COLORES_CUARTIL, reverse=True):
        sub = datos[datos["q"] == q]
        color = COLORES_CUARTIL[q]
        fig.add_trace(go.Choropleth(
            locations=sub["iso3"], z=[q] * len(sub), locationmode="ISO-3", showscale=False,
            colorscale=[[0, color], [1, color]], name=textos[q],
            showlegend=True, marker=dict(line=dict(color="#ffffff", width=0.7)),
            customdata=[[p, valor_latam(v, ind)] for p, v in zip(sub.index, sub[ind])],
            hovertemplate="<b>%{customdata[0]}</b>: %{customdata[1]}<extra></extra>"))
    if pais_sel in datos.index:
        fig.add_trace(go.Choropleth(
            locations=[datos.loc[pais_sel, "iso3"]], z=[1], locationmode="ISO-3", showscale=False,
            colorscale=[[0, "rgba(0,0,0,0)"], [1, "rgba(0,0,0,0)"]], showlegend=False, hoverinfo="skip",
            marker=dict(line=dict(color=COLOR_AZUL_PPT, width=2.4))))
    fig.update_geos(projection_type="mercator", resolution=50, showframe=False, showcoastlines=False,
                    showcountries=True, countrycolor="#ffffff", showland=True, landcolor="#e4ebf2",
                    showocean=False, showlakes=False, bgcolor="rgba(0,0,0,0)",
                    lonaxis=dict(range=[-118, -33]), lataxis=dict(range=[-56, 33]))
    fig.update_layout(margin=dict(l=0, r=0, t=0, b=0), dragmode=False,
                      legend=dict(orientation="v", x=0.02, y=0.05, yanchor="bottom",
                                  font=dict(size=12, color=C["tinta"]),
                                  bgcolor=hex_a_rgba(C["superficie"], 0.85)))
    return fig


def grafico_periodos_latam(latam, pais_sel):
    e = latam["eventos"]
    antes = e[e["anio"] <= CORTE_PERIODOS].groupby("pais").size()
    despues = e[e["anio"] > CORTE_PERIODOS].groupby("pais").size()
    tabla = pd.DataFrame({"antes": antes, "despues": despues}).fillna(0)
    tabla = tabla.loc[tabla.sum(axis=1).sort_values(ascending=False).index]   # mismo orden que el Gráfico 1
    etiquetados = set(tabla.index[:5]) | {pais_sel}   # cifras solo en los 5 primeros y el país elegido
    a0, a1 = int(e["anio"].min()), int(e["anio"].max())
    fig = go.Figure()
    for col, nombre, color in [("antes", f"{a0}-{CORTE_PERIODOS}", COLOR_PERIODO_1),
                               ("despues", f"{CORTE_PERIODOS + 1}-{a1}", COLOR_PERIODO_2)]:
        fig.add_bar(x=tabla.index, y=tabla[col], name=nombre,
                    marker=dict(color=color, line=dict(color=[COLOR_AZUL_PPT if p == pais_sel else color
                                                              for p in tabla.index],
                                                       width=[2.5 if p == pais_sel else 0 for p in tabla.index])),
                    text=[numero(v, 0) if p in etiquetados else "" for p, v in tabla[col].items()],
                    textposition="outside", cliponaxis=False, textangle=0, constraintext="none",
                    textfont=dict(size=10, color=C["tinta_2"]),
                    hovertemplate="<b>%{x}</b> " + nombre + ": %{y}<extra></extra>")
    fig.update_layout(barmode="group", bargap=0.25, bargroupgap=0.05, margin=dict(l=4, r=4, t=36, b=4),
                      legend=dict(y=1.04),
                      xaxis=dict(tickangle=-45, tickvals=list(tabla.index),
                                 ticktext=[f"<b>{p}</b>" if p == pais_sel else p for p in tabla.index],
                                 tickfont=dict(size=10.5, color=C["tinta"])),
                      yaxis=dict(title="N.º de desastres", range=[0, tabla.values.max() * 1.15]))
    return fig


def grafico_dispersion_latam(latam, a0, a1):
    e = eventos_periodo(latam, a0, min(a1, ULTIMO_ANIO_COSTO))
    anual = e.groupby("anio").agg(n=("id", "size"), costo=("dano_miles_usd", "sum"))
    anual["costo"] = anual["costo"] / 1e6  # miles de millones de US$
    if len(anual) < 3:
        return figura_vacia("Elige un periodo más largo para ver la relación.")
    fig = go.Figure()
    fig.add_scatter(x=anual["n"], y=anual["costo"], mode="markers", showlegend=False,
                    marker=dict(size=11, color=COLOR_DISPERSION, opacity=0.9, line=dict(color="#ffffff", width=1)),
                    customdata=[[a, numero(c, 1)] for a, c in zip(anual.index, anual["costo"])],
                    hovertemplate="<b>%{customdata[0]}</b><br>%{x} desastres<br>"
                                  "US$ %{customdata[1]} mil millones<extra></extra>")
    pendiente, intercepto = np.polyfit(anual["n"], anual["costo"], 1)
    xs = np.array([anual["n"].min(), anual["n"].max()])
    fig.add_scatter(x=xs, y=intercepto + pendiente * xs, mode="lines", hoverinfo="skip", showlegend=False,
                    line=dict(color=COLOR_TENDENCIA, width=2, dash="dash"))
    for anio in anual["costo"].nlargest(3).index:
        fig.add_annotation(x=anual.loc[anio, "n"], y=anual.loc[anio, "costo"], text=str(anio), showarrow=False,
                           yshift=14, font=dict(size=11, color=C["tinta"]))
    r = np.corrcoef(anual["n"], anual["costo"])[0, 1]
    fig.add_annotation(x=0.01, y=0.98, xref="paper", yref="paper", xanchor="left", yanchor="top", showarrow=False,
                       text=f"Correlación: <b>{numero(r, 2)}</b>", font=dict(size=12, color=C["tinta"]),
                       bgcolor=C["niebla"], borderpad=6)
    fig.update_layout(margin=dict(l=4, r=4, t=8, b=4), hovermode="closest",
                      xaxis=dict(title="N.º de desastres anuales", showgrid=True, gridcolor=C["rejilla"]),
                      yaxis=dict(title="Costo anual (miles de millones de US$)", rangemode="tozero"))
    return fig


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
                    text=[f"{numero(v, 1)}%" if v >= 4 else "" for v in partes[sub].fillna(0)], textposition="inside",
                    insidetextanchor="middle", textangle=0,
                    textfont=dict(size=9.5, color="#ffffff" if oscuro else C["tinta"]),
                    hovertemplate="<b>%{x}</b> · " + sub + ": %{y:.1f} %<extra></extra>")
    fig.update_layout(barmode="stack", bargap=0.28, margin=dict(l=4, r=4, t=8, b=4),
                      legend=dict(orientation="h", y=-0.12, yanchor="top", x=0.5, xanchor="center",
                                  traceorder="normal", font=dict(size=12)),
                      xaxis=dict(tickangle=-35 if len(paises) > 8 else 0, tickvals=list(partes.index),
                                 ticktext=[f"<b>{p}</b>" if p == pais_sel else p for p in partes.index],
                                 tickfont=dict(size=10.5, color=C["tinta"])),
                      yaxis=dict(range=[0, 100], ticksuffix=" %"))
    return fig


def grafico_tipos_pais(latam, pais, a0, a1, medida):
    e = eventos_periodo(latam, a0, a1)
    e = e[e["pais"] == pais]
    pares = latam["eventos"][["subgrupo", "tipo"]].drop_duplicates()
    if medida == "eventos":
        valores = e.groupby("tipo").size()
    else:
        pbi = latam["pbi"].loc[pais, "pbi_millones_usd"] * 1000
        valores = e.groupby("tipo")["dano_miles_usd"].sum() / pbi * 100
    pares = pares.assign(valor=pares["tipo"].map(valores).fillna(0).astype(float))
    pares["orden"] = pares["subgrupo"].map({s: i for i, s in enumerate(ORDEN_SUBGRUPOS_PAIS)})
    pares = pares.sort_values(["orden", "valor"], ascending=[True, False])
    pares["nombre"] = pares["tipo"].map(nombre_corto)
    fig = go.Figure()
    for sub in ORDEN_SUBGRUPOS_PAIS:
        d = pares[pares["subgrupo"] == sub]
        textos = [numero(v, 0) for v in d["valor"]] if medida == "eventos" else [porcentaje_corto(v) for v in d["valor"]]
        fig.add_bar(y=d["nombre"], x=d["valor"], orientation="h", name=sub, marker_color=COLORES_TIPO[sub],
                    text=textos, textposition="outside", cliponaxis=False,
                    textfont=dict(size=11, color=COLOR_AZUL_PPT), customdata=textos,
                    hovertemplate="<b>%{y}</b> (" + sub + "): %{customdata}<extra></extra>")
    tope = pares["valor"].max() or 1
    fig.update_layout(margin=dict(l=4, r=4, t=4, b=4), bargap=0.3, barmode="overlay",
                      legend=dict(orientation="h", y=-0.04, yanchor="top", x=0.5, xanchor="center",
                                  traceorder="normal", font=dict(size=12)),
                      xaxis=dict(visible=False, range=[0, tope * 1.15]),
                      yaxis=dict(categoryorder="array", categoryarray=list(pares["nombre"]), autorange="reversed",
                                 tickfont=dict(size=11.5, color=C["tinta"]), showline=True, linecolor=C["borde"],
                                 showgrid=False))
    return fig


def grafico_anual_pais(latam, pais, a0, a1, medida):
    """Número: barras por tipo + promedio móvil de 5 años. Costo: millones de US$ con los años más costosos."""
    costo = medida != "eventos"
    if costo:
        a1 = min(a1, ULTIMO_ANIO_COSTO)
        if a0 > a1:
            return figura_vacia(f"EM-DAT aún no tiene costos de {a0}.")
    e = eventos_periodo(latam, a0, a1)
    e = e[e["pais"] == pais]
    if costo:
        tabla = e.pivot_table(index="anio", columns="subgrupo", values="dano_miles_usd", aggfunc="sum") / 1000
    else:
        tabla = e.pivot_table(index="anio", columns="subgrupo", values="id", aggfunc="count")
    tabla = tabla.reindex(index=range(a0, a1 + 1), columns=ORDEN_SUBGRUPOS_PAIS).fillna(0)
    fig = go.Figure()
    formato = ":,.0f"
    for sub in ORDEN_SUBGRUPOS_PAIS:
        fig.add_bar(x=tabla.index, y=tabla[sub], name=sub, marker=dict(color=COLORES_TIPO[sub], line=dict(width=0)),
                    hovertemplate=sub + f": %{{y{formato}}}" + (" millones US$" if costo else "") + "<extra></extra>")
    total = tabla.sum(axis=1)
    if costo:  # los 3 años más costosos, con el desastre que más costó
        for anio in total[total > 0].nlargest(3).index:
            fila = e[e["anio"] == anio].nlargest(1, "dano_miles_usd").iloc[0]
            fig.add_annotation(x=anio, y=total[anio], yshift=12, showarrow=False,
                               text=f"<b>{anio}</b> · {nombre_corto(fila['tipo'])}",
                               font=dict(size=10.5, color=COLOR_AZUL_PPT), bgcolor=hex_a_rgba(C["superficie"], 0.85))
        titulo_y = "Millones de US$"
    else:
        movil = total.rolling(5, min_periods=1).mean()
        fig.add_scatter(x=movil.index, y=movil, name="Promedio de 5 años", mode="lines",
                        line=dict(color=COLOR_AZUL_PPT, width=2.4, shape="spline", smoothing=0.6),
                        hovertemplate="Promedio de 5 años: %{y:.1f}<extra></extra>")
        titulo_y = "N.º de desastres"
    fig.update_layout(barmode="stack", bargap=0.18, hovermode="x unified", margin=dict(l=4, r=4, t=36, b=4),
                      legend=dict(y=1.04, traceorder="normal"),
                      xaxis=dict(dtick=5 if a1 - a0 > 20 else (2 if a1 - a0 > 10 else 1), tickfont=dict(size=10.5)),
                      yaxis=dict(title=titulo_y, rangemode="tozero"))
    if costo:
        fig.update_yaxes(range=[0, max(total.max(), 1) * 1.15])
    return fig




# =============================================================================
# 5. DISEÑO  (íconos, piezas comunes, portada, una función por página y estilos CSS)
# =============================================================================
# ---- 5.1 Íconos (dibujados en SVG de 24 x 24) ---------------------------------------
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
    "sismo": "<path d='M2 12h4l2-6 3 12 3-9 2 5 2-2h4'/>",
    "nube": "<path d='M7 17a4.5 4.5 0 1 1 .9-8.9A6 6 0 0 1 19 10a3.5 3.5 0 0 1-.5 7z'/><path d='M12 13l-2 4h3l-2 4'/>",
    "virus": "<circle cx='12' cy='12' r='5'/><path d='M12 2v5'/><path d='M12 17v5'/><path d='M2 12h5'/><path d='M17 12h5'/>"
             "<path d='M5 5l3.5 3.5'/><path d='M15.5 15.5 19 19'/><path d='M19 5l-3.5 3.5'/><path d='M8.5 15.5 5 19'/>",
    "sol": "<circle cx='12' cy='12' r='4.5'/><path d='M12 2v2.5'/><path d='M12 19.5V22'/><path d='M2 12h2.5'/>"
           "<path d='M19.5 12H22'/><path d='M4.9 4.9l1.8 1.8'/><path d='M17.3 17.3l1.8 1.8'/><path d='M19.1 4.9l-1.8 1.8'/>"
           "<path d='M6.7 17.3l-1.8 1.8'/>",
    "moneda": "<ellipse cx='12' cy='6' rx='7' ry='3'/><path d='M5 6v6c0 1.7 3.1 3 7 3s7-1.3 7-3V6'/>"
              "<path d='M5 12v6c0 1.7 3.1 3 7 3s7-1.3 7-3v-6'/>",
    "puente": "<path d='M2 18h20'/><path d='M4 18v-5'/><path d='M20 18v-5'/><path d='M2 13c4-5 16-5 20 0'/>"
              "<path d='M9 18v-6'/><path d='M15 18v-6'/>",
    "escudo": "<path d='M12 3l8 3v6c0 4.5-3.4 8.2-8 9-4.6-.8-8-4.5-8-9V6z'/>",
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


# ---- 5.2 Piezas comunes ------------------------------------------------------------------
def panel(id_grafico, titulo, area, alto=380, numero_grafico=None, unidad=None, nota=None, fuente=None,
          leyenda=None):
    """Recuadro con un gráfico:
        Gráfico N.º 1: Título del gráfico     <- en rojo, en una sola línea (id: <id_grafico>-titulo)
        (unidad)                              <- id: <id_grafico>-unidad
        ... gráfico ...
        Nota y fuente"""
    titulo_html = [html.Span(f"Gráfico N.º {numero_grafico}: ", className="panel-numero")] if numero_grafico else []
    titulo_html.append(html.Span(titulo, id=f"{id_grafico}-titulo"))
    hijos = [html.Header(className="panel-cabecera", children=[
        html.H3(titulo_html, className="panel-titulo"),
        html.P(unidad or "", id=f"{id_grafico}-unidad", className="panel-unidad")])]
    if leyenda is not None:
        hijos.append(leyenda)
    hijos.append(html.Div(className="grafico", style={"height": f"{alto}px"}, children=dcc.Graph(
        id=id_grafico, config=CONFIG_GRAFICO, responsive=True, style={"height": "100%"})))
    if nota or fuente:
        hijos.append(html.P(([html.Strong("Nota: "), nota, html.Br()] if nota else []) + [fuente or ""],
                            className="panel-pie"))
    return html.Section(className=f"panel area-{area}", children=hijos)


def tarjeta_kpi(id_base, nombre_icono, titulo="", color=None, destacada=False, alto_grafico=None):
    """Tarjeta de arriba: ícono + título, valor, nota y (opcional) un gráfico pequeño.
    ids: <id_base>-titulo, <id_base>-valor, <id_base>-nota, <id_base>-graf"""
    estilo = {"borderTop": f"4px solid {color}"} if color else {}
    estilo_icono = {"background": hex_a_rgba(color, 0.14), "color": color} if color else {}
    hijos = [html.Div(className="kpi-cabeza", children=[
                 html.Span(icono(nombre_icono), className="kpi-icono", style=estilo_icono),
                 html.Span(titulo, id=f"{id_base}-titulo", className="kpi-titulo")]),
             html.Div(id=f"{id_base}-valor", className="kpi-valor"),
             html.Div(id=f"{id_base}-nota", className="kpi-delta")]
    if alto_grafico:
        hijos.append(dcc.Graph(id=f"{id_base}-graf", config=CONFIG_MINI, style={"height": f"{alto_grafico}px"}))
    return html.Div(className="kpi" + (" kpi-destacado" if destacada else ""), style=estilo, children=hijos)


def delta_texto(actual, previo, etiqueta_previa, decimales=2):
    """▲ +0,05 vs 2024 (rojo si sube, azul si baja)."""
    if es_vacio(actual) or es_vacio(previo):
        return html.Span("Sin dato para comparar", className="delta delta-neutro")
    cambio = round(float(actual) - float(previo), decimales) + 0.0
    if cambio == 0:
        return html.Span(f"= igual que {etiqueta_previa}", className="delta delta-neutro")
    sube = cambio > 0
    return html.Span([html.Span("▲" if sube else "▼", className="delta-icono"),
                      f" {con_signo(cambio, decimales)} vs {etiqueta_previa}"],
                     className="delta " + ("delta-sube" if sube else "delta-baja"))


def cabecera_pagina(antetitulo="", titulo="", subtitulo="", prefijo=None):
    """Encabezado de cada tablero, como la presentación: antetítulo rojo, título azul, subtítulo y línea roja.
    Con prefijo, las partes tienen id (<prefijo>-antetitulo, <prefijo>-titulo, <prefijo>-sub) y cambian con filtros."""
    def ids(parte):
        return {"id": f"{prefijo}-{parte}"} if prefijo else {}
    return html.Header(className="cabecera", children=[
        html.P(antetitulo, className="cabecera-antetitulo", **ids("antetitulo")),
        html.H1(titulo, **ids("titulo")),
        html.P(subtitulo, className="cabecera-sub", **ids("sub"))])


def boton_portada():
    return dcc.Link([icono("volver"), "Volver a la portada"], href="/", className="lateral-volver")


def selector_indices(activo):
    """Botones para pasar del Índice de Desastres al INFORM Risk (los dos índices de desastres)."""
    return html.Nav(className="selector-indices", children=[
        dcc.Link("Índice de Desastres", href=RUTA_ID,
                 className="selector-opcion" + (" activa" if activo == "id" else "")),
        dcc.Link("INFORM Risk", href=RUTA_INFORM,
                 className="selector-opcion" + (" activa" if activo == "inform" else ""))])


# ---- 5.3 Portada --------------------------------------------------------------------------
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


def crear_portada(clima=None, base_id=None, inform=None, latam=None):
    datos = {}
    if clima is not None:
        datos[RUTA_CLIMA] = [html.Span(f"{clima['etiqueta'].iloc[0]} – {clima['etiqueta'].iloc[-1]} · "),
                             html.Span(f"{len(INDICES_CLIMA)} índices mensuales")]
    if base_id is not None:
        ultimo = lista_anios(base_id)[-1]
        partes = [f"ID promedio {ultimo}: {numero(base_id.loc[base_id[COL_ANIO] == ultimo, COL_ID].mean(), 1)}"]
        if inform is not None and PAIS_INICIAL_INFORM in inform.index:
            partes.append(f"INFORM Risk {inform.loc[PAIS_INICIAL_INFORM, 'NOMBRE']}: "
                          f"{numero(inform.loc[PAIS_INICIAL_INFORM, 'INFORM RISK'], 1)}")
        datos[RUTA_ID] = [html.Span(p + (" · " if i < len(partes) - 1 else "")) for i, p in enumerate(partes)]
    if latam is not None:
        e = latam["eventos"]
        datos[RUTA_LATAM] = [html.Span(f"{e['anio'].min()}–{e['anio'].max()} · "),
                             html.Span(f"{e['pais'].nunique()} países · "),
                             html.Span(f"{numero(len(e), 0)} desastres")]
    tarjetas = [tarjeta_seccion(i, s, datos.get(s["ruta"], "En preparación"))
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


# ---- 5.4 Página: Índices climáticos ---------------------------------------------------
def barra_lateral_clima(clima):
    """Portada · año y mes elegidos · filtros (índice, año, mes) · fuente."""
    anios = sorted(clima["anio"].unique())
    a0, a1 = int(anios[0]), int(anios[-1])
    marcas_anio = {a: str(a) for a in range(a0, a1 + 1) if (a - a0) % 4 == 0 and a1 - a >= 3}
    marcas_anio[a1] = str(a1)
    marcas_mes = {m: MESES[m - 1][0] for m in range(1, 13)}  # E F M A M J J A S O N D
    ultimo = clima.iloc[-1]
    return html.Aside(className="lateral", children=[
        dcc.Link([icono("volver"), "Volver a la portada"], href="/", className="lateral-volver"),
        html.Div(className="marca", children=[html.Span("ICM", className="marca-sigla"),
                                              html.Span("Índices climáticos mensuales", className="marca-nombre")]),
        html.Div(className="seleccion", children=[html.Div(id="clima-lateral-anio", className="seleccion-anio"),
                                                  html.Div(id="clima-lateral-mes", className="seleccion-dpto")]),
        html.Div(className="controles", children=[
            html.Label([icono("indicador"), "Índice"], className="control-etiqueta"),
            dcc.RadioItems(id="clima-indice", value=list(INDICES_CLIMA)[0], className="control-radio",
                           options=[{"label": k, "value": k} for k in INDICES_CLIMA]),
            html.Label([icono("calendario"), "Año"], className="control-etiqueta"),
            dcc.Slider(id="clima-anio", min=a0, max=a1, step=1, value=int(ultimo["anio"]), marks=marcas_anio,
                       included=False, tooltip={"placement": "bottom", "always_visible": False}),
            html.Div(className="control-fila", children=[
                html.Label([icono("calendario"), "Mes"], className="control-etiqueta"),
                html.Button("▶ Reproducir", id="clima-btn-play", n_clicks=0, className="boton-play")]),
            dcc.Slider(id="clima-mes", min=1, max=12, step=1, value=int(ultimo["mes"]), marks=marcas_mes,
                       included=False),
            html.Button(f"Ir al último mes ({ultimo['etiqueta']})", id="clima-btn-ultimo",
                        n_clicks=0, className="boton-secundario")]),
        html.P(FUENTE_CLIMA, className="lateral-fuente"),
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
    unidad = cfg.get("unidad", "")
    cabecera = html.Thead(html.Tr([html.Th("Categoría"), html.Th("Magnitud" if varias else ""),
                                   html.Th(f"Intervalo ({unidad})" if unidad else "Intervalo")]))
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


def panel_definicion():
    """Cuadro con la definición del índice y su criterio (tabla). No lleva número: no es un gráfico."""
    return html.Section(className="panel area-definicion", children=[
        html.Header(className="panel-cabecera", children=[
            html.H3([html.Span("Definición y criterio: ", className="panel-numero"),
                     html.Span(id="clima-definicion-titulo")], className="panel-titulo")]),
        html.H4("Definición", className="subtitulo-panel"),
        html.P(id="clima-definicion", className="definicion-texto"),
        html.H4("Criterio", className="subtitulo-panel"),
        html.Div(id="clima-criterio", className="criterio-contenedor"),
        html.P(FUENTE_CLIMA_CORTA, className="panel-pie")])


def crear_layout_clima(clima):
    """título · 4 tarjetas · [Definición y criterio | Gráfico 1] · [Gráfico 2 | Gráfico 3]"""
    primero, ultimo = clima["etiqueta"].iloc[0], clima["etiqueta"].iloc[-1]
    cabecera = cabecera_pagina("Índices climáticos mensuales · Condición del mar frente al Perú", "",
                               f"Cuatro índices mensuales que dicen si el mar está en condición cálida, neutra o "
                               f"fría, {primero} – {ultimo}.", prefijo="clima")
    fila_1 = html.Div(className="fila fila-2", children=[
        panel_definicion(),
        panel("clima-graf-evolucion", "", "evolucion", alto=380, numero_grafico=1, fuente=FUENTE_CLIMA_CORTA,
              nota="rojo: por encima de cero (más cálido de lo normal); azul: por debajo. Líneas punteadas: "
                   "umbrales del criterio.")])
    fila_2 = html.Div(className="fila fila-2", children=[
        panel("clima-graf-estacional", "", "estacional", alto=400, numero_grafico=2, fuente=FUENTE_CLIMA_CORTA,
              nota="cada línea es un año. Azul oscuro: el año elegido; rojo: el año con el valor más alto; "
                   "franja gris: rango neutro."),
        panel("clima-graf-calendario", "", "calendario", alto=400, numero_grafico=3, fuente=FUENTE_CLIMA_CORTA,
              nota="cada celda es un mes, con el color de su categoría.")])
    contenido = html.Main(className="contenido", children=[
        cabecera, html.Section(className="clima-tarjetas", children=[tarjeta_clima(k) for k in INDICES_CLIMA]),
        fila_1, fila_2])
    return html.Div(className="app", children=[barra_lateral_clima(clima), contenido])


# ---- 5.5 Página: Índice de Desastres (ID) ------------------------------------------------
def barra_lateral_id(base):
    anios = lista_anios(base)
    opciones = [{"label": "Todo el Perú (promedio)", "value": NACIONAL}] + [
        {"label": nombre_visible(d), "value": d} for d in sorted(base[COL_DPTO].unique())]
    marcas = {a: str(a) for a in anios if (a - anios[0]) % 5 == 0 and anios[-1] - a >= 3}
    marcas[anios[-1]] = str(anios[-1])
    return html.Aside(className="lateral", children=[
        boton_portada(), selector_indices("id"),
        html.Div(className="marca", children=[html.Span("ID", className="marca-sigla"),
                                              html.Span("Índice de Desastres", className="marca-nombre")]),
        html.Div(className="seleccion", children=[html.Div(id="id-lateral-anio", className="seleccion-anio"),
                                                  html.Div(id="id-lateral-dpto", className="seleccion-dpto")]),
        html.Div(className="controles", children=[
            html.Label([icono("mapa"), "Departamento"], className="control-etiqueta"),
            dcc.Dropdown(id="id-dpto", options=opciones, value=NACIONAL, clearable=False),
            html.Div(className="control-fila", children=[
                html.Label([icono("calendario"), "Año"], className="control-etiqueta"),
                html.Button("▶ Reproducir", id="id-btn-play", n_clicks=0, className="boton-play")]),
            dcc.Slider(id="id-anio", min=anios[0], max=anios[-1], step=1, value=anios[-1], marks=marcas,
                       included=False, tooltip={"placement": "bottom", "always_visible": False}),
            html.Label([icono("indicador"), "Indicador"], className="control-etiqueta"),
            dcc.RadioItems(id="id-indicador", value=COL_ID, className="control-radio",
                           options=[{"label": v["corto"], "value": k} for k, v in INDICADORES_ID.items()]),
            html.Button(f"Volver a Perú, {anios[-1]}", id="id-btn-reiniciar", n_clicks=0,
                        className="boton-secundario")]),
        html.P(FUENTE_ID, className="lateral-fuente"),
        dcc.Interval(id="id-intervalo", interval=MS_POR_ANIO, disabled=True, n_intervals=0),
    ])


def fila_kpis_id():
    """Velocímetro del ID (promedio del Perú o departamento) + frecuencia + los 4 pilares de la severidad."""
    principal = html.Div(className="kpi kpi-destacado kpi-velocimetro", children=[
        html.Div(className="kpi-cabeza", children=[html.Span(icono("medidor"), className="kpi-icono"),
                                                   html.Span(id="id-kpi-0-titulo", className="kpi-titulo")]),
        dcc.Graph(id="id-velocimetro", config=CONFIG_MINI, style={"height": "140px"}),
        html.Div(id="id-kpi-0-nota", className="velocimetro-texto")])
    tarjetas = [tarjeta_kpi("id-kpi-1", "alerta", "Índice de frecuencia", alto_grafico=52)]
    for i, (nombre, p) in enumerate(PILARES.items(), start=2):
        tarjetas.append(tarjeta_kpi(f"id-kpi-{i}", p["icono"], p["tarjeta"], color=p["color"],
                                    alto_grafico=52))
    return html.Section(className="fila-kpis", children=[principal] + tarjetas)


def crear_layout_id(base):
    """título + definición · velocímetro y 5 tarjetas · [G1 ranking | G2 mapa] · [G3 evolución | G4 relación]
       · [G5 mapa de calor]"""
    cabecera = cabecera_pagina("", "Índice de Desastres (ID)", DEFINICION_ID, prefijo="id")
    fila_1 = html.Div(className="fila fila-2", children=[
        panel("id-graf-1", "", "id-1", alto=500, numero_grafico=1, fuente=FUENTE_ID),
        panel("id-graf-2", "", "id-2", alto=500, numero_grafico=2, fuente=FUENTE_ID)])
    fila_2 = html.Div(className="fila fila-2", children=[
        panel("id-graf-3", "", "id-3", alto=420, numero_grafico=3, fuente=FUENTE_ID),
        panel("id-graf-4", "", "id-4", alto=420, numero_grafico=4, fuente=FUENTE_ID,
              nota="cada burbuja es un departamento; su tamaño es el ID y su color, el cuartil. Las curvas "
                   "punteadas unen puntos con el mismo ID.")])
    fila_3 = html.Div(className="fila", children=[
        panel("id-graf-5", "", "id-5", alto=580, numero_grafico=5, fuente=FUENTE_ID,
              nota="arriba, los departamentos que pasan más años en los cuartiles más altos.")])
    contenido = html.Main(className="contenido", children=[cabecera, fila_kpis_id(), fila_1, fila_2, fila_3])
    return html.Div(className="app", children=[barra_lateral_id(base), contenido])


# ---- 5.6 Página: INFORM Risk --------------------------------------------------------------
def barra_lateral_inform(inform):
    paises = sorted(((iso, n) for iso, n in PAISES_INFORM.items() if iso in inform.index), key=lambda x: x[1])
    return html.Aside(className="lateral", children=[
        boton_portada(), selector_indices("inform"),
        html.Div(className="marca", children=[html.Span("IR", className="marca-sigla"),
                                              html.Span("Índice INFORM Risk", className="marca-nombre")]),
        html.Div(className="seleccion", children=[html.Div(id="inf-lateral-pais", className="seleccion-pais"),
                                                  html.Div(str(ANIO_INFORM), className="seleccion-dpto")]),
        html.Div(className="controles", children=[
            html.Label([icono("mapa"), "País"], className="control-etiqueta"),
            dcc.Dropdown(id="inf-pais", options=[{"label": n, "value": iso} for iso, n in paises],
                         value=PAIS_INICIAL_INFORM, clearable=False),
            html.Label([icono("indicador"), "Indicador"], className="control-etiqueta"),
            dcc.RadioItems(id="inf-indicador", value="INFORM RISK", className="control-radio",
                           options=[{"label": v, "value": k} for k, v in INDICADORES_INFORM.items()]),
            html.Button(f"Volver a {PAISES_INFORM[PAIS_INICIAL_INFORM]}", id="inf-btn-reiniciar", n_clicks=0,
                        className="boton-secundario")]),
        html.P(FUENTE_INFORM, className="lateral-fuente"),
    ])


def tarjetas_inform():
    """6 tarjetas: INFORM Risk (con su escala), las 3 dimensiones, amenazas naturales y promedios por grupo."""
    tarjetas = [tarjeta_kpi("inf-kpi-0", "escudo", "INFORM Risk", destacada=True, alto_grafico=40)]
    for i, (col, (texto, color)) in enumerate(DIMENSIONES_INFORM.items(), start=1):
        tarjetas.append(tarjeta_kpi(f"inf-kpi-{i}", {1: "alerta", 2: "casa", 3: "escudo"}[i], texto, color=color))
    tarjetas.append(tarjeta_kpi("inf-kpi-4", "sismo", "Amenazas naturales", color=COLOR_ROJO_PPT))
    tarjetas.append(tarjeta_kpi("inf-kpi-5", "globo", "Promedio de América Latina (21 países)", color="#7193C9"))
    return html.Section(className="kpis-6", children=tarjetas)


def crear_layout_inform(inform):
    """título + definición · 6 tarjetas · [G1 ranking | G2 mapa] · [G3 grupos | G4 amenazas naturales]
       · [G5 tres dimensiones | G6 vulnerabilidad y capacidad]"""
    cabecera = cabecera_pagina("", "Índice INFORM Risk", DEFINICION_INFORM, prefijo="inf")

    def grafico(n, alto, nota=None):
        return panel(f"inf-graf-{n}", "", f"inf-{n}", alto=alto, numero_grafico=n, fuente=FUENTE_INFORM, nota=nota)

    nota_21 = ("21 países: no se incluyen Puerto Rico e Islas Caimán (no tienen puntaje INFORM) ni Jamaica, "
               "Belice y Surinam.")
    fila_1 = html.Div(className="fila fila-2", children=[grafico(1, 480, nota_21), grafico(2, 480)])
    fila_2 = html.Div(className="fila fila-2", children=[
        grafico(3, 380, "INFORM 2026 clasifica al Perú (4,9) en la categoría «Medium» (Medio)."),
        grafico(4, 380, "riesgo de cada amenaza natural en una escala de 0 a 10.")])
    fila_3 = html.Div(className="fila fila-2", children=[
        grafico(5, 500, "INFORM Risk = (Amenaza y exposición)^1/3 × (Vulnerabilidad)^1/3 × "
                        "(Falta de capacidad de respuesta)^1/3."),
        grafico(6, 500, "a mayor valor, mayor vulnerabilidad o menor capacidad de respuesta.")])
    contenido = html.Main(className="contenido", children=[cabecera, tarjetas_inform(), fila_1, fila_2, fila_3])
    return html.Div(className="app", children=[barra_lateral_inform(inform), contenido])


# ---- 5.7 Página: América Latina ------------------------------------------------------------
def barra_lateral_latam(latam):
    anios = latam["eventos"]["anio"]
    a0, a1 = int(anios.min()), int(anios.max())
    marcas = {a: str(a) for a in range(a0, a1 + 1) if (a - a0) % 10 == 0 and a1 - a >= 8}
    marcas[a1] = str(a1)
    paises = sorted(latam["eventos"]["pais"].unique())
    return html.Aside(className="lateral", children=[
        boton_portada(),
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
            html.Label([icono("indicador"), "Indicador"], className="control-etiqueta"),
            dcc.RadioItems(id="latam-indicador", value="eventos", className="control-radio control-radio-grande",
                           options=[{"label": v, "value": k} for k, v in INDICADORES_LATAM.items()]),
            html.Button(f"Volver a {PAIS_INICIAL}, {a0}-{a1}", id="latam-btn-reiniciar", n_clicks=0,
                        className="boton-secundario")]),
        html.P(FUENTE_LATAM, className="lateral-fuente"),
    ])


def tarjetas_latam():
    """6 tarjetas: el total y los 5 tipos de desastre (sus textos cambian con el indicador)."""
    tarjetas = [tarjeta_kpi("latam-kpi-0", "alerta", destacada=True)]
    for i, sub in enumerate(ORDEN_SUBGRUPOS_PAIS, start=1):
        tarjetas.append(tarjeta_kpi(f"latam-kpi-{i}", ICONOS_TIPO[sub], color=COLORES_TIPO[sub]))
    return html.Section(className="kpis-6", children=tarjetas)


def crear_layout_latam(latam):
    """título · 6 tarjetas · [G1 ranking | G2 mapa] · [G3 periodos o correlación | G4 % por tipo]
       · [G5 país por tipo y subtipo | G6 país año por año]"""
    cabecera = cabecera_pagina(prefijo="latam")

    def grafico(n, alto):
        return panel(f"latam-graf-{n}", "", f"latam-{n}", alto=alto, numero_grafico=n, fuente=FUENTE_LATAM)

    fila_1 = html.Div(className="fila fila-2", children=[grafico(1, 500), grafico(2, 500)])
    fila_2 = html.Div(className="fila fila-2", children=[grafico(3, 380), grafico(4, 380)])
    fila_3 = html.Div(className="fila fila-2", children=[grafico(5, 420), grafico(6, 420)])
    contenido = html.Main(className="contenido", children=[
        cabecera, tarjetas_latam(), fila_1, fila_2, fila_3,
        html.Footer(className="pie", children=[html.P(
            "Fuente: EM-DAT (archivo EMDAT_LATAM.xlsx). Costos ajustados por inflación; el % del PBI compara los "
            "daños acumulados del periodo con el PBI del último año disponible. Hay costos hasta 2025.")]),
    ])
    return html.Div(className="app", children=[barra_lateral_latam(latam), contenido])


# ---- 5.8 Estilos CSS ------------------------------------------------------------------------
CSS = """
:root{--pacifico:#0b3a66;--pacifico-medio:#1f6fb2;--niebla:#edf3f9;--azul-suave:#e3eefa;
--superficie:#fff;--borde:#dde6f0;--tinta:#0f2f50;--tinta-2:#3f6488;--tinta-3:#6f8cab;
--rojo:#d62f2f;--azul-ppt:#0b2a5b;--sube:#c93a39;--baja:#2367b8;
--fuente:"Archivo","Segoe UI",Roboto,Helvetica,Arial,sans-serif;
--radio:16px;--espacio:18px;
--sombra:0 1px 2px rgba(11,42,91,.05),0 6px 20px rgba(11,42,91,.07);
--sombra-alta:0 2px 4px rgba(11,42,91,.06),0 14px 32px rgba(11,42,91,.12);
--Dash-Fill-Interactive-Strong:#1f6fb2;--Dash-Text-Strong:#0f2f50}
*{box-sizing:border-box}
html,body{max-width:100%;overflow-x:hidden}
body{margin:0;background:var(--niebla);color:var(--tinta);font-family:var(--fuente);font-size:15px;line-height:1.5}
h1,h2,h3,h4,p{margin:0}
.icono{display:inline-block;width:18px;height:18px;flex:none;background-color:currentColor;
-webkit-mask:var(--i) center/contain no-repeat;mask:var(--i) center/contain no-repeat}
.app{display:grid;grid-template-columns:288px minmax(0,1fr);min-height:100vh}
.contenido{padding:26px 30px 40px;display:flex;flex-direction:column;gap:var(--espacio);min-width:0}

/* Barra lateral (todos los tableros) */
.lateral{position:sticky;top:0;height:100vh;overflow-y:auto;color:#fff;
background:linear-gradient(180deg,#0b3a66 0%,#0a2f54 100%);
padding:24px 22px 22px;display:flex;flex-direction:column;gap:16px}
.lateral-volver{display:flex;align-items:center;justify-content:center;gap:9px;color:#fff;font-size:15px;
font-weight:650;text-decoration:none;padding:10px 14px;border-radius:12px;background:rgba(255,255,255,.14);
border:1px solid rgba(255,255,255,.55);transition:background .2s ease}
.lateral-volver:hover{background:rgba(255,255,255,.40)}
.lateral-volver .icono{width:19px;height:19px}
.selector-indices{display:grid;grid-template-columns:1fr 1fr;gap:4px;padding:4px;border-radius:12px;
background:rgba(255,255,255,.1);border:1px solid rgba(255,255,255,.2)}
.selector-opcion{text-align:center;font-size:13px;font-weight:650;line-height:1.25;color:rgba(255,255,255,.85);
text-decoration:none;padding:8px 6px;border-radius:9px;transition:background .2s ease}
.selector-opcion:hover{background:rgba(255,255,255,.14)}
.selector-opcion.activa{background:#fff;color:var(--azul-ppt);box-shadow:0 2px 8px rgba(0,0,0,.18)}
.marca{display:flex;align-items:baseline;gap:10px}
.marca-sigla{font-weight:800;font-stretch:125%;font-size:15px;padding:2px 7px;border:1.5px solid rgba(255,255,255,.55);border-radius:6px}
.marca-nombre{font-size:13px;color:rgba(255,255,255,.78)}
.seleccion-anio{font-size:80px;line-height:.9;font-weight:800;font-stretch:125%;letter-spacing:-.035em;font-variant-numeric:tabular-nums}
.seleccion-pais{font-size:44px;line-height:1;font-weight:800;font-stretch:118%;letter-spacing:-.02em}
.seleccion-dpto{margin-top:8px;font-size:24px;font-weight:500;color:#9fcbf2}
.controles{background:var(--superficie);color:var(--tinta);border-radius:var(--radio);padding:16px 16px 18px;
display:flex;flex-direction:column;gap:8px;box-shadow:0 8px 24px rgba(0,0,0,.18)}
.control-etiqueta{display:flex;align-items:center;gap:7px;font-size:13px;font-weight:650;color:var(--pacifico);margin-top:6px}
.control-etiqueta .icono{width:16px;height:16px;color:var(--pacifico-medio)}
.control-fila{display:flex;align-items:center;justify-content:space-between;margin-top:6px}
.control-fila .control-etiqueta{margin-top:0}
.control-radio{display:flex;flex-direction:column;gap:4px;font-size:14px}
.control-radio label{display:flex;align-items:center;gap:8px;cursor:pointer;color:var(--tinta)!important}
.control-radio input{accent-color:var(--pacifico-medio)}
.control-radio-grande{gap:6px;font-size:14.5px;font-weight:600}
.boton-play,.boton-secundario{font-family:var(--fuente);font-size:13px;font-weight:600;border-radius:999px;cursor:pointer}
.boton-play{border:0;background:var(--pacifico-medio);color:#fff;padding:5px 12px}
.boton-play:hover{background:var(--pacifico)}
.boton-secundario{margin-top:8px;border:1px solid var(--borde);background:var(--superficie);color:var(--tinta);padding:7px 14px}
.boton-secundario:hover{background:var(--niebla)}
.lateral-fuente{margin-top:auto;font-size:12.5px;line-height:1.45;color:rgba(255,255,255,.75);
border-top:1px solid rgba(255,255,255,.2);padding-top:12px}
.rc-slider-mark-text{font-size:11px;color:var(--tinta-3)!important}
.rc-slider-handle{border-color:var(--pacifico-medio)!important}

/* Encabezado de los tableros (como la presentación) */
.cabecera{padding-bottom:14px;border-bottom:2px solid var(--rojo)}
.cabecera-antetitulo{font-size:12.5px;font-weight:750;letter-spacing:.12em;text-transform:uppercase;color:var(--rojo);
margin-bottom:6px}
.cabecera-antetitulo:empty{display:none}
.cabecera h1{font-size:32px;line-height:1.12;font-weight:750;font-stretch:112%;letter-spacing:-.015em;color:var(--azul-ppt)}
.cabecera-sub{margin-top:6px;font-size:15px;color:var(--tinta-2);max-width:120ch}

/* Tarjetas de arriba (KPI) */
.fila-kpis{display:grid;gap:14px;grid-template-columns:minmax(0,1.35fr) repeat(5,minmax(0,1fr))}
.kpis-6{display:grid;gap:14px;grid-template-columns:repeat(6,minmax(0,1fr))}
.kpi{background:var(--superficie);border:1px solid var(--borde);border-radius:var(--radio);padding:14px 16px 12px;min-width:0;
display:flex;flex-direction:column;box-shadow:var(--sombra)}
.kpi .dash-graph{margin-top:auto}
.kpi.kpi-destacado{border:2px solid var(--pacifico-medio);background:linear-gradient(180deg,#f3f8fd 0%,#fff 60%)}
.kpi-velocimetro{padding:12px 12px 8px}
.kpi-cabeza{display:flex;align-items:center;gap:10px;min-height:2.9em}
.kpi-icono{width:38px;height:38px;border-radius:11px;background:var(--azul-suave);color:var(--pacifico-medio);
display:grid;place-items:center;flex:none}
.kpi-icono .icono{width:21px;height:21px}
.kpi-titulo{font-size:13px;line-height:1.3;color:var(--tinta-2);font-weight:550}
.kpi-valor{margin-top:6px;font-size:28px;font-weight:700;font-stretch:110%;line-height:1.2;font-variant-numeric:tabular-nums;
white-space:nowrap;color:var(--azul-ppt)}
.kpi-unidad{font-size:14px;font-weight:600;color:var(--tinta-2)}
.kpi-delta{font-size:12.5px;line-height:1.4;min-height:19px;color:var(--tinta-2)}
.delta-sube{color:var(--sube)}.delta-baja{color:var(--baja)}.delta-neutro{color:var(--tinta-3)}
.delta-icono{font-size:10px}
.velocimetro-texto{font-size:12px;line-height:1.35;color:var(--tinta-2);text-align:center;padding:0 2px}
.clase-chip{display:inline-block;font-size:12px;font-weight:700;padding:1px 9px;border-radius:999px;margin-left:6px;
vertical-align:middle}

/* Paneles con gráfico */
.panel{background:var(--superficie);border:1px solid var(--borde);border-radius:var(--radio);padding:16px 18px 12px;min-width:0;
box-shadow:var(--sombra);transition:box-shadow .2s ease}
.panel:hover{box-shadow:var(--sombra-alta)}
.panel-cabecera{padding-bottom:9px;margin-bottom:8px;border-bottom:1px solid var(--borde)}
.panel-titulo{font-size:16.5px;font-weight:650;line-height:1.3;color:var(--rojo)}
.panel-numero{font-weight:800}
.panel-unidad{margin-top:2px;font-size:12.5px;color:var(--tinta-3)}
.panel-unidad:empty{display:none}
.panel-pie{margin-top:6px;padding-top:7px;border-top:1px dashed var(--borde);font-size:11.5px;line-height:1.45;
color:var(--tinta-3)}
.panel-pie strong{color:var(--tinta-2)}
.grafico{width:100%}
.fila{display:grid;gap:var(--espacio);grid-template-columns:minmax(0,1fr)}
.fila-2{grid-template-columns:repeat(2,minmax(0,1fr))}
.pie{color:var(--tinta-3);font-size:12.5px}

/* Portada */
.portada{min-height:100vh;display:flex;flex-direction:column}
.portada-hero{position:relative;overflow:hidden;color:#fff;padding:80px 0 190px;min-height:74vh;
background:#0b3a66 url("/portada.jpg") center 35%/cover no-repeat}
.portada-hero::before{content:"";position:absolute;inset:0;pointer-events:none;
background:linear-gradient(180deg,rgba(237,243,249,0) 62%,#edf3f9 100%),
linear-gradient(100deg,rgba(6,30,54,.88) 0%,rgba(6,30,54,.55) 42%,rgba(6,30,54,0) 72%)}
.portada-hero h1,.portada-sub{text-shadow:0 2px 14px rgba(0,0,0,.35)}
.portada-interior,.portada-tarjetas,.portada-pie{width:100%;max-width:1400px;margin-left:auto;margin-right:auto}
.portada-interior{position:relative;z-index:1;padding:0 48px}
.portada-antetitulo{font-size:13px;font-weight:650;letter-spacing:.09em;text-transform:uppercase;color:#9fcbf2}
.portada h1{margin-top:14px;font-size:54px;line-height:1.04;font-weight:800;font-stretch:118%;letter-spacing:-.025em;max-width:20ch}
.portada-sub{margin-top:18px;font-size:18px;line-height:1.5;color:rgba(255,255,255,.82);max-width:58ch}
.portada-tarjetas{position:relative;z-index:2;margin-top:-150px;padding:0 48px;display:grid;gap:16px;
grid-template-columns:repeat(4,minmax(0,1fr))}
/* Tarjetas planas: ícono y título arriba, texto, y abajo el dato y el botón */
.tarjeta{display:grid;grid-template-columns:44px minmax(0,1fr);grid-template-rows:auto 1fr auto auto;column-gap:12px;row-gap:6px;
align-items:center;background:rgba(248,251,255,.93);-webkit-backdrop-filter:blur(10px);backdrop-filter:blur(10px);
border:1px solid rgba(255,255,255,.9);border-radius:18px;padding:16px 18px;color:var(--tinta);text-decoration:none;
box-shadow:0 8px 24px rgba(11,58,102,.12);transition:transform .2s ease,box-shadow .2s ease,background .2s ease}
.tarjeta:hover{transform:translateY(-3px);background:rgba(240,249,245,.97);box-shadow:0 14px 32px rgba(11,58,102,.18)}
.tarjeta:focus-visible{outline:3px solid #9fcbf2;outline-offset:3px}
.tarjeta-cabeza{display:flex;align-items:center;justify-content:space-between;gap:12px}
.tarjeta .tarjeta-cabeza{grid-column:1;grid-row:1}
.tarjeta-icono{width:44px;height:44px;border-radius:13px;background:var(--azul-suave);color:var(--pacifico-medio);
display:grid;place-items:center;flex:none}
.tarjeta-icono .icono{width:24px;height:24px}
.tarjeta .tarjeta-numero,.tarjeta .etiquetas{display:none}
.tarjeta .estado{display:none}
.tarjeta-pendiente .estado{display:inline-flex;grid-column:1/3;grid-row:3;justify-self:start}
.tarjeta h2{grid-column:2;grid-row:1;font-size:17.5px;line-height:1.2;font-weight:700;font-stretch:108%;color:var(--pacifico)}
.tarjeta .tarjeta-texto{grid-column:1/3;align-self:start;font-size:13px;line-height:1.4;color:var(--tinta-2)}
.tarjeta-numero{font-size:28px;font-weight:800;font-stretch:118%;color:var(--borde);font-variant-numeric:tabular-nums}
.estado{align-self:flex-start;display:inline-flex;align-items:center;font-size:12px;font-weight:650;padding:3px 10px;border-radius:999px}
.estado-listo{background:#e2f3ea;color:#0f6b48}
.estado-pendiente{background:#fff2dc;color:#8a5200}
.tarjeta-texto{font-size:14.5px;color:var(--tinta-2)}
.etiquetas{list-style:none;margin:0;padding:0;display:flex;flex-wrap:wrap;gap:6px}
.etiquetas li{font-size:12px;font-weight:600;color:var(--pacifico-medio);background:var(--niebla);
border:1px solid var(--borde);border-radius:6px;padding:2px 8px}
.tarjeta-dato{grid-column:1/3;min-width:0;padding-top:8px;border-top:1px solid rgba(15,47,80,.12);
font-size:12px;line-height:1.35;color:var(--tinta-2);font-variant-numeric:tabular-nums}
.tarjeta-pendiente .tarjeta-dato{display:none}
.tarjeta-boton{grid-column:1/3;justify-self:start;display:inline-flex;align-items:center;justify-content:center;gap:6px;
border-radius:999px;padding:7px 14px;font-size:13px;font-weight:650;background:#1f8a70;color:#fff;
border:1.5px solid #1f8a70;transition:background .2s ease;white-space:nowrap}
.tarjeta:hover .tarjeta-boton{background:#176b57;border-color:#176b57}
.tarjeta-pendiente .tarjeta-boton{background:transparent;color:var(--pacifico-medio);border-color:var(--pacifico-medio)}
.tarjeta-boton .icono{width:16px;height:16px;transition:transform .2s ease}
.tarjeta:hover .tarjeta-boton .icono{transform:translateX(3px)}
.portada-pie{padding:22px 48px 32px;color:var(--tinta-3);font-size:13px}

/* Índices climáticos */
.clima-tarjetas{display:grid;gap:14px;grid-template-columns:repeat(4,minmax(0,1fr))}
.clima-tarjeta{background:var(--superficie);border:2px solid var(--borde);border-radius:var(--radio);
padding:14px 14px 12px;min-width:0;cursor:pointer;display:flex;flex-direction:column;gap:4px;box-shadow:var(--sombra);
transition:border-color .2s ease,box-shadow .2s ease}
.clima-tarjeta:hover{box-shadow:var(--sombra-alta)}
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
.subtitulo-panel{margin-top:12px;font-size:12px;font-weight:700;letter-spacing:.06em;text-transform:uppercase;
color:var(--pacifico-medio)}
.definicion-texto{margin-top:6px;font-size:14px;line-height:1.55;color:var(--tinta)}
.criterio-contenedor{margin-top:6px;overflow-x:auto}
.criterio{width:100%;border-collapse:collapse;font-size:13.5px}
.criterio th{text-align:left;font-size:12px;font-weight:700;color:var(--pacifico);text-transform:uppercase;
letter-spacing:.04em;padding:7px 8px;border-bottom:2px solid var(--borde)}
.criterio td{padding:6px 8px;border-bottom:1px solid var(--borde);color:var(--tinta);vertical-align:middle}
.criterio td.num{text-align:right;font-variant-numeric:tabular-nums;white-space:nowrap}
.criterio-grupo{font-weight:600;color:var(--tinta-2)}
.criterio-color{display:inline-block;width:22px;height:16px;border-radius:3px;border:1px solid #9aa9b8;
vertical-align:middle;margin-right:8px}

/* Páginas en preparación */
.pagina-simple{max-width:760px;margin:0 auto;padding:40px 24px 56px}
.enlace-volver{display:inline-flex;align-items:center;gap:8px;color:var(--pacifico-medio);font-size:14px;
font-weight:600;text-decoration:none}
.enlace-volver:hover{color:var(--pacifico)}
.enlace-volver .icono{width:17px;height:17px}
.aviso{margin-top:18px;background:var(--superficie);border:1px solid var(--borde);border-radius:18px;
padding:32px;display:flex;flex-direction:column;gap:14px;box-shadow:var(--sombra)}
.aviso h1{font-size:30px;line-height:1.2;font-weight:700;font-stretch:108%;color:var(--pacifico)}
.aviso-nota{display:flex;align-items:center;gap:12px;margin-top:8px;padding:14px 16px;border-radius:12px;
background:var(--niebla);color:var(--tinta-2);font-size:14px}
.aviso-nota-icono{color:var(--pacifico-medio);display:grid;place-items:center}

/* Pantallas medianas */
@media (max-width:1280px){
.clima-tarjetas{grid-template-columns:repeat(2,minmax(0,1fr))}
.kpis-6,.fila-kpis{grid-template-columns:repeat(3,minmax(0,1fr))}
.portada-tarjetas{grid-template-columns:repeat(2,minmax(0,1fr))}}

/* Celulares */
@media (max-width:900px){
.app{grid-template-columns:minmax(0,1fr)}
.lateral{position:static;height:auto;padding:18px 16px;gap:12px}
.seleccion{display:flex;align-items:baseline;gap:12px;flex-wrap:wrap}
.seleccion-anio{font-size:54px}
.seleccion-pais{font-size:34px}
.seleccion-dpto{margin-top:0;font-size:22px}
.contenido{padding:18px 12px 28px;gap:14px}
.cabecera h1{font-size:24px}
.cabecera-sub{font-size:13.5px}
.kpis-6,.fila-kpis{grid-template-columns:repeat(2,minmax(0,1fr));gap:10px}
.fila-kpis .kpi-velocimetro{grid-column:span 2}
.kpi{padding:12px 12px 8px}
.kpi-icono{width:32px;height:32px}
.kpi-icono .icono{width:18px;height:18px}
.kpi-titulo{font-size:12px}
.kpi-valor{font-size:22px}
.fila-2{grid-template-columns:minmax(0,1fr)}
.panel{padding:12px 10px 8px}
.panel-titulo{font-size:15px}
.area-id-5 .grafico{height:640px!important}
.portada-hero{padding:40px 0 130px;min-height:0}
.portada-interior{padding:0 16px}
.portada h1{font-size:32px}
.portada-sub{font-size:15.5px}
.portada-tarjetas{grid-template-columns:minmax(0,1fr);padding:0 12px;margin-top:-90px;gap:12px}
.tarjeta{padding:14px 14px}
.portada-pie{padding:22px 16px 28px}
.pagina-simple{padding:24px 12px 40px}
.aviso{padding:22px 18px}
.aviso h1{font-size:24px}
.clima-tarjetas{grid-template-columns:minmax(0,1fr);gap:10px}}
"""

CSS = CSS + css_iconos()

PAGINA = ("<!DOCTYPE html><html lang='es'><head>{%metas%}<title>{%title%}</title>{%favicon%}{%css%}"
          "<style>" + CSS + "</style></head><body>{%app_entry%}<footer>{%config%}{%scripts%}"
          "{%renderer%}</footer></body></html>")


# =============================================================================
# 6. INTERACTIVIDAD (callbacks)
# =============================================================================
# ---- 6.1 Navegación entre páginas ---------------------------------------------------
def registrar_navegacion(app, clima, base_id, inform, latam):
    """Muestra la portada o el tablero según la dirección de la página."""
    secciones = {s["ruta"]: s for s in SECCIONES}

    @app.callback(Output("pagina", "children"), Input("url", "pathname"))
    def mostrar_pagina(ruta):
        ruta = (ruta or "/").rstrip("/") or "/"
        if ruta == RUTA_CLIMA and clima is not None:
            return crear_layout_clima(clima)
        if ruta in (RUTA_ID, "/ird") and base_id is not None:
            return crear_layout_id(base_id)
        if ruta == RUTA_INFORM and inform is not None:
            return crear_layout_inform(inform)
        if ruta == RUTA_LATAM and latam is not None:
            return crear_layout_latam(latam)
        if ruta in secciones:
            return pagina_en_preparacion(secciones[ruta])
        return crear_portada(clima, base_id, inform, latam)

    # Al cambiar de página, empezar desde arriba
    app.clientside_callback("function(ruta){window.scrollTo(0, 0); return ruta;}",
                            Output("arriba", "data"), Input("url", "pathname"))


# ---- 6.2 Índices climáticos ------------------------------------------------------------
def registrar_callbacks_clima(app, clima):
    claves = list(INDICES_CLIMA)
    ultimo = len(clima) - 1
    posicion = {(a, m): i for i, (a, m) in enumerate(zip(clima["anio"], clima["mes"]))}
    por_fecha = {f.strftime("%Y-%m"): i for i, f in enumerate(clima["fecha"])}
    primero_txt, ultimo_txt = clima["etiqueta"].iloc[0], clima["etiqueta"].iloc[-1]
    a0, a1 = int(clima["anio"].min()), int(clima["anio"].max())

    def ubicar(anio, mes):
        """Fila del mes elegido; si ese mes aún no tiene dato, el último mes con dato de ese año."""
        if (anio, mes) in posicion:
            return posicion[(anio, mes)]
        meses = clima.index[clima["anio"] == anio]
        return int(meses[-1]) if len(meses) else ultimo

    def a_filtros(i):
        return int(clima.loc[i, "anio"]), int(clima.loc[i, "mes"])

    # Tarjetas, clics en gráficos, reproducir, último mes y año/mes -> índice, año y mes
    @app.callback(Output("clima-indice", "value"), Output("clima-anio", "value"), Output("clima-mes", "value"),
                  [Input(f"clima-tarjeta-{k}", "n_clicks") for k in claves]
                  + [Input("clima-graf-calendario", "clickData"), Input("clima-graf-evolucion", "clickData"),
                     Input("clima-graf-estacional", "clickData"), Input("clima-intervalo", "n_intervals"),
                     Input("clima-btn-ultimo", "n_clicks"), Input("clima-anio", "value"), Input("clima-mes", "value")],
                  prevent_initial_call=True)
    def sincronizar_clima(*args):
        anio, mes = args[-2], args[-1]
        origen = ctx.triggered_id
        clic = ctx.triggered[0]["value"] if ctx.triggered else None
        puntos = (clic or {}).get("points") if isinstance(clic, dict) else None
        punto = puntos[0] if puntos else {}

        def ir_a(i):
            return (no_update,) + a_filtros(i) if i is not None else (no_update, no_update, no_update)

        if isinstance(origen, str) and origen.startswith("clima-tarjeta-"):
            return origen.replace("clima-tarjeta-", ""), no_update, no_update
        if origen == "clima-btn-ultimo":
            return ir_a(ultimo)
        if origen == "clima-intervalo":
            i = ubicar(anio, mes)
            return ir_a(i + 1 if i < ultimo else 0)
        if origen == "clima-graf-calendario":
            return ir_a(posicion.get((punto.get("y"), punto.get("x"))))
        if origen == "clima-graf-evolucion":
            return ir_a(por_fecha.get(str(punto.get("x", ""))[:7]))
        if origen == "clima-graf-estacional":
            custom = punto.get("customdata") or [None]
            return ir_a(posicion.get((custom[0], punto.get("x"))))
        if origen in ("clima-anio", "clima-mes") and (anio, mes) not in posicion:
            return ir_a(ubicar(anio, mes))  # mes sin dato: se pasa al último mes disponible
        return no_update, no_update, no_update

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

    @app.callback(salidas, Input("clima-anio", "value"), Input("clima-mes", "value"), Input("clima-indice", "value"))
    def actualizar_tarjetas_clima(anio, mes, clave_sel):
        idx = ubicar(anio, mes)
        fila = clima.loc[idx]
        previa = clima.loc[idx - 1] if idx > 0 else None
        nombre_mes = MESES[fila["mes"] - 1]
        salida = [str(fila["anio"]), nombre_mes, f"Índices climáticos, {nombre_mes.lower()} de {fila['anio']}"]
        for k in claves:
            cfg = INDICES_CLIMA[k]
            valor, dec, unidad = fila[k], cfg["decimales"], cfg.get("unidad", "")
            cat = clasificar_clima(valor, cfg)
            if cat is None:
                valor_html = html.Span(f"Sin dato en {fila['etiqueta']}")
                estado = html.Span("Sin dato")
            else:
                seg = segmentos_clima(cfg)[cat]
                valor_html = ["Valor preliminar" if fila[k + "_prelim"] else "Valor",
                              html.Sub(f" {fila['etiqueta']}"), " = ", html.Strong(valor_con_unidad(valor, cfg))]
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
                                       f" {con_signo(cambio, dec)} {unidad} vs {previa['etiqueta']}"],
                                      className="delta " + ("delta-sube" if sube else "delta-baja"))
            salida += [grafico_velocimetro_clima(cfg, valor), valor_html, estado, delta,
                       "clima-tarjeta activa" if k == clave_sel else "clima-tarjeta"]
        return salida

    # Definición y criterio + Gráficos N.º 1, 2 y 3 (figura, título y unidad)
    @app.callback(Output("clima-definicion-titulo", "children"), Output("clima-definicion", "children"),
                  Output("clima-criterio", "children"),
                  [Output(f"clima-graf-{g}{parte}", prop) for g in ("evolucion", "estacional", "calendario")
                   for parte, prop in (("", "figure"), ("-titulo", "children"), ("-unidad", "children"))],
                  Input("clima-anio", "value"), Input("clima-mes", "value"), Input("clima-indice", "value"))
    def actualizar_graficos_clima(anio, mes, clave):
        idx = ubicar(anio, mes)
        cfg = INDICES_CLIMA[clave]
        unidad = f"({cfg['unidad']})" if cfg.get("unidad") else ""
        return (f"{clave}: {cfg['nombre']}", cfg.get("definicion", ""), tabla_criterio(clave),
                grafico_evolucion_clima(clima, clave, idx),
                f"Evolución del {clave}, {primero_txt} – {ultimo_txt}", unidad,
                grafico_estacional_clima(clima, clave, idx),
                f"{clave} de cada año, mes a mes, {a0}-{a1}", unidad,
                grafico_calendario_clima(clima, clave, idx),
                f"Calendario del {clave}, {a0}-{a1}", f"{unidad} · color según el criterio".strip(" ·"))


# ---- 6.3 Índice de Desastres (ID) --------------------------------------------------------
def registrar_callbacks_id(app, base, geojson):
    anios = lista_anios(base)
    centros = centroides(geojson)
    serie_nacional = serie_id(base, NACIONAL)
    departamentos = set(base[COL_DPTO])
    por_nombre = {nombre_visible(d): d for d in departamentos}
    titulos = {COL_ID: ("ID del Perú", "del ID del Perú", "del Índice de Desastres"),
               COL_FREC: ("Índice de frecuencia", "del índice de frecuencia", "del índice de frecuencia"),
               COL_SEV: ("Severidad global", "de la severidad global", "de la severidad global")}

    # Clics en los gráficos, reproducir y reiniciar -> departamento y año
    @app.callback(Output("id-dpto", "value"), Output("id-anio", "value"),
                  Input("id-graf-1", "clickData"), Input("id-graf-2", "clickData"),
                  Input("id-graf-3", "clickData"), Input("id-graf-4", "clickData"),
                  Input("id-graf-5", "clickData"), Input("id-intervalo", "n_intervals"),
                  Input("id-btn-reiniciar", "n_clicks"),
                  State("id-dpto", "value"), State("id-anio", "value"), prevent_initial_call=True)
    def sincronizar_id(*args):
        dpto_actual, anio_actual = args[-2], args[-1]
        origen = ctx.triggered_id
        clic = ctx.triggered[0]["value"] if ctx.triggered else None
        puntos = (clic or {}).get("points") if isinstance(clic, dict) else None
        punto = puntos[0] if puntos else {}

        def alternar(dpto):  # segundo clic en el mismo departamento = volver a Perú
            if dpto not in departamentos:
                return no_update
            return NACIONAL if dpto == dpto_actual else dpto

        def anio_valido(valor):
            return int(valor) if valor in anios else no_update

        if origen == "id-btn-reiniciar":
            return NACIONAL, anios[-1]
        if origen == "id-intervalo":
            return no_update, (anio_actual + 1 if anio_actual < anios[-1] else anios[0])
        if origen == "id-graf-1":
            return alternar(punto.get("y")), no_update
        if origen == "id-graf-2":
            return alternar(punto.get("location")), no_update
        if origen == "id-graf-3":
            return no_update, anio_valido(punto.get("x"))
        if origen == "id-graf-4":
            custom = punto.get("customdata")
            return (alternar(custom[0]) if custom else no_update), no_update
        if origen == "id-graf-5":
            return alternar(por_nombre.get(punto.get("y"))), anio_valido(punto.get("x"))
        return no_update, no_update

    @app.callback(Output("id-intervalo", "disabled"), Output("id-btn-play", "children"),
                  Input("id-btn-play", "n_clicks"), State("id-intervalo", "disabled"), prevent_initial_call=True)
    def alternar_reproduccion_id(_n, desactivado):
        reproducir = bool(desactivado)
        return (not reproducir), ("❚❚ Pausar" if reproducir else "▶ Reproducir")

    # Barra lateral, antetítulo, velocímetro y las 5 tarjetas
    salidas = [Output("id-lateral-anio", "children"), Output("id-lateral-dpto", "children"),
               Output("id-antetitulo", "children"), Output("id-kpi-0-titulo", "children"),
               Output("id-velocimetro", "figure"), Output("id-kpi-0-nota", "children")]
    salidas += [Output(f"id-kpi-{i}-{parte}", prop) for i in range(1, 6)
                for parte, prop in (("valor", "children"), ("nota", "children"), ("graf", "figure"))]

    @app.callback(salidas, Input("id-dpto", "value"), Input("id-anio", "value"))
    def actualizar_resumen_id(dpto, anio):
        serie = serie_id(base, dpto)
        corte = corte_id(base, anio, COL_ID)
        cortes = clases_cuartil(corte[COL_ID], METODO_CUARTIL_ID)[1]
        nacional = dpto == NACIONAL
        lugar = "Perú" if nacional else nombre_visible(dpto)
        valor, previo = serie[COL_ID].get(anio), serie[COL_ID].get(anio - 1)
        frec, sev = serie[COL_FREC].get(anio), serie[COL_SEV].get(anio)
        formula = f"Frecuencia {numero(frec, 2)} × Severidad global {numero(sev, 2)}"
        if nacional:
            titulo = f"ID promedio del Perú, {anio}"
            nota = f"Promedio de los {len(corte)} departamentos · {formula}"
            referencia = None
        else:
            fila = corte.set_index(COL_DPTO).loc[dpto]
            titulo = f"ID de {lugar}, {anio}"
            referencia = serie_nacional.loc[anio, COL_ID]
            nota = (f"Puesto {int(fila[col_puesto(COL_ID)])} de {len(corte)} · Cuartil {int(fila[col_cuartil(COL_ID)])}"
                    f" · {formula}. Marca azul: promedio del Perú ({numero(referencia, 1)})")
        antetitulo = f"Índice 2: Índice de Desastres (ID) del Perú, {anio}" + ("" if nacional else f" · {lugar}")
        salida = [str(anio), lugar, antetitulo, titulo, grafico_velocimetro_id(valor, previo, cortes, referencia), nota]
        columnas = [COL_FREC] + [p["columna"] for p in PILARES.values()]
        for i, col in enumerate(columnas, start=1):
            actual, anterior = serie[col].get(anio), serie[col].get(anio - 1)
            nota = [delta_texto(actual, anterior, str(anio - 1))]
            if col == COL_FREC:
                nota += [html.Br(), f"{numero(serie[COL_INCIDENCIAS].get(anio), 0)} incidencias registradas"]
            else:
                nota += [html.Br(), next(p["detalle"] for p in PILARES.values() if p["columna"] == col)]
            salida += [numero(actual, 2), nota, grafico_sparkline(serie[col], anio, 2)]
        return salida

    # Los 5 gráficos numerados (figura, título y unidad)
    @app.callback([Output(f"id-graf-{n}{parte}", prop) for n in range(1, 6)
                   for parte, prop in (("", "figure"), ("-titulo", "children"), ("-unidad", "children"))],
                  Input("id-dpto", "value"), Input("id-anio", "value"), Input("id-indicador", "value"))
    def actualizar_graficos_id(dpto, anio, col):
        corte = corte_id(base, anio, col)
        info = INDICADORES_ID[col]
        nombre, del_mapa, del_indice = titulos[col]
        nacional = dpto == NACIONAL
        periodo = f"{anios[0]}-{anios[-1]}"
        if nacional:
            titulo_3 = f"Evolución {del_indice} promedio, {periodo}"
        else:
            titulo_3 = f"Evolución {del_indice}: {nombre_visible(dpto)} frente al promedio del Perú, {periodo}"
        return (grafico_ranking_id(corte, dpto, col), f"{nombre} por departamento, {anio}", info["unidad"],
                grafico_mapa_id(corte, geojson, centros, dpto, col), f"Mapa {del_mapa}, {anio}", info["unidad"],
                grafico_evolucion_id(serie_nacional, anio, col, None if nacional else serie_id(base, dpto),
                                     nombre_visible(dpto)), titulo_3, info["unidad"],
                grafico_relacion_id(base, dpto, anio, col), f"Relación entre frecuencia y severidad global, {anio}",
                "(Índices de 0 a 1)",
                grafico_calor_id(base, dpto, anio, col),
                f"{nombre} por departamento y año, {periodo}", "(Cuartil del año)")


# ---- 6.4 INFORM Risk ------------------------------------------------------------------------
def registrar_callbacks_inform(app, inform):
    alc = tabla_alc(inform)
    por_nombre = {n: iso for iso, n in zip(alc.index, alc["NOMBRE"])}
    grupos = {g: [i for i in paises if i in inform.index] for g, paises in GRUPOS_INFORM.items()}

    def puesto_alc(col, iso):
        return int((alc[col] > alc.loc[iso, col]).sum()) + 1

    # Clic en el ranking, el mapa o el gráfico de dimensiones -> país; botón -> Perú
    @app.callback(Output("inf-pais", "value"),
                  Input("inf-graf-1", "clickData"), Input("inf-graf-2", "clickData"),
                  Input("inf-graf-5", "clickData"), Input("inf-btn-reiniciar", "n_clicks"),
                  prevent_initial_call=True)
    def sincronizar_inform(*_):
        origen = ctx.triggered_id
        if origen == "inf-btn-reiniciar":
            return PAIS_INICIAL_INFORM
        clic = ctx.triggered[0]["value"] if ctx.triggered else None
        punto = ((clic or {}).get("points") or [{}])[0]
        iso = punto.get("location") if origen == "inf-graf-2" else por_nombre.get(punto.get("y"))
        return iso if iso in alc.index else no_update

    # Barra lateral, antetítulo y las 6 tarjetas
    @app.callback([Output("inf-lateral-pais", "children"), Output("inf-antetitulo", "children"),
                   Output("inf-kpi-0-graf", "figure")]
                  + [Output(f"inf-kpi-{i}-{parte}", "children") for i in range(6) for parte in ("titulo", "valor", "nota")],
                  Input("inf-pais", "value"))
    def actualizar_resumen_inform(iso):
        fila = inform.loc[iso]
        valor = fila["INFORM RISK"]
        clase, color = clase_inform(valor)
        salida = [fila["NOMBRE"], f"Índice 1: INFORM Risk — {fila['NOMBRE']} en América Latina y el Caribe, {ANIO_INFORM}",
                  grafico_escala_inform(valor)]
        chip = html.Span(clase, className="clase-chip",
                         style={"background": color, "color": "#fff" if clase in ("Alto", "Muy alto") else C["tinta"]})
        salida += [f"INFORM Risk de {fila['NOMBRE']}", [numero(valor, 1), chip],
                   f"Puesto {puesto_alc('INFORM RISK', iso)} de {len(alc)} en la región · "
                   f"{int(fila['Rank'])} de {len(inform)} en el mundo"]
        for col, (texto, _) in DIMENSIONES_INFORM.items():
            salida += [texto, numero(fila[col], 1),
                       f"Puesto {puesto_alc(col, iso)} de {len(alc)} · promedio de la región: {numero(alc[col].mean(), 1)}"]
        amenazas = pd.Series({n: fila[c] for c, n in AMENAZAS_NATURALES.items()}).sort_values(ascending=False)
        salida += ["Amenazas naturales", numero(fila["Natural"], 1),
                   f"La mayor: {amenazas.index[0].lower()} ({numero(amenazas.iloc[0], 1)}) · "
                   f"región: {numero(alc['Natural'].mean(), 1)}"]
        promedios = {g: inform.loc[i, "INFORM RISK"].mean() for g, i in grupos.items()}
        salida += ["INFORM Risk promedio de la región (21 países)", numero(alc["INFORM RISK"].mean(), 1),
                   f"Emergentes: {numero(promedios['Economías emergentes'], 1)} · "
                   f"avanzadas: {numero(promedios['Economías avanzadas'], 1)}"]
        return salida

    # Los 6 gráficos numerados (figura, título y unidad)
    @app.callback([Output(f"inf-graf-{n}{parte}", prop) for n in range(1, 7)
                   for parte, prop in (("", "figure"), ("-titulo", "children"), ("-unidad", "children"))],
                  Input("inf-pais", "value"), Input("inf-indicador", "value"))
    def actualizar_graficos_inform(iso, col):
        nombre = INDICADORES_INFORM[col]
        pais = inform.loc[iso, "NOMBRE"]
        if col == "INFORM RISK":
            titulo, mapa = "Índice INFORM Risk", "Mapa del índice INFORM Risk"
        else:
            titulo, mapa = f"{nombre} (INFORM Risk)", f"Mapa de {nombre.lower()} (INFORM Risk)"
        puntaje = "(Puntaje de 0 a 10)"
        return (grafico_ranking_inform(alc, iso, col), f"{titulo} en América Latina y el Caribe, {ANIO_INFORM}", puntaje,
                grafico_mapa_inform(inform, iso, col), f"{mapa}, {ANIO_INFORM}", "(Nivel de riesgo)",
                grafico_grupos_inform(inform, iso, col),
                f"{titulo}: América Latina frente a economías avanzadas y emergentes, {ANIO_INFORM}", puntaje,
                grafico_amenazas_inform(inform, iso),
                f"Amenazas naturales en {pais} frente al promedio de América Latina, {ANIO_INFORM}", puntaje,
                grafico_dimensiones_inform(alc, iso),
                f"Las tres dimensiones del INFORM Risk en América Latina y el Caribe, {ANIO_INFORM}",
                "(Puntaje de 0 a 10; la raya negra es el INFORM Risk)",
                grafico_subcomponentes_inform(inform, iso),
                f"Vulnerabilidad y falta de capacidad de respuesta en {pais}, {ANIO_INFORM}",
                "(Puntaje de 0 a 10; el rombo es el promedio de América Latina)")


# ---- 6.5 América Latina ----------------------------------------------------------------------
def registrar_callbacks_latam(app, latam):
    e = latam["eventos"]
    a0, a1 = int(e["anio"].min()), int(e["anio"].max())
    por_iso = dict(zip(e["iso3"], e["pais"]))
    paises = set(e["pais"])
    zonas = dict(zip(e["pais"], e["zona"]))
    en_zona = {"Sudamérica": "en Sudamérica", "Caribe": "en el Caribe"}

    def periodo_efectivo(periodo, ind):
        """El costo solo llega a 2025: con «costo», el periodo se recorta."""
        p0, p1 = periodo
        if ind == "costo_pbi":
            p1 = min(p1, ULTIMO_ANIO_COSTO)
            p0 = min(p0, p1)
        return p0, p1

    def con_articulo(pais):
        return ARTICULOS.get(pais, pais)

    # Clic en el ranking, el mapa, los periodos o los tipos por zona -> país; botón -> reinicio
    @app.callback(Output("latam-pais", "value"), Output("latam-periodo", "value"),
                  Input("latam-graf-1", "clickData"), Input("latam-graf-2", "clickData"),
                  Input("latam-graf-3", "clickData"), Input("latam-graf-4", "clickData"),
                  Input("latam-btn-reiniciar", "n_clicks"), prevent_initial_call=True)
    def sincronizar_latam(*_):
        origen = ctx.triggered_id
        if origen == "latam-btn-reiniciar":
            return PAIS_INICIAL, [a0, a1]
        clic = ctx.triggered[0]["value"] if ctx.triggered else None
        punto = ((clic or {}).get("points") or [{}])[0]
        if origen == "latam-graf-1":
            pais = punto.get("y")
        elif origen == "latam-graf-2":
            pais = por_iso.get(punto.get("location"))
        else:  # en el gráfico de correlación la x es un número: no cambia nada
            pais = punto.get("x")
        return (pais if pais in paises else no_update), no_update

    # Barra lateral, encabezado y las 6 tarjetas
    salidas = [Output("latam-lateral-pais", "children"), Output("latam-lateral-periodo", "children"),
               Output("latam-antetitulo", "children"), Output("latam-titulo", "children"),
               Output("latam-sub", "children")]
    salidas += [Output(f"latam-kpi-{i}-{parte}", "children") for i in range(6) for parte in ("titulo", "valor", "nota")]

    @app.callback(salidas, Input("latam-pais", "value"), Input("latam-periodo", "value"),
                  Input("latam-indicador", "value"))
    def actualizar_resumen_latam(pais, periodo, ind):
        p0, p1 = periodo_efectivo(periodo, ind)
        texto = texto_periodo(p0, p1).replace("–", "-")
        tabla = resumen_paises(latam, p0, p1)
        fila = tabla.loc[pais]
        n_paises = len(tabla)
        ev = eventos_periodo(latam, p0, p1)
        ev = ev[ev["pais"] == pais]
        donde = con_articulo(pais)

        def puesto(col):
            return f"Puesto {int((tabla[col] > fila[col]).sum()) + 1} de {n_paises} países"

        salida = [pais, texto_periodo(*periodo).replace("–", "-"), TEXTOS_LATAM[ind]["indicador"]]
        if ind == "eventos":
            salida += [f"Desastres naturales en {donde}, {texto}",
                       f"¿Qué mide? {TEXTOS_LATAM[ind]['que_mide']} {n_paises} países de América Latina y el Caribe."]
            antes = len(e[(e["pais"] == pais) & (e["anio"] <= CORTE_PERIODOS)])
            despues = len(e[(e["pais"] == pais) & (e["anio"] > CORTE_PERIODOS)])
            cambio = variacion_pct(despues, antes)
            nota = [html.Span(puesto("eventos"))]
            if cambio is not None:
                nota += [html.Br(), html.Span(f"{antes} → {despues} ({con_signo(cambio, 1)} %) entre {a0}-{CORTE_PERIODOS} "
                                              f"y {CORTE_PERIODOS + 1}-{a1}",
                                              className="delta-sube" if cambio > 0 else "delta-baja")]
            salida += ["Desastres registrados", numero(fila["eventos"], 0), nota]
            total = max(len(ev), 1)
            for sub in ORDEN_SUBGRUPOS_PAIS:
                del_tipo = ev[ev["subgrupo"] == sub]
                nota = f"{numero(len(del_tipo) / total * 100, 1)} % del total"
                if len(del_tipo):
                    principal = del_tipo["tipo"].value_counts()
                    nota += f" · {nombre_corto(principal.index[0])} ({principal.iloc[0]})"
                salida += [f"Desastres {sub.lower()}s", numero(len(del_tipo), 0), nota]
        else:
            pbi = latam["pbi"].loc[pais, "pbi_millones_usd"] * 1000
            zona = zonas[pais]
            promedio_zona = tabla.loc[tabla["zona"] == zona, "costo_pbi"].mean()
            salida += [f"Costo económico de los desastres en {donde}, {texto}",
                       f"¿Qué mide? {TEXTOS_LATAM[ind]['que_mide']}"]
            nota = [html.Span(puesto("costo_pbi")), html.Br(),
                    html.Span(f"US$ {numero(fila['dano_miles'] / 1000, 0)} millones · "
                              f"promedio {en_zona.get(zona, 'en su zona')}: {numero(promedio_zona, 1)} %")]
            salida += ["Costo económico", [numero(fila["costo_pbi"], 1), html.Span(" % del PBI", className="kpi-unidad")],
                       nota]
            total = fila["dano_miles"] or np.nan
            for sub in ORDEN_SUBGRUPOS_PAIS:
                del_tipo = ev[ev["subgrupo"] == sub]
                costo = del_tipo["dano_miles_usd"].sum()
                parte = 0 if es_vacio(total) else costo / total * 100
                nota = f"{numero(parte, 1)} % del costo total"
                if costo > 0:
                    principal = del_tipo.groupby("tipo")["dano_miles_usd"].sum().sort_values(ascending=False)
                    nota += f" · {nombre_corto(principal.index[0])}"
                salida += [f"Costo de desastres {sub.lower()}s",
                           [numero(costo / pbi * 100, 1), html.Span(" % del PBI", className="kpi-unidad")], nota]
        return salida

    # Los 6 gráficos numerados (figura, título y unidad de cada uno)
    @app.callback([Output(f"latam-graf-{n}{parte}", prop) for n in range(1, 7)
                   for parte, prop in (("", "figure"), ("-titulo", "children"), ("-unidad", "children"))],
                  Input("latam-pais", "value"), Input("latam-periodo", "value"), Input("latam-indicador", "value"))
    def actualizar_graficos_latam(pais, periodo, ind):
        p0, p1 = periodo_efectivo(periodo, ind)
        texto = texto_periodo(p0, p1).replace("–", "-")
        tabla = resumen_paises(latam, p0, p1)
        zona = zonas[pais]
        lugar = en_zona.get(zona, "en México y Centroamérica")
        donde = con_articulo(pais)
        if ind == "eventos":
            return (grafico_ranking_latam(tabla, pais, ind), f"Número de desastres naturales, {texto}",
                    "(N.º de desastres)",
                    grafico_mapa_latam(tabla, pais, ind), f"Mapa del número de desastres, {texto}", "(N.º de desastres)",
                    grafico_periodos_latam(latam, pais),
                    f"Número de desastres naturales en países de América Latina y el Caribe, "
                    f"{a0}-{CORTE_PERIODOS} y {CORTE_PERIODOS + 1}-{a1}", "(N.º de desastres; no depende del periodo)",
                    grafico_tipos_zona(latam, p0, p1, zona, pais, "eventos"),
                    f"Desastres naturales por tipo según su origen {lugar}, {texto}", "(% del total de desastres)",
                    grafico_tipos_pais(latam, pais, p0, p1, "eventos"),
                    f"Número de desastres registrados en {donde} por tipo y subtipo, {texto}", "(N.º de desastres)",
                    grafico_anual_pais(latam, pais, p0, p1, "eventos"),
                    f"Número anual de desastres en {donde} por tipo, {texto}",
                    "(N.º de desastres y promedio móvil de 5 años)")
        return (grafico_ranking_latam(tabla, pais, ind), f"Costo económico de los desastres, {texto}", "(% del PBI)",
                grafico_mapa_latam(tabla, pais, ind), f"Mapa del costo económico, {texto}", "(% del PBI)",
                grafico_dispersion_latam(latam, p0, p1),
                f"Número anual de desastres y costo económico en América Latina y el Caribe, {texto}",
                "(Miles de millones de US$)",
                grafico_tipos_zona(latam, p0, p1, zona, pais, "costo"),
                f"Distribución del costo económico de los desastres por tipo según su origen {lugar}, {texto}",
                "(% del costo total)",
                grafico_tipos_pais(latam, pais, p0, p1, "costo"),
                f"Costo de los desastres en {donde} por tipo y subtipo, {texto}", "(% del PBI)",
                grafico_anual_pais(latam, pais, p0, p1, "costo"),
                f"Costo económico anual de los desastres en {donde}, {texto}",
                "(Millones de US$, ajustados por inflación)")




# =============================================================================
# 7. EJECUTAR
# =============================================================================
GEOJSON = cargar_geojson()
CLIMA = cargar_clima()
BASE_ID = cargar_id()
INFORM = cargar_inform()
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
paginas = [crear_marco(), crear_portada(CLIMA, BASE_ID, INFORM, LATAM)]
if CLIMA is not None:
    paginas.append(crear_layout_clima(CLIMA))
if BASE_ID is not None:
    paginas.append(crear_layout_id(BASE_ID))
if INFORM is not None:
    paginas.append(crear_layout_inform(INFORM))
if LATAM is not None:
    paginas.append(crear_layout_latam(LATAM))
app.validation_layout = html.Div(paginas)

registrar_navegacion(app, CLIMA, BASE_ID, INFORM, LATAM)
if CLIMA is not None:
    registrar_callbacks_clima(app, CLIMA)
if BASE_ID is not None:
    registrar_callbacks_id(app, BASE_ID, GEOJSON)
if INFORM is not None:
    registrar_callbacks_inform(app, INFORM)
if LATAM is not None:
    registrar_callbacks_latam(app, LATAM)

if __name__ == "__main__":
    # En JupyterLab: "external" muestra el enlace http://127.0.0.1:8050 (ábrelo en el navegador).
    # Para verlo en tu celular (misma red WiFi): cambia a host="0.0.0.0" y abre en el celular
    # http://IP-DE-TU-PC:8050  (la IP la ves con el comando ipconfig en Windows).
    app.run(jupyter_mode="external", host="127.0.0.1", port=8050, debug=True)
