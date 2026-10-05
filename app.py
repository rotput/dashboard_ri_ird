# =============================================================================
#  DASHBOARD: ÍNDICE DE RIESGO DE DESASTRES (IRD) - PERÚ 2003-2025
#  Pega todo en UNA celda de JupyterLab y ejecútala.
#  Se abre en: http://127.0.0.1:8050
#
#  Librerías (una sola vez, en otra celda):
#      !pip install -U dash plotly pandas openpyxl
#
#  Secciones:
#    1. CONFIGURACIÓN   -> ruta del Excel, colores, KPIs, textos (cambia aquí)
#    2. FORMATO         -> números y nombres
#    3. DATOS           -> carga del Excel y cálculos
#    4. GRÁFICOS        -> una función por gráfico
#    5. DISEÑO          -> íconos, tarjetas, paneles y estilos CSS
#    6. INTERACTIVIDAD  -> callbacks
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
    "contexto": "#a7bcd2", "mejora": "#2a78d6", "empeora": "#e34948",
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

/* Pantallas medianas */
@media (max-width:1280px){
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
.area-calor .grafico{height:560px!important}}
"""

CSS = CSS + css_iconos()

PAGINA = ("<!DOCTYPE html><html lang='es'><head>{%metas%}<title>{%title%}</title>{%favicon%}{%css%}"
          "<style>" + CSS + "</style></head><body>{%app_entry%}<footer>{%config%}{%scripts%}"
          "{%renderer%}</footer></body></html>")


# =============================================================================
# 6. INTERACTIVIDAD (callbacks)
# =============================================================================
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

app = Dash(__name__, title=f"{TITULO} | Perú", external_stylesheets=[URL_FUENTES],
           meta_tags=[{"name": "viewport", "content": "width=device-width, initial-scale=1"}])
app.index_string = PAGINA
server = app.server
app.layout = crear_layout(BASE)
registrar_callbacks(app, BASE, GEOJSON)

if __name__ == "__main__":
    # En JupyterLab: "external" muestra el enlace http://127.0.0.1:8050 (ábrelo en el navegador).
    # Para verlo en tu celular (misma red WiFi): cambia a host="0.0.0.0" y abre en el celular
    # http://IP-DE-TU-PC:8050  (la IP la ves con el comando ipconfig en Windows).
    app.run(jupyter_mode="external", host="127.0.0.1", port=8050, debug=True)
