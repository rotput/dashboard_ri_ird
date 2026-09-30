import dash
from dash import dcc, html, Input, Output
import dash_bootstrap_components as dbc
import plotly.express as px
import pandas as pd
import urllib.request
import json

# =====================================================================
# 1. CONFIGURACIÓN DE TU BASE DE DATOS (¡Modifica esto según tu Excel!)
# =====================================================================
ARCHIVO_EXCEL = "BASE_ANUAL_DPTO_INCIDENCIAS_TOTALES.xlsx"
COL_DPTO = "DEPARTAMENTO"      # Nombre de tu columna de departamentos
COL_COSTO = "COSTO_ECONOMICO"  # Nombre de tu columna de costo en soles/dólares
COL_ANIO = "ANIO"              # Nombre de tu columna de años
# =====================================================================

# 2. CARGA Y LIMPIEZA DE DATOS
df = pd.read_excel(ARCHIVO_EXCEL)

# Limpieza básica para que los nombres coincidan perfectamente con el mapa
df[COL_DPTO] = df[COL_DPTO].astype(str).str.strip().str.upper()
# Si tienes valores vacíos en costo, los rellenamos con 0
if COL_COSTO in df.columns:
    df[COL_COSTO] = pd.to_numeric(df[COL_COSTO], errors='coerce').fillna(0)

# Descargamos el mapa oficial (GeoJSON) de los departamentos del Perú
url_geojson = "https://raw.githubusercontent.com/juaneladio/peru-geojson/master/peru_departamental_simple.geojson"
with urllib.request.urlopen(url_geojson) as response:
    peru_mapa = json.loads(response.read().decode())

# Obtenemos la lista de departamentos para el menú desplegable
lista_departamentos = df[COL_DPTO].unique().tolist()
lista_departamentos.sort()
lista_departamentos.insert(0, "Nacional (Todo el Perú)")

# 3. INICIALIZACIÓN DE LA APP (Tema oscuro y elegante)
app = dash.Dash(__name__, external_stylesheets=[dbc.themes.CYBORG])
server = app.server

# 4. DISEÑO DE LA PÁGINA (Layout)
app.layout = dbc.Container([
    # Encabezado
    dbc.Row([
        dbc.Col(html.H2("Tablero de Control: Incidencias de Origen Natural", 
                        className="text-center text-primary mt-4 mb-2")),
    ]),
    dbc.Row([
        dbc.Col(html.P("Monitoreo de Impacto y Costo Económico por Departamento", 
                       className="text-center text-muted mb-4")),
    ]),

    # Fila de Filtros
    dbc.Row([
        dbc.Col([
            html.Label("Seleccionar Departamento:", className="fw-bold text-info"),
            dcc.Dropdown(
                id='filtro-dpto',
                options=[{'label': d, 'value': d} for d in lista_departamentos],
                value="Nacional (Todo el Perú)", # Valor por defecto
                clearable=False,
                style={'color': '#000000'} # Letra negra para que se lea en fondo blanco del filtro
            )
        ], width=12, md=6, className="mb-4 mx-auto"),
    ]),

    # Fila de Tarjetas (KPIs)
    dbc.Row([
        dbc.Col(dbc.Card([
            dbc.CardBody([
                html.H5("Total de Incidencias", className="card-title text-secondary"),
                html.H3(id="kpi-incidencias", className="text-white")
            ])
        ], color="dark", inverse=True), width=12, md=6),
        
        dbc.Col(dbc.Card([
            dbc.CardBody([
                html.H5("Costo Económico Total (S/)", className="card-title text-secondary"),
                html.H3(id="kpi-costo", className="text-warning")
            ])
        ], color="dark", inverse=True), width=12, md=6),
    ], className="mb-4"),

    # Fila de Gráficos (Mapa a la izquierda, Tendencia a la derecha)
    dbc.Row([
        # El Mapa
        dbc.Col(dcc.Graph(id='grafico-mapa'), width=12, lg=6, className="mb-4"),
        # Gráfico de Líneas (Tendencia)
        dbc.Col(dcc.Graph(id='grafico-tendencia'), width=12, lg=6, className="mb-4"),
    ])

], fluid=True, style={'padding': '20px'})


# 5. EL CEREBRO DE LA APP (Callbacks que actualizan todo al cambiar el filtro)
@app.callback(
    [Output('kpi-incidencias', 'children'),
     Output('kpi-costo', 'children'),
     Output('grafico-mapa', 'figure'),
     Output('grafico-tendencia', 'figure')],
    [Input('filtro-dpto', 'value')]
)
def actualizar_dashboard(dpto_seleccionado):
    
    # 1. Filtrar la base de datos
    if dpto_seleccionado == "Nacional (Todo el Perú)":
        dff = df.copy()
        titulo_zona = "Nivel Nacional"
    else:
        dff = df[df[COL_DPTO] == dpto_seleccionado].copy()
        titulo_zona = dpto_seleccionado

    # 2. Calcular KPIs
    total_incidencias = len(dff)
    
    if COL_COSTO in dff.columns:
        costo_total = dff[COL_COSTO].sum()
        texto_costo = f"S/ {costo_total:,.2f}"
    else:
        texto_costo = "Sin datos de costo"

    # 3. Crear el Mapa
    # Agrupamos los datos por departamento para que el mapa se coloree
    if COL_COSTO in df.columns:
        df_mapa = dff.groupby(COL_DPTO)[COL_COSTO].sum().reset_index()
        variable_color = COL_COSTO
    else:
        df_mapa = dff.groupby(COL_DPTO).size().reset_index(name='CONTEO')
        variable_color = 'CONTEO'

    fig_mapa = px.choropleth_mapbox(
        df_mapa, geojson=peru_mapa, 
        locations=COL_DPTO, featureidkey="properties.NOMBDEP",
        color=variable_color,
        color_continuous_scale="Reds",
        mapbox_style="carto-darkmatter",
        zoom=3.8, center={"lat": -9.189, "lon": -75.015},
        title=f"Distribución de Impacto - {titulo_zona}"
    )
    fig_mapa.update_layout(margin={"r":0,"t":40,"l":0,"b":0}, template="plotly_dark", plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)')

    # 4. Crear Gráfico de Tendencia Anual
    if COL_ANIO in dff.columns:
        df_tendencia = dff.groupby(COL_ANIO).size().reset_index(name='Frecuencia')
        fig_tendencia = px.line(
            df_tendencia, x=COL_ANIO, y='Frecuencia', markers=True,
            title=f"Tendencia Anual de Incidencias - {titulo_zona}",
            color_discrete_sequence=['#00f5d4']
        )
    else:
        # Si no hay columna de año, muestra un gráfico vacío
        fig_tendencia = px.line(title="Falta la columna de Años para la tendencia")
        
    fig_tendencia.update_layout(template="plotly_dark", plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)')

    return f"{total_incidencias:,.0f}", texto_costo, fig_mapa, fig_tendencia


if __name__ == '__main__':
    app.run_server(debug=True)
