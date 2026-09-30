import dash
from dash import dcc, html, Input, Output
import dash_bootstrap_components as dbc
import plotly.express as px
import pandas as pd

# ============================================================
# 1. CARGA DE LA BASE DE DATOS REAL
# ============================================================
# Render leerá este archivo que subiste a GitHub
ruta_excel = "BASE_ANUAL_DPTO_INCIDENCIAS_TOTALES.xlsx"
df = pd.read_excel(ruta_excel, sheet_name="MATRIZ_ANUAL_DPTO")

# Identificar automáticamente las variables numéricas de daños
variables_metrica = [col for col in df.columns if col not in ['AÑO', 'DPTO.']]
df['AÑO'] = df['AÑO'].astype(int)

# ============================================================
# 2. CONFIGURACIÓN DEL DASHBOARD Y SERVIDOR (NUBE)
# ============================================================
app = dash.Dash(__name__, external_stylesheets=[dbc.themes.CYBORG])
server = app.server # ¡ESTA LÍNEA ES CRÍTICA PARA QUE RENDER FUNCIONE!

# ============================================================
# 3. ESTRUCTURA VISUAL (LAYOUT)
# ============================================================
app.layout = dbc.Container([
    # Encabezado
    dbc.Row([
        dbc.Col(html.Div([
            html.H2("⚡ Sistema de Monitoreo de Riesgos Naturales", className="text-warning mt-4 text-center fw-bold"),
            html.P("Análisis Histórico de Incidencias y Daños a Nivel Departamental", className="text-center text-light mb-4")
        ]), width=12)
    ]),

    # Panel principal
    dbc.Row([
        # --- BARRA LATERAL DE CONTROLES ---
        dbc.Col([
            dbc.Card([
                dbc.CardHeader("⚙️ Filtros de Análisis", className="fw-bold text-dark bg-warning"),
                dbc.CardBody([
                    html.Label("Métrica Principal:", className="text-light fw-bold"),
                    dcc.Dropdown(
                        id="drop-metrica",
                        options=[{"label": v.replace("_", " ").title(), "value": v} for v in variables_metrica],
                        value="NUMERO_DE_INCIDENCIAS",
                        clearable=False,
                        className="text-dark mb-4"
                    ),
                    
                    html.Label("Departamentos:", className="text-light fw-bold"),
                    dcc.Dropdown(
                        id="drop-dpto",
                        options=[{"label": d, "value": d} for d in sorted(df['DPTO.'].unique())],
                        value=sorted(df['DPTO.'].unique())[:5],
                        multi=True,
                        className="text-dark mb-4"
                    ),
                    
                    html.Label("Rango de Años:", className="text-light fw-bold"),
                    dcc.RangeSlider(
                        id="slider-años",
                        min=df['AÑO'].min(), max=df['AÑO'].max(),
                        step=1,
                        marks={int(y): str(y) for y in range(df['AÑO'].min(), df['AÑO'].max()+1, 2)},
                        value=[df['AÑO'].min(), df['AÑO'].max()],
                        tooltip={"placement": "bottom", "always_visible": True}
                    )
                ])
            ], className="border-secondary shadow-lg h-100")
        ], width=3),

        # --- ZONA DE GRÁFICOS Y KPIS ---
        dbc.Col([
            # Fila de KPIs
            dbc.Row([
                dbc.Col(dbc.Card(dbc.CardBody([html.H6("Total de Registros", className="text-secondary"), html.H3(id="kpi-total", className="text-info fw-bold")])), width=4),
                dbc.Col(dbc.Card(dbc.CardBody([html.H6("Promedio Anual", className="text-secondary"), html.H3(id="kpi-promedio", className="text-success fw-bold")])), width=4),
                dbc.Col(dbc.Card(dbc.CardBody([html.H6("Departamento más Crítico", className="text-secondary"), html.H3(id="kpi-max-dpto", className="text-danger fw-bold")])), width=4),
            ], className="mb-4"),

            # Gráfico de Tendencia (Líneas)
            dbc.Row([
                dbc.Col(dbc.Card(dbc.CardBody(dcc.Graph(id="graf-tendencia", style={"height": "350px"}))), width=12)
            ], className="mb-4"),

            # Gráficos Inferiores (Barras y Treemap)
            dbc.Row([
                dbc.Col(dbc.Card(dbc.CardBody(dcc.Graph(id="graf-barras", style={"height": "350px"}))), width=6),
                dbc.Col(dbc.Card(dbc.CardBody(dcc.Graph(id="graf-treemap", style={"height": "350px"}))), width=6),
            ])
        ], width=9)
    ])
], fluid=True, className="pb-5")

# ============================================================
# 4. MOTOR DE INTERACTIVIDAD (CALLBACKS)
# ============================================================
@app.callback(
    [Output("kpi-total", "children"), Output("kpi-promedio", "children"), Output("kpi-max-dpto", "children"),
     Output("graf-tendencia", "figure"), Output("graf-barras", "figure"), Output("graf-treemap", "figure")],
    [Input("drop-metrica", "value"), Input("drop-dpto", "value"), Input("slider-años", "value")]
)
def actualizar_dashboard(metrica, dptos, rango_años):
    dff = df[(df['DPTO.'].isin(dptos)) & (df['AÑO'] >= rango_años[0]) & (df['AÑO'] <= rango_años[1])]
    
    if dff.empty:
        return "0", "0", "N/A", {}, {}, {}

    total = dff[metrica].sum()
    promedio = total / len(dff['AÑO'].unique()) if len(dff['AÑO'].unique()) > 0 else 0
    agrupado_dpto = dff.groupby('DPTO.')[metrica].sum().reset_index()
    dpto_top = agrupado_dpto.loc[agrupado_dpto[metrica].idxmax()]['DPTO.'] if not agrupado_dpto.empty else "N/A"

    formato_total = f"{total:,.0f}"
    formato_prom = f"{promedio:,.1f}"

    df_tendencia = dff.groupby(['AÑO', 'DPTO.'])[metrica].sum().reset_index()
    fig_tendencia = px.line(df_tendencia, x="AÑO", y=metrica, color="DPTO.", markers=True,
                            title=f"📈 Evolución Histórica: {metrica.replace('_', ' ').title()}")
    fig_tendencia.update_layout(template="plotly_dark", plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
                                margin=dict(l=20, r=20, t=40, b=20), legend_title_text="Dpto.")

    dpto_top10 = agrupado_dpto.nlargest(10, metrica).sort_values(metrica, ascending=True)
    fig_barras = px.bar(dpto_top10, x=metrica, y="DPTO.", orientation='h', color=metrica, color_continuous_scale="Reds",
                        title="🏆 Ranking Acumulado")
    fig_barras.update_layout(template="plotly_dark", plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
                             margin=dict(l=20, r=20, t=40, b=20), coloraxis_showscale=False)

    fig_treemap = px.treemap(agrupado_dpto, path=[px.Constant("Perú"), "DPTO."], values=metrica, color=metrica,
                             color_continuous_scale="Viridis", title="🗺️ Mapa de Concentración")
    fig_treemap.update_layout(template="plotly_dark", plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
                              margin=dict(l=10, r=10, t=40, b=10))

    return formato_total, formato_prom, dpto_top, fig_tendencia, fig_barras, fig_treemap

# ============================================================
# 5. EJECUCIÓN DEL SERVIDOR
# ============================================================
if __name__ == "__main__":
    app.run_server(debug=False)
