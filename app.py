import dash
from dash import dcc, html, Input, Output, State
import dash_bootstrap_components as dbc
import plotly.express as px
import pandas as pd
import numpy as np

# 1. INICIALIZACIÓN DE LA APP (Tema Corporativo)
app = dash.Dash(__name__, external_stylesheets=[dbc.themes.FLATLY])
server = app.server  # <-- PIEZA CLAVE PARA RENDER (Gunicorn)

# 2. DISEÑO DE LA PÁGINA
app.layout = dbc.Container([
    
    # Encabezado
    dbc.Row([
        dbc.Col(
            html.Div([
                html.H2("Simulador DSGE: Modelo Neokeynesiano", className="text-white fw-bold m-0"),
                html.P("Funciones de Impulso-Respuesta (IRF) con parámetros estructurales personalizables", className="text-light m-0 mt-2")
            ], className="p-4 rounded shadow-sm text-center mb-4 mt-3", style={'backgroundColor': '#2c3e50'})
        )
    ]),

    dbc.Row([
        
        # ================= PANEL DE CONTROL (PARÁMETROS) =================
        dbc.Col([
            dbc.Card([
                dbc.CardHeader(html.H5("Parámetros Estructurales", className="fw-bold m-0")),
                dbc.CardBody([
                    
                    html.Label("Sensibilidad IS a la tasa (1/σ):", className="fw-bold mt-2 text-primary"),
                    dcc.Input(id='param-sigma', type='number', value=0.5, step=0.1, className="form-control mb-2"),
                    
                    html.Label("Pendiente Curva Phillips (κ):", className="fw-bold mt-2 text-primary"),
                    dcc.Input(id='param-kappa', type='number', value=0.15, step=0.05, className="form-control mb-2"),

                    html.Hr(),
                    html.H6("Regla de Taylor (Banco Central)", className="fw-bold text-secondary"),
                    
                    html.Label("Aversión a la Inflación (φ_π):", className="fw-bold mt-1"),
                    dcc.Input(id='param-phipi', type='number', value=1.5, step=0.1, className="form-control mb-2"),
                    
                    html.Label("Respuesta al PBI (φ_y):", className="fw-bold mt-1"),
                    dcc.Input(id='param-phiy', type='number', value=0.5, step=0.1, className="form-control mb-2"),
                    
                    html.Label("Suavización de Tasa (ρ):", className="fw-bold mt-1"),
                    dcc.Input(id='param-rho', type='number', value=0.7, step=0.05, className="form-control mb-3"),

                    html.Hr(),
                    html.H6("Configuración del Choque", className="fw-bold text-danger"),
                    dcc.Dropdown(
                        id='tipo-choque',
                        options=[
                            {'label': 'Shock de Oferta (Inflacionario)', 'value': 'oferta'},
                            {'label': 'Shock de Demanda (Consumo)', 'value': 'demanda'},
                            {'label': 'Shock de Política Monetaria (Tasa)', 'value': 'monetario'}
                        ],
                        value='oferta', clearable=False, className="mb-2"
                    ),
                    
                    html.Label("Magnitud del Choque:", className="fw-bold mt-1"),
                    dcc.Input(id='magnitud-choque', type='number', value=1.0, step=0.1, className="form-control mb-4"),

                    # EL BOTÓN DE EJECUCIÓN (Gatillo)
                    html.Div(
                        html.Button('EJECUTAR SIMULACIÓN', id='btn-correr', n_clicks=0, 
                                    className="btn btn-success btn-lg w-100 fw-bold shadow-sm"),
                        className="d-grid gap-2"
                    )

                ])
            ], className="shadow-sm border-0", style={"height": "100%"})
        ], width=12, lg=3, className="mb-4"),
        
        # ================= PANEL DE RESULTADOS (IRFs) =================
        dbc.Col([
            dbc.Card([
                dbc.CardBody([
                    html.H4("Funciones de Impulso-Respuesta (Desviación % del EE)", className="card-title text-secondary fw-bold mb-4"),
                    dcc.Graph(id='grafico-irf', style={"height": "650px"})
                ])
            ], className="shadow-sm border-0", style={"height": "100%"})
        ], width=12, lg=9, className="mb-4")
        
    ])
], fluid=True, style={'backgroundColor': '#f8f9fa', 'minHeight': '100vh', 'padding': '20px'})


# 3. EL CEREBRO MATEMÁTICO (Resolver y Graficar)
@app.callback(
    Output('grafico-irf', 'figure'),
    Input('btn-correr', 'n_clicks'), # Solo el botón detona el cálculo
    [State('param-sigma', 'value'),  # Los State extraen los valores silenciosamente
     State('param-kappa', 'value'),
     State('param-phipi', 'value'),
     State('param-phiy', 'value'),
     State('param-rho', 'value'),
     State('tipo-choque', 'value'),
     State('magnitud-choque', 'value')]
)
def calcular_modelo(n_clicks, sigma, kappa, phi_pi, phi_y, rho, tipo, magnitud):
    
    # Validaciones de seguridad por si dejas un cuadro vacío
    sigma = sigma or 0.5
    kappa = kappa or 0.15
    phi_pi = phi_pi or 1.5
    phi_y = phi_y or 0.5
    rho = rho or 0.7
    magnitud = magnitud or 1.0

    T = 40 # 40 trimestres
    y = np.zeros(T)   # Brecha del PBI
    pi = np.zeros(T)  # Inflación
    i = np.zeros(T)   # Tasa de Interés

    # Impacto inicial en t=0 según el tipo de choque
    if tipo == 'oferta':
        pi[0] = magnitud          # Sube la inflación directo
        y[0] = -0.5 * magnitud    # Cae el producto
    elif tipo == 'demanda':
        y[0] = magnitud           # Sube el consumo/producto
        pi[0] = 0.3 * magnitud    # Presiona precios al alza
    elif tipo == 'monetario':
        i[0] = magnitud           # El BCR sube la tasa sorpresivamente
        y[0] = -sigma * magnitud  # Contrae la demanda
        pi[0] = -kappa * magnitud # Enfría los precios

    # Sistema dinámico recursivo (VAR Estructural Linealizado)
    for t in range(1, T):
        # 1. Regla de Taylor (El BCR reacciona a la inflación y PBI del periodo anterior para suavizar)
        i[t] = rho * i[t-1] + (1 - rho) * (phi_pi * pi[t-1] + phi_y * y[t-1])
        
        # 2. Curva IS (El producto depende de la tasa de interés real esperada/pasada)
        tasa_real = i[t-1] - pi[t-1]
        y[t] = 0.8 * y[t-1] - sigma * tasa_real
        
        # 3. Curva de Phillips (La inflación depende de la inercia y la brecha del producto actual)
        pi[t] = 0.6 * pi[t-1] + kappa * y[t]

    # Consolidar datos para el gráfico
    df_irf = pd.DataFrame({
        'Trimestre': np.tile(np.arange(T), 3),
        'Desviación (%)': np.concatenate([y, pi, i]),
        'Variable': ['1. Brecha de Producción (y)']*T + ['2. Inflación (π)']*T + ['3. Tasa de Política (i)']*T
    })

    # Construir Subplots con Plotly Express
    fig = px.line(df_irf, x='Trimestre', y='Desviación (%)', facet_col='Variable', color='Variable',
                  color_discrete_sequence=['#2980b9', '#c0392b', '#27ae60'])
    
    fig.update_layout(template="plotly_white", showlegend=False, 
                      margin=dict(t=50, l=20, r=20, b=20),
                      font=dict(size=13))
    
    # Ajustes visuales de las escalas (cada gráfico tiene su propio eje Y)
    fig.update_yaxes(matches=None, showticklabels=True) 
    fig.add_hline(y=0, line_dash="solid", line_color="black", opacity=0.4, line_width=1.5)
    
    # Limpiar los títulos automáticos molestos de Plotly
    fig.for_each_annotation(lambda a: a.update(text=a.text.split("=")[-1], font=dict(weight='bold')))

    return fig

if __name__ == '__main__':
    app.run(debug=True)
