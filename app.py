import dash
from dash import dcc, html, Input, Output
import dash_bootstrap_components as dbc
import plotly.express as px
import pandas as pd
import numpy as np
import os

# 1. CARGA DE LA BASE DE DATOS REAL (Nota que solo va el nombre del archivo)
ruta_excel = "BASE_ANUAL_DPTO_INCIDENCIAS_TOTALES.xlsx"
df = pd.read_excel(ruta_excel, sheet_name="MATRIZ_ANUAL_DPTO")

# 2. CONFIGURACIÓN DEL DASHBOARD
app = dash.Dash(__name__, external_stylesheets=[dbc.themes.CYBORG])
server = app.server # <--- ESTA ES LA LÍNEA MÁS IMPORTANTE PARA LA NUBE
