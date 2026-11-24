import re
import time
from datetime import datetime
from urllib.parse import urljoin, urlparse, parse_qs

import pandas as pd
import requests
from bs4 import BeautifulSoup
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# =============================================================================
# CONFIGURACIÓN
# =============================================================================
URL_SENAMHI = "https://www.senamhi.gob.pe/?p=pronostico-meteorologico"
BASE_URL = "https://www.senamhi.gob.pe/"
ARCHIVO_SALIDA = "senamhi_pronostico_nacional.xlsx"
ARCHIVO_DEBUG = "senamhi_debug.html"  # se guarda solo si no se extrae nada
DIAS_ESPERADOS = 3

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/128.0.0.0 Safari/537.36"
    ),
    "Referer": "https://www.senamhi.gob.pe/?p=pronostico-meteorologico",
}

SELECTOR_DETALLE = 'a[href*="pronostico-detalle"]'


# =============================================================================
# UTILIDADES
# =============================================================================
def limpiar_numero(texto: str):
    """Extrae el primer número (negativos y decimales con . o ,) de un texto."""
    if not texto:
        return None
    m = re.search(r"-?\d+(?:[.,]\d+)?", texto)
    return float(m.group(0).replace(",", ".")) if m else None


def texto_limpio(tag) -> str:
    return " ".join(tag.get_text(" ", strip=True).split()) if tag else ""


def primero(nodo, *selectores):
    """Devuelve el primer selector que encuentre algo, EN EL ORDEN DADO.
    (select_one("a, b") no respeta prioridad: devuelve lo que aparezca antes en el HTML.)"""
    for sel in selectores:
        tag = nodo.select_one(sel)
        if tag:
            return tag
    return None


def obtener_soup(html: str):
    """Usa lxml si está disponible; si no, html.parser. Devuelve (soup, nombre_parser)."""
    try:
        return BeautifulSoup(html, "lxml"), "lxml"
    except Exception:
        return BeautifulSoup(html, "html.parser"), "html.parser"


def pertenece_a_fila(nodo, fila) -> bool:
    """True si `nodo` es de `fila` y no de otra fila de localidad anidada dentro de ella.
    Con html.parser, un <tr> sin cerrar hace que la fila siguiente quede DENTRO de la
    anterior; esto reemplaza el truco [:3], que fallaba si una localidad traía < 3 días."""
    for ancestro in nodo.parents:
        if ancestro is fila:
            return True
        if ancestro.name == "tr" and ancestro.select_one(SELECTOR_DETALLE):
            return False
    return False


def crear_sesion() -> requests.Session:
    sesion = requests.Session()
    sesion.headers.update(HEADERS)
    reintentos = Retry(
        total=3,
        backoff_factor=2,  # espera 2 s, 4 s, 8 s
        status_forcelist=[429, 500, 502, 503, 504],
        allowed_methods=["GET"],
    )
    sesion.mount("https://", HTTPAdapter(max_retries=reintentos))
    sesion.mount("http://", HTTPAdapter(max_retries=reintentos))
    return sesion


# =============================================================================
# DESCARGA
# =============================================================================
def descargar_html() -> str:
    resp = crear_sesion().get(URL_SENAMHI, timeout=40)
    resp.raise_for_status()
    # Solo adivinar la codificación si el servidor no la declaró
    # (requests pone ISO-8859-1 por defecto cuando falta el charset).
    if not resp.encoding or resp.encoding.lower() == "iso-8859-1":
        resp.encoding = resp.apparent_encoding or "utf-8"
    return resp.text


# =============================================================================
# PARSEO
# =============================================================================
def parsear_html(html: str) -> pd.DataFrame:
    soup, parser = obtener_soup(html)
    filas_tr = soup.select("table tr")
    print(f"Parser HTML            : {parser}")
    print(f"Filas <tr> detectadas  : {len(filas_tr)}")

    registros = []

    for tr in filas_tr:
        # 1. Enlace de detalle PROPIO de esta fila (contiene dp y localidad)
        enlace = next((a for a in tr.select(SELECTOR_DETALLE) if pertenece_a_fila(a, tr)), None)
        if not enlace:
            continue

        url_detalle = urljoin(BASE_URL, enlace.get("href", "").strip())
        qs = parse_qs(urlparse(url_detalle).query)
        cod_dp = qs.get("dp", [""])[0].zfill(2)
        cod_localidad = qs.get("localidad", [""])[0].zfill(4)

        # 2. Departamento y localidad (primera celda <td> propia)
        celdas = [td for td in tr.find_all("td") if pertenece_a_fila(td, tr)]
        if not celdas:
            continue

        textos_loc = [t for t in celdas[0].stripped_strings if t]
        if len(textos_loc) >= 2:
            departamento, localidad = textos_loc[0].upper(), textos_loc[1].upper()
        elif len(textos_loc) == 1:
            departamento, localidad = "", textos_loc[0].upper()
        else:
            continue

        # 3. Solo los bloques de día que pertenecen a ESTA fila
        bloques_dia = [b for b in tr.select("div.row.m-3") if pertenece_a_fila(b, tr)]

        for bloque in bloques_dia:
            tag_fecha = primero(bloque, "div.col-sm-3", "strong")
            tag_tmax = primero(bloque, "h1.text-danger", "span.text-danger", ".text-danger")
            tag_tmin = primero(bloque, "h1.text-primary", "span.text-primary", ".text-primary")
            tag_desc = primero(bloque, "div.col-sm-5", "p")
            tag_img = bloque.select_one("img")

            src_img = (tag_img.get("src") or tag_img.get("data-src") or "") if tag_img else ""

            registros.append({
                "cod_dp": cod_dp,
                "departamento": departamento,
                "cod_localidad": cod_localidad,
                "localidad": localidad,
                "fecha_texto": texto_limpio(tag_fecha),
                "temp_max_c": limpiar_numero(tag_tmax.get_text(strip=True)) if tag_tmax else None,
                "temp_min_c": limpiar_numero(tag_tmin.get_text(strip=True)) if tag_tmin else None,
                "descripcion": texto_limpio(tag_desc),
                "icono": src_img.split("/")[-1].strip(),
                "url_detalle": url_detalle,
            })

    return pd.DataFrame(registros)


# =============================================================================
# FLUJO PRINCIPAL
# =============================================================================
def scrapear_senamhi() -> pd.DataFrame:
    t0 = time.time()
    print("Descargando tabla nacional del SENAMHI...")
    html = descargar_html()

    df = parsear_html(html)

    if df.empty:
        with open(ARCHIVO_DEBUG, "w", encoding="utf-8") as f:
            f.write(html)
        raise ValueError(
            "No se extrajo ningún registro. Revisa la URL o la estructura HTML del SENAMHI. "
            f"Se guardó el HTML recibido en {ARCHIVO_DEBUG} para revisarlo."
        )

    # Fecha y hora de la consulta (el pronóstico cambia cada día)
    df.insert(0, "fecha_consulta", datetime.now().replace(microsecond=0))

    # Blindaje final contra duplicados
    filas_antes = len(df)
    df = df.drop_duplicates(
        subset=["cod_dp", "cod_localidad", "fecha_texto"],
        keep="first",
    ).reset_index(drop=True)
    duplicados_eliminados = filas_antes - len(df)

    # =========================================================================
    # CONTROL DE CALIDAD
    # =========================================================================
    conteo_loc = df.groupby(["cod_dp", "departamento", "cod_localidad", "localidad"]).size()
    anomalas = conteo_loc[conteo_loc != DIAS_ESPERADOS]
    sin_temp = df["temp_max_c"].isna().sum() + df["temp_min_c"].isna().sum()

    print("-" * 55)
    print(f"Tiempo total           : {time.time() - t0:.2f} s")
    print(f"Filas finales          : {len(df):,}")
    print(f"Localidades únicas     : {len(conteo_loc):,}")
    print(f"Departamentos únicos   : {df['departamento'].nunique()}")
    print(f"Duplicados eliminados  : {duplicados_eliminados}")
    print(f"Temperaturas vacías    : {sin_temp}")
    print(f"Localidades != {DIAS_ESPERADOS} días  : {len(anomalas)}")
    if len(anomalas) > 0:
        print(f"Aviso — localidades con cantidad distinta de {DIAS_ESPERADOS} días:")
        print(anomalas)
    print("-" * 55)

    return df


if __name__ == "__main__":
    df_final = scrapear_senamhi()

    with pd.ExcelWriter(ARCHIVO_SALIDA, engine="openpyxl") as writer:
        df_final.to_excel(writer, sheet_name="Pronostico", index=False)

    print(f"Archivo guardado correctamente: {ARCHIVO_SALIDA}")
