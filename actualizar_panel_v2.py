# -*- coding: utf-8 -*-
"""
Script de actualización del Panel de Licitaciones GEMCO — v2.

Cambios respecto a v1:
- Cruza PPublicas con la hoja Checklist (por ID Licitación) para traer
  los datos de garantía (Requiere / Tipo / Monto / Moneda / Banco / Fechas).
- La sección de garantía solo se muestra si "Requiere garantía" = SI
  (comparación normalizada: sin importar mayúsculas/espacios).
- El panel ahora se puede pausar con clic o barra espaciadora.
- Header con el logo real de GEMCO sobre una franja oscura.
- Slide de isotipo GEMCO (5 segundos) al final de cada ciclo completo.
- El Excel se descarga desde SharePoint vía Microsoft Graph API (MSAL,
  Client Credentials Flow) en vez de leerse desde una ruta local
  sincronizada por OneDrive. Ver descargar_excel_desde_sharepoint().

Requiere: pandas, openpyxl, msal, requests  (pip install -r requirements.txt)

Variables de entorno requeridas:
- SP_TENANT_ID     : Tenant ID de Entra ID
- SP_CLIENT_ID     : Client ID del App Registration
- SP_CLIENT_SECRET : Client Secret del App Registration
- SP_FILE_URL      : Link de "Compartir" del archivo Excel en SharePoint
                      (clic derecho sobre el archivo > Compartir > Copiar enlace)
"""

import pandas as pd
import subprocess
import sys
import os
import json
import base64
import time
import requests
import msal
from io import BytesIO
from datetime import datetime, timedelta
from pathlib import Path
from zoneinfo import ZoneInfo

# =========================================================
# CONFIGURACIÓN
# =========================================================

# --- Acceso al Excel vía SharePoint / Microsoft Graph API ---
SP_TENANT_ID = os.environ["SP_TENANT_ID"]
SP_CLIENT_ID = os.environ["SP_CLIENT_ID"]
SP_CLIENT_SECRET = os.environ["SP_CLIENT_SECRET"]
SP_FILE_URL = os.environ["SP_FILE_URL"]

GRAPH_AUTHORITY = f"https://login.microsoftonline.com/{SP_TENANT_ID}"
GRAPH_SCOPE = ["https://graph.microsoft.com/.default"]

HOJA_PRINCIPAL = "PPublicas"
HOJA_CHECKLIST = "Checklist"
FILA_ENCABEZADOS_PRINCIPAL = 2   # fila 2 en PPublicas
FILA_ENCABEZADOS_CHECKLIST = 2   # fila 2 en Checklist
DIAS_VENTANA = 20
ESTADOS_VALIDOS = ["Confirmada", "Sin confirmar"]

# Columnas de PPublicas
COL_ESTADO = "ESTADO"
COL_LINEA = "LINEA DE NEGOCIO"
COL_ID = "ID"
COL_FECHA_CIERRE = "FECHA DE CIERRE"
COL_CLIENTE = "CLIENTE LICITACIÓN"
COL_EQUIPAMIENTO = "EQUIPAMIENTO"
COL_VENDEDOR = "VENDEDOR"
COL_EMPRESA = "EMPRESA"

# Columnas de Checklist
COL_CK_ID = "ID Licitación"
COL_CK_REQUIERE = "Requiere garantía de seriedad (SI/NO)"
COL_CK_TIPO = "Tipo de garantía"
COL_CK_MONTO = "Monto garantía"
COL_CK_MONEDA = "Moneda garantía"
COL_CK_BANCO = "Banco/empresa emisora"
COL_CK_FECHA_INICIO = "Fecha inicio garantía"
COL_CK_FECHA_FIN = "Fecha fin garantía"

CARPETA_REPO = Path(__file__).parent
ARCHIVO_SALIDA = CARPETA_REPO / "index.html"
RUTA_LOGO_GEMCO = "PNG Empresas/Gemco-logo-blanco-v2.png"
RUTA_ISOTIPO_GEMCO = "PNG Empresas/Isotipos - copia.png"
RUTA_LOGO_INCARDIA = "PNG Empresas/INCARDIA-logo-blanco-v2.png"
RUTA_LOGO_MMQ = "PNG Empresas/MMQ-logo-blanco-v2.png"

FILAS_POR_SLIDE = 6
SEGUNDOS_POR_SLIDE = 12
SEGUNDOS_ISOTIPO = 5

INTENTOS_LECTURA_EXCEL = 3
ESPERA_ENTRE_INTENTOS_SEG = 5

ZONA_CHILE = ZoneInfo("America/Santiago")


def ahora_chile():
    """Hora actual en Chile (naive, sin tzinfo) para que se pueda comparar
    directamente con las fechas del Excel, que tampoco tienen tzinfo.
    El runner de GitHub Actions corre en UTC, por eso no alcanza con
    datetime.now()."""
    return datetime.now(ZONA_CHILE).replace(tzinfo=None)


def normalizar(valor):
    """Convierte a texto, mayúsculas, sin espacios extra. Maneja None."""
    if valor is None:
        return ""
    return str(valor).strip().upper()


def valor_o_guion(valor, defecto="—"):
    """Texto limpio del valor, o el valor por defecto si está vacío/NaN.
    Ojo: `valor or defecto` no sirve para esto porque NaN es "truthy" en Python
    (str(NaN) da "nan" en vez de caer al default)."""
    if pd.isna(valor):
        return defecto
    texto = str(valor).strip()
    return texto if texto else defecto


def _obtener_token_graph():
    """Autentica contra Entra ID vía MSAL (Client Credentials Flow) y devuelve un access token para Graph API."""
    app = msal.ConfidentialClientApplication(
        SP_CLIENT_ID, authority=GRAPH_AUTHORITY, client_credential=SP_CLIENT_SECRET
    )
    resultado = app.acquire_token_for_client(scopes=GRAPH_SCOPE)
    if "access_token" not in resultado:
        raise RuntimeError(
            f"No se pudo obtener token de Graph API: "
            f"{resultado.get('error')} - {resultado.get('error_description')}"
        )
    return resultado["access_token"]


def _codificar_url_para_graph(url):
    """Codifica una URL de SharePoint en el formato 'sharing token' que espera el endpoint /shares de Graph.
    Ref: https://learn.microsoft.com/graph/api/shares-get"""
    b64 = base64.urlsafe_b64encode(url.encode("utf-8")).decode("utf-8").rstrip("=")
    return "u!" + b64


def descargar_excel_desde_sharepoint():
    """Descarga el Excel desde SharePoint vía Microsoft Graph API y lo devuelve en memoria (BytesIO),
    sin escribirlo a disco."""
    token = _obtener_token_graph()
    headers = {"Authorization": f"Bearer {token}"}
    share_id = _codificar_url_para_graph(SP_FILE_URL)

    resp = requests.get(
        f"https://graph.microsoft.com/v1.0/shares/{share_id}/driveItem",
        headers=headers, timeout=30,
    )
    resp.raise_for_status()
    drive_item = resp.json()

    download_url = drive_item.get("@microsoft.graph.downloadUrl")
    if download_url:
        # Link prefirmado directo al contenido: no necesita el header de autorización.
        contenido = requests.get(download_url, timeout=60)
    else:
        # Fallback: pedir el contenido explícitamente vía Graph si no vino el downloadUrl.
        drive_id = drive_item["parentReference"]["driveId"]
        item_id = drive_item["id"]
        contenido = requests.get(
            f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/content",
            headers=headers, timeout=60,
        )
    contenido.raise_for_status()
    return BytesIO(contenido.content)


def leer_y_cruzar_licitaciones():
    """Descarga el Excel desde SharePoint, lee PPublicas y Checklist, cruza por ID, filtra y devuelve confirmadas/sin_confirmar."""
    print("Descargando Excel desde SharePoint (Graph API)...")
    df = None
    df_ck = None
    for intento in range(1, INTENTOS_LECTURA_EXCEL + 1):
        try:
            excel_bytes = descargar_excel_desde_sharepoint()
            libro = pd.ExcelFile(excel_bytes, engine="openpyxl")
            df = pd.read_excel(libro, sheet_name=HOJA_PRINCIPAL, header=FILA_ENCABEZADOS_PRINCIPAL - 1)
            df_ck = pd.read_excel(libro, sheet_name=HOJA_CHECKLIST, header=FILA_ENCABEZADOS_CHECKLIST - 1)
            break
        except requests.exceptions.RequestException as e:
            if intento < INTENTOS_LECTURA_EXCEL:
                print(f"Error de red/Graph API al descargar el Excel (intento {intento}/{INTENTOS_LECTURA_EXCEL}): {e}. "
                      f"Reintentando en {ESPERA_ENTRE_INTENTOS_SEG}s...")
                time.sleep(ESPERA_ENTRE_INTENTOS_SEG)
            else:
                print("ERROR: no se pudo descargar el Excel desde SharePoint tras varios intentos.")
                print(f"Detalle: {e}")
                sys.exit(1)

    ahora = ahora_chile()
    limite = ahora + timedelta(days=DIAS_VENTANA)

    # Filtro principal sobre PPublicas
    df = df[df[COL_ESTADO].isin(ESTADOS_VALIDOS)]
    df = df[df[COL_FECHA_CIERRE].notna()]
    df[COL_FECHA_CIERRE] = pd.to_datetime(df[COL_FECHA_CIERRE], errors="coerce")
    df = df[(df[COL_FECHA_CIERRE] >= ahora) & (df[COL_FECHA_CIERRE] <= limite)]
    df = df.sort_values(COL_FECHA_CIERRE)

    # Normalizar clave de cruce en ambos lados
    df["_id_norm"] = df[COL_ID].astype(str).str.strip()
    df_ck["_id_norm"] = df_ck[COL_CK_ID].astype(str).str.strip()

    # Cruce tipo BUSCARV: left join, nos quedamos con la primera coincidencia por ID
    df_ck_unico = df_ck.drop_duplicates(subset="_id_norm", keep="first")
    df = df.merge(
        df_ck_unico[[
            "_id_norm", COL_CK_REQUIERE, COL_CK_TIPO, COL_CK_MONTO,
            COL_CK_MONEDA, COL_CK_BANCO, COL_CK_FECHA_INICIO, COL_CK_FECHA_FIN
        ]],
        on="_id_norm", how="left"
    )

    registros = []
    for _, fila in df.iterrows():
        requiere_norm = normalizar(fila.get(COL_CK_REQUIERE))
        requiere_garantia = requiere_norm == "SI"

        garantia = None
        if requiere_garantia:
            f_ini = fila.get(COL_CK_FECHA_INICIO)
            f_fin = fila.get(COL_CK_FECHA_FIN)
            garantia = {
                "requiere": requiere_norm,
                "tipo": valor_o_guion(fila.get(COL_CK_TIPO)),
                "monto": None if pd.isna(fila.get(COL_CK_MONTO)) else float(fila.get(COL_CK_MONTO)),
                "moneda": valor_o_guion(fila.get(COL_CK_MONEDA), ""),
                "banco": valor_o_guion(fila.get(COL_CK_BANCO)),
                "fecha_inicio": f_ini.strftime("%Y-%m-%d") if pd.notna(f_ini) else None,
                "fecha_fin": f_fin.strftime("%Y-%m-%d") if pd.notna(f_fin) else None,
            }

        empresa_norm = normalizar(fila.get(COL_EMPRESA))
        if empresa_norm in ("INCARDIA", "MMQ"):
            empresa = empresa_norm
        else:
            empresa = "GEMCO"

        registros.append({
            "id": str(fila.get(COL_ID, "")),
            "linea": valor_o_guion(fila.get(COL_LINEA)),
            "cliente": valor_o_guion(fila.get(COL_CLIENTE), ""),
            "equipamiento": valor_o_guion(fila.get(COL_EQUIPAMIENTO), ""),
            "fecha": fila[COL_FECHA_CIERRE].strftime("%Y-%m-%dT%H:%M"),
            "vendedor": valor_o_guion(fila.get(COL_VENDEDOR)),
            "estado": fila[COL_ESTADO],
            "empresa": empresa,
            "garantia": garantia,
        })

    confirmadas = [r for r in registros if r["estado"] == "Confirmada"]
    sin_confirmar = [r for r in registros if r["estado"] == "Sin confirmar"]
    con_garantia = len([r for r in registros if r["garantia"]])
    print(f"Confirmadas: {len(confirmadas)} | Sin confirmar: {len(sin_confirmar)} | Con garantía requerida: {con_garantia}")
    return confirmadas, sin_confirmar


def imagen_a_base64(ruta_relativa):
    """Convierte una imagen a base64 para incrustarla directo en el HTML (evita problemas de rutas)."""
    ruta_completa = CARPETA_REPO / ruta_relativa
    if not ruta_completa.exists():
        print(f"ADVERTENCIA: no se encontró la imagen {ruta_completa}. Se omite.")
        return ""
    datos = base64.b64encode(ruta_completa.read_bytes()).decode("utf-8")
    return f"data:image/png;base64,{datos}"


def generar_html(confirmadas, sin_confirmar):
    datos = {
        "confirmadas": confirmadas,
        "sin_confirmar": sin_confirmar,
        "generado": ahora_chile().strftime("%Y-%m-%d %H:%M"),
    }
    datos_json = json.dumps(datos, ensure_ascii=False)

    logo_b64 = imagen_a_base64(RUTA_LOGO_GEMCO)
    isotipo_b64 = imagen_a_base64(RUTA_ISOTIPO_GEMCO)
    logo_incardia_b64 = imagen_a_base64(RUTA_LOGO_INCARDIA)
    logo_mmq_b64 = imagen_a_base64(RUTA_LOGO_MMQ)

    html = HTML_TEMPLATE.replace("__DATOS_JSON__", datos_json)
    html = html.replace("__LOGO_GEMCO__", logo_b64)
    html = html.replace("__ISOTIPO_GEMCO__", isotipo_b64)
    html = html.replace("__LOGO_INCARDIA__", logo_incardia_b64)
    html = html.replace("__LOGO_MMQ__", logo_mmq_b64)
    html = html.replace("__FILAS_POR_SLIDE__", str(FILAS_POR_SLIDE))
    html = html.replace("__SEGUNDOS_POR_SLIDE__", str(SEGUNDOS_POR_SLIDE))
    html = html.replace("__SEGUNDOS_ISOTIPO__", str(SEGUNDOS_ISOTIPO))
    ARCHIVO_SALIDA.write_text(html, encoding="utf-8")
    print(f"HTML generado en: {ARCHIVO_SALIDA}")


def subir_a_github():
    def correr(comando):
        resultado = subprocess.run(
            comando, cwd=CARPETA_REPO, capture_output=True, text=True, shell=True,
            encoding="utf-8", errors="replace"
        )
        print(resultado.stdout.strip())
        if resultado.returncode != 0:
            print("STDERR:", resultado.stderr.strip())
        return resultado.returncode

    print("Subiendo cambios a GitHub...")
    correr("git add index.html")
    mensaje = f"Actualizacion panel {ahora_chile().strftime('%Y-%m-%d %H:%M')}"
    codigo = correr(f'git commit -m "{mensaje}"')
    if codigo != 0:
        print("Nota: puede que no haya cambios nuevos que subir (esto no es un error).")
        return
    correr("git push")
    print("Listo. GitHub Pages va a reflejar el cambio en uno o dos minutos.")


HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<title>GEMCO · Licitaciones próximas</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=Oswald:wght@500;600;700&family=JetBrains+Mono:wght@500;600&display=swap" rel="stylesheet">
<style>
  :root{
    --bg:#f4f6fb; --card:#fff; --border:#e5e9f4; --text:#0f1a3c; --text-dim:#7c86a6;
    --header-bg-1:#0a1030; --header-bg-2:#2450e8;
    --brand:#2f6fed; --brand-soft:#e8f0fe; --teal:#0ea5b7; --teal-soft:#e1f7f9;
    --urgent:#ef4360; --urgent-soft:#fde7ec; --soon:#f59e0b; --soon-soft:#fef3e0;
    --ok:#22a06b; --shadow:0 4px 20px rgba(15,26,60,0.07);
    --font-display:'Oswald','Segoe UI',Arial,sans-serif; --font-body:'Inter','Segoe UI',Arial,sans-serif;
    --font-mono:'JetBrains Mono','Consolas',monospace;
  }
  *{box-sizing:border-box;margin:0;padding:0}
  html,body{height:100%;overflow:hidden}
  body{background:var(--bg);color:var(--text);font-family:var(--font-body);display:flex;flex-direction:column;cursor:pointer;user-select:none}

  header{
    display:flex;justify-content:space-between;align-items:center;
    padding:18px 40px;flex-shrink:0;position:relative;overflow:hidden;
    background:radial-gradient(circle at 12% -80%,rgba(66,133,244,.55),transparent 60%),linear-gradient(120deg,var(--header-bg-1),var(--header-bg-2));
  }
  .brand{display:flex;align-items:center;gap:16px;position:relative}
  .brand img{height:34px;width:auto;display:block}
  .brand .subtitle{color:#aeb9e8;font-size:13.5px;font-family:var(--font-body);border-left:1px solid rgba(255,255,255,.18);padding-left:16px}
  .meta{text-align:right;color:#aeb9e8;font-size:12.5px;font-family:var(--font-mono);position:relative}
  .meta strong{color:#fff}

  .pause-indicator{
    position:fixed;top:20px;right:40px;z-index:50;
    display:none;align-items:center;gap:8px;
    background:rgba(10,16,48,0.92);color:#fff;
    padding:8px 16px;border-radius:20px;font-size:13px;font-weight:600;
    box-shadow:0 4px 16px rgba(0,0,0,0.2);
  }
  .pause-indicator.visible{display:flex}

  .main-content{flex:1;display:flex;flex-direction:column;padding:24px 40px 18px;min-height:0}

  .kpis{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-bottom:16px;flex-shrink:0}
  .kpi-card{background:var(--card);border:1px solid var(--border);border-radius:14px;padding:12px 16px;box-shadow:var(--shadow);display:flex;align-items:center;gap:12px}
  .kpi-card .icon{width:32px;height:32px;border-radius:9px;display:flex;align-items:center;justify-content:center;font-size:15px;flex-shrink:0}
  .kpi-card.brand .icon{background:var(--brand-soft);color:var(--brand)}
  .kpi-card.teal .icon{background:var(--teal-soft);color:var(--teal)}
  .kpi-card.urgent .icon{background:var(--urgent-soft);color:var(--urgent)}
  .kpi-card.soon .icon{background:var(--soon-soft);color:var(--soon)}
  .kpi-text{display:flex;flex-direction:column;gap:1px;min-width:0}
  .kpi-label{font-size:11.5px;color:var(--text-dim);font-weight:500;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
  .kpi-value{font-family:var(--font-display);font-size:19px;font-weight:600;line-height:1.15;letter-spacing:.2px}
  .kpi-sub{font-size:11px;color:var(--text-dim);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}

  .slide-header{display:flex;align-items:center;gap:12px;margin-bottom:12px;flex-shrink:0}
  .slide-header h2{font-family:var(--font-display);font-weight:600;font-size:18px;letter-spacing:.3px;text-transform:uppercase}
  .slide-header .badge{font-size:12px;font-weight:600;color:#fff;background:var(--brand);padding:3px 11px;border-radius:20px}
  .slide-header.pending .badge{background:var(--soon)}
  .slide-header .subtitle2{color:var(--text-dim);font-size:13px}

  .stage{position:relative;flex:1;min-height:0}
  .slide{position:absolute;inset:0;opacity:0;transform:translateY(14px);transition:opacity .6s ease,transform .6s ease;pointer-events:none}
  .slide.active{opacity:1;transform:translateY(0);pointer-events:auto}

  .card-table{background:var(--card);border:1px solid var(--border);border-radius:16px;box-shadow:var(--shadow);overflow:hidden;height:100%;display:flex;flex-direction:column;position:relative}
  .card-table::before{
    content:'';position:absolute;inset:0;z-index:0;pointer-events:none;
    background-image:url("__ISOTIPO_GEMCO__");background-repeat:no-repeat;
    background-position:right -80px bottom -80px;background-size:520px auto;
    filter:brightness(0);opacity:.045;
  }
  .card-table>table{position:relative;z-index:1}
  table{width:100%;border-collapse:collapse;table-layout:fixed}
  thead th{text-align:left;font-size:11.5px;letter-spacing:.4px;text-transform:uppercase;color:var(--text-dim);font-weight:600;padding:12px 18px;border-bottom:1px solid var(--border)}
  col.col-urg{width:5px}
  col.col-empresa{width:52px}
  col.col-id{width:11%}
  col.col-linea{width:8.5%}
  col.col-nombre{width:20%}
  col.col-equip{width:23.5%}
  col.col-fecha{width:16%}
  col.col-vendedor{width:16%}
  tbody tr{border-bottom:1px solid var(--border)} tbody tr:last-child{border-bottom:none}
  tbody td{padding:14px 18px;font-size:14.5px;vertical-align:middle;overflow:hidden}
  .col-urgencia{width:5px;padding:0}
  .urgencia-bar{width:5px;height:100%;display:block;min-height:40px;border-radius:3px}
  .urgencia-bar.urgent{background:var(--urgent)} .urgencia-bar.soon{background:var(--soon)} .urgencia-bar.ok{background:var(--ok)}
  .id-cell{font-family:var(--font-mono);font-size:12px;color:var(--text-dim);white-space:normal;word-break:break-all;line-height:1.3}
  .col-empresa{padding:8px 4px 8px 18px}
  .empresa-chip{width:44px;height:26px;border-radius:7px;background:#1c2440;display:flex;align-items:center;justify-content:center;overflow:hidden}
  .empresa-chip img{max-width:34px;max-height:18px;width:auto;height:auto;object-fit:contain}
  .linea-tag{display:inline-block;font-size:10.5px;font-weight:600;padding:4px 9px;border-radius:20px;background:var(--brand-soft);color:var(--brand);white-space:nowrap}
  .nombre{font-weight:600;line-height:1.3;display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden}
  .equipamiento{font-size:13px;color:var(--text);line-height:1.35;display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden}
  .fecha{font-family:var(--font-mono);font-size:13px;white-space:nowrap}
  .fecha .dias{display:block;font-size:11px;margin-top:3px;font-family:var(--font-body);font-weight:600}
  .dias.urgent{color:var(--urgent)} .dias.soon{color:var(--soon)} .dias.ok{color:var(--text-dim);font-weight:500}
  .vendedor{color:var(--text-dim);font-size:13px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
  tbody tr.filler{visibility:hidden}
  tbody tr.garantia-row td{background:#fdf7ec;padding:9px 18px;font-size:12px;color:#8a6d1f;border-bottom:1px solid var(--border)}
  tbody tr.garantia-row b{color:#6b4e0f}
  .garantia-grid{display:flex;flex-direction:column;gap:2px}
  .garantia-grid div{line-height:1.5}

  footer{flex-shrink:0;display:flex;justify-content:center;align-items:center;gap:8px;margin-top:16px}
  .progress-dot{width:30px;height:5px;border-radius:3px;background:var(--border);overflow:hidden;position:relative}
  .progress-dot .fill{position:absolute;inset:0;background:var(--brand);transform:scaleX(0);transform-origin:left}
  .progress-dot.done .fill{transform:scaleX(1);transition:none}
  .progress-dot.active .fill{transform:scaleX(0);animation:fillbar var(--dur,12s) linear forwards}
  .progress-dot.paused .fill{animation-play-state:paused}
  @keyframes fillbar{from{transform:scaleX(0)}to{transform:scaleX(1)}}

  .isotipo-slide{
    position:absolute;inset:0;opacity:0;pointer-events:none;
    transition:opacity .6s ease;
    display:flex;align-items:center;justify-content:center;
    background:var(--bg);
  }
  .isotipo-slide.active{opacity:1;pointer-events:auto}
  .isotipo-slide img{height:180px;width:auto;animation:pulse-isotipo 2.5s ease-in-out infinite}
  @keyframes pulse-isotipo{0%,100%{transform:scale(1)}50%{transform:scale(1.05)}}
</style>
</head>
<body>

<div class="pause-indicator" id="pause-indicator">⏸ Panel en pausa — clic o espacio para continuar</div>

<header>
  <div class="brand">
    <img src="__LOGO_GEMCO__" alt="GEMCO">
    <span class="subtitle">Licitaciones · Panel de seguimiento</span>
  </div>
  <div class="meta">Actualizado<br><strong id="ts-generado">—</strong></div>
</header>

<div class="main-content">
  <div class="kpis" id="kpis"></div>

  <div class="slide-header" id="slide-header">
    <h2 id="slide-title">—</h2>
    <span class="badge" id="slide-badge">0</span>
    <span class="subtitle2" id="slide-subtitle"></span>
  </div>

  <div class="stage" id="stage">
    <div class="isotipo-slide" id="isotipo-slide"><img src="__ISOTIPO_GEMCO__" alt="GEMCO"></div>
  </div>

  <footer id="progress"></footer>
</div>

<script>
const LICITACIONES_DATA = __DATOS_JSON__;
const FILAS_POR_SLIDE = __FILAS_POR_SLIDE__;
const SEGUNDOS_POR_SLIDE = __SEGUNDOS_POR_SLIDE__;
const SEGUNDOS_ISOTIPO = __SEGUNDOS_ISOTIPO__;
const EMPRESA_LOGOS = {
  GEMCO: '__LOGO_GEMCO__',
  INCARDIA: '__LOGO_INCARDIA__',
  MMQ: '__LOGO_MMQ__',
};

function urgenciaClase(f){const a=new Date(LICITACIONES_DATA.generado.replace(' ','T'));const d=(new Date(f)-a)/864e5;return d<=2?'urgent':d<=5?'soon':'ok'}
function formatFecha(f){return new Date(f).toLocaleString('es-CL',{day:'2-digit',month:'short',hour:'2-digit',minute:'2-digit'}).replace('.','')}
function diasRestantes(f){
  const ahora=new Date(LICITACIONES_DATA.generado.replace(' ','T'));
  const cierre=new Date(f);
  const inicioHoy=new Date(ahora.getFullYear(),ahora.getMonth(),ahora.getDate());
  const inicioCierre=new Date(cierre.getFullYear(),cierre.getMonth(),cierre.getDate());
  const d=Math.round((inicioCierre-inicioHoy)/864e5);
  return d<=0?'Cierra hoy':d===1?'Cierra mañana':`Cierra en ${d} días`;
}
function formatMontoGarantia(m,moneda){if(!m)return '';return (moneda==='UF'?'UF ':'$')+m.toLocaleString('es-CL')}

function construirKPIs(){
  const c=LICITACIONES_DATA.confirmadas,p=LICITACIONES_DATA.sin_confirmar;
  const u=[...c,...p].filter(x=>urgenciaClase(x.fecha)==='urgent');
  const conGarantia=[...c,...p].filter(x=>x.garantia);
  const cards=[
    {clase:'brand',icono:'✓',label:'Confirmadas',valor:c.length,sub:'en ventana de 20 días'},
    {clase:'soon',icono:'?',label:'Por confirmar',valor:p.length,sub:'pendientes de estado'},
    {clase:'urgent',icono:'!',label:'Cierran en ≤2 días',valor:u.length,sub:'requieren atención'},
    {clase:'teal',icono:'§',label:'Requieren garantía',valor:conGarantia.length,sub:'de seriedad de oferta'},
  ];
  document.getElementById('kpis').innerHTML=cards.map(x=>`<div class="kpi-card ${x.clase}"><div class="icon">${x.icono}</div><div class="kpi-text"><div class="kpi-label">${x.label}</div><div class="kpi-value">${x.valor}</div><div class="kpi-sub">${x.sub}</div></div></div>`).join('');
}

function campoValido(v){
  if(v===null||v===undefined) return false;
  const s=String(v).trim();
  if(s==='' || s==='—') return false;
  const su=s.toUpperCase();
  if(su==='N/A' || su==='N/D' || su==='NA' || su==='NAN') return false;
  return true;
}

function filaGarantia(g){
  if(!g) return '';
  const monto=formatMontoGarantia(g.monto,g.moneda);
  const campos=[
    ['REQUIERE GARANTÍA DE SERIEDAD (SI/NO)', g.requiere],
    ['TIPO DE GARANTÍA', g.tipo],
    ['MONTO GARANTÍA', monto],
    ['BANCO/EMPRESA EMISORA', g.banco],
    ['FECHA INICIO GARANTÍA', g.fecha_inicio],
    ['FECHA FIN GARANTÍA', g.fecha_fin],
  ];
  const lineas=campos.filter(([,v])=>campoValido(v)).map(([label,v])=>`<div><b>${label}:</b> ${v}</div>`);
  if(lineas.length===0) return '';
  return `<tr class="garantia-row"><td></td><td></td><td colspan="6"><div class="garantia-grid">${lineas.join('')}</div></td></tr>`;
}

function empresaChip(item){
  const src=item.empresa && EMPRESA_LOGOS[item.empresa];
  if(!src) return '';
  return `<span class="empresa-chip"><img src="${src}" alt="${item.empresa}"></span>`;
}

function construirTablaHTML(lista){
  const filas=lista.map(item=>{
    const urg=urgenciaClase(item.fecha);
    const filaGar=filaGarantia(item.garantia);
    return `<tr><td class="col-urgencia"><span class="urgencia-bar ${urg}"></span></td><td class="col-empresa">${empresaChip(item)}</td><td class="id-cell">${item.id}</td><td><span class="linea-tag">${item.linea}</span></td><td><div class="nombre">${item.cliente||'(sin cliente)'}</div></td><td><div class="equipamiento">${item.equipamiento||'—'}</div></td><td class="fecha">${formatFecha(item.fecha)}<span class="dias ${urg}">${diasRestantes(item.fecha)}</span></td><td class="vendedor">${item.vendedor}</td></tr>${filaGar}`;
  }).join('');
  let fillers='';
  for(let i=0;i<FILAS_POR_SLIDE-lista.length;i++) fillers+='<tr class="filler"><td></td><td></td><td></td><td></td><td>&nbsp;</td><td></td><td></td><td></td></tr>';
  return `<div class="card-table"><table><colgroup><col class="col-urg"><col class="col-empresa"><col class="col-id"><col class="col-linea"><col class="col-nombre"><col class="col-equip"><col class="col-fecha"><col class="col-vendedor"></colgroup><thead><tr><th></th><th></th><th>ID</th><th>Línea</th><th>Cliente</th><th>Equipamiento</th><th>Cierre</th><th>Vendedor</th></tr></thead><tbody>${filas}${fillers}</tbody></table></div>`;
}

function chunk(a,s){const o=[];for(let i=0;i<a.length;i+=s)o.push(a.slice(i,i+s));return o}

const cC=chunk(LICITACIONES_DATA.confirmadas,FILAS_POR_SLIDE);
const cP=chunk(LICITACIONES_DATA.sin_confirmar,FILAS_POR_SLIDE);
const slidesDatos=[];
cC.forEach((c,i)=>slidesDatos.push({tipo:'confirmadas',datos:c,titulo:'Confirmadas',parte:cC.length>1?`${i+1}/${cC.length}`:null,total:LICITACIONES_DATA.confirmadas.length}));
cP.forEach((c,i)=>slidesDatos.push({tipo:'pendientes',datos:c,titulo:'Por confirmar',parte:cP.length>1?`${i+1}/${cP.length}`:null,total:LICITACIONES_DATA.sin_confirmar.length}));

const stage=document.getElementById('stage'),slideHeader=document.getElementById('slide-header'),slideTitle=document.getElementById('slide-title'),slideBadge=document.getElementById('slide-badge'),slideSubtitle=document.getElementById('slide-subtitle'),progress=document.getElementById('progress'),isotipoSlide=document.getElementById('isotipo-slide'),pauseIndicator=document.getElementById('pause-indicator');

construirKPIs();
slidesDatos.forEach((s,i)=>{const d=document.createElement('div');d.className='slide';d.id=`slide-${i}`;d.innerHTML=construirTablaHTML(s.datos);stage.insertBefore(d,isotipoSlide)});

// secuencia total = slides de datos + 1 slide isotipo al final (índice = slidesDatos.length)
const TOTAL_PASOS=slidesDatos.length+1;
const ES_ISOTIPO=(idx)=>idx===slidesDatos.length;

slidesDatos.forEach((s,i)=>{const d=document.createElement('div');d.className='progress-dot';d.id=`dot-${i}`;d.innerHTML='<span class="fill"></span>';progress.appendChild(d)});
const dotIsotipo=document.createElement('div');dotIsotipo.className='progress-dot';dotIsotipo.id=`dot-${slidesDatos.length}`;dotIsotipo.innerHTML='<span class="fill"></span>';progress.appendChild(dotIsotipo);

let current=0;
let pausado=false;
let timer=null;

function duracion(idx){ return ES_ISOTIPO(idx) ? SEGUNDOS_ISOTIPO : SEGUNDOS_POR_SLIDE; }

function showSlide(idx){
  slidesDatos.forEach((s,i)=>{
    const el=document.getElementById(`slide-${i}`);
    if(el) el.classList.toggle('active',i===idx);
  });
  isotipoSlide.classList.toggle('active', ES_ISOTIPO(idx));

  for(let i=0;i<TOTAL_PASOS;i++){
    const dot=document.getElementById(`dot-${i}`);
    dot.classList.remove('active');
    dot.classList.toggle('done', i<idx);
    dot.style.removeProperty('--dur');
  }
  const dotActive=document.getElementById(`dot-${idx}`);
  dotActive.classList.remove('done');
  dotActive.style.setProperty('--dur', duracion(idx)+'s');
  void dotActive.offsetWidth;
  dotActive.classList.add('active');
  dotActive.classList.toggle('paused', pausado);

  if(!ES_ISOTIPO(idx)){
    const s=slidesDatos[idx];
    slideTitle.textContent=s.titulo;
    slideBadge.textContent=s.total;
    slideSubtitle.textContent=s.parte?`Pantalla ${s.parte}`:'';
    slideHeader.classList.toggle('pending',s.tipo==='pendientes');
    slideHeader.style.visibility='visible';
  } else {
    slideHeader.style.visibility='hidden';
  }
}

function iniciarTimer(){
  clearTimeout(timer);
  timer=setTimeout(avanzar, duracion(current)*1000);
}

function irA(idx){
  current=((idx%TOTAL_PASOS)+TOTAL_PASOS)%TOTAL_PASOS;
  showSlide(current);
  if(pausado){ clearTimeout(timer); } else { iniciarTimer(); }
}

function avanzar(){ irA(current+1); }
function siguiente(){ irA(current+1); }
function anterior(){ irA(current-1); }

function togglePausa(){
  pausado=!pausado;
  pauseIndicator.classList.toggle('visible',pausado);
  const dotActive=document.getElementById(`dot-${current}`);
  dotActive.classList.toggle('paused',pausado);
  if(pausado){
    clearTimeout(timer);
  } else {
    iniciarTimer();
  }
}

document.body.addEventListener('click', togglePausa);
document.addEventListener('keydown', (e)=>{
  if(e.code==='Space'){ e.preventDefault(); togglePausa(); }
  else if(e.code==='ArrowRight'){ e.preventDefault(); siguiente(); }
  else if(e.code==='ArrowLeft'){ e.preventDefault(); anterior(); }
});

if(TOTAL_PASOS>0){
  showSlide(0);
  iniciarTimer();
}
document.getElementById('ts-generado').textContent=LICITACIONES_DATA.generado;

// Auto-refresh cada 30 min para levantar datos nuevos (se salta si está pausado)
setInterval(()=>{ if(!pausado) location.reload(); }, 30*60*1000);
</script>

</body>
</html>
"""


if __name__ == "__main__":
    confirmadas, sin_confirmar = leer_y_cruzar_licitaciones()
    generar_html(confirmadas, sin_confirmar)
    subir_a_github()
