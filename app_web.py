"""
Extractor de Documentación Contable — La Nación
App web Streamlit — diseño minimalista corporativo.
"""

import io
import re
import base64
from datetime import datetime

import pdfplumber
import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ══════════════════════════════════════════════════════════════════════
#  LOGO EMBEBIDO
# ══════════════════════════════════════════════════════════════════════

def get_logo_b64():
    for ruta in ["La_Nacion_Logo.png", "/mnt/user-data/uploads/La_Nacion_Logo.png"]:
        try:
            with open(ruta, "rb") as f:
                return base64.b64encode(f.read()).decode()
        except Exception:
            pass
    return ""


# ══════════════════════════════════════════════════════════════════════
#  EXTRACCIÓN DE CAMPOS DESDE PDF
# ══════════════════════════════════════════════════════════════════════

def extraer_emisor(texto):
    lineas = [l.strip() for l in texto.split('\n') if l.strip()]
    m = re.search(r'[Rr]az[oó]n\s+social[:\s]+([^\n\r]{3,60})', texto)
    if m:
        val = m.group(1).strip()
        if val and 'LA NACION' not in val.upper():
            return val
    keywords = ['S.A.', 'SA ', 'S.R.L.', 'SRL', 'CARGO', 'TRANSPORTES', 'SERVICIOS', 'AEROLINEAS', 'HANDYWAY', 'CRUZ DEL SUR']
    for linea in lineas[:20]:
        if any(kw in linea.upper() for kw in keywords) and 'LA NACION' not in linea.upper() and len(linea) > 5:
            return linea.strip()
    return ""


def extraer_fecha(texto):
    m = re.search(r'[Ff]echa[:\s]+(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4})', texto)
    if m:
        return m.group(1)
    m = re.search(r'(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{4})', texto)
    return m.group(1) if m else ""


def extraer_numero_factura(texto):
    patrones = [
        r'(?:Comp\. Nro|Comprobante N[°º]?|Factura N[°º]?)[:\s]*([A-Z0-9\-]+)',
        r'([0-9]{4}-[0-9]{8})',
        r'(?:N[°º]|Nro\.?)[:\s]*([A-Z0-9\-\/]+)',
    ]
    for p in patrones:
        m = re.search(p, texto)
        if m:
            return m.group(1).strip()
    return ""


def extraer_cuit(texto):
    m = re.search(r'(?:CUIT|C\.U\.I\.T\.)[:\s]*(\d{2}[-\s]?\d{8}[-\s]?\d)', texto)
    return m.group(1).replace(' ', '').replace('-', '') if m else ""


def limpiar_monto(texto_monto):
    texto_monto = texto_monto.replace('\xa0', '').strip()
    texto_monto = re.sub(r'[^\d,\.]', '', texto_monto)
    if ',' in texto_monto and '.' in texto_monto:
        if texto_monto.rfind(',') > texto_monto.rfind('.'):
            texto_monto = texto_monto.replace('.', '').replace(',', '.')
        else:
            texto_monto = texto_monto.replace(',', '')
    elif ',' in texto_monto:
        partes = texto_monto.split(',')
        if len(partes) == 2 and len(partes[1]) <= 2:
            texto_monto = texto_monto.replace(',', '.')
        else:
            texto_monto = texto_monto.replace(',', '')
    try:
        return float(texto_monto)
    except Exception:
        return 0.0


def extraer_montos(texto):
    neto, iva, total = 0.0, 0.0, 0.0
    m = re.search(r'(?:subtotal|base imponible|neto grabado|importe neto)[:\s$]*([0-9.,\xa0]+)', texto, re.IGNORECASE)
    if m:
        neto = limpiar_monto(m.group(1))
    m = re.search(r'(?:I\.?V\.?A\.?|impuesto)[:\s$]*(?:\d+[\.,]?\d*\s*%\s*)?[:\s$]*([0-9.,\xa0]+)', texto, re.IGNORECASE)
    if m:
        iva = limpiar_monto(m.group(1))
    patrones_total = [
        r'(?:total\s+a\s+pagar|importe\s+total|total\s+factura)[:\s$]*([0-9.,\xa0]+)',
        r'(?:^|\n)\s*total[:\s$]+([0-9.,\xa0]+)',
    ]
    for p in patrones_total:
        m = re.search(p, texto, re.IGNORECASE | re.MULTILINE)
        if m:
            total = limpiar_monto(m.group(1))
            break
    return neto, iva, total


def extraer_kilos(texto):
    """Extrae kilogramos aforados desde el texto de la factura."""
    patrones = [
        r'(?:kilos?|kg\.?|peso)\s+aforados?[:\s]*([0-9.,]+)',
        r'(?:kilos?|kg\.?)\s*[:\-]\s*([0-9.,]+)',
        r'([0-9.,]+)\s*(?:kilos?|kg)',
    ]
    for p in patrones:
        m = re.search(p, texto, re.IGNORECASE)
        if m:
            try:
                return float(m.group(1).replace(',', '.'))
            except Exception:
                pass
    return 0.0


def extraer_tramo(texto):
    """Detecta origen-destino del tramo desde el texto."""
    m = re.search(r'(?:tramo|trayecto|ruta)[:\s]*([^\n\r]{5,50})', texto, re.IGNORECASE)
    return m.group(1).strip() if m else ""


def extraer_tipo_despacho(texto):
    for t in ['ENVIO', 'ENVÍO', 'DEVOLUCION', 'DEVOLUCIÓN', 'REPOSICION', 'REPOSICIÓN']:
        if t in texto.upper():
            return t.replace('Í', 'I').replace('Ó', 'O')
    return ""


def extraer_datos(pdf_bytes, nombre_archivo):
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            texto = "\n".join(p.extract_text() or "" for p in pdf.pages)
    except Exception as e:
        return {"error": str(e), "archivo": nombre_archivo}

    neto, iva, total = extraer_montos(texto)
    kilos = extraer_kilos(texto)

    return {
        "archivo":          nombre_archivo,
        "fecha":            extraer_fecha(texto),
        "origen":           "",
        "destino":          "",
        "kilos_aforados":   kilos,
        "clase":            "",
        "fac_total":        total if total else (neto + iva),
        "mes":              "",
        "proveedor":        extraer_emisor(texto),
        "tramo":            extraer_tramo(texto),
        "tipo_despacho":    extraer_tipo_despacho(texto),
        "año":              datetime.now().year,
        "numero_factura":   extraer_numero_factura(texto),
        "cuit":             extraer_cuit(texto),
        "neto":             neto,
        "iva":              iva,
        "error":            "",
    }


# ══════════════════════════════════════════════════════════════════════
#  GENERACIÓN DEL EXCEL (formato modelo de tabla de kilos)
# ══════════════════════════════════════════════════════════════════════

COLS = [
    ("Fecha",              18),
    ("Origen",             20),
    ("Destino",            20),
    ("kilos aforados",     16),
    ("Clase",              10),
    ("Fac Total",          16),
    ("Mes",                8),
    ("PROVEEDOR",          20),
    ("TRAMO",              35),
    ("TIPO DE DESPACHO",   20),
    ("AÑO",                8),
    ("NUMERO DE FACTURA",  22),
    ("PRECIO POR KILO",    18),
]

HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
HEADER_FONT = Font(name="Arial", bold=True, color="FFFFFF", size=10)
HEADER_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)

STRIPE_FILL = PatternFill("solid", fgColor="D6E4F0")
NORMAL_FILL = PatternFill("solid", fgColor="FFFFFF")

BORDER = Border(
    left=Side(style="thin", color="BDC3C7"),
    right=Side(style="thin", color="BDC3C7"),
    top=Side(style="thin", color="BDC3C7"),
    bottom=Side(style="thin", color="BDC3C7"),
)

DATA_FONT = Font(name="Arial", size=10)
DATA_ALIGN_CENTER = Alignment(horizontal="center", vertical="center")
DATA_ALIGN_LEFT   = Alignment(horizontal="left",   vertical="center")
DATA_ALIGN_RIGHT  = Alignment(horizontal="right",  vertical="center")


def generar_excel(registros):
    wb = Workbook()
    ws = wb.active
    ws.title = "Reporte de Kilos"

    # Encabezados
    for col_idx, (nombre, ancho) in enumerate(COLS, start=1):
        cell = ws.cell(row=1, column=col_idx, value=nombre)
        cell.font   = HEADER_FONT
        cell.fill   = HEADER_FILL
        cell.alignment = HEADER_ALIGN
        cell.border = BORDER
        ws.column_dimensions[get_column_letter(col_idx)].width = ancho

    ws.row_dimensions[1].height = 30

    # Filas de datos
    for row_idx, d in enumerate(registros, start=2):
        fill = STRIPE_FILL if row_idx % 2 == 0 else NORMAL_FILL

        fac_total   = d.get("fac_total", 0) or 0
        kilos       = d.get("kilos_aforados", 0) or 0
        col_fac     = get_column_letter(6)   # F → Fac Total
        col_kilos   = get_column_letter(4)   # D → kilos aforados
        col_precio  = get_column_letter(13)  # M → PRECIO POR KILO

        valores = [
            d.get("fecha", ""),
            d.get("origen", ""),
            d.get("destino", ""),
            kilos if kilos else "",
            d.get("clase", ""),
            fac_total if fac_total else "",
            d.get("mes", ""),
            d.get("proveedor", ""),
            d.get("tramo", ""),
            d.get("tipo_despacho", ""),
            d.get("año", ""),
            d.get("numero_factura", ""),
            None,  # placeholder → se pone fórmula abajo
        ]

        for col_idx, valor in enumerate(valores, start=1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.fill   = fill
            cell.border = BORDER
            cell.font   = DATA_FONT

            if col_idx == 13:
                # Fórmula: Fac Total / kilos aforados (con IFERROR para evitar #DIV/0!)
                cell.value = f"=IFERROR({col_fac}{row_idx}/{col_kilos}{row_idx},\"\")"
                cell.alignment = DATA_ALIGN_RIGHT
                cell.number_format = '#,##0.00'
            elif col_idx in (4, 7, 11):  # kilos, Mes, AÑO → número centrado
                cell.value = valor
                cell.alignment = DATA_ALIGN_CENTER
            elif col_idx in (6,):        # Fac Total → moneda derecha
                cell.value = valor
                cell.alignment = DATA_ALIGN_RIGHT
                cell.number_format = '#,##0.00'
            else:
                cell.value = valor
                cell.alignment = DATA_ALIGN_LEFT

    # Fila de totales
    total_row = len(registros) + 2
    ws.cell(row=total_row, column=1, value="TOTAL").font = Font(name="Arial", bold=True, size=10)
    ws.cell(row=total_row, column=1).alignment = DATA_ALIGN_CENTER

    col_kilos_letra  = get_column_letter(4)
    col_fac_letra    = get_column_letter(6)
    data_start = 2
    data_end   = len(registros) + 1

    sum_kilos = ws.cell(row=total_row, column=4,
                        value=f"=SUM({col_kilos_letra}{data_start}:{col_kilos_letra}{data_end})")
    sum_kilos.font   = Font(name="Arial", bold=True, size=10)
    sum_kilos.alignment = DATA_ALIGN_CENTER
    sum_kilos.fill   = PatternFill("solid", fgColor="2E75B6")
    sum_kilos.font   = Font(name="Arial", bold=True, color="FFFFFF", size=10)
    sum_kilos.border = BORDER

    sum_fac = ws.cell(row=total_row, column=6,
                      value=f"=SUM({col_fac_letra}{data_start}:{col_fac_letra}{data_end})")
    sum_fac.font          = Font(name="Arial", bold=True, color="FFFFFF", size=10)
    sum_fac.fill          = PatternFill("solid", fgColor="2E75B6")
    sum_fac.alignment     = DATA_ALIGN_RIGHT
    sum_fac.number_format = '#,##0.00'
    sum_fac.border        = BORDER

    # Precio promedio por kilo en total_row col 13
    precio_prom = ws.cell(row=total_row, column=13,
                          value=f"=IFERROR({col_fac_letra}{total_row}/{col_kilos_letra}{total_row},\"\")")
    precio_prom.font          = Font(name="Arial", bold=True, color="FFFFFF", size=10)
    precio_prom.fill          = PatternFill("solid", fgColor="2E75B6")
    precio_prom.alignment     = DATA_ALIGN_RIGHT
    precio_prom.number_format = '#,##0.00'
    precio_prom.border        = BORDER

    ws.freeze_panes = "A2"

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()


# ══════════════════════════════════════════════════════════════════════
#  INTERFAZ STREAMLIT
# ══════════════════════════════════════════════════════════════════════

st.set_page_config(
    page_title="Extractor Contable · La Nación",
    page_icon="📄",
    layout="centered",
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600&display=swap');

html, body, [class*="css"] {
    font-family: 'Inter', sans-serif !important;
    background-color: #F4F7FB !important;
}
.ln-title {
    font-size: 1.55rem;
    font-weight: 600;
    color: #1F4E79;
    line-height: 1.35;
    letter-spacing: -0.01em;
}
.ln-subtitle {
    font-size: 0.9rem;
    color: #6B7A99;
    margin-top: 4px;
}
.ln-logo-wrap {
    display: flex;
    justify-content: center;
    margin: 18px 0 10px;
}
.ln-logo-wrap img {
    height: 48px;
    object-fit: contain;
}
hr.ln-rule {
    border: none;
    border-top: 1.5px solid #D0DAF0;
    margin: 18px 0 22px;
}
.ln-upload-lbl {
    font-size: 0.88rem;
    font-weight: 500;
    color: #1F4E79;
    margin-bottom: 4px;
}
div[data-testid="stButton"] > button {
    background-color: #1F4E79 !important;
    color: white !important;
    border: none !important;
    border-radius: 6px !important;
    font-weight: 500 !important;
    font-size: 0.9rem !important;
    padding: 10px 24px !important;
    transition: background 0.2s;
}
div[data-testid="stButton"] > button:hover {
    background-color: #254D8F !important;
}
div[data-testid="stDownloadButton"] > button {
    background-color: #1B3A6B !important;
    color: white !important;
    border: none !important;
    border-radius: 6px !important;
    font-weight: 500 !important;
}
.ln-footer {
    text-align: center;
    font-size: 0.72rem;
    color: #B0BAD0;
    margin-top: 40px;
}
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div style="text-align:center; padding: 32px 0 0;">
  <p class="ln-title">Extractor de datos de documentación<br>contable de La Nación</p>
  <p class="ln-subtitle">Procesamiento automático de facturas y comprobantes</p>
</div>
""", unsafe_allow_html=True)

logo_b64 = get_logo_b64()
if logo_b64:
    st.markdown(f"""
    <div class="ln-logo-wrap">
      <img src="data:image/png;base64,{logo_b64}" alt="La Nación"/>
    </div>
    """, unsafe_allow_html=True)
else:
    st.markdown('<div style="height:24px"></div>', unsafe_allow_html=True)

st.markdown('<hr class="ln-rule">', unsafe_allow_html=True)

st.markdown('<p class="ln-upload-lbl">📎 &nbsp; Adjunte su documento</p>', unsafe_allow_html=True)

archivos = st.file_uploader(
    "PDFs",
    type=["pdf"],
    accept_multiple_files=True,
    label_visibility="collapsed",
    help="Puede seleccionar múltiples archivos PDF a la vez.",
)

if archivos:
    n = len(archivos)
    st.caption(f"✔  {n} archivo{'s' if n > 1 else ''} seleccionado{'s' if n > 1 else ''}.")
else:
    st.caption("Formatos admitidos: PDF · Factura electrónica AFIP")

st.markdown("<br>", unsafe_allow_html=True)

procesar = st.button(
    "Procesar documentos",
    disabled=not archivos,
    use_container_width=True,
)

if procesar and archivos:
    registros, resultados_ui = [], []
    total = len(archivos)
    prog  = st.progress(0, text="Iniciando...")

    for i, archivo in enumerate(archivos):
        prog.progress(i / total, text=f"Procesando {archivo.name}…")
        datos = extraer_datos(archivo.read(), archivo.name)
        registros.append(datos)
        ok = not datos.get("error")
        resultados_ui.append((archivo.name, ok, datos.get("error", ""), datos))

    prog.progress(1.0, text="Completado.")

    st.markdown("---")
    st.subheader("Resultados")

    exitosos = [d for _, ok, _, d in resultados_ui if ok]
    fallidos  = [(n, e) for n, ok, e, _ in resultados_ui if not ok]

    for nombre, ok, error, datos in resultados_ui:
        if ok:
            st.success(f"✅ **{nombre}**")
            cols = st.columns(4)
            cols[0].metric("Proveedor",       datos.get("proveedor", "—") or "—")
            cols[1].metric("N° Factura",      datos.get("numero_factura", "—") or "—")
            cols[2].metric("Total",           f"$ {datos.get('fac_total', 0):,.2f}" if datos.get('fac_total') else "—")
            cols[3].metric("Kilos aforados",  f"{datos.get('kilos_aforados', 0):,.0f} kg" if datos.get('kilos_aforados') else "—")
        else:
            st.error(f"❌ **{nombre}** — {error}")

    if exitosos:
        st.markdown("---")
        excel_bytes = generar_excel(exitosos)
        fecha_hoy = datetime.now().strftime("%Y%m%d")
        st.download_button(
            label="⬇️  Descargar Excel",
            data=excel_bytes,
            file_name=f"reporte_kilos_{fecha_hoy}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    if fallidos:
        st.warning(f"⚠️  {len(fallidos)} archivo(s) no pudieron procesarse.")

st.markdown('<div class="ln-footer">La Nación · Documentación Contable · Uso interno</div>', unsafe_allow_html=True)
