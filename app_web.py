"""
Extractor de Documentación Contable — La Nación
App web Streamlit — diseño minimalista corporativo.
"""

import io
import re
import base64
from pathlib import Path
from datetime import datetime

import pdfplumber
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ══════════════════════════════════════════════════════════════════════
#  LOGO EMBEBIDO
# ══════════════════════════════════════════════════════════════════════

def get_logo_b64():
    # Logo embebido en base64 para Streamlit Cloud
    return "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"


# ══════════════════════════════════════════════════════════════════════
#  EXTRACCIÓN OPTIMIZADA PARA FACTURAS DE LA NACIÓN
# ══════════════════════════════════════════════════════════════════════

def extraer_emisor(texto):
    lineas = [l.strip() for l in texto.split('\n') if l.strip()]

    # Razón social explícita (HandyWay style)
    m = re.search(r'[Rr]az[oó]n\s+social[:\s]+([^\n\r]{3,60})', texto)
    if m:
        val = m.group(1).strip()
        if val and 'LA NACION' not in val.upper() and 'NACION' not in val.upper():
            return val

    # Empresa en primeras líneas con S.A., CARGO, TRANSPORTES, etc.
    keywords = ['S.A.', 'SA ', ' SA\n', 'S.R.L.', 'SRL', 'CARGO', 'TRANSPORTES', 'SERVICIOS']
    for linea in lineas[:20]:
        tiene_kw = any(kw in linea.upper() for kw in keywords)
        if tiene_kw and 'LA NACION' not in linea.upper() and len(linea) > 5:
            return linea.strip()

    return ""


def extraer_fecha(texto):
    # "Fecha: DD/MM/YYYY"
    m = re.search(r'[Ff]echa[:\s]+(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4})', texto)
    if m:
        return m.group(1).strip()

    # Aerolíneas: "06 03 2026" (tres bloques separados en el encabezado)
    m = re.search(r'\b(\d{2})\s+(\d{2})\s+(20\d{2})\b', texto)
    if m:
        return f"{m.group(1)}/{m.group(2)}/{m.group(3)}"

    # Genérico DD/MM/YYYY
    m = re.search(r'\b(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](20\d{2})\b', texto)
    if m:
        return f"{m.group(1)}/{m.group(2)}/{m.group(3)}"

    return ""


def parsear_monto(t):
    try:
        t = t.strip().replace(' ', '')
        if re.match(r'^\d{1,3}(\.\d{3})+(,\d{1,2})$', t):
            return float(t.replace('.', '').replace(',', '.'))
        if re.match(r'^\d{1,3}(,\d{3})+(\.\d{1,2})$', t):
            return float(t.replace(',', ''))
        if re.match(r'^\d+(,\d{1,2})$', t):
            return float(t.replace(',', '.'))
        return float(t.replace(',', ''))
    except Exception:
        return None


def formatear_monto(valor):
    try:
        partes = f"{valor:,.2f}".split('.')
        entero = partes[0].replace(',', '.')
        return f"$ {entero},{partes[1]}"
    except Exception:
        return str(valor)


def extraer_importe(texto):
    patrones = [
        r'[Ii]mporte\s+[Tt]otal\s+\$?\s*([\d\.,]+)',
        r'TOTAL\s+EN\s+PESOS\s+([\d\.,]+)',
        r'\bTOTAL\s+\$?\s*([\d\.,]+)',
        r'[Tt]otal\s*\$\s*([\d\.,]+)',
    ]
    candidatos = []
    for pat in patrones:
        for m in re.finditer(pat, texto):
            val = parsear_monto(m.group(1))
            if val and val > 100:
                candidatos.append(val)

    if not candidatos:
        return ""
    return formatear_monto(max(candidatos))


def extraer_numero_factura(texto):
    patrones = [
        r'[Nn]ro\.?:?\s*(\d{4,5}-\d{5,10})',
        r'[Cc]omprob\.?\s*[Nn]º?:?\s*(\d{4}-\d{5,10})',
        r'FACTURA\s*[:\s]*(\d{4}-\d{5,10})',
        r'(\d{4}-\d{6,10})',
    ]
    for pat in patrones:
        m = re.search(pat, texto)
        if m:
            return m.group(1).strip()
    return ""


def extraer_cuit_emisor(texto):
    for m in re.finditer(r'CUIT[:\s#°Nº]*(\d{2}[-\s]?\d{8}[-\s]?\d)', texto):
        cuit = m.group(1).strip()
        if '50008962' not in cuit:
            return cuit
    return ""


def extraer_datos(pdf_bytes, nombre_archivo):
    resultado = {
        "archivo": nombre_archivo,
        "emisor": "",
        "fecha_emision": "",
        "importe_total": "",
        "numero_factura": "",
        "cuit_emisor": "",
        "error": "",
    }
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            texto = "\n".join(p.extract_text() or "" for p in pdf.pages[:2])
        if not texto.strip():
            resultado["error"] = "PDF escaneado — sin texto extraíble"
            return resultado
        resultado["emisor"]         = extraer_emisor(texto)
        resultado["fecha_emision"]  = extraer_fecha(texto)
        resultado["importe_total"]  = extraer_importe(texto)
        resultado["numero_factura"] = extraer_numero_factura(texto)
        resultado["cuit_emisor"]    = extraer_cuit_emisor(texto)
    except Exception as e:
        resultado["error"] = str(e)
    return resultado


# ══════════════════════════════════════════════════════════════════════
#  EXCEL
# ══════════════════════════════════════════════════════════════════════

COLUMNAS = {
    "archivo":        "Archivo",
    "emisor":         "Empresa / Emisor",
    "fecha_emision":  "Fecha de Emisión",
    "importe_total":  "Importe Total",
    "numero_factura": "N° Comprobante",
    "cuit_emisor":    "CUIT Emisor",
    "error":          "Observaciones",
}


def generar_excel_bytes(registros):
    filas = [{COLUMNAS[k]: r.get(k, "") for k in COLUMNAS} for r in registros]
    buf = io.BytesIO()
    pd.DataFrame(filas).to_excel(buf, index=False, sheet_name="Facturas")
    buf.seek(0)

    wb = load_workbook(buf)
    ws = wb.active

    hf  = Font(name="Calibri", bold=True, color="FFFFFF", size=10)
    nf  = Font(name="Calibri", size=10, color="1A1A2E")
    bf  = Font(name="Calibri", size=10, bold=True, color="1B3A6B")
    brd = Border(
        left=Side(style="thin", color="D0D7E2"),
        right=Side(style="thin", color="D0D7E2"),
        top=Side(style="thin", color="D0D7E2"),
        bottom=Side(style="thin", color="D0D7E2"),
    )

    for cell in ws[1]:
        cell.font      = hf
        cell.fill      = PatternFill("solid", fgColor="1B3A6B")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border    = brd

    col_obs = list(COLUMNAS.keys()).index("error") + 1
    col_imp = list(COLUMNAS.keys()).index("importe_total") + 1

    for ri, row in enumerate(ws.iter_rows(min_row=2), 2):
        tiene_error = bool(ws.cell(ri, col_obs).value)
        bg = "FFDAD6" if tiene_error else ("F0F4FA" if ri % 2 == 0 else "FFFFFF")
        for cell in row:
            cell.font      = nf
            cell.fill      = PatternFill("solid", fgColor=bg)
            cell.alignment = Alignment(vertical="center")
            cell.border    = brd
        ws.cell(ri, col_imp).font = bf

    anchos = {
        "Archivo": 30, "Empresa / Emisor": 35, "Fecha de Emisión": 16,
        "Importe Total": 20, "N° Comprobante": 22, "CUIT Emisor": 18,
        "Observaciones": 30,
    }
    for i, name in enumerate(COLUMNAS.values(), 1):
        ws.column_dimensions[get_column_letter(i)].width = anchos.get(name, 18)

    ws.row_dimensions[1].height = 30
    ws.freeze_panes = "A2"

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out.read()


# ══════════════════════════════════════════════════════════════════════
#  UI — DISEÑO MINIMALISTA LA NACIÓN
# ══════════════════════════════════════════════════════════════════════

st.set_page_config(
    page_title="Extractor Contable · La Nación",
    page_icon="📋",
    layout="centered",
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Merriweather:wght@400;700&family=Inter:wght@300;400;500;600&display=swap');

html, body, [class*="css"] {
    font-family: 'Inter', sans-serif;
}

.ln-wrap {
    max-width: 680px;
    margin: 0 auto;
    padding: 36px 0 60px;
    text-align: center;
}

.ln-title {
    font-family: 'Merriweather', Georgia, serif;
    font-size: 1.25rem;
    font-weight: 700;
    color: #1B3A6B;
    line-height: 1.5;
    margin: 0 0 4px;
}

.ln-subtitle {
    font-size: 0.82rem;
    color: #8A97B0;
    font-weight: 300;
    margin: 0 0 20px;
    letter-spacing: 0.03em;
}

.ln-logo-wrap {
    display: flex;
    justify-content: center;
    align-items: center;
    padding: 12px 0 24px;
}
.ln-logo-wrap img {
    max-height: 48px;
    max-width: 220px;
    object-fit: contain;
}

.ln-rule {
    border: none;
    border-top: 1px solid #E2E8F0;
    margin: 0 0 28px;
}

.ln-upload-lbl {
    font-size: 0.78rem;
    font-weight: 600;
    color: #4A5568;
    text-transform: uppercase;
    letter-spacing: 0.08em;
    text-align: left;
    margin-bottom: 6px;
}

.ln-stat {
    background: white;
    border: 1px solid #E2E8F0;
    border-radius: 8px;
    padding: 14px 8px;
    text-align: center;
}
.ln-stat-n { font-size: 1.9rem; font-weight: 700; color: #1B3A6B; line-height: 1; }
.ln-stat-l { font-size: 0.68rem; color: #A0AABF; margin-top: 4px; text-transform: uppercase; letter-spacing: 0.06em; }

div[data-testid="stFileUploader"] > label { display: none; }

div[data-testid="stButton"] > button {
    background-color: #1B3A6B !important;
    color: white !important;
    border: none !important;
    border-radius: 6px !important;
    font-family: 'Inter', sans-serif !important;
    font-weight: 500 !important;
    font-size: 0.9rem !important;
    padding: 10px 24px !important;
    letter-spacing: 0.04em !important;
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
    font-family: 'Inter', sans-serif !important;
    font-weight: 500 !important;
}

.ln-footer {
    text-align: center;
    font-size: 0.72rem;
    color: #B0BAD0;
    margin-top: 40px;
    letter-spacing: 0.03em;
}
</style>
""", unsafe_allow_html=True)

# ── Título ───────────────────────────────────────────────────────────
st.markdown("""
<div style="text-align:center; padding: 32px 0 0;">
  <p class="ln-title">Extractor de datos de documentación<br>contable de La Nación</p>
  <p class="ln-subtitle">Procesamiento automático de facturas y comprobantes</p>
</div>
""", unsafe_allow_html=True)

# ── Logo ─────────────────────────────────────────────────────────────
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

# ── Upload ───────────────────────────────────────────────────────────
st.markdown('<p class="ln-upload-lbl">📎 &nbsp; Adjunte su documento</p>', unsafe_allow_html=True)

# Ocultar uploader nativo y mostrar texto en español
st.markdown("""
<style>
div[data-testid="stFileUploaderDropzone"] > div > div:first-child {
    font-size: 0 !important;
}
div[data-testid="stFileUploaderDropzone"] > div > div:first-child::after {
    content: "Arrastre su archivo aquí";
    font-size: 0.95rem;
    color: #4A5568;
    font-family: "Inter", sans-serif;
}
div[data-testid="stFileUploaderDropzone"] > div > small {
    font-size: 0 !important;
}
div[data-testid="stFileUploaderDropzone"] > div > small::after {
    content: "Límite 200MB por archivo · PDF";
    font-size: 0.78rem;
    color: #8A97B0;
}
</style>
""", unsafe_allow_html=True)

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

# ── Botón ────────────────────────────────────────────────────────────
procesar = st.button(
    "Procesar documentos",
    disabled=not archivos,
    use_container_width=True,
)

# ── Procesamiento ────────────────────────────────────────────────────
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
    st.markdown("<br>", unsafe_allow_html=True)

    # Estadísticas
    procesadas = sum(1 for _, ok, _, _ in resultados_ui if ok)
    con_error  = total - procesadas

    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown(f'<div class="ln-stat"><div class="ln-stat-n">{total}</div><div class="ln-stat-l">Documentos</div></div>', unsafe_allow_html=True)
    with c2:
        st.markdown(f'<div class="ln-stat"><div class="ln-stat-n" style="color:#2D6A4F">{procesadas}</div><div class="ln-stat-l">Procesados</div></div>', unsafe_allow_html=True)
    with c3:
        col = "#C0392B" if con_error else "#2D6A4F"
        st.markdown(f'<div class="ln-stat"><div class="ln-stat-n" style="color:{col}">{con_error}</div><div class="ln-stat-l">Con advertencias</div></div>', unsafe_allow_html=True)

    # Tabla
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("**Datos extraídos**")

    filas = []
    for _, ok, err, d in resultados_ui:
        filas.append({
            "Archivo":          d["archivo"],
            "Empresa / Emisor": d["emisor"] or "—",
            "Fecha de Emisión": d["fecha_emision"] or "—",
            "Importe Total":    d["importe_total"] or "—",
            "N° Comprobante":   d["numero_factura"] or "—",
            "Observaciones":    err or "OK",
        })

    st.dataframe(pd.DataFrame(filas), use_container_width=True, hide_index=True)

    # Descarga
    st.markdown("<br>", unsafe_allow_html=True)
    excel_bytes  = generar_excel_bytes(registros)
    nombre_excel = f"LaNacion_facturas_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"

    st.download_button(
        label="⬇  Descargar Excel",
        data=excel_bytes,
        file_name=nombre_excel,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

# ── Footer ───────────────────────────────────────────────────────────
st.markdown("""
<p class="ln-footer">La Nación &nbsp;·&nbsp; Documentación Contable &nbsp;·&nbsp; Uso interno</p>
""", unsafe_allow_html=True)
