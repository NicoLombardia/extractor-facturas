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
    return ""


# ══════════════════════════════════════════════════════════════════════
#  EXTRACCIÓN OPTIMIZADA PARA FACTURAS DE LA NACIÓN
# ══════════════════════════════════════════════════════════════════════

def extraer_emisor(texto):
    lineas = [l.strip() for l in texto.split('\n') if l.strip()]

    EXCLUIR = {'LA NACION', 'NACION', 'ORIGINAL', 'DUPLICADO', 'FACTURA',
               'DOMICILIO', 'RAZON', 'RAZÓN', 'RESPONSABLE', 'INSCRIPTO',
               'LIBERTADOR', 'VICENTE LOPEZ', 'BUENOS AIRES', 'COD.', 'AEP',
               'IVA', 'INGRESOS', 'INICIO', 'COMPROBANTE', 'PERÍODO',
               'APELLIDO', 'CANTIDAD', 'PRODUCTO'}

    def es_nombre_empresa(linea):
        l = linea.strip()
        if len(l) < 4:
            return False
        if re.match(r'^[\d\s\.\-\/]+$', l):
            return False
        for exc in EXCLUIR:
            if exc in l.upper():
                return False
        return True

    # ── Caso HandyWay: "Razón social:\nCARGO SA HANDY WAY CARGO SA A Nro.:"
    # En línea 2 aparece "Razón social: FACTURA A" y en línea 3 "CARGO SA HANDY WAY CARGO SA A Nro.:"
    m = re.search(r'[Rr]az[oó]n\s+social:\s*FACTURA\s+A\n([^\n]+)', texto)
    if m:
        linea = m.group(1)
        # Extraer "HANDY WAY CARGO SA" que está entre "CARGO SA " y " A Nro"
        m2 = re.search(r'CARGO SA (.+?) A Nro', linea)
        if m2:
            return m2.group(1).strip()
        # Alternativa: extraer texto antes de " A\b" o "COD"
        limpio = re.sub(r'\s+A\s+Nro.*$', '', linea).strip()
        limpio = re.sub(r'^CARGO SA\s+', '', limpio).strip()
        if len(limpio) > 4:
            return limpio

    # ── Caso Aerolíneas / Cruz del Sur: nombre en primeras líneas cerca del CUIT
    EXCLUIR_DIR = [r'\d{4,}', r'CP \(', r'T/F', r'TEL', r'INFO@',
                   r'WWW\.', r'TAPIALES', r'MERCADO', r'RICCHERI',
                   r'PROVINCIA', r'JORGE NEWBERY', r'AEROPARQUE',
                   r'\(C\d{4}', r'AV\.', r'AU\.']
    
    def es_dir(linea):
        return any(re.search(p, linea, re.IGNORECASE) for p in EXCLUIR_DIR)
    
    for i, linea in enumerate(lineas[:15]):
        if re.search(r'C\.?U\.?I\.?T\.?[:\s#N°]', linea, re.IGNORECASE):
            for j in range(i - 1, max(i - 8, -1), -1):
                candidato = lineas[j].strip()
                if es_nombre_empresa(candidato) and not es_dir(candidato):
                    limpio = re.sub(r'\s+(Av\.|Au\.|Ing\.|AEP\b|Aero|CP\s*\(|info@).*$', '', candidato, flags=re.IGNORECASE).strip()
                    if len(limpio) > 4 and es_nombre_empresa(limpio) and not es_dir(limpio):
                        return limpio
                    elif len(candidato) > 4 and not es_dir(candidato):
                        return candidato
            break

    # ── Estrategia final: recorrer todas las primeras 12 líneas
    # y elegir la primera que parezca nombre de empresa (sin dirección ni números)
    keywords_empresa = ['S.A.', 'S.A', 'S.R.L.', 'CARGO', 'TRANSPORTES',
                        'AEROLÍNEAS', 'AEROLINEAS', 'VICTOR', 'MASSON',
                        'HANDYWAY', 'HANDY']
    keywords_dir = ['AV.', 'AU. ', 'CALLE ', 'CP (', 'INFO@', 'WWW.',
                    'T/F', 'TEL.', 'TAPIALES', 'MERCADO CENTRAL',
                    'MUÑECAS', 'RICCHERI', 'OBLIGADO', 'AEROPARQUE',
                    'JORGE NEWBERY', 'PROVINCIA DE']

    for linea in lineas[:12]:
        tiene_empresa = any(kw in linea.upper() for kw in keywords_empresa)
        tiene_dir = any(kw in linea.upper() for kw in keywords_dir)
        if tiene_empresa and not tiene_dir and es_nombre_empresa(linea):
            limpio = re.sub(r'\s+(Nro\.|COD\.|AEP\b|Cód\.).*$', '', linea).strip()
            return limpio if len(limpio) > 4 else linea

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
    # Matchea "CUIT", "C.U.I.T." y "C.U.I.T. / D.N.I." — excluye CUIT de La Nación
    patron = r'C\.?U\.?I\.?T\.?[\s:\/DNI\.#°Nº]*\s*(\d{2}[-\s]?\d{8}[-\s]?\d)'
    for m in re.finditer(patron, texto, re.IGNORECASE):
        cuit = m.group(1).strip()
        if '50008962' not in cuit and '5000896' not in cuit:
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
    padding: 12px 0 40px;
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
    margin: 8px 0 16px;
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
<div style="text-align:center; padding: 16px 0 0;">
  <p class="ln-title">Extractor de datos de documentación<br>contable de La Nación</p>
  <p class="ln-subtitle">Procesamiento automático de facturas y comprobantes</p>
</div>
""", unsafe_allow_html=True)

# Logo eliminado — espacio compacto

st.markdown('<hr class="ln-rule">', unsafe_allow_html=True)

# ── Upload ───────────────────────────────────────────────────────────
st.markdown('<p class="ln-upload-lbl">📎 &nbsp; Adjunte su documento</p>', unsafe_allow_html=True)

# Traducir uploader al español con CSS
st.markdown("""
<style>
/* Ocultar texto original en inglés */
div[data-testid="stFileUploaderDropzone"] span[data-testid="stMarkdownContainer"] p,
div[data-testid="stFileUploaderDropzone"] > div > span,
div[data-testid="stFileUploaderDropzone"] > div > div > span,
div[data-testid="stFileUploadDropzone"] span {
    font-size: 0 !important;
    color: transparent !important;
}
div[data-testid="stFileUploaderDropzone"] > div > div:first-child::before {
    content: "Arrastre su archivo aquí";
    font-size: 0.95rem;
    color: #4A5568;
    display: block;
    margin-bottom: 4px;
}
div[data-testid="stFileUploaderDropzone"] > div > div:last-child::before {
    content: "Límite 200MB · PDF";
    font-size: 0.78rem;
    color: #8A97B0;
    display: block;
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
