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



# ══════════════════════════════════════════════════════════════════════
#  EXTRACCIÓN DE KILOS Y PRECIO/KG — POR PROVEEDOR
# ══════════════════════════════════════════════════════════════════════

def extraer_kilos_precio_aerolineas(texto_completo):
    """Aerolíneas: kilos aforados totales y precio/kg desde páginas de detalle."""
    m = re.search(r'TOTAL KILOS AFORADOS:\s*([\d\.]+)', texto_completo)
    total_kg = float(m.group(1)) if m else None

    guias = re.findall(
        r'044-\d+\s+\d{2}-\d{2}-\d{4}\s+[\d,]+\s+(\d+)\s+\w+\s+\w+\s+\w+\s+\w+-\w+\s+\d+\s+([\d\.]+)',
        texto_completo
    )
    if guias and total_kg and total_kg > 0:
        total_imp = sum(float(g[1].replace('.','').replace(',','.')) for g in guias)
        return total_kg, round(total_imp / total_kg, 2)
    return total_kg, None


def extraer_kilos_precio_handyway_liq(texto_liq):
    """HandyWay liquidación: n° liq, kilos totales, precio/kg predominante."""
    m_num = re.search(r'LIQUIDACION[:\s#]+\s*(\d+)', texto_liq, re.IGNORECASE)
    n_liq = m_num.group(1).strip() if m_num else None

    m_tot = re.search(r'\d+\s+([\d\.]+)\s+[\d\.]+\s+\$\s*([\d\.]+)', texto_liq)
    total_kg  = float(m_tot.group(1)) if m_tot else None
    total_imp = float(m_tot.group(2)) if m_tot else None

    precios = re.findall(r'\$([\d\.]+)/kg', texto_liq)
    precio_kg = None
    if precios:
        from collections import Counter
        precio_kg = float(Counter(precios).most_common(1)[0][0])
    elif total_kg and total_imp and total_kg > 0:
        precio_kg = round(total_imp / total_kg, 2)

    return n_liq, total_kg, precio_kg


def extraer_kilos_precio_cruzdelsur_excel(excel_bytes, n_factura_pdf):
    """Cruz del Sur Excel: kilos facturados y precio/kg, unido por NumeroDeFactura."""
    try:
        import openpyxl as _opx
        wb = _opx.load_workbook(io.BytesIO(excel_bytes))
        ws = wb.active
        headers = [cell.value for cell in ws[1]]
        idx = {h: i for i, h in enumerate(headers) if h}

        for req in ['NumeroDeFactura', 'KilogramosFacturados', 'Flete']:
            if req not in idx:
                return None, None, f"Excel sin columna '{req}'"

        total_kg = 0.0
        total_flete = 0.0
        coincide = False
        n_pdf_clean = re.sub(r'[-\s]', '', n_factura_pdf or '')

        for row in ws.iter_rows(min_row=2, values_only=True):
            n_excel = re.sub(r'[-\s]', '', str(row[idx['NumeroDeFactura']] or ''))
            if n_pdf_clean and n_excel and (n_pdf_clean in n_excel or n_excel in n_pdf_clean):
                coincide = True
                total_kg    += float(row[idx['KilogramosFacturados']] or 0)
                total_flete += float(row[idx['Flete']] or 0)

        if not coincide:
            return None, None, f"N° factura no coincide con Excel"

        precio_kg = round(total_flete / total_kg, 2) if total_kg > 0 else None
        return total_kg, precio_kg, None
    except Exception as e:
        return None, None, str(e)


def es_liquidacion_handyway(texto):
    """Detecta si el PDF es una liquidación de HandyWay (no una factura)."""
    return bool(re.search(r'LIQUIDACION[:\s#]+\s*\d+', texto, re.IGNORECASE)
                and 'Handyway' in texto or 'HANDYWAY' in texto or 'handyway' in texto.lower())


def es_excel_cruzdelsur(excel_bytes):
    """Detecta si el Excel corresponde a Cruz del Sur (tiene NumeroDeFactura)."""
    try:
        import openpyxl as _opx
        wb = _opx.load_workbook(io.BytesIO(excel_bytes))
        headers = [cell.value for cell in wb.active[1]]
        return 'NumeroDeFactura' in headers and 'KilogramosFacturados' in headers
    except Exception:
        return False


def extraer_datos(pdf_bytes, nombre_archivo, archivos_complementarios=None):
    """
    archivos_complementarios: dict {nombre: bytes} de archivos .xlsx o PDF de liquidación
    subidos junto a la factura para enriquecer los datos.
    """
    resultado = {
        "archivo":        nombre_archivo,
        "emisor":         "",
        "fecha_emision":  "",
        "importe_total":  "",
        "numero_factura": "",
        "cuit_emisor":    "",
        "total_kilos":    "",
        "precio_kg":      "",
        "error":          "",
    }
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            texto_p1  = "\n".join(p.extract_text() or "" for p in pdf.pages[:2])
            texto_all = "\n".join(p.extract_text() or "" for p in pdf.pages)

        if not texto_p1.strip():
            resultado["error"] = "PDF escaneado — sin texto extraíble"
            return resultado

        resultado["emisor"]         = extraer_emisor(texto_p1)
        resultado["fecha_emision"]  = extraer_fecha(texto_p1)
        resultado["importe_total"]  = extraer_importe(texto_p1)
        resultado["numero_factura"] = extraer_numero_factura(texto_p1)
        resultado["cuit_emisor"]    = extraer_cuit_emisor(texto_p1)

        emisor_up = resultado["emisor"].upper()

        # ── Aerolíneas: kilos y precio/kg desde páginas de detalle del mismo PDF
        if "AERO" in emisor_up or "30-64140555" in texto_p1:
            kg, p_kg = extraer_kilos_precio_aerolineas(texto_all)
            if kg:
                resultado["total_kilos"] = str(int(kg)) if kg == int(kg) else str(kg)
            if p_kg:
                resultado["precio_kg"] = formatear_monto(p_kg)

        # ── HandyWay: buscar PDF de liquidación en complementarios
        elif "HANDY" in emisor_up or "30711164932" in texto_p1:
            if archivos_complementarios:
                for nombre_comp, bytes_comp in archivos_complementarios.items():
                    if nombre_comp.lower().endswith('.pdf'):
                        try:
                            with pdfplumber.open(io.BytesIO(bytes_comp)) as pdf_c:
                                texto_liq = "\n".join(p.extract_text() or "" for p in pdf_c.pages)
                            if re.search(r'LIQUIDACION[:\s#]+\s*\d+', texto_liq, re.IGNORECASE):
                                # Verificar que la liquidación corresponde a esta factura
                                n_liq_en_factura = re.search(r'Liquidaci[oó]n\s*#?(\d+)', texto_p1)
                                n_liq_en_liq     = re.search(r'LIQUIDACION[:\s#]+\s*(\d+)', texto_liq, re.IGNORECASE)
                                if (n_liq_en_factura and n_liq_en_liq and
                                        n_liq_en_factura.group(1) == n_liq_en_liq.group(1)):
                                    _, kg, p_kg = extraer_kilos_precio_handyway_liq(texto_liq)
                                    if kg:
                                        resultado["total_kilos"] = str(kg)
                                    if p_kg:
                                        resultado["precio_kg"] = formatear_monto(p_kg)
                        except Exception:
                            pass

        # ── Cruz del Sur: buscar Excel en complementarios, unir por N° factura
        elif "CRUZ DEL SUR" in emisor_up or "MASSON" in emisor_up or "30-55656579" in texto_p1:
            if archivos_complementarios:
                for nombre_comp, bytes_comp in archivos_complementarios.items():
                    if nombre_comp.lower().endswith('.xlsx'):
                        kg, p_kg, err = extraer_kilos_precio_cruzdelsur_excel(
                            bytes_comp, resultado["numero_factura"]
                        )
                        if err:
                            resultado["error"] = err
                        else:
                            if kg:
                                resultado["total_kilos"] = str(kg)
                            if p_kg:
                                resultado["precio_kg"] = formatear_monto(p_kg)

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
    "total_kilos":    "Total Kilos",
    "precio_kg":      "Precio por Kg",
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
        "Archivo": 28, "Empresa / Emisor": 32, "Fecha de Emisión": 14,
        "Importe Total": 18, "Total Kilos": 14, "Precio por Kg": 16,
        "N° Comprobante": 20, "CUIT Emisor": 18, "Observaciones": 28,
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

# ── Upload facturas ─────────────────────────────────────────────────
st.markdown('<p class="ln-upload-lbl">📎 &nbsp; 1. Adjunte las facturas PDF</p>', unsafe_allow_html=True)

archivos = st.file_uploader(
    "Facturas PDF",
    type=["pdf"],
    accept_multiple_files=True,
    label_visibility="collapsed",
    help="Seleccione una o varias facturas en PDF.",
)

if archivos:
    n = len(archivos)
    st.caption(f"✔  {n} factura{'s' if n > 1 else ''} cargada{'s' if n > 1 else ''}.")
else:
    st.caption("Formatos admitidos: PDF · Factura electrónica AFIP")

# ── Upload complementarios ───────────────────────────────────────────
st.markdown("<br>", unsafe_allow_html=True)
st.markdown('<p class="ln-upload-lbl">📎 &nbsp; 2. Adjunte archivos complementarios (opcional)</p>', unsafe_allow_html=True)

complementarios = st.file_uploader(
    "Complementarios",
    type=["pdf", "xlsx"],
    accept_multiple_files=True,
    label_visibility="collapsed",
    help="HandyWay: liquidación PDF · Cruz del Sur: Excel de detalle",
)

if complementarios:
    st.caption(f"✔  {len(complementarios)} archivo{'s' if len(complementarios)>1 else ''} complementario{'s' if len(complementarios)>1 else ''} cargado{'s' if len(complementarios)>1 else ''}.")
else:
    st.caption("HandyWay: liquidación PDF · Cruz del Sur: Excel de detalle")

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

    # Preparar dict de complementarios {nombre: bytes}
    dict_comp = {}
    if complementarios:
        for comp in complementarios:
            dict_comp[comp.name] = comp.read()

    for i, archivo in enumerate(archivos):
        prog.progress(i / total, text=f"Procesando {archivo.name}…")
        datos = extraer_datos(archivo.read(), archivo.name, dict_comp)
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
            "Total Kilos":      d["total_kilos"] or "—",
            "Precio por Kg":    d["precio_kg"] or "—",
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
