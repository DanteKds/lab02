
from pathlib import Path
import re
import pdfplumber
import pandas as pd
import unicodedata
from datetime import datetime

# -----------------------------------------------------
# Esquema normalizado de salida
# -----------------------------------------------------
COLUMNAS = [
    "archivo_pdf",        # nombre del archivo
    "empresa",            # Metrogas / Enel / Aguas Andinas
    "nro_documento",      # número de boleta o factura
    "total_a_pagar",      # monto total
    "id_cliente",         # número o código de cliente
    "fecha_emision",      # fecha de emisión
    "fecha_vencimiento",  # fecha de vencimiento
    "consumo_periodo",    # consumo del período (m3, kWh, etc.)
    "estado"              # OK / PARCIAL / FALLA_EXTRACCION / OCR / etc.
]

# -----------------------------------------------------
# Utilidades
# -----------------------------------------------------
def _normalize_filename(s: str) -> str:
    s = s.strip().lower()
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    return s

def _listar_pdfs(directorio: Path):
    return sorted([p for p in Path(directorio).glob("*.pdf")])

def _leer_texto_pdf(path_pdf: Path) -> str:
    texto = ""
    with pdfplumber.open(str(path_pdf)) as pdf:
        for page in pdf.pages:
            try:
                texto += page.extract_text() or ""
            except Exception:
                texto += ""
    return texto

def _limpiar_monto(x: str | None):
    if not x:
        return None
    x = x.replace(".", "").replace(",", ".").strip()
    try:
        return int(float(x))
    except Exception:
        return x

def _estado_por_campos(dic, claves_minimas):
    if all(dic.get(k) for k in claves_minimas):
        return "OK"
    return "PARCIAL"

# -----------------------------------------------------
# Extractores (mapean a COLUMNAS)
# -----------------------------------------------------
def extraer_metrogas(path_pdf: Path) -> dict:
    out = {k: None for k in COLUMNAS}
    out["archivo_pdf"] = path_pdf.name
    out["empresa"] = "Metrogas"
    try:
        texto = _leer_texto_pdf(path_pdf)

        # id_cliente (nro cuenta / cliente)
        m_id = re.search(r"(?:Nro|N[º°]|Número)\s+de\s+cuenta\s*[:\-]?\s*([\d\-kK]{7,12})", texto, re.IGNORECASE)
        if not m_id:
            m_id = re.search(r"n[uú]mero\s+de\s+cuent[ao].*?([\d\-kK]{7,12})", texto, re.IGNORECASE|re.DOTALL)
        out["id_cliente"] = m_id.group(1) if m_id else None

        # nro_documento
        m_ndoc = re.search(r"BOLETA\s+ELECTR[ÓO]NICA\s*N[º°]?\s*(\d+)", texto, re.IGNORECASE)
        out["nro_documento"] = m_ndoc.group(1) if m_ndoc else None

        # fechas
        m_emision = re.search(r"FECHA\s+EMISI[ÓO]N[:\s]*(\d{2}[-/]\w{3}[-/]\d{4})", texto, re.IGNORECASE)
        out["fecha_emision"] = m_emision.group(1) if m_emision else None

        m_venc = re.search(r"VENCIMIENTO\s*(\d{2}[-/]\w{3}[-/]\d{4})", texto, re.IGNORECASE)
        out["fecha_vencimiento"] = m_venc.group(1) if m_venc else None

        # consumo (m3)
        m_cons = re.search(r"CONSUMO\s+TOTAL\s*([\d\.,]+)\s*m3", texto, re.IGNORECASE)
        out["consumo_periodo"] = (
            m_cons.group(1).replace(",", ".") + " m3" if m_cons else None
        )

        # total a pagar
        m_total = re.search(r"Total\s+a\s+pagar.*?\$?\s*([\d\.]{1,3}(?:\.[\d]{3})*)", texto, re.IGNORECASE|re.DOTALL)
        out["total_a_pagar"] = _limpiar_monto(m_total.group(1)) if m_total else None

        out["estado"] = _estado_por_campos(out, ["nro_documento", "total_a_pagar", "id_cliente", "fecha_emision", "fecha_vencimiento"])
    except Exception:
        out["estado"] = "FALLA_EXTRACCION"
    return out


def extraer_enel(path_pdf: Path) -> dict:
    out = {k: None for k in COLUMNAS}
    out["archivo_pdf"] = path_pdf.name
    out["empresa"] = "Enel"
    try:
        texto = _leer_texto_pdf(path_pdf)

        # id_cliente
        m_id = re.search(r"(?:N°|Nº|Número)\s+Cliente\s*[:\-]?\s*(\d{7,12})", texto, re.IGNORECASE)
        out["id_cliente"] = m_id.group(1) if m_id else None

        # nro_documento
        m_ndoc = re.search(r"Boleta\s+Electr[oó]nica\s*N[º°]?\s*(\d+)", texto, re.IGNORECASE)
        out["nro_documento"] = m_ndoc.group(1) if m_ndoc else None

        # fechas
        m_emision = re.search(r"Fecha\s+de\s+Emisi[oó]n\s*[:\-]?\s*(\d{1,2}\s+\w{3}\s+\d{4})", texto, re.IGNORECASE)
        out["fecha_emision"] = m_emision.group(1) if m_emision else None

        m_venc = re.search(r"Fecha\s+de\s+vencimi(?:ento|miento)\s*[:\-]?\s*(\d{1,2}\s+\w{3}\s+\d{4})", texto, re.IGNORECASE)
        out["fecha_vencimiento"] = m_venc.group(1) if m_venc else None

        # consumo (kWh)
        m_cons = re.search(r"Consumo\s+total\s+del\s+mes\s*=?\s*(\d+)\s*(kWh?)", texto, re.IGNORECASE)
        if m_cons:
            out["consumo_periodo"] = f"{m_cons.group(1)} {m_cons.group(2)}"

        # total a pagar
        m_total = re.search(r"Total\s+a\s+pagar.*?\$?\s*([\d\.]{1,3}(?:\.[\d]{3})*)", texto, re.IGNORECASE|re.DOTALL)
        out["total_a_pagar"] = _limpiar_monto(m_total.group(1)) if m_total else None

        out["estado"] = _estado_por_campos(out, ["nro_documento", "total_a_pagar", "id_cliente", "fecha_emision", "fecha_vencimiento"])
    except Exception:
        out["estado"] = "FALLA_EXTRACCION"
    return out


def extraer_aguas_andinas(path_pdf: Path) -> dict:
    out = {k: None for k in COLUMNAS}
    out["archivo_pdf"] = path_pdf.name
    out["empresa"] = "Aguas Andinas"
    try:
        texto = _leer_texto_pdf(path_pdf)

        # id_cliente
        m_id = re.search(r"(?:Nro|N[º°]|Número)\s+de\s+cuenta\s*[:\-]?\s*([\d\-kK]{7,12})", texto, re.IGNORECASE)
        out["id_cliente"] = m_id.group(1) if m_id else None

        # nro_documento
        m_ndoc = re.search(r"BOLETA\s+ELECTR[ÓO]NICA\s*N[º°]?\s*(\d+)", texto, re.IGNORECASE)
        out["nro_documento"] = m_ndoc.group(1) if m_ndoc else None

        # fechas
        m_emision = re.search(r"FECHA\s+EMISI[ÓO]N[:\s]*(\d{2}[-/]\w{3}[-/]\d{4})", texto, re.IGNORECASE)
        out["fecha_emision"] = m_emision.group(1) if m_emision else None

        m_venc = re.search(r"VENCIMIENTO\s*(\d{2}[-/]\w{3}[-/]\d{4})", texto, re.IGNORECASE)
        out["fecha_vencimiento"] = m_venc.group(1) if m_venc else None

        # consumo (m3)
        m_cons = re.search(r"CONSUMO\s+TOTAL\s*([\d\.,]+)\s*m3", texto, re.IGNORECASE)
        out["consumo_periodo"] = (
            m_cons.group(1).replace(",", ".") + " m3" if m_cons else None
        )

        # total a pagar
        m_total = re.search(r"Total\s+a\s+pagar.*?\$?\s*([\d\.]{1,3}(?:\.[\d]{3})*)", texto, re.IGNORECASE|re.DOTALL)
        out["total_a_pagar"] = _limpiar_monto(m_total.group(1)) if m_total else None

        out["estado"] = _estado_por_campos(out, ["nro_documento", "total_a_pagar", "id_cliente", "fecha_emision", "fecha_vencimiento"])
    except Exception:
        out["estado"] = "FALLA_EXTRACCION"
    return out

# -----------------------------------------------------
# Clasificador por nombre y proceso batch
# -----------------------------------------------------
def _tipo_por_nombre(path_pdf: Path) -> str | None:
    n = _normalize_filename(path_pdf.stem)
    if "metrogas" in n or ("metro" in n and "gas" in n):
        return "metrogas"
    if "enel" in n:
        return "enel"
    if "aguas" in n and "andinas" in n:
        return "aguas_andinas"
    return None

def procesar_boletas(carpeta_boletas: Path, carpeta_salida: Path | None = None) -> Path:
    carpeta_boletas = Path(carpeta_boletas)
    carpeta_salida = Path(carpeta_salida) if carpeta_salida else carpeta_boletas

    resultados = []
    for pdf in _listar_pdfs(carpeta_boletas):
        tipo = _tipo_por_nombre(pdf)
        if tipo == "metrogas":
            resultados.append(extraer_metrogas(pdf))
        elif tipo == "enel":
            resultados.append(extraer_enel(pdf))
        elif tipo == "aguas_andinas":
            resultados.append(extraer_aguas_andinas(pdf))
        else:
            resultados.append({
                "archivo_pdf": pdf.name,
                "empresa": None,
                "nro_documento": None,
                "total_a_pagar": None,
                "id_cliente": None,
                "fecha_emision": None,
                "fecha_vencimiento": None,
                "consumo_periodo": None,
                "estado": "PARCIAL"
            })

    if not resultados:
        raise RuntimeError("No se encontraron PDFs en la carpeta.")

    df = pd.DataFrame(resultados)[COLUMNAS]

    carpeta_salida.mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_csv = carpeta_salida / f"boletas_extraidas_{ts}.csv"
    out_xlsx = carpeta_salida / f"boletas_extraidas_{ts}.xlsx"
    df.to_csv(out_csv, index=False, encoding="utf-8")
    try:
        df.to_excel(out_xlsx, index=False)
    except Exception:
        pass
    return out_csv
