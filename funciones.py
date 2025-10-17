
from pathlib import Path
import re
import os
import pdfplumber
import pandas as pd
import unicodedata
from datetime import datetime

# -------------------------------
# Utilidades
# -------------------------------
def normalizar_nombre(s: str) -> str:
    s = s.strip().lower()
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    return s

def obtener_pdfs_en_directorio(directorio: Path):
    return sorted([p for p in Path(directorio).glob("*.pdf")])

def _leer_texto_pdf(path_pdf: Path) -> str:
    texto = ""
    with pdfplumber.open(str(path_pdf)) as pdf:
        for page in pdf.pages:
            try:
                texto += page.extract_text() or ""
            except Exception as e:
                # Continuamos aunque una página falle
                texto += ""
    return texto

def _limpiar_montos(x: str | None):
    if not x: 
        return None
    # Normaliza ".000" y "1.234.567" a entero
    x = x.replace(".", "").replace(",", ".").strip()
    try:
        # Muchos montos en CLP son enteros
        return int(float(x))
    except:
        return x

# -------------------------------
# Extractores por proveedor
# (basados en la lógica de tus notebooks)
# -------------------------------
def extraer_datos_metrogas(path_pdf: Path) -> dict:
    texto = _leer_texto_pdf(path_pdf)

    out = {
        "proveedor": "Metrogas",
        "archivo": path_pdf.name,
        "rut_empresa": None,
        "numero_cuenta": None,
        "numero_boleta": None,
        "fecha_emision": None,
        "periodo_facturado": None,
        "consumo_total_m3": None,
        "iva": None,
        "total_a_pagar": None,
        "vencimiento": None,
        "estado": "OK",
    }

    try:
        # patrones robustos (aproximación fiel a los originales)
        rut = re.search(r"R\.U\.T\.[:\s]*([\d\.]+-[\dkK])", texto, re.IGNORECASE)
        num_cta = re.search(r"(?:Nro|N[º°]|Número)\s+de\s+cuenta\s*[:\-]?\s*([\d\-kK]{7,12})", texto, re.IGNORECASE)
        if not num_cta:
            num_cta = re.search(r"n[uú]mero\s+de\s+cuent[ao].*?([\d\-kK]{7,12})", texto, re.IGNORECASE|re.DOTALL)
        n_boleta = re.search(r"BOLETA\s+ELECTR[ÓO]NICA\s*N[º°]?\s*(\d+)", texto, re.IGNORECASE)
        f_emision = re.search(r"FECHA\s+EMISI[ÓO]N[:\s]*(\d{2}[-/]\w{3}[-/]\d{4})", texto, re.IGNORECASE)
        per_fact = re.search(r"Considera\s+movimientos\s+hasta\s*(\d{2}[-/]\d{2}[-/]\d{4})", texto, re.IGNORECASE)
        consumo = re.search(r"CONSUMO\s+TOTAL\s*([\d\.,]+)\s*m3", texto, re.IGNORECASE)
        iva = re.search(r"\bIVA\b.*?\$?\s*([\d\.]{1,3}(?:\.[\d]{3})*)", texto, re.IGNORECASE|re.DOTALL)
        total = re.search(r"Total\s+a\s+pagar.*?\$?\s*([\d\.]{1,3}(?:\.[\d]{3})*)", texto, re.IGNORECASE|re.DOTALL)
        venc = re.search(r"VENCIMIENTO\s*(\d{2}[-/]\w{3}[-/]\d{4})", texto, re.IGNORECASE)

        if rut: out["rut_empresa"] = rut.group(1)
        if num_cta: out["numero_cuenta"] = num_cta.group(1)
        if n_boleta: out["numero_boleta"] = n_boleta.group(1)
        if f_emision: out["fecha_emision"] = f_emision.group(1)
        if per_fact: out["periodo_facturado"] = per_fact.group(1)
        if consumo: out["consumo_total_m3"] = consumo.group(1).replace(",", ".")
        if iva: out["iva"] = _limpiar_montos(iva.group(1))
        if total: out["total_a_pagar"] = _limpiar_montos(total.group(1))
        if venc: out["vencimiento"] = venc.group(1)

        claves = ["numero_cuenta","rut_empresa","numero_boleta","fecha_emision","total_a_pagar","vencimiento"]
        if not all(out.get(k) for k in claves):
            out["estado"] = "ERROR: Datos incompletos"
    except Exception as e:
        out["estado"] = f"ERROR: {e}"

    return out


def extraer_datos_enel(path_pdf: Path) -> dict:
    texto = _leer_texto_pdf(path_pdf)

    out = {
        "proveedor": "Enel",
        "archivo": path_pdf.name,
        "rut_empresa": None,
        "numero_cuenta": None,
        "numero_boleta": None,
        "fecha_emision": None,
        "monto_periodo": None,
        "consumo_total": None,
        "total_a_pagar": None,
        "fecha_vencimiento": None,
        "estado": "OK",
    }

    try:
        rut = re.search(r"R\.U\.T\.\s*[:\-]?\s*([\d\.]+-[\dkK])", texto, re.IGNORECASE)
        num_cte = re.search(r"(?:N°|Nº|Número)\s+Cliente\s*[:\-]?\s*(\d{7,12})", texto, re.IGNORECASE)
        n_boleta = re.search(r"Boleta\s+Electr[oó]nica\s*N[º°]?\s*(\d+)", texto, re.IGNORECASE)
        f_emision = re.search(r"Fecha\s+de\s+Emisi[oó]n\s*[:\-]?\s*(\d{1,2}\s+\w{3}\s+\d{4})", texto, re.IGNORECASE)
        monto_periodo = re.search(r"Monto\s+del\s+per[ií]odo\s*[:\-]?\s*\$?\s*([\d\.]{1,3}(?:\.[\d]{3})*)", texto, re.IGNORECASE)
        consumo = re.search(r"Consumo\s+total\s+del\s+mes\s*=?\s*(\d+\s*kWh?)", texto, re.IGNORECASE)
        total = re.search(r"Total\s+a\s+pagar.*?\$?\s*([\d\.]{1,3}(?:\.[\d]{3})*)", texto, re.IGNORECASE|re.DOTALL)
        f_venc = re.search(r"Fecha\s+de\s+vencimi(?:ento|miento)\s*[:\-]?\s*(\d{1,2}\s+\w{3}\s+\d{4})", texto, re.IGNORECASE)

        if rut: out["rut_empresa"] = rut.group(1)
        if num_cte: out["numero_cuenta"] = num_cte.group(1)
        if n_boleta: out["numero_boleta"] = n_boleta.group(1)
        if f_emision: out["fecha_emision"] = f_emision.group(1)
        if monto_periodo: out["monto_periodo"] = _limpiar_montos(monto_periodo.group(1))
        if consumo: out["consumo_total"] = consumo.group(1)
        if total: out["total_a_pagar"] = _limpiar_montos(total.group(1))
        if f_venc: out["fecha_vencimiento"] = f_venc.group(1)

        claves = ["numero_cuenta","rut_empresa","numero_boleta","fecha_emision","total_a_pagar","fecha_vencimiento"]
        if not all(out.get(k) for k in claves):
            out["estado"] = "ERROR: Datos incompletos"
    except Exception as e:
        out["estado"] = f"ERROR: {e}"

    return out


def extraer_datos_aguas_andinas(path_pdf: Path) -> dict:
    texto = _leer_texto_pdf(path_pdf)

    out = {
        "proveedor": "Aguas Andinas",
        "archivo": path_pdf.name,
        "rut_empresa": None,
        "numero_cuenta": None,
        "numero_boleta": None,
        "fecha_emision": None,
        "periodo_facturado": None,
        "consumo_total_m3": None,
        "iva": None,
        "total_a_pagar": None,
        "vencimiento": None,
        "estado": "OK",
    }

    try:
        rut = re.search(r"R\.U\.T\.[:\s]*([\d\.]+-[\dkK])", texto, re.IGNORECASE)
        num_cta = re.search(r"(?:Nro|N[º°]|Número)\s+de\s+cuenta\s*[:\-]?\s*([\d\-kK]{7,12})", texto, re.IGNORECASE)
        n_boleta = re.search(r"BOLETA\s+ELECTR[ÓO]NICA\s*N[º°]?\s*(\d+)", texto, re.IGNORECASE)
        f_emision = re.search(r"FECHA\s+EMISI[ÓO]N[:\s]*(\d{2}[-/]\w{3}[-/]\d{4})", texto, re.IGNORECASE)
        per_fact = re.search(r"Considera\s+movimientos\s+hasta\s*(\d{2}[-/]\d{2}[-/]\d{4})", texto, re.IGNORECASE)
        consumo = re.search(r"CONSUMO\s+TOTAL\s*([\d\.,]+)\s*m3", texto, re.IGNORECASE)
        iva = re.search(r"\bIVA\b.*?\$?\s*([\d\.]{1,3}(?:\.[\d]{3})*)", texto, re.IGNORECASE|re.DOTALL)
        total = re.search(r"Total\s+a\s+pagar.*?\$?\s*([\d\.]{1,3}(?:\.[\d]{3})*)", texto, re.IGNORECASE|re.DOTALL)
        venc = re.search(r"VENCIMIENTO\s*(\d{2}[-/]\w{3}[-/]\d{4})", texto, re.IGNORECASE)

        if rut: out["rut_empresa"] = rut.group(1)
        if num_cta: out["numero_cuenta"] = num_cta.group(1)
        if n_boleta: out["numero_boleta"] = n_boleta.group(1)
        if f_emision: out["fecha_emision"] = f_emision.group(1)
        if per_fact: out["periodo_facturado"] = per_fact.group(1)
        if consumo: out["consumo_total_m3"] = consumo.group(1).replace(",", ".")
        if iva: out["iva"] = _limpiar_montos(iva.group(1))
        if total: out["total_a_pagar"] = _limpiar_montos(total.group(1))
        if venc: out["vencimiento"] = venc.group(1)

        claves = ["numero_cuenta","rut_empresa","numero_boleta","fecha_emision","total_a_pagar","vencimiento"]
        if not all(out.get(k) for k in claves):
            out["estado"] = "ERROR: Datos incompletos"
    except Exception as e:
        out["estado"] = f"ERROR: {e}"

    return out

# -------------------------------
# Clasificador por nombre de archivo y proceso batch
# -------------------------------
def _tipo_por_nombre(path_pdf: Path) -> str | None:
    nombre = normalizar_nombre(path_pdf.stem)
    if "metrogas" in nombre or ("metro" in nombre and "gas" in nombre):
        return "metrogas"
    if "enel" in nombre:
        return "enel"
    if ("aguas" in nombre) and ("andinas" in nombre):
        return "aguas_andinas"
    return None

def procesar_boletas(carpeta_boletas: Path, carpeta_salida: Path | None = None) -> Path:
    carpeta_boletas = Path(carpeta_boletas)
    carpeta_salida = Path(carpeta_salida) if carpeta_salida else carpeta_boletas

    resultados = []
    for pdf in obtener_pdfs_en_directorio(carpeta_boletas):
        tipo = _tipo_por_nombre(pdf)
        if tipo == "metrogas":
            resultados.append(extraer_datos_metrogas(pdf))
        elif tipo == "enel":
            resultados.append(extraer_datos_enel(pdf))
        elif tipo == "aguas_andinas":
            resultados.append(extraer_datos_aguas_andinas(pdf))
        else:
            resultados.append({
                "proveedor": "DESCONOCIDO",
                "archivo": pdf.name,
                "estado": "Tipo desconocido por nombre de archivo",
            })

    if not resultados:
        raise RuntimeError("No se encontraron PDFs en la carpeta.")

    df = pd.DataFrame(resultados)
    carpeta_salida.mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_csv = carpeta_salida / f"boletas_extraidas_{ts}.csv"
    out_xlsx = carpeta_salida / f"boletas_extraidas_{ts}.xlsx"
    df.to_csv(out_csv, index=False, encoding="utf-8")
    try:
        df.to_excel(out_xlsx, index=False)
    except Exception:
        # openpyxl no siempre está disponible
        pass
    return out_csv
