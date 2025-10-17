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
    "archivo_pdf",
    "empresa",
    "nro_documento",
    "total_a_pagar",
    "id_cliente",
    "fecha_emision",
    "fecha_vencimiento",
    "consumo_periodo",
    "estado"
]

# -----------------------------------------------------
# Utilidades
# -----------------------------------------------------
def _normalizar_nombre(texto: str) -> str:
    texto = texto.strip().lower()
    texto = unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode("ascii")
    return texto

def _listar_pdfs(directorio: Path):
    return sorted([p for p in Path(directorio).glob("*.pdf")])

def _leer_texto_pdf(path_pdf: Path) -> str:
    texto = ""
    with pdfplumber.open(str(path_pdf)) as pdf:
        for pagina in pdf.pages:
            try:
                texto += pagina.extract_text() or ""
            except Exception:
                texto += ""
    return texto

def _limpiar_monto(valor: str | None):
    if not valor:
        return None
    valor = valor.replace(".", "").replace(",", ".").strip()
    try:
        return int(float(valor))
    except Exception:
        return valor

def _estado_por_campos(dic, claves_minimas):
    if all(dic.get(k) for k in claves_minimas):
        return "OK"
    return "PARCIAL"

# -----------------------------------------------------
# Ayudantes (Aguas Andinas)
# -----------------------------------------------------
def _preprocesar_texto(s: str) -> str:
    s = unicodedata.normalize("NFKC", s)
    s = s.replace("\u00A0", " ")
    s = s.replace("m³", "m3")
    s = re.sub(r"\s+", " ", s)
    return s

def _primer_patron(patrones, s: str):
    for p in patrones:
        m = re.search(p, s, re.IGNORECASE)
        if m:
            return m.group(1)
    return None

def _ventana_derecha(patron_label: str, s: str, ancho: int = 180) -> str | None:
    m = re.search(patron_label, s, re.IGNORECASE)
    if not m:
        return None
    inicio = m.end()
    return s[inicio:inicio+ancho]

# -----------------------------------------------------
# Extractores
# -----------------------------------------------------
def extraer_metrogas(path_pdf: Path) -> dict:
    salida = {k: None for k in COLUMNAS}
    salida["archivo_pdf"] = path_pdf.name
    salida["empresa"] = "Metrogas"
    try:
        texto = _leer_texto_pdf(path_pdf)
        m_id = re.search(r"(?:Nro|N[º°]|Número)\s+de\s+cuenta\s*[:\-]?\s*([\d\-kK]{7,12})", texto, re.IGNORECASE)
        if not m_id:
            m_id = re.search(r"n[uú]mero\s+de\s+cuent[ao].*?([\d\-kK]{7,12})", texto, re.IGNORECASE|re.DOTALL)
        salida["id_cliente"] = m_id.group(1) if m_id else None
        m_ndoc = re.search(r"BOLETA\s+ELECTR[ÓO]NICA\s*N[º°]?\s*(\d+)", texto, re.IGNORECASE)
        salida["nro_documento"] = m_ndoc.group(1) if m_ndoc else None
        m_emision = re.search(r"FECHA\s+EMISI[ÓO]N[:\s]*(\d{2}[-/]\w{3}[-/]\d{4})", texto, re.IGNORECASE)
        salida["fecha_emision"] = m_emision.group(1) if m_emision else None
        m_venc = re.search(r"VENCIMIENTO\s*(\d{2}[-/]\w{3}[-/]\d{4})", texto, re.IGNORECASE)
        salida["fecha_vencimiento"] = m_venc.group(1) if m_venc else None
        m_cons = re.search(r"CONSUMO\s+TOTAL\s*([\d\.,]+)\s*m3", texto, re.IGNORECASE)
        salida["consumo_periodo"] = (m_cons.group(1).replace(",", ".") + " m3" if m_cons else None)
        m_total = re.search(r"Total\s+a\s+pagar.*?\$?\s*([\d\.]{1,3}(?:\.[\d]{3})*)", texto, re.IGNORECASE|re.DOTALL)
        salida["total_a_pagar"] = _limpiar_monto(m_total.group(1)) if m_total else None
        salida["estado"] = _estado_por_campos(salida, ["nro_documento", "total_a_pagar", "id_cliente", "fecha_emision", "fecha_vencimiento"])
    except Exception:
        salida["estado"] = "FALLA_EXTRACCION"
    return salida

def extraer_enel(path_pdf: Path) -> dict:
    salida = {k: None for k in COLUMNAS}
    salida["archivo_pdf"] = path_pdf.name
    salida["empresa"] = "Enel"
    try:
        texto = _leer_texto_pdf(path_pdf)
        m_id = re.search(r"(?:N°|Nº|Número)\s+Cliente\s*[:\-]?\s*(\d{7,12})", texto, re.IGNORECASE)
        salida["id_cliente"] = m_id.group(1) if m_id else None
        m_ndoc = re.search(r"Boleta\s+Electr[oó]nica\s*N[º°]?\s*(\d+)", texto, re.IGNORECASE)
        salida["nro_documento"] = m_ndoc.group(1) if m_ndoc else None
        m_emision = re.search(r"Fecha\s+de\s+Emisi[oó]n\s*[:\-]?\s*(\d{1,2}\s+\w{3}\s+\d{4})", texto, re.IGNORECASE)
        salida["fecha_emision"] = m_emision.group(1) if m_emision else None
        m_venc = re.search(r"Fecha\s+de\s+vencimi(?:ento|miento)\s*[:\-]?\s*(\d{1,2}\s+\w{3}\s+\d{4})", texto, re.IGNORECASE)
        salida["fecha_vencimiento"] = m_venc.group(1) if m_venc else None
        m_cons = re.search(r"Consumo\s+total\s+del\s+mes\s*=?\s*(\d+)\s*(kWh?)", texto, re.IGNORECASE)
        if m_cons:
            salida["consumo_periodo"] = f"{m_cons.group(1)} {m_cons.group(2)}"
        m_total = re.search(r"Total\s+a\s+pagar.*?\$?\s*([\d\.]{1,3}(?:\.[\d]{3})*)", texto, re.IGNORECASE|re.DOTALL)
        salida["total_a_pagar"] = _limpiar_monto(m_total.group(1)) if m_total else None
        salida["estado"] = _estado_por_campos(salida, ["nro_documento", "total_a_pagar", "id_cliente", "fecha_emision", "fecha_vencimiento"])
    except Exception:
        salida["estado"] = "FALLA_EXTRACCION"
    return salida

def extraer_aguas_andinas(path_pdf: Path) -> dict:
    salida = {k: None for k in COLUMNAS}
    salida["archivo_pdf"] = path_pdf.name
    salida["empresa"] = "Aguas Andinas"
    try:
        # Texto lineal (para regex normales)
        texto_raw = _leer_texto_pdf(path_pdf)
        texto = _preprocesar_texto(texto_raw)
        # Texto compacto (para tus patrones de palabras pegadas)
        texto_compacto = re.sub(r"\s+", "", texto)

        # --- patrones en cascada (lineal) ---
        patrones_id = [
            r"(?:Nro|N[º°]|Número)\s+de\s+cuenta\s*[:\-]?\s*([\d\-kK]{6,})",
            r"(?:Nro|N[º°]|Número)\s+cliente\s*[:\-]?\s*([\d\-kK]{6,})",
            r"(?:Nro|N[º°]|Número)\s+servicio\s*[:\-]?\s*([\d\-kK]{6,})",
            r"(?:Cuenta\s+Contrato|Contrato)\s*[:\-]?\s*([\d\-kK]{6,})",
            r"Su\s+n[uú]mero\s+de\s+Cuenta\s+es\s*[:\-]?\s*([\d\-kK]{6,})",
            r"Nro\s+de\s+cuenta\s*[:\-]?\s*([\d\-kK]{6,})",
        ]
        patrones_ndoc = [
            r"BOLETA\s+ELECTR[ÓO]NICA\s*N[º°]?\s*(\d+)",
            r"Folio\s*[:\-]?\s*(\d{5,})",
            r"(?:Documento|Doc\.?)\s*(?:N[º°]|N°|#)?\s*[:\-]?\s*(\d{5,})",
            r"\bN[º°]\s*(\d{5,})",
        ]
        patrones_f_emision = [
            r"FECHA\s+EMISI[ÓO]N[:\s]*([0-3]?\d[-/]\w{3}[-/]\d{4})",
            r"FECHA\s+DE\s+EMISI[ÓO]N[:\s]*([0-3]?\d[-/][01]?\d[-/]\d{4})",
            r"EMISI[ÓO]N[:\s]*([0-3]?\d\s+\w+\s+\d{4})",
        ]
        patrones_f_venc = [
            r"VENCIMIENTO[:\s]*([0-3]?\d[-/]\w{3}[-/]\d{4})",
            r"Vencimiento[:\s]*([0-3]?\d[-/]\w{3}[-/]\d{4})",
            r"FECHA\s+DE\s+VENCIM(?:IENTO|MIENTO)\s*[:\-]?\s*([0-3]?\d\s+\w+\s+\d{4})",
        ]
        patrones_total = [
            r"Total\s*a\s*pagar(?:.{0,60})?\$?\s*([\d\.\s]{1,18})",
            r"TOTAL\s*A\s*PAGAR(?:.{0,60})?\$?\s*([\d\.\s]{1,18})",
            r"TOTAL\s*A\s*PAGAR\s*[^\d$]{0,20}\$?\s*([\d\.\s]{1,18})",
        ]
        patrones_consumo = [
            r"CONSUMO\s+TOTAL\s*([\d\.,]+)\s*m3",
            r"CONSUMO\s+DEL\s+PER[IÍ]ODO\s*([\d\.,]+)\s*m3",
            r"CONSUMO\s+FACTURADO\s*([\d\.,]+)\s*m3",
            r"DIFERENCIA\s+DE\s+LECTURAS\s*([\d\.,]+)\s*m3",
        ]

        # id_cliente (lineal)
        if salida["id_cliente"] is None:
            salida["id_cliente"] = _primer_patron(patrones_id, texto)
        if salida["id_cliente"] is None:
            w = _ventana_derecha(r"(Su\s+n[uú]mero\s+de\s+Cuenta\s+es|Nro\s+de\s+cuenta)", texto) or ""
            salida["id_cliente"] = _primer_patron(patrones_id, w)

        # nro_documento (lineal)
        if salida["nro_documento"] is None:
            salida["nro_documento"] = _primer_patron(patrones_ndoc, texto)

        # fechas (lineal)
        if salida["fecha_emision"] is None:
            salida["fecha_emision"] = _primer_patron(patrones_f_emision, texto)
        if salida["fecha_vencimiento"] is None:
            salida["fecha_vencimiento"] = _primer_patron(patrones_f_venc, texto)

        # total_a_pagar (lineal + ventana)
        if salida["total_a_pagar"] is None:
            w = _ventana_derecha(r"TOTAL\s*A\s*PAGAR", texto) or texto
            v = _primer_patron(patrones_total, w)
            if v:
                v = re.sub(r"\s", "", v)
                salida["total_a_pagar"] = _limpiar_monto(v)

        # consumo (lineal)
        if salida["consumo_periodo"] is None:
            c = _primer_patron(patrones_consumo, texto)
            if c:
                c_norm = c.replace(".", "").replace(",", ".")
                salida["consumo_periodo"] = f"{c_norm} m3"

        # --- FALLBACKS sobre texto compacto (tus patrones) ---
        if salida["nro_documento"] is None:
            patrones_nro_doc_compacto = [
                r"BOLETAELECTR[ÓO]NICA\s*N[º°]?\s*([\d]+)",
                r"N[º°]\s*([\d]+)"
            ]
            salida["nro_documento"] = _primer_patron(patrones_nro_doc_compacto, texto_compacto)

        if salida["total_a_pagar"] is None:
            patrones_total_compacto = [
                r"TOTALAPAGAR\s*\$*([\d\.\,]+)",
                r"TOTAL\s*A\s*PAGAR\s*\$?\s*([\d\.\,]+)"
            ]
            v = _primer_patron(patrones_total_compacto, texto_compacto)
            if v:
                v = re.sub(r"\s", "", v)
                salida["total_a_pagar"] = _limpiar_monto(v)

        if salida["id_cliente"] is None:
            patrones_id_compacto = [
                r"SunúmerodeCuentaes:([\d\-kK]+)",
                r"Su\s*númerode\s*Cuenta\s*es:\s*([\d\-kK]+)",
                r"(?:Nro\s*de\s*cuenta|Número\s*de\s*Cuenta)\s*([\d\-kK]+)"
            ]
            salida["id_cliente"] = _primer_patron([patrones_id_compacto[0]], texto_compacto) \                                   or _primer_patron(patrones_id_compacto[1:], texto)

        if salida["fecha_emision"] is None:
            patrones_femi_compacto = [
                r"FECHAEMISI[ÓO]N:(\d{2}-\w{3}-20\d{2})",
                r"Fechadeemisión:\s*(\d{2}\s*\w{3}\s*\d{4})"
            ]
            salida["fecha_emision"] = _primer_patron(patrones_femi_compacto, texto_compacto)

        if salida["fecha_vencimiento"] is None:
            patrones_fven_compacto = [
                r"VENCIMIENTO\s*(\d{2}-\w{3}-20\d{2})",
                r"PAGARHASTA\s*(\d{2}-\w{3}-20\d{2})"
            ]
            salida["fecha_vencimiento"] = _primer_patron(patrones_fven_compacto, texto_compacto)

        if salida["consumo_periodo"] is None:
            patrones_cons_compacto = [
                r"CONSUMOAGUAPOTABLENOPUNTA\s*([\d\.,]+)",
                r"CONSUMO\s*AGUA\s*([\d\.,]+)",
                r"CONSUMOTOTAL\s*\([\s\w\d]+\)\s*=\s*([\d\.,]+)",
                r"CONSUMOTOTAL\s*([\d\.,]+)"
            ]
            c = _primer_patron(patrones_cons_compacto, texto_compacto)
            if c:
                c_norm = c.replace(".", "").replace(",", ".")
                salida["consumo_periodo"] = f"{c_norm} m3"

        # Estado + validaciones
        salida["estado"] = _estado_por_campos(salida, [
            "nro_documento", "total_a_pagar", "id_cliente", "fecha_emision", "fecha_vencimiento"
        ])

        def _es_num_esperado(s):
            return bool(re.fullmatch(r"[0-9\-\–kK]+", s)) if s else False
        if salida["id_cliente"] and len(salida["id_cliente"]) < 6:
            salida["id_cliente"] = None
        if salida["id_cliente"] and not _es_num_esperado(salida["id_cliente"]):
            salida["id_cliente"] = None
        if salida["nro_documento"] and len(re.sub(r"\D", "", salida["nro_documento"])) < 5:
            salida["nro_documento"] = None
        if isinstance(salida["total_a_pagar"], int):
            if not (0 < salida["total_a_pagar"] < 1_000_000_000):
                salida["total_a_pagar"] = None

        salida["estado"] = _estado_por_campos(salida, [
            "nro_documento", "total_a_pagar", "id_cliente", "fecha_emision", "fecha_vencimiento"
        ])

    except Exception:
        salida["estado"] = "FALLA_EXTRACCION"
    return salida

# -----------------------------------------------------
# Clasificador por nombre y proceso por lotes
# -----------------------------------------------------
def _tipo_por_nombre(path_pdf: Path) -> str | None:
    n = _normalizar_nombre(path_pdf.stem)
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
