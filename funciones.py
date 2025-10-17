
from pathlib import Path
import re
import os
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
    "estado",
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

# --- recolector de montos en una cadena (enteros) ---
def _montos_en_texto(s: str):
    """
    Devuelve lista de montos encontrados como enteros a partir de patrones tipo
    $ 1.234.567, 1.234.567 o 1234567[,..]. Ignora símbolos y espacios.
    """
    candidatos = []
    patron = r"\$?\s*([0-9]{1,3}(?:[.\s][0-9]{3})+|[0-9]+)(?:,[0-9]{1,2})?"
    for m in re.finditer(patron, s):
        bruto = m.group(1)
        limpio = (
            bruto.replace(".", "")
                 .replace(" ", "")
                 .replace(" ", "")
        )
        try:
            val = int(float(limpio.replace(",", ".")))
            candidatos.append(val)
        except Exception:
            pass
    return candidatos

# -----------------------------------------------------
# Ayudantes (Aguas Andinas)
# -----------------------------------------------------
def _preprocesar_texto(s: str) -> str:
    s = unicodedata.normalize("NFKC", s)
    s = s.replace(" ", " ")
    s = s.replace("m³", "m3")  # m³ -> m3
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
    return s[inicio : inicio + ancho]
# -----------------------------------------------------
# ayudantes enel
# -----------------------------------------------------
def _ventana_alrededor(patron_label: str, s: str, izq: int = 150, der: int = 260) -> str | None:
    """
    Devuelve una subcadena que toma 'izq' caracteres a la izquierda y 'der' a la derecha
    del patrón encontrado. Útil cuando el monto puede estar antes o después del label.
    """
    m = re.search(patron_label, s, re.IGNORECASE)
    if not m:
        return None
    ini = max(0, m.start() - izq)
    fin = m.end() + der
    return s[ini:fin]

# -----------------------------------------------------
# ayudantes enelmetrogas
# -----------------------------------------------------


def _ventana_alrededor(patron_label: str, s: str, izq: int = 220, der: int = 360) -> str | None:
    """Subcadena que toma 'izq' chars a la izquierda y 'der' a la derecha del label."""
    m = re.search(patron_label, s, re.IGNORECASE)
    if not m:
        return None
    ini = max(0, m.start() - izq)
    fin = m.end() + der
    return s[ini:fin]

    m = re.search(patron_label, s, re.IGNORECASE)
    if not m:
        return None
    ini = max(0, m.start() - izq)
    fin = m.end() + der
    return s[ini:fin]

def _candidatos_total_por_label(texto: str) -> list[int]:
    """
    Busca 'TOTAL A PAGAR' y variantes y recoge montos alrededor (antes y después).
    Devuelve lista de enteros candidatos.
    """
    etiquetas = r"(?:TOTAL\s*A\s*PAGAR|Total\s*a\s*pagar|TOTAL\s+BOLETA|Total\s+boleta|Monto\s+total|Total\s+Cuenta)"
    candidatos = []
    for m in re.finditer(etiquetas, texto, re.IGNORECASE):
        w = _ventana_alrededor(etiquetas, texto, izq=220, der=360)
        if not w:
            continue
        candidatos += _montos_en_texto(w)
    return candidatos

def _candidatos_total_por_words_metrogas(path_pdf: Path) -> list[int]:
    """
    Usa pdfplumber para buscar la palabra 'Total'/'TOTAL'/'TOTAL A PAGAR' y captura montos
    en la misma línea y en la línea siguiente (izq/dcha).
    """
    etiquetas = re.compile(r"(?:TOTAL|Total)(?:\s*A\s*PAGAR|)|Total\s+boleta|Monto\s+total", re.IGNORECASE)
    candidatos = []
    try:
        with pdfplumber.open(str(path_pdf)) as pdf:
            # Priorizamos página 1
            paginas = [pdf.pages[0]] + list(pdf.pages[1:])
            for pagina in paginas:
                try:
                    words = pagina.extract_words(use_text_flow=True, keep_blank_chars=False) or []
                except Exception:
                    words = []
                # Agrupar por renglones aproximando Y
                lineas = {}
                for w in words:
                    y = round(w.get("top", 0), 1)
                    lineas.setdefault(y, []).append(w)
                ys = sorted(lineas.keys())
                for idx, y in enumerate(ys):
                    ws = sorted(lineas[y], key=lambda z: z.get("x0", 0))
                    # Texto de la línea
                    linea_txt = " ".join([w.get("text", "") for w in ws])
                    if etiquetas.search(linea_txt):
                        # Buscar montos a derecha e izquierda en la MISMA línea
                        texto_linea = " ".join([w.get("text", "") for w in ws])
                        candidatos += _montos_en_texto(texto_linea)
                        # Y en la línea inmediatamente siguiente (algunos diseños imprimen el monto debajo)
                        if idx + 1 < len(ys):
                            ws_next = sorted(lineas[ys[idx+1]], key=lambda z: z.get("x0", 0))
                            texto_next = " ".join([w.get("text", "") for w in ws_next])
                            candidatos += _montos_en_texto(texto_next)
    except Exception:
        pass
    return candidatos


    """
    Dado un conjunto de montos candidatos, elige el más probable para 'TOTAL A PAGAR'.
    Heurísticas:
      - preferir >= 10.000 (evita que agarre 5 o 999 por ruido),
      - entre los >= 10.000 devolver el máximo,
      - si no hay >= 10.000, devolver el máximo de todos.
    """
    if not cands:
        return None
    altos = [c for c in cands if c >= 10_000]
    return max(altos) if altos else max(cands)

def _elegir_total_preferente(cands: list[int]) -> int | None:
    """
    Heurística: si hay montos >= 10.000, usa el mayor de ellos (evita 5/999).
    Si no, usa el mayor disponible.
    """
    if not cands:
        return None
    altos = [c for c in cands if c >= 10_000]
    return max(altos) if altos else max(cands)


# -----------------------------------------------------
# Extractores
# -----------------------------------------------------
def extraer_metrogas(path_pdf: Path) -> dict:
    salida = {k: None for k in COLUMNAS}
    salida["archivo_pdf"] = path_pdf.name
    salida["empresa"] = "Metrogas"
    try:
        # Texto base
        texto_raw = _leer_texto_pdf(path_pdf)
        texto = _preprocesar_texto(texto_raw)
        texto_compacto = re.sub(r"\s+", "", texto)

        # -------- Patrones --------
        # id_cliente (en boletas Metrogas aparece como "Número Interno" y también como "Nº Cliente" en otras)
        patrones_id = [
            r"N[uú]mero\s+Interno\s*[:\-]?\s*([0-9]{6,12})",
            r"(?:N[°º]|Nº|Nro)\s*Cliente\s*[:\-]?\s*([0-9]{6,12})",
            r"Cliente\s*(?:N[°º]|Nº|#)?\s*[:\-]?\s*([0-9]{6,12})",
        ]
        patrones_id_compacto = [
            r"NumeroInterno[:\-]?([0-9]{6,12})",
            r"(?:N[°º]|Nº|Nro)Cliente[:\-]?([0-9]{6,12})",
        ]

        # nro_documento / folio
        patrones_ndoc = [
            r"BOLETA\s+ELECTR[ÓO]NICA\s*N[°º]?\s*(\d{5,})",
            r"Folio\s*[:\-]?\s*(\d{5,})",
        ]
        patrones_ndoc_compacto = [
            r"BOLETAELECTR[ÓO]NICA\s*N[°º]?\s*(\d{5,})",
            r"Folio[:\-]?(\d{5,})",
        ]

        # fechas
        patrones_f_emision = [
            r"Fecha\s+de\s+emisi[oó]n\s*[:\-]?\s*([0-3]?\d\s+\w{3}\s+\d{4})",
            r"Fecha\s+de\s+emisi[oó]n\s*[:\-]?\s*([0-3]?\d[-/]\w{3}[-/]\d{4})",
        ]
        patrones_f_venc = [
            r"(?:Fecha\s+de\s+vencimi(?:ento|miento)|Vencimiento)\s*[:\-]?\s*([0-3]?\d\s+\w{3}\s+\d{4})",
            r"(?:Fecha\s+de\s+vencimi(?:ento|miento)|Vencimiento)\s*[:\-]?\s*([0-3]?\d[-/]\w{3}[-/]\d{4})",
        ]

        # -------- TOTAL A PAGAR (robusto para Metrogas) --------
        patron_total = r"(?:VENCIMIENTO\s+)?TOTAL\s*A\s*PAGAR|Total\s*a\s*pagar|TOTAL\s+BOLETA|Total\s+boleta|Monto\s+total"

        # A) monto a la derecha del label (ruta normal)
        if salida["total_a_pagar"] is None:
            w = _ventana_derecha(patron_total, texto, ancho=320) or ""
            v = _primer_patron([
                r"(?:TOTAL|Total).{0,80}\$?\s*([\d\.\s]{1,18})",
                r"\$?\s*([\d\.\s]{1,18})"   # por si la regex anterior no pesca el número
            ], w)
            if v:
                salida["total_a_pagar"] = _limpiar_monto(re.sub(r"\s", "", v))

        # B) monto ANTES del label (ej.: "$ 6.840.780  Total a pagar")
        if salida["total_a_pagar"] is None:
            w_bi = _ventana_alrededor(patron_total, texto, izq=240, der=360) or ""
            v = _primer_patron([
                r"\$?\s*([\d\.\s]{1,18})\s*(?:Total\s*a\s*pagar|TOTAL\s*A\s*PAGAR|TOTAL\s+BOLETA|Total\s+boleta|Monto\s+total)"
            ], w_bi)
            if v:
                salida["total_a_pagar"] = _limpiar_monto(re.sub(r"\s", "", v))

        # C) compacto (todo pegado / sin espacios)
        if salida["total_a_pagar"] is None:
            v = _primer_patron([
                r"TOTALAPAGAR\s*\$*([\d\.\,]+)",
                r"TOTAL\s*A\s*PAGAR\s*\$?\s*([\d\.\,]+)",
                r"TOTALBOLETA\s*\$?\s*([\d\.\,]+)",
                r"\$?\s*([\d\.\,]+)\s*TOTALAPAGAR"
            ], texto_compacto)
            if v:
                salida["total_a_pagar"] = _limpiar_monto(re.sub(r"\s", "", v))

        # D) pool de candidatos alrededor del label (misma línea y línea siguiente) y elegir preferente
        if salida["total_a_pagar"] is None or (isinstance(salida["total_a_pagar"], int) and salida["total_a_pagar"] < 10000):
            candidatos = []

            # 1) Texto plano: ventana alrededor
            w2 = _ventana_alrededor(patron_total, texto, izq=260, der=400) or ""
            candidatos += _montos_en_texto(w2)

            # 2) Palabras por renglón (pdfplumber): misma línea y la siguiente
            try:
                with pdfplumber.open(str(path_pdf)) as pdf:
                    paginas = [pdf.pages[0]] + list(pdf.pages[1:])
                    for pagina in paginas:
                        try:
                            words = pagina.extract_words(use_text_flow=True, keep_blank_chars=False) or []
                        except Exception:
                            words = []
                        # Agrupar por y/top ≈ línea
                        lineas = {}
                        for w in words:
                            y = round(w.get("top", 0), 1)
                            lineas.setdefault(y, []).append(w)
                        ys = sorted(lineas.keys())
                        for idx, y in enumerate(ys):
                            ws = sorted(lineas[y], key=lambda z: z.get("x0", 0))
                            linea = " ".join([w.get("text","") for w in ws])
                            if re.search(patron_total, linea, re.IGNORECASE):
                                candidatos += _montos_en_texto(linea)
                                if idx + 1 < len(ys):
                                    linea_next = " ".join([w.get("text","") for w in sorted(lineas[ys[idx+1]], key=lambda z: z.get("x0",0))])
                                    candidatos += _montos_en_texto(linea_next)
            except Exception:
                pass

            elegido = _elegir_total_preferente(candidatos)
            if elegido and (not isinstance(salida["total_a_pagar"], int) or salida["total_a_pagar"] < 10000):
                salida["total_a_pagar"] = elegido

        # consumo (m3s)
        patrones_consumo = [
            r"Gas\s+consumido\s*\(\s*([\d\.,]+)\s*m3s?\s*\)",
            r"Consumo\s+Total\s*\(\s*[^\)]*\)\s*=\s*([\d\.,]+)\s*(?:m3s?|m3)",
        ]
        patrones_consumo_compacto = [
            r"Gasconsumido\(([\d\.,]+)m3s?\)",
            r"ConsumoTotal\([^\)]*\)=([\d\.,]+)(?:m3s?|m3)",
        ]

        # -------- Extracciones --------
        # id_cliente
        salida["id_cliente"] = _primer_patron(patrones_id, texto)
        if salida["id_cliente"] is None:
            salida["id_cliente"] = _primer_patron(patrones_id_compacto, texto_compacto)

        # nro_documento
        salida["nro_documento"] = _primer_patron(patrones_ndoc, texto)
        if salida["nro_documento"] is None:
            salida["nro_documento"] = _primer_patron(patrones_ndoc_compacto, texto_compacto)

        # fechas
        if salida["fecha_emision"] is None:
            salida["fecha_emision"] = _primer_patron(patrones_f_emision, texto)
        if salida["fecha_vencimiento"] is None:
            salida["fecha_vencimiento"] = _primer_patron(patrones_f_venc, texto)

        # total a pagar: derecha del label
        if salida["total_a_pagar"] is None:
            w = _ventana_derecha(r"(?:VENCIMIENTO\s+)?TOTAL\s*A\s*PAGAR|Total\s*a\s*pagar", texto, ancho=280) or texto
            v = _primer_patron(patrones_total_linea, w)
            if v:
                v = re.sub(r"\s", "", v)
                salida["total_a_pagar"] = _limpiar_monto(v)

        # total a pagar: monto ANTES del label (ej.: "$ 6.840.780  Total a pagar")
        if salida["total_a_pagar"] is None:
            w_bi = _ventana_alrededor(r"(?:VENCIMIENTO\s+)?TOTAL\s*A\s*PAGAR|Total\s*a\s*pagar", texto, izq=200, der=300) or ""
            v = _primer_patron([
                r"\$?\s*([\d\.\s]{1,18})\s*Total\s*a\s*pagar",
                r"\$?\s*([\d\.\s]{1,18})\s*TOTAL\s*A\s*PAGAR",
            ], w_bi)
            if v:
                v = re.sub(r"\s", "", v)
                salida["total_a_pagar"] = _limpiar_monto(v)

        # total a pagar: compacto
        if salida["total_a_pagar"] is None:
            v = _primer_patron(patrones_total_compacto, texto_compacto)
            if v:
                v = re.sub(r"\s", "", v)
                salida["total_a_pagar"] = _limpiar_monto(v)

        # selección robusta: mayor monto alrededor del label si quedó <1000 o None
        if salida["total_a_pagar"] is None or (isinstance(salida["total_a_pagar"], int) and salida["total_a_pagar"] < 1000):
            w2 = _ventana_alrededor(r"(?:VENCIMIENTO\s+)?TOTAL\s*A\s*PAGAR|Total\s*a\s*pagar", texto, izq=240, der=360) or ""
            cand = _montos_en_texto(w2)
            if cand:
                elegido = max(cand)
                if not isinstance(salida["total_a_pagar"], int) or salida["total_a_pagar"] < 1000:
                    salida["total_a_pagar"] = elegido

        # consumo (m3s)
        if salida["consumo_periodo"] is None:
            c = _primer_patron(patrones_consumo, texto)
            if c:
                salida["consumo_periodo"] = f"{c.replace('.', '').replace(',', '.')} m3s"
        if salida["consumo_periodo"] is None:
            c = _primer_patron(patrones_consumo_compacto, texto_compacto)
            if c:
                salida["consumo_periodo"] = f"{c.replace('.', '').replace(',', '.')} m3s"

        # -------- Validaciones + estado --------
        def _es_num_ok(s): return bool(re.fullmatch(r"[0-9\-–kK]+", s)) if s else False

        if salida["id_cliente"] and (len(salida["id_cliente"]) < 6 or not _es_num_ok(salida["id_cliente"])):
            salida["id_cliente"] = None
        if salida["nro_documento"] and len(re.sub(r"\D", "", salida["nro_documento"])) < 5:
            salida["nro_documento"] = None
        if isinstance(salida["total_a_pagar"], int) and not (0 < salida["total_a_pagar"] < 1_000_000_000):
            salida["total_a_pagar"] = None

        salida["estado"] = _estado_por_campos(
            salida, ["nro_documento", "total_a_pagar", "id_cliente", "fecha_emision", "fecha_vencimiento"]
        )
    except Exception:
        salida["estado"] = "FALLA_EXTRACCION"
    return salida


def extraer_enel(path_pdf: Path) -> dict:
    salida = {k: None for k in COLUMNAS}
    salida["archivo_pdf"] = path_pdf.name
    salida["empresa"] = "Enel"
    try:
        # Texto base
        texto_raw = _leer_texto_pdf(path_pdf)
        texto = _preprocesar_texto(texto_raw)
        texto_compacto = re.sub(r"\s+", "", texto)

        # --------- Patrones (todos definidos AQUÍ para evitar 'no definido') ---------
        # ID cliente
        patrones_id = [
            r"(?:N[°º]|Nº|Número)\s*Cliente\s*[:\-]?\s*(\d{6,12})",
            r"Cliente\s*(?:N[°º]|Nº|#)?\s*[:\-]?\s*(\d{6,12})",
            r"(?:ID|Identificador)\s*Cliente\s*[:\-]?\s*(\d{6,12})",
            r"(?:N[°º]|Nº)\s*(\d{6,12})\s*Cliente",
        ]

        # Nro documento / folio
        patrones_ndoc = [
            r"Boleta\s+Electr[oó]nica\s*N[°º]?\s*(\d{5,})",
            r"Folio\s*[:\-]?\s*(\d{5,})",
            r"(?:Documento|Doc\.?)\s*(?:N[°º]|Nº|#)?\s*[:\-]?\s*(\d{5,})",
            r"\bN[°º]\s*(\d{5,})\b",
        ]
        patrones_ndoc_compacto = [
            r"BoletaElectr[oó]nicaN[°º]?\s*(\d{5,})",
            r"Folio[:\-]?\s*(\d{5,})",
            r"N[°º]\s*(\d{5,})",
        ]

        # Fechas
        patrones_f_emision = [
            r"Fecha\s+de\s+Emisi[oó]n\s*[:\-]?\s*([0-3]?\d\s+\w{3}\s+\d{4})",
            r"Fecha\s+de\s+Emisi[oó]n\s*[:\-]?\s*([0-3]?\d[-/]\w{3}[-/]\d{4})",
            r"Fecha\s+Emisi[oó]n\s*[:\-]?\s*([0-3]?\d[-/][01]?\d[-/]\d{4})",
        ]
        patrones_f_venc = [
            r"Fecha\s+de\s+Vencimi(?:ento|miento)\s*[:\-]?\s*([0-3]?\d\s+\w{3}\s+\d{4})",
            r"Fecha\s+de\s+Vencimi(?:ento|miento)\s*[:\-]?\s*([0-3]?\d[-/]\w{3}[-/]\d{4})",
            r"Vencimiento\s*[:\-]?\s*([0-3]?\d[-/][01]?\d[-/]\d{4})",
        ]

        # Total a pagar
        patrones_total_linea = [
            r"Total\s*a\s*pagar(?:.{0,80})?\$?\s*([\d\.\s]{1,18})",
            r"TOTAL\s*A\s*PAGAR(?:.{0,80})?\$?\s*([\d\.\s]{1,18})",
            r"TOTAL\s*A\s*PAGAR\s*[^\d$]{0,40}\$?\s*([\d\.\s]{1,18})",
        ]
        patrones_total_compacto = [
            r"TOTALAPAGAR\s*\$*([\d\.\,]+)",
            r"TOTAL\s*A\s*PAGAR\s*\$?\s*([\d\.\,]+)",
            r"\$?\s*([\d\.\,]+)\s*TOTALAPAGAR",  # monto antes del label, pegado
        ]

        # Consumo
        patrones_consumo = [
            r"Consumo\s+total\s+del\s+mes\s*=?\s*([\d\.,]+)\s*(kW?h?)",
            r"Consumo\s*:?[\s=]*([\d\.,]+)\s*(kW?h?)",
        ]
        patrones_consumo_compacto = [
            r"CONSUMOTOTALDELMES\s*=?\s*([\d\.,]+)(kWh?)",
            r"Consumo\s*total\s*:?=*\s*([\d\.,]+)(kWh?)",
        ]

        # --------- Extracciones ---------
        # id_cliente
        salida["id_cliente"] = _primer_patron(patrones_id, texto)
        if salida["id_cliente"] is None:
            w = _ventana_derecha(r"(?:N[°º]|Nº|Número|ID)\s*Cliente|Cliente", texto, ancho=200) or ""
            salida["id_cliente"] = _primer_patron(patrones_id, w)

        # nro_documento
        salida["nro_documento"] = _primer_patron(patrones_ndoc, texto)
        if salida["nro_documento"] is None:
            salida["nro_documento"] = _primer_patron(patrones_ndoc_compacto, texto_compacto)

        # Fechas
        if salida["fecha_emision"] is None:
            salida["fecha_emision"] = _primer_patron(patrones_f_emision, texto)
        if salida["fecha_vencimiento"] is None:
            salida["fecha_vencimiento"] = _primer_patron(patrones_f_venc, texto)

        # -------- Total a pagar (Enel) --------
        # 1) Monto a la derecha del label
        if salida["total_a_pagar"] is None:
            w = _ventana_derecha(r"(?:VENCIMIENTO\s+)?TOTAL\s*A\s*PAGAR|Total\s*a\s*pagar", texto, ancho=260) or texto
            v = _primer_patron(patrones_total_linea, w)
            if v:
                v = re.sub(r"\s", "", v)
                salida["total_a_pagar"] = _limpiar_monto(v)

        # 2) Monto ANTES del label (ej. "$ 1.457.759Total a pagar")
        if salida["total_a_pagar"] is None:
            w_bi = _ventana_alrededor(r"(?:VENCIMIENTO\s+)?TOTAL\s*A\s*PAGAR|Total\s*a\s*pagar", texto, izq=180, der=260) or ""
            v = _primer_patron([
                r"\$?\s*([\d\.\s]{1,18})\s*Total\s*a\s*pagar",
                r"\$?\s*([\d\.\s]{1,18})\s*TOTAL\s*A\s*PAGAR",
            ], w_bi)
            if v:
                v = re.sub(r"\s", "", v)
                salida["total_a_pagar"] = _limpiar_monto(v)

        # 3) Compacto
        if salida["total_a_pagar"] is None:
            v = _primer_patron(patrones_total_compacto, texto_compacto)
            if v:
                v = re.sub(r"\s", "", v)
                salida["total_a_pagar"] = _limpiar_monto(v)

        # 4) Alternativa "Total boleta"
        if salida["total_a_pagar"] is None:
            v = _primer_patron([
                r"Total\s+boleta\s*\$?\s*([\d\.\s]{1,18})",
                r"TOTAL\s+BOLETA\s*\$?\s*([\d\.\s]{1,18})",
            ], texto)
            if v:
                v = re.sub(r"\s", "", v)
                salida["total_a_pagar"] = _limpiar_monto(v)

        # 5) Selección robusta (si quedó <1000 o None): mayor monto alrededor del label
        if salida["total_a_pagar"] is None or (isinstance(salida["total_a_pagar"], int) and salida["total_a_pagar"] < 1000):
            w2 = _ventana_alrededor(r"(?:VENCIMIENTO\s+)?TOTAL\s*A\s*PAGAR|Total\s*a\s*pagar", texto, izq=220, der=320) or ""
            candidatos = _montos_en_texto(w2)
            if candidatos:
                elegido = max(candidatos)
                if not isinstance(salida["total_a_pagar"], int) or salida["total_a_pagar"] < 1000:
                    salida["total_a_pagar"] = elegido

        # Consumo
        if salida["consumo_periodo"] is None:
            mm = None
            for p in patrones_consumo:
                mm = re.search(p, texto, re.IGNORECASE)
                if mm:
                    salida["consumo_periodo"] = f"{mm.group(1).replace('.', '').replace(',', '.')} {mm.group(2)}"
                    break
        if salida["consumo_periodo"] is None:
            mm = None
            for p in patrones_consumo_compacto:
                mm = re.search(p, texto_compacto, re.IGNORECASE)
                if mm:
                    salida["consumo_periodo"] = f"{mm.group(1).replace('.', '').replace(',', '.')} {mm.group(2)}"
                    break

        # -------- Validaciones + estado --------
        def _es_num_estricto(s):
            return bool(re.fullmatch(r"[0-9\-–kK]+", s)) if s else False

        if salida["id_cliente"] and (len(salida["id_cliente"]) < 6 or not _es_num_estricto(salida["id_cliente"])):
            salida["id_cliente"] = None
        if salida["nro_documento"] and len(re.sub(r"\D", "", salida["nro_documento"])) < 5:
            salida["nro_documento"] = None
        if isinstance(salida["total_a_pagar"], int) and not (0 < salida["total_a_pagar"] < 1_000_000_000):
            salida["total_a_pagar"] = None

        salida["estado"] = _estado_por_campos(
            salida, ["nro_documento", "total_a_pagar", "id_cliente", "fecha_emision", "fecha_vencimiento"]
        )

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
        # Texto compacto (para patrones de palabras pegadas)
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
            r"N[º°]\s*(\d{5,})",
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
            r"Total\s*a\s*pagar(?:.{0,80})?\$?\s*([\d\.\s]{1,18})",
            r"TOTAL\s*A\s*PAGAR(?:.{0,80})?\$?\s*([\d\.\s]{1,18})",
            r"TOTAL\s*A\s*PAGAR\s*[^\d$]{0,40}\$?\s*([\d\.\s]{1,18})",
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
            w = _ventana_derecha(r"(?:VENCIMIENTO\s+)?TOTAL\s*A\s*PAGAR", texto) or texto
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

        # --- FALLBACKS sobre texto compacto (palabras pegadas) ---
        if salida["total_a_pagar"] is None:
            v = _primer_patron(
                [
                    r"TOTALAPAGAR\s*\$*([\d\.\,]+)",
                    r"TOTAL\s*A\s*PAGAR\s*\$?\s*([\d\.\,]+)",
                ],
                texto_compacto,
            )
            if v:
                v = re.sub(r"\s", "", v)
                salida["total_a_pagar"] = _limpiar_monto(v)

        if salida["nro_documento"] is None:
            patrones_nro_doc_compacto = [
                r"BOLETAELECTR[ÓO]NICA\s*N[º°]?\s*([\d]+)",
                r"N[º°]\s*([\d]+)",
            ]
            salida["nro_documento"] = _primer_patron(patrones_nro_doc_compacto, texto_compacto)

        if salida["id_cliente"] is None:
            patrones_id_compacto = [
                r"SunúmerodeCuentaes:([\d\-kK]+)",
                r"Su\s*númerode\s*Cuenta\s*es:\s*([\d\-kK]+)",
                r"(?:Nro\s*de\s*cuenta|Número\s*de\s*Cuenta)\s*([\d\-kK]+)",
            ]
            salida["id_cliente"] = (
                _primer_patron([patrones_id_compacto[0]], texto_compacto)
                or _primer_patron(patrones_id_compacto[1:], texto)
            )

        if salida["fecha_emision"] is None:
            patrones_femi_compacto = [
                r"FECHAEMISI[ÓO]N:(\d{2}-\w{3}-20\d{2})",
                r"Fechadeemisión:\s*(\d{2}\s*\w{3}\s*\d{4})",
            ]
            salida["fecha_emision"] = _primer_patron(patrones_femi_compacto, texto_compacto)

        if salida["fecha_vencimiento"] is None:
            patrones_fven_compacto = [
                r"VENCIMIENTO\s*(\d{2}-\w{3}-20\d{2})",
                r"PAGARHASTA\s*(\d{2}-\w{3}-20\d{2})",
            ]
            salida["fecha_vencimiento"] = _primer_patron(patrones_fven_compacto, texto_compacto)

        if salida["consumo_periodo"] is None:
            patrones_cons_compacto = [
                r"CONSUMOAGUAPOTABLENOPUNTA\s*([\d\.,]+)",
                r"CONSUMO\s*AGUA\s*([\d\.,]+)",
                r"CONSUMOTOTAL\s*\([\s\w\d]+\)\s*=\s*([\d\.,]+)",
                r"CONSUMOTOTAL\s*([\d\.,]+)",
            ]
            c = _primer_patron(patrones_cons_compacto, texto_compacto)
            if c:
                c_norm = c.replace(".", "").replace(",", ".")
                salida["consumo_periodo"] = f"{c_norm} m3"

        # === Fallback específico para nro_documento en tablas y palabras ===
        if salida["nro_documento"] is None:
            try:
                with pdfplumber.open(str(path_pdf)) as pdf_f:
                    encontrado = None
                    for pagina in pdf_f.pages:
                        try:
                            tablas = pagina.extract_tables() or []
                        except Exception:
                            tablas = []
                        for t in tablas:
                            for fila in t:
                                if not fila:
                                    continue
                                celdas = [(c or "").strip() for c in fila]
                                for i, celda in enumerate(celdas):
                                    if re.search(r"(Folio|Documento|Boleta|N[º°])", celda, re.IGNORECASE):
                                        if i + 1 < len(celdas):
                                            val = (celdas[i + 1] or "").strip()
                                            val_num = re.sub(r"\D", "", val)
                                            if len(val_num) >= 5:
                                                encontrado = val_num
                                                break
                                if encontrado:
                                    break
                            if encontrado:
                                break
                    if encontrado:
                        salida["nro_documento"] = encontrado

                    if salida["nro_documento"] is None:
                        encontrado2 = None
                        for pagina in pdf_f.pages:
                            try:
                                words = pagina.extract_words(use_text_flow=True, keep_blank_chars=False)
                            except Exception:
                                words = []
                            lineas = {}
                            for w in words:
                                y = round(w.get("top", 0), 1)
                                lineas.setdefault(y, []).append(w)
                            for y, ws in lineas.items():
                                ws = sorted(ws, key=lambda z: z.get("x0", 0))
                                for i, w in enumerate(ws):
                                    if re.search(r"(Folio|Documento|Boleta|N[º°])", w.get("text", ""), re.IGNORECASE):
                                        for j in range(i + 1, len(ws)):
                                            derecha = ws[j]
                                            dx = derecha.get("x0", 0) - w.get("x1", 0)
                                            if 0 <= dx <= 150:
                                                val = derecha.get("text", "").strip()
                                                val_num = re.sub(r"\D", "", val)
                                                if len(val_num) >= 5:
                                                    encontrado2 = val_num
                                                    break
                                            else:
                                                break
                                if encontrado2:
                                    break
                            if encontrado2:
                                salida["nro_documento"] = encontrado2
            except Exception:
                pass

        # --- SELECCIÓN ROBUSTA: mayor monto en ventana TOTAL A PAGAR ---
        if salida["total_a_pagar"] is None or (isinstance(salida["total_a_pagar"], int) and salida["total_a_pagar"] < 1000):
            w2 = _ventana_derecha(r"(?:VENCIMIENTO\s+)?TOTAL\s*A\s*PAGAR", texto, ancho=240) or ""
            candidatos = _montos_en_texto(w2)
            if candidatos:
                candidato = max(candidatos)
                if not isinstance(salida["total_a_pagar"], int) or salida["total_a_pagar"] < 1000:
                    salida["total_a_pagar"] = candidato

        # Estado + validaciones
        salida["estado"] = _estado_por_campos(
            salida, ["nro_documento", "total_a_pagar", "id_cliente", "fecha_emision", "fecha_vencimiento"]
        )

        def _es_num_esperado(s):
            return bool(re.fullmatch(r"[0-9\-–kK]+", s)) if s else False

        if salida["id_cliente"] and len(salida["id_cliente"]) < 6:
            salida["id_cliente"] = None
        if salida["id_cliente"] and not _es_num_esperado(salida["id_cliente"]):
            salida["id_cliente"] = None
        if salida["nro_documento"] and len(re.sub(r"\D", "", salida["nro_documento"])) < 5:
            salida["nro_documento"] = None
        if isinstance(salida["total_a_pagar"], int):
            if not (0 < salida["total_a_pagar"] < 1_000_000_000):
                salida["total_a_pagar"] = None

        salida["estado"] = _estado_por_campos(
            salida, ["nro_documento", "total_a_pagar", "id_cliente", "fecha_emision", "fecha_vencimiento"]
        )

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
            resultados.append(
                {
                    "archivo_pdf": pdf.name,
                    "empresa": None,
                    "nro_documento": None,
                    "total_a_pagar": None,
                    "id_cliente": None,
                    "fecha_emision": None,
                    "fecha_vencimiento": None,
                    "consumo_periodo": None,
                    "estado": "PARCIAL",
                }
            )

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
