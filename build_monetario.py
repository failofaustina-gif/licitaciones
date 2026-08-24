#!/usr/bin/env python3
"""
Genera monetario.html: explorador interactivo de indicadores monetarios del BCRA
a partir de series.xlsm (https://www.bcra.gob.ar/datos-monetarios-diarios/),
cruzado con datos_informe.xlsx (licitaciones del Tesoro).

Uso: python3 build_monetario.py [ruta_xlsm] [ruta_xlsx_licitaciones]
"""
import sys, re, json, datetime
import openpyxl
from pathlib import Path

HERE = Path(__file__).parent
XLSM = Path(sys.argv[1]) if len(sys.argv) > 1 else HERE / "series.xlsm"
XLSX_LIC = Path(sys.argv[2]) if len(sys.argv) > 2 else HERE / "datos_informe.xlsx"

MAX_POINTS = 1000  # ~4 años hábiles por serie, para no inflar el JSON

SHEETS = {
    "BASE MONETARIA": "BM",
    "RESERVAS": "RES",
    "DEPOSITOS": "DEP",
    "PRESTAMOS": "PRE",
    "TASAS DE MERCADO": "TAS",
    "INSTRUMENTOS DEL BCRA": "INS",
}

# ---------- 1. leer series.xlsm (todas las columnas) ----------
wb = openpyxl.load_workbook(XLSM, data_only=True, read_only=False, keep_vba=False)

def resolve_headers(ws, header_rows=(4, 5, 6, 7)):
    """Arma el texto de cada columna combinando las filas de encabezado,
    respetando las celdas combinadas (merge) reales de la hoja."""
    grid = {}
    for r in header_rows:
        for c in range(1, ws.max_column + 1):
            v = ws.cell(row=r, column=c).value
            if v not in (None, ""):
                grid[(r, c)] = str(v).replace("\n", " ").strip()
    for m in ws.merged_cells.ranges:
        if m.max_row < min(header_rows) or m.min_row > max(header_rows):
            continue
        top_val = ws.cell(row=m.min_row, column=m.min_col).value
        if top_val in (None, ""):
            continue
        top_val = str(top_val).replace("\n", " ").strip()
        for r in range(max(m.min_row, min(header_rows)), min(m.max_row, max(header_rows)) + 1):
            for c in range(m.min_col, m.max_col + 1):
                grid[(r, c)] = top_val
    labels = {}
    for c in range(2, ws.max_column + 1):
        parts = []
        for r in header_rows:
            v = grid.get((r, c))
            if v and (not parts or parts[-1] != v):
                parts.append(v)
        labels[c] = " - ".join(parts) if parts else f"Columna {c}"
    return labels

def classify_kind(sheet, label):
    low = label.lower()
    last_segment = label.split(" - ")[-1].strip().lower()
    if sheet == "RESERVAS" and "tipo de cambio" in low:
        return "fx"
    if sheet == "TASAS DE MERCADO" or last_segment in ("tna", "tea") or "tasa de política monetaria" in low or "tasa de interés" in low:
        return "rate"
    if sheet == "RESERVAS":
        return "usd"
    if "expresados en dólares" in low:
        return "usd"
    if "expresados en pesos" in low:
        return "ars"
    if "en dólares" in low or "total dólares" in low or "total \ndólares".replace("\n", " ") in low:
        return "usd"
    return "ars"

def read_sheet(sheet_name, prefix):
    ws = wb[sheet_name]
    labels = resolve_headers(ws)
    cols = [c for c in range(2, ws.max_column + 1) if labels.get(c) and "tipo de serie" not in labels[c].lower()]
    buffers = {c: [] for c in cols}
    last_date = None
    for row in ws.iter_rows(min_row=8, values_only=True):
        d = row[0]
        if not isinstance(d, (datetime.date, datetime.datetime)):
            continue
        if last_date is not None and d < last_date:
            break
        last_date = d
        ds = d.strftime("%Y-%m-%d")
        for c in cols:
            idx = c - 1
            v = row[idx] if idx < len(row) else None
            if isinstance(v, (int, float)):
                buffers[c].append((ds, round(v, 4)))
    out = []
    for c in cols:
        pts = buffers[c][-MAX_POINTS:]
        if not pts:
            continue
        out.append({
            "id": f"{prefix}_{c}",
            "sheet": sheet_name,
            "label": labels[c],
            "kind": classify_kind(sheet_name, labels[c]),
            "data": pts,
        })
    return out

all_series = []
for sheet_name, prefix in SHEETS.items():
    all_series.extend(read_sheet(sheet_name, prefix))

as_of = None
for s in all_series:
    if s["data"]:
        d = s["data"][-1][0]
        if as_of is None or d > as_of:
            as_of = d

# ---------- 2. leer datos_informe.xlsx (cruce licitaciones) ----------
HIST_DIR = HERE / "licitaciones_historial"

licitacion = None
if XLSX_LIC.exists():
    wb2 = openpyxl.load_workbook(XLSX_LIC, data_only=True)

    def sheet_rows(name, start=2):
        ws = wb2[name]
        for row in ws.iter_rows(min_row=start, values_only=True):
            if row[0] is None:
                continue
            yield row

    portada = {r[0]: r[1] for r in wb2["Portada"].iter_rows(min_row=1, values_only=True) if r[0]}
    rollover = {r[0]: r[1] for r in sheet_rows("Rollover")}
    monet = [{"concepto": r[0], "valor": r[1], "unidad": r[2]} for r in sheet_rows("Monetizacion") if r[2]]
    composicion = [{"categoria": r[0], "monto": r[1], "pct": r[2], "color": r[3]}
                   for r in sheet_rows("Composicion") if r[3]]
    timeline = [{"instrumento": r[0], "fecha_antes": str(r[1]) if r[1] else None,
                 "fecha_ahora": str(r[2]) if r[2] else None, "caption": r[3]}
                for r in sheet_rows("Timeline")]

    MESES = {"enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6, "julio": 7,
             "agosto": 8, "septiembre": 9, "octubre": 10, "noviembre": 11, "diciembre": 12}
    fecha_lic = None
    subt = portada.get("Subtítulo", "") or ""
    m = re.search(r"(\d{1,2})\s+de\s+([a-zA-Zñáéíóú]+)\s+de\s+(\d{4})", subt)
    if m:
        dia, mes_txt, anio = m.groups()
        mes = MESES.get(mes_txt.lower())
        if mes:
            fecha_lic = f"{anio}-{mes:02d}-{int(dia):02d}"

    licitacion = {
        "titulo": portada.get("Título"),
        "subtitulo": portada.get("Subtítulo"),
        "fecha": fecha_lic,
        "rollover": rollover,
        "monetizacion": monet,
        "composicion": composicion,
        "timeline": timeline,
    }

    # Archivar esta licitación para no perderla cuando datos_informe.xlsx se
    # reemplace por la próxima. Se guarda un JSON por fecha en licitaciones_historial/.
    if fecha_lic:
        HIST_DIR.mkdir(exist_ok=True)
        (HIST_DIR / f"{fecha_lic}.json").write_text(
            json.dumps(licitacion, ensure_ascii=False, indent=2), encoding="utf-8"
        )
    else:
        print("Aviso: no se pudo determinar la fecha de la licitación, no se archivó en licitaciones_historial/.")

# ---------- 2a-bis. leer colocaciones_deuda.xlsx (registro de la Oficina Nacional de --------
# ---------- Crédito Público) y armar UNA licitación por cada fecha de colocación -----------
# Fuente mucho más robusta que datos_informe.xlsx/el PDF: es una tabla de columnas fijas
# (Nombre del Instrumento, Fecha colocación, Valor Nominal, Valor Efectivo, etc.) que no
# cambia de formato de una publicación a otra. No trae ofertas recibidas ni montos ofertados
# (eso solo lo tiene el comunicado en PDF de la Secretaría de Finanzas) — únicamente lo
# efectivamente adjudicado. Para actualizar: reemplazar colocaciones_deuda.xlsx por la
# versión nueva que publique la Oficina Nacional de Crédito Público y volver a correr esto
# (o simplemente pushear — el workflow lo hace solo).
XLSX_COLOC = Path(sys.argv[3]) if len(sys.argv) > 3 else HERE / "colocaciones_deuda.xlsx"

COLOC_SHEETS = {
    "Bonos": "bono", "Letras": "letra",
    "Otras Operaciones": "otra_operacion",
    # OJO: dos hojas quedan afuera a propósito.
    # - "Canjes- Conversiones": estructura totalmente distinta (dos sub-tablas apiladas,
    #   "Canje" y "Conversiones", esta última con columnas "Instrumento a entregar" /
    #   "Instrumento a dar de baja" en vez de "Valor Efectivo") — no hay un VE colocado
    #   comparable con las demás hojas, y tratar de encajarla en el mismo parseo termina
    #   leyendo nombres de instrumentos donde se esperan números.
    # - "Letras ISP" (intra-sector público): por definición son deuda entre organismos del
    #   Estado (p. ej. letras técnicas emitidas directamente al BCRA), no colocaciones de
    #   mercado — y de hecho mezclan magnitudes completamente distintas al resto (algunas
    #   filas de "LETRA/U$S/BCRA/..." aparecen con valores ~1000x más grandes que instrumentos
    #   comparables), así que no son comparables ni siquiera dentro de la propia hoja.
}
TIPO_LABELS_COMPOSICION = {
    "bono": "Bonos", "letra": "Letras", "letra_isp": "Letras ISP",
    "otra_operacion": "Otras operaciones", "canje": "Canjes/conversiones",
}

def excel_date_to_iso(v):
    if isinstance(v, (datetime.date, datetime.datetime)):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, (int, float)):
        return (datetime.date(1899, 12, 30) + datetime.timedelta(days=v)).strftime("%Y-%m-%d")
    return None

def find_header_row(ws, max_row=10):
    for r in range(1, max_row + 1):
        v = ws.cell(row=r, column=1).value
        if v is not None and str(v).strip() == "Nombre del Instrumento":
            return r
    return None

def parse_colocaciones_excel(path):
    wb3 = openpyxl.load_workbook(path, data_only=True)
    out = []
    for sheet_name, tipo in COLOC_SHEETS.items():
        if sheet_name not in wb3.sheetnames:
            continue
        ws = wb3[sheet_name]
        header_row = find_header_row(ws)
        if header_row is None:
            continue
        headers = []
        for c in range(1, ws.max_column + 1):
            v = ws.cell(row=header_row, column=c).value
            headers.append(re.sub(r"\s+", " ", str(v)).strip() if v not in (None, "") else "")

        def col(name):
            for i, h in enumerate(headers):
                if h.startswith(name):
                    return i
            return -1

        c_nombre, c_emision, c_venc = col("Nombre del Instrumento"), col("Fecha de emisión"), col("Vencimiento")
        c_moneda, c_moneda_origen = col("Tipo Moneda"), col("Moneda de Origen")
        c_coloc, c_vn, c_ve = col("Fecha colocación"), col("Valor Nominal"), col("Valor Efectivo")
        c_vta, c_precio, c_vida = col("Valor Técnico Adjudicado"), col("Precio de emisión"), col("Vida Promedio")
        c_resol = col("Resolución") if col("Resolución") != -1 else col("Norma")
        c_numero = col("Número")

        for row in ws.iter_rows(min_row=header_row + 1, values_only=True):
            if c_nombre == -1 or c_nombre >= len(row) or row[c_nombre] in (None, ""):
                continue
            fecha_colocacion = excel_date_to_iso(row[c_coloc]) if 0 <= c_coloc < len(row) else None
            if not fecha_colocacion:
                continue
            def get(idx):
                return row[idx] if 0 <= idx < len(row) else None
            def get_num(idx):
                # defensivo: si la columna esperada no es numérica (p. ej. el layout de la
                # hoja no coincide con lo esperado), se descarta en vez de romper el build.
                v = get(idx)
                return v if isinstance(v, (int, float)) else None
            out.append({
                "tipo": tipo, "hoja": sheet_name, "nombre": str(row[c_nombre]).strip(),
                "fecha_emision": excel_date_to_iso(get(c_emision)) if c_emision != -1 else None,
                "vencimiento": excel_date_to_iso(get(c_venc)) if c_venc != -1 else None,
                "moneda": get(c_moneda), "moneda_origen": get(c_moneda_origen),
                "fecha_colocacion": fecha_colocacion,
                "valor_nominal": get_num(c_vn), "precio_emision": get_num(c_precio), "vida_promedio": get_num(c_vida),
                "valor_efectivo": get_num(c_ve) if get_num(c_ve) is not None else get_num(c_vta),
                "resolucion": get(c_resol), "numero": get(c_numero),
            })
    return out

def es_instrumento_nuevo(nombre, fecha_colocacion, todas):
    return not any(c["nombre"] == nombre and c["fecha_colocacion"] < fecha_colocacion for c in todas)

def fmt_date_ddmmyyyy(iso):
    y, m, d = iso.split("-")
    return f"{d}/{m}/{y}"

def build_licitaciones_from_colocaciones(colocaciones):
    fechas = sorted({c["fecha_colocacion"] for c in colocaciones if c["fecha_colocacion"]})
    licitaciones_out = []
    for fecha in fechas:
        items = [dict(c, nuevo=es_instrumento_nuevo(c["nombre"], c["fecha_colocacion"], colocaciones))
                 for c in colocaciones if c["fecha_colocacion"] == fecha]
        ve_ars = sum((c["valor_efectivo"] or 0) for c in items if c["moneda"] == "MONEDA NACIONAL")
        ve_usd = sum((c["valor_efectivo"] or 0) for c in items if c["moneda"] == "MONEDA EXTRANJERA")
        by_tipo = {}
        for c in items:
            if c["moneda"] != "MONEDA NACIONAL" or not c["valor_efectivo"]:
                continue
            by_tipo[c["tipo"]] = by_tipo.get(c["tipo"], 0) + c["valor_efectivo"]
        composicion = [{"categoria": TIPO_LABELS_COMPOSICION.get(t, t), "monto": m, "pct": m / ve_ars}
                       for t, m in by_tipo.items()] if ve_ars else []
        licitaciones_out.append({
            "origen": "excel", "fecha": fecha,
            "titulo": f"Licitación del Tesoro — {fmt_date_ddmmyyyy(fecha)}",
            "subtitulo": f"Fuente: registro de Colocaciones de Deuda ({len(items)} instrumento{'s' if len(items)!=1 else ''} colocado{'s' if len(items)!=1 else ''}).",
            "resumenExcel": {"veAdjudicadoArs": ve_ars, "veAdjudicadoUsd": ve_usd, "cantidadInstrumentos": len(items)},
            "composicion": composicion, "instrumentosExcel": items,
        })
    return licitaciones_out

if XLSX_COLOC.exists():
    colocaciones_all = parse_colocaciones_excel(XLSX_COLOC)
    if colocaciones_all:
        nuevas_lic = build_licitaciones_from_colocaciones(colocaciones_all)
        HIST_DIR.mkdir(exist_ok=True)
        for lic in nuevas_lic:
            path = HIST_DIR / f"{lic['fecha']}.json"
            existing = {}
            if path.exists():
                try:
                    existing = json.loads(path.read_text(encoding="utf-8"))
                except (json.JSONDecodeError, OSError):
                    existing = {}
            merged = dict(existing)
            merged["fecha"] = lic["fecha"]
            merged["titulo"] = existing.get("titulo") or lic["titulo"]
            merged["subtitulo"] = existing.get("subtitulo") or lic["subtitulo"]
            merged["resumenExcel"] = lic["resumenExcel"]
            merged["instrumentosExcel"] = lic["instrumentosExcel"]
            if not existing.get("rollover"):
                merged["composicion"] = lic["composicion"]
            else:
                merged.setdefault("composicion", lic["composicion"])
            ex_origen = existing.get("origen")
            merged["origen"] = "mixto" if (ex_origen and ex_origen != "excel") else "excel"
            path.write_text(json.dumps(merged, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"Colocaciones de deuda: {len(colocaciones_all)} filas -> {len(nuevas_lic)} licitaciones archivadas en licitaciones_historial/.")
    else:
        print(f"Aviso: {XLSX_COLOC.name} no tiene ninguna fila reconocible (¿hojas Bonos/Letras/etc. con formato distinto?).")

# ---------- 2b. cargar todo el historial de licitaciones archivadas ----------
licitaciones = []
if HIST_DIR.exists():
    for f in HIST_DIR.glob("*.json"):
        try:
            licitaciones.append(json.loads(f.read_text(encoding="utf-8")))
        except (json.JSONDecodeError, OSError) as e:
            print(f"Aviso: no se pudo leer {f.name} del historial de licitaciones: {e}")
licitaciones.sort(key=lambda x: x.get("fecha") or "", reverse=True)

# ---------- 2c. PBI nominal trimestral (para expresar series como % del PBI) ----------
# Se actualiza a mano en pbi_trimestral.json (INDEC publica con ~3 meses de rezago y no tiene
# un endpoint estable para automatizar la descarga). Los valores de ese archivo son los mismos
# que las columnas "I/II/III/IV trimestre" del Cuadro 8 de INDEC: cada uno YA está a tasa
# anualizada (no es el flujo real del trimestre), así que se usan directamente como denominador
# anual por trimestre — NO hay que sumar los 4 trimestres de un año (eso daría ~4x de más).
QTR_START = {"Q1": "01-01", "Q2": "04-01", "Q3": "07-01", "Q4": "10-01"}
pbi_quarters = []
pbi_path = HERE / "pbi_trimestral.json"
if pbi_path.exists():
    pbi_raw = json.loads(pbi_path.read_text(encoding="utf-8"))["trimestres"]
    for q in sorted(pbi_raw.keys()):
        year, qn = q.split("-Q")
        pbi_quarters.append({"trimestre": q, "desde": f"{year}-{QTR_START[f'Q{qn}']}", "anualizado": round(pbi_raw[q], 1)})
    pbi_quarters.sort(key=lambda x: x["desde"])
else:
    print("Aviso: no se encontró pbi_trimestral.json, no se va a poder expresar nada como % del PBI.")

dashboard_data = {
    "generated_at": datetime.datetime.now().strftime("%Y-%m-%dT%H:%M:%S"),
    "as_of": as_of,
    "series": all_series,
    "licitaciones": licitaciones,
    "pbi": pbi_quarters,
}

json_str = json.dumps(dashboard_data, ensure_ascii=False, separators=(",", ":"))
print(f"Series exportadas: {len(all_series)} | as_of={as_of} | JSON: {len(json_str)/1024:.0f} KB")

# ---------- 3. armar monetario.html ----------
TEMPLATE = (HERE / "_monetario_template.html").read_text(encoding="utf-8")
html = TEMPLATE.replace("__DASHBOARD_DATA__", json_str).replace("__AS_OF__", as_of or "s/d")

out_html = HERE / "monetario.html"
out_html.write_text(html, encoding="utf-8")
print(f"HTML generado: {out_html}")
