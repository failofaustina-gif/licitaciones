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

# ---------- 2b. cargar todo el historial de licitaciones archivadas ----------
licitaciones = []
if HIST_DIR.exists():
    for f in HIST_DIR.glob("*.json"):
        try:
            licitaciones.append(json.loads(f.read_text(encoding="utf-8")))
        except (json.JSONDecodeError, OSError) as e:
            print(f"Aviso: no se pudo leer {f.name} del historial de licitaciones: {e}")
licitaciones.sort(key=lambda x: x.get("fecha") or "", reverse=True)

dashboard_data = {
    "generated_at": datetime.datetime.now().strftime("%Y-%m-%dT%H:%M:%S"),
    "as_of": as_of,
    "series": all_series,
    "licitaciones": licitaciones,
}

json_str = json.dumps(dashboard_data, ensure_ascii=False, separators=(",", ":"))
print(f"Series exportadas: {len(all_series)} | as_of={as_of} | JSON: {len(json_str)/1024:.0f} KB")

# ---------- 3. armar monetario.html ----------
TEMPLATE = (HERE / "_monetario_template.html").read_text(encoding="utf-8")
html = TEMPLATE.replace("__DASHBOARD_DATA__", json_str).replace("__AS_OF__", as_of or "s/d")

out_html = HERE / "monetario.html"
out_html.write_text(html, encoding="utf-8")
print(f"HTML generado: {out_html}")
