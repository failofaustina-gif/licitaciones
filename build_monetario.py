#!/usr/bin/env python3
"""
Genera monetario.html a partir de series.xlsm (BCRA - datos monetarios diarios)
https://www.bcra.gob.ar/datos-monetarios-diarios/

Uso: python3 build_monetario.py [ruta_xlsm]
"""
import sys, datetime
import openpyxl
from pathlib import Path

HERE = Path(__file__).parent
XLSM = Path(sys.argv[1]) if len(sys.argv) > 1 else HERE / "series.xlsm"

# ---------- qué columna de cada hoja mostramos ----------
# (columna 1-indexada tal cual aparece en series.xlsm; "kind" define el formato)
SERIES_CONFIG = [
    dict(sheet="BASE MONETARIA", col=30, key="bm_stock",
         label="Base Monetaria (circulante + cta. cte. en el BCRA)", group="Base Monetaria", kind="ars"),
    dict(sheet="BASE MONETARIA", col=3, key="bm_var",
         label="Variación diaria de la Base Monetaria", group="Base Monetaria", kind="ars"),
    dict(sheet="RESERVAS", col=3, key="reservas_stock",
         label="Reservas Internacionales (stock)", group="Reservas Internacionales", kind="usd"),
    dict(sheet="RESERVAS", col=16, key="tc_ref",
         label="Tipo de cambio de referencia (BCRA)", group="Reservas Internacionales", kind="fx"),
    dict(sheet="DEPOSITOS", col=23, key="dep_total",
         label="Depósitos totales del sistema", group="Depósitos", kind="ars"),
    dict(sheet="DEPOSITOS", col=24, key="dep_privado",
         label="Depósitos totales - sector privado", group="Depósitos", kind="ars"),
    dict(sheet="PRESTAMOS", col=21, key="prest_total",
         label="Préstamos al sector privado (total, en pesos)", group="Préstamos", kind="ars"),
    dict(sheet="TASAS DE MERCADO", col=2, key="tasa_pf",
         label="Plazo fijo - tasa promedio general en pesos", group="Tasas de Mercado", kind="rate"),
    dict(sheet="TASAS DE MERCADO", col=9, key="tamar",
         label="TAMAR en pesos", group="Tasas de Mercado", kind="rate"),
    dict(sheet="TASAS DE MERCADO", col=12, key="badlar",
         label="BADLAR en pesos", group="Tasas de Mercado", kind="rate"),
    dict(sheet="INSTRUMENTOS DEL BCRA", col=11, key="tpm",
         label="Tasa de política monetaria", group="Instrumentos del BCRA", kind="rate"),
    dict(sheet="INSTRUMENTOS DEL BCRA", col=2, key="pases_pasivos",
         label="Pases pasivos en pesos (stock)", group="Instrumentos del BCRA", kind="ars"),
    dict(sheet="INSTRUMENTOS DEL BCRA", col=7, key="lebac_nobac",
         label="Letras y notas del BCRA en pesos, excl. LELIQ (stock)", group="Instrumentos del BCRA", kind="ars"),
]
GROUP_ORDER = ["Base Monetaria", "Reservas Internacionales", "Depósitos", "Préstamos",
               "Tasas de Mercado", "Instrumentos del BCRA"]

# ---------- lectura de series.xlsm ----------
wb = openpyxl.load_workbook(XLSM, data_only=True, read_only=True)

def read_columns(sheet_name, cols, min_row=8):
    """Una sola pasada secuencial por la hoja. Corta apenas la fecha deja de
    ser creciente: al final de cada hoja el BCRA agrega filas anuales/mensuales
    de referencia que no son parte de la serie diaria continua."""
    ws = wb[sheet_name]
    out = {c: [] for c in cols}
    last_date = None
    for row in ws.iter_rows(min_row=min_row, values_only=True):
        d = row[0]
        if not isinstance(d, (datetime.date, datetime.datetime)):
            continue
        if last_date is not None and d < last_date:
            break
        last_date = d
        for c in cols:
            idx = c - 1
            v = row[idx] if idx < len(row) else None
            if isinstance(v, (int, float)):
                out[c].append((d, v))
    return out

by_sheet = {}
for cfg in SERIES_CONFIG:
    by_sheet.setdefault(cfg["sheet"], set()).add(cfg["col"])

sheet_data = {sheet: read_columns(sheet, sorted(cols)) for sheet, cols in by_sheet.items()}

as_of = None
for cfg in SERIES_CONFIG:
    pts = sheet_data[cfg["sheet"]][cfg["col"]]
    if pts:
        d = pts[-1][0]
        if as_of is None or d > as_of:
            as_of = d

# ---------- helpers de formato (es-AR) ----------
def ar_number(n, decimals=0):
    s = f"{n:,.{decimals}f}"
    s = s.replace(",", "§").replace(".", ",").replace("§", ".")
    return s

def fmt_value(v, kind):
    av = abs(v)
    if kind == "rate":
        return f"{ar_number(v, 2)}%"
    if kind == "fx":
        return f"${ar_number(v, 2)}"
    if kind == "usd":
        return f"US$ {ar_number(v, 0)} M"
    if kind == "ars":
        if av >= 1_000_000:
            return f"$ {ar_number(v / 1_000_000, 2)} bn"
        if av >= 1_000:
            return f"$ {ar_number(v / 1_000, 1)} mil M"
        return f"$ {ar_number(v, 0)} M"
    return ar_number(v, 2)

def fmt_date(d):
    return d.strftime("%d/%m/%Y")

MESES_ABR = {1:"ene",2:"feb",3:"mar",4:"abr",5:"may",6:"jun",7:"jul",8:"ago",9:"sep",10:"oct",11:"nov",12:"dic"}
def fmt_date_short(d):
    return f"{d.day:02d} {MESES_ABR[d.month]}"

def pct_delta(curr, prev):
    if prev == 0:
        return None
    return (curr - prev) / abs(prev) * 100

# ---------- sparkline SVG ----------
SPARK_W, SPARK_H, SPARK_PAD = 320, 64, 4

def sparkline_svg(points, color):
    vals = [v for _, v in points]
    if len(vals) < 2:
        return ""
    lo, hi = min(vals), max(vals)
    span = (hi - lo) or 1
    n = len(vals)
    xs = [SPARK_PAD + i * (SPARK_W - 2*SPARK_PAD) / (n - 1) for i in range(n)]
    ys = [SPARK_H - SPARK_PAD - (v - lo) / span * (SPARK_H - 2*SPARK_PAD) for v in vals]
    line_pts = " ".join(f"{x:.1f},{y:.1f}" for x, y in zip(xs, ys))
    area_pts = f"{xs[0]:.1f},{SPARK_H} " + line_pts + f" {xs[-1]:.1f},{SPARK_H}"
    return f"""<svg viewBox="0 0 {SPARK_W} {SPARK_H}" class="spark" preserveAspectRatio="none">
      <polygon points="{area_pts}" fill="{color}" opacity="0.12"/>
      <polyline points="{line_pts}" fill="none" stroke="{color}" stroke-width="2" stroke-linejoin="round" stroke-linecap="round"/>
      <circle cx="{xs[-1]:.1f}" cy="{ys[-1]:.1f}" r="3.2" fill="{color}"/>
    </svg>"""

COLORS = {"Base Monetaria":"#2C5FA8","Reservas Internacionales":"#1E8A5F","Depósitos":"#5B8FD9",
          "Préstamos":"#C2571B","Tasas de Mercado":"#8A5FC2","Instrumentos del BCRA":"#0B2547"}

# ---------- armar tarjetas ----------
DAILY_WINDOW = 130   # ~6 meses hábiles para el sparkline
MONTH_BACK = 21       # ~1 mes hábil para la variación mensual

def delta_badge(curr, prev, kind):
    if prev is None:
        return ""
    if kind == "rate":
        d = curr - prev
        cls = "up" if d > 0 else ("down" if d < 0 else "flat")
        arrow = "▲" if d > 0 else ("▼" if d < 0 else "—")
        return f'<span class="chg {cls}">{arrow} {ar_number(abs(d),2)} p.p.</span>'
    pct = pct_delta(curr, prev)
    if pct is None:
        return ""
    cls = "up" if pct > 0 else ("down" if pct < 0 else "flat")
    arrow = "▲" if pct > 0 else ("▼" if pct < 0 else "—")
    return f'<span class="chg {cls}">{arrow} {ar_number(abs(pct),1)}%</span>'

cards_by_group = {g: [] for g in GROUP_ORDER}
for cfg in SERIES_CONFIG:
    pts = sheet_data[cfg["sheet"]][cfg["col"]]
    if not pts:
        continue
    last_d, last_v = pts[-1]
    prev_v = pts[-2][1] if len(pts) >= 2 else None
    month_v = pts[-1 - MONTH_BACK][1] if len(pts) > MONTH_BACK else None
    color = COLORS[cfg["group"]]
    spark = sparkline_svg(pts[-DAILY_WINDOW:], color)
    card = f"""
    <div class="card">
      <div class="card-label">{cfg['label']}</div>
      <div class="card-value">{fmt_value(last_v, cfg['kind'])}</div>
      <div class="card-date">al {fmt_date(last_d)}</div>
      <div class="card-chart">{spark}</div>
      <div class="card-deltas">
        <div>vs. dato anterior {delta_badge(last_v, prev_v, cfg['kind'])}</div>
        <div>vs. ~1 mes {delta_badge(last_v, month_v, cfg['kind'])}</div>
      </div>
    </div>"""
    cards_by_group[cfg["group"]].append(card)

sections_html = ""
for g in GROUP_ORDER:
    if not cards_by_group[g]:
        continue
    sections_html += f"""
  <div class="section"><div class="section-label">{g}</div></div>
  <div class="grid">{''.join(cards_by_group[g])}</div>"""

as_of_str = fmt_date(as_of) if as_of else "s/d"
generated_str = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")

HTML = f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Indicadores Monetarios - BCRA</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;600;700&family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
<style>
  :root{{
    --navy:#0B2547; --blue:#2C5FA8; --blue-light:#E7EFFB; --sky:#5B8FD9;
    --green:#1E8A5F; --green-light:#E6F5EE; --amber:#C2571B; --amber-light:#FCEEE3;
    --paper:#F6F7FB; --ink:#0B2547; --ink-soft:#5B6472; --line:#DCE3EF;
  }}
  *{{box-sizing:border-box;}}
  body{{margin:0;background:#DCE3EF;font-family:'Inter',sans-serif;color:var(--ink);display:flex;justify-content:center;padding:32px 12px;}}
  .page{{width:1040px;max-width:100%;background:var(--paper);border-radius:22px;overflow:hidden;box-shadow:0 30px 60px -20px rgba(11,37,71,.35);}}
  .hero{{background:linear-gradient(135deg,var(--navy) 0%,#173B72 60%,var(--blue) 100%);color:#fff;padding:36px 44px 28px;position:relative;overflow:hidden;}}
  .hero::after{{content:"";position:absolute;right:-60px;top:-60px;width:260px;height:260px;border-radius:50%;background:rgba(255,255,255,.06);}}
  .eyebrow{{font-size:12.5px;letter-spacing:.14em;text-transform:uppercase;color:#AFC7EE;font-weight:600;margin-bottom:10px;}}
  h1{{font-family:'Space Grotesk',sans-serif;font-size:26px;line-height:1.2;margin:0 0 8px;font-weight:700;max-width:680px;}}
  .hero p{{margin:0;color:#D7E3F6;font-size:14px;max-width:640px;}}
  .hero a{{color:#fff;}}
  .section{{padding:26px 44px 0;}}
  .section-label{{font-family:'Space Grotesk',sans-serif;font-size:13.5px;letter-spacing:.06em;text-transform:uppercase;color:var(--blue);font-weight:700;margin-bottom:14px;}}
  .grid{{padding:0 44px 8px;display:grid;grid-template-columns:repeat(auto-fit,minmax(280px,1fr));gap:16px;}}
  .card{{background:#fff;border:1px solid var(--line);border-radius:14px;padding:18px 18px 16px;}}
  .card-label{{font-size:12.5px;color:var(--ink-soft);line-height:1.35;min-height:32px;}}
  .card-value{{font-family:'Space Grotesk',sans-serif;font-size:23px;font-weight:700;color:var(--navy);margin-top:6px;}}
  .card-date{{font-size:11px;color:var(--ink-soft);margin-top:2px;}}
  .card-chart{{margin:10px 0 4px;}}
  .spark{{width:100%;height:56px;display:block;}}
  .card-deltas{{display:flex;justify-content:space-between;font-size:11px;color:var(--ink-soft);border-top:1px solid var(--line);padding-top:8px;margin-top:4px;}}
  .chg{{font-weight:700;margin-left:4px;}}
  .chg.up{{color:var(--blue);}} .chg.down{{color:var(--amber);}} .chg.flat{{color:var(--ink-soft);}}
  .footer{{padding:24px 44px 30px;font-size:11px;color:#8A93A3;line-height:1.7;}}
  .footer a{{color:#8A93A3;}}
</style>
</head>
<body>
<div class="page">
  <div class="hero">
    <div class="eyebrow">BCRA · Datos Monetarios Diarios</div>
    <h1>Seguimiento de indicadores monetarios de la Argentina</h1>
    <p>Base monetaria, reservas, depósitos, préstamos, tasas de mercado e instrumentos del BCRA — actualizado
       automáticamente a partir de <a href="https://www.bcra.gob.ar/datos-monetarios-diarios/">series.xlsm</a>
       del Banco Central. Último dato disponible en el archivo: <b>{as_of_str}</b>.</p>
    <p style="margin-top:10px;"><a href="index.html" style="color:#AFC7EE;">← Ver informe de licitaciones</a></p>
  </div>
  {sections_html}
  <div class="footer">
    Fuente: Gerencia de Estadísticas Monetarias - Banco Central de la República Argentina (series.xlsm,
    datos provisorios sujetos a revisión). "vs. dato anterior" compara con la observación previa de cada
    serie (diaria); "vs. ~1 mes" compara con la observación de ~21 días hábiles atrás. Página generada
    automáticamente el {generated_str} (hora del servidor de GitHub Actions).
  </div>
</div>
</body>
</html>
"""

out_html = HERE / "monetario.html"
out_html.write_text(HTML, encoding="utf-8")
print(f"HTML generado: {out_html} (as_of={as_of_str})")
