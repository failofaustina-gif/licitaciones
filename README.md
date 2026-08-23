# Seguimiento de licitaciones e indicadores monetarios

Dos páginas sobre la coyuntura monetaria argentina, publicadas en el mismo sitio:

1. **Licitaciones del Tesoro** — informe por licitación de deuda de la Secretaría de Finanzas:
   rollover, tasas vs. licitación anterior, composición de lo colocado, vencimientos e impacto monetario.
2. **Indicadores monetarios (BCRA)** — explorador interactivo con las ~150 series de
   [`series.xlsm`](https://www.bcra.gob.ar/datos-monetarios-diarios/) (base monetaria, reservas, depósitos,
   préstamos, tasas de mercado e instrumentos del BCRA), actualizado automáticamente todos los días hábiles.
   Se puede arrastrar cualquier variable a un gráfico, armar cálculos propios (por ejemplo, una aproximación
   de reservas netas, o de dinero interno vs. externo / inside-outside money, con dos casilleros arrastrables
   y coeficientes editables), expresar cualquier gráfico o cálculo como % del PBI, ver composiciones en gráfico
   de torta, y elegir qué licitación del Tesoro (de todas las que se fueron cargando en `datos_informe.xlsx`
   a lo largo del tiempo) cruzar con el movimiento de la base monetaria e instrumentos del BCRA alrededor de
   esa fecha.

## Ver el sitio

- Licitaciones: **https://failofaustina-gif.github.io/licitaciones/**
- Indicadores monetarios: **https://failofaustina-gif.github.io/licitaciones/monetario.html**

## Cómo se actualiza

- **Licitaciones**: manual. Subís un `datos_informe.xlsx` nuevo y la página se regenera sola (ver más abajo).
- **Indicadores monetarios**: automático. Un workflow de GitHub Actions corre de lunes a viernes a las 19:00
  (hora Argentina), descarga la última versión de `series.xlsm` desde la web del BCRA, la commitea al repo
  si cambió, y regenera `monetario.html`. También se puede disparar a mano desde la pestaña **Actions** →
  *Actualizar sitio* → **Run workflow**.

## Cómo actualizar la licitación (para la próxima)

Hay dos formas: la manual de siempre (queda guardada para todo el mundo) y un cargador rápido en la
propia página (solo para vos, no se guarda).

### A. Manual — queda publicada en el sitio para siempre

1. Editá `datos_informe.xlsx` con los datos de la licitación nueva (hojas: Portada, Rollover, Tasas,
   Coyuntura, Composición, Timeline, Monetización, Fuentes).
2. Subilo al repo reemplazando el archivo actual: **Add file → Upload files** → arrastrás el Excel → **Commit changes**.
3. Andá a la pestaña **Actions** y esperá el tilde verde (1-2 minutos).
4. Refrescá la URL de arriba — ya está actualizada.

Al regenerar la página, la licitación que estaba en `datos_informe.xlsx` se archiva automáticamente en
`licitaciones_historial/` (un JSON por fecha), así queda disponible en el selector de la pestaña
**Licitación ↔ Monetario** aunque después subas una licitación nueva que sobreescriba el Excel.

### B. Cargador en la página — al toque, pero solo en tu navegador

En la pestaña **Licitación ↔ Monetario** hay un cargador que acepta directamente los archivos oficiales,
sin pasar por `datos_informe.xlsx`:

- **Comunicado de resultados (PDF)** de la Secretaría de Finanzas → se parsea en el navegador (pdf.js) y
  arma automáticamente ofertas recibidas, VE ofertado/adjudicado, el detalle por instrumento y la
  composición de lo emitido.
- **Colocaciones de deuda (Excel, opcional)** de la Oficina Nacional de Crédito Público → se parsea con
  SheetJS y, si alguna fila coincide con la fecha de la licitación que estás viendo, marca qué instrumentos
  son nuevos y cuáles son reaperturas.

Todo el procesamiento pasa en el navegador — nada se sube a ningún servidor. Por eso mismo **no queda
guardado**: es visible solo en esa pestaña del navegador mientras no se recargue la página, y no lo ve
nadie más que entre al sitio. Para que quede publicado de forma permanente, hay que pasar el PDF (y el
Excel, si aplica) para sumarlos al repositorio con la vía A.

Limitación conocida: como ninguna de las dos fuentes trae los "vencimientos del día", el cargador no
calcula rollover automáticamente (esa cifra sigue siendo manual, vía `datos_informe.xlsx`). Tampoco cruza
por nombre de instrumento entre PDF y Excel — lo hace por fecha de colocación, así que si las fechas de
los dos archivos no se solapan, no va a mostrar ningún cruce (no es un error, simplemente no hay nada que
cruzar).

## Estructura

- `datos_informe.xlsx` — los datos de la licitación más reciente (lo único que hay que tocar para actualizar).
- `pbi_trimestral.json` — PBI nominal trimestral de INDEC (a tasa anualizada, igual que el Cuadro 8 de INDEC),
  usado para el toggle "% del PBI" en Gráfico, Calculadora e Insights. Se actualiza a mano cada vez que INDEC
  publica un trimestre nuevo (~3 meses de rezago) — no hay un endpoint estable para automatizarlo. Instrucciones
  de dónde sacar el dato nuevo adentro del archivo.
- `licitaciones_historial/` — un JSON por cada licitación ya procesada (nombrado por fecha), generado y
  commiteado automáticamente por el workflow; es lo que alimenta el selector de licitaciones. No hace falta
  tocarlo a mano.
- `series.xlsm` — datos monetarios diarios del BCRA, se actualiza solo (no hace falta tocarlo a mano).
- `build_informe.py` — genera `informe.html`/`informe.pdf` (licitaciones) a partir de `datos_informe.xlsx`.
- `build_monetario.py` — vuelca **todas** las columnas numéricas de `series.xlsm` (más el cruce con
  `datos_informe.xlsx` y el historial en `licitaciones_historial/`) a un JSON embebido en `monetario.html`;
  ese JSON alimenta el explorador interactivo (arrastrar y soltar, calculadora, torta, insights, selector de
  licitaciones). Las tarjetas fijas de la pestaña "Resumen" y del cruce con licitaciones están definidas por
  id de serie (`RESUMEN_IDS` en Python y en el `<script>` del HTML) — se pueden cambiar ahí.
- `_monetario_template.html` — plantilla HTML/CSS/JS del explorador (todo el frontend); `build_monetario.py`
  le inyecta el JSON de datos. Si se quiere tocar el diseño o la lógica de la página, es este archivo.
  Incluye el cargador de PDF/Excel de licitaciones (carga diferida de pdf.js y SheetJS desde CDN, parseo
  100% client-side, no persiste — ver sección B de arriba).
- `.github/workflows/build.yml` — automatización: descarga `series.xlsm`, corre ambos scripts y publica en
  GitHub Pages, tanto en cada push a `main` como todos los días hábiles por horario.
- `informe_28-07-2026.pdf` — copia del primer informe generado.

