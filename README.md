# Seguimiento de licitaciones e indicadores monetarios

Dos páginas sobre la coyuntura monetaria argentina, publicadas en el mismo sitio:

1. **Licitaciones del Tesoro** — informe por licitación de deuda de la Secretaría de Finanzas:
   rollover, tasas vs. licitación anterior, composición de lo colocado, vencimientos e impacto monetario.
2. **Indicadores monetarios (BCRA)** — explorador interactivo con las ~150 series de
   [`series.xlsm`](https://www.bcra.gob.ar/datos-monetarios-diarios/) (base monetaria, reservas, depósitos,
   préstamos, tasas de mercado e instrumentos del BCRA), actualizado automáticamente todos los días hábiles.
   Se puede arrastrar cualquier variable a un gráfico, armar cálculos propios (por ejemplo, una aproximación
   de reservas netas), ver composiciones en gráfico de torta, ver la evolución de dinero interno vs. externo
   (inside/outside money), y elegir qué licitación del Tesoro (de todas las que se fueron cargando en
   `datos_informe.xlsx` a lo largo del tiempo) cruzar con el movimiento de la base monetaria e instrumentos
   del BCRA alrededor de esa fecha.

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

1. Editá `datos_informe.xlsx` con los datos de la licitación nueva (hojas: Portada, Rollover, Tasas,
   Coyuntura, Composición, Timeline, Monetización, Fuentes).
2. Subilo al repo reemplazando el archivo actual: **Add file → Upload files** → arrastrás el Excel → **Commit changes**.
3. Andá a la pestaña **Actions** y esperá el tilde verde (1-2 minutos).
4. Refrescá la URL de arriba — ya está actualizada.

Al regenerar la página, la licitación que estaba en `datos_informe.xlsx` se archiva automáticamente en
`licitaciones_historial/` (un JSON por fecha), así queda disponible en el selector de la pestaña
**Licitación ↔ Monetario** aunque después subas una licitación nueva que sobreescriba el Excel.

## Estructura

- `datos_informe.xlsx` — los datos de la licitación más reciente (lo único que hay que tocar para actualizar).
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
- `.github/workflows/build.yml` — automatización: descarga `series.xlsm`, corre ambos scripts y publica en
  GitHub Pages, tanto en cada push a `main` como todos los días hábiles por horario.
- `informe_28-07-2026.pdf` — copia del primer informe generado.

