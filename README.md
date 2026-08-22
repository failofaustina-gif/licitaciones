# Seguimiento de licitaciones e indicadores monetarios

Dos páginas sobre la coyuntura monetaria argentina, publicadas en el mismo sitio:

1. **Licitaciones del Tesoro** — informe por licitación de deuda de la Secretaría de Finanzas:
   rollover, tasas vs. licitación anterior, composición de lo colocado, vencimientos e impacto monetario.
2. **Indicadores monetarios (BCRA)** — base monetaria, reservas internacionales, depósitos, préstamos,
   tasas de mercado e instrumentos del BCRA, actualizados automáticamente todos los días hábiles a partir
   del archivo [`series.xlsm`](https://www.bcra.gob.ar/datos-monetarios-diarios/) del Banco Central.

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

## Estructura

- `datos_informe.xlsx` — los datos de licitaciones (lo único que hay que tocar para actualizar ese informe).
- `series.xlsm` — datos monetarios diarios del BCRA, se actualiza solo (no hace falta tocarlo a mano).
- `build_informe.py` — genera `informe.html`/`informe.pdf` (licitaciones) a partir de `datos_informe.xlsx`.
- `build_monetario.py` — genera `monetario.html` (indicadores BCRA) a partir de `series.xlsm`. Las series
  que muestra están definidas en `SERIES_CONFIG`, al principio del archivo — se puede sumar o sacar
  variables editando esa lista (hoja, columna y etiqueta de cada una).
- `.github/workflows/build.yml` — automatización: descarga `series.xlsm`, corre ambos scripts y publica en
  GitHub Pages, tanto en cada push a `main` como todos los días hábiles por horario.
- `informe_28-07-2026.pdf` — copia del primer informe generado.

