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
   de torta, y elegir qué licitación del Tesoro (de todas las que se fueron cargando, sobre todo desde el
   registro de Colocaciones de Deuda) cruzar con el movimiento de la base monetaria e instrumentos del BCRA
   alrededor de esa fecha.

## Ver el sitio

- Licitaciones: **https://failofaustina-gif.github.io/licitaciones/**
- Indicadores monetarios: **https://failofaustina-gif.github.io/licitaciones/monetario.html**

## Cómo se actualiza

- **Licitaciones**: manual. Subís un `datos_informe.xlsx` nuevo y la página se regenera sola (ver más abajo).
- **Indicadores monetarios**: automático. Un workflow de GitHub Actions corre de lunes a viernes a las 19:00
  (hora Argentina), descarga la última versión de `series.xlsm` desde la web del BCRA, la commitea al repo
  si cambió, y regenera `monetario.html`. También se puede disparar a mano desde la pestaña **Actions** →
  *Actualizar sitio* → **Run workflow**.

## Cómo actualizar las licitaciones (para la próxima)

Hay tres formas. Las dos primeras quedan **publicadas para todo el mundo**; la tercera es un cargador
rápido solo en tu navegador, para previsualizar antes de decidir si vale la pena subir el archivo.

### A. Colocaciones de deuda (Excel) — recomendada, cubre muchas licitaciones de una

Esta es la fuente principal. `colocaciones_deuda.xlsx` es el registro de la Oficina Nacional de Crédito
Público (hojas Bonos/Letras/Otras Operaciones) — una tabla con columnas fijas, no cambia de formato de una
publicación a otra. Cada vez que se genera el sitio, `build_monetario.py` lo lee entero y arma **una
licitación por cada fecha de colocación distinta que tenga adentro**, archivándolas en
`licitaciones_historial/` — así una sola carga puede agregar docenas de licitaciones de golpe.

1. Cuando tengas una versión nueva del registro (la Oficina Nacional de Crédito Público la va actualizando
   periódicamente), reemplazá `colocaciones_deuda.xlsx` en el repo: **Add file → Upload files** → arrastrás
   el Excel nuevo (con ese nombre exacto) → **Commit changes**.
2. El deploy se dispara solo en cualquier push a `main`. Esperá el tilde verde en **Actions** (1-2 minutos).
3. Refrescá `monetario.html` — las licitaciones nuevas ya están en el selector.

Limitación: el registro solo tiene lo efectivamente **adjudicado** (VN, VE, moneda, vencimiento) — no
incluye ofertas recibidas ni montos ofertados, ni vencimientos del día (por eso no calcula rollover
automáticamente). Tampoco se usan las hojas "Letras ISP" (son deuda intra-Estado, no de mercado, y mezclan
magnitudes no comparables) ni "Canjes- Conversiones" (tiene una estructura de tabla totalmente distinta).

### B. Comunicado de resultados (PDF) — complementa a la A con ofertas/ofertado

`build_informe.py` sigue usando `datos_informe.xlsx` (con los datos del comunicado en PDF de la Secretaría
de Finanzas: ofertas recibidas, montos ofertados, rollover) para el informe de licitaciones y para sumar
esos datos — que el Excel de colocaciones no tiene — a la licitación de esa fecha puntual.

1. Editá `datos_informe.xlsx` con los datos de la licitación nueva (hojas: Portada, Rollover, Tasas,
   Coyuntura, Composición, Timeline, Monetización, Fuentes).
2. Subilo al repo reemplazando el archivo actual: **Add file → Upload files** → arrastrás el Excel → **Commit changes**.
3. Esperá el tilde verde en **Actions** (1-2 minutos) y refrescá.

Al regenerar la página, esta licitación se archiva automáticamente en `licitaciones_historial/` y se
combina con lo que ya haya ahí para la misma fecha (por ejemplo, si el Excel de colocaciones ya tenía esa
fecha, no se pisan entre sí — se completan).

### C. Cargador en la página — al toque, pero solo en tu navegador (para previsualizar)

En la pestaña **Licitación ↔ Monetario** hay un cargador que acepta los mismos dos archivos (Excel de
colocaciones y/o PDF del comunicado) directamente en el navegador, sin pasar por el repositorio. Sirve para
ver cómo va a quedar antes de decidir si lo subís de verdad (vías A/B) — pero **no queda guardado**: es
visible solo en esa pestaña mientras no se recargue la página, y no lo ve nadie más que entre al sitio.
El PDF, además, cambia de formato de un comunicado a otro y el parseo es heurístico (busca patrones de
texto, no lee una tabla real); cuando el formato no coincide con el que sabe leer, hay un chequeo de
sanidad que corta el proceso y avisa en vez de mostrar datos mezclados o incorrectos.

## Estructura

- `colocaciones_deuda.xlsx` — registro de Colocaciones de Deuda de la Oficina Nacional de Crédito Público
  (fuente principal de licitaciones — ver sección A de arriba). Se reemplaza entero cada vez que hay una
  versión nueva; `build_monetario.py` lo relee completo en cada build.
- `datos_informe.xlsx` — los datos del comunicado de resultados en PDF de la licitación más reciente
  (ofertas, ofertado, rollover — ver sección B de arriba).
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

