# Seguimiento de licitaciones e indicadores monetarios

Dos páginas sobre la coyuntura monetaria argentina, publicadas en el mismo sitio, actualizadas
automáticamente a partir de **dos únicas fuentes oficiales**:

- **`colocaciones_deuda.xlsx`** — registro de Colocaciones de Deuda de la Oficina Nacional de Crédito
  Público. Fuente de todo lo relacionado con licitaciones del Tesoro.
- **`series.xlsm`** — datos monetarios diarios del BCRA. Fuente de todo lo relacionado con indicadores
  monetarios (base monetaria, reservas, depósitos, préstamos, tasas, instrumentos del BCRA).

No hay PDFs ni carga manual de ningún otro archivo: ambas fuentes son tablas con columnas fijas que no
cambian de formato de una publicación a otra, así que todo el sitio se regenera solo.

1. **Licitaciones del Tesoro** — portada con el balance simplificado del BCRA (activo/pasivo) y las
   últimas licitaciones: valor efectivo adjudicado, composición e instrumentos colocados.
2. **Indicadores monetarios (BCRA)** — explorador interactivo con las ~150 series de `series.xlsm`,
   actualizado automáticamente todos los días hábiles. Se puede arrastrar cualquier variable a un gráfico,
   armar cálculos propios (por ejemplo, una aproximación de reservas netas, o de dinero interno vs. externo),
   expresar cualquier gráfico o cálculo como % del PBI, ver composiciones en gráfico de torta, y elegir qué
   licitación del Tesoro cruzar con el movimiento de la base monetaria e instrumentos del BCRA alrededor de
   esa fecha (con tarjetas que muestran cuánto varió cada serie en esa ventana), y una pestaña
   **Calendario** con las próximas publicaciones (IPC y PBI trimestral de INDEC, licitaciones del Tesoro).

## Ver el sitio

- Licitaciones: **https://failofaustina-gif.github.io/licitaciones/**
- Indicadores monetarios: **https://failofaustina-gif.github.io/licitaciones/monetario.html**

## Cómo se actualiza

Todo es automático. Un workflow de GitHub Actions corre de lunes a viernes a las 19:00 (hora Argentina),
descarga la última versión de `series.xlsm` desde la web del BCRA, la commitea al repo si cambió, lee
`colocaciones_deuda.xlsx` tal como esté en el repo, y regenera `informe.html` (portada) y `monetario.html`.
También se puede disparar a mano desde la pestaña **Actions** → *Actualizar sitio* → **Run workflow**, y
corre solo en cualquier push a `main`.

## Cómo actualizar las licitaciones

`colocaciones_deuda.xlsx` es el registro de la Oficina Nacional de Crédito Público (hojas Bonos, Letras,
Otras Operaciones) — una tabla con columnas fijas. Cada vez que se genera el sitio, `build_monetario.py` lo
lee entero y arma **una licitación por cada fecha de colocación distinta que tenga adentro**, archivándolas
en `licitaciones_historial/` — así una sola carga puede agregar docenas de licitaciones de golpe.

1. Cuando tengas una versión nueva del registro (la Oficina Nacional de Crédito Público la va actualizando
   periódicamente), reemplazá `colocaciones_deuda.xlsx` en el repo: **Add file → Upload files** → arrastrás
   el Excel nuevo (con ese nombre exacto) → **Commit changes**.
2. Esperá el tilde verde en **Actions** (1-2 minutos) y refrescá el sitio — las licitaciones nuevas ya están
   en el selector y en la portada.

Limitación: el registro solo tiene lo efectivamente **adjudicado** (VN, VE, moneda, vencimiento) — no
incluye ofertas recibidas ni montos ofertados. Tampoco se usan las hojas "Letras ISP" (son deuda
intra-Estado, no de mercado, y mezclan magnitudes no comparables) ni "Canjes- Conversiones" (tiene una
estructura de tabla totalmente distinta).

Cuando la única colocación de una fecha viene de la hoja "Otras Operaciones" (que no informa Valor
Efectivo), la tarjeta de esa licitación muestra "s/d" en vez de un monto en pesos — no es un error, es que
esa hoja no reporta ese dato.

### Cargador en la página — al toque, pero solo en tu navegador (para previsualizar)

En la pestaña **Licitación ↔ Monetario** hay un cargador que acepta el mismo Excel de colocaciones
directamente en el navegador, sin pasar por el repositorio. Sirve para ver cómo va a quedar antes de
decidir si lo subís de verdad — pero **no queda guardado**: es visible solo en esa pestaña mientras no se
recargue la página, y no lo ve nadie más que entre al sitio.

## Estructura

- `colocaciones_deuda.xlsx` — registro de Colocaciones de Deuda de la Oficina Nacional de Crédito Público
  (única fuente de licitaciones). Se reemplaza entero cada vez que hay una versión nueva; `build_monetario.py`
  lo relee completo en cada build.
- `series.xlsm` — datos monetarios diarios del BCRA (única fuente de indicadores monetarios), se actualiza
  solo (no hace falta tocarlo a mano).
- `pbi_trimestral.json` — PBI nominal trimestral de INDEC (a tasa anualizada, igual que el Cuadro 8 de INDEC),
  usado para el toggle "% del PBI" en Gráfico, Calculadora e Insights. Se actualiza a mano cada vez que INDEC
  publica un trimestre nuevo (~3 meses de rezago) — no hay un endpoint estable para automatizarlo. Instrucciones
  de dónde sacar el dato nuevo adentro del archivo.
- `calendario.json` — eventos de la pestaña **Calendario**: fechas de IPC y PBI trimestral (INDEC) y de
  licitaciones del Tesoro (Secretaría de Finanzas). Igual que `pbi_trimestral.json`, es 100% manual — ninguna
  de las tres fuentes tiene un endpoint o cronograma público estable para automatizar. Para agregar una fecha
  nueva, sumá un objeto a `"eventos"` con `fecha` (`AAAA-MM-DD`), `tipo` (`ipc`, `pbi` o `licitacion`),
  `titulo` y opcionalmente `detalle`, y subí el archivo al repo — el sitio se regenera solo. Instrucciones de
  dónde sacar cada fecha (calendario semestral de INDEC, anuncios de licitación de la Secretaría de Finanzas)
  adentro del archivo.
- `licitaciones_historial/` — un JSON por cada licitación ya procesada (nombrado por fecha), generado y
  commiteado automáticamente por el workflow; es lo que alimenta el selector de licitaciones y la portada.
  No hace falta tocarlo a mano.
- `build_monetario.py` — lee `series.xlsm` y `colocaciones_deuda.xlsx`, y genera tanto `monetario.html`
  (explorador interactivo, con el JSON de datos embebido) como `informe.html` (portada con el balance
  simplificado del BCRA y las últimas licitaciones).
- `_monetario_template.html` — plantilla HTML/CSS/JS del explorador (todo el frontend); `build_monetario.py`
  le inyecta el JSON de datos. Si se quiere tocar el diseño o la lógica de la página, es este archivo.
  Incluye el cargador de Excel de licitaciones (carga diferida de SheetJS desde CDN, parseo 100%
  client-side, no persiste — ver sección de arriba).
- `.github/workflows/build.yml` — automatización: descarga `series.xlsm`, corre `build_monetario.py` y
  publica en GitHub Pages, tanto en cada push a `main` como todos los días hábiles por horario.
