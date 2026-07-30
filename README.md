# Seguimiento de licitaciones del Tesoro

Informe visual (una página) sobre cada licitación de deuda de la Secretaría de Finanzas:
rollover, tasas vs. licitación anterior, composición de lo colocado, vencimientos e impacto monetario.

## Ver el informe

**https://failofaustina-gif.github.io/licitaciones/**

Esta página se actualiza sola cada vez que se sube un `datos_informe.xlsx` nuevo (no hace falta descargar nada ni correr comandos).

## Cómo actualizarlo (para la próxima licitación)

1. Editá `datos_informe.xlsx` con los datos de la licitación nueva (hojas: Portada, Rollover, Tasas,
   Coyuntura, Composición, Timeline, Monetización, Fuentes).
2. Subilo al repo reemplazando el archivo actual: **Add file → Upload files** → arrastrás el Excel → **Commit changes**.
3. Andá a la pestaña **Actions** y esperá el tilde verde (1-2 minutos).
4. Refrescá la URL de arriba — ya está actualizada.

## Estructura

- `datos_informe.xlsx` — los datos (lo único que hay que tocar para actualizar).
- `build_informe.py` — genera el HTML/PDF a partir del Excel.
- `.github/workflows/build.yml` — automatización: corre el script y publica en GitHub Pages en cada actualización del Excel.
- `informe_28-07-2026.pdf` — copia del primer informe generado.

