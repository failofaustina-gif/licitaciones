# Seguimiento de licitaciones del Tesoro

Informe visual (una página) sobre cada licitación de deuda de la Secretaría de Finanzas:
rollover, tasas vs. licitación anterior, composición de lo colocado, vencimientos e impacto monetario.

## Cómo actualizarlo

1. Abrí `datos_informe.xlsx` y editá los datos de la licitación nueva (hojas: Portada, Rollover, Tasas,
   Coyuntura, Composición, Timeline, Monetización, Fuentes).
2. Corré:
   ```bash
   pip install openpyxl playwright
   playwright install chromium
   python3 build_informe.py
   ```
3. Esto genera `informe.html` e `informe.pdf` actualizados en la misma carpeta.

## Estructura

- `datos_informe.xlsx` — los datos (lo único que hay que tocar para actualizar).
- `build_informe.py` — genera el HTML/PDF a partir del Excel.
- `informe_28-07-2026.pdf` — último informe generado.

## Colaborar

Ambos pueden editar `datos_informe.xlsx` y correr el script. Recomendado: cada licitación en un
commit separado, nombrando el PDF con la fecha (`informe_DD-MM-AAAA.pdf`).
