# Open Order — Fechas de carga estimada

App web para cruzar el **Open Order Report de Ashley** con el **Trasladar (stock.picking) de Odoo** y proyectar fechas de llegada.

## Cómo usar

1. Abre el link de la app
2. Sube el archivo `OpenOrderReport*.xls` (exportado de Ashley)
3. Sube el archivo `Trasladar*.xlsx` (exportado de Odoo → stock.picking)
4. Clic en **Procesar archivos**
5. Descarga el Excel con las fechas asignadas

## Lógica

- Cruza por **Referencia de proveedor + SKU** (match exacto)
- Solo incluye los pedidos del backlog (los que están en el Open Order)
- Agrega `Fecha carga est.` y `Fecha est. llegada` (carga + 70 días)
- Verde = fecha asignada · Naranja = artículo en tránsito sin fecha en backlog

## Deploy local

```bash
pip install -r requirements.txt
streamlit run app.py
```
