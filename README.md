# PAVIMAX — Panel

Panel web para que el operario vea pedidos pendientes de PAVIMAX (asfalto en frío embolsado), los marque como entregados (con foto del remito) y registre la producción diaria. Lleva el control de stock automático.

- **App pública**: https://marcoskatz-cmd.github.io/pavimax/
- **Backend**: Google Sheets + Apps Script
- **Fotos**: Google Drive

## Estructura

- `index.html` — la web (frontend completo, sin build)
- `Code.gs` — backend de Apps Script (pegar en script.google.com)
- `appsscript.json` — manifiesto del Apps Script
- `SETUP.md` — pasos para configurar desde cero

## Cómo se actualiza el panel

Editar `index.html` → `git add . && git commit -m "..." && git push origin main`. GitHub Pages redeploya solo en 1-3 min.

## Cómo se actualiza el backend

Editar `Code.gs` localmente → copiarlo al editor de Apps Script en Google → **Implementar → Administrar implementaciones → editar → Versión nueva → Implementar**. La URL no cambia.

## Si un pedido "desaparece" (no está en pendientes ni en entregados)

Hasta nov-2026 las tandas se escribían en un orden fijo de columnas; como la columna `granel_tn` quedó al final de la hoja **ENTREGAS_PARCIALES**, los pedidos nuevos se guardaban corridos una columna y el panel no los mostraba. El backend ya escribe según el encabezado real de la hoja. Para arreglar las filas que quedaron corridas, una vez (es idempotente, se puede repetir sin riesgo):

- en el editor de Apps Script, ejecutar la función `repararTandasCorridas`, o
- abrir en el navegador `URL_DEL_WEB_APP?action=reparar` (devuelve cuántas filas arregló).

También agrega la fila `granel_tn` en la hoja PRODUCTO si falta (para poder cargarle costo y precio).
