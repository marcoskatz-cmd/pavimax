# Setup del panel PAVIMAX

Sigue estos pasos UNA SOLA VEZ (Marcos, ~15 minutos).

## 1. Crear la planilla de Google

1. Andá a https://sheets.new y nombrala **PAVIMAX**.
2. Copiá el ID de la URL (el pedazo entre `/d/` y `/edit`):
   `https://docs.google.com/spreadsheets/d/`**`ESTE_ID`**`/edit`

## 2. Crear la carpeta de Drive para los remitos

1. Andá a https://drive.google.com → Nuevo → Carpeta → llamala **PAVIMAX - Remitos**.
2. Abrila y copiá el ID de la URL:
   `https://drive.google.com/drive/folders/`**`ESTE_ID`**

## 3. Crear el Apps Script

1. Desde la planilla abierta: **Extensiones → Apps Script**.
2. Borrá el `Code.gs` por defecto y pegá todo el contenido de `Code.gs` de este repo.
3. En las primeras líneas reemplazá:
   ```js
   const SHEET_ID  = 'PEGAR_AQUI_EL_ID_DE_LA_PLANILLA';
   const FOLDER_ID = 'PEGAR_AQUI_EL_ID_DE_LA_CARPETA_REMITOS';
   ```
   con los IDs de los pasos 1 y 2.
4. Guardá (Ctrl+S). Le ponés nombre al proyecto: **PAVIMAX Backend**.
5. En el panel izquierdo abrí el archivo de manifest (`appsscript.json`) — si no aparece, andá a ⚙️ → "Mostrar archivo de manifiesto appsscript.json", y pegá el contenido de `appsscript.json` de este repo.

## 4. Crear las hojas con sus columnas

En el editor de Apps Script:
1. Seleccioná la función `initSheets` en el dropdown de arriba.
2. Click en **Ejecutar**. Te va a pedir autorizar permisos (Sheets + Drive) — autorizá con tu cuenta de Google.
3. Volvé a la planilla — vas a ver tres hojas creadas con headers: **PEDIDOS**, **PRODUCCION**, **STOCK**.

## 5. Ajustar el stock inicial

En la hoja **STOCK**, celda **B1**, poné el número de bolsas que tenés HOY en stock. Las celdas B2, B3 y B4 ya tienen fórmulas automáticas, no las toques.

## 6. Desplegar el Apps Script como Web App

1. En el editor de Apps Script: arriba a la derecha, **Implementar → Nueva implementación**.
2. Tipo: **Aplicación web** (engranaje al lado de "Seleccionar tipo").
3. Configurá:
   - **Descripción**: `v1`
   - **Ejecutar como**: Yo (tu mail)
   - **Quién tiene acceso**: **Cualquier persona**
4. **Implementar** → autorizá si te lo pide → te da una URL tipo `https://script.google.com/macros/s/AKfyc.../exec`. **Copiala**.

> ⚠ Cada vez que modifiques el código tenés que crear una **Nueva implementación** (o "Administrar implementaciones → editar → versión nueva"). La URL se mantiene si editás una existente.

## 7. Pegar las URLs en index.html

Abrí `index.html` y al principio del `<script>` editá:

```js
const CONFIG = {
  API_URL:   'https://script.google.com/macros/s/.../exec',
  SHEET_URL: 'https://docs.google.com/spreadsheets/d/.../edit',
  DRIVE_URL: 'https://drive.google.com/drive/folders/...',
  POLL_MS:   30000
};
```

## 8. Subir a GitHub Pages

1. Andá a https://github.com/new
2. Repository name: `pavimax`
3. Public, sin README, sin .gitignore, sin licencia (todo vacío).
4. Crear.
5. Desde esta carpeta (`C:\Users\Usuario\Downloads\pavimax\`):
   ```
   git init
   git add .
   git commit -m "init"
   git branch -M main
   git remote add origin https://github.com/marcoskatz-cmd/pavimax.git
   git push -u origin main
   ```
6. En el repo → **Settings → Pages** → Source: `Deploy from a branch`, Branch: `main` / `/ (root)`. Save.
7. En 1-3 minutos el panel queda en: **https://marcoskatz-cmd.github.io/pavimax/**

## 9. Probar

1. Cargá un pedido a mano en la planilla, hoja PEDIDOS:
   - id: `1` (o cualquier número único)
   - fecha_carga: hoy
   - cliente: `Cliente Prueba`
   - cantidad_bolsas: `10`
   - observacion_pedido: `bolsa blanca`
   - estado: `pendiente`
   - (resto en blanco)
2. Abrí el panel en el celular del operario y se tiene que ver el pedido.
3. Tocá ENTREGAR, sacá foto, confirmar → debería aparecer en ENTREGADOS y en la carpeta de Drive.
4. En la vista "Pedidos", tocá el botón 🔔 sonido una vez para activar las notificaciones del navegador.

## Uso diario

- **Oficina**: carga una fila por pedido en la planilla con estado=pendiente.
- **Operario**: ve los pedidos en el panel (botón **Pedidos**), los marca como entregados con foto + observación. La producción del día se carga desde el botón **Producción**.
- **Stock**: se recalcula solo. Lo ves desde la vista "Producción" del panel o abriendo la planilla directo.

## Cambios después

- Cambios al HTML: editar `index.html`, `git add . && git commit -m "..." && git push`. Pages se redeploya solo.
- Cambios al Apps Script: editar en el editor de Google, **Implementar → Administrar implementaciones → editar → Versión nueva → Implementar**. La URL no cambia.
