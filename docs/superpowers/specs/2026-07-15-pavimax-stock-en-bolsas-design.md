# PAVIMAX — Stock en bolsas (dos stocks separados)

Fecha: 2026-07-15

## Problema

El STOCK hoy se lleva en **kg** (un solo número que combina bolsas×25 + bigbag×1000).
Marcos quiere ver el stock en **unidades físicas**: bolsas de 25kg y big bags de
1000kg como **dos stocks separados**, no en kilos.

Fuera de alcance: GANANCIAS (sigue en kg, es otra métrica) y CAPACIDAD (ya está en
bolsas/día). La emulsión sigue sin stock.

## Modelo nuevo

Dos productos stockeables, cada uno con su propio inicial / producido / vendido /
actual, expresados en **unidades** (bolsas y big bags), sin conversión a kg.

- `actual = inicial + producido − vendido` (por producto).
- Producido y vendido salen de contar unidades por producto (helpers ya existentes).

## Cambios

### 1. Hoja STOCK → tabla por producto

Deja de ser un dashboard en kg (`Concepto | Kg` con 4 filas) y pasa a una tabla:

| Producto | Inicial | Producido | Vendido | Actual |
|---|---|---|---|---|
| Bolsas 25kg | *(editable, default 0)* | auto | auto | auto |
| Big bag 1000kg | *(editable, default 0)* | auto | auto | auto |

- Encabezado fila 1: `Producto | Inicial | Producido | Vendido | Actual`.
- Fila 2: bolsas_25kg. Fila 3: bigbag_1000kg.
- Columnas `Inicial` (B) editables por Marcos; `Producido/Vendido/Actual` (C/D/E) las
  reescribe `getStock()` en cada lectura.
- Formato numérico sin sufijo (unidades), no `" kg"`.
- `CONFIG.stock_unidad = 'unidades'`.

### 2. Backend `getStock()`

Devuelve por producto en vez de un único número:

```js
{
  bolsas_25kg:   { inicial, producido, vendido, actual },
  bigbag_1000kg: { inicial, producido, vendido, actual }
}
```

- `inicial`: se lee de la columna B (bolsas B2, bigbag B3).
- `producido`: de `sumProducidoPorEnvase_()` → `{ bolsas_25kg, bigbag_1000kg }`.
- `vendido`: de `sumVendidoPorProducto_()` (unidades entregadas por producto).
- `actual = inicial + producido − vendido`.
- Reescribe C/D/E de las filas 2 y 3 con producido/vendido/actual.

Se dejan de usar (para stock) `sumProducido_()` y `sumVendidoKg_()` — siguen
existiendo porque GANANCIAS los usa.

### 3. Frontend (app operario, `index.html`)

La caja única "Stock" pasa a **dos cajas** siempre visibles:

- **Bolsas 25kg**: Inicial / Producido / Vendido / **Actual** (total resaltado).
- **Big bag 1000kg**: mismas 4 filas.

Números en unidades, sin sufijo " kg". `renderStock()` lee la nueva estructura
`state.stock.bolsas_25kg.*` y `state.stock.bigbag_1000kg.*`.

### 4. Migración en `initSheets()`

- Reconstruir el layout de STOCK a la tabla nueva (idempotente).
- **Resetear los iniciales a 0**: el inicial viejo estaba en kg y no se puede
  repartir limpio entre bolsas y big bag. Marcos carga a mano los valores reales en
  B2/B3 después del deploy.
- Setear `CONFIG.stock_unidad = 'unidades'` (para no reactivar la vieja migración
  bolsas→kg).
- Ajustar el formato visual de STOCK (`formatDashboard_` → tabla / `formatList_`),
  sin sufijo kg.

## Riesgos / notas

- **Ventana de corte**: al redeployar backend + migrar planilla, el frontend viejo
  queda incompatible unos minutos (ver memoria pavimax). Hacerlo tranquilo, Ctrl+Shift+R.
- `initSheets()` se corre a mano en el editor (pide autorización). Idempotente.
- El reset de iniciales borra el stock inicial actual — avisado y aceptado.
