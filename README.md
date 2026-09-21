# Cargas · Grupo Depor

Dashboard de órdenes de compra (OC) enviadas a tienda y su recepción reportada por las supervisoras. Misma arquitectura que `levantamientos`: página estática en GitHub Pages, datos en JSON dentro del repositorio, subidos desde el navegador con un token.

**Sitio publicado:** https://jmunozjm93-blip.github.io/cargas/

| Archivo | Descripción |
|---|---|
| `index.html` | Dashboard. Pestaña **Mis cargas** (por supervisora: pendientes / reportadas / con observación, botón *Reportar*) y pestaña **Resumen** (solo lectura: avance por supervisora y por OC, respuestas sin cruzar) |
| `cargar.html` | App de carga: valida el Excel de cargas, lo convierte y lo sube a GitHub |
| `js/convertir.js` | Conversión Excel → JSON (hoja `Datos`) |
| `js/worker-convertir.js` | Web Worker que lee el Excel sin congelar la pantalla |
| `data/manifiesto.json` | Todas las OC cargadas con su resumen por tienda (lo que lee el dashboard) |
| `data/oc/<oc>.json` | Detalle de cada OC: modelos y unidades por tienda (se pide al tocar una tienda) |

## De dónde salen los datos

1. **OC enviadas**: el Excel de cargas (`Cargas_<Depto>_<Cliente>.xlsx`) que trae la hoja `Datos` con una fila por OC · modelo · talla · tienda (`NumAtCard · Cliente · Departamento · modelo · Descripcion modelo · ShipToCode · Tienda · Supervisor · Quantity …`). Se sube por `cargar.html`.
2. **Recepción**: las respuestas del formulario de Google caen en el Sheet *Consolidado Final* (pestaña de respuestas, `gid=0`). El dashboard las lee directo del Sheet cada vez que se abre (`gviz/tq?tqx=out:csv`), no hay que subir nada. El Sheet debe seguir con acceso "cualquiera con el enlace puede ver".
3. **Fotos de los modelos**: el catálogo `data/imagenes.json` del repositorio `levantamientos` (se actualiza desde allá con `Excel_Macro.xlsx`).

## Cómo se cruzan

Cada fila del Excel es una **OC × tienda**. Se busca una respuesta del formulario con ese mismo número de OC (se aceptan varios números en una celda) y el mismo **código de tienda** (los dígitos con que empieza el nombre: `10084-Copiapo` → `10084`).

- Hay respuesta → **Reportada**; si el comentario dice "no llegó", "incompleta", "falta", etc. → **Con observación**.
- No hay respuesta → **Pendiente**, con los días desde que se cargó la OC.
- Respuesta con la OC correcta pero una tienda que no está en esa OC: se asigna si la supervisora tiene una sola tienda pendiente en la OC (marcó mal la tienda), y se muestra la tienda que puso.
- Tiendas sin supervisor en el Excel (`SIN ASIGNAR`) se listan como **Sin supervisora** y no cuentan en el avance.
- Respuestas que no cruzan con ninguna OC cargada (desde 2 semanas antes de la primera OC) se listan en el Resumen con el motivo.

## Botón "Reportar"

Abre el formulario de Google con retail, tienda, categoría, marca y N° de OC ya rellenados; la supervisora solo escribe el comentario y sube fotos. Se configura en `index.html`:

```js
var FORM={ url:'https://docs.google.com/forms/d/e/…/viewform', campos:{ retail:'entry.…', tienda:'entry.…', categoria:'entry.…', marca:'entry.…', oc:'entry.…' } };
```

Los `entry.NNNN` se sacan del formulario: ⋮ → *Obtener enlace de relleno previo*, llenar cualquier valor en cada campo y copiar el link. Mientras `url` esté vacío el botón no aparece. Para que el prellenado funcione con la tienda, el valor de la lista del formulario debe ser exactamente el nombre que trae el Excel (`10084-Copiapo`).

## Rutina de carga

1. Abrir **https://jmunozjm93-blip.github.io/cargas/cargar.html**.
2. Arrastrar el Excel de cargas (uno o varios). Se muestran las OC encontradas con tiendas, unidades y tiendas sin supervisor.
3. Pulsar **Subir a GitHub**. En 1–2 minutos aparece en el dashboard como pendiente para cada supervisora.
4. Volver a subir un Excel con una OC ya cargada la reemplaza. **Quitar** (en la tabla de OC publicadas) la saca del dashboard.

Los Excel no se suben: se quedan en el PC (ignorados por git). La app necesita un token de GitHub fine-grained con *Contents: Read and write* sobre este repositorio; si ya existe el de `levantamientos`, basta con agregarle este repositorio.

## Estructura de `data/manifiesto.json`

```json
{ "ocs": { "2320562": { "oc": "2320562", "cliente": "Ripley", "depto": "Deporte", "marca": "Converse", "comentario": "Day one runner",
    "udsOC": 707, "uds": 573, "modelos": 3, "fuente": "Ordenes_Picking 27.08.xlsx", "archivo": "Cargas_Deporte_Ripley.xlsx", "subido": "2026-09-01T12:00:00Z",
    "sinSupervisor": [], "tiendas": [{ "cod": "10084", "nombre": "10084-Copiapo", "sup": "ERIKA ZUAZUA", "uds": 68, "modelos": 3 }] } },
  "actualizado": "2026-09-01T12:00:00Z" }
```
