# Cargas · Grupo Depor

Dashboard de órdenes de compra (OC) enviadas a tienda y su recepción reportada por las supervisoras **desde la misma página** (sin formulario de Google). Misma arquitectura que `levantamientos`: página estática en GitHub Pages, las OC en JSON dentro del repositorio subidas desde el navegador con un token; los reportes van a un Google Sheet a través de un Apps Script.

**Sitio publicado:** https://jmunozjm93-blip.github.io/cargas/

| Archivo | Descripción |
|---|---|
| `index.html` | Dashboard. Pestaña **Mis cargas** (por supervisora: pendientes / reportadas / con observación, botón *Reportar* con estado, comentario y fotos) y pestaña **Resumen** (solo lectura: avance por supervisora y por OC, reportes sin OC cargada) |
| `armar.html` | Arma el Excel de cargas desde el listado de OC (Marca · Departamento · N°DEOC · Familia · UND OC) + `Ordenes_Picking` + maestro de supervisoras + fotos (ver *Armar el Excel de cargas*) |
| `js/armar.js` · `js/worker-armar.js` | Cruce OC × picking × supervisoras y escritura del Excel (ExcelJS); el worker lee los Excel sin congelar la pantalla |
| `cargar.html` | App de carga: valida el Excel de cargas, lo convierte y lo sube a GitHub |
| `js/convertir.js` | Conversión Excel → JSON (hoja `Datos`) |
| `js/worker-convertir.js` | Web Worker que lee el Excel sin congelar la pantalla |
| `data/manifiesto.json` | Todas las OC cargadas con su resumen por tienda (lo que lee el dashboard) |
| `data/oc/<oc>.json` | Detalle de cada OC: modelos y unidades por tienda (se pide al tocar una tienda) |
| `data/imagenes.json` | Modelo → ID de foto en Google Drive (desde `Excel_Macro.xlsx`, hoja Imagenes) |
| `data/supervisores.json` | Maestro de supervisoras por tienda (desde `Supervisores_Consolidado.xlsx`, hoja Supervisores); manda sobre la columna `Supervisor` del Excel de cargas |
| `apps-script/Codigo.gs` | Buzón de reportes (Google Apps Script): guarda las fotos en Drive y cada reporte como fila del Sheet; la página lo lee de ahí |

## De dónde salen los datos

1. **OC enviadas**: el Excel de cargas (`Cargas_<Depto>_<Cliente>.xlsx`) que trae la hoja `Datos` con una fila por OC · modelo · talla · tienda (`NumAtCard · Cliente · modelo · Descripcion modelo · ShipToCode · Tienda · Supervisor · Quantity …`; la columna `Departamento` es opcional: si no viene se toma del nombre del archivo, ej. `Cargas_21.09_Paris-deporte.xlsx` → Deporte). Se sube por `cargar.html`.
2. **Recepción**: la supervisora toca la carga pendiente → **Reportar** → elige *Llegó completa / Llegó incompleta / No llegó*, escribe un comentario y saca fotos. La página achica las fotos (lado mayor 1280 px, JPG, ~150–250 KB) y manda todo al Apps Script, que guarda las fotos en la carpeta de Drive `Cargas - Fotos` y agrega una fila en la hoja `Reportes` del Sheet `Cargas - Reportes`. Al abrir la página se leen los reportes desde el mismo script (`?accion=reportes`).
3. **Supervisora de cada tienda**: `data/supervisores.json`, el maestro que sale de `Supervisores_Consolidado.xlsx` (columnas `Cadena · Cod · Tienda · Supervisor`). Se sube por `cargar.html` y **manda sobre lo que traiga el Excel de cargas**: corrige nombres distintos, `SIN ASIGNAR` y `RUTA NUEVO`, y se aplica también a las OC ya cargadas (no hay que volver a subirlas). Si una tienda no está en el maestro, se usa lo que diga su Excel.
4. **Fotos de los modelos**: `data/imagenes.json` de este mismo repositorio (`MODELO` → ID de foto en Google Drive). Se actualiza arrastrando `Excel_Macro.xlsx` (hoja `Imagenes`) en `cargar.html`.

## Cómo se cruzan

Cada fila del Excel es una **OC × tienda**. El reporte se hace desde esa misma fila, así que trae la OC y el código de tienda exactos: no hay nada que escribir a mano ni tiendas mal elegidas.

- Hay reporte → si fue *Llegó completa* la carga se **archiva sola** (sale de la lista "Activas" de la supervisora y queda en "Archivadas"); si fue *incompleta* o *no llegó* queda como **Con observación**, a la vista. Si hay varios reportes manda el último; todos se ven en el detalle. En el Resumen se ven todas.
- No hay reporte → **Pendiente**, con los días desde que se cargó la OC.
- Tiendas que quedan sin supervisora (no están en el maestro y su Excel dice `SIN ASIGNAR` / `RUTA NUEVO`) **no se suben ni se muestran**: nadie las reportaría. La app de carga avisa cuáles quedaron fuera.
- Reportes cuya OC ya no está cargada (se quitó) se listan en el Resumen como "Reportes sin OC cargada".

En el Resumen, la columna **Estado** es un semáforo: 🔴 *Sin reportar* (falta alguna tienda, con el conteo), 🟡 *Con observación* (todas reportaron pero alguna quedó incompleta o no llegó), 🟢 *Todo ok* y ⚪ *Sin asignar* (tiendas sin supervisora).

## Instalar el buzón de reportes (una vez)

Con la cuenta de Google donde quieres que quede el Sheet:

1. Entrar a [script.google.com](https://script.google.com) → **Nuevo proyecto** → pegar el contenido de `apps-script/Codigo.gs` (reemplazando lo que haya) → Guardar.
2. Arriba elegir la función **configurar** → **Ejecutar** → autorizar. Crea el Sheet `Cargas - Reportes` (hoja `Reportes`) y la carpeta de Drive `Cargas - Fotos`; el registro muestra los links.
3. **Implementar → Nueva implementación → Aplicación web**: *Ejecutar como* **Yo** · *Quién tiene acceso* **Cualquier usuario** → Implementar.
4. Copiar la **URL de la aplicación web** (termina en `/exec`) y pegarla en `index.html`:

```js
var API_URL='https://script.google.com/macros/s/…/exec';
```

Si se cambia el código del script hay que crear una **versión nueva** de la implementación para que la URL use el código nuevo. Mientras `API_URL` esté vacío la página avisa y no deja reportar.

Todo lo del sistema vive en la carpeta de Drive **Cargas - Grupo Depor** (el Sheet de reportes, la carpeta `Cargas - Fotos` y el propio script). La función `organizar()` del script la crea y mete todo dentro; se puede volver a ejecutar cuando sea.

Columnas del Sheet: `Fecha · OC · Cliente · Departamento · Cod tienda · Tienda · Supervisora · Estado (completa / incompleta / nollego) · Comentario · Fotos (IDs de Drive) · Unidades · Origen`.

## Armar el Excel de cargas

**https://jmunozjm93-blip.github.io/cargas/armar.html**

1. Pegar el listado de OC copiado desde Excel (`Marca · Departamento · N°DEOC · Familia · UND OC`, con o sin títulos; también sirve una OC por línea o arrastrar el Excel del listado).
2. Arrastrar el `Ordenes_Picking <dd.mm>.xlsx` del día. El maestro de supervisoras y el catálogo de fotos se toman de lo publicado (`data/supervisores.json`, `data/imagenes.json`); arrastrar `Supervisores_Consolidado.xlsx` o `Excel_Macro.xlsx` los reemplaza solo para esa sesión.
3. Revisar la tabla: por OC muestra cliente, tiendas, modelos, UND OC vs unidades del picking, estado en el picking y tiendas sin supervisora. Si una OC no está en el picking sugiere las parecidas (primero las que calzan con UND OC, ej. `232336` → `2323362`) con un botón para corregirla.
4. **Generar Excel** → descarga `Cargas_<dd.mm>_<Clientes>[-<depto>].xlsx` con hojas `Resumen`, `Datos` y una por OC (modelo × tienda, supervisora arriba, foto). Ese archivo se sube en `cargar.html`.

Reglas: cliente por `CardName` (CENCOSUD → Paris, COMERCIAL ECCSA → Ripley, Falabella, La Polar, Hites); se ignoran líneas `CANCELED = Y`; sin filtro por `StatusOrden`; si una OC viene con varios `DocNum`, por OC + tienda + ItemCode manda el más reciente. Las fotos se bajan de `lh3.googleusercontent.com` de a 3 con reintentos (Google responde 429 si se piden muchas seguidas); las que no bajan quedan con un link "ver foto".

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
