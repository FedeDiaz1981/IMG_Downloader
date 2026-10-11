# Historial de desarrollo - IMG Downloader

Este archivo resume el trabajo realizado hasta el 10 de octubre de 2026 para poder retomar el proyecto sin depender del historial de chat.

## Objetivo original

Herramienta local en Flask para:

1. Importar un Excel de productos.
2. Agruparlos por marca.
3. Elegir las marcas a procesar y la web de cada fabricante.
4. Buscar, descargar y organizar imagenes de producto.
5. Preparar las imagenes para el catalogo de Pinto Fruta.

El proyecto se ejecuta desde esta carpeta y se abre en `http://127.0.0.1:<puerto>`.

## Funcionalidades incorporadas

### Importacion y descarga de imagenes

- Deteccion automatica de encabezados en Excel.
- Seleccion de columnas de marca, producto, presentacion adicional y SKU/codigo.
- Agrupacion por marca con URL de fabricante opcional.
- Casillas para marcar o desmarcar todas las marcas.
- Busqueda dentro del sitio del fabricante y busqueda complementaria en Google Images mediante Serper.
- Priorizacion de imagenes de mayor resolucion cuando las URLs del sitio exponen variantes.
- Validacion opcional con Gemini para descartar imagenes que no representan el producto.
- Reintento de productos que conservaron el archivo `SIN_IMAGEN.txt`.
- Enlaces de busqueda manual en Google Images para los productos que no obtuvieron imagen valida.

Las descargas se organizan bajo `downloads/<fecha>/<marca>/<producto>/`.

### Preparacion de imagen comercial

- Herramienta para recorrer carpetas de productos descargados.
- Seleccion de carpeta local de origen y carpeta destino con el Explorador de Windows.
- Generacion de `producto_comercial.png` a partir de una imagen encontrada, con fondo removido cuando es posible.
- El original no se modifica; la salida se crea en una subcarpeta nueva.

### Busqueda desde un listado de faltantes

- Pantalla para importar CSV o XLSX de productos pendientes.
- Seleccion de carpeta de salida.
- Busqueda, descarga y preparacion de imagenes en estructura apta para posterior carga.

### Comparacion y duplicados

- Normalizacion de texto que ignora acentos, guiones, espacios y puntuacion.
- Comparacion basada en marca, producto y presentacion, evitando que una descripcion extensa afecte el match.
- Generacion de CSV de productos duplicados a partir de listados locales.
- En Pinto Fruta: listado de duplicados y limpieza automatizada, conservando primero el producto con imagen y, en empate, el ID mas reciente.

### Carga en Pinto Fruta

- Conexion a Edge por CDP en `http://127.0.0.1:9222` para reutilizar una sesion autenticada.
- Boton `Abrir Pinto Fruta` que abre Edge con un perfil local en `.pintofruta-edge`.
- Busqueda de la ficha de cada producto por los datos derivados de la carpeta.
- Carga de imagen cuando falta y reemplazo cuando ya existe.
- Resolucion opcional de coincidencias ambiguas mediante Gemini.
- Pantalla manual de eleccion cuando hay varias fichas candidatas.
- Listado de productos sin imagen.

## Integracion de operaciones administrativas de Pinto Fruta

Antes de esta integracion se usaron scripts puntuales. Sus capacidades relevantes quedaron incorporadas a la interfaz de IMG Downloader.

### Auditoria de Instagram

Panel: `Instagram y disponibilidad`.

1. Lee publicaciones de `@distribuidora_pintofruta` usando la sesion abierta de Edge.
2. Captura la publicacion y combina su imagen con su texto.
3. Gemini identifica productos visibles.
4. Compara los productos detectados contra el catalogo de Pinto Fruta.
5. Genera el CSV `catalogo-no-detectado-instagram-<fecha>.csv`.

La auditoria requiere:

- Sesion de Pinto Fruta abierta en Edge.
- `GEMINI_API_KEY` configurada en el entorno de Windows.

La deteccion por publicaciones es una ayuda operativa, no una garantia de que un producto no se venda: las publicaciones antiguas, historias y productos que no aparecen claramente en una imagen pueden no ser detectados.

### Inactivacion de productos desde el CSV

Panel: `Instagram y disponibilidad`.

- Recibe un CSV de auditoria con la columna `ID Pinto Fruta`.
- Busca el producto en el administrador y desmarca `ACTIVO`.
- No borra registros ni imagenes.
- Genera `productos-inactivados-instagram-<fecha>.csv` con el resultado por producto.

Durante el trabajo se inactivaron los productos detectados como ausentes de Instagram. El reporte historico queda en `jobs/`.

### Marcas sin productos activos

Panel: `Marcas sin productos activos`.

Flujo:

1. Seleccionar el archivo `.env.local` del proyecto de Pinto Fruta.
2. Usar `Vista previa y CSV` para revisar las marcas candidatas.
3. Usar `Inactivar marcas vacias` para desactivar las marcas que no tengan productos activos.

La operacion:

- Lee solo `DATABASE_URL` del archivo seleccionado durante la ejecucion.
- No guarda la cadena de conexion ni la incluye en reportes.
- Consulta `brands` y los `products` cuyo estado no es `inactive`.
- Normaliza nombres de marca antes de comparar.
- Establece `active = false` y `featured = false` para las marcas sin productos activos.
- No borra marcas ni productos.

Verificacion realizada el 10/10/2026:

- Se inactivaron 36 marcas sin productos activos.
- Luego de la limpieza quedaron 31 marcas con productos activos.
- Una vista previa posterior devolvio 0 marcas activas sin productos activos.

## Reportes persistentes

Los resultados no dependen de los trabajos activos en memoria:

- Todos los CSV relevantes se guardan en `jobs/`.
- El panel `Reportes guardados` lista los ultimos 15 reportes de Pinto Fruta tras reiniciar la aplicacion.
- Las tareas activas y sus logs se pierden al reiniciar Flask, pero sus salidas y reportes se conservan en disco.

Prefijos principales:

- `pintofruta-sin-imagen-`
- `pintofruta-productos-duplicados-`
- `pintofruta-duplicados-eliminados-`
- `catalogo-no-detectado-instagram-`
- `productos-inactivados-instagram-`
- `pintofruta-marcas-sin-productos-activos-`
- `pintofruta-marcas-inactivadas-`

## Configuracion y dependencias

Dependencias principales en `requirements.txt`:

- Flask
- Playwright
- openpyxl
- Pillow
- rembg
- requests
- `psycopg[binary]` para el proceso de marcas sin productos activos

Variables de entorno usadas:

- `SERPER_API_KEY`: busqueda de Google Images.
- `GEMINI_API_KEY`: validacion de imagenes, desempates y auditoria de Instagram.
- `PINFOFRUTA_ENV_FILE` opcional: ruta por defecto al `.env.local` de Pinto Fruta.
- `PORT` opcional: puerto de Flask; por defecto `5000`.

No agregar claves, cadenas de conexion ni tokens a este repositorio ni a este archivo.

## Ejecucion

```powershell
py -m pip install -r requirements.txt
py app.py
```

Si `5000` ya esta ocupado, iniciar una instancia en otro puerto:

```powershell
$env:PORT = '5002'
py app.py
```

La ultima version verificada durante esta etapa se inicio en `http://127.0.0.1:5002/` porque una instancia anterior continuaba usando el puerto `5000`.

## Estado tecnico conocido

### Correccion pendiente de desplegar en Pinto Fruta

Se detecto un defecto en el backend del proyecto Pinto Fruta: la accion administrativa de guardar marcas recibia el valor `active`, pero no lo incluia en el `insert/upsert` de `saveBrand`.

Se corrigio localmente el archivo del proyecto Pinto Fruta:

`E:\FD\WBI\maxi\Proyectos\en curso\Pintofruta\PintoFruta_Store\src\app\admin\actions.ts`

La correccion hace que `active` se lea, se conserve cuando no se envia y se guarde al crear o actualizar una marca. Esa modificacion pertenece al proyecto Pinto Fruta y necesita su proceso normal de despliegue para reflejarse en el administrador web.

La herramienta IMG Downloader evita ese problema para la limpieza de marcas porque realiza la actualizacion directamente en la base de datos, con una vista previa disponible antes de modificar.

### Scripts historicos

Permanecen en esta carpeta como referencia de las operaciones iniciales:

- `instagram_catalog_audit.py`
- `deactivate_instagram_missing.py`
- `deactivate_empty_brands.py`

No son necesarios para el flujo habitual: sus operaciones estan disponibles desde la interfaz de IMG Downloader. Antes de eliminar estos scripts conviene conservar una copia versionada o confirmar que no se usan en otra automatizacion.

## Recomendaciones para retomar

1. Abrir esta bitacora y `README.md` antes de modificar los flujos de Pinto Fruta.
2. Ejecutar una vista previa antes de cualquier inactivacion masiva.
3. Descargar el CSV de cada operacion y conservarlo como respaldo.
4. Para cargas masivas, abrir Edge desde la aplicacion e iniciar sesion antes de iniciar el proceso.
5. Revisar los productos ambiguos desde la pantalla de eleccion manual en lugar de forzar coincidencias por texto.
6. Si cambia la estructura visual de `/admin`, ajustar los selectores de Playwright en `app.py`.
