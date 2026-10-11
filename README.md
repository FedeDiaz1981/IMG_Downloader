# IMG Downloader

Herramienta local para importar una lista de productos desde Excel, agruparlos por marca y descargar imagenes desde las paginas web de los fabricantes.

## Uso rapido

1. Instalar dependencias:

   ```powershell
   py -m pip install -r requirements.txt
   ```

2. Ejecutar la app:

   ```powershell
   py app.py
   ```

3. Abrir `http://127.0.0.1:5000`.

Para usar el boton de validacion con IA, configura una clave de Gemini solo en el entorno local:

```powershell
[Environment]::SetEnvironmentVariable('GEMINI_API_KEY', 'tu_clave', 'User')
```

Para los botones de Google Images y validacion con IA, crea una clave en https://serper.dev/ y pegala en el formulario. Tambien puedes guardarla en Windows:

```powershell
[Environment]::SetEnvironmentVariable('SERPER_API_KEY', 'tu_clave', 'User')
```

## Flujo

1. Subi el Excel.
2. Elegi las columnas de marca, producto y, opcionalmente, codigo/SKU.
3. Carga la web oficial para cada marca.
4. Ejecuta la descarga o usa "Buscar y validar con IA" para que Gemini descarte imagenes de otros productos.

Las imagenes se guardan en `downloads/<fecha>/<marca>/<producto>/` y tambien se genera un `.zip` para descargar todo junto.

## Pinto Fruta

La pantalla principal tambien concentra las tareas de catalogo de Pinto Fruta:

- carga y reemplazo de imagenes mediante la sesion abierta en Edge;
- listado y eliminacion de productos duplicados;
- auditoria del catalogo contra las publicaciones de Instagram con Gemini;
- inactivacion de productos desde el CSV de esa auditoria, sin borrarlos;
- vista previa e inactivacion de marcas que no tengan productos activos.

Para la limpieza de marcas, selecciona el archivo `.env.local` del proyecto de Pinto Fruta. La app lee solamente `DATABASE_URL` en el momento de ejecutar: no guarda credenciales en sus reportes. Los CSV quedan en `jobs/` y se muestran de nuevo al reiniciar la app.

## Notas

- La busqueda usa textos de la pagina, enlaces, nombres de archivo, atributos `alt/title` y metadatos como `og:image`.
- Algunos sitios bloquean scraping o cargan imagenes con JavaScript; en esos casos puede hacer falta adaptar el conector para esa marca.
- Conviene usar la pagina del catalogo o buscador del fabricante, no solo la home, cuando sea posible.
