# Conectividad y publicación en Etsy

## 1. Registrar el callback OAuth

En Etsy Developer Portal, abre la aplicación cuyas credenciales están en `.env` y registra exactamente:

```text
http://localhost:8501/etsy/callback
```

Debe coincidir con `ETSY_REDIRECT_URI`.

## 2. Autorizar la tienda

Desde `D:\Proyectos\Creador de Excels` ejecuta:

```powershell
.venv\Scripts\python.exe etsy_oauth.py
```

Se abrirá Etsy en el navegador. Autoriza los permisos `listings_r` y `listings_w`. El script valida `state` y PKCE y guarda el `ETSY_REFRESH_TOKEN` en `.env` sin imprimirlo.

## 3. Comprobar la taxonomía

La taxonomía se puede recalcular con:

```powershell
.venv\Scripts\python.exe etsy_taxonomy.py "budget planner spreadsheet" --write-env
```

El script consulta el árbol oficial de taxonomía y solo guarda la coincidencia si hay una diferencia clara frente al segundo candidato.

## 4. Revisar el paquete

Antes de una escritura real hay que confirmar:

- Que el texto y el archivo son propios o tienen licencia de uso.
- Que la descripción y las etiquetas no contienen afirmaciones engañosas.
- Que el precio y la moneda son correctos.
- Que el paquete contiene una imagen de portada para poder activar el listing.

El paquete actual más reciente está en `outputs\listings\product-7`. Está bloqueado hasta completar la revisión de política y derechos.

## 5. Crear un borrador en Etsy

La aplicación sigue protegida por defecto. Tras completar OAuth y la revisión, cambia temporalmente en `.env`:

```text
ETSY_DRY_RUN=false
ETSY_PUBLISH_ENABLED=true
```

Luego crea el listing como borrador y carga el XLSX:

```powershell
.venv\Scripts\python.exe scripts\\etsy_publish.py --manifest outputs\listings\product-7\listing.json --execute
```

La activación pública es una acción separada y requiere una imagen de portada:

```powershell
.venv\Scripts\python.exe scripts\\etsy_publish.py --manifest outputs\listings\product-7\listing.json --execute --activate
```

No uses `--activate` hasta revisar el borrador en Etsy.

