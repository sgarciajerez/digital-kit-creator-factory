# Conectividad y publicación en Etsy

## 1. Registrar el callback OAuth

En Etsy Developer Portal, registra exactamente:

```text
http://localhost:8501/etsy/callback
```

Debe coincidir con `ETSY_REDIRECT_URI`.

## 2. Autorizar la tienda

Desde `D:\Proyectos\Creador de Excels`:

```powershell
.venv\Scripts\python.exe -m scripts.etsy.etsy_oauth
```

Autoriza los permisos `listings_r` y `listings_w`. El script valida `state` y
PKCE y guarda el token de renovación en `.env` sin imprimirlo.

## 3. Revisar el paquete

Ejecuta la factoría para el producto seleccionado:

```powershell
.venv\Scripts\python.exe -m scripts.run_autonomous_factory --product N --no-open
```

El flujo genera y valida el kit, prepara el paquete de listing y, como etapa
obligatoria, crea las capturas, mockups y demo del Excel para ese producto. Los
recursos se guardan en `outputs\listings\product-N\draft_package`; el manifiesto
`listing.json` apunta a todos ellos. No se usan rutas fijas a otro producto.

Antes de publicar, revisa los derechos, el texto comercial, el precio, la
portada y las vistas del libro. La generación de assets no sustituye esa revisión.

## 4. Crear y activar el listing

En la ejecución normal de la factoría, crear el listing, subir sus medios y
activarlo son etapas automáticas para productos elegibles. Los comandos
manuales de abajo permanecen sujetos a los flags globales y requieren
habilitación explícita en el entorno.

`.env` mantiene `ETSY_DRY_RUN=true` y `ETSY_PUBLISH_ENABLED=false` por defecto.
La factoría puede habilitar las escrituras temporalmente, solo dentro de sus
subprocesos finales y para los productos autorizados por `auto_publish_policy`
en `config/FACTORY_PIPELINE.json`. No publica productos antiguos: solo se
admiten excepciones puntuales o productos nuevos desde el ID configurado.

Primero se crea el listing como borrador y se carga el ZIP:

```powershell
.venv\Scripts\python.exe -m scripts.etsy.etsy_publish --manifest outputs\listings\product-N\draft_package\listing.json --execute
```

Después, `etsy_activate.py` carga la portada, las imágenes de marketing y el
vídeo declarados en el manifiesto antes de activar el anuncio. No lo ejecutes
hasta revisar el borrador y los recursos visuales en el paquete local:

```powershell
.venv\Scripts\python.exe -m scripts.etsy.etsy_activate --manifest outputs\listings\product-N\draft_package\listing.json --execute
```

La activación exige que las imágenes y el vídeo existan en el paquete y que el
token conserve `listings_w`. El manifiesto registra los IDs de los recursos
subidos para poder auditar la publicación.
