# Exploración automática de nichos

## Propósito y límites

El Explorador de nichos ayuda a comparar temas candidatos antes de invertir en un producto. Puede partir del catálogo inicial configurable de `config/market_discovery.json` o de consultas escritas por la persona usuaria. El catálogo es un conjunto de hipótesis editoriales para iniciar investigación, no una lista de tendencias ni una señal de demanda.

La factoría consulta la Etsy Open API para observar anuncios activos, muestras de productos digitales, títulos y precios. Estas mediciones describen oferta/competencia; no miden búsquedas, ventas, conversión, saturación real ni intención de compra. Los totales pueden incluir anuncios físicos y los resultados corresponden a una muestra limitada. No generar una puntuación de oportunidad a partir solo del recuento de anuncios.

## Flujo

1. Escoger mercado e idioma, y hasta ocho temas del catálogo base o consultas propias.
2. Consultar los anuncios activos de Etsy mediante la API oficial. No usar scraping.
3. Guardar informe fechado, consultas, fuente, precio mediano de la muestra, cantidad de anuncios observados y sus limitaciones. Clasificarlo como evidencia observada de oferta, nunca como evidencia de demanda.
4. Mostrar frases repetidas en títulos como vocabulario de competidores, no como palabras clave probadas.
5. Dejar todas las ideas en `needs_human_validation`. No aceptar oportunidades, generar productos ni publicar como efecto de este escaneo.
6. Antes de promover una hipótesis, definir comprador, problema y solución y reunir evidencia independiente de demanda/intención de compra. Aplicar después las puertas descritas en `PRODUCT_SELECTION.md`.

## Catálogo inicial y configuración

El catálogo en `config/market_discovery.json` ofrece veinticuatro temas de partida en inglés y español. La aplicación propone primero los que todavía no se han escaneado para ese idioma; una vez recorridos, se pueden repetir o añadir consultas propias. Se puede editar el catálogo sin cambiar el código; cada tema tiene un identificador, una etiqueta visible y una consulta para Etsy. No se debe describir el catálogo como descubrimiento de tendencias: es una fuente inicial de hipótesis que la factoría contrasta con datos de oferta.

El barrido admite hasta ocho consultas para controlar cuotas y coste operativo. Se excluyen los anuncios de la propia tienda cuando `ETSY_SHOP_ID` está configurado. El acceso Etsy se lee del `.env` local; no guardar secretos en informes ni en el repositorio.

## Marketplace Insights

Marketplace Insights queda fuera del flujo de la factoría: no hay automatización ni proceso de transcripción/carga manual. No se debe presentar como fuente de tendencia activa ni condicionar oportunidades, selección o desarrollo a datos introducidos manualmente. Si Etsy ofrece en el futuro una API o autorización escrita para acceder a esos datos, evaluar la integración por separado antes de incorporarla.

## Google Ads

La conexión OAuth de Google Ads quedó probada y se conserva en `scripts/research/authorize_google_ads.py` y `scripts/research/test_google_ads_connection.py` para diagnóstico de credenciales. No se usa para inferir demanda de Etsy. Keyword Planner es una fuente de estimaciones para campañas de Google Ads y su método requiere un nivel de acceso concreto; además, la política actual de Google dirige el alta de acceso de API a Google Cloud y recomienda dejar de enviar developer tokens heredados. No conectar esas estimaciones al selector de nichos ni presentarlas como búsquedas o ventas de Etsy.

Referencias oficiales: [política y acceso a Google Ads API](https://developers.google.com/google-ads/api/docs/api-policy/developer-token), [niveles de acceso](https://developers.google.com/google-ads/api/docs/api-policy/access-levels), [Keyword Plan Idea Service](https://developers.google.com/google-ads/api/docs/keyword-planning/generate-keyword-ideas) y [referencia Etsy Open API](https://developers.etsy.com/documentation/reference).

## Evolución

Para incorporar nuevas fuentes de demanda, usar una API/documentación oficial, verificar que su licencia y propósito permiten investigación comercial de productos, almacenar fecha/fuente/consulta/mercado/métrica/limitaciones y mantener las señales de oferta separadas de las de demanda. No automatizar scraping ni aprobación/publicación.
