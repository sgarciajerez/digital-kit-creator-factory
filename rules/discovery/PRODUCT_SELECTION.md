# Selección de nuevos productos

La factoría recomienda temas a partir de oportunidades registradas y evidencia asociada. No interpreta ideas del modelo, anuncios parecidos ni volumen de resultados como prueba suficiente de demanda.

## Flujo de decisión

1. Registrar una hipótesis con público, problema y solución propuesta.
2. Importar observaciones con fecha, mercado, consulta, métrica, tipo de evidencia, limitaciones y referencia verificable.
3. Puntuar la oportunidad con criterios explicables y vincular a la puntuación los identificadores de evidencia usados.
4. Ejecutar `.venv\\Scripts\\python.exe -m scripts.research.product_selection` o la acción equivalente de la interfaz.
5. Revisar `outputs/reports/product_selection_report.json`: recomendadas, bloqueadas y motivos.
6. Aprobar manualmente el experimento y definir después el producto/kit concreto; la recomendación no publica ni genera por sí sola un kit.

## Filtros obligatorios

- Decisión `accept` y puntuación comercial mínima configurada.
- Tres observaciones recientes como mínimo (ventana inicial de 180 días), observadas y con referencia.
- Dos fuentes independientes como mínimo, calculadas por dominio de URL o nombre de fuente.
- Al menos una observación que represente demanda y otra de competencia/precio.
- Un candidato no puede superar el umbral de similitud léxica frente al catálogo generado/activo.
- En un mismo lote se recomienda como máximo una idea por categoría.

Las búsquedas devueltas, anuncios activos y recuentos de resultados se clasifican como oferta/competencia; no demuestran que compradores busquen o compren. Para la señal de demanda se necesitan métricas pertinentes, como búsquedas mensuales, ventas, conversiones, solicitudes o entrevistas, con sus límites declarados. Las estimaciones e hipótesis pueden informar una revisión, pero no satisfacen la puerta de evidencia observada.

## Ranking y aprendizaje

El ranking combina puntuación comercial (70 %), novedad respecto al catálogo (20 %) y resultados previos de la tienda por categoría (10 %). Si no hay métricas de tienda utilizables, el peso se redistribuye entre las señales disponibles; no se penaliza automáticamente una categoría nueva por carecer de historial.

Los resultados de tienda solo se usan cuando se enlazan a una categoría, no están marcados como datos de prueba y alcanzan los mínimos configurables de visitas y pedidos. Se informa la conversión observada relativa entre categorías disponibles; no se interpreta una muestra pequeña como fracaso.

La similitud usa Jaccard sobre términos normalizados de nombre, público, problema, solución y funciones del catálogo. Es un filtro explicable para detectar empaquetados repetidos, no un modelo semántico. Las alertas y los candidatos pospuestos por diversidad son visibles para revisión humana.

Los umbrales y pesos iniciales están en `config/product_selection.json`. Son hipótesis operativas y deben ajustarse con las métricas reales, no retocarse para justificar una idea preferida. Etsy SEO y comparación de precios siguen siendo pasos posteriores a elegir una oportunidad: no sustituyen la investigación de demanda.

## Límites actuales

- La evidencia entra por importación manual/CSV; el selector no obtiene automáticamente volúmenes de búsqueda de Etsy.
- La cartera conocida se deriva del catálogo de producto y carpetas generadas localmente. Listings externos aún no sincronizados requieren registrarse/importarse para que aparezcan en la comparación.
- El recomendador ordena oportunidades ya investigadas; no inventa señales de mercado ni autoacepta ideas.
- El autopiloto heredado solo sabe generar planificadores de presupuesto. Se bloquea si la oportunidad seleccionada no es compatible, hasta que exista un adaptador específico para su tipo de kit.
