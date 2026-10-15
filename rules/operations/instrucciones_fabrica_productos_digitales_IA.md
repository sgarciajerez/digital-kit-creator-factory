# Especificación maestra: Fábrica autónoma de productos digitales con IA

**Versión:** 1.0  
**Entorno inicial:** Windows, ejecución local y supervisada  
**Presupuesto operativo inicial:** 10–20 € al mes para IA y experimentos, sin contar comisiones de plataformas  
**Objetivo:** construir un sistema integrado que investigue mercados, detecte oportunidades, decida qué merece la pena producir, genere y valide productos digitales, prepare/publica anuncios en plataformas permitidas y aprenda de visitas, ventas, costes y resultados.

---

## 0. Instrucciones para la IA que vaya a ejecutar este proyecto

Actúa como **ingeniero de software principal, arquitecto de automatización, analista de mercado y responsable de experimentación comercial**. Tu misión es construir el sistema de forma incremental, verificable, económica y segura.

No te limites a dar recomendaciones o pseudocódigo: inspecciona el entorno disponible, crea la estructura del proyecto, implementa funcionalidades reales, ejecuta pruebas, corrige errores y documenta cómo usarlo. Si una acción requiere credenciales, aprobación humana o acceso que no tienes, implementa una interfaz segura y explica exactamente qué necesita el usuario.

### Reglas innegociables

1. **No inventes datos de mercado.** Distingue datos observados, estimaciones, hipótesis y datos ausentes. Guarda la fuente, fecha, método y nivel de confianza de cada señal.
2. **No confundas generar con vender.** Optimiza por utilidad, demanda validada y beneficio neto, no por cantidad de archivos producidos.
3. **Mantén al usuario al mando al principio.** La investigación, clasificación, generación y análisis pueden automatizarse; la publicación debe requerir aprobación humana hasta que se hayan probado las integraciones y controles.
4. **Respeta las condiciones de cada plataforma.** Usa API oficiales e integraciones autorizadas. No hagas scraping prohibido, no eludas límites, CAPTCHA, controles de acceso ni medidas anti-bots. Si no existe una vía autorizada, prepara un borrador para publicación manual.
5. **No publiques contenido en masa sin control de calidad.** Cada producto debe superar pruebas automáticas y, al principio, revisión humana.
6. **Respeta derechos de autor y licencias.** No copies productos de competidores, marcas, personajes, diseños protegidos, textos o imágenes. No uses recursos sin licencia comercial compatible.
7. **No uses reseñas, ventas, visitas ni testimonios inventados.** No generes escasez falsa ni afirmaciones engañosas.
8. **Protege credenciales.** No guardes claves API en código, logs, repositorio ni base de datos sin protección. Usa variables de entorno o un almacén local de secretos.
9. **Controla el gasto.** Establece límites de gasto por día, mes, lote y producto. Detén las llamadas si se alcanza el presupuesto.
10. **Registra las decisiones.** Toda oportunidad aceptada, descartada, generada, publicada o pausada debe tener motivo, puntuación, fecha y evidencia disponible.
11. **No afirmes que algo está terminado sin ejecutarlo.** Informa de pruebas realizadas, resultados, limitaciones y próximos pasos.
12. **No automatices acciones irreversibles o con coste** (publicar, comprar anuncios, pagar servicios, borrar anuncios) sin una autorización explícita hasta que el usuario active conscientemente una política de autonomía.

---

## 1. Visión y objetivo del producto

Crear una aplicación local que coordine un ciclo de negocio completo:

1. Investigar posibles nichos y necesidades.
2. Proponer categorías y microcategorías.
3. Recopilar señales de demanda y competencia de fuentes autorizadas.
4. Puntuar las oportunidades con criterios explícitos.
5. Aprobar, rechazar o enviar a revisión las ideas según umbrales configurables.
6. Diseñar el producto digital más apropiado para cada oportunidad.
7. Generar los archivos, documentación, imágenes y ficha comercial.
8. Validar la calidad técnica, funcional, visual y legal.
9. Crear borradores y, cuando esté autorizado y sea técnicamente posible, publicar anuncios.
10. Recopilar impresiones, visitas, favoritos, conversiones, ventas, devoluciones, comisiones y costes.
11. Estimar qué está funcionando, evitar conclusiones prematuras y proponer nuevos experimentos.
12. Repetir el ciclo con aprendizaje documentado.

### Principio central

El sistema debe comportarse como un **laboratorio comercial autónomo con supervisión**, no como una máquina que produce miles de productos indiscriminadamente.

El objetivo es descubrir productos que:
- resuelvan un problema concreto;
- sean originales y útiles;
- puedan producirse y entregarse con coste bajo;
- tengan una oportunidad de mercado razonable;
- logren visitas y compras reales;
- dejen un margen positivo después de costes, comisiones, devoluciones e impuestos aplicables.

---

## 2. Contexto técnico y restricciones iniciales

- Sistema operativo: Windows.
- Equipo de inicio: PC personal, aproximadamente 16 GB de RAM.
- Ejecución: local; el sistema trabaja mientras el PC está encendido y el usuario está presente.
- Presupuesto adicional de IA: objetivo inicial de 10–20 €/mes.
- No se requiere modelo local ni GPU dedicada.
- Se permite usar API de modelos alojados, seleccionando modelos económicos para tareas rutinarias y modelos más capaces solo para decisiones importantes.
- No contratar servicios recurrentes ni publicidad sin aprobación.
- No se necesita desplegar en la nube en la primera versión.
- Diseñar la arquitectura para que pueda migrarse a un servidor más adelante sin reescribir el núcleo.
- Idioma inicial de productos y anuncios: español. La arquitectura debe permitir añadir otros idiomas posteriormente.
- La primera familia de productos recomendada son **hojas de cálculo, calculadoras y plantillas funcionales para un público específico**, porque pueden validarse con pruebas automáticas y no requieren generar imágenes en cada iteración.

---

## 3. Qué entra en el MVP y qué queda para más adelante

### MVP obligatorio

1. Aplicación local con interfaz.
2. Configuración de proveedor de IA y límites de gasto.
3. Base de datos local.
4. Registro de ideas, fuentes, observaciones y puntuaciones.
5. Flujo de investigación semiautomático.
6. Motor explicable de decisión: aceptar, rechazar o revisar.
7. Generación de hojas de cálculo/plantillas con contenido estructurado.
8. Validación automática de los archivos generados.
9. Generación de títulos, descripciones, etiquetas, instrucciones y archivos de vista previa.
10. Panel de aprobación humana.
11. Exportación de un paquete listo para subir a un marketplace.
12. Registro manual o importación inicial de métricas de visitas y ventas.
13. Informe de experimentos y recomendaciones.
14. Pruebas automatizadas y documentación.

### Segunda fase

- Integración con API oficial de una plataforma elegida.
- Creación de anuncios en modo borrador.
- Sincronización de métricas mediante API o importación CSV.
- Comparación de variantes de título, precio, imágenes y producto.
- Generación de PDF, documentos y otros formatos.
- Automatización de lotes más grandes, manteniendo límites.

### Fases posteriores

- Varias plataformas.
- Más familias de productos.
- Localización a varios idiomas.
- Programación de tareas y trabajos de larga duración.
- Ejecución en nube, solo si el negocio lo justifica.
- Modelos predictivos propios cuando haya suficientes datos reales.
- Automatización ampliada de publicación tras aprobación explícita y periodo de funcionamiento estable.

**No intentes implementar todo a la vez.** Cada fase debe ser utilizable por sí misma.

---

## 4. Arquitectura propuesta

Mantén una arquitectura modular y sencilla. No uses microservicios en el MVP.

### Tecnologías sugeridas

- **Python 3.12 o versión estable compatible**.
- **Streamlit** para una interfaz local rápida; si aparece una limitación real, valorar FastAPI + interfaz web ligera.
- **SQLite** para persistencia local.
- **Pydantic** para validar entradas y salidas estructuradas.
- **OpenAI API u otro proveedor configurable** para investigación, clasificación, generación y crítica.
- **OpenPyXL** para generar y comprobar archivos XLSX.
- **pytest** para pruebas.
- **Pandas** para análisis de métricas cuando aporte valor.
- **Jinja2** para plantillas de documentos y fichas.
- **Pillow** únicamente si se necesita procesamiento básico de imágenes.
- **Git** para control de versiones.
- **python-dotenv** para variables de entorno locales.
- **Logging** estándar de Python, con rotación y sin secretos.
- **httpx/requests** solo para integraciones autorizadas.
- **Tenacity** o una estrategia equivalente para reintentos limitados con backoff, sin repetir acciones de publicación de forma insegura.

Antes de instalar paquetes, comprueba Python, Git y el entorno actual. No reinstales herramientas existentes sin necesidad. Fija versiones compatibles en `requirements.txt` y, preferiblemente, genera un archivo de bloqueo.

### Módulos lógicos

1. `research`: fuentes, consultas, recopilación y normalización.
2. `opportunities`: agrupación de necesidades y cálculo de puntuaciones.
3. `decision_engine`: reglas, umbrales y explicación de decisiones.
4. `product_planning`: brief de producto y especificaciones.
5. `generation`: contenido y archivos.
6. `quality`: pruebas funcionales, contenido, duplicados y controles de riesgo.
7. `listing`: títulos, descripciones, etiquetas, precio y paquete de publicación.
8. `platforms`: adaptadores separados para cada plataforma.
9. `analytics`: métricas, embudos, márgenes y cohortes.
10. `learning`: hipótesis, experimentos y propuestas de actualización.
11. `budget`: seguimiento del consumo y límites de gasto.
12. `approval`: aprobaciones humanas y registro de acciones.
13. `ui`: interfaz de control.
14. `scheduler`: trabajos bajo demanda y tareas periódicas locales.
15. `audit`: eventos, decisiones y trazabilidad.

### Estructura de carpetas objetivo

```text
digital-product-factory/
├── app.py
├── README.md
├── requirements.txt
├── .env.example
├── .gitignore
├── config/
│   ├── settings.yaml
│   ├── scoring.yaml
│   ├── product_types.yaml
│   └── platform_policies.yaml
├── src/
│   └── factory/
│       ├── research/
│       ├── opportunities/
│       ├── decision_engine/
│       ├── product_planning/
│       ├── generation/
│       ├── quality/
│       ├── listing/
│       ├── platforms/
│       ├── analytics/
│       ├── learning/
│       ├── budget/
│       ├── approval/
│       ├── scheduler/
│       ├── audit/
│       └── common/
├── templates/
│   ├── spreadsheet/
│   ├── listings/
│   └── reports/
├── data/
│   ├── .gitkeep
│   ├── imports/
│   └── exports/
├── outputs/
│   ├── products/
│   ├── previews/
│   ├── listings/
│   └── reports/
├── logs/
│   └── .gitkeep
├── tests/
│   ├── unit/
│   ├── integration/
│   └── fixtures/
└── scripts/
    ├── setup_windows.ps1
    ├── run_app.ps1
    ├── backup.ps1
    └── import_metrics.py
```

La base de datos, archivos generados, registros y copias de seguridad deben estar fuera del repositorio o excluidos por `.gitignore`. Nunca versionar `.env`, claves ni tokens.

---

## 5. Interfaz de usuario

La aplicación debe abrirse en el navegador local y mostrar estas secciones:

### Panel principal
- gasto de IA este mes;
- presupuesto restante;
- ideas investigadas;
- oportunidades aceptadas/rechazadas/en revisión;
- productos generados y aprobados;
- anuncios en borrador y publicados;
- visitas, ventas e ingresos observados;
- avisos de calidad, integración o presupuesto;
- tareas en curso, fallidas y pendientes.

### Investigación
- iniciar una investigación;
- seleccionar o excluir fuentes;
- ver las fuentes y fechas;
- consultar términos, competencia y evidencias;
- filtrar por país, idioma, público y familia de producto;
- diferenciar evidencia real de hipótesis generadas por IA.

### Oportunidades
- tabla de oportunidades ordenable;
- puntuación total y puntuaciones parciales;
- nivel de confianza y calidad de datos;
- recomendación: aceptar, rechazar o revisar;
- explicación legible de los factores que influyeron;
- posibilidad de que el usuario anule la decisión y deje un motivo.

### Producción
- cola de productos;
- brief, público objetivo y problema resuelto;
- progreso por etapas;
- coste estimado y coste real;
- archivos y previsualizaciones;
- errores y resultados de pruebas;
- botones de regenerar, aprobar o descartar.

### Publicación
- plataforma de destino;
- vista previa de título, descripción, etiquetas, precio e imágenes;
- validaciones de política y derechos;
- estado: local, listo, pendiente de aprobación, borrador remoto, publicado, pausado o error;
- aprobación explícita antes de publicar.

### Analítica
- impresiones/visitas según las métricas disponibles;
- clics, favoritos, ventas y conversión cuando puedan medirse;
- facturación bruta, comisiones, devoluciones, coste de IA y margen estimado;
- resultados por categoría, producto, fecha y variante;
- nivel de confianza de cada conclusión.

### Configuración
- proveedor/modelo de IA;
- presupuesto mensual y límites por trabajo;
- países, idiomas y familias de productos;
- umbrales del motor de decisión;
- fuentes habilitadas;
- plataforma seleccionada;
- modo de autonomía;
- copia de seguridad y exportación de datos.

---

## 6. Investigación de mercados y detección de oportunidades

El sistema debe investigar oportunidades con **fuentes lícitas, documentadas y técnicamente accesibles**.

### Fuentes posibles

Priorizar, en este orden:

1. APIs oficiales de marketplaces y herramientas de palabras clave.
2. Tendencias y herramientas públicas de búsqueda.
3. Catálogos y páginas públicas accesibles respetando condiciones de uso.
4. Reseñas y preguntas públicas, si se pueden consultar legítimamente.
5. Datos que el usuario importe manualmente, incluidos CSV de visitas/ventas.
6. Fuentes de pago solo con aprobación y dentro del presupuesto.

No des por hecho que una plataforma ofrece acceso a todas sus métricas. Investiga la documentación vigente antes de implementar una integración. No automatices scraping prohibido.

### Datos que debe guardar cada señal

- nombre y tipo de fuente;
- URL o identificador;
- fecha/hora de consulta;
- consulta utilizada;
- país e idioma;
- observación exacta o resumen fiel;
- dato cuantitativo y unidad, si existe;
- método de obtención;
- limitaciones;
- fiabilidad estimada;
- si se trata de un dato observado, una estimación o una hipótesis.

### Generación de ideas

La IA puede proponer ideas nuevas a partir de:
- necesidades recurrentes;
- quejas o problemas sin resolver;
- tareas repetitivas;
- búsquedas específicas;
- productos que ahorren dinero o tiempo;
- públicos profesionales definidos;
- huecos observables en la oferta;
- productos existentes que puedan mejorarse sin copiarlos.

Cada idea debe expresarse como una hipótesis comprobable:

> “El público X necesita resolver Y. Las soluciones actuales presentan el problema Z. Un producto digital con la funcionalidad A podría aportar valor y venderse a un precio B. Evidencia disponible: […]. Incertidumbres: […]. Prueba más barata para validarlo: […].”

### Evitar falsas señales

- Muchos anuncios competidores no demuestran que exista demanda rentable.
- Un producto con muchas reseñas puede tener ventas históricas, pero no prueba la demanda actual.
- El número de resultados de búsqueda no equivale al volumen de búsquedas.
- Las cifras de ventas inferidas por terceros deben etiquetarse como estimaciones.
- La IA no puede inventar cifras de volumen, conversión, ventas ni precios observados.
- Una categoría sin datos suficientes debe quedar en “revisión”, no aprobarse por intuición.
- Guarda cuándo se actualizó cada señal y evita mezclar datos de periodos distintos sin advertirlo.

---

## 7. Motor de puntuación y decisiones

El sistema debe calcular una puntuación explicable. No debe dejar que una respuesta libre del modelo sea la única base para invertir dinero o publicar.

### Puntuación inicial sugerida

Puntúa cada criterio de 0 a 5. El valor ponderado se calcula así:

`puntuación_total = suma((puntuación_criterio / 5) * peso_criterio)`

| Criterio | Peso inicial |
|---|---:|
| Evidencia de demanda | 25 % |
| Intensidad del problema y utilidad | 15 % |
| Competencia y diferenciación posible | 15 % |
| Viabilidad de producir con calidad | 15 % |
| Margen potencial tras costes y comisiones | 15 % |
| Facilidad de llegar al público objetivo | 10 % |
| Riesgo legal, de plataforma y mantenimiento (puntuación alta = riesgo bajo) | 5 % |
| **Total** | **100 %** |

Los pesos son una hipótesis inicial configurable, no una verdad estadística. Deben revisarse a partir de los resultados reales y no cambiarse para justificar una idea preferida.

### Umbrales iniciales

- **75–100:** candidata a aprobación para producir un prototipo, si la calidad de evidencia es suficiente.
- **55–74:** revisar o realizar una investigación adicional de bajo coste.
- **0–54:** descartar temporalmente o archivar.
- **Bloqueo obligatorio:** riesgo legal, falta de derechos, incumplimiento de políticas, imposibilidad de validar una afirmación importante o coste fuera del presupuesto.

Requisitos adicionales:
- Si la evidencia de demanda es débil, la puntuación alta no permite publicación automática.
- La confianza de datos debe mostrarse separada de la puntuación comercial.
- El usuario puede ajustar umbrales desde configuración.
- Cada decisión debe explicar qué factores aumentaron o redujeron la puntuación.
- Los descartes se guardan para no repetir constantemente la misma idea sin evidencia nueva.
- No descartar definitivamente una idea solo por falta de datos; puede quedar en “pendiente de investigación”.

### Comparación de productos

Evaluar la oportunidad a dos niveles:
1. **Categoría/nicho:** ¿merece la pena investigar y probar esta familia?
2. **Producto concreto:** ¿este producto resuelve el problema mejor que las alternativas y es viable producirlo?

Evitar generar muchos productos similares antes de probar el primero.

La implementación operativa del selector, sus filtros de evidencia, revisión
de solapamiento, diversidad y uso de resultados de tienda está documentada en
`rules/discovery/PRODUCT_SELECTION.md`. No se debe confundir el SEO/precio de un producto
ya elegido con la investigación necesaria para decidir qué producto crear.

---

## 8. Estrategia inicial de producto: hojas de cálculo y plantillas funcionales

Comienza con una familia de productos que pueda verificarse objetivamente.

### Candidatos de ejemplo (no asumir demanda sin investigar)

- calculadoras de presupuesto;
- comparadores de ofertas o costes;
- planificadores para autónomos;
- control de inventario y margen para pequeños negocios;
- hojas de análisis de rentabilidad de alquileres;
- control de gastos y ahorro;
- planificadores de proyectos para profesiones específicas.

Para cada candidato, investigar antes:
- quién lo compra;
- qué problema resuelve;
- qué alternativas hay;
- qué nivel de precisión requiere;
- si existen obligaciones legales o fiscales;
- qué precio puede probarse;
- cómo se puede demostrar su utilidad.

Evita empezar con consejos médicos, asesoramiento legal, recomendaciones de inversión personalizadas u otros ámbitos de alto riesgo. Si se crean calculadoras financieras o fiscales, deben indicar supuestos, limitaciones, jurisdicción y fecha, y no presentarse como asesoramiento profesional.

### Brief obligatorio antes de generar

Cada producto debe tener un documento `product_brief.json` con:
- identificador;
- nombre provisional;
- categoría y subcategoría;
- público objetivo;
- problema concreto;
- propuesta de valor;
- funciones incluidas y excluidas;
- datos que debe introducir el usuario;
- resultado que recibirá;
- idioma y jurisdicción;
- precio de prueba y justificación;
- alternativas competidoras investigadas;
- diferenciación;
- fuentes y nivel de evidencia;
- criterios de aceptación;
- riesgos y advertencias;
- licencias y recursos utilizados;
- versión y fecha.

No iniciar generación si el brief está incompleto o la idea no ha superado los controles.

---

## 9. Generación de productos

### Proceso estándar

1. Crear el brief y criterios de aceptación.
2. Generar una especificación técnica.
3. Crear el archivo mediante código o plantilla, no solo texto libre del modelo.
4. Generar instrucciones de uso.
5. Crear una ficha comercial basada únicamente en funciones existentes.
6. Generar una vista previa.
7. Ejecutar pruebas.
8. Revisar duplicados y similitud con productos ya creados.
9. Calcular coste real y tiempo de producción.
10. Enviar a revisión humana.
11. Aprobar, corregir o descartar.
12. Crear un paquete reproducible con todos los archivos y metadatos.

### Requisitos de calidad

- El archivo debe abrirse correctamente.
- Las fórmulas deben probarse con casos normales, extremos, vacíos e inválidos.
- Las entradas deben validarse.
- Los resultados deben tener unidades y formato comprensibles.
- No debe haber referencias rotas ni contenido de ejemplo presentado como real.
- Las instrucciones deben coincidir con la versión del archivo.
- Los productos deben tener una diferencia funcional clara entre sí.
- La ficha no debe prometer características que no existen.
- Los archivos deben tener nombres claros y versión.
- Incluir accesibilidad básica, instrucciones legibles y una estructura coherente.
- Verificar que no se incluyan datos personales o secretos en los archivos.

### Calidad de hojas de cálculo

- Mantener fórmulas y cálculos en funciones de Python que puedan probarse, cuando sea posible.
- Crear una hoja de instrucciones y otra de configuración/entradas claramente diferenciadas.
- Usar celdas de entrada distinguibles y proteger fórmulas si resulta práctico.
- Evitar macros salvo necesidad demostrada.
- Probar divisiones por cero, valores negativos, celdas vacías, redondeos y límites.
- Probar en LibreOffice o Excel si el entorno lo permite; no afirmar compatibilidad universal sin probarla.
- Incluir ejemplos sintéticos identificados claramente como ejemplos.
- Generar un informe de pruebas con resultados.

### Duplicados y catálogo

Calcular similitud entre productos por:
- título y descripción;
- estructura de archivos;
- funcionalidad;
- imágenes;
- público objetivo;
- propuesta de valor.

Si dos productos son casi iguales, generar una mejora significativa o no producir el segundo. No crear variaciones superficiales solo para inflar el catálogo.

---

## 10. Fichas comerciales y presentación

Cada ficha debe generarse a partir de datos reales del producto.

Campos:
- título claro;
- descripción;
- público recomendado;
- problema que resuelve;
- contenido incluido;
- formato y requisitos;
- instrucciones de descarga/uso;
- limitaciones;
- etiquetas/categoría adecuadas;
- precio de prueba;
- imágenes de vista previa;
- declaración de uso de IA cuando la plataforma lo exija;
- información de licencia y uso permitido;
- política de soporte y actualización, cuando corresponda.

No utilizar:
- testimonios inventados;
- insignias o certificados falsos;
- capturas engañosas;
- promesas de resultados garantizados;
- marcas ajenas para atraer búsquedas sin autorización;
- palabras clave irrelevantes;
- material copiado de competidores.

Las vistas previas deben representar fielmente el producto final. Genera gráficos simples mediante código cuando sea suficiente; utiliza generación de imágenes solo si aporta valor real.

---

## 11. Publicación e integraciones

### Diseño independiente de plataforma

Define una interfaz común, por ejemplo:
- `validate_listing()`
- `create_draft()`
- `upload_files()`
- `publish_listing()` (solo tras autorización)
- `fetch_metrics()`
- `pause_listing()` (requiere política de autorización apropiada)

Cada plataforma debe tener un adaptador separado y documentar las capacidades reales disponibles.

### Reglas

1. Investigar la documentación oficial vigente antes de implementar.
2. Confirmar si la API admite productos digitales, borradores, archivos, precios y estadísticas.
3. Confirmar permisos, requisitos de aplicación y límites de uso.
4. No simular que se ha publicado si solo se ha generado un archivo local.
5. No usar automatización de navegador para saltarse una API o condiciones de uso.
6. Si no hay integración autorizada, exportar un paquete preparado para publicación manual.
7. Usar claves de prueba o modo simulación siempre que exista.
8. Registrar identificador remoto, estado, fecha y respuesta de cada acción.
9. No reintentar una publicación a ciegas tras un timeout: comprobar primero si el anuncio se creó para evitar duplicados.
10. La primera publicación real requiere aprobación humana explícita.

### Flujo de aprobación

`local → validado → pendiente de aprobación → borrador remoto → aprobado para publicación → publicado`

Un fallo de validación no puede saltarse automáticamente. Toda publicación debe ser trazable y reversible cuando la plataforma lo permita.

---

## 12. Analítica y retroalimentación comercial

El sistema debe distinguir entre:
- producto creado;
- producto aprobado;
- producto anunciado;
- anuncio con impresiones;
- anuncio con visitas;
- anuncio con ventas;
- producto rentable.

No asumas que todos los marketplaces proporcionan todas estas métricas. Registra únicamente las que se observen o se importen, con su fuente.

### Métricas fundamentales

- impresiones;
- visitas/clics;
- favoritos o señales de interés disponibles;
- pedidos;
- unidades vendidas;
- conversión = pedidos / visitas, cuando las definiciones sean compatibles;
- ingresos brutos;
- comisiones;
- devoluciones y reembolsos;
- gasto en IA;
- gasto en herramientas y publicidad atribuible;
- margen de contribución estimado;
- tiempo de producción;
- beneficio por hora de trabajo;
- coste por producto aprobado;
- coste por producto publicado.

### Modelo de rentabilidad

`margen_contribucion = ingresos_brutos - reembolsos - comisiones - costes_variables - coste_IA_atribuible - publicidad_atribuible`

Registra por separado:
- costes fijos mensuales;
- impuestos que no se puedan calcular automáticamente con fiabilidad;
- trabajo del usuario;
- costes no atribuibles.

No llames “beneficio neto” a una cifra que no incluya todos los costes e impuestos pertinentes. Muestra las hipótesis.

### Interpretación prudente

- Cero ventas con cero visitas no demuestra que el producto sea malo.
- Muchas visitas sin compras puede indicar problemas de precio, presentación, confianza, ajuste con la intención de búsqueda o producto.
- Una venta aislada no demuestra que exista una demanda estable.
- Una tasa de conversión calculada con pocas visitas es muy incierta.
- Compara periodos equivalentes y registra cambios de precio, título, imagen y promoción.
- No cambies simultáneamente muchos factores si se pretende identificar qué produjo una mejora.
- Conserva un historial de versiones para poder comparar.

### Experimentos

Cada experimento debe especificar:
- hipótesis;
- variable que se cambia;
- productos afectados;
- métrica principal;
- duración o criterio mínimo de datos;
- presupuesto máximo;
- resultado observado;
- limitaciones;
- decisión siguiente.

El sistema debe sugerir experimentos pequeños y baratos, no hacer cambios masivos sin control.

---

## 13. Sistema de aprendizaje y retroalimentación

La IA debe usar los datos históricos para proponer decisiones, pero no debe modificar reglas críticas sin dejar rastro.

### Ciclo de aprendizaje

1. Importar nuevas métricas.
2. Validar y normalizar datos.
3. Detectar datos ausentes, duplicados y cambios de definición.
4. Calcular métricas por producto y categoría.
5. Comparar resultados con hipótesis previas.
6. Estimar incertidumbre.
7. Identificar productos que necesitan más tráfico/datos.
8. Proponer una acción: seguir observando, mejorar ficha, cambiar precio, mejorar producto, crear variante o pausar.
9. Estimar coste y riesgo de la acción.
10. Pedir aprobación si afecta a publicación, gasto o contenido visible.
11. Ejecutar la acción autorizada.
12. Medir el resultado y guardar la lección.

### Evitar el sobreajuste

- No construir un modelo predictivo complejo con unos pocos datos.
- Comenzar con reglas transparentes y estadísticas sencillas.
- No tomar decisiones por una sola venta.
- Mantener un grupo de productos o decisiones de referencia cuando sea viable.
- Registrar decisiones tomadas por el usuario frente a recomendaciones automáticas.
- Recalcular puntuaciones con datos nuevos, conservando el historial de versiones.
- No tratar correlación como causalidad.
- La IA debe poder concluir “no hay datos suficientes”.

---

## 14. Control de presupuesto de IA

Implementa un módulo de presupuesto antes de ejecutar generación por lotes.

### Límites iniciales recomendados

- Presupuesto objetivo: 10–20 € al mes.
- Límite de gasto mensual configurable y obligatorio.
- Límite por investigación.
- Límite por producto.
- Límite por lote.
- Estimación previa del coste de cada tarea.
- Registro de modelo, tokens o unidades consumidas, coste estimado y coste real si el proveedor lo facilita.
- Alertas al 50 %, 75 %, 90 % y 100 % del límite.
- Bloqueo automático de nuevas llamadas de pago al alcanzar el límite.
- Nunca continuar gastando silenciosamente.
- Preferir modelos económicos para clasificación, extracción y borradores.
- Reservar modelos más caros para revisión final, casos ambiguos o tareas que realmente lo justifiquen.
- Limitar longitud de entradas y salidas.
- Cachear respuestas reutilizables.
- Evitar volver a generar archivos si no ha cambiado el brief.
- No llamar al modelo para cálculos deterministas que pueda realizar Python.

Las estimaciones de coste no son facturas definitivas. Indicar siempre el proveedor y los precios vigentes configurados.

---

## 15. Seguridad, privacidad, derechos y cumplimiento

- Guardar credenciales en `.env` local ignorado por Git, o usar el almacén de credenciales del sistema si está disponible.
- Incluir `.env.example` sin valores reales.
- No mostrar claves en la interfaz ni en logs.
- No enviar datos personales de clientes al modelo salvo necesidad y autorización apropiadas.
- No recoger datos de usuarios de manera innecesaria.
- Respetar GDPR y las obligaciones de protección de datos aplicables si se procesan datos personales.
- Respetar licencias comerciales de fuentes, tipografías, imágenes, iconos y plantillas.
- Registrar la procedencia de los recursos.
- Investigar los requisitos fiscales y de información al consumidor que correspondan antes de operar comercialmente; no improvisar asesoramiento fiscal o jurídico.
- Evitar productos que generen riesgos elevados o requieran credenciales profesionales.
- No usar contenido protegido por terceros sin derechos.
- Cumplir las políticas vigentes de la plataforma sobre productos generados con IA, declaraciones obligatorias y contenido digital.
- Mantener registro de qué partes del producto se generaron con IA y qué revisiones se realizaron.
- Añadir mecanismos para corregir, pausar o retirar un anuncio si se detecta un error.

---

## 16. Registro, trazabilidad y gestión de errores

Cada ejecución debe registrar:
- identificador de trabajo;
- fecha/hora de inicio y fin;
- módulo y etapa;
- entradas y versión del prompt/plantilla;
- modelo utilizado;
- coste;
- resultado;
- error, si lo hubo;
- reintentos;
- aprobaciones;
- archivos generados;
- decisiones tomadas.

Usa identificadores únicos y operaciones idempotentes cuando sea posible. Un reintento no debe crear duplicados ni repetir cargos innecesarios.

Implementa:
- logs con rotación;
- mensajes de error claros;
- reintentos limitados;
- pausa de trabajos cuando falte configuración;
- cola de trabajos fallidos para revisión;
- copia de seguridad de SQLite y metadatos;
- exportación de datos a CSV/JSON;
- recuperación documentada.

No registres claves API, tokens, datos de pago ni información privada.

---

## 17. Esquema de datos inicial

Usa SQLite con migraciones sencillas. Como mínimo, contempla estas entidades:

### `research_runs`
- `id`
- `started_at`
- `completed_at`
- `query_config`
- `status`
- `cost_estimate`
- `cost_actual`
- `summary`

### `sources`
- `id`
- `source_type`
- `name`
- `url`
- `access_method`
- `terms_checked_at`
- `notes`

### `evidence`
- `id`
- `source_id`
- `research_run_id`
- `observed_at`
- `query`
- `market`
- `language`
- `metric_name`
- `metric_value`
- `unit`
- `evidence_type` (observed/estimated/hypothesis)
- `confidence`
- `limitations`
- `raw_reference`

### `opportunities`
- `id`
- `name`
- `category`
- `subcategory`
- `target_audience`
- `problem`
- `proposed_solution`
- `score`
- `confidence`
- `decision`
- `decision_reason`
- `created_at`
- `updated_at`
- `next_review_at`

### `score_components`
- `id`
- `opportunity_id`
- `criterion`
- `weight`
- `score`
- `rationale`
- `evidence_ids`

### `products`
- `id`
- `opportunity_id`
- `name`
- `product_type`
- `version`
- `status`
- `brief_json`
- `output_path`
- `created_at`
- `quality_score`
- `quality_report`
- `generation_cost`
- `production_time_seconds`

### `listings`
- `id`
- `product_id`
- `platform`
- `external_listing_id`
- `title`
- `description`
- `price`
- `currency`
- `status`
- `created_at`
- `published_at`
- `last_synced_at`
- `policy_review`

### `metrics_daily`
- `id`
- `listing_id`
- `date`
- `impressions`
- `visits`
- `favorites`
- `orders`
- `units_sold`
- `gross_revenue`
- `refunds`
- `fees`
- `ad_spend`
- `source`
- `data_quality_notes`

### `expenses`
- `id`
- `date`
- `category`
- `provider`
- `amount`
- `currency`
- `attributed_product_id`
- `attributed_research_run_id`
- `notes`

### `experiments`
- `id`
- `name`
- `hypothesis`
- `variable`
- `primary_metric`
- `budget_cap`
- `start_date`
- `end_date`
- `status`
- `result`
- `decision`

### `approvals`
- `id`
- `action_type`
- `target_type`
- `target_id`
- `requested_at`
- `approved_at`
- `approved_by`
- `decision`
- `notes`

### `audit_events`
- `id`
- `timestamp`
- `actor`
- `event_type`
- `entity_type`
- `entity_id`
- `details_redacted`

Usa claves foráneas, índices útiles y validaciones. Evita guardar blobs innecesarios; conserva los archivos en disco y guarda sus rutas y hashes.

---

## 18. Modos de autonomía

Implementa un selector de modo de funcionamiento:

### Modo 0 — Simulación
- No realiza llamadas de publicación.
- Puede usar datos de prueba.
- No incurre en costes de IA salvo que el usuario active una prueba.
- Sirve para probar el flujo completo.

### Modo 1 — Investigación y propuesta
- Investiga y puntúa.
- Genera briefs.
- No produce archivos finales ni publica sin aprobación.

### Modo 2 — Producción supervisada (modo predeterminado)
- Puede generar productos dentro de los límites.
- Ejecuta controles de calidad.
- Pide aprobación para publicar.
- No compra anuncios ni contrata servicios.

### Modo 3 — Publicación de borradores
- Puede crear borradores remotos mediante API autorizada.
- No los publica definitivamente.
- Solicita aprobación para el paso final.

### Modo 4 — Automatización limitada
- Solo disponible tras pruebas estables y activación explícita del usuario.
- Puede ejecutar acciones previamente autorizadas dentro de límites concretos.
- Debe detenerse ante errores, cambios de política, costes inesperados o baja confianza.
- No permite superar presupuesto, crear anuncios engañosos ni ignorar controles de derechos.

Registrar cada cambio de modo y sus límites.

---

## 19. Plan de implementación por hitos

### Hito 0 — Inspección del entorno
- Confirmar Windows, Python, Git, espacio en disco y permisos.
- Verificar si existe ya un proyecto.
- No sobrescribir archivos sin inspeccionarlos.
- Crear una rama Git o copia de seguridad antes de cambios relevantes.
- Resultado: informe de entorno y plan de instalación.

### Hito 1 — Esqueleto ejecutable
- Crear estructura de carpetas.
- Crear entorno virtual.
- Instalar dependencias mínimas.
- Crear `README.md`, configuración, logs y base de datos.
- Ejecutar la aplicación local.
- Resultado: interfaz que arranca y muestra el estado del sistema.

### Hito 2 — Persistencia y auditoría
- Crear esquema SQLite.
- Implementar operaciones CRUD para oportunidades y productos.
- Registrar eventos y costes.
- Probar reinicio y recuperación.
- Resultado: datos persistentes y exportables.

### Hito 3 — Proveedor de IA y presupuesto
- Crear interfaz abstracta de proveedor.
- Añadir un proveedor inicial configurable.
- Validar respuestas estructuradas con Pydantic.
- Registrar coste y limitar gasto.
- Probar fallos de red, límite de cuota y respuestas inválidas.
- Resultado: una llamada controlada, trazable y con límite.

### Hito 4 — Investigación y puntuación
- Crear al menos un método de investigación autorizado y otro de importación CSV/manual.
- Guardar evidencia con fuentes y fechas.
- Implementar puntuación y explicación.
- Crear ejemplos de prueba claramente ficticios.
- Resultado: lista de oportunidades ordenada y justificable.

### Hito 5 — Primer producto
- Seleccionar una sola familia de producto.
- Crear brief.
- Generar una hoja de cálculo útil.
- Implementar pruebas funcionales.
- Crear instrucciones y vista previa.
- Resultado: un producto descargable que pasa las pruebas.

### Hito 6 — Lote pequeño y revisión
- Generar un lote de 5–10 productos, no cientos.
- Medir coste, tiempo, calidad y similitud.
- Enviar a revisión humana.
- Resultado: catálogo inicial de calidad conocida.

### Hito 7 — Paquete de publicación
- Generar título, descripción, etiquetas, precio propuesto y archivos.
- Implementar validaciones de política y contenido.
- Exportar un paquete listo para subir manualmente.
- Resultado: publicación reproducible sin integración de plataforma.

### Hito 8 — Métricas y aprendizaje
- Importar CSV o introducir visitas y ventas manualmente.
- Calcular conversiones y margen con advertencias de incertidumbre.
- Generar informes.
- Proponer próximos experimentos.
- Resultado: ciclo de retroalimentación operativo.

### Hito 9 — Primera integración de plataforma
- Elegir una plataforma tras investigar demanda y condiciones.
- Consultar documentación oficial actual.
- Implementar solo las funciones realmente permitidas.
- Probar primero con borradores y operaciones idempotentes.
- Mantener aprobación humana.
- Resultado: integración fiable, sin publicaciones accidentales.

**No avances al siguiente hito hasta que el actual tenga pruebas ejecutadas y un criterio de aceptación cumplido.**

---

## 20. Criterios de aceptación del MVP

El MVP no se considera terminado hasta que:

1. La aplicación se inicia localmente en Windows siguiendo instrucciones claras.
2. El usuario puede crear y guardar una investigación.
3. Cada señal relevante tiene fuente, fecha y etiqueta de fiabilidad.
4. El motor puntúa oportunidades de forma determinista a partir de criterios configurados.
5. La decisión se explica y se puede anular con motivo.
6. Se puede generar al menos una hoja de cálculo útil desde un brief.
7. El archivo generado se abre y supera pruebas funcionales.
8. Se produce una ficha comercial que coincide con el producto real.
9. El coste de IA se registra y se aplica un límite de gasto.
10. El sistema no publica sin aprobación explícita en los modos iniciales.
11. Se pueden importar visitas y ventas desde un CSV de ejemplo.
12. Los informes distinguen datos reales, estimaciones y datos ficticios.
13. El sistema calcula ingresos y margen estimado sin llamar “beneficio neto” a datos incompletos.
14. Los errores no destruyen datos ni generan duplicados.
15. Las pruebas automatizadas pasan.
16. La documentación permite que el usuario vuelva a iniciar el sistema y haga una copia de seguridad.

---

## 21. Protocolo para medir si una categoría merece escalar

Para cada categoría, ejecutar un experimento acotado:

1. Investigar y documentar las señales disponibles.
2. Seleccionar entre 3 y 5 productos distintos que cubran necesidades específicas.
3. Producir solo después de superar el umbral de decisión.
4. Validar y revisar los archivos.
5. Publicar los productos aprobados en una plataforma.
6. Medir durante un periodo suficiente para observar tráfico; no fijar conclusiones solo por calendario.
7. Separar falta de exposición de falta de conversión.
8. Cambiar una variable por experimento cuando sea posible.
9. Calcular costes reales y margen.
10. Decidir entre escalar, mejorar, mantener en observación o detener.

No impongas un número universal de visitas o una tasa de conversión “buena” sin datos de referencia del sector, la plataforma, el precio y la fuente de tráfico. Define umbrales iniciales explícitamente como hipótesis y revísalos cuando exista una muestra útil.

### Reglas de escala

Escalar solo si:
- hay señales de demanda repetidas;
- el producto mantiene calidad;
- el margen es razonable;
- el coste de soporte no se dispara;
- no hay alertas legales o de plataforma;
- los resultados no dependen exclusivamente de una observación aislada.

---

## 22. Informe de cada sesión de desarrollo

Al finalizar cada sesión, entrega un resumen con:

1. **Qué se ha implementado.**
2. **Archivos creados o modificados.**
3. **Comandos ejecutados.**
4. **Pruebas ejecutadas y resultado exacto.**
5. **Limitaciones conocidas.**
6. **Coste aproximado de IA o servicios.**
7. **Riesgos pendientes.**
8. **Cómo ejecutar el proyecto.**
9. **Siguiente tarea concreta.**

No marques como completado algo que no se haya probado.

---

## 23. Primera tarea que debes ejecutar ahora

Empieza por el **Hito 0 y Hito 1**. No te pongas a generar productos en masa todavía.

1. Inspecciona el entorno de Windows y determina si Python y Git están instalados.
2. Comprueba el directorio de trabajo y evita sobrescribir archivos existentes.
3. Propón una ubicación clara para el proyecto, por ejemplo `C:\Users\<usuario>\Projects\digital-product-factory`, adaptándola al entorno real.
4. Crea el esqueleto del proyecto y un entorno virtual.
5. Añade configuración segura, `.gitignore`, `README.md` y pruebas mínimas.
6. Implementa una interfaz local básica que muestre estado del sistema, presupuesto configurado, trabajos y módulos pendientes.
7. Inicializa SQLite y registra una primera ejecución.
8. Ejecuta las pruebas y arranca la aplicación para verificar que funciona.
9. **No solicites ni escribas una clave API en el código.** Si todavía no existe, crea `.env.example` y deja la configuración pendiente.
10. No hagas compras, no publiques anuncios, no contrates servicios y no realices acciones externas sin autorización.
11. Al terminar, presenta el informe de sesión según la sección 22 y espera la siguiente instrucción antes de añadir complejidad.

El objetivo de esta primera tarea es entregar una base local ejecutable, comprobada y ampliable, no una maqueta ni una lista de planes futuros.
