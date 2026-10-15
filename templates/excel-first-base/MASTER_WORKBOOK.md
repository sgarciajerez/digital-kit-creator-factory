# Ficha de plantilla: `master_workbook.xlsx`

> Perfil funcional y guía de selección de la plantilla maestra de presupuesto familiar. Esta ficha describe el libro que existe hoy y propone criterios para decidir si es una base adecuada para un nuevo producto; no sustituye el contrato funcional ni las pruebas del Excel generado.

## Identidad de la plantilla

| Campo | Valor |
|---|---|
| Archivo | [`master_workbook.xlsx`](master_workbook.xlsx) |
| ID actual | `family-budget-workbook-v1` |
| Familia | Finanzas personales y del hogar |
| Tipo de modelo | Registro operativo + planificación mensual + análisis y seguimiento |
| Idioma de la base | Español |
| Libro | `.xlsx` nativo, sin macros requeridas |
| Capacidad observada | 9 hojas, 8 tablas de Excel, aproximadamente 305 fórmulas, 2 gráficos y un recurso gráfico por hoja |
| Fuente funcional | [`model_behavior.md`](model_behavior.md), [`model_contract.json`](model_contract.json) y [`native_base_contract.json`](native_base_contract.json) |
| Manifiesto | [`template_manifest.json`](template_manifest.json) |
| Especificación visual de referencia | [`design_spec.json`](design_spec.json) |

Los recuentos técnicos son una fotografía del archivo actual. Si se modifica el maestro, esta ficha, el manifiesto, los contratos y las pruebas deben revisarse juntos.

## Qué resuelve

El modelo permite registrar movimientos de dinero una sola vez y utilizar esos registros para observar el gasto real, compararlo con el plan, seguir el flujo de caja y controlar metas de ahorro y deudas. Su eje operativo es la hoja `Movimientos`: la tabla de gastos reales y varios indicadores mensuales se calculan a partir de sus entradas.

La plantilla es más que una colección de tablas visuales. Incluye una pequeña aplicación de planificación financiera dentro de Excel, con distintas capas:

1. **Entrada:** movimientos, presupuesto, parámetros, metas y deudas.
2. **Transformación:** agregaciones mensuales y por categoría, totales y cálculos de seguimiento.
3. **Lectura:** resumen comparativo, flujo de caja y visualizaciones.
4. **Orientación:** hoja de inicio, textos de ayuda, navegación y listas editables.

## Inventario de hojas

| Hoja | Papel en el modelo | Qué contiene y cómo se usa |
|---|---|---|
| `Inicio` | Orientación | Explica el propósito y los primeros pasos. Incluye una imagen/identidad gráfica y referencias rápidas a parámetros del libro. |
| `Resumen` | Panel de decisión | Presenta indicadores del mes seleccionado y compara plan, gasto real, diferencia y porcentaje gastado por categoría. Incluye una tabla y un gráfico nativo. |
| `Presupuesto` | Plan anual | Organiza las categorías en columnas de enero a diciembre y calcula totales mensuales y anuales. Las categorías se vinculan a la lista de parámetros. |
| `Gastos reales` | Resultado observado | Resume los gastos registrados por mes y categoría. Sus fórmulas agregan las entradas de `Movimientos`; no está pensada como una segunda tabla de captura manual. |
| `Movimientos` | Registro de eventos | Recibe fecha, mes, tipo, categoría, concepto, importe, pagador y estado. Es la fuente operativa para los gastos e ingresos reales y los movimientos de ahorro. |
| `Flujo de caja` | Evolución mensual | Reúne ingreso, gasto y ahorro previstos y reales, y calcula el disponible. Contiene los doce meses y un total, además de un gráfico nativo. |
| `Metas y fondo` | Ahorro dirigido | Permite fijar metas, calcular el objetivo necesario, registrar lo ahorrado, mostrar lo pendiente y el porcentaje de progreso y consultar un estado. |
| `Deudas` | Seguimiento de pasivos | Registra saldo, tipo de interés, pago y plazo; calcula un indicador de intereses estimados y el estado de cada deuda. |
| `Listas` | Configuración | Centraliza parámetros del hogar, categorías, meses, tipos y otros valores que alimentan las fórmulas y las listas desplegables. Debe editarse con criterio porque afecta a varias hojas. |

## Funcionamiento del modelo

### Flujo principal de datos

```text
Listas ───────────────► configuración, categorías y meses
   │                              │
   ├──────────────► Presupuesto ──┼──► Resumen
   │                              │
Movimientos ─► Gastos reales ─────┤
       └──────► Flujo de caja ────┘

Metas y fondo ─► cálculo de objetivo, pendiente y progreso
Deudas ────────► cálculo de interés estimado y estado
```

- **Presupuesto** expresa el plan por categoría y mes; sus importes se agregan por mes y por año.
- **Movimientos** recoge la actividad real. El tipo de movimiento distingue, al menos, ingresos, gastos y ahorro. Las fórmulas `SUMIFS` agregan importes según el tipo, categoría y mes.
- **Gastos reales** obtiene sus valores a partir del registro, de modo que la misma operación no debería introducirse también manualmente en esta hoja.
- **Resumen** selecciona un mes y compara el presupuesto con lo observado. Calcula diferencia como plan menos real y porcentaje gastado como real dividido por plan, con protección frente a divisiones por cero.
- **Flujo de caja** presenta la evolución mensual y obtiene el disponible restando del ingreso real los gastos y el ahorro real.
- **Metas y fondo** calcula el importe objetivo a partir de gasto mensual y meses, limita el pendiente a cero como mínimo y deriva el progreso del ahorro acumulado.
- **Deudas** deriva un indicador de interés a partir de saldo, porcentaje y plazo. Es una estimación orientativa de la plantilla, no un cuadro de amortización contractual.

Las fórmulas de Excel se recalculan al abrir el libro. Las celdas derivadas no deben reemplazarse por valores fijos al adaptar una copia.

## Estructura y controles observados

- Nueve pestañas, con nombres funcionales en español.
- Ocho tablas nativas (`ListObject`) con encabezados semánticos y rangos definidos. La hoja de inicio no es una tabla de datos.
- Dos gráficos nativos: una comparación de presupuesto frente a gasto real y una lectura del flujo de caja.
- Validaciones desplegables concentradas en `Listas`, sobre parámetros que condicionan otras hojas.
- Fórmulas distribuidas en varias hojas, incluidas referencias cruzadas, `SUM`, `SUMIFS`, `INDEX`, `MATCH`, `IF`, `IFERROR` y `MAX`.
- Imágenes/recursos de identidad presentes en las hojas, sin requerir iconos SVG externos en la derivación descrita por el contrato nativo.
- La geometría del maestro está sujeta a contratos: posiciones de cabecera, áreas de tablas, columnas, gráficos y totales no se deben mover de forma improvisada.
- El rango operativo de `Movimientos` es finito: el contrato tabular actual llega a la fila 27, mientras que ciertas fórmulas consultan hasta la fila 200. Al ampliar filas hay que redimensionar la tabla, mantener las validaciones y comprobar que las fórmulas, filtros y rangos sigan cubriendo las nuevas entradas.

## Qué puede personalizar un producto derivado

La regla segura es clonar el maestro y declarar la adaptación en su especificación; el archivo maestro no se edita para cada producto. Según el contrato y las pruebas del modelo, son puntos de variación previstos:

- textos, idioma y ayudas;
- paleta visual coherente del producto;
- nombres de hojas y propósito, si se conserva el contrato funcional equivalente;
- categorías, encabezados y columnas de entrada;
- datos de ejemplo realistas y coherentes con el caso de uso;
- valores de listas desplegables;
- fórmulas y rangos, solo si se actualizan también el contrato y las pruebas;
- formatos numéricos según significado: importes, porcentajes, cantidades y fechas;
- utilidad, anclaje y rangos de los gráficos.

La adaptación no debe dejar fórmulas de presupuesto familiar en un libro cuya temática haya cambiado, ni cambiar solo las etiquetas de pantalla y dar por hecho que el modelo ya representa el nuevo dominio.

## Para qué productos es ideal

### Encaje alto

Selecciona esta base cuando el producto requiera la mayoría de estos elementos:

- planificación frente a valores reales;
- movimientos o transacciones que puedan clasificarse por fecha, categoría y tipo;
- lectura periódica mensual, con posibilidad de resumen anual;
- totales, diferencias, importes disponibles o progreso hacia objetivos;
- un resumen de gestión más una o varias hojas de captura estructurada;
- controles de ingresos, gastos, ahorro, fondos, pagos o deudas.

Ejemplos de productos que podrían derivarse, siempre que el modelo se rediseñe y valide para su nuevo contexto:

- presupuesto doméstico por persona, pareja o familia;
- registro de gastos compartidos o liquidación de cuentas del hogar;
- planificador de ahorro y fondo de emergencia;
- planificador de pago de deudas;
- control de ingresos y gastos para un proyecto pequeño o autónomo;
- presupuesto de una boda, mudanza o reforma cuando el usuario necesite comparar partidas presupuestadas con pagos reales, fechas y saldo pendiente.

Los dos últimos ejemplos pueden necesitar columnas, estados, hitos, depósitos o lógica de pagos que no están resueltos por el maestro de forma automática. Deben tratarse como una adaptación funcional, no como un simple cambio de color o de nombres.

### Encaje parcial

Puede servir como punto de partida si el producto comparte solo una parte del patrón, por ejemplo un seguimiento de metas, una agenda de pagos o un control por períodos. Antes de elegirla, decide qué hojas sobran, cuáles faltan y qué relaciones nuevas son necesarias. Si el producto no requiere transacciones clasificadas ni comparación entre plan y real, probablemente se heredaría complejidad innecesaria.

### No es la opción adecuada

Evítala como plantilla principal para:

- planificadores centrados en calendario, citas o cronogramas;
- inventario, ventas por producto o logística de pedidos;
- trackers de hábitos, rutinas, lectura, estudio o salud sin dimensión presupuestaria;
- listas de invitados, seating charts o directorios;
- calculadoras puras o modelos de simulación técnica;
- formularios, hojas de trabajo imprimibles o kits cuyo elemento central no sea una tabla de datos;
- dominios donde el estado depende de hitos, dependencias o unidades incompatibles con importes y meses.

Esos productos necesitan una plantilla distinta o una estructura diseñada expresamente. La similitud superficial —tener filas, columnas y un dashboard— no basta para justificar reutilizar este modelo.

## Regla de selección dentro de la biblioteca

Para decidir entre esta y futuras plantillas, puntúa cada criterio de 0 a 2: `0` = no aplica; `1` = aplica en parte; `2` = es central para el producto.

| Criterio | Pregunta de selección |
|---|---|
| Registro de eventos | ¿El cliente registra transacciones o sucesos repetidos con fecha, categoría e importe? |
| Plan frente a real | ¿Debe comparar una previsión con resultados observados? |
| Periodicidad | ¿El mes o el año es una unidad central de revisión? |
| Análisis financiero | ¿Necesita saldos, totales, flujo de caja, ahorro o deuda? |
| Reutilización funcional | ¿Puede aprovechar una parte sustancial de las fórmulas y relaciones ya probadas? |

Orientación inicial: **8–10 puntos**, candidata prioritaria; **5–7**, candidata condicionada a una adaptación descrita; **0–4**, buscar otra plantilla. Además del total, descarta esta base si el modelo financiero no es central, aunque comparta etiquetas como “resumen” o “progreso”. El puntaje ayuda a ordenar candidatas; no sustituye el juicio sobre seguridad, facilidad de uso, mantenimiento y validación.

Para comparar plantillas futuras, cada ficha debería declarar al menos: identificador y versión, propósito, modelo de datos, hojas y rol, tipos de entrada, fórmulas/relaciones, límites de capacidad, capacidades visuales, dependencias, productos adecuados, productos excluidos y pruebas obligatorias. Así la factoría podrá elegir por compatibilidad funcional y no por parecido visual.

## Límites y riesgos de adaptación

1. **Complejidad:** nueve pestañas y cientos de fórmulas son demasiadas para un producto pequeño. No se deben mostrar hojas que el cliente no necesita.
2. **Rangos acotados:** las tablas no son infinitas. Si aumenta la capacidad, redimensiona las tablas y actualiza todas las fórmulas, gráficos, filtros, validaciones y pruebas afectadas.
3. **Semántica de categorías:** las fórmulas dependen de nombres, meses y tipos concretos. Cambiar una etiqueta aislada puede romper las coincidencias.
4. **Localización:** comprueba formatos de fecha y moneda, separadores, nombres visibles y valores de listas en Excel de la región de destino.
5. **Cálculo:** bibliotecas de lectura/escritura pueden conservar fórmulas sin calcular sus resultados. Valida el archivo final en Excel de escritorio cuando el flujo lo requiera.
6. **Interés estimado:** la hoja de deuda simplifica el cálculo; no debe promocionarse como simulador exacto de amortización o asesoramiento financiero.
7. **Integridad contractual:** todo cambio en relaciones o fórmulas exige actualizar `model_contract.json`, `model_behavior.md` y los casos de QA; no basta con que el archivo abra sin errores.

## Selección recomendada para la factoría

**Etiqueta de catálogo:** `finanzas_personales_hogar`  
**Prioridad sugerida:** alta solo para productos donde el presupuesto, los movimientos y el seguimiento financiero sean el núcleo.  
**Nivel de adaptación:** medio/alto cuando se salga de presupuesto familiar.  
**Reutilizar como:** modelo funcional y visual de finanzas; no como plantilla universal de hojas de cálculo.  
**Condición mínima antes de generar:** confirmar el encaje del caso de uso, declarar las hojas y datos que se conservan, sustituir las relaciones que ya no aplican y ejecutar las pruebas funcionales y visuales sobre el libro derivado.

## Archivos que se deben mantener sincronizados

- `master_workbook.xlsx`: artefacto ejecutable de referencia.
- `template_manifest.json`: identidad, lineage, capacidad y reglas de clonación.
- `native_base_contract.json`: geometría, tablas y reglas de presentación nativa.
- `model_contract.json`: entradas, resultados, relaciones, fórmulas y casos de aceptación.
- `model_behavior.md`: explicación legible del comportamiento esperado.
- `design_spec.json`: diseño y contenido declarativo de una familia/producto de referencia; no debe asumirse que reemplaza el contrato funcional.
- Esta ficha: perfil de selección comprensible para la factoría y para quien mantiene la biblioteca.

Cuando se cambie una capacidad, primero decide si es una revisión compatible de esta familia o si merece una nueva plantilla con otro `template_id`. Si se altera la geometría o el modelo de forma incompatible, versiona la plantilla en lugar de actualizarla silenciosamente.
