# Base de presupuesto familiar: síntesis de fuentes

Esta carpeta contiene la plantilla maestra visual y estructural. El libro no
reproduce literalmente ninguna fuente: toma sus encuadres y patrones de uso,
y los normaliza en una retícula única para que las siguientes plantillas
puedan heredarla.

## Decisiones comunes incorporadas

- Cabecera nativa en las filas 2–4, con título, subtítulo y nombre de pantalla.
- Resumen compacto en `B6:M7`, inspirado en los paneles de resultados y hojas
  de consolidación.
- Tabla principal con encabezado de color, entradas amarillas y fórmulas
  verdes; el encuadre se conserva aunque cambie la identidad de cada libro.
- Comparación planificado/real y diferencia, común en presupuestos personales,
  empresariales, de ventas, de obras y de eventos.
- Vista mensual/anual, registro de movimientos, categorías y listas de
  validación, tomados de los libros de flujo de efectivo y planificación.
- Pestañas específicas para deudas, metas y fondo de emergencia, sin forzar
  esos bloques en el dashboard.
- Un gráfico nativo opcional, anclado en `H9`, y pestañas nativas de Excel como
  navegación; no se añaden botones dentro de las hojas.

## Dimensiones y distribución

La inspección de las hojas legibles muestra encabezados y bloques de tabla con
alturas habituales de aproximadamente 15–30 puntos, columnas de contenido
entre 10 y 20 caracteres y una zona de resumen separada del detalle. La base
lo convierte en medidas estables y deliberadas:

- zoom de apertura del 95 % y lienzo útil de `B:M`;
- filas 2–4 para la cabecera, a 28 puntos; filas de datos a 22 puntos;
- resumen en dos filas (`B6:M7`), con cuatro bloques de tres columnas;
- dashboard: tabla en `B:E`, gutter en `F:G` y gráfico en `H:M`;
- pantallas de tabla: anchos propios para que una quinta, sexta o séptima
  columna no caiga dentro de un gutter estrecho;
- nota inferior en la fila 28 y pestañas nativas fuera de la retícula.

Así se conserva el encuadre y la escala de los ejemplos sin arrastrar sus
anchos accidentales ni sus combinaciones de celdas específicas.

## Trazabilidad por fuente

| Fuente | Patrón aprovechado |
|---|---|
| `BALANCE DE PÉRDIDAS Y GANANCIAS.xlsx` | panel de resumen, categorías y comparación entre periodo anterior, presupuesto y real |
| `COMPARACIÓN DE HIPOTECAS.xlsx` | entradas compactas de deuda y comparación gráfica |
| `costo-de-startups-en-excel.xlsx` | estimado frente a real y diferencia |
| `cronograma-de-actividades-en-excel.xlsx` | tabla de seguimiento con responsable, fecha y estado |
| `FLUJO DE EFECTIVO PERSONAL.xlsx` | flujo diario, mensual y anual separado por ingresos, gastos, discrecional y ahorro |
| `fondo-de-emergencia-en-excel.xlsx` | meta de ahorro, pendiente y progreso |
| `formato-de-cotizacion-en-excel.xlsx` | cabecera de datos y tabla de conceptos con importe |
| `gastos-de-empresas-en-excel.xlsx` | presupuesto, real, diferencia monetaria y porcentual |
| `planilla-de-excel-de-control-de-stock-y-gestion-de-cobranzas.xlsm` | registro de movimientos, catálogos y estados derivados |
| `planilla-de-excel-de-planeamiento-de-compras.xlsx` | presupuesto de compras y detalle por concepto |
| `planilla-de-excel-de-presupuesto-de-fiesta.xlsx` | presupuesto contra gasto real y distribución por categorías |
| `planilla-de-excel-de-presupuesto-de-navidad.xlsx` | varias tablas de categorías y comparación presupuestada/real |
| `planilla-de-excel-de-presupuesto-de-obra-de-vivienda.xlsx` | detalle de partidas, cantidades, precios unitarios y totales |
| `planilla-de-excel-de-presupuesto-de-ventas.xlsx` | matriz mensual y total anual |
| `planilla-de-excel-para-hoja-de-presupuesto.xlsx` | bloque de datos, líneas de detalle y total de documento |
| `planilla-de-excel-para-planeamiento-de-ventas.xlsx` | registro de oportunidades, estados y reporte mensual |
| `planilla-de-excel-para-presupuesto-anual-mensual.xlsx` | doble vista planeado/real y diferencia |
| `planilla-de-excel-para-presupuesto-de-casamiento.xlsx` | presupuesto por categorías de evento |
| `planilla-de-excel-para-presupuesto-de-obra.xlsx` | cabecera de presupuesto y detalle de partidas |
| `Presupuesto de gastos empresariales.xlsx` | gastos planificados, reales, desviaciones y análisis |
| `presupuesto-cascada-en-excel.xlsx` | plan inicial, actualización mensual y seguimiento de desviación |
| `presupuesto-de-calidad-y-control-en-excel.xlsx` | registro mensual de conceptos y panel de progreso |
| `presupuesto-de-capacitacion-en-excel.xlsx` | presupuesto por periodos y ejecutado frente a diferencia |
| `presupuesto-de-gastos-personales-en-excel.xlsx` | categorías personales, presupuesto/real y reportes mensual y por categoría |
| `presupuesto-de-ingresos-y-costos-en-excel.xlsx` | detalle mensual planeado/real y consolidado |
| `presupuesto-de-rrhh-en-excel.xlsx` | matriz mensual, mes seleccionado y progreso |
| `Programación de la amortización del préstamo.xlsx` | parámetros, pago mensual, interés y meses restantes |
| `PROYECCIÓN DE PÉRDIDAS Y BENEFICIOS.xlsx` | proyección mensual, total anual y porcentajes |
| `remodelacion-de-cocina-en-excel.xls` | presupuesto de obra y detalle de conceptos; formato antiguo no legible por `openpyxl` |

La fuente de presupuesto de casamiento contiene XML de hojas inválido para
`openpyxl`; se ha conservado en el inventario y se ha incorporado su patrón
visible de presupuesto por categorías a partir del nombre y del conjunto de
fuentes relacionadas. Las decisiones de estructura se apoyan en los demás
libros legibles y no en un estilo exclusivo de una sola fuente.

## Regla para futuros libros

Se puede cambiar el tema visual, el texto, las categorías, las filas y las
fórmulas. No se debe cambiar el encuadre compartido ni crear una composición
visual distinta por pestaña sin subir la versión de la plantilla maestra.
