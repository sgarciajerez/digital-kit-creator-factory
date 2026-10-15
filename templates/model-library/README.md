# Biblioteca de modelos nativos

Esta carpeta conserva modelos recuperados de la factoría para reutilizarlos en
nuevos productos. No son listings activos: son referencias de estructura,
dimensiones, fórmulas y comportamiento.

## Modelos disponibles

- `native-master-base`: retícula común, navegación, icono, colores de sistema y
  ranuras de cabecera, tabla, resumen y nota.
- `family-budget-sources-base`: modelo construido directamente desde los
  libros de `templates/sources`; reúne presupuesto, movimientos, flujo de caja,
  metas, deudas y listas.
- `wedding-planner-base`: especialización nativa del modelo familiar para
  presupuesto de boda, pagos, proveedores, invitados y checklist.

## Regla de uso

Un producto nuevo debe seleccionar uno o más modelos de esta biblioteca,
declarar las fuentes concretas de `templates/sources` y conservar las
dimensiones útiles del modelo elegido. Se pueden cambiar textos, paleta,
identidad SVG, cabeceras, listas, datos de ejemplo y fórmulas justificadas por
el caso de uso; no se deben copiar columnas con otro significado ni arrastrar
formatos de porcentaje, fechas o importes.

La fuente visual y estructural es siempre el XLSX nativo y su contrato. No se
generan representaciones paralelas para proponer, traducir o comparar diseños.
