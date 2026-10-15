# Sistema de diseño Excel nativo

Este sistema se aplica a productos nuevos junto con
`rules/FACTORY_NATIVE_SOURCE_WORKFLOW.md`.

## Prioridades

1. utilidad y sentido funcional;
2. replicabilidad desde sources y modelos nativos;
3. legibilidad de entradas, fórmulas y estados;
4. identidad visual coherente.

La complejidad objetivo es 2/10 y el máximo 4/10.

## Gramática Excel

Cada hoja usa una cabecera breve, una tabla o bloque de entrada, una zona de
resumen cuando sea útil, una nota corta y navegación nativa. Los gráficos son
opcionales y nativos, con rangos de fórmulas explícitos.

La geometría se expresa con columnas, filas, rangos, anclajes y zoom de Excel.
La única referencia visual y estructural es
`templates/excel-first-base/master_workbook.xlsx`.

## Biblioteca

- `config/EXCEL_TEMPLATE_LIBRARY.json`: primitivas y variaciones permitidas;
- `templates/excel-first-base/`: retícula común;
- `templates/model-library/`: modelos históricos recuperados;
- `templates/sources/`: ejemplos que justifican estructura y distribución.

## Reglas

- tablas nativas con cabeceras semánticas;
- totales fuera de `ListObject`;
- formatos por significado del campo;
- validaciones conectadas a rangos declarados;
- icono SVG propio o adaptado en todas las hojas;
- paleta única por producto;
- botones redondeados nativos; paneles congelados solo si se declaran por hoja en `freeze_panes`;
- anchuras y alturas explícitas, sin desbordamientos.

La calidad se decide leyendo el XLSX y ejecutando pruebas funcionales y
estructurales. No se usan HTML, CSS, PNG, capturas ni comparaciones de píxeles.
