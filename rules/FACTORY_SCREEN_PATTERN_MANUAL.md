# Manual del patrón nativo

El patrón común se calibra sobre `templates/excel-first-base/master_workbook.xlsx`
y los modelos de `templates/model-library/`. Las pantallas nuevas heredan esa
retícula y solo cambian slots declarados.

## Método

1. Elegir una hoja representativa del master o del modelo funcional.
2. Declarar columnas, filas, rangos, formas, SVG, tablas y navegación.
3. Validar el master y el XLSX final con el contrato nativo.
4. Aplicar el mismo contrato al resto de hojas.
5. Reparar el constructor o el XLSX cuando falle una comprobación.

## Invariantes

- misma cabecera, márgenes y ritmo de filas;
- tablas dentro de rangos declarados y con cabeceras semánticas;
- columnas suficientemente anchas para sus valores visibles;
- padding y alineación explícitos;
- totales y notas fuera del rango de entrada;
- icono SVG con proporción intacta en cada hoja;
- botones redondeados completos, con separación y pestaña activa;
- `freeze_panes` coincide con la adaptación declarada por hoja; ausencia de líneas residuales;
- paleta homogénea por producto.

## Variaciones

Solo se permiten cantidad de filas, rangos de tarjetas, copy, fórmulas,
validaciones, gráfico declarado y omisiones justificadas. Cambiar la retícula
común exige versionar el modelo nativo.

## Orden de reparación

1. significado de tablas y columnas;
2. fórmulas y rangos;
3. columnas, filas y encuadre;
4. cabecera, icono y navegación;
5. bordes, color y detalles.

## Validación

El orden es contrato, estructura, funcionalidad y QA nativo del XLSX. No hay
capturas, inventario HTML ni auditoría HTML-to-XLSX.
