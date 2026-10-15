# Contrato de diseño Excel nativo

```json
{
  "master_workbook": "templates/excel-first-base/master_workbook.xlsx",
  "native_reference_mode": "master_workbook_contract",
  "source_of_truth": "master_workbook.xlsx + model_contract.json + model_behavior.md"
}
```

El master fija la retícula, las dimensiones, la navegación, la posición de las
tablas y los componentes permitidos. Cada producto puede cambiar textos,
columnas semánticas, fórmulas, paleta, icono y datos de ejemplo dentro de los
slots declarados.

## Comprobaciones obligatorias

- Las pestañas y su orden coinciden con la ficha.
- Las columnas tienen nombres concretos y significado correcto.
- Las tablas tienen rangos completos, encabezados únicos y totales fuera de
  los datos.
- Los formatos de porcentaje solo aparecen en ratios.
- Las fórmulas apuntan a la tabla o columna que representan.
- Las listas de validación solo aceptan valores compatibles.
- Las dimensiones de filas y columnas conservan la distribución del master.
- Las formas de navegación y el SVG de identidad están en todas las hojas.
- El `model_behavior.md` describe el comportamiento esperado y sus pruebas.

La validación se hace sobre el XLSX. No hay capturas, comparación de píxeles,
HTML ni navegador.
