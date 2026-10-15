# Arquitectura de diseño de la factoría

La arquitectura separa tres entradas que no deben mezclarse:

```text
templates/sources/                 patrones estudiados
templates/model-library/            modelos funcionales reutilizables
templates/excel-first-base/
  master_workbook.xlsx              única base visual y estructural
  native_base_contract.json         contrato del master
        |
        v
outputs/listings/product-N/
  design_spec.json
  model_contract.json
  model_behavior.md
  app/*.xlsx                        producto final
```

## Contrato por producto

Cada producto declara:

- `master_workbook`: normalmente `templates/excel-first-base/master_workbook.xlsx`;
- `model_base` y `source_lineage`;
- `screen_contracts`, pestañas y componentes permitidos;
- `model_contract.json` y `model_behavior.md`;
- paleta, icono SVG, tablas, fórmulas, listas y gráficos nativos.

El `master_workbook.xlsx` fija la distribución, encuadre, dimensiones y
navegación. Las variaciones se limitan a los slots declarados. Un modelo
funcional puede aportar hojas o lógica de dominio, pero no sustituye el master
visual sin versionar explícitamente la arquitectura.

## Validación

La validación se hace leyendo los libros XLSX con comprobaciones de estructura,
semántica y funcionalidad. Debe detectar encabezados incorrectos, rangos de
tablas mal dimensionados, porcentajes aplicados a columnas que no son ratios,
fórmulas incoherentes, listas incompatibles, líneas o bordes no declarados y
variaciones de la retícula.

No se generan ni se exigen archivos HTML, CSS, PNG, capturas, referencias de
píxeles o sesiones automatizadas de Excel.
