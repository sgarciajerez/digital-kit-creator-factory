# Flujo nativo de la factoría Excel

Este es el flujo único para crear productos nuevos. El punto de partida visual
y estructural es siempre `templates/excel-first-base/master_workbook.xlsx`.
Los libros de `templates/sources/` y `templates/model-library/` aportan
patrones funcionales y modelos recuperados, pero ningún archivo de salida se
convierte en una referencia intermedia.

## Orden obligatorio

1. Detectar la oportunidad y definir el producto.
2. Inventariar los sources pertinentes.
3. Elegir el modelo funcional reutilizable de `templates/model-library/`.
4. Clonar `master_workbook.xlsx` como base visual y estructural.
5. Declarar `model_contract.json`, `model_behavior.md`, `app/README.md`, `source_lineage` y las
   variaciones permitidas.
6. Adaptar textos, columnas, fórmulas, listas, paleta, icono SVG y gráficos
   nativos sin romper la retícula del master.
7. Generar el XLSX final desde el master y el contrato funcional.
8. Validar directamente el XLSX: pestañas, dimensiones, tablas, encabezados,
   rangos, formatos semánticos, fórmulas, listas, gráficos, iconos y casos de
   uso declarados.
9. Generar la ficha del producto y la guía funcional independiente.
10. Solo después de superar todas las puertas anteriores, generar los assets
    publicitarios a partir del XLSX final validado, incluido un vídeo demo de
    uso real sobre una copia de trabajo.

## Reglas de referencia

- `master_workbook.xlsx` es la única referencia de encuadre, distribución y
  dimensiones.
- `templates/sources/` se estudia para absorber patrones; no se copia a ciegas.
- `templates/model-library/` se usa para recuperar modelos funcionales ya
  corregidos.
- `model_behavior.md` se redacta desde lo que la plantilla debe hacer, antes
  de comprobar el XLSX. Si el libro no cumple la guía, falla.
- Durante el diseño y la QA no se crean capturas ni referencias visuales
  paralelas para decidir si el libro está bien.
- Las capturas y mockups publicitarios solo se permiten después de la QA,
  dentro de `marketing_assets/`, y nunca modifican ni validan el XLSX. El
  paquete debe incluir un vídeo que muestre navegación, entrada de datos y una
  actualización real de fórmula o resultado.
- La revisión visual opcional la hace una persona abriendo el XLSX; no es una
  etapa automatizada ni una condición basada en imágenes.

## Puertas de calidad

- El master existe y abre correctamente.
- El producto conserva las pestañas, la retícula y los anclajes declarados.
- Las tablas tienen nombres y encabezados semánticos.
- Las filas de totales quedan fuera de los rangos de datos.
- Las fórmulas apuntan a las columnas correctas y las listas solo validan
  entradas compatibles.
- El icono SVG es específico o recoloreado para el producto y está presente
  en todas las hojas.
- El resultado pasa `scripts/audit_design_contract.py`,
  `scripts/preflight_visual_contract.py` y la validación funcional del modelo.
