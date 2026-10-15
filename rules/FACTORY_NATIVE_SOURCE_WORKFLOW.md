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
5. Declarar `model_contract.json`, `model_behavior.md`, `app/README.md`, una copia PDF de la guía,
   el perfil de kit digital, `source_lineage` y las variaciones permitidas.
6. Adaptar textos, columnas, fórmulas, listas, paleta, icono SVG y gráficos
   nativos sin romper la retícula del master.
7. Generar el XLSX final desde el master y el contrato funcional.
8. Validar directamente el XLSX: pestañas, dimensiones, tablas, encabezados,
   rangos, formatos semánticos, fórmulas, listas, gráficos, iconos y casos de
   uso declarados.
9. Generar la ficha del producto y la guía funcional independiente.
10. Investigar SEO para el anuncio Etsy en su idioma y mercado, documentar las
    fuentes y decisiones en `seo_keyword_research.json`, y construir título y
    etiquetas claros con frases relevantes para el producto real.
11. Solo después de superar todas las puertas anteriores, generar los assets
    publicitarios a partir del XLSX final validado, incluido un vídeo demo de
    uso real sobre una copia de trabajo.

## Composición del kit digital

La factoría genera las dos variantes del XLSX, `README.md` y su PDF legible.
`config/digital_kit_catalog.json` declara las necesidades del comprador y las
reglas que seleccionan complementos. Cada módulo se elige cuando sus
disparadores coinciden con las necesidades del producto; la decisión y el
motivo quedan en `digital-kit-manifest.json`. Un producto sin perfil o con
necesidades sin cubrir no puede marcarse listo.

El e-book es opcional u obligatorio según `ebook_required` en el perfil. Los
complementos específicos (por ejemplo, hojas imprimibles) se declaran como
módulos con contenido y condiciones de selección. El kit final incluye un
índice PDF, manifiesto de hashes y un ZIP cuya integridad se comprueba antes
de habilitar el paquete comercial.

## Regla histórica del e-book (sustituida por el catálogo de kits)

Cada producto incluye, ademÃ¡s de las versiones de Excel y la guÃ­a Markdown,
un PDF tipo e-book con contenido original y prÃ¡ctico sobre su tema. La ficha
de diseÃ±o declara un manuscrito de exactamente diez puntos y la ruta de una
portada especÃ­fica generada con IA. La portada usa la paleta del Excel y el
texto se compone en negro a 12 puntos. El documento lleva portada, una pÃ¡gina
de respeto en blanco, un Ã­ndice de diez entradas y salto de pÃ¡gina entre
puntos. Cada punto ocupa al menos una pÃ¡gina; no se aceptan continuaciones
residuales de unas pocas lÃ­neas. La factorÃ­a valida el PDF y lo incorpora al
paquete del producto; si falta el manuscrito o la portada, la generaciÃ³n falla.

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

## SelecciÃ³n de herramienta HTML para el kit

La factorÃ­a compara las necesidades declaradas del producto con las
herramientas candidatas de `digital_kit_catalog.json` y selecciona la de mayor
coincidencia. Si ninguna aporta valor claro, no fuerza una herramienta. Cuando
hay coincidencia, genera un Ãºnico `.html` autocontenido, adaptable a mÃ³vil,
coherente con la paleta y el idioma del producto, y lo incluye en el ZIP,
`digital-kit-manifest.json`, el Ã­ndice y la ficha/inventario del kit.

La primera familia son calculadoras de una sola tarea (por ejemplo, reserva de
mudanza, imprevistos de reforma o separaciÃ³n de plantas). Cada herramienta
explica sus entradas y fÃ³rmula, valida los valores y funciona sin conexiÃ³n ni
servicios externos. No transmite ni guarda datos del cliente. Las candidatas
se eligen mediante disparadores explÃ­citos de necesidades, no por coincidencias
casuales en el texto del tÃ­tulo; los perfiles nuevos pueden aÃ±adir candidatas
declarativas y una operaciÃ³n compatible.

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
