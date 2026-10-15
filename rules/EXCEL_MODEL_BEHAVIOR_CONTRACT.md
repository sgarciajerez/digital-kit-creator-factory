# Contrato funcional de modelos Excel

Cada base y cada producto Excel debe incluir un `model_behavior.md` junto a su
ficha. Ese documento es una especificación normativa y una guía de uso: define
lo que la plantilla debería hacer antes de que se genere el XLSX.

## Regla de independencia

El Markdown se escribe desde un contrato declarativo de diseño, no leyendo el
libro generado. El XLSX se compara después con ese contrato. Nunca se corrige
el Markdown para que coincida con una fórmula equivocada, una columna mal
nombrada o un total que haya acabado dentro de una tabla.

## Contenido obligatorio

Cada `model_behavior.md` debe incluir:

- objetivo de la plantilla y guía breve para el usuario;
- hojas esperadas;
- tablas, rangos y columnas con nombres semánticos;
- rol de cada columna (`input` o `derived`) y significado (`amount`, `count`,
  `percent`, `date` o `text`);
- relaciones entre hojas y fórmulas que deben existir;
- casos de prueba con entradas y resultados esperados;
- criterios explícitos de fallo;
- una declaración de que el documento es independiente del XLSX generado.

## Verificación de implementación

La factoría debe comprobar, como mínimo:

1. que existe cada hoja y tabla declarada;
2. que el rango y las cabeceras coinciden con el contrato;
3. que las filas de total quedan fuera del rango filtrable;
4. que las fórmulas declaradas están en sus celdas y conservan sus relaciones;
5. que los formatos numéricos corresponden al significado de la columna;
6. que los casos de prueba se ejecutan sobre una copia temporal cuando hay
   Excel disponible.

Si una comprobación falla, el producto queda bloqueado hasta corregir el
contrato de diseño o la implementación. No se permite cambiar el resultado
esperado para ocultar el fallo.

## Ficha del producto

La ficha (`app_design_spec.md` y `design_spec.json`) debe enlazar el archivo
`model_behavior.md` y explicar al usuario que puede editar entradas, consultar
resultados calculados y usar los casos de verificación como comprobación de
funcionamiento. La ficha no sustituye al contrato: lo presenta y lo hace
usable.

