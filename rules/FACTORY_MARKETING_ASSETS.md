# Assets publicitarios de productos Excel

Los mockups, capturas, portadas y vídeos son una fase posterior a la creación
y obligatoria cuando se prepara el paquete publicitario. No forman parte del
diseño ni de la validación funcional o estructural.

## Precondiciones

Solo se pueden generar cuando el producto ya tiene:

- el XLSX final en `app/`;
- la auditoría semántica aprobada;
- el preflight nativo aprobado;
- la validación funcional aprobada;
- la ficha y el contrato del producto actualizados.

## Reglas

- La entrada es siempre el XLSX final validado.
- La salida se guarda en `outputs/listings/product-N/marketing_assets/`.
- El paquete debe incluir capturas originales del Excel tomadas desde el XLSX
  final y desde el propio proyecto.
- Las capturas deben estar bien encuadradas sobre el área útil de Excel: sin
  bordes blancos, márgenes sobrantes, navegador, escritorio, barras ajenas,
  overlays de depuración, cursores visibles ni elementos extraños.
- Las capturas reutilizadas del proyecto solo son válidas si corresponden a la
  misma versión del XLSX que ha superado la QA; nunca se reciclan capturas de
  una versión anterior o de otro producto.
- Conservar los originales en `marketing_assets/source_captures/` y usar las
  versiones derivadas en `marketing_assets/mockups/` o para la portada.
- Se pueden crear mockups de dispositivo, portada y otros recursos a partir de
  esas capturas originales, pero no deben ocultar por completo el Excel real.
- Cada paquete publicitario debe incluir un vídeo demo del XLSX final.
- El vídeo debe mostrar el libro en uso: introducir o modificar datos de
  ejemplo, navegar por sus hojas y comprobar al menos una actualización de
  fórmula o resultado relevante.
- El vídeo debe grabarse sobre una copia de trabajo del XLSX validado para que
  la demostración no altere el archivo entregable.
- Los assets no se leen como evidencia de QA y no pueden cambiar el XLSX.
- Si cambia el libro, primero se repite la QA y después se regeneran los assets.
- Los nombres y textos publicitarios deben describir solo funciones que existen
  en el contrato y en el libro final.
