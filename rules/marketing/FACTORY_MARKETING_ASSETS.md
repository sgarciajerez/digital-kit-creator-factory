# Assets publicitarios de productos Excel

Las capturas del libro, los mockups y los vídeos demo forman una etapa común y
obligatoria de la factoría. No deben depender del ID de un producto concreto.
Se generan después de superar la QA y antes de preparar el anuncio para Etsy;
no forman parte de la validación funcional o estructural.

## Precondiciones

Solo se generan cuando el producto ya tiene:

- el XLSX final en `app/`;
- auditoría semántica, preflight nativo y validación funcional aprobados;
- ficha y contrato del producto actualizados.

## Reglas

- La entrada son el XLSX final validado, la especificación y el contrato del
  producto. No se pueden inferir funciones que no existan en esos artefactos.
- La etapa común se ejecuta como `MARKETING_ASSETS` en el flujo de factoría.
- La salida se incorpora al `draft_package` del producto: `source_captures/`,
  `marketing_mockups/` y `marketing_video/`.
- Las capturas originales se toman del XLSX final abierto en Microsoft Excel y
  muestran hojas funcionales declaradas en la especificación. Nunca se reciclan
  capturas de otra versión del libro u otro producto.
- Los mockups deben conservar una vista reconocible del Excel real y seguir la
  paleta declarada por el producto. Sus titulares y beneficios deben proceder
  de las funciones del contrato y del libro final.
- Todo texto visible en portadas, mockups, iconos y recuentos debe estar en el
  mismo idioma que el Excel, según `design_language` en la especificación.
  Mantener los nombres reales de archivo tal como se entregan; traducir las
  etiquetas descriptivas y no mezclar idiomas dentro de una misma galería.
- Cada producto necesita una identidad de marketing propia: nombre y motivo
  visual explícitos en `visual_identity`, paleta, y arte/ilustración del tema.
  La factoría bloquea identidades repetidas entre productos activos; cambiar
  únicamente el título o dos colores no cuenta como identidad distinta.
- La identidad de marketing debe especificar también `marketing_art_direction`,
  `marketing_layout` y `marketing_artwork_file`. La dirección describe el mundo
  visual, la composición, el medio y los elementos prohibidos; el layout elige
  una composición de galería no utilizada por otro producto activo; el archivo
  contiene arte original generado para ese producto. Un nombre, un motivo corto
  y una recoloración no satisfacen esta regla.
- La portada debe dar protagonismo al arte específico del producto (la portada
  del e-book cuando exista; si no, una ilustración de su motivo) y mostrar el
  nombre/motivo de identidad. No se reutilizan la misma imagen, composición
  temática o portada genérica en productos distintos.
- La portada de cada kit debe tener composición propia, no solo cambiar título,
  ilustración o colores dentro de una plantilla común. Debe integrar el arte
  como parte de una dirección de arte y mostrar una vista real del Excel.
- Los demás mockups conservan el branding y los colores propios del producto,
  pero no deben sustituir su tema por una plantilla visual indiferenciada.
- Los mockups comunes son materia prima, no diseños finales inmutables: se
  recortan, reorganizan, superponen o recomponen según el lenguaje del producto.
  Cada kit define una secuencia y una organización visual propias; no se copia
  el mismo orden ni la misma composición de tarjetas de otro producto activo.
- El generador integra el arte específico con capturas existentes del Excel y
  datos verificables del paquete. El arte puede ambientar o enmarcar, pero nunca
  reemplaza la captura real ni inventa pantallas, cifras o funciones.
- Antes de crear mockups, la etapa de marketing debe producir el arte específico
  del producto con generación de imagen y guardarlo en el paquete. Si falta arte,
  dirección o composición única, la etapa falla; no se degrada silenciosamente a
  una portada genérica. Al regenerar solo marketing se pueden reutilizar las
  capturas y el vídeo cuya huella ya conste en el manifiesto, sin reabrir el XLSX.
- La primera imagen debe ser una portada comercial del kit completo, no solo la
  cubierta del e-book: debe mostrar el producto/Excel real y enumerar entregables
  confirmados en el paquete.
- Si el ZIP incluye una herramienta HTML, la portada debe mostrar juntos y con
  peso visual comparable el Excel real y una captura real de la calculadora en
  el navegador. No relegar la herramienta a una lista pequeña de entregables.
  Añadir además un mockup dedicado con la interfaz real, nombre de la herramienta
  y la indicación de que funciona offline cuando ese dato esté validado.
- Generar mockups específicos para las dos ediciones de Excel (con datos y en
  blanco) y, cuando existan, para el e-book y los demás PDF incluidos. Usar la
  cubierta real del e-book dentro del mockup de recursos y no presentarla como
  si fuera la portada de todo el producto.
- Generar siempre un mockup dedicado de inventario completo, por ejemplo
  `03-whats-included-mockup.png`. Debe enumerar todos los archivos que recibe el
  cliente a partir del manifiesto, el paquete y el contenido del ZIP, sin omitir
  entregables por conveniencia ni anunciar recursos inexistentes. Debe mostrar
  cantidades por tipo (Excel, PDF, guía, etc.), el nombre de cada recurso y un
  icono reconocible; puede usar un diseño informativo genérico, pero sus formas,
  iconos y composición deben respetar la dirección de arte y la paleta del kit.
  Si no hay herramienta HTML, el inventario aparece inmediatamente después de la
  portada. Si sí la hay, poner primero la portada conjunta, después el mockup
  destacado de HTML y a continuación el inventario completo.
- Si el ZIP incluye una herramienta HTML, mostrarla en el inventario con un
  icono HTML y una etiqueta comprensible; no anunciarla como app online: es un
  archivo local que el comprador abre en su navegador y puede usar sin red.
- Los nombres y tipos de archivo anunciados deben derivarse del manifiesto y del
  ZIP final; no añadir recursos hipotéticos.
- Cada paquete publicitario incluye un vídeo demo del XLSX final. El demo navega
  por las hojas, cambia una entrada numérica declarada en el contrato y enseña
  la vista relacionada después del recálculo. Si no existe una entrada segura
  para demostrar, el manifiesto debe indicar esa limitación.
- Si el ZIP incluye una herramienta HTML, generar también un vídeo aparte que
  cargue el archivo entregado en el navegador, introduzca valores de ejemplo
  válidos, dispare la interacción real de la herramienta y enseñe el resultado
  actualizado. Debe ser distinto del vídeo de navegación del Excel y usar el
  idioma del producto. No simular resultados con pantallas inventadas.
- Registrar los dos vídeos en `video_files`, conservar `video_file` como el demo
  Excel por compatibilidad y registrar por separado `html_tool_video_file`, la
  captura de navegador y sus hashes. Validar duración, resolución y hashes de
  cada vídeo; al subirlos juntos a Etsy, habilitar el modo de vídeo múltiple
  para que la demo HTML no sustituya al vídeo del Excel.
- El vídeo y las capturas se generan sobre una copia de trabajo del XLSX; el
  archivo entregable no se modifica.
- El manifiesto debe registrar `image_files`, `screenshot_files`, `mockup_files`,
  `cover_image`, `video_file`, `video_files` y la huella del XLSX fuente. Cuando
  haya herramienta HTML, también su captura, su mockup y su vídeo independiente.
- Etsy admite hasta diez imágenes por anuncio: conservar la portada del kit como
  primera imagen y priorizar la calculadora HTML (si existe), el inventario,
  mockups de entregables y después capturas útiles dentro del límite.
- Los cargadores de Etsy leen los recursos desde el manifiesto, sin rutas ni
  textos fijos a `product-7`.
- Si el paquete incluye un e-book en PDF, renderizar su primera pÃ¡gina real y
  guardar la captura como asset verificable del paquete.
- Incluir esa captura en un mockup dedicado que destaque la portada del e-book;
  presentar los demÃ¡s PDF como recursos complementarios y no reemplazar la
  portada real por arte conceptual.
- La galerÃ­a debe dar visibilidad al conjunto del kit digital (Excel, e-book,
  guÃ­as y plantillas), no organizar todos los mensajes alrededor del Excel.
  La portada principal presenta el kit completo y el mockup dedicado vende
  explÃ­citamente el e-book.
- Los assets no son evidencia de QA ni pueden cambiar el XLSX. Si cambia el
  libro, primero se repite la QA y después se regeneran los assets.
