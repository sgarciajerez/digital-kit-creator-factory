# Checklist de creación de productos Excel

## Idea y fuentes

- [ ] Definir usuario, problema y resultado útil.
- [ ] Revisar métricas y oportunidades antes de elegir el producto.
- [ ] Inventariar `templates/sources/`.
- [ ] Elegir un modelo de `templates/model-library/`.
- [ ] Confirmar que la base visual es `templates/excel-first-base/master_workbook.xlsx`.

## Modelo y diseño

- [ ] Definir tablas con columnas semánticas y rangos claros.
- [ ] Definir fórmulas, listas, formatos y casos de prueba.
- [ ] Escribir `model_behavior.md` desde el comportamiento esperado.
- [ ] Crear `model_contract.json` independiente del XLSX.
- [ ] Declarar paleta, SVG, navegación, dimensiones y variaciones.
- [ ] Generar el XLSX final clonando el master.

## QA directo del XLSX

- [ ] Ejecutar `scripts/validate_design_base.py`.
- [ ] Ejecutar `scripts/audit_design_contract.py --product N`.
- [ ] Ejecutar `scripts/preflight_visual_contract.py --product N`.
- [ ] Ejecutar la validación funcional del modelo.
- [ ] Añadir registros de ejemplo y comprobar que los totales cambian con
      sentido.
- [ ] Confirmar que las columnas de porcentaje son ratios y no importes o
      conteos.
- [ ] Confirmar que no hay líneas, bordes o paneles fuera del contrato.

## Publicidad posterior a la validación

- [ ] Generar capturas o mockups solo después de superar toda la QA del XLSX.
- [ ] Guardarlos en `outputs/listings/product-N/marketing_assets/`.
- [ ] Incluir capturas originales del Excel tomadas desde el propio proyecto.
- [ ] Recortar las capturas al área útil, sin bordes blancos ni elementos
      ajenos al Excel.
- [ ] Comprobar que las capturas pertenecen a la misma versión validada.
- [ ] Conservar los originales en `marketing_assets/source_captures/`.
- [ ] Definir una dirección de arte de marketing propia: concepto, composición,
      medio visual y elementos que se deben evitar.
- [ ] Crear arte original generado para el producto; no basar la identidad solo
      en nombre, motivo breve, colores o una plantilla común.
- [ ] Elegir una composición y una secuencia de galería que no copie otro kit
      activo; recomponer los mockups comunes para que encajen en esa identidad.
- [ ] Si el kit tiene herramienta HTML, mostrar en la portada una composición
      clara con la captura real del Excel y la calculadora del navegador juntas.
- [ ] Crear un mockup destacado con la calculadora real; ordenar la galería como
      portada, herramienta HTML, inventario completo y después los demás recursos.
- [ ] Integrar capturas reales del libro sin inventar pantallas, cifras o
      prestaciones; reutilizarlas en una regeneración solo de marketing.
- [ ] Incluir un mockup de inventario que enumere todos los entregables reales,
      con cantidad por tipo, nombre de archivo e icono, coherente con el diseño.
- [ ] Confirmar que todo el texto visible de la galería usa el idioma declarado
      para el Excel; conservar sin traducir los nombres reales de archivo.
- [ ] Crear un vídeo demo usando una copia del XLSX validado.
- [ ] Mostrar en el vídeo navegación, entrada de datos y una actualización
      real de fórmula o resultado.
- [ ] No usar esos archivos para aprobar fórmulas, dimensiones o estructura.
- [ ] No modificar el XLSX validado durante la generación publicitaria.

## E-book y equilibrio del kit digital

- [ ] Si hay e-book en PDF, renderizar su primera pÃ¡gina real y guardar la
      captura como asset verificable.
- [ ] Crear un mockup dedicado que destaque esa portada real y muestre los
      demÃ¡s PDF como recursos complementarios.
- [ ] Revisar que la galerÃ­a presenta el kit completo (Excel, e-book, guÃ­as y
      plantillas), no solo distintas variaciones del Excel.

## Herramienta HTML del kit

- [ ] Evaluar necesidades declaradas y elegir la calculadora/herramienta HTML
      con mayor coincidencia; no añadir una utilidad genÃ©rica solo por cumplir.
- [ ] Mantener el `.html` autocontenido, adaptable a mÃ³vil, en el idioma y
      paleta del producto, sin CDNs, llamadas de red ni almacenamiento de datos.
- [ ] Explicar entradas, validaciones y fÃ³rmula; probar resultados y casos
      lÃ­mite con una prueba independiente del XLSX.
- [ ] Incluir el archivo en el ZIP y manifiesto con hash, Ã­ndice e inventario
      visual; describirlo en la ficha solo si realmente estÃ¡ empaquetado.

- [ ] Abrir el mismo HTML del ZIP en un navegador, capturar la herramienta y
      generar un vídeo demostrativo independiente que introduzca valores y enseñe
      el resultado calculado. Subir ambos vídeos en modo múltiple al publicar en
      Etsy para que ninguno sustituya al otro.

## SEO y ficha del anuncio Etsy

- [ ] Investigar palabras clave por idioma y mercado usando Marketplace
      Insights cuando esté disponible, además de sugerencias y anuncios
      comparables relevantes.
- [ ] Guardar `seo_keyword_research.json` con fecha, fuentes, términos, datos
      observados, intención y decisiones de uso/descarte.
- [ ] Construir un título conciso que nombre el producto y sus rasgos
      principales; incluir `Digital Kit` o el formato digital solo si describe
      fielmente el paquete y la investigación lo respalda.
- [ ] Completar las 13 etiquetas con frases distintas, relevantes y de 20 caracteres como
      máximo; cubrir búsquedas específicas, formato, usos y sinónimos.
- [ ] Verificar límites y guía Etsy vigentes, y comprobar que título, etiquetas,
      descripción y archivos del kit cuentan la misma historia.
