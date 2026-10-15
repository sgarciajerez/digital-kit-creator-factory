# Operativa de fábrica y publicación

La fábrica crea y verifica primero el XLSX. La autopublicación queda limitada
a productos nuevos elegibles: ID 15 en adelante, más las excepciones puntuales
que figuren en `auto_publish_policy.one_time_product_ids`. Los productos 1 y 2
y el resto de anuncios existentes no se publican automáticamente. La ficha solo
se activa después de superar todas las puertas de calidad y subir el ZIP, las
imágenes y el vídeo. Las herramientas Etsy ejecutadas manualmente conservan su
bloqueo global por defecto.

## Paquete de producto

Cada producto debe contener:

- `app_design_spec.md`;
- `design_spec.json`;
- `model_contract.json`;
- `model_behavior.md`;
- `identity_assets/*.svg`;
- `app/*.xlsx`;
- informes de auditoría funcional y estructural.

## Puerta de calidad

El XLSX solo avanza si pasan:

1. la autoevaluación del producto;
2. la arquitectura y el lineage de sources;
3. la auditoría de tablas y fórmulas;
4. el preflight nativo de dimensiones, navegación, iconos y tablas;
5. los casos declarados en `model_behavior.md`.

La validación lee el libro directamente. No se hacen capturas para aprobar la
interfaz ni se usan imágenes como criterio de QA.

## Assets publicitarios posteriores

Una vez superadas todas las puertas de calidad, se generan las capturas,
mockups, portada y vídeo promocional desde el XLSX final. Las imágenes
principales deben ser capturas originales del Excel, tomadas desde el propio
proyecto y recortadas al área útil, sin bordes blancos, navegador, escritorio,
cursores ni elementos ajenos. Solo se pueden reutilizar capturas del proyecto
que correspondan a la misma versión validada del libro.

Cada paquete debe incluir un vídeo demo que pruebe el libro: introducir datos
de ejemplo, navegar por las hojas y mostrar una actualización real de una
fórmula o resultado. Estos archivos deben guardarse en `marketing_assets/`, no
se mezclan con los informes de QA y no pueden cambiar el XLSX ya validado. Si
el libro cambia, los assets deben regenerarse después de repetir la validación.

## Publicación

El paquete de listing contiene el XLSX, la ficha, el precio y el texto. Antes
de cerrar el título y las etiquetas, la factoría debe investigar términos de
búsqueda pertinentes y guardar fuentes, fecha, métricas observadas y decisiones
en `seo_keyword_research.json`. Puede usar expresiones como `Digital Kit` si
describen el paquete y la investigación las respalda; no debe rellenar el título
con palabras clave repetidas. Se aplican los límites vigentes de Etsy y se
revisa coherencia entre título, etiquetas, descripción y archivos.

La política de autopublicación se configura en
`config/FACTORY_PIPELINE.json`. La excepción del producto 12 se consume tras
una publicación satisfactoria. Para los productos nuevos (ID 15 o superior), la
factoría habilita las escrituras de Etsy solo en sus subprocesos finales; no
modifica `.env`, que conserva `ETSY_DRY_RUN=true` y
`ETSY_PUBLISH_ENABLED=false`. Un fallo de QA impide subir o activar el anuncio.

## Comandos de fábrica

```powershell
.venv\Scripts\python.exe -m scripts.generation.create_next_product_designs
.venv\Scripts\python.exe -m scripts.validation.validate_design_architecture
.venv\Scripts\python.exe -m scripts.validation.evaluate_excel_products --no-open
.venv\Scripts\python.exe -m scripts.generation.generate_autonomous_apps --no-open
.venv\Scripts\python.exe -m scripts.validation.audit_design_contract --product N
.venv\Scripts\python.exe -m scripts.validation.preflight_visual_contract --product N
.venv\Scripts\python.exe -m scripts.run_autonomous_factory --product N --no-open
```
