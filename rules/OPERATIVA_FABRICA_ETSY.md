# Operativa de fábrica y publicación

La fábrica crea y verifica primero el XLSX. La publicación se mantiene
desactivada hasta una revisión humana independiente.

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

El paquete de listing contiene el XLSX, la ficha, el precio y el texto. La
activación está deshabilitada por defecto. Antes de publicar, una persona debe
abrir el XLSX, confirmar que el contenido tiene sentido y aprobar manualmente
la ficha y los datos comerciales.

## Comandos de fábrica

```powershell
.venv\Scripts\python.exe scripts/create_next_product_designs.py
.venv\Scripts\python.exe scripts/validate_design_architecture.py
.venv\Scripts\python.exe scripts/evaluate_excel_products.py --no-open
.venv\Scripts\python.exe scripts/generate_autonomous_apps.py --no-open
.venv\Scripts\python.exe scripts/audit_design_contract.py --product N
.venv\Scripts\python.exe scripts/preflight_visual_contract.py --product N
.venv\Scripts\python.exe scripts/run_autonomous_factory.py --product N --no-open
```
