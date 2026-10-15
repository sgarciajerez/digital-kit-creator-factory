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
- [ ] Crear un vídeo demo usando una copia del XLSX validado.
- [ ] Mostrar en el vídeo navegación, entrada de datos y una actualización
      real de fórmula o resultado.
- [ ] No usar esos archivos para aprobar fórmulas, dimensiones o estructura.
- [ ] No modificar el XLSX validado durante la generación publicitaria.
