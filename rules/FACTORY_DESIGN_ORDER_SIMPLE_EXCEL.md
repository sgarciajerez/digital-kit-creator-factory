# Orden de diseño para productos Excel nativos

Estas reglas sustituyen cualquier flujo anterior basado en HTML o referencias
de imagen.

## 1. Elegir fuentes y modelo

1. Revisar los libros de `templates/sources/` relevantes.
2. Estudiar hojas, dimensiones, tablas, fórmulas, listas y estilos útiles.
3. Elegir un modelo de `templates/model-library/`.
4. Usar siempre `templates/excel-first-base/master_workbook.xlsx` como base
   visual y estructural.
5. Registrar `master_workbook`, `model_base` y `source_lineage` en la ficha.

## 2. Construir el contrato

Antes del XLSX se declaran hojas, orden, columnas, tipos, formatos, fórmulas,
listas, dimensiones, anclajes, zoom, navegación, identidad SVG, paleta y
casos de verificación funcional. `model_behavior.md` debe ser independiente
del libro y describir lo que la plantilla debería hacer.

## 3. Crear el XLSX

Se clona el master y se cambian únicamente los slots declarados. No se crea un
preview alternativo: el master es la referencia que permanece en la factoría.
No se crean HTML, CSS, PNG, capturas ni carpetas de comparación.

## 4. Variaciones permitidas

- títulos, subtítulos y textos de ayuda;
- nombres y propósito de hojas;
- una paleta coherente por producto;
- SVG propio o adaptado;
- cabeceras y columnas con significado equivalente;
- datos de ejemplo, fórmulas y rangos justificados;
- listas y estados;
- hasta un gráfico nativo por pantalla cuando ayude a decidir.

## 5. QA

- validar contrato y master;
- generar el XLSX final desde el master;
- ejecutar casos funcionales en una copia de datos;
- auditar tablas, fórmulas, formatos, dimensiones, iconos, botones y líneas
  mediante el propio XLSX;
- guardar los bloqueos de lógica en `mejoras de factoria.md`.
