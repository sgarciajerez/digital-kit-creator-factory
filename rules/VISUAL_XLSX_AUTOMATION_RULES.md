# Reglas de automatización visual XLSX

La apariencia se implementa en celdas, dimensiones, estilos, formas SVG y
gráficos nativos del libro. El master_workbook es la fuente visual; no se
utilizan HTML, navegador, PNG ni capturas de pantalla para decidir si el
producto está bien.

## Verificación directa

- Comprobar el orden y nombre de las pestañas.
- Comprobar anchos de columnas, alturas de filas, celdas combinadas y ausencia
  de paneles inmovilizados.
- Comprobar la posición y tamaño de las tablas nativas.
- Comprobar encabezados, formatos numéricos, rellenos, bordes, alineaciones y
  colores de la paleta.
- Comprobar que las formas de navegación y el SVG de identidad están en todas
  las hojas.
- Comprobar que los gráficos usan rangos respaldados por fórmulas.
- Comprobar fórmulas y casos funcionales desde `model_behavior.md`.

## Regla de reparación

Si falla una comprobación, se repara el constructor o el XLSX. Nunca se
modifica una imagen de referencia para ocultar el fallo.

## Estilo de texto de grÃ¡ficos

Los tÃ­tulos y tÃ­tulos de ejes usan un relleno degradado claro tomado del
fondo de la paleta del producto hacia blanco, con texto negro y sin bordes ni
contorno tipogrÃ¡fico. Leyendas, etiquetas y valores usan texto negro sin
contorno.

## Comandos

```powershell
.venv\Scripts\python.exe scripts/validate_design_base.py
.venv\Scripts\python.exe scripts/audit_design_contract.py --product N
.venv\Scripts\python.exe scripts/preflight_visual_contract.py --product N
.venv\Scripts\python.exe scripts/preflight_visual_contract.py --product N
```

