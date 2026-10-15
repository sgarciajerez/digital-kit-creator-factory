# Creador de Excels

Factoría local para investigar oportunidades, crear plantillas Excel nativas y
validarlas funcional y estructuralmente.

## Fuentes y modelos

- `templates/sources/`: ejemplos de Excel para estudiar estructura,
  dimensiones, fórmulas y distribución.
- `templates/excel-first-base/master_workbook.xlsx`: única base visual y
  estructural compartida.
- `templates/model-library/`: modelos recuperados de productos anteriores.
- `outputs/listings/`: productos activos en proceso de creación.

Los productos nuevos se crean clonando el master y aplicando el modelo
funcional elegido. El master y el XLSX final son los únicos artefactos de
diseño y validación estructural. Después de superar la QA se generan los
assets publicitarios dentro de `marketing_assets/`, incluido un vídeo demo
del XLSX en uso.

La regla completa está en `rules/FACTORY_NATIVE_SOURCE_WORKFLOW.md`.

## Reglas de modelos

- Cabeceras semánticas y tablas nativas.
- Totales fuera del rango de datos.
- Porcentajes solo en ratios y progreso.
- Fechas, cantidades e importes con formatos propios.
- Fórmulas declaradas antes de construir el XLSX y probadas con casos reales.
- Guía normativa `model_behavior.md` y contrato `model_contract.json` en cada
  base y producto.
- Una paleta coherente y un SVG propio o adaptado en todas las hojas.
- Paneles congelados solo cuando se declaran por hoja en `freeze_panes`; sin líneas grises no declaradas.

## Flujo principal

Requiere Windows, Python 3.12 y Microsoft Excel de escritorio. Ejecuta
`scripts/setup_windows.ps1` y comprueba el entorno con
`.venv\Scripts\python.exe scripts/check_environment.py --excel`.
Las dependencias Python se declaran en `requirements.txt`; Excel se instala aparte.
`constraints-windows-py312.txt` fija las versiones del entorno comprobado.

```powershell
.venv\Scripts\python.exe scripts/create_next_product_designs.py
.venv\Scripts\python.exe scripts/validate_design_architecture.py
.venv\Scripts\python.exe scripts/generate_autonomous_apps.py --no-open
.venv\Scripts\python.exe scripts/audit_design_contract.py --product N
.venv\Scripts\python.exe scripts/preflight_visual_contract.py --product N
.venv\Scripts\python.exe -m pytest -q
```

La publicación Etsy permanece desactivada por defecto y requiere aprobación
humana.

## Calidad y entrega

El constructor nativo está en `scripts/native_project_builder.py`; el comportamiento
de la familia de proyectos, en `scripts/project_model.py`. Los casos funcionales
se ejecutan sobre copias independientes e incluyen crecimiento de 111 filas,
borrado, devolución, vaciado y fechas vencidas.

Cada aprobación contiene SHA-256 del XLSX, master, contrato y especificación.
Modificar cualquiera invalida la aprobación y obliga a regenerar y validar.
El paquete incluye un libro de ejemplo, una copia `-blank.xlsx` validada y una guía.
Las filas nuevas se insertan dentro de las tablas, nunca debajo del total.

La inspección con openpyxl es de lectura: puede avisar de extensiones que no
interpreta. No se guardan los libros nativos con esa biblioteca; Excel conserva
sus objetos y validaciones durante las modificaciones.

## Alcance de la autonomía

La producción y la QA están automatizadas. La investigación actual recibe
evidencias mediante CSV y el proveedor de IA incluido es una simulación `mock`.
No se obtiene demanda real ni se hacen llamadas de pago automáticamente.
Integrar una fuente externa exige elegir una fuente y configurar sus credenciales;
una hipótesis nunca se presenta como demanda observada.
