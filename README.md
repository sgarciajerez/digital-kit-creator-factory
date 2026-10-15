# Creador de Excels

Factoría local para investigar oportunidades, crear plantillas Excel nativas y
validarlas funcional y estructuralmente.

## Descubrimiento de nichos

El panel incluye un explorador que toma hipótesis del catálogo configurable o de consultas propias y contrasta anuncios activos de Etsy (señal de oferta, no demanda). La factoría no automatiza Marketplace Insights ni ofrece un flujo de transcripción manual. Google Ads OAuth queda como diagnóstico de conexión y no se usa para inferir demanda de Etsy. La configuración está en `rules/MARKET_DISCOVERY.md`.

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

## Marketing y publicación

La factoría ejecuta `MARKETING_ASSETS` y `MARKETING_QA` después de validar el
libro y antes de permitir subir el kit a Etsy. Captura el XLSX final en Excel,
crea mockups con su paleta —incluida una portada del kit y vistas de los archivos
realmente incluidos— y graba una demo interactiva sobre una copia temporal.
Las capturas, imágenes y el vídeo quedan declarados en
`outputs/listings/product-N/draft_package/listing.json`; Etsy publica usando
únicamente esas rutas.

Antes de cerrar el anuncio, la factoría también debe documentar la investigación
SEO de Etsy y usar palabras clave pertinentes en el título y las etiquetas.
Consulta `rules/FACTORY_ETSY_SEO.md`.

## Reglas de modelos

- Cabeceras semánticas y tablas nativas.
- Totales fuera del rango de datos.
- Porcentajes solo en ratios y progreso.
- Fechas, cantidades e importes con formatos propios.
- Fórmulas declaradas antes de construir el XLSX y probadas con casos reales.
- Quotes.Selected = Yes + Decision = Preferred se refleja como previsión por área; solo Purchases aumenta el gasto real.
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

La publicación Etsy permanece bloqueada por defecto en `.env`. La factoría solo
autopublica productos nuevos elegibles (desde el ID 15) y excepciones puntuales
configuradas en `config/FACTORY_PIPELINE.json`, siempre tras superar QA. Los
productos existentes no se autopublican.

## Calidad y entrega

El constructor nativo está en `scripts/native_project_builder.py`; el comportamiento
de la familia de proyectos, en `scripts/project_model.py`. Los casos funcionales
se ejecutan sobre copias independientes e incluyen crecimiento de 111 filas,
borrado, devolución, vaciado y fechas vencidas.

Cada aprobación contiene SHA-256 del XLSX, master, contrato y especificación.
Modificar cualquiera invalida la aprobación y obliga a regenerar y validar.
El kit digital incluye un libro de ejemplo, una copia `-blank.xlsx` validada,
`README.md` y su guía PDF, los complementos que seleccione el perfil de producto,
un índice de entrega y un ZIP verificado. `config/digital_kit_catalog.json`
declara las necesidades, reglas de selección, e-book opcional/obligatorio y
puertas de preparación comercial por producto; una necesidad sin módulo o un
perfil ausente bloquea el kit en vez de producir una entrega genérica.
Las filas nuevas se insertan dentro de las tablas, nunca debajo del total.

La inspección con openpyxl es de lectura: puede avisar de extensiones que no
interpreta. No se guardan los libros nativos con esa biblioteca; Excel conserva
sus objetos y validaciones durante las modificaciones.

## Alcance de la autonomía

La producción y la QA están automatizadas. La investigación recibe evidencias
mediante CSV y el proveedor de IA incluido es una simulación `mock`. El selector
de producto contrasta la evidencia, la novedad del catálogo, la diversidad por
categoría y las métricas de tienda que existan. No obtiene automáticamente
volúmenes de búsqueda de Etsy: resultados/listados son señales de competencia,
no prueba de demanda. Consulta `rules/PRODUCT_SELECTION.md` y
`config/product_selection.json`. Una hipótesis nunca se presenta como demanda
observada y las recomendaciones requieren revisión humana.
