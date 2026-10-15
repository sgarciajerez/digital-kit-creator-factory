# Mejoras de factoría

Registro de problemas que afectan a Excel, los modelos, las tablas, las
fórmulas, las dimensiones o la operación. Las capturas quedan reservadas a la
fase publicitaria posterior y nunca participan en la QA del XLSX.

## Mejoras aplicadas

### 001 - Lote activo explícito

- **Problema:** el pipeline apuntaba a productos antiguos que ya no estaban en
  `outputs`.
- **Solución:** `FACTORY_PIPELINE.json` limita el lote activo y separa modelos
  recuperados de productos en construcción.
- **Regla:** no usar carpetas históricas como inventario implícito.

### 002 - Sin paneles inmovilizados

- **Problema:** `freeze_panes` generaba líneas horizontales o verticales que
  atravesaban la hoja.
- **Solución:** los constructores compartidos y el contrato nativo fijan
  `freeze_panes = None`.
- **Regla:** comprobarlo en la auditoría directa del XLSX.

### 003 - Inicialización idempotente

- **Problema:** el diseño intentaba escribir antes de crear la carpeta del
  producto.
- **Solución:** cada fase crea sus carpetas mínimas y puede reintentarse.
- **Regla:** ningún script depende de una fase previa para crear directorios.

### 004 - No reescribir el master con OpenXML

- **Problema:** guardar la copia maestra con `openpyxl` dañaba validaciones y
  objetos nativos antes de abrirla con Excel.
- **Solución:** el constructor nativo usa Excel COM para editar y guardar la
  copia; `openpyxl` queda para inspección y contratos.
- **Regla:** cualquier modelo con objetos nativos se guarda mediante Excel.

### 005 - Adaptación explícita de tipos COM

- **Problema:** Excel COM rechazaba objetos `date` de Python.
- **Solución:** el puente convierte fechas a tipos compatibles y conserva el
  formato de fecha de la celda.
- **Regla:** adaptar siempre los tipos antes de cruzar una API nativa.

### 006 - Rangos de tabla explícitos

- **Problema:** Excel ampliaba `ListObject` mediante `CurrentRegion`, generaba
  cabeceras genéricas y dejaba fórmulas incompletas.
- **Solución:** cada tabla se redimensiona al rango declarado y se reescriben
  cabeceras, datos y fórmulas después de la normalización.
- **Regla:** validar nombre, cabeceras, rango y primera fila calculada.

### 007 - Totales fuera del ListObject

- **Problema:** escribir una fila TOTAL junto a una tabla podía volver a
  incluirla en el rango filtrable.
- **Solución:** cerrar la tabla con `Resize` después de escribir los totales y
  verificar de nuevo el rango.
- **Regla:** los totales nunca forman parte de los datos de entrada.

### 008 - Anchura basada en valores visibles

- **Problema:** medir el texto de una fórmula producía falsos avisos de ancho.
- **Solución:** la auditoría mide cabeceras y resultados visibles, no la cadena
  fuente de la fórmula.
- **Regla:** las dimensiones se evalúan desde la estructura del XLSX.

### 009 - Herencia nativa declarada

- **Problema:** los modelos derivados no declaraban qué base ni qué fuentes
  justificaban su encuadre.
- **Solución:** se añade `templates/model-library/index.json`, `model_base`,
  `master_workbook` y `source_lineage` al contrato de diseño.
- **Regla:** un producto sin lineage nativo no entra en generación.

### 010 - Alineación nativa explícita

- **Problema:** Excel guardaba `General`, haciendo depender la lectura del tipo
  de dato o de la configuración regional.
- **Solución:** el escritor fija la alineación de las celdas visibles y ajusta
  anchuras mínimas de fechas, conteos e importes.
- **Regla:** no depender de `General` para una tabla aprobada.

### 011 - Modelos históricos recuperables

- **Problema:** al limpiar `outputs`, los productos anteriores dejaron de estar
  disponibles como referencias reutilizables.
- **Solución:** se regeneraron y guardaron como modelos nativos en
  `templates/model-library/`.
- **Regla:** los productos retirados de listings pasan a la biblioteca de
  modelos o a `archive` si siguen siendo útiles.

### 012 - Master único para diseño

- **Problema:** el flujo había introducido un archivo intermedio por producto
  para representar el diseño.
- **Solución:** la única referencia visual y estructural vuelve a ser
  `templates/excel-first-base/master_workbook.xlsx`.
- **Regla:** los productos se generan directamente desde el master, sources,
  modelo funcional y contrato; no se crean previews alternativos.

### 013 - Pricing ausente como bloqueo explícito

- **Problema:** un producto nuevo sin comparables en
  `competition_pricing.json` provocaba un `KeyError` al preparar el borrador.
- **Solución:** la preparación del borrador registra
  `blocked_by_missing_pricing` y conserva el motivo sin interrumpir el resto
  de la auditoría del XLSX.
- **Regla:** no inventar precios; añadir comparables observados antes de
  intentar crear el paquete comercial.

### 014 - Capturas solo para publicidad final

- **Problema:** mezclar capturas de interfaz con la validación podía hacer que
  una imagen sustituyera al contrato real del XLSX.
- **Solución:** la QA sigue siendo directa sobre el libro; las capturas solo
  se permiten después, como assets de listing en `marketing_assets/`.
- **Regla:** una captura nunca aprueba ni corrige el Excel y cualquier cambio
  posterior en el libro obliga a repetir QA y regenerar la publicidad.

### 015 - Vídeo demo funcional para publicidad

- **Problema:** una imagen estática no demuestra que el comprador pueda usar
  realmente la plantilla.
- **Solución:** el paquete publicitario debe incluir un vídeo grabado sobre una
  copia del XLSX validado, con navegación, entrada de datos y actualización de
  al menos una fórmula o resultado.
- **Regla:** el vídeo sirve para explicar y vender el producto, nunca para
  sustituir las pruebas funcionales o estructurales.

### 016 - Capturas originales y encuadre publicitario

- **Problema:** un mockup podía ocultar el Excel real o mostrar márgenes,
  bordes blancos y elementos ajenos al producto.
- **Solución:** el paquete conserva capturas originales del XLSX final tomadas
  desde el propio proyecto y las usa como fuente de mockups y portada.
- **Regla:** las capturas deben estar recortadas al área útil de Excel y
  corresponder a la misma versión que pasó la QA.

## Criterio de cierre

El producto no se considera terminado hasta que su XLSX, contrato funcional,
guía `model_behavior.md`, lineage de sources, dimensiones, tablas, fórmulas y
validaciones pasan sin errores.
