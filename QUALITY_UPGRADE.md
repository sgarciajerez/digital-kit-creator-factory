# Revisión de calidad — 15 de octubre de 2026

## Cambios aplicados

- Aprobaciones vinculadas al SHA-256 de libro, master, contrato normativo y especificación. Preflight y empaquetado rechazan aprobaciones obsoletas o QA no ejecutada.
- Casos funcionales independientes: cada caso vuelve a abrir una copia sin cambios guardados. No se admite aprobar un contrato sin casos ni resultados esperados.
- Familia de proyectos: seis casos por producto, incluyendo 111 compras añadidas, eliminación, devolución, vaciado de importes y fecha vencida.
- Totales de compras y su imputación al presupuesto usan referencias estructuradas. El usuario debe insertar filas dentro de la tabla; escribir fuera de ella no constituye una compra registrada.
- Copia limpia validada y guía de uso, además del ejemplo. El empaquetado verifica también la huella de la copia limpia.
- Constructor nativo separado en native_project_builder.py y código heredado en legacy_workbook_builders.py. Comportamiento de familia en project_model.py.
- Columnas de jardín adaptadas: Equipment, Soil, Workshops, Planning y Location. Formatos monetarios nativos en EUR.
- Paneles congelados permitidos si coinciden con la declaración por hoja. Gráficos declarados en la ruta nativa se crean mediante Excel, preservando extensiones.
- Dependencias directas completas, restricciones con versiones verificadas y comprobación de Excel en instalación.
- Comando --no-open documentado compatible; selección de producto propagada al empaquetado.
- Las observaciones de precios ya no se recortan al aplicar el máximo comercial al precio sugerido.
- Directorio temporal de pruebas independiente, sin utilizar el enlace pytest-current de otro usuario.

## Verificación realizada

- Entorno: Python 3.12.10, Excel 16.0; comprobación satisfactoria.
- Suite completa con RUN_EXCEL_TESTS=1: 38 pruebas aprobadas.
- Dos pruebas de escritorio verifican aislamiento de casos, detección de errores Excel, gráficos nativos y paneles congelados.
- Productos 1 y 14 regenerados: seis casos funcionales aprobados por producto, auditoría semántica y preflight aprobados.
- Copias limpias: tres indicadores de Dashboard comprobados a cero.
- Producto 14: paquete local regenerado correctamente. No se ha publicado nada.
- Copias de los productos anteriores: backups/quality-upgrade-20261015-085312/.

## Límites y trabajo pendiente

- La investigación sigue usando CSV y el proveedor IA sigue siendo mock. Elegir proveedor, fuente, credenciales y condiciones de uso es necesario para integrar investigación externa; no se han inventado observaciones ni activado llamadas de pago.
- El producto 1 no tiene comparables de precio en competition_pricing.json. Su XLSX está validado, pero su paquete comercial permanece bloqueado por esa falta de datos.
- La suite de escritorio emite un diagnóstico nativo RPC_DISCONNECTED (0x80010108) al liberar Excel bajo pytest, aun terminando con código 0 y pruebas aprobadas. La inicialización COM explícita no eliminó el diagnóstico. Requiere seguimiento del entorno pywin32/Excel; no se considera resuelto.
- openpyxl avisa de una extensión que no interpreta al leer el master. Las modificaciones de la ruta nativa se realizan con Excel; no se ha eliminado ni ocultado esta advertencia.
- La geometría del master continúa siendo el punto de partida; las categorías presupuestarias mantienen las siete ranuras del modelo. El crecimiento de compras está comprobado; no se afirma crecimiento ilimitado de todas las familias y gráficos.
- Las reglas estructurales no sustituyen una evaluación humana de legibilidad y utilidad comercial. No se ha realizado una revisión visual humana.

## Repetir las pruebas

```powershell
.venv\Scripts\python.exe scripts/check_environment.py --excel
$env:RUN_EXCEL_TESTS='1'
.venv\Scripts\python.exe -m pytest -q -p no:cacheprovider
.venv\Scripts\python.exe scripts/generate_autonomous_apps.py --product 1 --no-open
.venv\Scripts\python.exe scripts/audit_design_contract.py --product 1
.venv\Scripts\python.exe scripts/preflight_visual_contract.py --product 1
```
