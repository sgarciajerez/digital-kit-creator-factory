# Changelog

Registro de cambios relevantes de la factoría. Añade una entrada en `Unreleased` para cada pull request; las notas de cada release también se generan automáticamente desde los pull requests fusionados.

## Unreleased

- Se incorpora el producto 17, un kit en inglés de mantenimiento y costes del vehículo, con identidad de taller propia, e-book con fuentes, hojas imprimibles y calculadora HTML de coste por milla. La investigación de Etsy se marca como preliminar y el precio final depende de comparables activos.
- Se amplía la factoría con la operación offline `vehicle_cost_per_mile` y una composición de marketing exclusiva tipo ficha de taller/registro de servicio.
- Se excluyen los recursos voluminosos de `assets/` del control de versiones sin borrar las copias locales.
- Se organizan los scripts y las reglas por responsabilidad, con índices de navegación y rutas actualizadas.
- Se retiran del repositorio los documentos temporales `QUALITY_UPGRADE.md` y `mejoras de factoria.md`.
- Se documenta `master_workbook.xlsx` como plantilla candidata y se incorpora `VERSION` para alinear el versionado SemVer, el changelog y las releases.
- Se amplía el flujo Git: cada cambio debe actualizar `changelog.md` y la versión objetivo de la siguiente integración.
- Se establece `develop` como rama de trabajo y `master` como rama estable.
- Se automatiza la creación de releases versionadas al integrar cambios de `develop` en `master`.
