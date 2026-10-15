# Workflow de este repositorio

- Trabaja siempre en la rama `develop`; no hagas cambios ni commits directamente en `master`.
- Usa `master` como rama estable. Cuando `develop` esté listo, intégralo mediante un pull request con destino `master`.
- Mantén `changelog.md` actualizado en cada cambio, agrupando las notas en `Unreleased` hasta integrarlas en `master`.
- Mantén `VERSION` en la raíz como versión objetivo de la siguiente integración: usa SemVer (`MAJOR.MINOR.PATCH`), parte de `0.1.0`, aumenta PATCH para correcciones/refactors compatibles, MINOR para capacidades compatibles y MAJOR para cambios incompatibles. Actualiza `VERSION` junto con `changelog.md` en el mismo cambio.
- La integración en `master` publica una GitHub Release con la versión de `VERSION` y notas generadas desde GitHub. No integres cambios si la versión ya existe como etiqueta; antes de otro ciclo de trabajo, incrementa `VERSION` respecto a la última release.
- Los cambios se trabajan y se confirman en `develop`; integra `develop` en `master` mediante un pull request cuando esté listo. Si el usuario pide expresamente una fusión directa, debe ser un merge commit `--no-ff` de `develop`, nunca un commit de trabajo directo en `master`.
- No incluyas secretos locales, bases de datos, `outputs/` ni instaladores en los commits.
