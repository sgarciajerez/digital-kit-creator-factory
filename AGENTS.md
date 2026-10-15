# Workflow de este repositorio

- Trabaja siempre en la rama `develop`; no hagas cambios ni commits directamente en `master`.
- Usa `master` como rama estable. Cuando `develop` esté listo, intégralo mediante un pull request con destino `master`.
- Mantén `changelog.md` actualizado en la sección `Unreleased`; indica explícitamente cuando un cambio no requiera entrada.
- Al fusionarse en `master` un pull request originado en `develop`, GitHub Actions crea una release con notas generadas por GitHub. La primera etiqueta será `v0.1.0`; las siguientes incrementarán la versión de parche.
- No incluyas secretos locales, bases de datos, `outputs/` ni instaladores en los commits.
