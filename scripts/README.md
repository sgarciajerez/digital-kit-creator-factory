# Scripts

- `core/`: configuración, base de datos, contratos y servicios compartidos.
- `etsy/`: conexión, anuncios, precios y publicación en Etsy.
- `excel/`: construcción y representación de libros Excel.
- `generation/`: generación de productos y kits digitales.
- `marketing/`: capturas, mockups, vídeos y paquetes de marketing.
- `research/`: exploración de mercado, selección de productos y diagnósticos OAuth.
- `validation/`: validadores y controles de calidad.
- `maintenance/`: configuración local, comprobaciones, copias de seguridad y lanzadores.

Los comandos Python se ejecutan como módulos desde la raíz del proyecto, por
ejemplo: `.venv\\Scripts\\python.exe -m scripts.run_autonomous_factory`.
`app.py` y el orquestador principal se conservan en la raíz de `scripts/` como
puntos de entrada estables.
