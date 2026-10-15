# Autoevaluación de productos Excel

La factoría ejecuta `evaluate_excel_products.py` antes de generar el XLSX. La
nota es reproducible y se apoya en el contrato del master y en el modelo
funcional, no en imágenes.

## Nota y pesos

| Área | Peso | Qué comprueba |
|---|---:|---|
| Funcionalidades planteadas | 35 | Alcance, módulos, flujo y presencia del contrato funcional. |
| Diseño nativo | 30 | Master declarado, lineage, geometría de celdas, identidad SVG y componentes nativos. |
| Utilidad de la plantilla | 25 | Comprador, promesa, utilidad y flujo de trabajo. |
| Posibles mejoras | 10 | Mejoras futuras concretas y separadas del núcleo. |
| **Total** | **100** | Suma de las cuatro áreas. |

## Regla de paso

- `score >= 95`: `ready_for_app_generation`.
- `score < 95`: `needs_design_improvement`.

Además de la nota, se validan arquitectura, tablas, fórmulas, listas, SVG,
dimensiones y la guía `model_behavior.md`.

La revisión humana sigue siendo obligatoria antes de publicar, pero no depende
de capturas automatizadas.

```powershell
.venv\Scripts\python.exe -m scripts.validation.evaluate_excel_products --no-open
.venv\Scripts\python.exe -m scripts.run_autonomous_factory --no-open
```
