# Native Excel design contract

```json
{
  "master_workbook": "templates/excel-first-base/master_workbook.xlsx",
  "native_reference_mode": "master_workbook_contract",
  "source_of_truth": "master_workbook.xlsx + model_contract.json + model_behavior.md"
}
```

The master workbook defines the grid, dimensions, navigation, table anchors and
allowed components. A product may change semantic text, columns, formulas,
palette, icon and example data only inside declared variation slots.

## Required checks

- Sheet names and order match the product brief.
- Columns have explicit names and the right meaning.
- Tables have complete ranges, unique headers and totals outside data ranges.
- Percentage formats are used only for ratio fields.
- Formulas point to the table or column they represent.
- Validation lists accept only compatible values.
- Row and column dimensions preserve the master's composition.
- Navigation shapes and the identity SVG exist on every worksheet.
- `model_behavior.md` states the expected behaviour and verification cases.

Validation reads the XLSX directly. There are no screenshots, pixel comparisons,
HTML files or browser dependencies.
