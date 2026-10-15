# Estructura de salidas

```text
templates/
├── sources/                         sources estudiados
├── model-library/                   modelos XLSX recuperados
└── excel-first-base/
    ├── master_workbook.xlsx         única base visual y estructural
    ├── native_base_contract.json
    └── template_manifest.json

outputs/listings/product-N/
├── app_design_spec.md               ficha y contrato editorial
├── design_spec.json                 contrato de arquitectura
├── model_contract.json              contrato de tablas y fórmulas
├── model_behavior.md                guía independiente de uso y QA
├── identity_assets/*.svg            identidad del producto
├── app/*.xlsx                       libro final
├── app/README.md                    instrucciones de uso en el idioma del producto
├── marketing_assets/                capturas originales, mockups y vídeo demo
├── design_contract_report.json      auditoría semántica
├── native_xlsx_preflight_report.json auditoría estructural nativa
└── app_generation_report.json       resultado de generación
```

`marketing_assets/` solo se crea después de que el XLSX haya pasado las
puertas funcional y estructural. Sus archivos, incluido el vídeo demo, no son
evidencias de QA ni pueden modificar el libro validado.
