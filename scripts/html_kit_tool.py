"""Select and render small, offline HTML helpers for digital kits."""
from __future__ import annotations

import html
import json
import re
from pathlib import Path


ROOT = Path(__file__).resolve().parent.parent
SUPPORTED_OPERATIONS = {"moving_budget", "renovation_contingency", "garden_spacing", "aquarium_volume", "puppy_first_year"}

COPY = {
    "en": {
        "eyebrow": "A QUICK, PRACTICAL TOOL", "open_note": "Works offline in your browser. Your entries stay on this device and are not saved or sent anywhere.",
        "calculate": "Update estimate", "result": "Your estimate", "reset": "Reset values", "currency": "Amounts in your chosen currency",
        "invalid": "Enter valid, non-negative numbers to see your estimate.",
        "titles": {"moving_budget": "Moving Budget Buffer Calculator", "renovation_contingency": "Renovation Contingency Calculator", "garden_spacing": "Garden Bed Planting Calculator", "aquarium_volume": "Freshwater Tank Volume Calculator", "puppy_first_year": "Puppy First-Year Cost Calculator"},
        "moving": {"intro": "Give your moving budget a little breathing room before the boxes arrive.", "labels": ["Available moving budget", "Estimated moving costs", "Contingency reserve (%)"], "result_labels": ["Suggested reserve", "Budget including reserve", "Left after this estimate"], "formula": "Reserve = estimated costs × reserve rate. Remaining budget = available budget − estimated costs − reserve."},
        "renovation": {"intro": "Check whether your contingency is proportionate to the current project estimate.", "labels": ["Current project estimate", "Contingency rate (%)", "Reserve already set aside"], "result_labels": ["Suggested contingency", "Additional amount to set aside", "Project estimate with contingency"], "formula": "Suggested contingency = estimate × rate. Additional amount = max(0, suggested contingency − reserve already set aside)."},
        "garden": {"intro": "Turn bed dimensions and plant spacing into a quick starting-point count.", "labels": ["Bed length (m)", "Bed width (m)", "Spacing between plants (cm)"], "result_labels": ["Approximate plants per row", "Approximate rows", "Approximate plants in the bed"], "formula": "Plants per row = floor(length ÷ spacing) + 1. Rows = floor(width ÷ spacing) + 1. Total = plants per row × rows."},
        "aquarium": {"intro": "Estimate a rectangular freshwater tank's water volume from its inside measurements.", "labels": ["Inside length (cm)", "Inside width (cm)", "Water depth (cm)", "Estimated fill factor (%)"], "result_labels": ["Estimated water volume (litres)", "Approximate US gallons", "Gross volume before fill factor (litres)"], "formula": "Gross litres = inside length × inside width × water depth ÷ 1,000. Estimate = gross litres × fill factor. This does not account for substrate, hardscape or equipment displacement."},
        "puppy": {"intro": "Build a first-year estimate from your own local quotes and monthly figures; no prices or care schedule are assumed.", "labels": ["Adoption or acquisition fee", "Initial setup costs", "Food per month", "Routine care per month", "Insurance per month"], "result_labels": ["Estimated first-year total", "Average monthly equivalent", "Up-front costs", "Recurring costs across 12 months"], "formula": "First-year estimate = adoption fee + setup costs + 12 × (monthly food + routine care + insurance). Use local quotes and your veterinarian's advice; this is a budgeting aid, not a care or insurance recommendation."},
    },
    "es": {
        "eyebrow": "HERRAMIENTA PRÁCTICA", "open_note": "Funciona sin conexión en tu navegador. Los datos se quedan en este dispositivo: no se guardan ni se envían.",
        "calculate": "Actualizar cálculo", "result": "Tu estimación", "reset": "Restablecer valores", "currency": "Importes en la moneda que elijas",
        "invalid": "Introduce números válidos y no negativos para ver el cálculo.",
        "titles": {"moving_budget": "Calculadora del presupuesto de mudanza", "renovation_contingency": "Calculadora de imprevistos de reforma", "garden_spacing": "Calculadora de plantación en bancal", "aquarium_volume": "Calculadora del volumen de un acuario de agua dulce", "puppy_first_year": "Calculadora del primer año con un cachorro"},
        "moving": {"intro": "Deja un pequeño margen en el presupuesto de la mudanza antes de que lleguen las cajas.", "labels": ["Presupuesto disponible para la mudanza", "Costes estimados de la mudanza", "Reserva para imprevistos (%)"], "result_labels": ["Reserva sugerida", "Presupuesto con reserva", "Restante tras esta estimación"], "formula": "Reserva = costes estimados × porcentaje de reserva. Presupuesto restante = presupuesto disponible − costes estimados − reserva."},
        "renovation": {"intro": "Comprueba si la reserva para imprevistos guarda proporción con la estimación actual de la obra.", "labels": ["Estimación actual del proyecto", "Porcentaje para imprevistos (%)", "Reserva ya apartada"], "result_labels": ["Reserva sugerida", "Importe adicional que conviene apartar", "Estimación del proyecto con reserva"], "formula": "Reserva sugerida = estimación × porcentaje. Importe adicional = max(0, reserva sugerida − reserva ya apartada)."},
        "garden": {"intro": "Convierte las medidas del bancal y la separación entre plantas en un recuento orientativo.", "labels": ["Largo del bancal (m)", "Ancho del bancal (m)", "Separación entre plantas (cm)"], "result_labels": ["Plantas aproximadas por fila", "Filas aproximadas", "Plantas aproximadas en el bancal"], "formula": "Plantas por fila = floor(largo ÷ separación) + 1. Filas = floor(ancho ÷ separación) + 1. Total = plantas por fila × filas."},
        "aquarium": {"intro": "Estima el volumen de agua de un acuario rectangular con sus medidas interiores.", "labels": ["Largo interior (cm)", "Ancho interior (cm)", "Profundidad del agua (cm)", "Factor de llenado estimado (%)"], "result_labels": ["Volumen estimado (litros)", "Galones estadounidenses aproximados", "Volumen bruto antes del factor (litros)"], "formula": "Litros brutos = largo interior × ancho interior × profundidad del agua ÷ 1.000. Estimación = litros brutos × factor de llenado. No descuenta el desplazamiento del sustrato, rocas ni equipo."},
        "puppy": {"intro": "Prepara una estimación del primer año con tus precios locales y gastos mensuales, sin asumir precios ni pautas de cuidado.", "labels": ["Adopción o adquisición", "Gastos iniciales de preparación", "Alimentación mensual", "Cuidados habituales al mes", "Seguro mensual"], "result_labels": ["Estimación total del primer año", "Promedio mensual equivalente", "Gastos iniciales", "Gastos recurrentes de 12 meses"], "formula": "Estimación anual = adopción + preparación + 12 × (alimentación mensual + cuidados habituales + seguro). Usa presupuestos locales y consulta al veterinario; es una ayuda presupuestaria, no una recomendación veterinaria ni de seguros."},
    },
}

FIELD_IDS = {
    "moving_budget": ["budget", "costs", "rate"],
    "renovation_contingency": ["estimate", "rate", "reserve"],
    "garden_spacing": ["length", "width", "spacing"],
    "aquarium_volume": ["length", "width", "depth", "fill"],
    "puppy_first_year": ["adoption", "setup", "food", "care", "insurance"],
}


def select_html_tool(product: dict, candidates: list[dict]) -> dict | None:
    """Select the most relevant declared tool by exact overlap with product needs."""
    needs = set(product.get("needs", []))
    ranked = []
    for candidate in candidates:
        matched = sorted(needs.intersection(candidate.get("triggers", [])))
        if matched and candidate.get("operation") in SUPPORTED_OPERATIONS:
            ranked.append((len(matched), int(candidate.get("priority", 0)), candidate, matched))
    if not ranked:
        return None
    _, _, candidate, matched = max(ranked, key=lambda item: (item[0], item[1], str(item[2].get("id", ""))))
    return {
        "id": candidate["id"],
        "title": candidate["title"],
        "filename": candidate["filename"],
        "operation": candidate["operation"],
        "matched_needs": matched,
        "reason": "Selected because it best matches declared product needs: " + ", ".join(matched),
    }


def generate_html_tool(spec: dict, output_path: Path, tool: dict) -> dict:
    """Generate one single-file, responsive calculator with no network dependencies."""
    operation = str(tool.get("operation", ""))
    if operation not in SUPPORTED_OPERATIONS:
        raise ValueError(f"Unsupported digital-kit HTML operation: {operation}")
    filename = str(tool.get("filename", ""))
    if Path(filename).name != filename or Path(filename).suffix.casefold() != ".html":
        raise ValueError("HTML tool filename must be a plain .html filename")

    language = "es" if str(spec.get("design_language") or spec.get("language") or "en").casefold().startswith("es") else "en"
    copy = COPY[language]
    tool_copy = copy["moving" if operation == "moving_budget" else "renovation" if operation == "renovation_contingency" else "garden" if operation == "garden_spacing" else "aquarium" if operation == "aquarium_volume" else "puppy"]
    tool_title = copy["titles"][operation]
    if language == "en":
        tool_title = str(tool.get("title", tool_title))
    palette = spec.get("palette") or {}

    def color(value: str, fallback: str) -> str:
        raw = str(value or fallback).strip()
        return raw if re.fullmatch(r"#[0-9a-fA-F]{6}", raw) else fallback

    accent = color(palette.get("accent"), "#2F7F79")
    secondary = color(palette.get("secondary"), "#E68A70")
    background = color(palette.get("background"), "#F4F2EA")
    ink = color(palette.get("ink"), "#202A33")
    title = str(spec.get("title", "Digital kit"))
    fields = []
    for field_id, label in zip(FIELD_IDS[operation], tool_copy["labels"]):
        initial = "10" if field_id in {"rate", "spacing"} else "0"
        minimum = "0.1" if field_id == "spacing" else "0"
        maximum = ' max="100"' if field_id == "rate" else ""
        if operation == "garden_spacing" and field_id == "spacing":
            initial = "30"
        if operation == "aquarium_volume":
            initial = {"length": "60", "width": "30", "depth": "28", "fill": "85"}[field_id]
            minimum = "0.1"
            maximum = ' max="100"' if field_id == "fill" else ""
        if operation == "puppy_first_year":
            initial = {"adoption": "450", "setup": "500", "food": "70", "care": "55", "insurance": "35"}[field_id]
        fields.append(
            f'<label for="{field_id}">{html.escape(label)}'
            f'<input id="{field_id}" name="{field_id}" type="number" inputmode="decimal" min="{minimum}"{maximum} step="any" value="{initial}" required></label>'
        )
    result_cards = "\n".join(
        f'<article class="result-card"><span>{html.escape(label)}</span><strong id="result-{index}">—</strong></article>'
        for index, label in enumerate(tool_copy["result_labels"])
    )
    config = {
        "operation": operation,
        "invalid": copy["invalid"],
        "locale": "es-ES" if language == "es" else "en-US",
        "currencyLabel": copy["currency"],
    }
    config_json = json.dumps(config, ensure_ascii=False).replace("<", "\\u003c").replace(">", "\\u003e").replace("&", "\\u0026")
    page = f'''<!doctype html>
<html lang="{language}">
<head>
<meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<meta http-equiv="Content-Security-Policy" content="default-src 'none'; style-src 'unsafe-inline'; script-src 'unsafe-inline'; form-action 'none'; base-uri 'none'">
<title>{html.escape(tool_title)} | {html.escape(title)}</title>
<style>
:root{{--accent:{accent};--secondary:{secondary};--paper:{background};--ink:{ink};--card:#fff;--muted:#667078}}
*{{box-sizing:border-box}}body{{margin:0;background:var(--paper);color:var(--ink);font:16px/1.55 system-ui,-apple-system,"Segoe UI",sans-serif}}
main{{width:min(920px,calc(100% - 32px));margin:34px auto 56px}}.topline{{height:8px;background:linear-gradient(90deg,var(--accent),var(--secondary));border-radius:9px;margin-bottom:25px}}
.eyebrow{{color:var(--accent);font-size:.78rem;font-weight:800;letter-spacing:.12em;text-transform:uppercase}}h1{{font-size:clamp(2rem,5vw,3.2rem);line-height:1.08;margin:.45rem 0 .7rem;letter-spacing:-.04em}}.intro{{font-size:1.08rem;max-width:66ch;color:#35414a}}
.privacy{{padding:12px 15px;border-left:4px solid var(--secondary);background:#ffffffa8;border-radius:0 8px 8px 0;font-size:.91rem}}
.layout{{display:grid;grid-template-columns:minmax(0,1fr) minmax(260px,.8fr);gap:20px;margin-top:25px}}section{{background:var(--card);border:1px solid #00000012;border-radius:16px;padding:22px;box-shadow:0 10px 30px #16202a0d}}h2{{font-size:1.15rem;margin:0 0 16px}}
form{{display:grid;gap:14px}}label{{display:grid;gap:6px;font-weight:650;font-size:.94rem}}input{{width:100%;border:1px solid #aeb6b8;border-radius:9px;padding:12px;font:inherit;color:var(--ink);background:#fff}}input:focus{{outline:3px solid color-mix(in srgb,var(--accent) 24%,transparent);border-color:var(--accent)}}
.actions{{display:flex;gap:10px;flex-wrap:wrap;margin-top:5px}}button{{border:0;border-radius:9px;padding:12px 16px;font:700 .94rem system-ui;cursor:pointer}}button.primary{{background:var(--accent);color:#fff}}button.secondary{{background:#edf0ed;color:var(--ink)}}.results{{display:grid;gap:10px}}.result-card{{padding:14px;border-radius:11px;background:var(--paper);border-bottom:3px solid var(--secondary)}}.result-card span{{display:block;color:var(--muted);font-size:.84rem}}.result-card strong{{display:block;font-size:1.5rem;margin-top:3px}}.formula{{font-size:.88rem;color:#48545a;margin:15px 0 0}}.error{{color:#9f3026;min-height:1.4em;font-size:.9rem}}
footer{{color:var(--muted);font-size:.83rem;margin-top:20px}}@media(max-width:700px){{main{{margin-top:20px}}.layout{{grid-template-columns:1fr}}section{{padding:18px}}}}
</style></head><body><main><div class="topline"></div>
<div class="eyebrow">{html.escape(copy["eyebrow"])}</div><h1>{html.escape(tool_title)}</h1>
<p class="intro">{html.escape(tool_copy["intro"])}</p><p class="privacy">{html.escape(copy["open_note"])}</p>
<div class="layout"><section><h2>{html.escape(copy["calculate"])}</h2><form id="tool-form" novalidate>{''.join(fields)}
<div class="actions"><button class="primary" type="submit">{html.escape(copy["calculate"])}</button><button class="secondary" id="reset" type="reset">{html.escape(copy["reset"])}</button></div><div id="error" class="error" role="status" aria-live="polite"></div></form></section>
<section><h2>{html.escape(copy["result"])}</h2><div class="results">{result_cards}</div><p class="formula">{html.escape(tool_copy["formula"])}</p></section></div>
<footer>{html.escape(copy["currency"]) if operation != "garden_spacing" else html.escape(title)} · {html.escape(str(spec.get("title", "Digital kit")))}</footer></main>
<script>
const cfg={config_json};const form=document.getElementById('tool-form');const errorBox=document.getElementById('error');
const read=(id)=>Number(document.getElementById(id).value);const put=(index,value,integer=false)=>{{document.getElementById('result-'+index).textContent=integer?String(Math.max(0,Math.floor(value))):new Intl.NumberFormat(cfg.locale,{{maximumFractionDigits:2}}).format(Math.max(0,value));}};
function calculate(event){{if(event)event.preventDefault();const values=[...form.querySelectorAll('input')].map(input=>input.value.trim()===''?NaN:Number(input.value));if(values.some(value=>!Number.isFinite(value)||value<0)||((cfg.operation==='moving_budget'||cfg.operation==='renovation_contingency')&&values[2]>100)){{errorBox.textContent=cfg.invalid;return;}}errorBox.textContent='';
if(cfg.operation==='moving_budget'){{const [budget,costs,rate]=values;const buffer=costs*rate/100;put(0,buffer);put(1,costs+buffer);const left=budget-costs-buffer;document.getElementById('result-2').textContent=new Intl.NumberFormat(cfg.locale,{{maximumFractionDigits:2}}).format(left);}}
else if(cfg.operation==='renovation_contingency'){{const [estimate,rate,reserve]=values;const target=estimate*rate/100;put(0,target);put(1,Math.max(0,target-reserve));put(2,estimate+target);}}
else if(cfg.operation==='aquarium_volume'){{const [length,width,depth,fill]=values;if(length<=0||width<=0||depth<=0||fill<=0||fill>100){{errorBox.textContent=cfg.invalid;return;}}const gross=length*width*depth/1000;const litres=gross*fill/100;put(0,litres);put(1,litres*0.219969);put(2,gross);}}
else if(cfg.operation==='puppy_first_year'){{const [adoption,setup,food,care,insurance]=values;const upfront=adoption+setup;const recurring=12*(food+care+insurance);const total=upfront+recurring;put(0,total);put(1,total/12);put(2,upfront);put(3,recurring);}}
else{{const [length,width,spacing]=values;if(spacing<=0){{errorBox.textContent=cfg.invalid;return;}}const perRow=Math.floor(length*100/spacing)+1;const rows=Math.floor(width*100/spacing)+1;put(0,perRow,true);put(1,rows,true);put(2,perRow*rows,true);}}
}}
form.addEventListener('input',()=>calculate());form.addEventListener('submit',calculate);document.getElementById('reset').addEventListener('click',()=>setTimeout(()=>calculate(),0));calculate();
</script></body></html>'''
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(page, encoding="utf-8")
    content = output_path.read_text(encoding="utf-8")
    if re.search(r"<(?:script|link)[^>]+(?:src|href)=[\"']https?://", content, re.IGNORECASE):
        output_path.unlink(missing_ok=True)
        raise ValueError("Generated HTML must not depend on external network resources")
    return {
        "passed": True,
        "file": str(output_path),
        "name": output_path.name,
        "title": tool_title,
        "operation": operation,
        "language": language,
        "standalone": True,
        "offline": True,
        "network_dependencies": [],
        "formula": tool_copy["formula"],
        "size_bytes": output_path.stat().st_size,
    }
