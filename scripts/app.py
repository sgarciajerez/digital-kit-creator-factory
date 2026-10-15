from __future__ import annotations

import streamlit as st

from autopilot import run_autopilot
from budget import BudgetManager
from factory_core import PROJECT_ROOT, create_opportunity, initialize_database, list_opportunities, load_settings, record_run, scalar
from product_selection import recommend_products
from market_discovery import discovery_catalog, load_discovery_config, run_market_discovery, unscanned_catalog
from etsy_client import load_env as load_etsy_env


def main() -> None:
    st.set_page_config(page_title="Fábrica de productos digitales", page_icon="🧪", layout="wide")

    settings = load_settings()
    initialize_database(settings.database_path)
    budget = BudgetManager(
        settings.database_path,
        monthly_limit_eur=settings.monthly_budget_eur,
        daily_limit_eur=settings.daily_budget_eur,
        task_limit_eur=settings.task_budget_eur,
        product_limit_eur=settings.product_budget_eur,
        batch_limit_eur=settings.batch_budget_eur,
    )
    if "startup_recorded" not in st.session_state:
        record_run(settings.database_path, source="streamlit")
        st.session_state.startup_recorded = True

    st.title("🧪 Fábrica de productos digitales con IA")
    st.caption("MVP local · producción supervisada · publicación bajo aprobación humana")

    budget_snapshot = budget.snapshot()
    expenses = budget_snapshot.spent_eur
    remaining = budget_snapshot.remaining_eur
    runs = int(scalar(settings.database_path, "SELECT COUNT(*) FROM app_runs"))
    opportunities = int(scalar(settings.database_path, "SELECT COUNT(*) FROM opportunities"))

    first, second, third, fourth = st.columns(4)
    first.metric("Gasto IA este mes", f"{expenses:.2f} €")
    second.metric("Presupuesto restante", f"{remaining:.2f} €")
    third.metric("Ejecuciones", runs)
    fourth.metric("Oportunidades", opportunities)

    st.divider()
    st.subheader("Oportunidades")
    opportunity_rows = list_opportunities(settings.database_path)
    if opportunity_rows:
        st.dataframe(
            [
                {
                    "Nombre": row["name"],
                    "Puntuación": row["score"],
                    "Confianza": row["confidence"],
                    "Decisión": row["decision"],
                    "Investigación Etsy": "Sí" if row.get("market_research_json") else "No",
                    "Motivo": row["decision_reason"],
                }
                for row in opportunity_rows
            ],
            use_container_width=True,
            hide_index=True,
        )
    else:
        st.info("Todavía no hay oportunidades registradas.")

    st.subheader("Explorador automático de nichos")
    st.caption(
        "Explora un catálogo inicial configurable de temas o tus propias ideas y contrástalos con anuncios "
        "digitales activos de Etsy. La API mide oferta y competencia observada, no búsquedas ni ventas. "
        "Los resultados son hipótesis y no crean ni publican productos."
    )
    discovery_config = load_discovery_config()
    country_codes = list(discovery_config["markets"])
    country = st.selectbox(
        "Mercado",
        country_codes,
        index=country_codes.index(discovery_config["default_market"]),
        format_func=lambda code: discovery_config["markets"][code]["label"],
    )
    language_codes = list(discovery_config["languages"])
    language = st.selectbox(
        "Idioma de búsqueda",
        language_codes,
        index=language_codes.index(discovery_config["default_language"]),
        format_func=lambda code: discovery_config["languages"][code]["label"],
    )
    with st.form("market_discovery_form"):
        source_mode = st.radio(
            "Cómo proponer temas",
            ["Catálogo inicial", "Mis propias ideas"],
            horizontal=True,
        )
        if source_mode == "Catálogo inicial":
            catalog = discovery_catalog(language, discovery_config)
            fresh_catalog = unscanned_catalog(language, settings.database_path, discovery_config)
            catalog_labels = {item["query"]: item["label"] for item in catalog}
            if not fresh_catalog:
                st.caption("Ya se han escaneado todos los temas del catálogo para este idioma; puedes repetirlos o añadir consultas propias.")
            selected_topics = st.multiselect(
                "Temas candidatos (máximo ocho por barrido)",
                options=list(catalog_labels),
                default=[item["query"] for item in fresh_catalog[:int(discovery_config["etsy_queries_per_scan"])]],
                format_func=lambda query: catalog_labels[query],
            )
            seed_text = ""
        else:
            selected_topics = []
            seed_text = st.text_area(
                "Consultas semilla (una por línea)",
                placeholder="ej. aquarium setup\npet adoption checklist\nsmall-space gardening",
            )
        submitted_discovery = st.form_submit_button("Explorar nichos")

    if submitted_discovery:
        seeds = selected_topics if source_mode == "Catálogo inicial" else [line.strip() for line in seed_text.splitlines() if line.strip()]
        try:
            discovery = run_market_discovery(
                settings.database_path,
                seeds,
                load_etsy_env(),
                shop_id=load_etsy_env().get("ETSY_SHOP_ID"),
                market=country,
                language=language,
                report_directory=PROJECT_ROOT / "outputs" / "research",
            )
            st.session_state["market_discovery_report"] = discovery
            st.success("Exploración guardada como evidencia de mercado; las ideas siguen pendientes de validación.")
        except Exception as error:
            st.error(str(error))

    discovery = st.session_state.get("market_discovery_report")
    if discovery:
        st.markdown("**Contraste de oferta en Etsy**")
        st.dataframe(
            [
                {
                    "Consulta": item["query"],
                    "Anuncios activos": item["total_active_results"],
                    "Descargas digitales en muestra": item["sampled_digital_competitors"],
                    "Mediana de precio de muestra": item["digital_price_median"],
                }
                for item in discovery["etsy_supply_scan"]["results"]
            ],
            use_container_width=True,
            hide_index=True,
        )
        clusters = discovery["etsy_supply_scan"]["supply_phrase_clusters"]
        if clusters:
            st.markdown("**Frases repetidas en títulos · hipótesis de oferta, no demanda**")
            st.dataframe(
                [{"Frase": item["phrase"], "Anuncios de apoyo": item["listing_support"], "Precio mediano": item["median_price"]} for item in clusters],
                use_container_width=True,
                hide_index=True,
            )
        with st.expander("Convertir un tema en hipótesis para investigar"):
            candidate_options = list(dict.fromkeys(
                [item["phrase"] for item in clusters]
                + [item["query"] for item in discovery["etsy_supply_scan"]["results"]]
            ))
            candidate = st.selectbox("Tema candidato", candidate_options)
            idea_audience = st.text_input("¿Quién compraría un kit útil sobre esto?")
            idea_problem = st.text_area("¿Qué problema concreto tendría?")
            idea_solution = st.text_input("¿Qué tipo de kit podría resolverlo?")
            if st.button("Guardar esta idea como hipótesis", key="save_discovered_idea"):
                if not idea_audience.strip() or not idea_problem.strip() or not idea_solution.strip():
                    st.error("Completa comprador, problema y solución; la oferta observada no prueba demanda.")
                else:
                    idea_id = create_opportunity(
                        settings.database_path,
                        name=candidate.title(),
                        category="discovery",
                        subcategory=candidate,
                        target_audience=idea_audience,
                        problem=idea_problem,
                        proposed_solution=idea_solution,
                        decision="review",
                        decision_reason=(
                            f"Hipótesis basada en exploración Etsy {discovery['research_run_id']}; "
                            "la evidencia recopilada describe oferta, no demanda del comprador."
                        ),
                    )
                    st.success(f"Hipótesis {idea_id} guardada. Reúne evidencia de demanda antes de puntuarla o producirla.")
        st.caption(f"Informe persistido: {discovery.get('report_path', 'consulta actual')}")

    st.subheader("Selección automática supervisada")
    st.caption(
        "Contrasta oportunidades aceptadas con evidencia reciente, novedad frente al catálogo y resultados "
        "de tienda disponibles. La recomendación no publica ni sustituye la revisión humana."
    )
    with st.expander("Registrar una hipótesis para investigar"):
        with st.form("new_opportunity_form", clear_on_submit=True):
            name = st.text_input("Idea / necesidad concreta")
            category = st.text_input("Categoría de mercado")
            subcategory = st.text_input("Subcategoría")
            audience = st.text_input("Público objetivo")
            problem = st.text_area("Problema que resolvería")
            proposed_solution = st.text_area("Solución o tipo de kit que imaginas")
            submitted = st.form_submit_button("Guardar como hipótesis")
            if submitted:
                if not name.strip() or not problem.strip():
                    st.error("Indica al menos la idea y el problema que resolvería.")
                else:
                    opportunity_id = create_opportunity(
                        settings.database_path,
                        name=name,
                        category=category or None,
                        subcategory=subcategory or None,
                        target_audience=audience or None,
                        problem=problem,
                        proposed_solution=proposed_solution or None,
                        decision="review",
                        decision_reason="Hipótesis pendiente de investigación; no es evidencia de demanda.",
                    )
                    st.success(f"Hipótesis {opportunity_id} guardada para investigar.")

    if st.button("Analizar oportunidades y recomendar próximos kits"):
        selection_path = PROJECT_ROOT / "outputs" / "reports" / "product_selection_report.json"
        selection = recommend_products(settings.database_path, report_path=selection_path)
        if selection.selected_opportunity_ids:
            st.success(f"Recomendaciones: {', '.join(map(str, selection.selected_opportunity_ids))}")
        else:
            st.info("No hay oportunidades que superen todos los filtros de evidencia, puntuación y diferenciación.")
        if selection.candidates:
            st.dataframe(
                [
                    {
                        "ID": row["opportunity_id"],
                        "Idea": row["name"],
                        "Categoría": row["category"],
                        "Puntuación": row["selection_score"],
                        "Confianza": row["confidence"],
                        "Solapamiento": row["catalog_similarity"],
                        "Evidencia": row["evidence"]["evidence_count"],
                        "Estado": row["status"],
                        "Motivos": " ".join(row["reasons"]),
                    }
                    for row in selection.candidates
                ],
                use_container_width=True,
                hide_index=True,
            )
        st.caption(f"Informe detallado: {selection_path}")

    if st.button("Preparar borradores locales compatibles"):
        report = run_autopilot(
            settings.database_path,
            PROJECT_ROOT / "outputs" / "products",
            PROJECT_ROOT / "outputs" / "listings",
            max_products=3,
        )
        if report.product_ids:
            st.success(
                f"Preparados {len(report.product_ids)} productos y "
                f"{len(report.listing_ids)} paquetes locales para revisión."
            )
        else:
            st.info(
                "No hay recomendaciones elegibles con evidencia suficiente, o las oportunidades "
                "requieren un tipo de kit distinto al generador heredado de presupuestos."
            )
        st.caption(f"Informe: {report.report_path}")

    st.divider()
    left, right = st.columns(2)
    with left:
        st.subheader("Estado del sistema")
        st.success("Base de datos inicializada")
        st.info(f"Modo de autonomía: {settings.autonomy_mode} — producción supervisada")
        st.caption(
            f"Presupuesto IA: {budget_snapshot.percentage:.1f}% utilizado "
            f"({budget_snapshot.spent_eur:.2f} € de {budget_snapshot.limit_eur:.2f} €)"
        )
        if budget_snapshot.alerts:
            st.warning(f"Alertas de presupuesto alcanzadas: {', '.join(map(str, budget_snapshot.alerts))}%")
        if settings.ai_provider == "mock":
            st.info("Proveedor de IA: mock; no realiza llamadas de pago")
        elif settings.ai_api_key_configured:
            st.success(f"Proveedor de IA configurado: {settings.ai_provider}")
        else:
            st.warning(f"Proveedor {settings.ai_provider} pendiente: falta AI_API_KEY")
    with right:
        st.subheader("Módulos del MVP")
        for label in ("Persistencia SQLite", "Registro de auditoría", "Control básico de presupuesto"):
            st.write(f"✅ {label}")
        for label in ("Investigación y puntuación", "Generación y validación", "Importación de métricas"):
            st.write(f"⏳ {label}")


if __name__ == "__main__":
    main()

