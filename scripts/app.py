from __future__ import annotations

import streamlit as st

from autopilot import register_idea_candidates, run_autopilot
from budget import BudgetManager
from factory_core import PROJECT_ROOT, initialize_database, list_opportunities, load_settings, record_run, scalar


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
                    "Motivo": row["decision_reason"],
                }
                for row in opportunity_rows
            ],
            use_container_width=True,
            hide_index=True,
        )
    else:
        st.info("Todavía no hay oportunidades registradas.")

    st.subheader("Selección automática supervisada")
    st.caption(
        "Registra hipótesis, selecciona solo oportunidades aceptadas y prepara borradores locales. "
        "No publica en Etsy ni valida por sí sola la demanda, política o derechos."
    )
    candidate_col, autopilot_col = st.columns(2)
    with candidate_col:
        if st.button("Generar ideas candidatas"):
            candidate_ids = register_idea_candidates(settings.database_path)
            if candidate_ids:
                st.success(f"Registradas {len(candidate_ids)} ideas como hipótesis para revisar.")
            else:
                st.info("No se han creado ideas nuevas; ya existían candidatas equivalentes.")
    with autopilot_col:
        if st.button("Preparar productos aceptados"):
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
                    "No hay oportunidades aceptadas con puntuación ≥75 y confianza ≥60, "
                    "o ya tienen un producto generado."
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

