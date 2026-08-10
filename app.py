
from __future__ import annotations

import io
from pathlib import Path

import pandas as pd
import streamlit as st

from ccs_engine import (
    CATEGORIES,
    economic_analysis,
    economic_gate,
    parse_reference,
    screen_fields,
    sensitivity_oat,
)

st.set_page_config(
    page_title="CCS Screening & Investment Decision Framework",
    page_icon="🌍",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    .main-title {font-size: 2.0rem; font-weight: 700; margin-bottom: 0.2rem;}
    .sub-title {color: #555; margin-bottom: 1.2rem;}
    .stage-box {padding: 0.8rem 1rem; border-radius: 0.6rem; background: #f2f7f9; border: 1px solid #d5e1e5;}
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown('<div class="main-title">CCS Screening & Investment Decision Framework</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="sub-title">Two-phase decision-support tool: Phase 1 geological/site screening → Phase 2 economic gate, economics and sensitivity analysis.</div>',
    unsafe_allow_html=True,
)


def init_state():
    defaults = {
        "phase1_ranked": pd.DataFrame(),
        "phase1_failed": pd.DataFrame(),
        "phase1_details": {},
        "reference_meta": {},
        "phase2_gate_passed": pd.DataFrame(),
        "phase2_gate_failed": pd.DataFrame(),
        "economic_df": pd.DataFrame(),
        "economic_raw": {},
        "sensitivity_df": pd.DataFrame(),
        "sensitivity_scenarios": {},
        "phase2_stage": 1,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


init_state()


def reset_phase2():
    for k in [
        "phase2_gate_passed",
        "phase2_gate_failed",
        "economic_df",
        "economic_raw",
        "sensitivity_df",
        "sensitivity_scenarios",
    ]:
        if isinstance(st.session_state[k], pd.DataFrame):
            st.session_state[k] = pd.DataFrame()
        else:
            st.session_state[k] = {}
    st.session_state["phase2_stage"] = 1


def export_xlsx(sheets: dict[str, pd.DataFrame]) -> bytes:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        for name, df in sheets.items():
            safe = name[:31]
            df.to_excel(writer, sheet_name=safe, index=False)
    return buffer.getvalue()


def show_reasons(df: pd.DataFrame, title: str):
    st.subheader(title)
    if df.empty:
        st.success("No eliminated fields at this stage.")
        return
    for _, row in df.iterrows():
        st.markdown(f"**{row['Field']}**")
        reason = str(row.get("Reason", "")).strip()
        for line in reason.splitlines():
            line = line.lstrip("• ").strip()
            if line:
                st.markdown(f"- {line}")


with st.sidebar:
    st.header("Framework Navigation")
    phase = st.radio(
        "Select phase",
        ["Phase 1 — Screening", "Phase 2 — Investment Economics"],
        index=0 if st.session_state.phase1_ranked.empty else 1,
    )

    st.divider()
    st.caption("Calculation conventions")
    st.caption(
        "• Numeric Phase-1 hard cut-off defaults to strict > unless the reference cell explicitly contains an operator such as >= or <=."
    )
    st.caption(
        "• Qualitative matching is case-insensitive; after the hard-cutoff gate, the engine searches all five SAW reference text/sentence cells."
    )
    st.caption(
        "• Numerical inputs must also fall inside a reference SAW range after passing the hard-cutoff gate."
    )
    st.caption(
        "• Blank, N/A, '-', and other unavailable/invalid inputs are treated as no usable value."
    )
    st.caption(
        "• Phase-1 ranking is Rank 1 = best; higher SAW score is better."
    )
    st.caption(
        "• Economic ranking equally aggregates ranks of NPV (higher better), IRR (higher better), and payback (lower better)."
    )
    st.caption(
        "• Sensitivity uses one-at-a-time low/high endpoint cases and ranks fields by robustness."
    )

# -------------------------- PHASE 1 --------------------------
if phase == "Phase 1 — Screening":
    st.header("Phase / Module 1 — Field Screening")

    with st.expander("How Phase 1 works", expanded=False):
        st.write(
            "Upload one active reference workbook/CSV and one or more field data files. "
            "The engine reads hard cut-offs, sub-parameter AHP weights, SAW ranges, and overall AHP weights from the reference data. "
            "A field must pass every hard cut-off and have a matching SAW score to enter the ranked board."
        )

    st.subheader("Step 1 — Reference Data")
    st.info(
        "You may upload multiple reference datasheets. Select exactly one as the active reference dataset for this screening run. "
        "The framework uses only the selected active reference dataset."
    )

    ref_files = st.file_uploader(
        "Upload reference datasheet(s)",
        type=["xlsx", "csv"],
        accept_multiple_files=True,
        key="reference_uploader",
        help="XLSX is recommended because the supplied template contains separate sheets for Overall and the five screening categories.",
    )

    if ref_files:
        ref_names = [f.name for f in ref_files]
        active_ref_name = st.selectbox(
            "Active reference datasheet",
            ref_names,
            key="active_reference_name",
        )
        active_ref = next(f for f in ref_files if f.name == active_ref_name)
        try:
            ref_bytes = active_ref.getvalue()
            reference = parse_reference(ref_bytes, active_ref.name)
            st.session_state["reference_meta"] = {
                "name": active_ref.name,
                "reference": reference,
            }
            st.success(
                f"Reference dataset loaded: {active_ref.name}. "
                f"{len(reference['categories']['Technical']) + len(reference['categories']['Environmental']) + len(reference['categories']['Regulatory']) + len(reference['categories']['Long Term Operation']) + len(reference['categories']['Risk'])} sub-parameters detected."
            )
            with st.expander("Show detected reference weights"):
                st.write("Overall AHP weights")
                st.dataframe(
                    pd.DataFrame(
                        [
                            {"Parameter": k, "Weight": v}
                            for k, v in reference["overall_weights"].items()
                        ]
                    ),
                    use_container_width=True,
                    hide_index=True,
                )
        except Exception as exc:
            st.error(f"Reference datasheet validation failed: {exc}")
            st.session_state["reference_meta"] = {}

    st.button(
        "➕ Create another reference datasheet",
        help="The uploader accepts multiple reference datasets. Only the selected active reference dataset is used for this run.",
    )

    st.subheader("Step 2 — Field Data")
    field_files = st.file_uploader(
        "Upload field datasheet(s)",
        type=["xlsx", "csv"],
        accept_multiple_files=True,
        key="field_uploader",
        help="Each XLSX should contain Technical, Environmental, Regulatory, Long Term Operation, Risk and Economics sheets.",
    )

    if field_files:
        st.write("Detected field files:")
        st.dataframe(
            pd.DataFrame({"Field name used by the framework": [f.name for f in field_files]}),
            use_container_width=True,
            hide_index=True,
        )

    if st.button(
        "🚀 Start the Screening Process",
        type="primary",
        disabled=not (ref_files and field_files and st.session_state.get("reference_meta")),
    ):
        with st.spinner("Running hard cut-off screening, SAW scoring and AHP aggregation..."):
            try:
                active_reference = st.session_state["reference_meta"]["reference"]
                payloads = [(f.name, f.getvalue()) for f in field_files]
                ranked, failed, details = screen_fields(active_reference, payloads)
                st.session_state["phase1_ranked"] = ranked
                st.session_state["phase1_failed"] = failed
                st.session_state["phase1_details"] = details
                st.session_state["phase2_stage"] = 1
                reset_phase2()
                st.success("Phase 1 screening completed.")
            except Exception as exc:
                st.error(f"Screening could not be completed: {exc}")

    if st.session_state.phase1_details:
        with st.expander("🔎 View value/text normalization and reference matching diagnostics", expanded=False):
            st.caption(
                "The engine first converts field inputs into canonical numeric values or canonical reference text. "
                "This diagnostic table shows exactly what was read from the field file and what reference value/range/text was selected."
            )
            diagnostic_rows = []
            for field_name, field_result in st.session_state.phase1_details.items():
                for category, records in field_result["Parameter Details"].items():
                    for rec in records:
                        diagnostic_rows.append({
                            "Field": field_name,
                            "Category": category,
                            "Parameter": rec["Parameter"],
                            "Raw Field Value": rec["Raw Field Value"],
                            "Field Unit": rec["Field Unit"],
                            "Normalized / Canonical Value": rec["Normalized / Canonical Value"],
                            "Normalization": rec["Normalization"],
                            "Hard Cut-Off": rec["Hard Cut-Off"],
                            "SAW Score": rec["SAW Score"],
                            "Matched SAW Reference": rec["Matched SAW Reference"],
                            "SAW Match Similarity (%)": rec["SAW Match Similarity (%)"],
                            "Status": rec["Status"],
                        })
            if diagnostic_rows:
                st.dataframe(pd.DataFrame(diagnostic_rows), use_container_width=True, hide_index=True)

    if not st.session_state.phase1_ranked.empty or not st.session_state.phase1_failed.empty:
        st.divider()
        st.header("Phase 1 Results")

        if not st.session_state.phase1_ranked.empty:
            st.subheader("Qualified Fields — Rank 1 = Best")
            st.dataframe(
                st.session_state.phase1_ranked,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Overall SAW Score": st.column_config.NumberColumn(format="%.4f"),
                },
            )
            st.download_button(
                "Download Phase 1 qualified ranking (CSV)",
                st.session_state.phase1_ranked.to_csv(index=False).encode("utf-8"),
                "phase1_qualified_ranking.csv",
                "text/csv",
            )
        else:
            st.warning("No field passed the complete Phase 1 screening gate.")

        show_reasons(st.session_state.phase1_failed, "Eliminated / Failed Fields")

        if not st.session_state.phase1_failed.empty:
            with st.expander("View detailed Phase 1 failure records"):
                st.dataframe(
                    st.session_state.phase1_failed,
                    use_container_width=True,
                    hide_index=True,
                )

# -------------------------- PHASE 2 --------------------------
else:
    if st.session_state.phase1_ranked.empty:
        st.warning("Complete Phase 1 first. No qualified field list is currently available.")
        st.stop()

    st.header("Phase / Module 2 — Investment Economics")

    stage = st.session_state.get("phase2_stage", 1)
    stage_names = {
        1: "Stage 1 — Expected vs Actual CAPEX/OPEX Gate",
        2: "Stage 2 — Economic Analysis",
        3: "Stage 3 — Sensitivity Analysis",
    }
    st.progress(stage / 3, text=stage_names[stage])

    if stage == 1:
        st.subheader("Step 1 — Select fields from the Phase 1 scoring board")
        phase1_fields = st.session_state.phase1_ranked["Field"].tolist()
        selected = st.multiselect(
            "Select one or more fields, or use Select All",
            options=phase1_fields,
            default=phase1_fields,
            key="phase2_selected_fields",
        )
        if st.button("Select All", key="select_all_phase2"):
            st.session_state["phase2_selected_fields"] = phase1_fields
            st.rerun()

        st.subheader("Step 2 — Expected CAPEX and OPEX")
        c1, c2 = st.columns(2)
        with c1:
            expected_capex = st.number_input(
                "Expected CAPEX",
                min_value=0.0,
                value=100_000_000.0,
                step=1_000_000.0,
                format="%.2f",
            )
        with c2:
            expected_opex = st.number_input(
                "Expected OPEX / year",
                min_value=0.0,
                value=10_000_000.0,
                step=500_000.0,
                format="%.2f",
            )

        st.caption(
            "Gate rule implemented from your flowchart: Expected CAPEX must be greater than Actual CAPEX AND Expected OPEX must be greater than Actual OPEX."
        )

        if st.button(
            "Run Expected vs Actual CAPEX/OPEX Gate",
            type="primary",
            disabled=not selected,
        ):
            passed, failed = economic_gate(
                selected,
                st.session_state.phase1_details,
                expected_capex,
                expected_opex,
            )
            st.session_state.phase2_gate_passed = passed
            st.session_state.phase2_gate_failed = failed
            st.session_state.phase2_gate_expected = {
                "CAPEX": expected_capex,
                "OPEX": expected_opex,
            }

        if not st.session_state.phase2_gate_passed.empty or not st.session_state.phase2_gate_failed.empty:
            st.divider()
            st.subheader("Gate Result")
            if not st.session_state.phase2_gate_passed.empty:
                st.success("Fields passing the expected-vs-actual economic gate")
                st.dataframe(
                    st.session_state.phase2_gate_passed,
                    use_container_width=True,
                    hide_index=True,
                )
            show_reasons(
                st.session_state.phase2_gate_failed,
                "Phase 2 Eliminated Fields — Gate Reasons",
            )

            eliminated = (
                st.session_state.phase2_gate_failed["Field"].tolist()
                if not st.session_state.phase2_gate_failed.empty
                else []
            )
            passed = (
                st.session_state.phase2_gate_passed["Field"].tolist()
                if not st.session_state.phase2_gate_passed.empty
                else []
            )

            if eliminated:
                st.warning(
                    "Your framework allows Phase 2 gate-eliminated fields to be manually overridden for further economic analysis."
                )
                overrides = st.multiselect(
                    "Select eliminated fields to continue anyway",
                    eliminated,
                    key="phase2_gate_overrides",
                )
            else:
                overrides = []

            economic_candidates = passed + overrides

            st.subheader("Proceed to Economic Analysis")
            st.write(f"Fields entering Stage 2: **{len(economic_candidates)}**")
            if economic_candidates and st.button("Continue to Economic Analysis →", type="primary"):
                st.session_state["phase2_economic_candidates"] = economic_candidates
                st.session_state["phase2_stage"] = 2
                st.rerun()

    elif stage == 2:
        st.subheader("Stage 2 — Economic Analysis")
        candidates = st.session_state.get("phase2_economic_candidates", [])
        if not candidates:
            st.warning("No fields are selected for economic analysis. Return to Stage 1.")
            if st.button("← Back to Stage 1"):
                st.session_state.phase2_stage = 1
                st.rerun()
            st.stop()

        st.write("Fields selected for economic analysis:")
        st.dataframe(
            pd.DataFrame({"Field": candidates}),
            use_container_width=True,
            hide_index=True,
        )

        st.subheader("Financial and Revenue Assumptions")
        c1, c2, c3 = st.columns(3)
        with c1:
            carbon_credit = st.number_input("Carbon credits ($/tCO₂)", min_value=0.0, value=20.0, step=1.0)
            government_subsidy = st.number_input("Government subsidy ($/tCO₂)", min_value=0.0, value=0.0, step=1.0)
            tax_incentive = st.number_input("Tax incentive ($/tCO₂)", min_value=0.0, value=0.0, step=1.0)
        with c2:
            storage_fee = st.number_input("Storage fee ($/tCO₂)", min_value=0.0, value=5.0, step=1.0)
            carbon_price = st.number_input("Carbon price ($/tCO₂)", min_value=0.0, value=50.0, step=1.0)
            discount_rate = st.number_input("Discount rate (%)", min_value=0.0, value=8.0, step=0.5)
        with c3:
            inflation_rate = st.number_input("Inflation rate (%)", min_value=0.0, value=2.5, step=0.25)
            project_lifetime = st.number_input("Project lifetime (years)", min_value=1, value=20, step=1)
            injection_rate_mtpa = st.number_input("CO₂ injection rate (MtCO₂/year)", min_value=0.001, value=1.0, step=0.1)

        inflate_revenues = st.checkbox(
            "Escalate revenue assumptions with inflation",
            value=True,
            help="If enabled, revenue per tonne and OPEX are escalated annually using the entered inflation rate.",
        )

        assumptions = {
            "carbon_credit": carbon_credit,
            "government_subsidy": government_subsidy,
            "tax_incentive": tax_incentive,
            "storage_fee": storage_fee,
            "carbon_price": carbon_price,
            "discount_rate": discount_rate,
            "inflation_rate": inflation_rate,
            "project_lifetime": int(project_lifetime),
            "injection_rate_mtpa": injection_rate_mtpa,
            "inflate_revenues": inflate_revenues,
        }

        if st.button("Calculate NPV, IRR & Payback Period", type="primary"):
            with st.spinner("Calculating field cashflows and economic ranking..."):
                try:
                    econ_df, raw = economic_analysis(
                        candidates,
                        st.session_state.phase1_details,
                        assumptions,
                    )
                    st.session_state.economic_df = econ_df
                    st.session_state.economic_raw = raw
                    st.session_state.economic_assumptions = assumptions
                    st.session_state.phase2_stage = 2
                except Exception as exc:
                    st.error(f"Economic calculation failed: {exc}")

        if not st.session_state.economic_df.empty:
            st.divider()
            st.subheader("Economic Ranking — Rank 1 = Best")
            st.dataframe(
                st.session_state.economic_df,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "NPV": st.column_config.NumberColumn(format="$%.2f"),
                    "IRR (%)": st.column_config.NumberColumn(format="%.2f%%"),
                    "Payback Period (years)": st.column_config.NumberColumn(format="%.2f"),
                },
            )
            st.caption(
                "Economic ranking is a transparent equal-weight aggregation of metric ranks: higher NPV, higher IRR and shorter payback are better. It is a screening ranking, not an investment recommendation."
            )

            st.download_button(
                "Download economic results (CSV)",
                st.session_state.economic_df.to_csv(index=False).encode("utf-8"),
                "phase2_economic_results.csv",
                "text/csv",
            )
            st.download_button(
                "Download economic results (Excel)",
                export_xlsx({"Economic Results": st.session_state.economic_df}),
                "phase2_economic_results.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

            st.divider()
            st.subheader("Proceed to Sensitivity Analysis?")
            b1, b2 = st.columns(2)
            with b1:
                if st.button("Go Ahead → Sensitivity Analysis", type="primary"):
                    st.session_state.phase2_stage = 3
                    st.rerun()
            with b2:
                if st.button("← Back to Stage 1"):
                    st.session_state.phase2_stage = 1
                    st.rerun()

    elif stage == 3:
        st.subheader("Stage 3 — Sensitivity Analysis")
        candidates = st.session_state.get("phase2_economic_candidates", [])
        if not candidates:
            st.warning("No economic-analysis candidates are available.")
            if st.button("← Back to Stage 2"):
                st.session_state.phase2_stage = 2
                st.rerun()
            st.stop()

        st.info(
            "Sensitivity method: one-at-a-time endpoint analysis. For each parameter, the low and high values are tested separately while all other base assumptions remain unchanged. This avoids hiding the effect of individual assumptions inside a 2⁷ combination table."
        )

        econ_fields = st.session_state.economic_df["Field"].tolist()
        selected_sens = st.multiselect(
            "Select fields for sensitivity analysis",
            econ_fields,
            default=econ_fields,
            key="sensitivity_selected_fields",
        )
        if st.button("Select All", key="select_all_sensitivity"):
            st.session_state["sensitivity_selected_fields"] = econ_fields
            st.rerun()

        base = st.session_state.get("economic_assumptions", {})
        if not base:
            st.warning("Run Stage 2 economic analysis first.")
            st.stop()

        st.subheader("Enter Low and High Values")
        parameters = [
            ("CAPEX", "CAPEX ($)", "Actual field CAPEX is varied"),
            ("OPEX", "OPEX ($/year)", "Actual field OPEX is varied"),
            ("Discount Rate", "Discount Rate (%)", "Base discount rate is varied"),
            ("Inflation Rate", "Inflation Rate (%)", "Base inflation rate is varied"),
            ("CO2 Injection Rate", "CO₂ Injection Rate (MtCO₂/year)", "Base injection rate is varied"),
            ("Project Lifetime", "Project Lifetime (years)", "Base lifetime is varied"),
            ("Carbon Credits", "Carbon Credits ($/tCO₂)", "Base carbon-credit value is varied"),
        ]

        ranges = {}
        for key, label, help_text in parameters:
            c1, c2 = st.columns(2)
            default_low = {
                "CAPEX": 0.8,
                "OPEX": 0.8,
                "Discount Rate": 0.8,
                "Inflation Rate": 0.8,
                "CO2 Injection Rate": 0.8,
                "Project Lifetime": 0.8,
                "Carbon Credits": 0.8,
            }[key]
            default_high = {
                "CAPEX": 1.2,
                "OPEX": 1.2,
                "Discount Rate": 1.2,
                "Inflation Rate": 1.2,
                "CO2 Injection Rate": 1.2,
                "Project Lifetime": 1.2,
                "Carbon Credits": 1.2,
            }[key]

            if key == "CAPEX":
                base_val = st.session_state.phase1_details[candidates[0]]["Actual CAPEX"] or 0
                low_default, high_default = base_val * default_low, base_val * default_high
                step = max(base_val * 0.01, 1.0)
            elif key == "OPEX":
                base_val = st.session_state.phase1_details[candidates[0]]["Actual OPEX"] or 0
                low_default, high_default = base_val * default_low, base_val * default_high
                step = max(base_val * 0.01, 1.0)
            elif key == "Discount Rate":
                base_val = base["discount_rate"]
                low_default, high_default = max(0, base_val * default_low), base_val * default_high
                step = 0.1
            elif key == "Inflation Rate":
                base_val = base["inflation_rate"]
                low_default, high_default = max(0, base_val * default_low), base_val * default_high
                step = 0.1
            elif key == "CO2 Injection Rate":
                base_val = base["injection_rate_mtpa"]
                low_default, high_default = max(0.001, base_val * default_low), base_val * default_high
                step = 0.01
            elif key == "Project Lifetime":
                base_val = base["project_lifetime"]
                low_default, high_default = max(1, round(base_val * default_low)), max(1, round(base_val * default_high))
                step = 1
            else:
                base_val = base["carbon_credit"]
                low_default, high_default = max(0, base_val * default_low), base_val * default_high
                step = 0.1

            with c1:
                low = st.number_input(f"{label} — Low", min_value=0.0, value=float(low_default), step=float(step), key=f"{key}_low")
            with c2:
                high = st.number_input(f"{label} — High", min_value=0.0, value=float(high_default), step=float(step), key=f"{key}_high")

            if high < low:
                st.error(f"{label}: high value must be greater than or equal to low value.")
            ranges[key] = (low, high)

        if st.button(
            "Run Sensitivity Analysis",
            type="primary",
            disabled=not selected_sens,
        ):
            invalid = [p for p, (lo, hi) in ranges.items() if hi < lo]
            if invalid:
                st.error("Correct the ranges before running sensitivity analysis.")
            else:
                with st.spinner("Running low/high one-at-a-time sensitivity calculations..."):
                    try:
                        sens_df, scenarios = sensitivity_oat(
                            selected_sens,
                            st.session_state.phase1_details,
                            base,
                            ranges,
                        )
                        st.session_state.sensitivity_df = sens_df
                        st.session_state.sensitivity_scenarios = scenarios
                    except Exception as exc:
                        st.error(f"Sensitivity calculation failed: {exc}")

        if not st.session_state.sensitivity_df.empty:
            st.divider()
            st.subheader("Sensitivity Results — Rank 1 = Most Robust")
            st.dataframe(
                st.session_state.sensitivity_df,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Sensitivity NPV Min": st.column_config.NumberColumn(format="$%.2f"),
                    "Sensitivity NPV Max": st.column_config.NumberColumn(format="$%.2f"),
                    "Sensitivity IRR Min (%)": st.column_config.NumberColumn(format="%.2f%%"),
                    "Sensitivity IRR Max (%)": st.column_config.NumberColumn(format="%.2f%%"),
                    "Sensitivity Payback Min (years)": st.column_config.NumberColumn(format="%.2f"),
                    "Sensitivity Payback Max (years)": st.column_config.NumberColumn(format="%.2f"),
                },
            )
            st.caption(
                "Sensitivity ranking prioritizes robustness: higher worst-case NPV, higher worst-case IRR, and lower worst-case payback. This is a screening/ranking output and should not be presented as an automatic investment decision."
            )

            st.subheader("Detailed One-at-a-Time Sensitivity Scenarios")
            for field, sdf in st.session_state.sensitivity_scenarios.items():
                with st.expander(field):
                    st.dataframe(sdf, use_container_width=True, hide_index=True)

            st.download_button(
                "Download sensitivity summary (CSV)",
                st.session_state.sensitivity_df.to_csv(index=False).encode("utf-8"),
                "phase2_sensitivity_summary.csv",
                "text/csv",
            )
            all_scenarios = (
                pd.concat(st.session_state.sensitivity_scenarios.values(), ignore_index=True)
                if st.session_state.sensitivity_scenarios
                else pd.DataFrame()
            )
            st.download_button(
                "Download sensitivity results (Excel)",
                export_xlsx(
                    {
                        "Sensitivity Summary": st.session_state.sensitivity_df,
                        "Scenario Details": all_scenarios,
                    }
                ),
                "phase2_sensitivity_results.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        st.divider()
        c1, c2 = st.columns(2)
        with c1:
            if st.button("← Back to Stage 2"):
                st.session_state.phase2_stage = 2
                st.rerun()
        with c2:
            if st.button("↻ Restart Phase 2"):
                reset_phase2()
                st.rerun()
