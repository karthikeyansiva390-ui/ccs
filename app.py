from __future__ import annotations

import io
import re
from typing import Any

import pandas as pd
import streamlit as st

from ccs_engine import (
    CATEGORIES,
    canonical_unit,
    economic_analysis,
    economic_gate,
    is_missing,
    norm_text,
    parse_number_strict,
    parse_reference,
    screen_one_field,
    sensitivity_oat,
    split_choices,
)

st.set_page_config(
    page_title="CCS Screening & Investment Decision Framework",
    page_icon="🌍",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ---------------------------------------------------------------------------
# UI styling
# ---------------------------------------------------------------------------
st.markdown(
    """
    <style>
    .main-title {font-size:2rem;font-weight:750;margin-bottom:.15rem;}
    .sub-title {color:#555;margin-bottom:1rem;}
    .input-header {
        font-weight:700;
        padding:.45rem .6rem;
        background:#f3f6f8;
        border:1px solid #d8e0e5;
        border-radius:.45rem;
    }
    .invalid-note {
        color:#b00020;
        font-weight:700;
        margin-top:-.45rem;
        margin-bottom:.35rem;
    }
    .valid-note {
        color:#147a3d;
        font-size:.82rem;
        margin-top:-.45rem;
        margin-bottom:.35rem;
    }
    .reference-note {
        color:#555;
        font-size:.85rem;
        margin-top:.15rem;
    }
    .stage-card {
        padding:.8rem 1rem;
        border-radius:.6rem;
        background:#f5f8fa;
        border:1px solid #d7e1e6;
        margin-bottom:.8rem;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    '<div class="main-title">CCS Screening & Investment Decision Framework</div>',
    unsafe_allow_html=True,
)
st.markdown(
    '<div class="sub-title">Manual field-data entry + reference-datasheet-driven Phase 1 screening → Phase 2 economic gate, economics and sensitivity analysis.</div>',
    unsafe_allow_html=True,
)


# ---------------------------------------------------------------------------
# Session state
# ---------------------------------------------------------------------------
def init_state():
    defaults = {
        "reference_meta": {},
        "manual_fields": {},
        "active_manual_field": None,
        "phase1_ranked": pd.DataFrame(),
        "phase1_failed": pd.DataFrame(),
        "phase1_details": {},
        "phase2_gate_passed": pd.DataFrame(),
        "phase2_gate_failed": pd.DataFrame(),
        "economic_df": pd.DataFrame(),
        "economic_raw": {},
        "sensitivity_df": pd.DataFrame(),
        "sensitivity_scenarios": {},
        "phase2_stage": 1,
        "invalid_input_messages": {},
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


init_state()


def reset_phase1_results():
    st.session_state.phase1_ranked = pd.DataFrame()
    st.session_state.phase1_failed = pd.DataFrame()
    st.session_state.phase1_details = {}
    st.session_state.phase2_gate_passed = pd.DataFrame()
    st.session_state.phase2_gate_failed = pd.DataFrame()
    st.session_state.economic_df = pd.DataFrame()
    st.session_state.economic_raw = {}
    st.session_state.sensitivity_df = pd.DataFrame()
    st.session_state.sensitivity_scenarios = {}
    st.session_state.phase2_stage = 1


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
    st.session_state.phase2_stage = 1


def export_xlsx(sheets: dict[str, pd.DataFrame]) -> bytes:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        for name, df in sheets.items():
            if df is None:
                continue
            df.to_excel(writer, sheet_name=name[:31], index=False)
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


def reference_unit_or_type(spec: dict[str, Any]) -> str:
    dtype = str(spec.get("data_type", "") or "").strip()
    return dtype if dtype else "Not specified"


def is_qualitative_spec(spec: dict[str, Any]) -> bool:
    dtype = norm_text(spec.get("data_type"))
    qualitative_tokens = {
        "qualitative",
        "text",
        "string",
        "categorical",
        "category",
        "yes/no",
        "yes no",
        "boolean",
        "bool",
    }
    return dtype in qualitative_tokens or any(
        token in dtype for token in ["qualitative", "categorical", "textual"]
    )


def qualitative_reference_choices(spec: dict[str, Any]) -> list[str]:
    """
    The dropdown is deliberately built from BOTH:
      - hard-cutoff reference text
      - every SAW start/end reference text

    Therefore the user is never forced to type a qualitative phrase manually.
    """
    raw_candidates = []

    if not is_missing(spec.get("hard_cutoff")):
        raw_candidates.extend(
            split_choices(spec.get("hard_cutoff"))
            or [str(spec.get("hard_cutoff"))]
        )

    for start, end, _score in spec.get("saw_ranges", []):
        if not is_missing(start):
            raw_candidates.extend(split_choices(start) or [str(start)])
        if not is_missing(end):
            raw_candidates.extend(split_choices(end) or [str(end)])

    result = []
    seen = set()
    for item in raw_candidates:
        text = str(item).strip()
        if not text or is_missing(text):
            continue
        key = re.sub(r"\s+", " ", text.casefold()).strip()
        if key not in seen:
            seen.add(key)
            result.append(text)

    return result


def validate_manual_value(raw: Any, spec: dict[str, Any]) -> tuple[bool, str]:
    """
    UI validation is intentionally stricter than the engine's final matching.

    Allowed:
      - blank
      - N/A
      - Nil
      - -
      - a valid numeric value for measurable fields
      - one of the reference-provided qualitative choices

    Invalid input is never silently sent to the screening engine.
    """
    if raw is None or str(raw).strip() == "":
        return True, ""

    text = str(raw).strip()

    if is_missing(text):
        return True, ""

    if is_qualitative_spec(spec):
        choices = qualitative_reference_choices(spec)
        normalized = re.sub(r"\s+", " ", text.casefold()).strip()
        allowed = {
            re.sub(r"\s+", " ", x.casefold()).strip()
            for x in choices
        }
        if normalized in allowed:
            return True, ""
        return (
            False,
            "Invalid input — select one of the qualitative values supplied by the reference datasheet, or leave blank / use N/A / Nil / -.",
        )

    # Numerical/measurable field.
    number = parse_number_strict(text)
    if number is None:
        return (
            False,
            "Invalid input — enter a numerical value in the format described by the reference datasheet, or leave blank / use N/A / Nil / -.",
        )

    if not pd.api.types.is_number(number):
        return False, "Invalid input."

    return True, ""


def manual_field_data(field_name: str, reference: dict[str, Any]) -> dict[str, Any]:
    """
    Convert the GUI state into exactly the structure expected by ccs_engine.
    """
    state = st.session_state.manual_fields[field_name]

    values = {cat: {} for cat in CATEGORIES}
    units = {cat: {} for cat in CATEGORIES}

    for cat in CATEGORIES:
        for spec in reference["categories"][cat]:
            parameter = spec["parameter"]
            key = norm_text(parameter)
            values[cat][key] = state["values"][cat].get(parameter, "")
            units[cat][key] = spec.get("data_type", "")

    return {
        "values": values,
        "units": units,
        "economics": {
            norm_text("Actual CAPEX"): state.get("actual_capex", ""),
            norm_text("Actual OPEX"): state.get("actual_opex", ""),
        },
    }


def create_manual_field(name: str, reference: dict[str, Any]):
    st.session_state.manual_fields[name] = {
        "values": {
            cat: {
                spec["parameter"]: ""
                for spec in reference["categories"][cat]
            }
            for cat in CATEGORIES
        },
        "actual_capex": "",
        "actual_opex": "",
    }



def clear_invalid_widget(widget_key: str, raw_value: str, spec: dict[str, Any]):
    """
    If a user types an unsupported value, clear the widget and retain a
    persistent warning for the next Streamlit rerun.
    """
    valid, message = validate_manual_value(raw_value, spec)
    messages = st.session_state.setdefault("invalid_input_messages", {})
    if not valid:
        st.session_state[widget_key] = ""
        messages[widget_key] = str(message)
    else:
        messages.pop(widget_key, None)


def render_manual_field_editor(field_name: str, reference: dict[str, Any]):
    """
    Main Phase-1 manual input dialog.

    The five common parameters are tabs. Every sub-parameter is presented as:
        [Sub-parameter name] [Input box] [Unit / Type]

    Qualitative parameters are selectboxes populated directly from the
    reference datasheet. Numerical parameters are text inputs because blank,
    N/A, Nil and '-' are valid "no-value" states.
    """
    state = st.session_state.manual_fields[field_name]

    st.markdown(
        f'<div class="stage-card"><b>Current field:</b> {field_name}<br>'
        "Enter the value for every sub-parameter. The reference datasheet controls the expected unit/type and qualitative vocabulary.</div>",
        unsafe_allow_html=True,
    )

    tabs = st.tabs(list(CATEGORIES) + ["Economics"])

    for tab, cat in zip(tabs[:5], CATEGORIES):
        with tab:
            st.caption(
                f"{cat}: enter values below. Qualitative values are selectable only from the reference datasheet."
            )

            h1, h2, h3 = st.columns([4, 5, 2])
            h1.markdown('<div class="input-header">Sub-parameter</div>', unsafe_allow_html=True)
            h2.markdown('<div class="input-header">Input value</div>', unsafe_allow_html=True)
            h3.markdown('<div class="input-header">Unit / Type</div>', unsafe_allow_html=True)

            for idx, spec in enumerate(reference["categories"][cat]):
                parameter = spec["parameter"]
                key_base = f"manual_{field_name}_{cat}_{idx}"
                current = state["values"][cat].get(parameter, "")

                c1, c2, c3 = st.columns([4, 5, 2])

                with c1:
                    st.write(parameter)

                with c2:
                    if is_qualitative_spec(spec):
                        choices = ["-- Select / No value --"] + qualitative_reference_choices(spec)
                        current_norm = str(current).strip().casefold()
                        selected_index = 0
                        for j, choice in enumerate(choices):
                            if choice.strip().casefold() == current_norm:
                                selected_index = j
                                break

                        selected = st.selectbox(
                            f"{parameter} input",
                            choices,
                            index=selected_index,
                            key=f"{key_base}_select",
                            label_visibility="collapsed",
                        )

                        new_value = "" if selected == choices[0] else selected
                        state["values"][cat][parameter] = new_value

                        if new_value:
                            st.markdown(
                                '<div class="valid-note">✓ Valid reference value selected</div>',
                                unsafe_allow_html=True,
                            )
                        else:
                            st.markdown(
                                '<div class="reference-note">Leave blank, or select N/A / Nil / - if those are available in the reference vocabulary.</div>',
                                unsafe_allow_html=True,
                            )

                    else:
                        widget_key = f"{key_base}_text"
                        aria_label = f"{field_name} | {cat} | {parameter} | input"

                        new_value = st.text_input(
                            aria_label,
                            value=str(current),
                            key=widget_key,
                            label_visibility="collapsed",
                            placeholder="Enter numerical value, or N/A / Nil / -",
                            on_change=clear_invalid_widget,
                            args=(widget_key, st.session_state.get(widget_key, ""), spec),
                        )

                        state["values"][cat][parameter] = new_value

                        valid, message = validate_manual_value(new_value, spec)
                        stored_message = st.session_state.get("invalid_input_messages", {}).get(widget_key)
                        if stored_message:
                            valid = False
                            message = stored_message

                        if not valid:
                            # The input is intentionally cleared immediately at the
                            # next Streamlit rerun by resetting the widget key.
                            # A CSS selector also paints the corresponding input
                            # outline red so the invalid box is visually obvious.
                            safe_aria = (
                                aria_label.replace("\\", "\\\\")
                                .replace('"', '\"')
                            )
                            st.markdown(
                                f"""
                                <style>
                                div[data-baseweb="input"]:has(input[aria-label="{safe_aria}"]) {{
                                    border: 2px solid #d62728 !important;
                                    border-radius: 0.45rem !important;
                                }}
                                input[aria-label="{safe_aria}"] {{
                                    border: 2px solid #d62728 !important;
                                    border-radius: 0.35rem !important;
                                }}
                                </style>
                                <div class="invalid-note">🔴 {message}</div>
                                """,
                                unsafe_allow_html=True,
                            )
                        elif str(new_value).strip():
                            st.markdown(
                                '<div class="valid-note">✓ Input format accepted</div>',
                                unsafe_allow_html=True,
                            )

                with c3:
                    st.caption(reference_unit_or_type(spec))


    # -------------------------------------------------------------------
    # Sixth tab: Economics
    # -------------------------------------------------------------------
    with tabs[5]:
        st.caption(
            "Enter the field's actual CAPEX and OPEX here. These values are "
            "stored with the field and are used by Phase 2's Expected vs Actual "
            "CAPEX/OPEX hard-cutoff gate."
        )

        ec1, ec2 = st.columns([4, 2])
        with ec1:
            st.markdown('<div class="input-header">Actual CAPEX</div>', unsafe_allow_html=True)
            state["actual_capex"] = st.text_input(
                "Actual CAPEX input",
                value=str(state.get("actual_capex", "")),
                key=f"{field_name}_actual_capex",
                placeholder="Enter CAPEX, or leave blank / N/A / Nil / -",
                label_visibility="collapsed",
            )
        with ec2:
            st.markdown('<div class="input-header">Unit / Type</div>', unsafe_allow_html=True)
            st.caption("Numerical / monetary")

        ec3, ec4 = st.columns([4, 2])
        with ec3:
            st.markdown('<div class="input-header">Actual OPEX</div>', unsafe_allow_html=True)
            state["actual_opex"] = st.text_input(
                "Actual OPEX input",
                value=str(state.get("actual_opex", "")),
                key=f"{field_name}_actual_opex",
                placeholder="Enter OPEX, or leave blank / N/A / Nil / -",
                label_visibility="collapsed",
            )
        with ec4:
            st.markdown('<div class="input-header">Unit / Type</div>', unsafe_allow_html=True)
            st.caption("Numerical / monetary")

        for label, value in [
            ("Actual CAPEX", state["actual_capex"]),
            ("Actual OPEX", state["actual_opex"]),
        ]:
            if str(value).strip() and not is_missing(value) and parse_number_strict(value) is None:
                st.error(f"{label}: invalid numerical input. Correct it before screening.")

        st.info(
            "CAPEX and OPEX entered here are the field's actual values. "
            "The expected CAPEX and expected OPEX used for the Phase-2 gate "
            "are entered later in Phase 2."
        )


# ---------------------------------------------------------------------------
# Sidebar
# ---------------------------------------------------------------------------
with st.sidebar:
    st.header("Framework Navigation")

    if not st.session_state.phase1_ranked.empty:
        phase_default = "Phase 2 — Investment Economics"
    else:
        phase_default = "Phase 1 — Screening"

    phase = st.radio(
        "Select phase",
        ["Phase 1 — Screening", "Phase 2 — Investment Economics"],
        index=0 if phase_default.startswith("Phase 1") else 1,
    )

    st.divider()
    st.caption("Input rules")
    st.caption("• The reference datasheet remains the source of truth.")
    st.caption("• Qualitative fields use a dropdown populated from the reference hard-cutoff + SAW text.")
    st.caption("• Numerical fields accept numbers only, plus blank / N/A / Nil / -.")
    st.caption("• Invalid numerical/text input is rejected before screening.")
    st.caption("• SAW scoring is checked first; the hard cut-off is used only when no SAW range/text matches.")
    st.caption("• No fuzzy matching or automatic unit conversion is performed.")
    st.caption("• Qualitative values are selected directly from reference SAW/hard-cutoff text.")
    st.caption("• Higher Phase-1 SAW score = better field.")


# ===========================================================================
# PHASE 1
# ===========================================================================
if phase == "Phase 1 — Screening":
    st.header("Phase / Module 1 — Screening")

    with st.expander("Phase-1 workflow", expanded=False):
        st.markdown(
            """
            **1. Reference data → 2. Manual field data → 3. Hard cut-off → 4. AHP-weighted SAW → 5. Ranking**

            The reference workbook/CSV supplies the hard cut-off values, AHP weights and all five SAW score ranges/texts. The user then enters the field data manually.
            """
        )

    # -----------------------------------------------------------------------
    # Reference data
    # -----------------------------------------------------------------------
    st.subheader("Step 1 — Upload Reference Datasheet")

    st.info(
        "You may upload multiple reference datasheets. Only one selected reference dataset is active for this Phase-1 run. "
        "Use 'Create another reference datasheet' by uploading another file if you want to compare different reference sets."
    )

    ref_files = st.file_uploader(
        "Upload reference datasheet(s)",
        type=["xlsx", "csv"],
        accept_multiple_files=True,
        key="reference_uploader",
        help="XLSX is recommended because the framework template contains Overall + five category sheets.",
    )

    reference = None

    if ref_files:
        names = [f.name for f in ref_files]
        active_name = st.selectbox(
            "Select active reference datasheet",
            names,
            key="active_reference_name",
        )

        active = next(f for f in ref_files if f.name == active_name)

        try:
            reference = parse_reference(active.getvalue(), active.name)
            st.session_state.reference_meta = {
                "name": active.name,
                "reference": reference,
            }
            st.success(
                f"Reference loaded successfully: {active.name}. "
                f"Detected {sum(len(reference['categories'][c]) for c in CATEGORIES)} sub-parameters."
            )
        except Exception as exc:
            st.session_state.reference_meta = {}
            st.error(f"Could not read the reference datasheet: {exc}")

    elif st.session_state.reference_meta:
        reference = st.session_state.reference_meta["reference"]
        st.info(
            f"Using previously loaded reference dataset: {st.session_state.reference_meta['name']}"
        )

    if reference is None:
        st.warning("Upload a reference datasheet before creating field input forms.")
        st.stop()

    # -----------------------------------------------------------------------
    # Create/manage fields
    # -----------------------------------------------------------------------
    st.subheader("Step 2 — Create Field Input Forms")

    st.write(
        "The field datasheet upload has intentionally been removed from the Phase-1 input workflow. "
        "Field values are now entered manually in the tabs below."
    )

    if not st.session_state.manual_fields:
        default_field = "Field 1"
        create_manual_field(default_field, reference)
        st.session_state.active_manual_field = default_field

    m1, m2 = st.columns([3, 1])
    with m1:
        field_names = list(st.session_state.manual_fields.keys())
        active_field = st.selectbox(
            "Select field to enter/edit",
            field_names,
            index=max(0, field_names.index(st.session_state.active_manual_field))
            if st.session_state.active_manual_field in field_names
            else 0,
            key="active_manual_field_select",
        )
        st.session_state.active_manual_field = active_field

    with m2:
        st.write("")
        st.write("")
        if st.button("＋ Add another field", use_container_width=True):
            new_number = 1
            while f"Field {new_number}" in st.session_state.manual_fields:
                new_number += 1
            new_name = f"Field {new_number}"
            create_manual_field(new_name, reference)
            st.session_state.active_manual_field = new_name
            reset_phase1_results()
            st.rerun()

    # Optional rename
    with st.expander("Rename current field"):
        rename_to = st.text_input(
            "New field name",
            value=active_field,
            key=f"rename_{active_field}",
        )
        if st.button("Save field name"):
            clean = rename_to.strip()
            if not clean:
                st.error("Field name cannot be blank.")
            elif clean != active_field and clean in st.session_state.manual_fields:
                st.error("That field name already exists.")
            else:
                st.session_state.manual_fields[clean] = st.session_state.manual_fields.pop(active_field)
                st.session_state.active_manual_field = clean
                reset_phase1_results()
                st.rerun()

    render_manual_field_editor(active_field, reference)

    # -----------------------------------------------------------------------
    # Screening validation and execution
    # -----------------------------------------------------------------------
    st.divider()
    st.subheader("Step 3 — Start Screening")

    invalid_records = []

    for field_name, field_state in st.session_state.manual_fields.items():
        for cat in CATEGORIES:
            for spec in reference["categories"][cat]:
                raw = field_state["values"][cat].get(spec["parameter"], "")
                valid, msg = validate_manual_value(raw, spec)
                if not valid:
                    invalid_records.append(
                        f"{field_name} → {cat} → {spec['parameter']}: {msg}"
                    )

        for econ_label in ["actual_capex", "actual_opex"]:
            raw = field_state.get(econ_label, "")
            if str(raw).strip() and not is_missing(raw) and parse_number_strict(raw) is None:
                invalid_records.append(
                    f"{field_name} → {econ_label.replace('_', ' ').title()}: invalid numerical value."
                )

    if invalid_records:
        st.error(
            f"{len(invalid_records)} invalid input(s) detected. Correct the red/invalid fields before starting the screening process."
        )
        with st.expander("View invalid inputs"):
            for msg in invalid_records:
                st.markdown(f"- {msg}")

    if st.button(
        "▶ Start the Screening Process",
        type="primary",
        use_container_width=True,
        disabled=bool(invalid_records),
    ):
        with st.spinner("Checking SAW ranges/text first, applying hard-cutoff fallback where needed, calculating AHP-weighted SAW scores and ranking fields..."):
            passed = []
            failed = []
            details = {}

            for field_name in st.session_state.manual_fields:
                data = manual_field_data(field_name, reference)
                result = screen_one_field(field_name, data, reference)
                details[field_name] = result

                if result["Passed"]:
                    passed.append(result)
                else:
                    failed.append(result)

            passed.sort(
                key=lambda x: x["Overall SAW Score"],
                reverse=True,
            )

            ranked_rows = []
            for rank, result in enumerate(passed, start=1):
                ranked_rows.append(
                    {
                        "Rank": rank,
                        "Field": result["Field"],
                        "Overall SAW Score": round(result["Overall SAW Score"], 4),
                        "Technical": round(result["Technical"], 4),
                        "Environmental": round(result["Environmental"], 4),
                        "Regulatory": round(result["Regulatory"], 4),
                        "Long Term Operation": round(result["Long Term Operation"], 4),
                        "Risk": round(result["Risk"], 4),
                    }
                )

            failed_rows = [
                {
                    "Field": result["Field"],
                    "Reason": "\n".join(f"• {r}" for r in result["Reasons"]),
                }
                for result in failed
            ]

            st.session_state.phase1_ranked = pd.DataFrame(ranked_rows)
            st.session_state.phase1_failed = pd.DataFrame(failed_rows)
            st.session_state.phase1_details = details
            reset_phase2()

        st.success("Phase-1 screening completed.")

    # -----------------------------------------------------------------------
    # Results
    # -----------------------------------------------------------------------
    if not st.session_state.phase1_ranked.empty or not st.session_state.phase1_failed.empty:
        st.divider()
        st.header("Phase-1 Results")

        if not st.session_state.phase1_ranked.empty:
            st.subheader("Qualified Field Ranking — Best to Worst")

            st.dataframe(
                st.session_state.phase1_ranked,
                use_container_width=True,
                hide_index=True,
            )

            st.download_button(
                "Download Phase-1 Ranking CSV",
                st.session_state.phase1_ranked.to_csv(index=False).encode("utf-8"),
                "phase1_ranking.csv",
                "text/csv",
            )

            st.markdown(
                "Higher **Overall SAW Score** is better. The ranking is descending because the best field has the highest score."
            )
        else:
            st.warning("No field passed all Phase-1 screening requirements.")

        show_reasons(
            st.session_state.phase1_failed,
            "Phase-1 Eliminated Fields and Reasons",
        )

        with st.expander("🔎 Diagnostic — what the framework understood from every input"):
            st.caption(
                "This diagnostic is intentionally exposed so the user can audit the normalization and reference matching."
            )

            diagnostic_rows = []
            for field_name, result in st.session_state.phase1_details.items():
                for cat in CATEGORIES:
                    diagnostic_rows.extend(
                        [
                            {
                                "Field": field_name,
                                "Common Parameter": cat,
                                **detail,
                            }
                            for detail in result["Parameter Details"][cat]
                        ]
                    )

            diagnostic_df = pd.DataFrame(diagnostic_rows)
            if not diagnostic_df.empty:
                st.dataframe(
                    diagnostic_df,
                    use_container_width=True,
                    hide_index=True,
                )

        if st.button("↻ Clear Phase-1 results and edit field inputs"):
            reset_phase1_results()
            st.rerun()


# ===========================================================================
# PHASE 2
# ===========================================================================
else:
    st.header("Phase / Module 2 — Investment Economics")

    if not st.session_state.phase1_details:
        st.warning("Complete Phase 1 first.")
        st.stop()

    details = st.session_state.phase1_details
    all_fields = list(details.keys())

    st.subheader("Stage 1 — Select Fields")

    # The framework intentionally allows Phase-2 gate-failed fields to be
    # selected again later. Therefore use ALL Phase-1 fields, not only ranked.
    selection_mode = st.radio(
        "Field selection",
        ["Select individual fields", "Select all fields"],
        horizontal=True,
        key="phase2_selection_mode",
    )

    if selection_mode == "Select all fields":
        selected_fields = all_fields
        st.success(f"{len(selected_fields)} field(s) selected.")
    else:
        selected_fields = st.multiselect(
            "Select one or more fields from the Phase-1 board",
            all_fields,
            default=[],
            key="phase2_selected_fields",
        )

    if not selected_fields:
        st.info("Select at least one field to continue.")
        st.stop()

    st.subheader("Stage 1 — Expected vs Actual CAPEX & OPEX Hard Gate")

    expected_capex = st.number_input(
        "Expected CAPEX",
        min_value=0.0,
        value=100.0,
        step=10.0,
        key="expected_capex_manual",
    )
    expected_opex = st.number_input(
        "Expected OPEX",
        min_value=0.0,
        value=10.0,
        step=1.0,
        key="expected_opex_manual",
    )

    if st.button("Apply CAPEX & OPEX Hard Cut-Off Gate", type="primary"):
        with st.spinner("Comparing expected and actual field economics..."):
            gate_passed, gate_failed = economic_gate(
                selected_fields,
                details,
                expected_capex,
                expected_opex,
            )
            st.session_state.phase2_gate_passed = gate_passed
            st.session_state.phase2_gate_failed = gate_failed
            st.session_state.phase2_stage = 2

            # Reset later-stage outputs.
            st.session_state.economic_df = pd.DataFrame()
            st.session_state.economic_raw = {}
            st.session_state.sensitivity_df = pd.DataFrame()
            st.session_state.sensitivity_scenarios = {}

        st.success("CAPEX/OPEX gate completed.")

    gate_passed = st.session_state.phase2_gate_passed
    gate_failed = st.session_state.phase2_gate_failed

    if not gate_passed.empty or not gate_failed.empty:
        st.subheader("CAPEX/OPEX Gate Result")

        if not gate_passed.empty:
            st.success("Fields passing the CAPEX/OPEX hard gate")
            st.dataframe(gate_passed, use_container_width=True, hide_index=True)

        if not gate_failed.empty:
            st.warning(
                "These fields failed the CAPEX/OPEX gate. They remain selectable for further economic analysis, exactly as specified in the framework."
            )
            st.dataframe(gate_failed, use_container_width=True, hide_index=True)

        # Allow all selected fields to continue, including gate-eliminated ones.
        st.subheader("Stage 2 — Economic Analysis Field Selection")

        eligible_for_economics = list(dict.fromkeys(selected_fields))

        economic_selection_mode = st.radio(
            "Economic-analysis field selection",
            ["Select individual fields", "Select all selected fields"],
            horizontal=True,
            key="economic_selection_mode",
        )

        if economic_selection_mode == "Select all selected fields":
            economic_fields = eligible_for_economics
        else:
            economic_fields = st.multiselect(
                "Choose fields for NPV / IRR / Payback",
                eligible_for_economics,
                key="economic_fields",
            )

        if not economic_fields:
            st.info("Select fields for economic analysis.")
            st.stop()

        st.subheader("Economic Assumptions")

        a1, a2, a3 = st.columns(3)
        with a1:
            carbon_credit = st.number_input(
                "Carbon credits ($/tCO₂)", min_value=0.0, value=20.0, step=1.0
            )
            government_subsidy = st.number_input(
                "Government subsidy ($/tCO₂)", min_value=0.0, value=0.0, step=1.0
            )
            tax_incentive = st.number_input(
                "Tax incentive ($/tCO₂)", min_value=0.0, value=0.0, step=1.0
            )
        with a2:
            storage_fee = st.number_input(
                "Storage fee ($/tCO₂)", min_value=0.0, value=5.0, step=1.0
            )
            carbon_price = st.number_input(
                "Carbon price ($/tCO₂)", min_value=0.0, value=50.0, step=1.0
            )
            discount_rate = st.number_input(
                "Discount rate (%)", min_value=0.0, value=8.0, step=0.5
            )
        with a3:
            inflation_rate = st.number_input(
                "Inflation rate (%)", min_value=0.0, value=2.5, step=0.25
            )
            project_lifetime = st.number_input(
                "Project lifetime (years)", min_value=1, value=20, step=1
            )
            injection_rate_mtpa = st.number_input(
                "CO₂ injection rate (MtCO₂/year)",
                min_value=0.001,
                value=1.0,
                step=0.1,
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
        }

        if st.button("Calculate NPV, IRR & Payback Period", type="primary"):
            try:
                with st.spinner("Calculating project economics..."):
                    econ_df, econ_raw = economic_analysis(
                        economic_fields,
                        details,
                        assumptions,
                    )
                    st.session_state.economic_df = econ_df
                    st.session_state.economic_raw = econ_raw
                    st.session_state.phase2_stage = 3
                    st.session_state.sensitivity_df = pd.DataFrame()
                    st.session_state.sensitivity_scenarios = {}
                st.success("Economic analysis completed.")
            except Exception as exc:
                st.error(f"Economic calculation failed: {exc}")

        if not st.session_state.economic_df.empty:
            st.divider()
            st.subheader("Economic Analysis Results — Best to Worst")

            st.dataframe(
                st.session_state.economic_df,
                use_container_width=True,
                hide_index=True,
            )

            st.download_button(
                "Download Economic Results CSV",
                st.session_state.economic_df.to_csv(index=False).encode("utf-8"),
                "phase2_economic_results.csv",
                "text/csv",
            )

            st.divider()
            st.subheader("Stage 3 — Sensitivity Analysis")

            do_sensitivity = st.radio(
                "Do you want to perform sensitivity analysis?",
                ["Go ahead", "Finish"],
                horizontal=True,
                key="do_sensitivity",
            )

            if do_sensitivity == "Go ahead":
                sensitivity_mode = st.radio(
                    "Sensitivity field selection",
                    ["Select individual fields", "Select all economic-analysis fields"],
                    horizontal=True,
                    key="sensitivity_selection_mode",
                )

                if sensitivity_mode == "Select all economic-analysis fields":
                    sensitivity_fields = economic_fields
                else:
                    sensitivity_fields = st.multiselect(
                        "Select fields for sensitivity analysis",
                        economic_fields,
                        key="sensitivity_fields",
                    )

                st.markdown(
                    "Enter **low and high** values. The framework will rerun NPV, IRR and Payback Period for the requested sensitivity range."
                )

                base = assumptions
                ranges = {}

                sensitivity_specs = [
                    ("CAPEX", "capex"),
                    ("OPEX", "opex"),
                    ("Discount rate (%)", "discount_rate"),
                    ("Inflation rate (%)", "inflation_rate"),
                    ("CO₂ Injection rate (MtCO₂/year)", "injection_rate_mtpa"),
                    ("Project lifetime (years)", "project_lifetime"),
                    ("Carbon credits ($/tCO₂)", "carbon_credit"),
                ]

                for label, key in sensitivity_specs:
                    if key in {"project_lifetime"}:
                        low_default = max(1, int(base[key]) - 5)
                        high_default = int(base[key]) + 5
                        step = 1
                    else:
                        low_default = float(base.get(key, 0.0))
                        high_default = float(base.get(key, 0.0))
                        step = 0.1 if key == "injection_rate_mtpa" else 1.0

                    s1, s2 = st.columns(2)
                    with s1:
                        low = st.number_input(
                            f"{label} — Low",
                            min_value=0.0 if key != "project_lifetime" else 1.0,
                            value=float(low_default),
                            step=float(step),
                            key=f"sens_{key}_low",
                        )
                    with s2:
                        high = st.number_input(
                            f"{label} — High",
                            min_value=0.0 if key != "project_lifetime" else 1.0,
                            value=float(high_default),
                            step=float(step),
                            key=f"sens_{key}_high",
                        )

                    ranges[key] = (
                        int(low) if key == "project_lifetime" else low,
                        int(high) if key == "project_lifetime" else high,
                    )

                if st.button("Run Sensitivity Analysis", type="primary"):
                    if not sensitivity_fields:
                        st.error("Select at least one field.")
                    else:
                        try:
                            with st.spinner("Running sensitivity scenarios..."):
                                sens_df, scenarios = sensitivity_oat(
                                    sensitivity_fields,
                                    details,
                                    base,
                                    ranges,
                                )
                                st.session_state.sensitivity_df = sens_df
                                st.session_state.sensitivity_scenarios = scenarios
                            st.success("Sensitivity analysis completed.")
                        except Exception as exc:
                            st.error(f"Sensitivity analysis failed: {exc}")

            if not st.session_state.sensitivity_df.empty:
                st.subheader("Sensitivity Analysis Results — Best to Worst")
                st.dataframe(
                    st.session_state.sensitivity_df,
                    use_container_width=True,
                    hide_index=True,
                )

                st.download_button(
                    "Download Sensitivity Results CSV",
                    st.session_state.sensitivity_df.to_csv(index=False).encode("utf-8"),
                    "phase2_sensitivity_results.csv",
                    "text/csv",
                )

        st.divider()
        st.subheader("Phase-2 Eliminated Fields — CAPEX/OPEX Gate Reasons")

        if gate_failed.empty:
            st.success("No field was eliminated by the Phase-2 CAPEX/OPEX gate.")
        else:
            for _, row in gate_failed.iterrows():
                st.markdown(f"**{row['Field']}**")
                for reason in str(row["Reason"]).splitlines():
                    reason = reason.lstrip("• ").strip()
                    if reason:
                        st.markdown(f"- {reason}")

        if st.button("↻ Restart Phase 2"):
            reset_phase2()
            st.rerun()
