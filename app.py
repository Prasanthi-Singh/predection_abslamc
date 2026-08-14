import io
import re
from typing import Dict, Tuple

import numpy as np
import pandas as pd
import streamlit as st


st.set_page_config(
    page_title="Employee Performance Forecast & Scenario Planner",
    page_icon="📊",
    layout="wide",
)

st.markdown(
    """
    <style>
    .block-container {padding-top: 1.4rem; padding-bottom: 2rem;}
    div[data-testid="stMetric"] {
        border: 1px solid rgba(128,128,128,.25);
        border-radius: 12px;
        padding: 14px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# -------------------------------------------------------------------
# Expected input columns
# -------------------------------------------------------------------
# Full structure expected from the employee performance Excel.
# Extra columns are still allowed.
EXPECTED_INPUT_COLUMNS = [
    "Emp Code",
    "ADID",
    "Status",
    "Type",
    "Employee Name",
    "DOJ",
    "ZONE",
    "REGION",
    "EM City",
    "MKT TYPE",
    "FY 26 TGT EQ",
    "YTD June EQ TGT",
    "Equity GS Ach YTD June",
    "Equity GS % Ach YTD June",
    "Quarterly Gap",
    "Scenario 1",
    "Scenario 2",
    "120% Target",
    "Scenario 3",
    "FY 26 TGT DT",
    "YTD June DT TGT",
    "Debt GS Ach",
    "Debt GS Ach%",
    "Quarterly Gap",
    "Scenario 1",
    "Scenario 2",
    "120% Target",
    "Scenario 3",
    "FY 26 TGT LIQ",
    "YTD June LIQ TGT",
    "Liquid GS Ach",
    "Liquid GS Ach %",
    "Quarterly Gap",
    "Scenario 1",
    "Scenario 2",
    "120% Target",
    "Scenario 3",
    "WGS % AchYTD June",
    "WGS Weightage",
    "WGS Score",
    "X",
    "FY 26 TGT EQ NS",
    "YTD June EQ NS TGT",
    "Equity NS Ach YTD June",
    "Equity NS Ach",
    "Quarterly Gap",
    "Scenario 1",
    "Scenario 2",
    "120% Target",
    "Scenario 3",
    "FY 26 TGT DT NS",
    "YTD June DT NS TGT",
    "Debt NS Ach",
    "Debt NS Ach",
    "Quarterly Gap",
    "Scenario 1",
    "Scenario 2",
    "120% Target",
    "Scenario 3",
    "FY 26 TGT LIQ NS",
    "YTD June LIQ NS TGT",
    "Liquid NS Ach",
    "Liquid NS Ach",
]

# Only these columns are mandatory for the scenario calculations.
REQUIRED_COLUMNS = [
    "Employee Name",
    "FY 26 TGT EQ",
    "Equity GS Ach YTD June",
    "FY 26 TGT DT",
    "Debt GS Ach",
    "FY 26 TGT LIQ",
    "Liquid GS Ach",
    "FY 26 TGT EQ NS",
    "Equity NS Ach YTD June",
    "FY 26 TGT DT NS",
    "Debt NS Ach",
    "FY 26 TGT LIQ NS",
    "Liquid NS Ach",
]

# Identification / organisation fields carried into result tables.
METADATA_COLUMNS = [
    "Emp Code",
    "ADID",
    "Status",
    "Type",
    "Employee Name",
    "DOJ",
    "ZONE",
    "REGION",
    "EM City",
    "MKT TYPE",
]

FORECAST_NAME = "Current Run Rate Forecast"

SCENARIO_NAMES = [
    FORECAST_NAME,
    "Scenario 1",
    "Scenario 2",
    "Scenario 3",
    "Scenario 4",
]


def normalize_column_name(name: object) -> str:
    """Normalize Excel headers so NBSP / repeated spaces do not break matching."""
    value = str(name).replace("\u00a0", " ").strip()
    return " ".join(value.split())


def normalize_input_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Return a copy with cleaned Excel column names."""
    cleaned = df.copy()
    cleaned.columns = [normalize_column_name(c) for c in cleaned.columns]
    return cleaned


def clean_numeric(series: pd.Series) -> pd.Series:
    """Convert Excel-style numeric columns safely to float."""
    return pd.to_numeric(series, errors="coerce").fillna(0.0)


def validate_input(df: pd.DataFrame) -> Tuple[bool, list]:
    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    return len(missing) == 0, missing


def proportional_non_equity_targets(
    eq_target: pd.Series,
    debt_target: pd.Series,
    liq_target: pd.Series,
    equity_multiplier: float,
    overall_multiplier: float,
) -> Tuple[pd.Series, pd.Series, pd.Series]:
    """
    Fix Equity at equity_multiplier and make total target equal
    overall_multiplier by distributing the remaining target between
    Debt and Liquid in proportion to their FY targets.
    """
    total_target = eq_target + debt_target + liq_target
    required_total = overall_multiplier * total_target
    required_equity = equity_multiplier * eq_target

    remaining = (required_total - required_equity).clip(lower=0)
    non_eq_base = debt_target + liq_target

    debt_share = np.where(
        non_eq_base > 0,
        debt_target / non_eq_base,
        0.0,
    )
    liq_share = np.where(
        non_eq_base > 0,
        liq_target / non_eq_base,
        0.0,
    )

    debt_final = remaining * debt_share
    liq_final = remaining * liq_share

    return required_equity, pd.Series(debt_final, index=eq_target.index), pd.Series(
        liq_final, index=eq_target.index
    )


def component_target_amounts(
    df: pd.DataFrame,
    prefix: str,
    scenario: str,
    assumptions: Dict[str, float],
) -> Tuple[pd.Series, pd.Series, pd.Series]:
    """
    Returns final target amounts for Equity, Debt and Liquid.

    prefix == "GS" uses Gross Sales columns.
    prefix == "NS" uses Net Sales columns.
    """
    if prefix == "GS":
        eq_tgt = clean_numeric(df["FY 26 TGT EQ"])
        dt_tgt = clean_numeric(df["FY 26 TGT DT"])
        liq_tgt = clean_numeric(df["FY 26 TGT LIQ"])
    else:
        eq_tgt = clean_numeric(df["FY 26 TGT EQ NS"])
        dt_tgt = clean_numeric(df["FY 26 TGT DT NS"])
        liq_tgt = clean_numeric(df["FY 26 TGT LIQ NS"])

    if scenario == "Scenario 1":
        multiplier = assumptions["scenario1_overall"]
        return (
            eq_tgt * multiplier,
            dt_tgt * multiplier,
            liq_tgt * multiplier,
        )

    if scenario == "Scenario 2":
        multiplier = assumptions["scenario2_overall"]
        return (
            eq_tgt * multiplier,
            dt_tgt * multiplier,
            liq_tgt * multiplier,
        )

    if scenario == "Scenario 3":
        return proportional_non_equity_targets(
            eq_tgt,
            dt_tgt,
            liq_tgt,
            assumptions["scenario3_equity"],
            assumptions["scenario3_overall"],
        )

    # Scenario 4:
    # Equity is fixed at 110% by default.
    # Debt + Liquid are balanced proportionally so total reaches
    # the minimum overall target (85% by default).
    return proportional_non_equity_targets(
        eq_tgt,
        dt_tgt,
        liq_tgt,
        assumptions["scenario4_equity"],
        assumptions["scenario4_overall_min"],
    )


def current_run_rate_projection(
    achievement: pd.Series,
    fy_target: pd.Series,
    months_completed: int,
) -> Tuple[pd.Series, pd.Series, pd.Series]:
    """Annualise current YTD pace without applying any scenario assumption."""
    completed = max(int(months_completed), 1)
    monthly_run_rate = achievement / completed
    projected_fy_achievement = monthly_run_rate * 12
    projected_fy_percent = pd.Series(
        np.where(fy_target > 0, projected_fy_achievement / fy_target, 0.0),
        index=achievement.index,
    )
    return monthly_run_rate, projected_fy_achievement, projected_fy_percent



def calculate_scenario(
    df: pd.DataFrame,
    scenario: str,
    months_remaining: int,
    assumptions: Dict[str, float],
) -> pd.DataFrame:
    result = pd.DataFrame(index=df.index)

    for column in METADATA_COLUMNS:
        if column in df.columns:
            result[column] = df[column]

    if "Employee Name" not in result.columns:
        result["Employee Name"] = df["Employee Name"].astype(str)

    months = max(int(months_remaining), 1)
    months_completed = max(12 - int(months_remaining), 1)

    # ---------------- Current Run Rate Forecast ----------------
    # No management scenario is imposed. Actual YTD performance is
    # annualised at the employee's current monthly pace.
    if scenario == FORECAST_NAME:
        # Gross Sales
        gs_eq_tgt = clean_numeric(df["FY 26 TGT EQ"])
        gs_dt_tgt = clean_numeric(df["FY 26 TGT DT"])
        gs_liq_tgt = clean_numeric(df["FY 26 TGT LIQ"])

        gs_eq_ach = clean_numeric(df["Equity GS Ach YTD June"])
        gs_dt_ach = clean_numeric(df["Debt GS Ach"])
        gs_liq_ach = clean_numeric(df["Liquid GS Ach"])

        gs_eq_rr, gs_eq_proj, gs_eq_pct = current_run_rate_projection(
            gs_eq_ach, gs_eq_tgt, months_completed
        )
        gs_dt_rr, gs_dt_proj, gs_dt_pct = current_run_rate_projection(
            gs_dt_ach, gs_dt_tgt, months_completed
        )
        gs_liq_rr, gs_liq_proj, gs_liq_pct = current_run_rate_projection(
            gs_liq_ach, gs_liq_tgt, months_completed
        )

        result["GS Equity FY Target"] = gs_eq_tgt
        result["GS Equity YTD Achievement"] = gs_eq_ach
        result["GS Equity Current Run Rate / Month"] = gs_eq_rr
        result["GS Equity Projected FY Achievement"] = gs_eq_proj
        result["GS Equity Projected FY %"] = gs_eq_pct

        result["GS Debt FY Target"] = gs_dt_tgt
        result["GS Debt YTD Achievement"] = gs_dt_ach
        result["GS Debt Current Run Rate / Month"] = gs_dt_rr
        result["GS Debt Projected FY Achievement"] = gs_dt_proj
        result["GS Debt Projected FY %"] = gs_dt_pct

        result["GS Liquid FY Target"] = gs_liq_tgt
        result["GS Liquid YTD Achievement"] = gs_liq_ach
        result["GS Liquid Current Run Rate / Month"] = gs_liq_rr
        result["GS Liquid Projected FY Achievement"] = gs_liq_proj
        result["GS Liquid Projected FY %"] = gs_liq_pct

        result["GS Total FY Target"] = gs_eq_tgt + gs_dt_tgt + gs_liq_tgt
        result["GS Total YTD Achievement"] = gs_eq_ach + gs_dt_ach + gs_liq_ach
        result["GS Total Current Run Rate / Month"] = gs_eq_rr + gs_dt_rr + gs_liq_rr
        result["GS Total Projected FY Achievement"] = gs_eq_proj + gs_dt_proj + gs_liq_proj
        result["GS Overall Projected FY %"] = np.where(
            result["GS Total FY Target"] > 0,
            result["GS Total Projected FY Achievement"] / result["GS Total FY Target"],
            0.0,
        )

        # Net Sales
        ns_eq_tgt = clean_numeric(df["FY 26 TGT EQ NS"])
        ns_dt_tgt = clean_numeric(df["FY 26 TGT DT NS"])
        ns_liq_tgt = clean_numeric(df["FY 26 TGT LIQ NS"])

        ns_eq_ach = clean_numeric(df["Equity NS Ach YTD June"])
        ns_dt_ach = clean_numeric(df["Debt NS Ach"])
        ns_liq_ach = clean_numeric(df["Liquid NS Ach"])

        ns_eq_rr, ns_eq_proj, ns_eq_pct = current_run_rate_projection(
            ns_eq_ach, ns_eq_tgt, months_completed
        )
        ns_dt_rr, ns_dt_proj, ns_dt_pct = current_run_rate_projection(
            ns_dt_ach, ns_dt_tgt, months_completed
        )
        ns_liq_rr, ns_liq_proj, ns_liq_pct = current_run_rate_projection(
            ns_liq_ach, ns_liq_tgt, months_completed
        )

        result["NS Equity FY Target"] = ns_eq_tgt
        result["NS Equity YTD Achievement"] = ns_eq_ach
        result["NS Equity Current Run Rate / Month"] = ns_eq_rr
        result["NS Equity Projected FY Achievement"] = ns_eq_proj
        result["NS Equity Projected FY %"] = ns_eq_pct

        result["NS Debt FY Target"] = ns_dt_tgt
        result["NS Debt YTD Achievement"] = ns_dt_ach
        result["NS Debt Current Run Rate / Month"] = ns_dt_rr
        result["NS Debt Projected FY Achievement"] = ns_dt_proj
        result["NS Debt Projected FY %"] = ns_dt_pct

        result["NS Liquid FY Target"] = ns_liq_tgt
        result["NS Liquid YTD Achievement"] = ns_liq_ach
        result["NS Liquid Current Run Rate / Month"] = ns_liq_rr
        result["NS Liquid Projected FY Achievement"] = ns_liq_proj
        result["NS Liquid Projected FY %"] = ns_liq_pct

        result["NS Total FY Target"] = ns_eq_tgt + ns_dt_tgt + ns_liq_tgt
        result["NS Total YTD Achievement"] = ns_eq_ach + ns_dt_ach + ns_liq_ach
        result["NS Total Current Run Rate / Month"] = ns_eq_rr + ns_dt_rr + ns_liq_rr
        result["NS Total Projected FY Achievement"] = ns_eq_proj + ns_dt_proj + ns_liq_proj
        result["NS Overall Projected FY %"] = np.where(
            result["NS Total FY Target"] > 0,
            result["NS Total Projected FY Achievement"] / result["NS Total FY Target"],
            0.0,
        )

        result["Scenario"] = scenario
        result["Months Completed"] = months_completed
        result["Months Remaining"] = int(months_remaining)
        return result

    # ---------------- Scenario planning: Gross Sales ----------------
    gs_eq_final, gs_dt_final, gs_liq_final = component_target_amounts(
        df, "GS", scenario, assumptions
    )

    gs_eq_ach = clean_numeric(df["Equity GS Ach YTD June"])
    gs_dt_ach = clean_numeric(df["Debt GS Ach"])
    gs_liq_ach = clean_numeric(df["Liquid GS Ach"])

    result["GS Equity Final Target"] = gs_eq_final
    result["GS Debt Final Target"] = gs_dt_final
    result["GS Liquid Final Target"] = gs_liq_final

    result["GS Equity Required / Month"] = (
        (gs_eq_final - gs_eq_ach).clip(lower=0) / months
    )
    result["GS Debt Required / Month"] = (
        (gs_dt_final - gs_dt_ach).clip(lower=0) / months
    )
    result["GS Liquid Required / Month"] = (
        (gs_liq_final - gs_liq_ach).clip(lower=0) / months
    )

    result["GS Total Required / Month"] = (
        result["GS Equity Required / Month"]
        + result["GS Debt Required / Month"]
        + result["GS Liquid Required / Month"]
    )

    gs_fy_total = (
        clean_numeric(df["FY 26 TGT EQ"])
        + clean_numeric(df["FY 26 TGT DT"])
        + clean_numeric(df["FY 26 TGT LIQ"])
    )
    gs_final_total = gs_eq_final + gs_dt_final + gs_liq_final
    result["GS Scenario Overall %"] = np.where(
        gs_fy_total > 0, gs_final_total / gs_fy_total, 0
    )

    # ---------------- Scenario planning: Net Sales ----------------
    ns_eq_final, ns_dt_final, ns_liq_final = component_target_amounts(
        df, "NS", scenario, assumptions
    )

    ns_eq_ach = clean_numeric(df["Equity NS Ach YTD June"])
    ns_dt_ach = clean_numeric(df["Debt NS Ach"])
    ns_liq_ach = clean_numeric(df["Liquid NS Ach"])

    result["NS Equity Final Target"] = ns_eq_final
    result["NS Debt Final Target"] = ns_dt_final
    result["NS Liquid Final Target"] = ns_liq_final

    result["NS Equity Required / Month"] = (
        (ns_eq_final - ns_eq_ach).clip(lower=0) / months
    )
    result["NS Debt Required / Month"] = (
        (ns_dt_final - ns_dt_ach).clip(lower=0) / months
    )
    result["NS Liquid Required / Month"] = (
        (ns_liq_final - ns_liq_ach).clip(lower=0) / months
    )

    result["NS Total Required / Month"] = (
        result["NS Equity Required / Month"]
        + result["NS Debt Required / Month"]
        + result["NS Liquid Required / Month"]
    )

    ns_fy_total = (
        clean_numeric(df["FY 26 TGT EQ NS"])
        + clean_numeric(df["FY 26 TGT DT NS"])
        + clean_numeric(df["FY 26 TGT LIQ NS"])
    )
    ns_final_total = ns_eq_final + ns_dt_final + ns_liq_final
    result["NS Scenario Overall %"] = np.where(
        ns_fy_total > 0, ns_final_total / ns_fy_total, 0
    )

    result["Scenario"] = scenario
    result["Months Completed"] = months_completed
    result["Months Remaining"] = months
    return result


def calculate_all_scenarios(
    df: pd.DataFrame,
    months_remaining: int,
    assumptions: Dict[str, float],
) -> pd.DataFrame:
    frames = [
        calculate_scenario(df, scenario, months_remaining, assumptions)
        for scenario in SCENARIO_NAMES
    ]
    return pd.concat(frames, ignore_index=True)


def make_download_excel(
    original_df: pd.DataFrame,
    selected_result: pd.DataFrame,
    all_results: pd.DataFrame,
    assumptions: Dict[str, float],
    selected_scenario: str,
    months_remaining: int,
) -> bytes:
    output = io.BytesIO()
    months_completed = max(12 - int(months_remaining), 1)

    control_df = pd.DataFrame(
        {
            "Setting": [
                "Selected Case",
                "Months Completed",
                "Months Remaining",
                "Scenario 1 Overall Target",
                "Scenario 2 Overall Target",
                "Scenario 3 Equity Target",
                "Scenario 3 Overall Target",
                "Scenario 4 Equity Target",
                "Scenario 4 Minimum Overall",
            ],
            "Value": [
                selected_scenario,
                months_completed,
                months_remaining,
                assumptions["scenario1_overall"],
                assumptions["scenario2_overall"],
                assumptions["scenario3_equity"],
                assumptions["scenario3_overall"],
                assumptions["scenario4_equity"],
                assumptions["scenario4_overall_min"],
            ],
        }
    )

    if selected_scenario == FORECAST_NAME:
        gs_target = selected_result["GS Total FY Target"].sum()
        gs_projected = selected_result["GS Total Projected FY Achievement"].sum()
        ns_target = selected_result["NS Total FY Target"].sum()
        ns_projected = selected_result["NS Total Projected FY Achievement"].sum()
        gs_pct = gs_projected / gs_target if gs_target > 0 else 0.0
        ns_pct = ns_projected / ns_target if ns_target > 0 else 0.0

        summary_df = pd.DataFrame(
            {
                "Metric": [
                    "Employees",
                    "Selected Case",
                    "Months Completed",
                    "GS Projected FY Achievement %",
                    "NS Projected FY Achievement %",
                ],
                "Value": [
                    len(selected_result),
                    selected_scenario,
                    months_completed,
                    gs_pct,
                    ns_pct,
                ],
            }
        )
    else:
        summary_df = pd.DataFrame(
            {
                "Metric": [
                    "Employees",
                    "Selected Scenario",
                    "Total GS Required / Month",
                    "Total NS Required / Month",
                ],
                "Value": [
                    len(selected_result),
                    selected_scenario,
                    selected_result["GS Total Required / Month"].sum(),
                    selected_result["NS Total Required / Month"].sum(),
                ],
            }
        )

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        original_df.to_excel(writer, sheet_name="Uploaded Data", index=False)
        control_df.to_excel(writer, sheet_name="Scenario Control", index=False)
        summary_df.to_excel(writer, sheet_name="Scenario Summary", index=False)
        selected_result.to_excel(writer, sheet_name="Selected Case", index=False)
        all_results.to_excel(writer, sheet_name="All Cases", index=False)

        for sheet_name in writer.book.sheetnames:
            ws = writer.book[sheet_name]
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions

            for cell in ws[1]:
                cell.font = cell.font.copy(bold=True)

            for column_cells in ws.columns:
                max_len = 0
                column_letter = column_cells[0].column_letter
                for cell in column_cells[:100]:
                    value = "" if cell.value is None else str(cell.value)
                    max_len = max(max_len, len(value))
                ws.column_dimensions[column_letter].width = min(
                    max(max_len + 2, 12), 34
                )

    output.seek(0)
    return output.getvalue()


def display_result_table(result: pd.DataFrame, mode: str) -> None:
    is_forecast = (
        "Scenario" in result.columns
        and not result.empty
        and result["Scenario"].iloc[0] == FORECAST_NAME
    )

    if is_forecast and mode == "Gross Sales":
        desired = [
            *METADATA_COLUMNS,
            "GS Equity FY Target",
            "GS Equity YTD Achievement",
            "GS Equity Current Run Rate / Month",
            "GS Equity Projected FY Achievement",
            "GS Equity Projected FY %",
            "GS Debt FY Target",
            "GS Debt YTD Achievement",
            "GS Debt Current Run Rate / Month",
            "GS Debt Projected FY Achievement",
            "GS Debt Projected FY %",
            "GS Liquid FY Target",
            "GS Liquid YTD Achievement",
            "GS Liquid Current Run Rate / Month",
            "GS Liquid Projected FY Achievement",
            "GS Liquid Projected FY %",
            "GS Total Current Run Rate / Month",
            "GS Total Projected FY Achievement",
            "GS Overall Projected FY %",
        ]
    elif is_forecast:
        desired = [
            *METADATA_COLUMNS,
            "NS Equity FY Target",
            "NS Equity YTD Achievement",
            "NS Equity Current Run Rate / Month",
            "NS Equity Projected FY Achievement",
            "NS Equity Projected FY %",
            "NS Debt FY Target",
            "NS Debt YTD Achievement",
            "NS Debt Current Run Rate / Month",
            "NS Debt Projected FY Achievement",
            "NS Debt Projected FY %",
            "NS Liquid FY Target",
            "NS Liquid YTD Achievement",
            "NS Liquid Current Run Rate / Month",
            "NS Liquid Projected FY Achievement",
            "NS Liquid Projected FY %",
            "NS Total Current Run Rate / Month",
            "NS Total Projected FY Achievement",
            "NS Overall Projected FY %",
        ]
    elif mode == "Gross Sales":
        desired = [
            *METADATA_COLUMNS,
            "GS Equity Final Target",
            "GS Equity Required / Month",
            "GS Debt Final Target",
            "GS Debt Required / Month",
            "GS Liquid Final Target",
            "GS Liquid Required / Month",
            "GS Total Required / Month",
            "GS Scenario Overall %",
        ]
    else:
        desired = [
            *METADATA_COLUMNS,
            "NS Equity Final Target",
            "NS Equity Required / Month",
            "NS Debt Final Target",
            "NS Debt Required / Month",
            "NS Liquid Final Target",
            "NS Liquid Required / Month",
            "NS Total Required / Month",
            "NS Scenario Overall %",
        ]

    columns = [c for c in desired if c in result.columns]
    display_df = result[columns].copy()

    format_map = {}
    for col in display_df.columns:
        if col.endswith("%"):
            format_map[col] = st.column_config.NumberColumn(col, format="%.1f%%")
            display_df[col] = display_df[col] * 100
        elif "Target" in col or "Achievement" in col or "/ Month" in col:
            format_map[col] = st.column_config.NumberColumn(col, format="%,.0f")

    st.dataframe(
        display_df,
        use_container_width=True,
        hide_index=True,
        column_config=format_map,
        height=520,
    )



# -------------------------------------------------------------------
# Five-sheet workbook dashboard helpers
# -------------------------------------------------------------------
REQUIRED_WORKBOOK_SHEETS = [
    "Summary",
    "Summary - Achievement",
    "RM Retail Sales",
    "RM DHNI",
    "VRM",
]

VERTICAL_SHEETS = {
    "Retail": "RM Retail Sales",
    "DHNI": "RM DHNI",
    "VRM": "VRM",
}

ACHIEVEMENT_BANDS = [
    "Greater than 100%",
    "80% - 100%",
    "50% - 80%",
    "30% - 50%",
    "below 30%",
    "NA",
]


def compact_raw_sheet(df: pd.DataFrame) -> pd.DataFrame:
    """Remove completely blank rows/columns while preserving workbook layout values."""
    if df is None or df.empty:
        return pd.DataFrame()
    out = df.copy()
    out = out.dropna(axis=0, how="all").dropna(axis=1, how="all")
    return out


def duplicate_family_columns(df: pd.DataFrame, base_name: str) -> list:
    """Return repeated Excel headers in left-to-right order (e.g. Scenario 1, Scenario 1.1...)."""
    pattern = re.compile(rf"^{re.escape(base_name)}(?:\.\d+)?$")
    return [str(c) for c in df.columns if pattern.match(str(c))]


def safe_sum(df: pd.DataFrame, column: str) -> float:
    if column not in df.columns:
        return 0.0
    return float(clean_numeric(df[column]).sum())


def scenario_component_sums_from_sheet(
    df: pd.DataFrame,
    assumptions: Dict[str, float],
) -> Dict[str, Dict[str, Dict[str, float]]]:
    """
    Aggregate the Scenario 1/2/3 run-rate columns already present in a detailed sheet.

    The employee workbook repeats Scenario 1/2/3 once for each component in this order:
    GS Equity, GS Debt, GS Liquid, NS Equity, NS Debt, NS Liquid.
    If those repeated source columns are unavailable, fall back to the app's calculation
    engine so the dashboard remains usable.
    """
    result = {
        "GS": {asset: {} for asset in ["Equity", "Debt", "Liquid"]},
        "NS": {asset: {} for asset in ["Equity", "Debt", "Liquid"]},
    }
    mapping = [
        ("GS", "Equity"),
        ("GS", "Debt"),
        ("GS", "Liquid"),
        ("NS", "Equity"),
        ("NS", "Debt"),
        ("NS", "Liquid"),
    ]

    source_ok = True
    scenario_families = {}
    for scenario in ["Scenario 1", "Scenario 2", "Scenario 3"]:
        cols = duplicate_family_columns(df, scenario)
        scenario_families[scenario] = cols
        if len(cols) < 6:
            source_ok = False

    if source_ok:
        for scenario, cols in scenario_families.items():
            for idx, (mode, asset) in enumerate(mapping):
                result[mode][asset][scenario] = safe_sum(df, cols[idx])
        return result

    # Fallback: use the application's scenario engine. Scenario 1/3 use 9 months;
    # Scenario 2 uses 7 months to mirror the workbook Summary horizon headings.
    fallback_months = {"Scenario 1": 9, "Scenario 2": 7, "Scenario 3": 9}
    result_columns = {
        "GS": {
            "Equity": "GS Equity Required / Month",
            "Debt": "GS Debt Required / Month",
            "Liquid": "GS Liquid Required / Month",
        },
        "NS": {
            "Equity": "NS Equity Required / Month",
            "Debt": "NS Debt Required / Month",
            "Liquid": "NS Liquid Required / Month",
        },
    }
    for scenario, months in fallback_months.items():
        calculated = calculate_scenario(df, scenario, months, assumptions)
        for mode in ["GS", "NS"]:
            for asset in ["Equity", "Debt", "Liquid"]:
                column = result_columns[mode][asset]
                result[mode][asset][scenario] = float(calculated[column].sum())
    return result


def build_sales_summary(
    vertical_frames: Dict[str, pd.DataFrame],
    mode: str,
    months_completed: int,
    assumptions: Dict[str, float],
) -> pd.DataFrame:
    """Build the Gross Sales / Net Sales vertical summary shown in the workbook dashboard."""
    completed = max(int(months_completed), 1)

    if mode == "GS":
        target_cols = {
            "Equity": "FY 26 TGT EQ",
            "Debt": "FY 26 TGT DT",
            "Liquid": "FY 26 TGT LIQ",
        }
        ach_cols = {
            "Equity": "Equity GS Ach YTD June",
            "Debt": "Debt GS Ach",
            "Liquid": "Liquid GS Ach",
        }
    else:
        target_cols = {
            "Equity": "FY 26 TGT EQ NS",
            "Debt": "FY 26 TGT DT NS",
            "Liquid": "FY 26 TGT LIQ NS",
        }
        ach_cols = {
            "Equity": "Equity NS Ach YTD June",
            "Debt": "Debt NS Ach",
            "Liquid": "Liquid NS Ach",
        }

    rows = []
    for vertical, df in vertical_frames.items():
        scenario_sums = scenario_component_sums_from_sheet(df, assumptions)

        target_by_asset = {asset: safe_sum(df, col) for asset, col in target_cols.items()}
        actual_by_asset = {asset: safe_sum(df, col) for asset, col in ach_cols.items()}
        current_by_asset = {
            asset: actual_by_asset[asset] / completed for asset in ["Equity", "Debt", "Liquid"]
        }

        row = {
            ("Overall", "Total FY Target"): sum(target_by_asset.values()),
            ("Overall", "Q1 Actuals"): sum(actual_by_asset.values()),
            ("Overall", "Q1 Actual Run Rate"): sum(actual_by_asset.values()) / completed,
        }

        for scenario in ["Scenario 1", "Scenario 2", "Scenario 3"]:
            row[("Overall", scenario)] = sum(
                scenario_sums[mode][asset][scenario] for asset in ["Equity", "Debt", "Liquid"]
            )

        for asset in ["Equity", "Debt", "Liquid"]:
            row[(asset, "Current")] = current_by_asset[asset]
            for scenario in ["Scenario 1", "Scenario 2", "Scenario 3"]:
                row[(asset, scenario)] = scenario_sums[mode][asset][scenario]

        rows.append((vertical, row))

    summary = pd.DataFrame([row for _, row in rows], index=[v for v, _ in rows])
    summary.index.name = "Vertical"
    summary.columns = pd.MultiIndex.from_tuples(summary.columns)

    if not summary.empty:
        summary.loc["Total"] = summary.sum(axis=0, numeric_only=True)

    ordered_columns = [
        ("Overall", "Total FY Target"),
        ("Overall", "Q1 Actuals"),
        ("Overall", "Q1 Actual Run Rate"),
        ("Overall", "Scenario 1"),
        ("Overall", "Scenario 2"),
        ("Overall", "Scenario 3"),
    ]
    for asset in ["Equity", "Debt", "Liquid"]:
        ordered_columns.extend(
            [
                (asset, "Current"),
                (asset, "Scenario 1"),
                (asset, "Scenario 2"),
                (asset, "Scenario 3"),
            ]
        )
    return summary.reindex(columns=pd.MultiIndex.from_tuples(ordered_columns))


def achievement_percentage(
    df: pd.DataFrame,
    target_column: str,
    achievement_column: str,
    explicit_percent_column: str | None = None,
) -> pd.Series:
    """Return decimal achievement ratios; NaN means the percentage is not applicable."""
    if explicit_percent_column and explicit_percent_column in df.columns:
        raw = pd.to_numeric(df[explicit_percent_column], errors="coerce")
        # Excel percentages are usually decimals, but tolerate files containing 80 instead of 0.80.
        non_null = raw.dropna()
        if not non_null.empty and non_null.abs().median() > 3:
            raw = raw / 100.0
        return raw

    if target_column not in df.columns or achievement_column not in df.columns:
        return pd.Series(np.nan, index=df.index, dtype=float)

    target = pd.to_numeric(df[target_column], errors="coerce")
    achievement = pd.to_numeric(df[achievement_column], errors="coerce")
    valid = target.notna() & (target > 0) & achievement.notna()
    ratio = pd.Series(np.nan, index=df.index, dtype=float)
    ratio.loc[valid] = achievement.loc[valid] / target.loc[valid]
    return ratio


def classify_achievement(value: float) -> str:
    if pd.isna(value):
        return "NA"
    if value > 1.0:
        return "Greater than 100%"
    if value >= 0.80:
        return "80% - 100%"
    if value >= 0.50:
        return "50% - 80%"
    if value >= 0.30:
        return "30% - 50%"
    return "below 30%"


def build_achievement_summary(
    vertical_frames: Dict[str, pd.DataFrame],
    mode: str,
) -> pd.DataFrame:
    """Count brokers in the requested achievement bands for Equity, Debt and Liquid."""
    if mode == "GS":
        specs = {
            "Equity": ("YTD June EQ TGT", "Equity GS Ach YTD June", "Equity GS % Ach YTD June"),
            "Debt": ("YTD June DT TGT", "Debt GS Ach", "Debt GS Ach%"),
            "Liquid": ("YTD June LIQ TGT", "Liquid GS Ach", "Liquid GS Ach %"),
        }
    else:
        specs = {
            "Equity": ("YTD June EQ NS TGT", "Equity NS Ach YTD June", None),
            "Debt": ("YTD June DT NS TGT", "Debt NS Ach", None),
            "Liquid": ("YTD June LIQ NS TGT", "Liquid NS Ach", None),
        }

    output = pd.DataFrame(index=ACHIEVEMENT_BANDS)
    output.index.name = "Achievement"

    for vertical, df in vertical_frames.items():
        employee_mask = pd.Series(True, index=df.index)
        if "Employee Name" in df.columns:
            names = df["Employee Name"].astype(str).str.strip()
            employee_mask = df["Employee Name"].notna() & names.ne("") & names.ne("nan")

        for asset, (target_col, achievement_col, pct_col) in specs.items():
            ratios = achievement_percentage(df, target_col, achievement_col, pct_col)
            bands = ratios.loc[employee_mask].map(classify_achievement)
            counts = bands.value_counts()
            output[(vertical, asset)] = [int(counts.get(band, 0)) for band in ACHIEVEMENT_BANDS]

    output.columns = pd.MultiIndex.from_tuples(output.columns)

    # Gross Sales in the user's workbook normally has no NA row. Hide it when it is zero everywhere.
    if mode == "GS" and "NA" in output.index and int(output.loc["NA"].sum()) == 0:
        output = output.drop(index="NA")

    output.loc["Total"] = output.sum(axis=0, numeric_only=True)
    return output


def flatten_multiindex_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Flatten MultiIndex columns for Excel export while keeping the Streamlit view grouped."""
    out = df.copy()
    if isinstance(out.columns, pd.MultiIndex):
        out.columns = [" | ".join(str(x) for x in col if str(x) != "") for col in out.columns]
    return out.reset_index()


def make_multisheet_download_excel(
    raw_summary: pd.DataFrame,
    raw_summary_achievement: pd.DataFrame,
    vertical_frames: Dict[str, pd.DataFrame],
    gross_summary: pd.DataFrame,
    net_summary: pd.DataFrame,
    gross_achievement: pd.DataFrame,
    net_achievement: pd.DataFrame,
    selected_scenario: str,
    months_remaining: int,
    assumptions: Dict[str, float],
) -> bytes:
    """Export source views plus recalculated dashboard tables and all existing case calculations."""
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        compact_raw_sheet(raw_summary).to_excel(
            writer, sheet_name="Source Summary", index=False, header=False
        )
        compact_raw_sheet(raw_summary_achievement).to_excel(
            writer, sheet_name="Source Summary-Ach", index=False, header=False
        )
        flatten_multiindex_columns(gross_summary).to_excel(
            writer, sheet_name="Calc Gross Summary", index=False
        )
        flatten_multiindex_columns(net_summary).to_excel(
            writer, sheet_name="Calc Net Summary", index=False
        )
        flatten_multiindex_columns(gross_achievement).to_excel(
            writer, sheet_name="Calc GS Achievement", index=False
        )
        flatten_multiindex_columns(net_achievement).to_excel(
            writer, sheet_name="Calc NS Achievement", index=False
        )

        for vertical, df in vertical_frames.items():
            safe_vertical = vertical[:20]
            df.to_excel(writer, sheet_name=f"{safe_vertical} Data"[:31], index=False)
            selected = calculate_scenario(df, selected_scenario, months_remaining, assumptions)
            selected.to_excel(
                writer, sheet_name=f"{safe_vertical} Selected"[:31], index=False
            )
            all_cases = calculate_all_scenarios(df, months_remaining, assumptions)
            all_cases.to_excel(
                writer, sheet_name=f"{safe_vertical} All Cases"[:31], index=False
            )

        for ws in writer.book.worksheets:
            ws.freeze_panes = "A2"
            if ws.max_row >= 1 and ws.max_column >= 1:
                ws.auto_filter.ref = ws.dimensions
            for cell in ws[1]:
                cell.font = cell.font.copy(bold=True)
            for column_cells in ws.columns:
                max_len = 0
                column_letter = column_cells[0].column_letter
                for cell in column_cells[:100]:
                    value = "" if cell.value is None else str(cell.value)
                    max_len = max(max_len, len(value))
                ws.column_dimensions[column_letter].width = min(max(max_len + 2, 12), 34)

    output.seek(0)
    return output.getvalue()


def render_case_explanation(
    selected_scenario: str,
    months_completed: int,
    months_remaining: int,
    s1: float,
    s2: float,
    s3_eq: float,
    s3_overall: float,
    s4_eq: float,
    s4_overall: float,
) -> None:
    if selected_scenario == FORECAST_NAME:
        st.info(
            f"**Current Run Rate Forecast:** Uses performance from the first {months_completed} completed month(s) to calculate the current monthly pace.  \n"
            "If the same pace continues for the full year, it projects the FY achievement percentage; no target scenario is imposed."
        )
    elif selected_scenario == "Scenario 1":
        st.info(
            f"**Scenario 1:** Assumes the employee should finish the year at {s1:.0f}% of the overall FY target.  \n"
            f"The app calculates the Equity, Debt and Liquid monthly requirement over the remaining {months_remaining} month(s)."
        )
    elif selected_scenario == "Scenario 2":
        st.info(
            f"**Scenario 2:** Assumes the employee should finish the year at {s2:.0f}% of the overall FY target.  \n"
            "The app increases the required run rate accordingly and shows the monthly Equity, Debt and Liquid requirement."
        )
    elif selected_scenario == "Scenario 3":
        st.info(
            f"**Scenario 3:** Fixes Equity achievement at {s3_eq:.0f}% while targeting {s3_overall:.0f}% overall FY achievement.  \n"
            "The remaining requirement is distributed between Debt and Liquid in proportion to their FY targets."
        )
    else:
        st.info(
            f"**Scenario 4:** Fixes Equity achievement at {s4_eq:.0f}% and requires at least {s4_overall:.0f}% overall FY achievement.  \n"
            "Debt and Liquid are adjusted proportionally so the employee reaches the minimum overall target."
        )


def render_vertical_dashboard(
    vertical: str,
    df: pd.DataFrame,
    selected_scenario: str,
    months_remaining: int,
    assumptions: Dict[str, float],
    s1: float,
    s2: float,
    s3_eq: float,
    s3_overall: float,
    s4_eq: float,
    s4_overall: float,
) -> None:
    valid, missing_columns = validate_input(df)
    if not valid:
        st.error(
            f"{vertical} is missing required calculation columns:\n\n"
            + "\n".join(f"- {c}" for c in missing_columns)
        )
        st.dataframe(df, use_container_width=True, hide_index=True, height=520)
        return

    selected_result = calculate_scenario(df, selected_scenario, months_remaining, assumptions)
    employees_count = len(selected_result)
    months_completed = max(12 - int(months_remaining), 1)

    st.subheader(f"{vertical} — Employee Performance")
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Selected Case", selected_scenario)
    m2.metric("Employees", f"{employees_count:,}")

    if selected_scenario == FORECAST_NAME:
        gs_target_total = selected_result["GS Total FY Target"].sum()
        gs_projected_total = selected_result["GS Total Projected FY Achievement"].sum()
        ns_target_total = selected_result["NS Total FY Target"].sum()
        ns_projected_total = selected_result["NS Total Projected FY Achievement"].sum()
        gs_pct = gs_projected_total / gs_target_total if gs_target_total > 0 else 0.0
        ns_pct = ns_projected_total / ns_target_total if ns_target_total > 0 else 0.0
        m3.metric("GS Projected FY %", f"{gs_pct * 100:.1f}%")
        m4.metric("NS Projected FY %", f"{ns_pct * 100:.1f}%")
    else:
        m3.metric("GS Required / Month", f'{selected_result["GS Total Required / Month"].sum():,.0f}')
        m4.metric("NS Required / Month", f'{selected_result["NS Total Required / Month"].sum():,.0f}')

    render_case_explanation(
        selected_scenario,
        months_completed,
        months_remaining,
        s1,
        s2,
        s3_eq,
        s3_overall,
        s4_eq,
        s4_overall,
    )

    gs_tab, ns_tab, raw_tab = st.tabs(["Gross Sales", "Net Sales", "Uploaded Data"])
    with gs_tab:
        display_result_table(selected_result, "Gross Sales")
    with ns_tab:
        display_result_table(selected_result, "Net Sales")
    with raw_tab:
        st.dataframe(df, use_container_width=True, hide_index=True, height=520)

    st.markdown("#### Employee drill-down")
    employee_options = selected_result["Employee Name"].dropna().astype(str).tolist()
    if not employee_options:
        st.info("No employee names are available for drill-down.")
        return

    employee = st.selectbox(
        "Choose employee",
        options=employee_options,
        key=f"employee_{vertical}",
    )
    emp = selected_result[selected_result["Employee Name"].astype(str) == employee].iloc[0]
    c1, c2 = st.columns(2)

    if selected_scenario == FORECAST_NAME:
        with c1:
            st.markdown("**Gross Sales projected FY achievement**")
            st.write(
                {
                    "Equity": f'{emp["GS Equity Projected FY %"] * 100:.1f}%',
                    "Debt": f'{emp["GS Debt Projected FY %"] * 100:.1f}%',
                    "Liquid": f'{emp["GS Liquid Projected FY %"] * 100:.1f}%',
                    "Overall": f'{emp["GS Overall Projected FY %"] * 100:.1f}%',
                }
            )
            st.caption(f'Current GS run rate: {emp["GS Total Current Run Rate / Month"]:,.2f} per month')
        with c2:
            st.markdown("**Net Sales projected FY achievement**")
            st.write(
                {
                    "Equity": f'{emp["NS Equity Projected FY %"] * 100:.1f}%',
                    "Debt": f'{emp["NS Debt Projected FY %"] * 100:.1f}%',
                    "Liquid": f'{emp["NS Liquid Projected FY %"] * 100:.1f}%',
                    "Overall": f'{emp["NS Overall Projected FY %"] * 100:.1f}%',
                }
            )
            st.caption(f'Current NS run rate: {emp["NS Total Current Run Rate / Month"]:,.2f} per month')
    else:
        with c1:
            st.markdown("**Gross Sales monthly requirement**")
            st.write(
                {
                    "Equity": round(emp["GS Equity Required / Month"], 2),
                    "Debt": round(emp["GS Debt Required / Month"], 2),
                    "Liquid": round(emp["GS Liquid Required / Month"], 2),
                    "Total": round(emp["GS Total Required / Month"], 2),
                }
            )
        with c2:
            st.markdown("**Net Sales monthly requirement**")
            st.write(
                {
                    "Equity": round(emp["NS Equity Required / Month"], 2),
                    "Debt": round(emp["NS Debt Required / Month"], 2),
                    "Liquid": round(emp["NS Liquid Required / Month"], 2),
                    "Total": round(emp["NS Total Required / Month"], 2),
                }
            )


# -------------------------------------------------------------------
# UI
# -------------------------------------------------------------------
st.title("Employee Performance Forecast & Scenario Planner")
st.caption(
    "Upload the five-sheet workbook to view management summaries, achievement-band counts, "
    "and the existing employee-level forecast/scenario analysis for Retail, DHNI and VRM."
)

uploaded_file = st.file_uploader(
    "Upload employee performance Excel",
    type=["xlsx", "xlsm"],
    help=(
        "Expected sheets: Summary, Summary - Achievement, RM Retail Sales, RM DHNI and VRM. "
        "All five are read by the application."
    ),
)

with st.expander("Expected detailed-sheet columns"):
    st.write(
        "RM Retail Sales, RM DHNI and VRM use the same employee-level structure. "
        "Extra columns are allowed."
    )
    st.code("\n".join(EXPECTED_INPUT_COLUMNS), language=None)
    st.markdown("**Mandatory columns used by the employee calculations**")
    st.code("\n".join(REQUIRED_COLUMNS), language=None)

if uploaded_file is None:
    st.info("Upload the five-sheet Excel workbook to start.")
    st.stop()

try:
    xls = pd.ExcelFile(uploaded_file, engine="openpyxl")
except Exception as exc:
    st.error(f"Could not read the Excel file: {exc}")
    st.stop()

missing_sheets = [name for name in REQUIRED_WORKBOOK_SHEETS if name not in xls.sheet_names]
if missing_sheets:
    st.error(
        "The workbook is missing these required sheets:\n\n"
        + "\n".join(f"- {name}" for name in missing_sheets)
        + "\n\nSheets found: "
        + ", ".join(xls.sheet_names)
    )
    st.stop()

# Read ALL five sheets. Summary sheets are loaded raw to preserve their multi-row/merged layout.
raw_summary = pd.read_excel(xls, sheet_name="Summary", header=None)
raw_summary_achievement = pd.read_excel(xls, sheet_name="Summary - Achievement", header=None)

vertical_frames = {}
for vertical, sheet in VERTICAL_SHEETS.items():
    detail_df = pd.read_excel(xls, sheet_name=sheet)
    vertical_frames[vertical] = normalize_input_columns(detail_df)

st.success(
    "Loaded all five sheets: Summary, Summary - Achievement, RM Retail Sales, RM DHNI and VRM."
)

# ---------------- Global controls retained from the earlier app ----------------
st.sidebar.header("Forecast / Scenario Controls")
selected_scenario = st.sidebar.selectbox("Selected Case", SCENARIO_NAMES)
months_remaining = st.sidebar.number_input(
    "Months Remaining",
    min_value=1,
    max_value=11,
    value=9,
    step=1,
    help=(
        "For Q1 / YTD June, keep 9 months remaining = 3 months completed. "
        "This drives the Current Run Rate Forecast and detailed scenario calculations."
    ),
)

st.sidebar.subheader("Current Run Rate Forecast")
st.sidebar.caption(
    "No target assumption is imposed. Actual YTD performance is annualised at the current monthly pace."
)

s1 = 100.0
s2 = 120.0
s3_eq = 100.0
s3_overall = 120.0
s4_eq = 110.0
s4_overall = 85.0

if selected_scenario == "Scenario 1":
    st.sidebar.subheader("Scenario 1")
    s1 = st.sidebar.number_input(
        "Overall target %", 0.0, 300.0, s1, 5.0, key="s1"
    )
elif selected_scenario == "Scenario 2":
    st.sidebar.subheader("Scenario 2")
    s2 = st.sidebar.number_input(
        "Overall target %", 0.0, 300.0, s2, 5.0, key="s2"
    )
elif selected_scenario == "Scenario 3":
    st.sidebar.subheader("Scenario 3")
    s3_eq = st.sidebar.number_input(
        "Equity target %", 0.0, 300.0, s3_eq, 5.0, key="s3_eq"
    )
    s3_overall = st.sidebar.number_input(
        "Overall target %", 0.0, 300.0, s3_overall, 5.0, key="s3_overall"
    )
elif selected_scenario == "Scenario 4":
    st.sidebar.subheader("Scenario 4")
    s4_eq = st.sidebar.number_input(
        "Equity target %", 0.0, 300.0, s4_eq, 5.0, key="s4_eq"
    )
    s4_overall = st.sidebar.number_input(
        "Minimum overall target %", 0.0, 300.0, s4_overall, 5.0, key="s4_overall"
    )

assumptions = {
    "scenario1_overall": s1 / 100,
    "scenario2_overall": s2 / 100,
    "scenario3_equity": s3_eq / 100,
    "scenario3_overall": s3_overall / 100,
    "scenario4_equity": s4_eq / 100,
    "scenario4_overall_min": s4_overall / 100,
}

months_completed = max(12 - int(months_remaining), 1)

# Validate detailed sheets, but keep the dashboard visible so problems are sheet-specific.
validation_issues = {}
for vertical, frame in vertical_frames.items():
    valid, missing = validate_input(frame)
    if not valid:
        validation_issues[vertical] = missing

if validation_issues:
    with st.expander("Detailed-sheet validation warnings", expanded=True):
        for vertical, missing in validation_issues.items():
            st.warning(f"{vertical}: missing " + ", ".join(missing))

# New management summaries.
valid_vertical_frames = {
    vertical: frame
    for vertical, frame in vertical_frames.items()
    if vertical not in validation_issues
}

gross_summary = build_sales_summary(valid_vertical_frames, "GS", months_completed, assumptions)
net_summary = build_sales_summary(valid_vertical_frames, "NS", months_completed, assumptions)
gross_achievement = build_achievement_summary(valid_vertical_frames, "GS")
net_achievement = build_achievement_summary(valid_vertical_frames, "NS")

summary_tab, achievement_tab, retail_tab, dhni_tab, vrm_tab = st.tabs(
    ["Summary", "Summary - Achievement", "RM Retail Sales", "RM DHNI", "VRM"]
)

with summary_tab:
    st.header("Summary")
    st.caption(
        "Gross Sales and Net Sales below are recalculated from RM Retail Sales, RM DHNI and VRM. "
        "The workbook's original Summary is also shown so Gross SIP+STP, Net SIP+STP and Alternate remain visible."
    )

    st.subheader("Gross Sales")
    st.dataframe(
        gross_summary.style.format("{:,.0f}"),
        use_container_width=True,
        height=260,
    )
    st.caption(
        f"Current = Q1 actual run rate using {months_completed} completed month(s). "
        "Scenario 1/2/3 use the scenario run-rate columns already present in each detailed sheet."
    )

    st.subheader("Net Sales")
    st.dataframe(
        net_summary.style.format("{:,.0f}"),
        use_container_width=True,
        height=260,
    )

    st.markdown("#### Source Summary sheet — includes SIP+STP and Alternate")
    source_summary_view = compact_raw_sheet(raw_summary)
    st.dataframe(source_summary_view, use_container_width=True, hide_index=True, height=620)

with achievement_tab:
    st.header("Summary - Achievement — Number of Brokers")
    st.caption(
        "Broker counts are recalculated from the three detailed vertical sheets using YTD achievement percentages. "
        "If a Net Sales YTD target is missing or non-positive, that broker/component is classified as NA."
    )

    st.subheader("Gross Sales")
    st.dataframe(gross_achievement, use_container_width=True, height=310)

    st.subheader("Net Sales")
    st.dataframe(net_achievement, use_container_width=True, height=340)

    with st.expander("Source Summary - Achievement sheet"):
        st.dataframe(
            compact_raw_sheet(raw_summary_achievement),
            use_container_width=True,
            hide_index=True,
            height=620,
        )

with retail_tab:
    render_vertical_dashboard(
        "Retail",
        vertical_frames["Retail"],
        selected_scenario,
        months_remaining,
        assumptions,
        s1,
        s2,
        s3_eq,
        s3_overall,
        s4_eq,
        s4_overall,
    )

with dhni_tab:
    render_vertical_dashboard(
        "DHNI",
        vertical_frames["DHNI"],
        selected_scenario,
        months_remaining,
        assumptions,
        s1,
        s2,
        s3_eq,
        s3_overall,
        s4_eq,
        s4_overall,
    )

with vrm_tab:
    render_vertical_dashboard(
        "VRM",
        vertical_frames["VRM"],
        selected_scenario,
        months_remaining,
        assumptions,
        s1,
        s2,
        s3_eq,
        s3_overall,
        s4_eq,
        s4_overall,
    )

# ---------------- Download: source + recalculated summaries + previous case outputs ----------------
excel_bytes = make_multisheet_download_excel(
    raw_summary=raw_summary,
    raw_summary_achievement=raw_summary_achievement,
    vertical_frames=valid_vertical_frames,
    gross_summary=gross_summary,
    net_summary=net_summary,
    gross_achievement=gross_achievement,
    net_achievement=net_achievement,
    selected_scenario=selected_scenario,
    months_remaining=months_remaining,
    assumptions=assumptions,
)

st.download_button(
    "Download five-sheet dashboard calculations",
    data=excel_bytes,
    file_name="employee_5sheet_dashboard_results.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True,
)