import re
from pathlib import Path

import numpy as np
import pandas as pd
import streamlit as st


st.set_page_config(page_title="GWP Analysis for WtE Recovery Processes", layout="wide")


# -----------------------------
# Constants and emission factors
# -----------------------------
GWP_CH4 = 23
GWP_N2O = 296

# Energy-related emission factors from your text
# Direct and indirect factors
# Units:
# CO2 direct, CH4 direct = g/MJ
# N2O direct = mg/MJ
# CO2 indirect, CH4 indirect = g/MJ
# N2O indirect = mg/MJ
ENERGY_FACTORS = {
    "Steam": {
        "co2_direct_g_mj": 113.87,
        "ch4_direct_g_mj": 0.29,
        "n2o_direct_mg_mj": 1.79,
        "co2_indirect_g_mj": 0.0,
        "ch4_indirect_g_mj": 0.0,
        "n2o_indirect_mg_mj": 0.0,
    },
    "Electricity": {
        "co2_direct_g_mj": 248.02,
        "ch4_direct_g_mj": 2.16,
        "n2o_direct_mg_mj": 0.62,
        "co2_indirect_g_mj": 0.0,
        "ch4_indirect_g_mj": 0.0,
        "n2o_indirect_mg_mj": 0.0,
    },
    "Natural gas": {
        "co2_direct_g_mj": 16.58,
        "ch4_direct_g_mj": 0.05,
        "n2o_direct_mg_mj": 0.12,
        "co2_indirect_g_mj": 55.612,
        "ch4_indirect_g_mj": 0.001,
        "n2o_indirect_mg_mj": 0.001,
    },
    "Gasoline": {
        "co2_direct_g_mj": 67.914,
        "ch4_direct_g_mj": 28.83,
        "n2o_direct_mg_mj": 0.08,
        "co2_indirect_g_mj": 0.09,
        "ch4_indirect_g_mj": 0.002,
        "n2o_indirect_mg_mj": 0.47,
    },
    "Diesel": {
        "co2_direct_g_mj": 72.585,
        "ch4_direct_g_mj": 27.87,
        "n2o_direct_mg_mj": 0.004,
        "co2_indirect_g_mj": 0.08,
        "ch4_indirect_g_mj": 0.028,
        "n2o_indirect_mg_mj": 0.44,
    },
}

TRANSPORT_MJ_PER_TKM = 1362 / 1000  # 1362 kJ/t·km = 1.362 MJ/t·km


# -----------------------------
# Utility functions
# -----------------------------
def clean_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df


def safe_numeric(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce")


def normalize_name(text: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(text).lower())


def detect_column(df: pd.DataFrame, keywords):
    norm_map = {c: normalize_name(c) for c in df.columns}
    for col, norm in norm_map.items():
        if all(k in norm for k in keywords):
            return col
    return None


def detect_process_column(df: pd.DataFrame):
    for candidate in df.columns:
        values = df[candidate].astype(str).str.strip()
        if values.str.contains(r"Process\s*[-–]?\s*\d+", case=False, regex=True).any():
            return candidate
    return None


def find_process_blocks(df: pd.DataFrame):
    """
    Detect rows like 'Process-1', 'Process 2', etc.
    Returns list of (process_name, start_idx, end_idx)
    """
    proc_col = detect_process_column(df)
    if proc_col is None:
        return []

    markers = []
    for idx, val in df[proc_col].astype(str).items():
        if re.search(r"Process\s*[-–]?\s*\d+", str(val), flags=re.I):
            markers.append((idx, str(val).strip()))

    if not markers:
        return []

    blocks = []
    for i, (start_idx, name) in enumerate(markers):
        end_idx = markers[i + 1][0] - 1 if i + 1 < len(markers) else len(df) - 1
        blocks.append((name, start_idx, end_idx))
    return blocks


def extract_first_number(val):
    if pd.isna(val):
        return np.nan
    if isinstance(val, (int, float, np.number)):
        return float(val)
    match = re.search(r"[-+]?\d*\.?\d+(?:[eE][-+]?\d+)?", str(val).replace(",", ""))
    return float(match.group()) if match else np.nan


def sum_row_values(row: pd.Series) -> float:
    nums = pd.to_numeric(row, errors="coerce")
    return float(np.nansum(nums.values))


def convert_mg_to_g(x_mg):
    return x_mg / 1000.0


def co2eq_kg_from_species_g(co2_g, ch4_g, n2o_g):
    return (co2_g + GWP_CH4 * ch4_g + GWP_N2O * n2o_g) / 1000.0


# -----------------------------
# Emission calculations
# -----------------------------
def energy_emissions_from_mj(mj: float, source_name: str):
    ef = ENERGY_FACTORS[source_name]

    co2_direct = mj * ef["co2_direct_g_mj"]
    ch4_direct = mj * ef["ch4_direct_g_mj"]
    n2o_direct = mj * convert_mg_to_g(mj * ef["n2o_direct_mg_mj"])

    co2_indirect = mj * ef["co2_indirect_g_mj"]
    ch4_indirect = mj * ef["ch4_indirect_g_mj"]
    n2o_indirect = mj * convert_mg_to_g(mj * ef["n2o_indirect_mg_mj"])

    return {
        "CO2_direct_g": co2_direct,
        "CH4_direct_g": ch4_direct,
        "N2O_direct_g": n2o_direct,
        "CO2_indirect_g": co2_indirect,
        "CH4_indirect_g": ch4_indirect,
        "N2O_indirect_g": n2o_indirect,
    }


def transport_emissions(throughput_t_day, distance_km, gasoline_share, diesel_share):
    tkm_day = throughput_t_day * distance_km
    total_mj = tkm_day * TRANSPORT_MJ_PER_TKM

    gas_mj = total_mj * gasoline_share
    diesel_mj = total_mj * diesel_share

    gas_em = energy_emissions_from_mj(gas_mj, "Gasoline")
    diesel_em = energy_emissions_from_mj(diesel_mj, "Diesel")

    out = {}
    for k in gas_em:
        out[k] = gas_em[k] + diesel_em[k]
    out["Transport_energy_MJ_day"] = total_mj
    out["Gasoline_energy_MJ_day"] = gas_mj
    out["Diesel_energy_MJ_day"] = diesel_mj
    return out


# -----------------------------
# Excel inspection helpers
# -----------------------------
def inspect_workbook_issues(xls: pd.ExcelFile):
    warnings = []

    for sheet in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet, header=None)

        flattened = " ".join(df.astype(str).fillna("").values.flatten().tolist())

        if "N20" in flattened:
            warnings.append(f"{sheet}: found 'N20', which likely should be 'N2O'.")

        if "Disel" in flattened:
            warnings.append(f"{sheet}: found 'Disel', which likely should be 'Diesel'.")

        if "68% gasoline" in flattened.lower() or "32% diesel" in flattened.lower():
            pass

        # Look for suspicious electricity conversions
        text = flattened.replace(" ", "")
        if "kW*24" in text or "kW×24" in text:
            warnings.append(
                f"{sheet}: found electricity conversion resembling 'kW × 24'. "
                f"If energy is needed in MJ/day, a consistent conversion is kW × 24 × 3.6."
            )
        if "kW*3.6" in text or "kW×3.6" in text:
            warnings.append(
                f"{sheet}: found electricity conversion resembling 'kW × 3.6'. "
                f"That converts kWh to MJ only after multiplying by hours."
            )

        # Try to detect contradictory gasoline/diesel shares
        if "32% gasoline" in flattened.lower() and "68% diesel" in flattened.lower():
            warnings.append(
                f"{sheet}: found transport fuel mix as 32% gasoline and 68% diesel. "
                f"Your paragraph states 32% gasoline and 68% diesel, so keep the text and formulas consistent."
            )

    return warnings


# -----------------------------
# Data extraction
# -----------------------------
def parse_sheet_for_analysis(df: pd.DataFrame):
    """
    This parser is intentionally flexible because real workbooks differ a lot.
    It tries to locate:
    - process blocks
    - direct emissions rows for CO2, CH4, N2O
    - electricity demand rows
    - throughput / transport rows if present
    """
    df = clean_columns(df)
    result_rows = []

    blocks = find_process_blocks(df)
    if not blocks:
        return pd.DataFrame()

    norm_cols = {c: normalize_name(c) for c in df.columns}

    # Guess columns
    electricity_col = None
    throughput_col = None
    distance_col = None

    for col, norm in norm_cols.items():
        if electricity_col is None and ("electric" in norm or "power" in norm):
            electricity_col = col
        if throughput_col is None and ("throughput" in norm or "feed" in norm or "massflow" in norm or "capacity" in norm):
            throughput_col = col
        if distance_col is None and "distance" in norm:
            distance_col = col

    first_col = df.columns[0]

    for process_name, start_idx, end_idx in blocks:
        block = df.loc[start_idx:end_idx].copy()

        # Find rows mentioning CO2 / CH4 / N2O
        co2_direct = 0.0
        ch4_direct = 0.0
        n2o_direct = 0.0
        electricity_kw = np.nan
        throughput_t_day = np.nan
        distance_km = np.nan

        for _, row in block.iterrows():
            row_text = " ".join([str(v) for v in row.values if pd.notna(v)])
            row_norm = normalize_name(row_text)

            if "co2" in row_norm and any(k in row_norm for k in ["direct", "emission", "stack", "vent"]):
                co2_direct += sum_row_values(row.drop(labels=[first_col], errors="ignore"))

            if "ch4" in row_norm and any(k in row_norm for k in ["direct", "emission", "stack", "vent"]):
                ch4_direct += sum_row_values(row.drop(labels=[first_col], errors="ignore"))

            if "n2o" in row_norm and any(k in row_norm for k in ["direct", "emission", "stack", "vent"]):
                n2o_direct += sum_row_values(row.drop(labels=[first_col], errors="ignore"))

            if pd.isna(electricity_kw) and ("electricity" in row_norm or "power" in row_norm):
                nums = pd.to_numeric(row, errors="coerce")
                if nums.notna().any():
                    electricity_kw = float(nums.dropna().iloc[0])

            if pd.isna(throughput_t_day) and any(k in row_norm for k in ["throughput", "feed", "capacity", "sludge"]):
                nums = pd.to_numeric(row, errors="coerce")
                if nums.notna().any():
                    throughput_t_day = float(nums.dropna().iloc[0])

            if pd.isna(distance_km) and "distance" in row_norm:
                nums = pd.to_numeric(row, errors="coerce")
                if nums.notna().any():
                    distance_km = float(nums.dropna().iloc[0])

        result_rows.append(
            {
                "Process": process_name,
                "CO2_WP_direct_g_day": co2_direct if np.isfinite(co2_direct) else 0.0,
                "CH4_WP_direct_g_day": ch4_direct if np.isfinite(ch4_direct) else 0.0,
                "N2O_WP_direct_g_day": n2o_direct if np.isfinite(n2o_direct) else 0.0,
                "Electricity_kW": electricity_kw,
                "Throughput_t_day": throughput_t_day,
                "Transport_distance_km": distance_km,
            }
        )

    return pd.DataFrame(result_rows)


def calculate_gwp_table(
    proc_df: pd.DataFrame,
    electricity_hours_per_day=24.0,
    electricity_source="Electricity",
    include_transport=True,
    transport_distance_default=50.0,
    throughput_default=1.0,
    gasoline_share=0.32,
    diesel_share=0.68,
):
    calc = proc_df.copy()

    # Electricity MJ/day = kW × h/day × 3.6
    calc["Electricity_MJ_day"] = (
        pd.to_numeric(calc["Electricity_kW"], errors="coerce").fillna(0.0)
        * electricity_hours_per_day
        * 3.6
    )

    elec_emissions = calc["Electricity_MJ_day"].apply(
        lambda mj: energy_emissions_from_mj(mj, electricity_source)
    )

    calc["CO2_EE_g_day"] = elec_emissions.apply(lambda x: x["CO2_direct_g"] + x["CO2_indirect_g"])
    calc["CH4_EE_g_day"] = elec_emissions.apply(lambda x: x["CH4_direct_g"] + x["CH4_indirect_g"])
    calc["N2O_EE_g_day"] = elec_emissions.apply(lambda x: x["N2O_direct_g"] + x["N2O_indirect_g"])

    # Transport emissions
    if include_transport:
        calc["Throughput_used_t_day"] = pd.to_numeric(calc["Throughput_t_day"], errors="coerce").fillna(throughput_default)
        calc["Distance_used_km"] = pd.to_numeric(calc["Transport_distance_km"], errors="coerce").fillna(transport_distance_default)

        transport_res = calc.apply(
            lambda r: transport_emissions(
                throughput_t_day=r["Throughput_used_t_day"],
                distance_km=r["Distance_used_km"],
                gasoline_share=gasoline_share,
                diesel_share=diesel_share,
            ),
            axis=1,
        )

        calc["CO2_transport_g_day"] = transport_res.apply(lambda x: x["CO2_direct_g"] + x["CO2_indirect_g"])
        calc["CH4_transport_g_day"] = transport_res.apply(lambda x: x["CH4_direct_g"] + x["CH4_indirect_g"])
        calc["N2O_transport_g_day"] = transport_res.apply(lambda x: x["N2O_direct_g"] + x["N2O_indirect_g"])
        calc["Transport_energy_MJ_day"] = transport_res.apply(lambda x: x["Transport_energy_MJ_day"])
    else:
        calc["CO2_transport_g_day"] = 0.0
        calc["CH4_transport_g_day"] = 0.0
        calc["N2O_transport_g_day"] = 0.0
        calc["Transport_energy_MJ_day"] = 0.0

    # Total species emissions
    calc["CO2_total_g_day"] = (
        calc["CO2_WP_direct_g_day"] + calc["CO2_EE_g_day"] + calc["CO2_transport_g_day"]
    )
    calc["CH4_total_g_day"] = (
        calc["CH4_WP_direct_g_day"] + calc["CH4_EE_g_day"] + calc["CH4_transport_g_day"]
    )
    calc["N2O_total_g_day"] = (
        calc["N2O_WP_direct_g_day"] + calc["N2O_EE_g_day"] + calc["N2O_transport_g_day"]
    )

    calc["GWP_kg_CO2eq_day"] = calc.apply(
        lambda r: co2eq_kg_from_species_g(
            r["CO2_total_g_day"], r["CH4_total_g_day"], r["N2O_total_g_day"]
        ),
        axis=1,
    )
    calc["GWP_t_CO2eq_day"] = calc["GWP_kg_CO2eq_day"] / 1000.0

    return calc


# -----------------------------
# Writing support
# -----------------------------
REVISED_TEXT = """
The determination of global warming potential (GWP) is widely used to evaluate the environmental impact of waste-to-energy (WtE) recovery processes, because global warming is closely associated with greenhouse gas (GHG) emissions released during these processes. Therefore, emissions of CO2, CH4, and N2O are considered in the GWP assessment based on the method proposed by Li and Cheng [51].

The GHG emissions are categorized into two groups: direct emissions from the TS upcycling processes and indirect emissions associated with energy consumption. The life-cycle system boundary of the TS upcycling processes is shown in Figure 4. Gaseous emissions may arise from TS transportation, treatment, and downstream separation steps used to obtain value-added products.

For transportation-related emissions, road transportation is assumed, with an energy consumption of 1362 kJ/t·km. The transport fuel is assumed to consist of 32% gasoline and 68% diesel. For process-energy requirements, electricity, steam, and natural gas are considered [51, 52].

To determine the global warming potential, the following correlation among greenhouse gases is used:

GHG^TE = CO2^WP + CO2^EE + 23(CH4^WP + CH4^EE) + 296(N2O^WP + N2O^EE)

where

CO2^EE = Σ_i S_i (D_i,CO2 + I_i,CO2)
CH4^EE = Σ_i S_i (D_i,CH4 + I_i,CH4)
N2O^EE = Σ_i S_i (D_i,N2O + I_i,N2O)

Here, D_i and I_i represent the direct and indirect emission factors, respectively, for energy source i, and S_i is the amount of energy consumed from source i. The direct and indirect emission factors for the different energy sources are listed in Table S4 of the Supplementary Information. The emissions data for CO2^WP, CH4^WP, and N2O^WP are obtained directly from the simulation models.
""".strip()


# -----------------------------
# Streamlit UI
# -----------------------------
st.title("Global Warming Potential (GWP) Analysis App")
st.write(
    "This app estimates GWP for WtE / TS upcycling processes from direct process emissions, "
    "energy-related emissions, and transportation emissions."
)

with st.expander("Method used", expanded=False):
    st.markdown(
        r"""
**Equation used**

\[
GHG^{TE} = CO_2^{WP} + CO_2^{EE} + 23\left(CH_4^{WP} + CH_4^{EE}\right) + 296\left(N_2O^{WP} + N_2O^{EE}\right)
\]

\[
CO_2^{EE} = \sum_i S_i \left(D_{i,CO_2} + I_{i,CO_2}\right)
\]

\[
CH_4^{EE} = \sum_i S_i \left(D_{i,CH_4} + I_{i,CH_4}\right)
\]

\[
N_2O^{EE} = \sum_i S_i \left(D_{i,N_2O} + I_{i,N_2O}\right)
\]

where:
- \(S_i\) = energy usage from source \(i\)
- \(D_i\) = direct emission factor
- \(I_i\) = indirect emission factor
        """
    )

uploaded_file = st.file_uploader("Upload Excel workbook", type=["xlsx", "xls"])

st.sidebar.header("Calculation settings")
electricity_hours_per_day = st.sidebar.number_input("Operating hours per day", min_value=1.0, value=24.0, step=1.0)
electricity_source = st.sidebar.selectbox("Electricity factor source", ["Electricity"])
include_transport = st.sidebar.checkbox("Include transportation emissions", value=True)
transport_distance_default = st.sidebar.number_input("Default transport distance (km)", min_value=0.0, value=50.0, step=1.0)
throughput_default = st.sidebar.number_input("Default throughput (t/day)", min_value=0.0, value=1.0, step=1.0)
gasoline_share = st.sidebar.number_input("Gasoline share", min_value=0.0, max_value=1.0, value=0.32, step=0.01)
diesel_share = st.sidebar.number_input("Diesel share", min_value=0.0, max_value=1.0, value=0.68, step=0.01)

share_sum = gasoline_share + diesel_share
if abs(share_sum - 1.0) > 1e-9:
    st.sidebar.warning(f"Fuel shares sum to {share_sum:.2f}, not 1.00.")

st.subheader("Revised writing")
st.text_area("Edited paragraph", REVISED_TEXT, height=340)

if uploaded_file is not None:
    try:
        xls = pd.ExcelFile(uploaded_file)
        st.success(f"Workbook loaded successfully. Sheets found: {', '.join(xls.sheet_names)}")

        warnings = inspect_workbook_issues(xls)
        if warnings:
            st.subheader("Workbook writing / formula issues detected")
            for w in warnings:
                st.warning(w)

        sheet_name = st.selectbox("Select sheet for analysis", xls.sheet_names)
        df = pd.read_excel(xls, sheet_name=sheet_name)
        st.subheader("Raw data preview")
        st.dataframe(df, use_container_width=True)

        parsed = parse_sheet_for_analysis(df)

        if parsed.empty:
            st.info(
                "No process blocks were detected automatically. "
                "Expected labels such as 'Process-1', 'Process-2', etc."
            )
        else:
            st.subheader("Parsed process data")
            st.dataframe(parsed, use_container_width=True)

            result = calculate_gwp_table(
                parsed,
                electricity_hours_per_day=electricity_hours_per_day,
                electricity_source=electricity_source,
                include_transport=include_transport,
                transport_distance_default=transport_distance_default,
                throughput_default=throughput_default,
                gasoline_share=gasoline_share,
                diesel_share=diesel_share,
            )

            st.subheader("GWP results")
            cols_to_show = [
                "Process",
                "CO2_WP_direct_g_day",
                "CH4_WP_direct_g_day",
                "N2O_WP_direct_g_day",
                "Electricity_MJ_day",
                "Transport_energy_MJ_day",
                "CO2_total_g_day",
                "CH4_total_g_day",
                "N2O_total_g_day",
                "GWP_kg_CO2eq_day",
                "GWP_t_CO2eq_day",
            ]
            st.dataframe(result[cols_to_show], use_container_width=True)

            st.subheader("Summary")
            total_gwp = result["GWP_kg_CO2eq_day"].sum()
            total_gwp_t = total_gwp / 1000.0

            c1, c2, c3 = st.columns(3)
            c1.metric("Processes analyzed", len(result))
            c2.metric("Total GWP (kg CO2-eq/day)", f"{total_gwp:,.3f}")
            c3.metric("Total GWP (t CO2-eq/day)", f"{total_gwp_t:,.6f}")

            csv = result.to_csv(index=False).encode("utf-8")
            st.download_button(
                "Download results as CSV",
                data=csv,
                file_name=f"gwp_results_{sheet_name}.csv",
                mime="text/csv",
            )

    except Exception as e:
        st.error(f"Failed to read or process the workbook: {e}")

else:
    st.info("Upload your Excel file to start the analysis.")
