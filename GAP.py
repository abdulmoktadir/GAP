import pandas as pd
import streamlit as st

st.set_page_config(page_title="Manual GWP Analysis", layout="wide")

GWP_CH4 = 23
GWP_N2O = 296

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

TRANSPORT_MJ_PER_TKM = 1.362  # 1362 kJ/t·km = 1.362 MJ/t·km


def mg_to_g(x_mg):
    return x_mg / 1000.0


def energy_emissions_from_mj(mj, source_name):
    ef = ENERGY_FACTORS[source_name]
    return {
        "co2_g": mj * (ef["co2_direct_g_mj"] + ef["co2_indirect_g_mj"]),
        "ch4_g": mj * (ef["ch4_direct_g_mj"] + ef["ch4_indirect_g_mj"]),
        "n2o_g": mj * mg_to_g(mj * 0) if False else mj * mg_to_g(ef["n2o_direct_mg_mj"] + ef["n2o_indirect_mg_mj"]),
    }


def transport_emissions(throughput_t_day, distance_km, gasoline_share, diesel_share):
    total_mj = throughput_t_day * distance_km * TRANSPORT_MJ_PER_TKM
    gas_mj = total_mj * gasoline_share
    diesel_mj = total_mj * diesel_share

    gas = energy_emissions_from_mj(gas_mj, "Gasoline")
    diesel = energy_emissions_from_mj(diesel_mj, "Diesel")

    return {
        "transport_energy_mj_day": total_mj,
        "co2_g": gas["co2_g"] + diesel["co2_g"],
        "ch4_g": gas["ch4_g"] + diesel["ch4_g"],
        "n2o_g": gas["n2o_g"] + diesel["n2o_g"],
    }


def calc_gwp_kg(co2_g, ch4_g, n2o_g):
    return (co2_g + 23 * ch4_g + 296 * n2o_g) / 1000.0


st.title("Manual Global Warming Potential (GWP) Calculator")

with st.expander("Equation", expanded=True):
    st.markdown(
        """
\[
GHG^{TE} = CO_2^{WP} + CO_2^{EE} + 23(CH_4^{WP} + CH_4^{EE}) + 296(N_2O^{WP} + N_2O^{EE})
\]

\[
CO_2^{EE} = \sum_i S_i(D_{i,CO_2} + I_{i,CO_2})
\]

\[
CH_4^{EE} = \sum_i S_i(D_{i,CH_4} + I_{i,CH_4})
\]

\[
N_2O^{EE} = \sum_i S_i(D_{i,N_2O} + I_{i,N_2O})
\]
"""
    )

st.sidebar.header("General settings")
num_processes = st.sidebar.number_input("Number of processes", min_value=1, value=1, step=1)
include_transport = st.sidebar.checkbox("Include transportation emissions", value=True)
gasoline_share = st.sidebar.number_input("Gasoline share", min_value=0.0, max_value=1.0, value=0.32, step=0.01)
diesel_share = st.sidebar.number_input("Diesel share", min_value=0.0, max_value=1.0, value=0.68, step=0.01)

if abs((gasoline_share + diesel_share) - 1.0) > 1e-9:
    st.sidebar.warning("Gasoline share + diesel share should equal 1.00")

results = []

for i in range(int(num_processes)):
    st.subheader(f"Process {i+1}")

    col1, col2 = st.columns(2)

    with col1:
        process_name = st.text_input(f"Process name {i+1}", value=f"Process-{i+1}")

        st.markdown("**Direct emissions from simulation model**")
        co2_wp = st.number_input(f"CO2^WP (g/day) - {process_name}", min_value=0.0, value=0.0, key=f"co2wp_{i}")
        ch4_wp = st.number_input(f"CH4^WP (g/day) - {process_name}", min_value=0.0, value=0.0, key=f"ch4wp_{i}")
        n2o_wp = st.number_input(f"N2O^WP (g/day) - {process_name}", min_value=0.0, value=0.0, key=f"n2owp_{i}")

    with col2:
        st.markdown("**Energy consumption**")
        electricity_mj = st.number_input(f"Electricity (MJ/day) - {process_name}", min_value=0.0, value=0.0, key=f"elec_{i}")
        steam_mj = st.number_input(f"Steam (MJ/day) - {process_name}", min_value=0.0, value=0.0, key=f"steam_{i}")
        natural_gas_mj = st.number_input(f"Natural gas (MJ/day) - {process_name}", min_value=0.0, value=0.0, key=f"ng_{i}")

        throughput_t_day = 0.0
        distance_km = 0.0
        if include_transport:
            st.markdown("**Transportation input**")
            throughput_t_day = st.number_input(f"Throughput (t/day) - {process_name}", min_value=0.0, value=0.0, key=f"throughput_{i}")
            distance_km = st.number_input(f"Transport distance (km) - {process_name}", min_value=0.0, value=0.0, key=f"distance_{i}")

    elec_em = energy_emissions_from_mj(electricity_mj, "Electricity")
    steam_em = energy_emissions_from_mj(steam_mj, "Steam")
    ng_em = energy_emissions_from_mj(natural_gas_mj, "Natural gas")

    co2_ee = elec_em["co2_g"] + steam_em["co2_g"] + ng_em["co2_g"]
    ch4_ee = elec_em["ch4_g"] + steam_em["ch4_g"] + ng_em["ch4_g"]
    n2o_ee = elec_em["n2o_g"] + steam_em["n2o_g"] + ng_em["n2o_g"]

    transport_energy = 0.0
    co2_tr = 0.0
    ch4_tr = 0.0
    n2o_tr = 0.0

    if include_transport:
        tr = transport_emissions(throughput_t_day, distance_km, gasoline_share, diesel_share)
        transport_energy = tr["transport_energy_mj_day"]
        co2_tr = tr["co2_g"]
        ch4_tr = tr["ch4_g"]
        n2o_tr = tr["n2o_g"]

    co2_total = co2_wp + co2_ee + co2_tr
    ch4_total = ch4_wp + ch4_ee + ch4_tr
    n2o_total = n2o_wp + n2o_ee + n2o_tr

    gwp_kg_day = calc_gwp_kg(co2_total, ch4_total, n2o_total)
    gwp_t_day = gwp_kg_day / 1000.0

    results.append(
        {
            "Process": process_name,
            "CO2_WP_g_day": co2_wp,
            "CH4_WP_g_day": ch4_wp,
            "N2O_WP_g_day": n2o_wp,
            "Electricity_MJ_day": electricity_mj,
            "Steam_MJ_day": steam_mj,
            "NaturalGas_MJ_day": natural_gas_mj,
            "Transport_MJ_day": transport_energy,
            "CO2_total_g_day": co2_total,
            "CH4_total_g_day": ch4_total,
            "N2O_total_g_day": n2o_total,
            "GWP_kg_CO2eq_day": gwp_kg_day,
            "GWP_t_CO2eq_day": gwp_t_day,
        }
    )

df = pd.DataFrame(results)

st.subheader("Results")
st.dataframe(df, use_container_width=True)

if not df.empty:
    total_gwp_kg = df["GWP_kg_CO2eq_day"].sum()
    total_gwp_t = total_gwp_kg / 1000.0

    c1, c2, c3 = st.columns(3)
    c1.metric("Processes analyzed", len(df))
    c2.metric("Total GWP (kg CO2-eq/day)", f"{total_gwp_kg:,.3f}")
    c3.metric("Total GWP (t CO2-eq/day)", f"{total_gwp_t:,.6f}")

csv = df.to_csv(index=False).encode("utf-8")
st.download_button(
    "Download results as CSV",
    data=csv,
    file_name="manual_gwp_results.csv",
    mime="text/csv",
)
