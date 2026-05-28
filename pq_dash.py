import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import os
from datetime import datetime

st.set_page_config(page_title="PQM Dashboard", layout="wide", page_icon="📊")

file_path = "Report.xlsx"

# ================= CSS =================

st.markdown("""
<style>
[data-testid="stAppViewContainer"] {
    background: linear-gradient(135deg,#eef3f8,#f8fbff);
}
[data-testid="stSidebar"] {
    background: linear-gradient(180deg,#071f4d,#0b3b82);
}
[data-testid="stSidebar"] * {
    color:white;
}
.main-title {
    background: linear-gradient(90deg,#071f4d,#0b5cad,#00a6d6);
    padding:24px;
    border-radius:22px;
    color:white;
    text-align:center;
    font-size:36px;
    font-weight:900;
    box-shadow:0 10px 28px rgba(0,0,0,0.18);
}
.kpi-card {
    background:white;
    padding:18px 8px;
    border-radius:22px;
    box-shadow:0 10px 28px rgba(0,0,0,0.12);
    text-align:center;
    height:128px;
    border-top:7px solid;
}
.kpi-title {
    font-size:13px;
    color:#666;
    font-weight:800;
}
.kpi-value {
    font-size:25px;
    color:#171717;
    font-weight:900;
    margin-top:16px;
    white-space:nowrap;
}
</style>
""", unsafe_allow_html=True)

# ================= LOAD DATA =================

@st.cache_data
def load_data():
    df = pd.read_excel(
        file_path,
        sheet_name="Branch Summary",
        header=None,
        skiprows=4,
        usecols="A:AD"
    )

    df.columns = [
        "Zone", "Region", "State", "DM_Name", "PQM_Name", "PQM_Emp_ID", "Branch",
        "Branch_GRT_Old", "PQM_GRT_Old", "Total_GRT_Old", "GRT_Old_Percent",
        "Branch_GRT", "PQM_GRT", "Total_GRT", "GRT_Percent",
        "Total_Disbursed", "Disbursed_Amount",
        "FTNR_PreApproved", "Received_HO", "FTNR_HO_Percent",
        "Branch_FTNR", "Branch_Total", "Branch_FTNR_Percent",
        "PQM_FTNR", "PQM_Total", "PQM_FTNR_Percent",
        "Same_Day", "T1", "T2", "Greater_2"
    ]

    df = df[df["Zone"].notna()]
    df = df[df["Zone"] != "Total"]

    text_cols = ["Zone", "Region", "State", "DM_Name", "PQM_Name", "PQM_Emp_ID", "Branch"]

    for col in df.columns:
        if col not in text_cols:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    percent_cols = [
        "GRT_Old_Percent", "GRT_Percent", "FTNR_HO_Percent",
        "Branch_FTNR_Percent", "PQM_FTNR_Percent",
        "Same_Day", "T1", "T2", "Greater_2"
    ]

    for col in percent_cols:
        df[col] = df[col].apply(lambda x: x * 100 if pd.notna(x) and x <= 1 else x)

    return df


df = load_data()

# ================= SIDEBAR =================

st.sidebar.title("📌 Filters")



zone_list = sorted(df["Zone"].dropna().unique())
zone_filter = st.sidebar.multiselect("Zone", zone_list)

if len(zone_filter) == 0:
    zone_filter = zone_list

state_list = sorted(df[df["Zone"].isin(zone_filter)]["State"].dropna().unique())
state_filter = st.sidebar.multiselect("State", state_list)

if len(state_filter) == 0:
    state_filter = state_list

branch_list = sorted(
    df[
        (df["Zone"].isin(zone_filter)) &
        (df["State"].isin(state_filter))
    ]["Branch"].dropna().unique()
)

branch_filter = st.sidebar.multiselect("Branch", branch_list)

if len(branch_filter) == 0:
    branch_filter = branch_list

filtered_df = df[
    (df["Zone"].isin(zone_filter)) &
    (df["State"].isin(state_filter)) &
    (df["Branch"].isin(branch_filter))
].copy()

# ================= DYNAMIC CHART LEVEL =================

if len(zone_filter) == 1 and len(state_filter) == 1:
    chart_level = "Branch"
elif len(zone_filter) == 1:
    chart_level = "State"
else:
    chart_level = "Zone"

# ================= HEADER =================

from datetime import datetime, timedelta

as_of_date = (datetime.today() - timedelta(days=1)).strftime("%d-%b-%Y")

st.markdown("""
<div class="main-title">
PQM PERFORMANCE DASHBOARD
</div>
""", unsafe_allow_html=True)

st.markdown(
    f"<div style='text-align:right;color:#666;font-weight:600;margin-top:8px;'>As of {as_of_date}</div>",
    unsafe_allow_html=True
)

# ================= KPI CALCULATION =================

total_branches = filtered_df["Branch"].nunique()
total_grt = filtered_df["Total_GRT_Old"].sum()
pqm_grt = filtered_df["PQM_GRT_Old"].sum()
pqm_grt_percent = pqm_grt / total_grt * 100 if total_grt else 0

branch_ftnr_percent = (
    filtered_df["Branch_FTNR"].sum()
    / filtered_df["Branch_Total"].sum()
    * 100
)

pqm_ftnr_percent = (
    filtered_df["PQM_FTNR"].sum()
    / filtered_df["PQM_Total"].sum()
    * 100
)

total_disbursed_amount = filtered_df["Disbursed_Amount"].sum()

def kpi_card(title, value, color):
    st.markdown(f"""
    <div class="kpi-card" style="border-top-color:{color};">
        <div class="kpi-title">{title}</div>
        <div class="kpi-value">{value}</div>
    </div>
    """, unsafe_allow_html=True)

# ================= NORMAL KPI =================

k1, k2, k3, k4, k5, k6, k7 = st.columns(7)

with k1:
    kpi_card("Branches", f"{total_branches:,.0f}", "#071f4d")
with k2:
    kpi_card("Total GRT", f"{total_grt:,.0f}", "#0070c0")
with k3:
    kpi_card("PQM GRT", f"{pqm_grt:,.0f}", "#7030a0")
with k4:
    kpi_card("PQM GRT %", f"{pqm_grt_percent:.0f}%", "#70ad47")
with k5:
    kpi_card("Branch FTNR %", f"{branch_ftnr_percent:.0f}%", "#c00000")
with k6:
    kpi_card("PQM FTNR %", f"{pqm_ftnr_percent:.2f}%", "#7030a0")
with k7:
    kpi_card("Pre-Approved DB Amt", f"₹ {total_disbursed_amount/10000000:.2f} Cr", "#ed7d31")

st.markdown("<br>", unsafe_allow_html=True)

# ================= CHARTS =================

c1, c2 = st.columns(2)

with c1:
    grt_data = filtered_df.groupby(chart_level, as_index=False).agg({
        "PQM_GRT_Old": "sum",
        "Total_GRT_Old": "sum"
    })

    grt_data["PQM_GRT_Percent"] = (
        grt_data["PQM_GRT_Old"] / grt_data["Total_GRT_Old"] * 100
    )

    fig1 = px.bar(
        grt_data,
        x=chart_level,
        y="PQM_GRT_Percent",
        text=grt_data["PQM_GRT_Percent"].round(0),
        title=f"PQM GRT % by {chart_level}",
        color=chart_level
    )

    fig1.update_traces(texttemplate="%{text}%", textposition="inside")
    fig1.update_layout(
        height=420,
        showlegend=False,
        plot_bgcolor="white",
        paper_bgcolor="white",
        title_font_size=22,
        xaxis_title="",
        yaxis_title="PQM GRT %"
    )

    st.plotly_chart(fig1, use_container_width=True)

with c2:
    pqm_data = filtered_df.groupby(chart_level, as_index=False).agg({
        "PQM_FTNR": "sum",
        "PQM_Total": "sum"
    })

    pqm_data["PQM_FTNR_Percent"] = (
        pqm_data["PQM_FTNR"] / pqm_data["PQM_Total"] * 100
    )

    pqm_data = pqm_data.sort_values("PQM_FTNR_Percent", ascending=True)

    fig2 = px.bar(
        pqm_data,
        x="PQM_FTNR_Percent",
        y=chart_level,
        orientation="h",
        text=pqm_data["PQM_FTNR_Percent"].round(2),
        title=f"PQM FTNR % by {chart_level}",
        color="PQM_FTNR_Percent"
    )

    fig2.update_traces(texttemplate="%{text}%")
    fig2.update_layout(
        height=420,
        plot_bgcolor="white",
        paper_bgcolor="white",
        title_font_size=22,
        xaxis_title="PQM FTNR %",
        yaxis_title=""
    )

    st.plotly_chart(fig2, use_container_width=True)

c3, c4 = st.columns(2)

with c3:
    disb_data = filtered_df.groupby(chart_level, as_index=False).agg({
        "Disbursed_Amount": "sum"
    })

    disb_data["Disbursed_Cr"] = disb_data["Disbursed_Amount"] / 10000000

    fig3 = px.pie(
        disb_data,
        names=chart_level,
        values="Disbursed_Amount",
        hole=0.55,
        title=f"Disbursed Amount by {chart_level}",
        custom_data=["Disbursed_Cr"]
    )

    fig3.update_traces(
        texttemplate="₹ %{customdata[0]:.2f} Cr",
        textposition="inside"
    )

    fig3.update_layout(
        height=420,
        paper_bgcolor="white",
        title_font_size=22
    )

    st.plotly_chart(fig3, use_container_width=True)

with c4:
    tat_data = filtered_df.groupby(chart_level, as_index=False).agg({
        "Same_Day": "mean",
        "T1": "mean",
        "T2": "mean",
        "Greater_2": "mean"
    })

    fig4 = go.Figure()

    for col in ["Same_Day", "T1", "T2", "Greater_2"]:
        fig4.add_trace(go.Bar(
            name=col,
            y=tat_data[chart_level],
            x=tat_data[col],
            orientation="h"
        ))

    fig4.update_layout(
        barmode="stack",
        title=f"CGT to GRT TAT by {chart_level}",
        height=420,
        xaxis_title="Percentage",
        yaxis_title="",
        plot_bgcolor="white",
        paper_bgcolor="white",
        title_font_size=22
    )

    st.plotly_chart(fig4, use_container_width=True)

# ================= TOP / BOTTOM =================

st.markdown("## 🔝 Top 10 & Bottom 10 Branches")

ranking_option = st.selectbox(
    "Ranking Base Select Kijiye",
    ["PQM GRT %", "PQM FTNR %", "Total Disbursement"]
)

rank_df = filtered_df.copy()

rank_df["PQM GRT No."] = rank_df["PQM_GRT_Old"]
rank_df["PQM GRT %"] = rank_df["PQM_GRT_Old"] / rank_df["Total_GRT_Old"] * 100
rank_df["PQM FTNR %"] = rank_df["PQM_FTNR"] / rank_df["PQM_Total"] * 100
rank_df["Total Disbursement"] = rank_df["Total_Disbursed"]

rank_df = rank_df[[
    "Zone", "State", "Branch",
    "PQM GRT No.", "PQM GRT %",
    "PQM FTNR %",
    "Total Disbursement"
]]

rank_df = rank_df.dropna(subset=[ranking_option])

if ranking_option == "PQM FTNR %":
    top10 = rank_df.sort_values(by=ranking_option, ascending=True).head(10)
    bottom10 = rank_df.sort_values(by=ranking_option, ascending=False).head(10)
else:
    top10 = rank_df.sort_values(by=ranking_option, ascending=False).head(10)
    bottom10 = rank_df.sort_values(by=ranking_option, ascending=True).head(10)

top10 = top10.round({
    "PQM GRT No.": 0,
    "PQM GRT %": 2,
    "PQM FTNR %": 2,
    "Total Disbursement": 0
})

bottom10 = bottom10.round({
    "PQM GRT No.": 0,
    "PQM GRT %": 2,
    "PQM FTNR %": 2,
    "Total Disbursement": 0
})

t1, t2 = st.columns(2)

with t1:
    st.markdown(f"### 🏆 Top 10 Branches by {ranking_option}")
    st.dataframe(top10, use_container_width=True, height=420)

with t2:
    st.markdown(f"### ⚠ Bottom 10 Branches by {ranking_option}")
    st.dataframe(bottom10, use_container_width=True, height=420)

# ================= SUMMARY =================

st.markdown("## 📋 Branch Wise Summary")

summary = filtered_df[[
    "Zone", "State", "Region", "DM_Name", "PQM_Name", "Branch",
    "Total_GRT_Old", "PQM_GRT_Old", "GRT_Old_Percent",
    "Total_GRT", "PQM_GRT", "GRT_Percent",
    "Total_Disbursed", "Disbursed_Amount",
    "Branch_FTNR_Percent",
    "PQM_FTNR_Percent",
    "Same_Day", "T1", "T2", "Greater_2"
]]

st.dataframe(summary, use_container_width=True, height=430)

with st.expander("View Raw Data"):
    st.dataframe(filtered_df, use_container_width=True)
