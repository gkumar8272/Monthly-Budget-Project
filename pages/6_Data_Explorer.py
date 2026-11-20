"""
pages/6_Data_Explorer.py - Relational Data Explorer & Multi-Format Exports (UI-07)
"""

import os
import io
import streamlit as st
import pandas as pd
from src.ui import render_header
from src.db import get_db_connection

st.set_page_config(page_title="Data Explorer & Export", page_icon="💾", layout="wide")

def main():
    render_header(
        title="Relational Data Explorer & Audit Downloads",
        subtitle="Inspect underlying SQLite master dimensions, fact tables, and the 34-column analytical view",
        badge_text="UI-07 Data Governance"
    )

    tables = {
        "view_cac_full_analysis (Joined View - 34 cols)": "view_cac_full_analysis",
        "fact_marketing_campaigns (Fact Table 1)": "fact_marketing_campaigns",
        "fact_customer_cohorts (Fact Table 2)": "fact_customer_cohorts",
        "dim_channels (Master Dimension)": "dim_channels",
        "dim_subscription_tiers (Master Dimension)": "dim_subscription_tiers"
    }

    selected_label = st.selectbox("Select Relational Entity to Audit", options=list(tables.keys()))
    selected_table = tables[selected_label]

    conn = get_db_connection(read_only=True)
    df = pd.read_sql_query(f"SELECT * FROM {selected_table};", conn)
    conn.close()

    st.markdown(f"**Showing `{len(df)}` rows and `{df.shape[1]}` columns for entity `{selected_table}`:**")
    st.dataframe(df, use_container_width=True)

    # Export Buttons
    col1, col2 = st.columns(2)
    with col1:
        csv_data = df.to_csv(index=False).encode("utf-8")
        st.download_button(
            label=f"📥 Download {selected_table}.csv",
            data=csv_data,
            file_name=f"{selected_table}.csv",
            mime="text/csv",
            use_container_width=True
        )

    with col2:
        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name=selected_table[:31], index=False)
        st.download_button(
            label=f"📊 Download {selected_table}.xlsx",
            data=excel_buffer.getvalue(),
            file_name=f"{selected_table}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

if __name__ == "__main__":
    main()
