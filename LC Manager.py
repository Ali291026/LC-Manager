import streamlit as st
import pandas as pd
import numpy as np
import os

# --- Page Configuration ---
st.set_page_config(
    page_title="LC Master Dashboard & Multi-Currency Tracker",
    page_icon="🌍",
    layout="wide"
)

st.title("🌍 Multi-Currency Letter of Credit (LC) Executive Dashboard")
st.markdown("Manage your global Letters of Credit portfolio, track shipments, monitor document statuses, and view multi-currency exposures in real time.")

# --- File Uploader Widget ---
st.sidebar.header("📁 Data Management")
uploaded_file = st.sidebar.file_uploader("Upload or Update your LC Excel File", type=["xlsx", "xls"])

# --- Foolproof Data Parser ---
def load_and_parse_lc_data(file_input):
    if file_input is None:
        return None

    try:
        xls = pd.ExcelFile(file_input)
    except Exception as e:
        st.error(f"Error opening Excel file: {e}")
        return None

    records = []

    for sheet_name in xls.sheet_names:
        sheet_upper = sheet_name.upper()
        
        # Determine Bank and Currency
        if "HABIB" in sheet_upper:
            bank_name = "Habib Bank"
            default_currency = "TZS"
        elif "DTB" in sheet_upper:
            bank_name = "Diamond Trust Bank (DTB)"
            default_currency = "USD"
        elif "EXIM" in sheet_upper:
            bank_name = "Exim Bank"
            default_currency = "USD"
        else:
            continue

        try:
            df = pd.read_excel(xls, sheet_name=sheet_name)
        except Exception:
            continue

        # Find header row dynamically
        header_idx = 0
        for i, row in df.iterrows():
            row_str = " ".join([str(val).upper() for val in row.values if pd.notna(val)])
            if any(k in row_str for k in ["SUPPLIER", "LC #", "DESCRIPTION", "PLASTIC", "CABLE", "IMU"]):
                header_idx = i
                break
        
        try:
            df_data = pd.read_excel(xls, sheet_name=sheet_name, skiprows=header_idx+1)
        except Exception:
            df_data = df

        for idx, row in df_data.iterrows():
            try:
                # Ensure row has enough columns
                if len(row) < 2:
                    continue

                first_col_val = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
                
                # Skip invalid or header/total rows
                if not first_col_val or any(k in first_col_val.upper() for k in ["TOTAL", "OPEN CREDIT", "COMMISSION", "VAT", "PLASTIC/NAILS", "INVOICE VALUE", "UNNAMED"]):
                    continue

                # Safe column getters with positional fallbacks
                supplier = row.iloc[2] if len(row) > 2 and pd.notna(row.iloc[2]) else "Unknown Supplier"
                description = row.iloc[3] if len(row) > 3 and pd.notna(row.iloc[3]) else "N/A"
                tenure = row.iloc[1] if len(row) > 1 and pd.notna(row.iloc[1]) else np.nan
                maturity = row.iloc[5] if len(row) > 5 and pd.notna(row.iloc[5]) else np.nan
                inv_value = row.iloc[6] if len(row) > 6 and pd.notna(row.iloc[6]) else np.nan
                lc_status = row.iloc[9] if len(row) > 9 and pd.notna(row.iloc[9]) else "NORMAL"
                shipment_status = row.iloc[10] if len(row) > 10 and pd.notna(row.iloc[10]) else "AWAITING SHIPMENT"
                sob_date = row.iloc[12] if len(row) > 12 and pd.notna(row.iloc[12]) else np.nan

                if pd.isna(inv_value):
                    continue

                expiry_date, maturity_date = None, None
                if pd.notna(maturity):
                    mat_str = str(maturity).strip()
                    if "EXPIRY" in mat_str.upper():
                        expiry_date = mat_str
                    else:
                        maturity_date = mat_str

                records.append({
                    "Bank Name": bank_name,
                    "LC No. / Ref": first_col_val,
                    "Supplier Name": supplier,
                    "Goods Description": description,
                    "Currency": default_currency,
                    "Invoice Value": inv_value,
                    "Tenure (Days)": tenure,
                    "LC Type/Status": lc_status,
                    "Shipment Status": shipment_status,
                    "SOB Date": sob_date,
                    "Document Status": "Pending Collection",
                    "Maturity Date": maturity_date,
                    "Expiration Date": expiry_date
                })
            except Exception:
                continue

    if not records:
        return None

    lc_df = pd.DataFrame(records)
    lc_df["Tenure (Days)"] = pd.to_numeric(lc_df["Tenure (Days)"], errors="coerce")
    lc_df["Invoice Value"] = pd.to_numeric(lc_df["Invoice Value"], errors="coerce")
    return lc_df

# --- Load Data ---
default_file = "LC Summary for Program.xlsx"
df = None

if uploaded_file is not None:
    df = load_and_parse_lc_data(uploaded_file)
elif os.path.exists(default_file):
    df = load_and_parse_lc_data(default_file)

# --- Render App UI ---
if df is None or df.empty:
    st.warning("⚠️ Could not load LC data automatically.")
    st.info("👉 **Please upload your 'LC Summary for Program.xlsx' file using the sidebar uploader.**")
else:
    st.success(f"Successfully loaded {len(df)} LC records across all banks!")
    
    # --- Sidebar Filters ---
    st.sidebar.markdown("---")
    st.sidebar.header("🔍 Advanced Filters")
    
    selected_bank = st.sidebar.selectbox("Bank Name", ["All"] + list(df["Bank Name"].unique()))
    selected_currency = st.sidebar.selectbox("Currency", ["All"] + list(df["Currency"].unique()))
    selected_status = st.sidebar.selectbox("Shipment Status", ["All"] + list(df["Shipment Status"].unique()))
    selected_type = st.sidebar.selectbox("LC Type", ["All"] + list(df["LC Type/Status"].unique()))

    # Apply Filters
    filtered_df = df.copy()
    if selected_bank != "All":
        filtered_df = filtered_df[filtered_df["Bank Name"] == selected_bank]
    if selected_currency != "All":
        filtered_df = filtered_df[filtered_df["Currency"] == selected_currency]
    if selected_status != "All":
        filtered_df = filtered_df[filtered_df["Shipment Status"] == selected_status]
    if selected_type != "All":
        filtered_df = filtered_df[filtered_df["LC Type/Status"] == selected_type]

    # --- Top KPI Summary Metrics ---
    st.markdown("### 📈 Executive Key Metrics")
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric("Total Active LCs", len(filtered_df))
    with col2:
        usd_total = filtered_df[filtered_df["Currency"] == "USD"]["Invoice Value"].sum()
        st.metric("Total Exposure (USD)", f"${usd_total:,.2f}")
    with col3:
        tzs_total = filtered_df[filtered_df["Currency"] == "TZS"]["Invoice Value"].sum()
        st.metric("Total Exposure (TZS)", f"TZS {tzs_total:,.2f}" if tzs_total > 0 else "TZS 0.00")
    with col4:
        avg_tenure = filtered_df["Tenure (Days)"].mean()
        st.metric("Avg. Credit Tenure", f"{avg_tenure:.1f} Days" if not np.isnan(avg_tenure) else "N/A")

    st.markdown("---")

    # --- Charts Section ---
    c1, c2 = st.columns(2)
    with c1:
        st.subheader("📊 Exposure by Bank & Currency")
        bank_currency_exp = filtered_df.groupby(["Bank Name", "Currency"])["Invoice Value"].sum().unstack(fill_value=0)
        st.bar_chart(bank_currency_exp)

    with c2:
        st.subheader("📦 Shipment Milestones")
        shipment_counts = filtered_df["Shipment Status"].value_counts()
        st.bar_chart(shipment_counts)

    st.markdown("---")

    # --- Interactive Detailed Table ---
    st.subheader("📋 Filtered LC Detailed Report")
    st.dataframe(filtered_df, use_container_width=True)

    # --- Export Report Button ---
    st.download_button(
        label="📥 Download Filtered Report as CSV",
        data=filtered_df.to_csv(index=False).encode('utf-8'),
        file_name='filtered_lc_multicurrency_report.csv',
        mime='text/csv',
    )