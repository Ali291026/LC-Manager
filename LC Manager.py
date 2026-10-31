def load_and_parse_lc_data(file_input):
    if file_input is None:
        return None

    try:
        xls = pd.ExcelFile(file_input)
    except Exception as e:
        st.error(f"Error reading Excel file: {e}")
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

        df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
        
        # Find where data starts by looking for rows that contain LC references or invoice numbers
        start_row = 0
        for i, row in df.iterrows():
            row_str = " ".join([str(val).upper() for val in row.values if pd.notna(val)])
            if any(k in row_str for k in ["SUPPLIER", "LC #", "DESCRIPTION", "PLASTIC", "CABLE", "IMU"]):
                start_row = i + 1
                break
        
        # Read data starting from the detected row (or scan all rows if header is flexible)
        df_data = pd.read_excel(xls, sheet_name=sheet_name, skiprows=start_row)
        
        for idx, row in df_data.iterrows():
            # Get values safely by column position or name
            row_vals = [str(val).strip() for val in row.values if pd.notna(val)]
            if not row_vals:
                continue
                
            first_col_val = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
            
            # Skip totals, headers, or empty rows
            if not first_col_val or any(k in first_col_val.upper() for k in ["TOTAL", "OPEN CREDIT", "COMMISSION", "VAT", "PLASTIC/NAILS", "INVOICE VALUE"]):
                continue

            # Extract fields dynamically
            supplier = row.get("SUPPLIERS NAME", row.iloc[2] if len(row) > 2 else np.nan)
            description = row.get("DECRIPTION OF GOODS", row.get("DESCRIPTION", row.iloc[3] if len(row) > 3 else np.nan))
            tenure = row.get("DAYS", row.iloc[1] if len(row) > 1 else np.nan)
            maturity = row.get("MATURITY DATE", row.iloc[5] if len(row) > 5 else np.nan)
            inv_value = row.get("Invoice Value", row.iloc[6] if len(row) > 6 else np.nan)
            lc_status = row.get("LC Status", row.iloc[9] if len(row) > 9 else "NORMAL")
            shipment_status = row.get("Shipment Status", row.iloc[10] if len(row) > 10 else "AWAITING SHIPMENT")
            sob_date = row.get("SOB DATE", row.iloc[12] if len(row) > 12 else np.nan)

            if pd.isna(supplier) and pd.isna(description) and pd.isna(inv_value):
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
                "LC Type/Status": lc_status if pd.notna(lc_status) else "NORMAL",
                "Shipment Status": shipment_status if pd.notna(shipment_status) else "AWAITING SHIPMENT",
                "SOB Date": sob_date,
                "Document Status": "Pending Collection",
                "Maturity Date": maturity_date,
                "Expiration Date": expiry_date
            })

    if not records:
        return None

    lc_df = pd.DataFrame(records)
    lc_df["Tenure (Days)"] = pd.to_numeric(lc_df["Tenure (Days)"], errors="coerce")
    lc_df["Invoice Value"] = pd.to_numeric(lc_df["Invoice Value"], errors="coerce")
    return lc_df