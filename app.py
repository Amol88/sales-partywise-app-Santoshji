import streamlit as st
import pandas as pd
import io
import re

# Page Configuration
st.set_page_config(page_title="Partywise Sales & Outstanding Automation", layout="wide")

st.title("📊 Party-wise Sales & Outstanding Excel Generator")
st.markdown("""
Yeh app aapke **Main Sales Data** ko upload karte hi:
1. Har **Unique Party** ke liye alag sheet banayega (Date-wise sorted).
2. Jin entries mein **Payment Received Date** blank hai, unka ek alag **Outstanding Sheet** (Date-wise) ready karega.
""")

# Step 1: Upload Excel File
uploaded_file = st.file_uploader("Apni Main Sales Excel File Upload Karein (.xlsx)", type=["xlsx", "xls"])

def clean_sheet_name(name):
    """Excel sheet name rules: max 31 chars, no invalid characters like \ / ? * : [ ]"""
    cleaned = re.sub(r'[\\/*?:\[\]]', '', str(name)).strip()
    return cleaned[:31] if cleaned else "Unknown_Party"

if uploaded_file is not None:
    try:
        # Read all sheet names to allow user to select the main sheet
        excel_file = pd.ExcelFile(uploaded_file)
        sheet_names = excel_file.sheet_names
        
        selected_sheet = st.selectbox("Main Sales Data wali Sheet select karein:", sheet_names)
        
        # Load data from selected sheet
        df = pd.read_excel(uploaded_file, sheet_name=selected_sheet)
        
        st.subheader("📋 Uploaded Data Preview")
        st.dataframe(df.head(5), use_container_width=True)
        
        columns = df.columns.tolist()
        
        st.subheader("⚙️ Columns Map Karein")
        col1, col2, col3, col4 = st.columns(4)
        
        # Smart default selection helper
        def get_default_index(cols, keywords):
            for idx, col in enumerate(cols):
                if any(k in str(col).lower() for k in keywords):
                    return idx
            return 0

        with col1:
            party_col = st.selectbox(
                "1. Party Name Column:", 
                columns, 
                index=get_default_index(columns, ['party', 'customer', 'client', 'name'])
            )
        with col2:
            date_col = st.selectbox(
                "2. Sale/Bill Date Column:", 
                columns, 
                index=get_default_index(columns, ['date', 'bill date', 'sale date', 'invoice date'])
            )
        with col3:
            payment_date_col = st.selectbox(
                "3. Payment Received Date Column:", 
                columns, 
                index=get_default_index(columns, ['payment', 'received', 'rec date', 'clear'])
            )
        with col4:
            amount_col = st.selectbox(
                "4. Sale/Bill Amount Column (Optional):", 
                ["None"] + columns, 
                index=get_default_index(["None"] + columns, ['amount', 'total', 'sale', 'value', 'amt'])
            )

        if st.button("🚀 Process Data & Generate Excel", type="primary"):
            with st.spinner("Sheets ready ho rahi hain, kripya wait karein..."):
                # Clean Party Column (Remove completely empty party rows)
                df_clean = df.dropna(subset=[party_col]).copy()
                
                # Convert Sale Date column to datetime for accurate sorting
                df_clean['_sort_date'] = pd.to_datetime(df_clean[date_col], errors='coerce')
                
                # Sort main data by Date
                df_clean = df_clean.sort_values(by='_sort_date', ascending=True)
                
                # Format date nicely if it got converted, or keep original
                # Drop helper column before exporting
                
                # ---------------------------------------------------------
                # 2. Extract Outstanding Data (Where Payment Received Date is Blank)
                # ---------------------------------------------------------
                is_blank_payment = (
                    df_clean[payment_date_col].isna() | 
                    (df_clean[payment_date_col].astype(str).str.strip() == '') | 
                    (df_clean[payment_date_col].astype(str).str.lower() == 'nan') |
                    (df_clean[payment_date_col].astype(str).str.lower() == 'none')
                )
                
                outstanding_df = df_clean[is_blank_payment].copy()
                outstanding_df = outstanding_df.sort_values(by='_sort_date', ascending=True)
                
                # Remove helper column
                df_export = df_clean.drop(columns=['_sort_date'])
                outstanding_export = outstanding_df.drop(columns=['_sort_date'])
                
                # Create Party-wise Outstanding Summary Table if Amount column is selected
                summary_df = None
                if amount_col != "None":
                    outstanding_export[amount_col] = pd.to_numeric(outstanding_export[amount_col], errors='coerce').fillna(0)
                    summary_df = outstanding_export.groupby(party_col, as_index=False)[amount_col].sum()
                    summary_df.rename(columns={amount_col: 'Total Outstanding Amount'}, inplace=True)
                    summary_df = summary_df.sort_values(by='Total Outstanding Amount', ascending=False)

                # ---------------------------------------------------------
                # Create Excel File in Memory
                # ---------------------------------------------------------
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    # Sheet 1: Main Sheet
                    df_export.to_excel(writer, sheet_name='Main_Sales_Data', index=False)
                    
                    # Sheet 2: Outstanding Details (Date-wise all parties where payment date is blank)
                    outstanding_export.to_excel(writer, sheet_name='Outstanding_Datewise', index=False)
                    
                    # Sheet 3: Party-wise Outstanding Summary (Optional bonus sheet)
                    if summary_df is not None:
                        summary_df.to_excel(writer, sheet_name='Partywise_Summary', index=False)
                    
                    # Sheet 4+: Unique Party Sheets
                    unique_parties = df_clean[party_col].unique()
                    used_sheet_names = {'main_sales_data', 'outstanding_datewise', 'partywise_summary'}
                    
                    for party in unique_parties:
                        party_data = df_clean[df_clean[party_col] == party].sort_values(by='_sort_date', ascending=True)
                        party_data_export = party_data.drop(columns=['_sort_date'])
                        
                        # Ensure unique & valid Excel sheet name
                        base_name = clean_sheet_name(party)
                        sheet_name = base_name
                        counter = 1
                        while sheet_name.lower() in used_sheet_names:
                            suffix = f"_{counter}"
                            sheet_name = f"{base_name[:31-len(suffix)]}{suffix}"
                            counter += 1
                        
                        used_sheet_names.add(sheet_name.lower())
                        party_data_export.to_excel(writer, sheet_name=sheet_name, index=False)

                output.seek(0)
                
                st.success(f"✅ Processing Complete! Total **{len(unique_parties)}** Party sheets aur **Outstanding** sheet ready ho gayi hai.")
                
                # Preview of Outstanding Data on Screen
                st.subheader("📌 Outstanding Data Preview (Payment Date Blank)")
                col_a, col_b = st.columns([2, 1])
                with col_a:
                    st.write("**Date-wise Outstanding Entries:**")
                    st.dataframe(outstanding_export, use_container_width=True)
                with col_b:
                    if summary_df is not None:
                        st.write("**Party-wise Total Outstanding:**")
                        st.dataframe(summary_df, use_container_width=True)
                
                # Download Button
                st.download_button(
                    label="📥 Ready Excel File Download Karein",
                    data=output,
                    file_name="Partywise_Sales_And_Outstanding_Report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    except Exception as e:
        st.error(f"File read karne mein error aaya: {e}")
