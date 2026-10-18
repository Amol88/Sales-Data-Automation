import streamlit as st
import pandas as pd
import io
import re
from datetime import datetime

st.set_page_config(page_title="Advanced Sales & Outstanding Splitter", layout="wide")

st.title("📊 Advanced Excel Splitter (With Overdue Sorting & Missing Parties)")
st.markdown("""
**New Features Added:**
- **Outstanding Sheet** mein data Party-wise alphabetically arrange hoga.
- Har party ke baad ek **Blank Row** aayegi.
- Jo Invoices **Overdue** ho chuke hain, woh party ke section mein pehle aayenge.
- Agar "Due Date" sheet mein kisi party ka date/days blank hai, toh **Missing Party Payment Due Date.xlsx** generate hogi.
""")

# File Uploader
uploaded_file = st.file_uploader("Apni Excel File Upload Karein (.xlsx)", type=["xlsx", "xls"])

def clean_sheet_name(name):
    cleaned = re.sub(r'[\\/*?:\[\]]', '', str(name)).strip()
    return cleaned[:31] if cleaned else "Unknown_Party"

if uploaded_file is not None:
    try:
        # Excel read karna aur sheet names nikalna
        xls = pd.ExcelFile(uploaded_file)
        sheet_names = xls.sheet_names
        
        st.write("### 📂 Sheets Select Karein")
        col_s1, col_s2 = st.columns(2)
        
        header_row = st.number_input("Headings (Columns) kis row number par hain? (Default 1 hota hai)", min_value=1, value=1) - 1
        
        main_sheet_name = col_s1.selectbox("1. Main Sales Data wali Sheet:", sheet_names)
        due_sheet_name = col_s2.selectbox("2. Partywise Due Date wali Sheet (Optional):", ["None"] + sheet_names)
        
        # Load Main Data
        df_main_original = pd.read_excel(uploaded_file, sheet_name=main_sheet_name, header=header_row)
        df_main = df_main_original.copy()
        columns = df_main.columns.tolist()
        
        st.write("### ⚙️ Main Sheet ke Columns Match Karein")
        c1, c2, c3, c4 = st.columns(4)
        party_col = c1.selectbox("Party Name Column:", columns)
        invoice_date_col = c2.selectbox("Invoice Date Column:", columns)
        payment_date_col = c3.selectbox("Payment Received Date Column:", columns)
        sr_no_col = c4.selectbox("Sr No. Column (Serial wise karne ke liye):", ["None"] + columns)
        
        # Load Due Date Data
        df_due = None
        due_dict = {}
        
        if due_sheet_name != "None":
            df_due = pd.read_excel(uploaded_file, sheet_name=due_sheet_name, header=header_row)
            st.write("### ⚙️ Due Date Sheet ke Columns Match Karein")
            dc1, dc2 = st.columns(2)
            due_cols = df_due.columns.tolist()
            due_party_col = dc1.selectbox("Due Sheet - Party Name:", due_cols)
            due_days_col = dc2.selectbox("Due Sheet - Payment Due Date (Days):", due_cols)
            
            # Convert due days to numeric, errors to NaN
            df_due[due_days_col] = pd.to_numeric(df_due[due_days_col], errors='coerce')
            due_dict = dict(zip(df_due[due_party_col].astype(str).str.strip(), df_due[due_days_col]))

        if st.button("🚀 Excel Generate Karein", type="primary"):
            with st.spinner("Data process ho raha hai, kripya wait karein..."):
                
                # ---------------------------------------------------------
                # 1. Missing Parties Check (Jinka Due Date cell blank hai)
                # ---------------------------------------------------------
                missing_parties = set()
                
                if due_sheet_name != "None":
                    # Check parties in Main Sheet that are missing or have blank due dates
                    all_main_parties = df_main[party_col].dropna().astype(str).str.strip().unique()
                    for p in all_main_parties:
                        if p not in due_dict or pd.isna(due_dict.get(p)):
                            missing_parties.add(p)
                    
                    # Check parties directly in the Due Date sheet whose due date is blank
                    for idx, row in df_due.iterrows():
                        d_party = str(row.get(due_party_col, "")).strip()
                        d_days = row.get(due_days_col)
                        if pd.isna(d_days) and d_party != "nan" and d_party != "":
                            missing_parties.add(d_party)
                            
                missing_parties = sorted(list(missing_parties))

                # Clean Main Data
                df_clean = df_main.dropna(subset=[party_col]).copy()
                df_clean[party_col] = df_clean[party_col].astype(str).str.strip()
                df_clean[invoice_date_col] = pd.to_datetime(df_clean[invoice_date_col], errors='coerce')
                
                # ---------------------------------------------------------
                # 2. Outstanding Data Formatting & Sorting
                # ---------------------------------------------------------
                is_blank_payment = (
                    df_clean[payment_date_col].isna() | 
                    (df_clean[payment_date_col].astype(str).str.strip() == '') |
                    (df_clean[payment_date_col].astype(str).str.lower() == 'nan')
                )
                
                out_df = df_clean[is_blank_payment].copy()
                today = pd.Timestamp.today().normalize()
                
                # Calculate Overdue
                def calculate_overdue(row):
                    p_name = row.get(party_col)
                    i_date = row.get(invoice_date_col)
                    if pd.notna(p_name) and pd.notna(i_date) and p_name in due_dict:
                        days = due_dict[p_name]
                        if pd.notna(days):
                            due_date = i_date + pd.Timedelta(days=days)
                            overdue = (today - due_date).days
                            if overdue > 0:
                                return overdue
                    return None 

                out_df['Overdue (Days)'] = out_df.apply(calculate_overdue, axis=1)
                
                # SORTING: Alphabetical Party -> Highest Overdue First -> Earliest Invoice First
                out_df = out_df.sort_values(by=[party_col, 'Overdue (Days)', invoice_date_col], 
                                            ascending=[True, False, True], 
                                            na_position='last')
                
                # Apply Sr No.
                if sr_no_col != "None":
                    out_df[sr_no_col] = range(1, len(out_df) + 1)
                
                # INSERT BLANK ROWS BETWEEN PARTIES
                spaced_rows = []
                prev_party = None
                
                for idx, row in out_df.iterrows():
                    curr_party = row[party_col]
                    if prev_party is not None and curr_party != prev_party:
                        # Insert completely blank row
                        empty_row = pd.Series([None] * len(out_df.columns), index=out_df.columns)
                        spaced_rows.append(empty_row)
                    spaced_rows.append(row)
                    prev_party = curr_party
                    
                out_df_spaced = pd.DataFrame(spaced_rows, columns=out_df.columns)
                
                # Highlighting Function (Blank rows ko skip karega)
                def highlight_row(row):
                    bg_color = ''
                    text_color = ''
                    overdue = row.get('Overdue (Days)')
                    party_name = row.get(party_col)
                    
                    if pd.isna(party_name): # Skip blank rows
                        return [''] * len(row)
                        
                    if pd.notna(overdue):
                        if overdue > 30:
                            bg_color = '#8B0000' # Dark Red
                            text_color = '#FFFFFF'
                        elif overdue > 0:
                            bg_color = '#FFC7CE' # Light Red
                            text_color = '#000000'
                                
                    if bg_color:
                        return [f'background-color: {bg_color}; color: {text_color}'] * len(row)
                    return [''] * len(row)

                out_styled = out_df_spaced.style.apply(highlight_row, axis=1)
                
                # ---------------------------------------------------------
                # 3. Main File Export
                # ---------------------------------------------------------
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter', datetime_format='dd/mm/yyyy') as writer:
                    
                    df_main_original.to_excel(writer, sheet_name='Main Sheet', index=False)
                    out_styled.to_excel(writer, sheet_name='Outstanding Sheet', index=False)
                    
                    unique_parties = df_clean[party_col].unique()
                    used_sheet_names = {'main sheet', 'outstanding sheet'}
                    
                    for party in sorted(unique_parties):
                        party_df = df_clean[df_clean[party_col] == party].copy()
                        
                        if sr_no_col != "None":
                            party_df[sr_no_col] = range(1, len(party_df) + 1)
                            
                        base_name = clean_sheet_name(party)
                        sheet_name = base_name
                        counter = 1
                        while sheet_name.lower() in used_sheet_names:
                            suffix = f"_{counter}"
                            sheet_name = f"{base_name[:31-len(suffix)]}{suffix}"
                            counter += 1
                        used_sheet_names.add(sheet_name.lower())
                        
                        if 'Overdue (Days)' in party_df.columns:
                            party_df = party_df.drop(columns=['Overdue (Days)'])
                            
                        party_df.to_excel(writer, sheet_name=sheet_name, index=False)
                        
                output.seek(0)
                st.success(f"✅ Main Report Processed! Outstanding aur {len(unique_parties)} Party Sheets ready hain.")
                
                st.download_button(
                    label="📥 Final Report Download Karein (Main & Partywise Data)",
                    data=output,
                    file_name="Final_Sales_Outstanding_Report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
                # ---------------------------------------------------------
                # 4. Missing Parties File Export
                # ---------------------------------------------------------
                if missing_parties:
                    st.error(f"⚠️ Dhyan Dein: {len(missing_parties)} parties aisi hain jinka 'Payment Due Date' blank (khali) hai!")
                    
                    missing_df = pd.DataFrame(missing_parties, columns=["Party Name"])
                    missing_output = io.BytesIO()
                    
                    with pd.ExcelWriter(missing_output, engine='xlsxwriter') as missing_writer:
                        missing_df.to_excel(missing_writer, sheet_name='Missing Due Dates', index=False)
                        
                    missing_output.seek(0)
                    
                    st.download_button(
                        label="📥 Missing Parties ki Excel Download Karein",
                        data=missing_output,
                        file_name="Missing Party Payment Due Date.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="missing_parties_btn" 
                    )

    except Exception as e:
        st.error(f"Error: {e}. Kripya check karein ki Excel mein data sahi format mein ho.")
