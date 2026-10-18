import streamlit as st
import pandas as pd
import io
import re
from datetime import datetime

st.set_page_config(page_title="Advanced Sales & Outstanding Splitter", layout="wide")

st.title("📊 Advanced Excel Splitter (With Overdue & Missing Parties)")
st.markdown("""
**New Features:**
- **Outstanding Sheet** mein aakhir mein **Overdue (Days)** ka column add hoga.
- Agar kisi party ka Due Date mention nahi hai, toh ek alag **Missing Parties Excel** file generate hogi.
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
        
        # Agar Excel mein top par extra title line ho toh header row change karne ka option
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
        
        # Load Due Date Data (Agar select ki hai)
        df_due = None
        due_dict = {}
        
        if due_sheet_name != "None":
            df_due = pd.read_excel(uploaded_file, sheet_name=due_sheet_name, header=header_row)
            st.write("### ⚙️ Due Date Sheet ke Columns Match Karein")
            dc1, dc2 = st.columns(2)
            due_cols = df_due.columns.tolist()
            due_party_col = dc1.selectbox("Due Sheet - Party Name:", due_cols)
            due_days_col = dc2.selectbox("Due Sheet - Due Days (Number of days):", due_cols)
            
            # Dictionary banana fast matching ke liye 
            df_due[due_days_col] = pd.to_numeric(df_due[due_days_col], errors='coerce')
            due_dict = dict(zip(df_due[due_party_col].astype(str).str.strip(), df_due[due_days_col]))

        if st.button("🚀 Excel Generate Karein", type="primary"):
            with st.spinner("Data process ho raha hai, kripya wait karein..."):
                
                # Clean Data
                df_clean = df_main.dropna(subset=[party_col]).copy()
                df_clean[party_col] = df_clean[party_col].astype(str).str.strip()
                
                # Missing parties find karna
                missing_parties = []
                if due_sheet_name != "None":
                    all_main_parties = df_clean[party_col].unique()
                    missing_parties = [p for p in all_main_parties if p not in due_dict]
                
                # Dates convert karna taaki math calculation ho sake
                df_clean[invoice_date_col] = pd.to_datetime(df_clean[invoice_date_col], errors='coerce')
                
                # ---------------------------------------------------------
                # Outstanding Data
                # ---------------------------------------------------------
                is_blank_payment = (
                    df_clean[payment_date_col].isna() | 
                    (df_clean[payment_date_col].astype(str).str.strip() == '') |
                    (df_clean[payment_date_col].astype(str).str.lower() == 'nan')
                )
                
                out_df = df_clean[is_blank_payment].copy()
                out_df = out_df.sort_values(by=invoice_date_col, ascending=True)
                
                if sr_no_col != "None":
                    out_df[sr_no_col] = range(1, len(out_df) + 1)
                
                today = pd.Timestamp.today().normalize()
                
                # Overdue (Days) calculate karna
                def calculate_overdue(row):
                    party_name = row.get(party_col)
                    inv_date = row.get(invoice_date_col)
                    if pd.notna(party_name) and pd.notna(inv_date) and party_name in due_dict:
                        days = due_dict[party_name]
                        if pd.notna(days):
                            due_date = inv_date + pd.Timedelta(days=days)
                            overdue = (today - due_date).days
                            if overdue > 0:
                                return overdue
                    return None # Agar overdue nahi hai ya due date nahi hai toh blank rahega

                # Naya column last mein add karna
                out_df['Overdue (Days)'] = out_df.apply(calculate_overdue, axis=1)
                
                # Highlighting Function
                def highlight_row(row):
                    bg_color = ''
                    text_color = ''
                    overdue = row.get('Overdue (Days)')
                    
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

                out_styled = out_df.style.apply(highlight_row, axis=1)
                
                # ---------------------------------------------------------
                # Main File Export
                # ---------------------------------------------------------
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter', datetime_format='dd/mm/yyyy') as writer:
                    
                    df_main_original.to_excel(writer, sheet_name='Main Sheet', index=False)
                    out_styled.to_excel(writer, sheet_name='Outstanding Sheet', index=False)
                    
                    unique_parties = df_clean[party_col].unique()
                    used_sheet_names = {'main sheet', 'outstanding sheet'}
                    
                    for party in unique_parties:
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
                        
                        # Party sheet mein 'Overdue (Days)' column nahi chahiye toh use drop kar rahe hain agar galti se aa jaye
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
                # Missing Parties File Export (Agar missing parties hain)
                # ---------------------------------------------------------
                if missing_parties:
                    st.warning(f"⚠️ Dhyan Dein: {len(missing_parties)} parties aisi hain jinka 'Due Date' set nahi hai!")
                    
                    missing_df = pd.DataFrame(missing_parties, columns=["Party Name"])
                    missing_output = io.BytesIO()
                    
                    with pd.ExcelWriter(missing_output, engine='xlsxwriter') as missing_writer:
                        missing_df.to_excel(missing_writer, sheet_name='Missing Parties', index=False)
                        
                    missing_output.seek(0)
                    
                    st.download_button(
                        label="📥 Missing Parties ki Excel Download Karein",
                        data=missing_output,
                        file_name="missing_partywise_due_date.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="missing_parties_btn" # Unique key button error rokne ke liye
                    )

    except Exception as e:
        st.error(f"Error: {e}. Kripya check karein ki Excel mein data sahi format mein ho.")
