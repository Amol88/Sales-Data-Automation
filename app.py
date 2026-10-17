import streamlit as st
import pandas as pd
import io
import re

# Page Configuration
st.set_page_config(page_title="Partywise Sales & Outstanding", layout="wide")

st.title("📊 Excel Party-wise & Outstanding Splitter")
st.markdown("""
**Yeh app aapki Main Sales Sheet se:**
1. Har **Party** ke liye ek alag sheet banayega aur saara data same headings ke sath usmein transfer karega.
2. Jin rows mein **Payment received date** khali (blank) hai, unhe copy karke ek nayi **Outstanding sheet** banayega aur use **Invoice date** ke hisaab se sort karega.
""")

# File Uploader
uploaded_file = st.file_uploader("Apni Excel File Upload Karein (.xlsx)", type=["xlsx", "xls"])

def clean_sheet_name(name):
    """Excel sheet names mein invalid characters allow nahi hote, ye function unhe clean karega"""
    cleaned = re.sub(r'[\\/*?:\[\]]', '', str(name)).strip()
    return cleaned[:31] if cleaned else "Unknown_Party"

if uploaded_file is not None:
    try:
        # Read Excel and Sheet Names
        excel_file = pd.ExcelFile(uploaded_file)
        selected_sheet = st.selectbox("Main Sales Data wali Sheet select karein:", excel_file.sheet_names)
        
        # Load the dataframe
        df = pd.read_excel(uploaded_file, sheet_name=selected_sheet)
        columns = df.columns.tolist()
        
        st.write("### ⚙️ Apne Columns Select Karein")
        col1, col2, col3 = st.columns(3)
        
        party_col = col1.selectbox("1. Party Name wala Column:", columns, index=0)
        invoice_date_col = col2.selectbox("2. Invoice Date wala Column:", columns, index=min(1, len(columns)-1))
        payment_date_col = col3.selectbox("3. Payment Received Date wala Column:", columns, index=min(2, len(columns)-1))
        
        if st.button("🚀 Excel Generate Karein", type="primary"):
            with st.spinner("Data process ho raha hai, kripya wait karein..."):
                
                # Khali Party Name wale rows hata rahe hain (taki kachra na aaye)
                df_clean = df.dropna(subset=[party_col]).copy()
                
                # ---------------------------------------------------------
                # Step 2: Outstanding Data Extract karna (Blank Payment Date)
                # ---------------------------------------------------------
                # Check where payment date is missing (NaN, empty string, etc.)
                is_blank_payment = (
                    df_clean[payment_date_col].isna() | 
                    (df_clean[payment_date_col].astype(str).str.strip() == '') |
                    (df_clean[payment_date_col].astype(str).str.lower() == 'nan') |
                    (df_clean[payment_date_col].astype(str).str.lower() == 'nat')
                )
                
                outstanding_df = df_clean[is_blank_payment].copy()
                
                # Outstanding sheet ko Invoice Date ke hisaab se sort karna
                outstanding_df['_temp_invoice_date'] = pd.to_datetime(outstanding_df[invoice_date_col], errors='coerce')
                outstanding_df = outstanding_df.sort_values(by='_temp_invoice_date', ascending=True)
                outstanding_df = outstanding_df.drop(columns=['_temp_invoice_date']) # Temporary column hata diya taki headings same rahein
                
                # ---------------------------------------------------------
                # Excel File Banana (Memory mein)
                # ---------------------------------------------------------
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    
                    # Sabse pehle Outstanding Sheet save karte hain
                    outstanding_df.to_excel(writer, sheet_name='Outstanding sheet', index=False)
                    
                    # ---------------------------------------------------------
                    # Step 1: Party-wise Data Transfer karna
                    # ---------------------------------------------------------
                    unique_parties = df_clean[party_col].unique()
                    used_sheet_names = {'outstanding sheet'}
                    
                    for party in unique_parties:
                        # Us specific party ka data filter karna
                        party_df = df_clean[df_clean[party_col] == party].copy()
                        
                        # Sheet ka naam create aur clean karna
                        base_name = clean_sheet_name(party)
                        sheet_name = base_name
                        counter = 1
                        # Agar same naam ki sheet pehle se hai toh uske aage _1, _2 laga dega
                        while sheet_name.lower() in used_sheet_names:
                            suffix = f"_{counter}"
                            sheet_name = f"{base_name[:31-len(suffix)]}{suffix}"
                            counter += 1
                            
                        used_sheet_names.add(sheet_name.lower())
                        
                        # Data Excel mein likhna (Sari headings exact same rahengi)
                        party_df.to_excel(writer, sheet_name=sheet_name, index=False)
                
                output.seek(0)
                
                st.success(f"✅ Kaam Ho Gaya! **Outstanding sheet** aur **{len(unique_parties)} Party sheets** ban chuki hain.")
                
                # Download Button
                st.download_button(
                    label="📥 Nayi Excel File Download Karein",
                    data=output,
                    file_name="Final_Sales_and_Outstanding.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
    except Exception as e:
        st.error(f"Error aaya hai: {e}")
