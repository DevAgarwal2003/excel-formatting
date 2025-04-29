import streamlit as st
import pandas as pd
from io import BytesIO
import re

def preprocess_excel(file):
    """Load and preprocess the uploaded Excel file."""
    data = pd.read_excel(file)
    df = pd.DataFrame(data)
    df = df.iloc[10:-5]
    if 'Unnamed: 2' in df.columns and df['Unnamed: 2'].astype(str).str.contains('Case Borrower Name').any():
        df = df.drop(['Unnamed: 0'], axis=1, errors='ignore')
    else:
        df = df.drop(['Unnamed: 0', 'Unnamed: 2'], axis=1, errors='ignore')
    df = df.reset_index(drop=True)
    df.columns = df.iloc[0]  # Set first row as column names
    df = df[1:].reset_index(drop=True)  # Remove first row and reset index
    return df

def format_dataframe(df):
    """Format the DataFrame by converting all date-like columns and expanding case numbers."""

    # Step 1: Clean up column headers (ensure unique names)
    df.columns = df.iloc[0].astype(str)  # Use first row as headers
    df = df[1:].reset_index(drop=True)

    # Ensure column names are unique
    def make_unique(col_list):
        counts = {}
        unique_cols = []
        for col in col_list:
            col = col.strip()
            if col in counts:
                counts[col] += 1
                unique_cols.append(f"{col}_{counts[col]}")
            else:
                counts[col] = 0
                unique_cols.append(col)
        return unique_cols

    df.columns = make_unique(df.columns)

    # Step 2: Convert all date-like columns
    for col in df.columns:
        try:
            parsed = pd.to_datetime(df[col], errors='coerce')
            if parsed.notna().sum() > 0:
                df[col] = parsed.dt.strftime('%d-%m-%Y')
        except Exception:
            continue

    # Step 3: Expand the case number column (assume it's the first column)
    case_col = df.columns[0]
    expanded_data = []

    for _, row in df.iterrows():
        cases = re.split(r'\s*/\s*', str(row[case_col]))
        for case in cases:
            new_row = row.copy()
            new_row[case_col] = case.strip()
            expanded_data.append(new_row)

    df_expanded = pd.DataFrame(expanded_data)

    return df_expanded



def to_excel(df):
    """Convert DataFrame to an Excel file in memory."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Processed Data')
    output.seek(0)
    return output

def main():
    st.markdown("""
    <style>
    body {
        background-color: black;
    }
    .title-style {
        font-size: 48px;
        color: white;
        background-color: #D2B48C;  
        padding: 20px;
        text-align: center;
        border-radius: 10px;
        width: 100%;
        margin: 0 auto;
    }
    </style>
    <div class="title-style">
        Excel Formatting App
    </div>
    """, unsafe_allow_html=True)

    uploaded_file = st.file_uploader("📤 Upload Excel File (.xlsx)", type=["xlsx"])
    
    if uploaded_file is not None:
        df = preprocess_excel(uploaded_file)
        df_expanded = format_dataframe(df)
        
        st.subheader("Processed Data")
        st.write(df_expanded)
        
        excel_data = to_excel(df_expanded)
        st.download_button(label="Download Processed Excel", data=excel_data, file_name="processed_data.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if __name__ == "__main__":
    main()
