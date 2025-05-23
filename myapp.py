import streamlit as st
import pandas as pd
from io import BytesIO
import re

def preprocess_excel(file):
    """Load and preprocess the uploaded Excel file."""
    data = pd.read_excel(file)
    df = pd.DataFrame(data)
    
    # Trim top 10 and bottom 5 rows (adjust as needed)
    df = df.iloc[10:-5]

    # Drop unwanted columns based on content
    if 'Unnamed: 2' in df.columns and df['Unnamed: 2'].astype(str).str.contains('Case Borrower Name').any():
        df = df.drop(['Unnamed: 0'], axis=1, errors='ignore')
    else:
        df = df.drop(['Unnamed: 0', 'Unnamed: 2'], axis=1, errors='ignore')

    df = df.reset_index(drop=True)

    # Safely set first row as header
    new_header = df.iloc[0]              # First row as header
    df = df[1:].copy()                   # Remove header row from data
    df.columns = new_header              # Assign new headers
    df = df.reset_index(drop=True)       # Reset index

    return df

def format_dataframe(df):
    """Dynamically detect and format date-like columns, and expand case numbers."""
    for col in df.columns:
        non_null_series = df[col].dropna().astype(str)
        parsed_dates = pd.to_datetime(non_null_series, errors='coerce')

        if parsed_dates.notna().sum() / len(non_null_series) > 0.5:
            df[col] = pd.to_datetime(df[col], errors='coerce').dt.strftime('%d/%m/%y')

    # Expand the case number column (assumed to be the first one)
    case_col_name = df.columns[0]
    other_cols = df.columns[1:]

    expanded_rows = []
    for _, row in df.iterrows():
        case_values = re.split(r'\s*/\s*', str(row[case_col_name]))
        for case in case_values:
            new_row = row.copy()
            new_row[case_col_name] = case.strip()
            expanded_rows.append(new_row)

    df_expanded = pd.DataFrame(expanded_rows)
    df_expanded = df_expanded[[case_col_name] + list(other_cols)]  # Keep original column order

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
        st.download_button(
            label="Download Processed Excel",
            data=excel_data,
            file_name="processed_data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

if __name__ == "__main__":
    main()
