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
    """Format the DataFrame using column indices for date and case number columns."""

    # Convert date columns using index positions (4, 6, 10, 11)
    date_indices = [1,2,3,4,5,6,7,8,9,10,11,12,13]
    for idx in date_indices:
        if idx < len(df.columns):
            col_name = df.columns[idx]
            df[col_name] = pd.to_datetime(df[col_name], errors='coerce').dt.strftime('%d-%m-%Y')

    # Get the case number column by index
    case_col_name = df.columns[0]
    other_cols = df.columns[1:]

    # Expand the case number column safely
    expanded_rows = []
    for _, row in df.iterrows():
        case_values = re.split(r'\s*/\s*', str(row[case_col_name]))
        for case in case_values:
            new_row = row.copy()
            new_row[case_col_name] = case.strip()
            expanded_rows.append(new_row)

    df_expanded = pd.DataFrame(expanded_rows)

    # Ensure correct column order
    final_columns = [case_col_name] + list(other_cols)
    df_expanded = df_expanded[final_columns]

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
