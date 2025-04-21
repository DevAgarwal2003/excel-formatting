import streamlit as st
import pandas as pd
from io import BytesIO

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
    """Format the DataFrame by splitting the 'Case No: Loan A/C No.' column and reformatting date columns."""
    
    # Convert date columns to 'dd-mm-yyyy' format
    date_columns = ['DM Filling date','Date of Filling', 'DM Order Date', 'Verification date','Next date of Hearing']
    for col in date_columns:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors='coerce').dt.strftime('%d-%m-%Y')

    # Split 'Case No: Loan A/C No.' column
    cols_to_keep = [col for col in df.columns if col != 'Case No: Loan A/C No.']
    
    df_expanded = df.set_index(cols_to_keep)['Case No: Loan A/C No.'].str.split(' / |/', expand=True).stack().reset_index(name='Case No: Loan A/C No.')
    df_expanded = df_expanded.loc[:, ~df_expanded.columns.str.contains('level')]
    
    # Restore original columns
    df_expanded = df_expanded.merge(df[cols_to_keep], on=cols_to_keep, how='left')
    
    # Reorder columns
    final_columns = ['Case No: Loan A/C No.'] + cols_to_keep
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
