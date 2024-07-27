import pandas as pd
import streamlit as st

# Streamlit app
st.title('Excel Sheets Merger')

# Upload Excel files
uploaded_files = st.file_uploader("Upload Excel files", type="xlsx", accept_multiple_files=True)

# Define the sheet names you want to merge
sheet_names = ['H1 ICD-10-CM (Non-IQVIA)', 'Cluster1', 'Cluster2', 'Cluster3', 'Cluster4', 'Cluster5', 'Cluster6', 'Cluster7', 'Cluster8', 'Cluster9', 'Cluster10']

# Define the columns you want to include in the merged DataFrame
columns = ['CODE', 'DESCRIPTION', 'Decision', 'Mendel ID', 'Concept Name', 'Missing Concept', 'Parent Mendel ID If Missing Concept', 'Parent Concept Name If Missing Concept']

if uploaded_files:
    # Initialize an empty DataFrame to hold the merged data
    merged_df = pd.DataFrame(columns=columns + ['source_sheet'])

    # Iterate over each uploaded file
    for uploaded_file in uploaded_files:
        # Read the uploaded file as a BytesIO object
        with uploaded_file:
            # Iterate over the specified sheet names
            for sheet in sheet_names:
                try:
                    # Read the sheet into a DataFrame with engine specified
                    df = pd.read_excel(uploaded_file, sheet_name=sheet, engine='openpyxl')

                    # Add a column for the source sheet
                    df['source_sheet'] = sheet

                    # Ensure only the desired columns are included, add missing columns if any
                    for col in columns:
                        if col not in df.columns:
                            df[col] = None
                    df = df[columns + ['source_sheet']]

                    # Concatenate the data
                    merged_df = pd.concat([merged_df, df], ignore_index=True)
                except Exception as e:
                    st.write(f"Error processing file {uploaded_file.name}, sheet {sheet}: {e}")

    # Save the merged DataFrame to a single Excel file
    output_file = 'merged_output.xlsx'
    merged_df.to_excel(output_file, index=False)

    st.success('Merging complete. Merged file saved as merged_output.xlsx.')

    # Provide a download link for the merged file
    with open(output_file, "rb") as file:
        btn = st.download_button(
            label="Download Merged File",
            data=file,
            file_name=output_file,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
