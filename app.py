import pandas as pd
import streamlit as st
import openpyxl  # Ensure openpyxl is imported

# Streamlit app
st.title('Excel Sheets Merger')

# User inputs for sheet names and columns
sheet_names_input = st.text_area("Enter sheet names (comma-separated)", value="H1 ICD-10-CM (Non-IQVIA), Cluster1, Cluster2, Cluster3, Cluster4, Cluster5, Cluster6, Cluster7, Cluster8, Cluster9, Cluster10")
columns_input = st.text_area("Enter column names (comma-separated)", value="CODE, DESCRIPTION, Decision, Mendel ID, Concept Name, Missing Concept, Parent Mendel ID If Missing Concept, Parent Concept Name If Missing Concept")

# Parse the user inputs
sheet_names = [name.strip() for name in sheet_names_input.split(',')]
columns = [col.strip() for col in columns_input.split(',')]

# Upload Excel files
uploaded_files = st.file_uploader("Upload Excel files", type="xlsx", accept_multiple_files=True)

if uploaded_files:
    # Initialize an empty DataFrame to hold the merged data
    merged_df = pd.DataFrame(columns=columns + ['source_sheet'])

    # Create a progress bar for the files
    progress_bar = st.progress(0)
    total_steps = len(uploaded_files) * len(sheet_names)
    step = 0

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
                    st.error(f"Error processing file {uploaded_file.name}, sheet {sheet}: {e}")

                # Update the progress bar
                step += 1
                progress_bar.progress(step / total_steps)

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
else:
    st.info("Please upload one or more Excel files.")
