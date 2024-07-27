import pandas as pd
import streamlit as st

# Streamlit app
st.title('Excel Sheets Merger')

# User inputs for sheet names and columns
sheet_names_input = st.text_area("Enter sheet names (comma-separated) e.g. 7amada, 7amada 1", value="")
columns_input = st.text_area("Enter column names (comma-separated) using the same pattern as before. Please blash tfty and write it as is or better copy and paste because it's case-sensitive", value="")

# Parse the user inputs
sheet_names = [name.strip() for name in sheet_names_input.split(',')]
columns = [col.strip() for col in columns_input.split(',')]

# Upload Excel files
uploaded_files = st.file_uploader("Upload Excel files", type="xlsx", accept_multiple_files=True)

# Input for the name of the output file
output_filename = st.text_input("Enter the name of the output file (without extension)", value="merged_output")

# Run button to start the merging process
if st.button('Merge'):
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
        output_file = f'{output_filename}.xlsx'
        merged_df.to_excel(output_file, index=False)

        st.success(f'Merging complete. Merged file saved as {output_file}.')

        # Provide a download link for the merged file
        with open(output_file, "rb") as file:
            btn = st.download_button(
                label="Download Merged Sheet",
                data=file,
                file_name=output_file,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.info("Please upload one or more Excel files.")
