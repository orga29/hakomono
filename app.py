import streamlit as st
import pandas as pd
from logic import load_source_data, process_data, write_to_template
import io
from datetime import datetime

st.set_page_config(page_title="Hakomono Aggregation App", layout="wide")

st.title("Hakomono Aggregation App")
st.markdown("Upload the source file and template to generate the aggregation report.")

# File Uploaders
col1, col2 = st.columns(2)
with col1:
    source_file = st.file_uploader("Source File (受注集計表)", type=["xlsx", "xlsm"])
with col2:
    template_file = st.file_uploader("Template File (集計表)", type=["xlsx", "xlsm"])

if source_file and template_file:
    if st.button("Run Aggregation"):
        try:
            with st.spinner("Processing..."):
                # Load Source
                filtered_df, col_mapping = load_source_data(source_file)
                st.success(f"Loaded source data. Found {len(filtered_df)} 'Box' items.")
                
                # Process
                df_koda, df_yamato, koda_headers, yamato_headers, yamato_delivery_types = process_data(filtered_df, col_mapping)
                st.info(f"Split data: {len(df_koda.columns)} cols for Koda, {len(df_yamato.columns)} cols for Yamato.")
                
                # Write to Template
                # We need to reset the template file pointer because it might have been read or we need a fresh copy
                template_file.seek(0)
                
                # Generate filename
                today_str = datetime.now().strftime("%y%m%d")
                output_filename = f"集計表{today_str}.xlsm"
                
                output_buffer = write_to_template(template_file, df_koda, df_yamato, koda_headers, yamato_headers, yamato_delivery_types, output_filename)
                
                st.success("Aggregation complete!")
                
                # Download Button
                st.download_button(
                    label="Download Result",
                    data=output_buffer,
                    file_name=output_filename,
                    mime="application/vnd.ms-excel.sheet.macroEnabled.12"
                )
                
                # Preview (Optional)
                with st.expander("Preview Data (Koda)"):
                    st.dataframe(df_koda.head())
                with st.expander("Preview Data (Yamato)"):
                    st.dataframe(df_yamato.head())
                    
        except Exception as e:
            st.error(f"An error occurred: {e}")
            st.exception(e)
