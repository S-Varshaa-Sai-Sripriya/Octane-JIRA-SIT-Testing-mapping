import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO

st.set_page_config(page_title="Octane ID - JIRA ID Mapper", layout="wide")

st.title("🔄 Octane ID - JIRA ID Mapper")
st.write("Upload an Excel file to map Octane IDs with JIRA IDs")

# File uploader
uploaded_file = st.file_uploader("Choose an Excel file", type=['xlsx', 'xls'])

if uploaded_file is not None:
    try:
        # Read the Excel file
        df = pd.read_excel(uploaded_file)
        
        st.subheader("Input Data Preview")
        st.dataframe(df.head(10), use_container_width=True)
        
        if st.button("💻 Compute Mapping", key="compute_btn"):
            with st.spinner("Processing..."):
                # Create output data
                output_rows = []
                
                # Iterate through each row
                for idx, row in df.iterrows():
                    test_team = row.get('Test Team', '')
                    octane_id = row.get('ID', '')
                    jira_ids = row.get('Test: JIRA ID', '')
                    
                    # Skip rows with missing data
                    if pd.isna(test_team) or pd.isna(octane_id):
                        continue
                    
                    # Convert to string and handle empty JIRA IDs
                    test_team = str(test_team).strip()
                    octane_id = str(int(octane_id)) if pd.notna(octane_id) else ''
                    jira_ids_str = str(jira_ids).strip() if pd.notna(jira_ids) else ''
                    
                    if not jira_ids_str or jira_ids_str == '':
                        # If no JIRA ID, still add the row
                        output_rows.append({
                            'Test Team': test_team,
                            'Octane ID': octane_id,
                            'JIRA ID': ''
                        })
                    else:
                        # Split JIRA IDs by comma and create a row for each
                        jira_list = [j.strip() for j in jira_ids_str.split(',')]
                        for jira_id in jira_list:
                            output_rows.append({
                                'Test Team': test_team,
                                'Octane ID': octane_id,
                                'JIRA ID': jira_id
                            })
                
                # Create output dataframe
                output_df = pd.DataFrame(output_rows)
                
                st.subheader("✅ Output Data")
                st.dataframe(output_df, use_container_width=True)
                
                st.success(f"✓ Successfully processed {len(output_df)} rows!")
                
                # Download button
                output_bytes = BytesIO()
                with pd.ExcelWriter(output_bytes, engine='openpyxl') as writer:
                    output_df.to_excel(writer, sheet_name='Mapped Data', index=False)
                output_bytes.seek(0)
                
                st.download_button(
                    label="📥 Download Output Excel",
                    data=output_bytes.getvalue(),
                    file_name="octane_jira_mapping_output.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
    except Exception as e:
        st.error(f"❌ Error processing file: {str(e)}")
else:
    st.info("👆 Please upload an Excel file to begin")
