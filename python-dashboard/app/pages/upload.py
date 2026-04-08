import streamlit as st
import pandas as pd
from pathlib import Path
from utils.excel_parser import parse_excel_file
from utils.data_manager import DataManager

def show():
    st.title("📤 Upload KPI Data")
    st.markdown("Import KPI data from Excel files")
    
    dm = DataManager()
    
    # Upload section
    st.subheader("Upload Excel File")
    uploaded_file = st.file_uploader(
        "Choose an Excel file (.xlsx, .csv)",
        type=['xlsx', 'csv']
    )
    
    if uploaded_file is not None:
        try:
            # Show file info
            col1, col2, col3 = st.columns(3)
            with col1:
                st.info(f"📁 File: {uploaded_file.name}")
            with col2:
                st.info(f"📊 Size: {uploaded_file.size / 1024:.2f} KB")
            with col3:
                st.info(f"⏰ Type: {uploaded_file.type}")
            
            st.markdown("---")
            
            # Parse file
            df = parse_excel_file(uploaded_file)
            
            if df is not None:
                st.success("✅ File parsed successfully!")
                
                # Show preview
                st.subheader("Data Preview")
                st.dataframe(df.head(10), width='stretch')
                
                # Validation
                st.subheader("Data Validation")
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    st.metric("Total Records", len(df))
                
                with col2:
                    valid_records = df.dropna().shape[0]
                    st.metric("Valid Records", valid_records)
                
                with col3:
                    invalid_records = len(df) - valid_records
                    st.metric("Invalid Records", invalid_records, delta_color="inverse")
                
                # Show any issues
                if invalid_records > 0:
                    st.warning(f"⚠️ {invalid_records} records have missing values")
                    invalid_data = df[df.isna().any(axis=1)]
                    st.dataframe(invalid_data, width='stretch')
                
                st.markdown("---")
                
                # Upload button
                col1, col2, col3 = st.columns([2, 1, 1])
                
                with col2:
                    if st.button("✅ Confirm & Upload", key="upload_btn"):
                        # Save to database
                        success, message = dm.save_kpi_data(df)
                        
                        if success:
                            st.success(f"✅ {message}")
                            st.balloons()
                        else:
                            st.error(f"❌ {message}")
                
                with col3:
                    if st.button("📥 Download Sample", key="download_sample"):
                        sample_data = pd.DataFrame({
                            'kpi_code': ['KPI001', 'KPI002'],
                            'program_code': ['GI', 'RAS'],
                            'service_unit_code': ['', 'SU01'],
                            'period_date': ['2025-12-31', '2025-12-31'],
                            'value': [150, 85.5],
                            'target': [160, 90],
                            'data_source': ['Programs', 'Services']
                        })
                        st.download_button(
                            label="📥 Sample.xlsx",
                            data=sample_data.to_csv(index=False),
                            file_name="kpi_sample.csv"
                        )
        
        except Exception as e:
            st.error(f"❌ Error processing file: {str(e)}")
    
    # Instructions
    st.markdown("---")
    st.subheader("📋 File Format Requirements")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("""
        ### Required Columns:
        - **kpi_code**: Unique KPI identifier
        - **period_date**: Date (YYYY-MM-DD)
        - **value**: Numeric value
        
        ### Optional Columns:
        - **program_code**: Program identifier
        - **service_unit_code**: Service unit identifier
        - **target**: Target value
        - **data_source**: Data source name
        """)
    
    with col2:
        st.markdown("""
        ### Example Data:
        | kpi_code | period_date | value | target |
        |----------|------------|-------|--------|
        | KPI001   | 2025-12-31 | 150   | 160    |
        | KPI002   | 2025-12-31 | 85.5  | 90     |
        | KPI003   | 2025-12-31 | 92    | 95     |
        """)
    
    # Upload history
    st.markdown("---")
    st.subheader("📝 Upload History")
    
    history = dm.get_upload_history()
    if history is not None and len(history) > 0:
        st.dataframe(history, width='stretch')
    else:
        st.info("No uploads yet")
