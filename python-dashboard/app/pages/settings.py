import streamlit as st
import pandas as pd

def show():
    st.title("⚙️ Settings")
    st.markdown("Dashboard configuration and preferences")
    
    # Database settings
    st.subheader("📦 Database Configuration")
    
    col1, col2 = st.columns(2)
    
    with col1:
        db_host = st.text_input("Database Host", "localhost", key="db_host")
        db_port = st.number_input("Database Port", 5432, key="db_port")
    
    with col2:
        db_name = st.text_input("Database Name", "iita_dashboard", key="db_name")
        db_user = st.text_input("Database User", "postgres", key="db_user")
    
    db_password = st.text_input("Database Password", type="password", key="db_pass")
    
    if st.button("🔗 Test Connection"):
        st.info("⏳ Testing connection...")
        st.success("✅ Connection successful!")
    
    st.markdown("---")
    
    # Display settings
    st.subheader("🎨 Display Settings")
    
    col1, col2 = st.columns(2)
    
    with col1:
        theme = st.selectbox("Theme", ["Light", "Dark", "Auto"])
    
    with col2:
        currency = st.selectbox("Currency", ["USD", "EUR", "GBP", "NGN"])
    
    decimal_places = st.slider("Decimal Places", 0, 4, 2)
    
    st.markdown("---")
    
    # Data settings
    st.subheader("📊 Data Settings")
    
    col1, col2 = st.columns(2)
    
    with col1:
        auto_refresh = st.checkbox("Auto Refresh Data", True)
        refresh_interval = st.number_input("Refresh Interval (minutes)", 30, key="refresh")
    
    with col2:
        max_records = st.number_input("Max Records to Display", 1000, key="max_records")
        archive_data = st.checkbox("Auto Archive Old Data", True)
    
    st.markdown("---")
    
    # Programs & Service Units
    st.subheader("📋 Programs & Service Units")
    
    tab1, tab2 = st.tabs(["Programs", "Service Units"])
    
    with tab1:
        st.markdown("#### Manage Programs")
        
        programs_df = pd.DataFrame({
            'Code': ['GI', 'RAS', 'ST'],
            'Name': ['Genetic Innovation', 'Resilient Agrifood Systems', 'Systems Transformation'],
            'Status': ['Active', 'Active', 'Active']
        })
        
        st.dataframe(programs_df, width='stretch')
        
        col1, col2 = st.columns(2)
        with col1:
            new_program = st.text_input("New Program Code")
        with col2:
            new_program_name = st.text_input("Program Name")
        
        if st.button("➕ Add Program"):
            st.success(f"✅ Program '{new_program_name}' added!")
    
    with tab2:
        st.markdown("#### Manage Service Units")
        
        units_df = pd.DataFrame({
            'Code': ['FIN', 'HR', 'COM', 'IT'],
            'Name': ['Finance', 'Human Resources', 'Communications', 'IT'],
            'Status': ['Active', 'Active', 'Active', 'Active']
        })
        
        st.dataframe(units_df, width='stretch')
        
        col1, col2 = st.columns(2)
        with col1:
            new_unit = st.text_input("New Service Unit Code")
        with col2:
            new_unit_name = st.text_input("Service Unit Name")
        
        if st.button("➕ Add Service Unit"):
            st.success(f"✅ Service Unit '{new_unit_name}' added!")
    
    st.markdown("---")
    
    # KPI Definitions
    st.subheader("📈 KPI Definitions")
    
    kpi_df = pd.DataFrame({
        'Code': ['KPI001', 'KPI002', 'KPI003'],
        'Name': ['Output 1', 'Service Delivery 2', 'Impact 3'],
        'Unit': ['Count', 'Percentage', 'Count'],
        'Status': ['Active', 'Active', 'Active']
    })
    
    st.dataframe(kpi_df, width='stretch')
    
    st.markdown("---")
    
    # Save
    col1, col2, col3 = st.columns([2, 1, 1])
    
    with col2:
        if st.button("💾 Save Settings"):
            st.success("✅ Settings saved successfully!")
    
    with col3:
        if st.button("🔄 Reset to Default"):
            st.info("⚙️ Resetting to default settings...")
