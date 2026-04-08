import streamlit as st
import pandas as pd
import plotly.express as px

def show():
    st.title("📚 KPI Library")
    st.markdown("Comprehensive KPI Definitions and Metadata")
    
    # Filter options
    col1, col2, col3 = st.columns([2, 2, 2])
    
    with col1:
        program = st.selectbox(
            "Filter by Program",
            ["All Programs", "Genetic Innovation", "Resilient Agrifood Systems", "Systems Transformation"]
        )
    
    with col2:
        category = st.selectbox(
            "Filter by Category",
            ["All", "Output Indicators", "Service Delivery", "Efficiency", "Impact", "Financial"]
        )
    
    with col3:
        status = st.selectbox(
            "Filter by Status",
            ["All", "Active", "Inactive"]
        )
    
    st.markdown("---")
    
    # KPI Library Data
    kpi_data = pd.DataFrame({
        'KPI Code': ['KPI001', 'KPI002', 'KPI003', 'KPI004', 'KPI005', 'KPI006', 'KPI007', 'KPI008', 'KPI009', 'KPI010'],
        'KPI Name': [
            'Crop Varieties Released',
            'Research Publications',
            'Farmers Reached',
            'Training Events',
            'Project Budget Utilization',
            'Staff Productivity Index',
            'Technology Adoption Rate',
            'Average Yield Improvement',
            'Community Satisfaction Score',
            'System Uptime'
        ],
        'Program': ['GI', 'GI', 'RAS', 'RAS', 'ST', 'ST', 'GI', 'RAS', 'ST', 'ST'],
        'Category': ['Output', 'Output', 'Output', 'Service', 'Financial', 'Efficiency', 'Impact', 'Impact', 'Impact', 'Efficiency'],
        'Unit': ['Count', 'Count', 'Count', 'Count', 'Percentage', 'Index', 'Percentage', 'Percentage', 'Score', 'Percentage'],
        'Target': [12, 50, 10000, 24, 95, 85, 75, 20, 4.5, 99.9],
        'Status': ['Active', 'Active', 'Active', 'Active', 'Active', 'Active', 'Active', 'Active', 'Active', 'Active']
    })
    
    st.subheader(f"📊 Available KPIs ({len(kpi_data)} Total)")
    
    # Tabs for different views
    tab1, tab2, tab3 = st.tabs(["All KPIs", "By Program", "By Category"])
    
    with tab1:
        st.dataframe(
            kpi_data,
            width='stretch',
            hide_index=True,
            column_config={
                "KPI Code": st.column_config.TextColumn("Code", width="small"),
                "KPI Name": st.column_config.TextColumn("KPI Name", width="medium"),
                "Program": st.column_config.TextColumn("Program", width="small"),
                "Category": st.column_config.TextColumn("Category", width="small"),
                "Unit": st.column_config.TextColumn("Unit", width="small"),
                "Target": st.column_config.NumberColumn("Target", format="%g"),
                "Status": st.column_config.TextColumn("Status", width="small")
            }
        )
    
    with tab2:
        # KPIs by program
        program_dist = kpi_data.groupby('Program').size().reset_index(name='Count')
        fig = px.bar(program_dist, x='Program', y='Count', title='KPIs by Program')
        st.plotly_chart(fig, width='stretch')
        
        # Detailed list
        for prog in ['GI', 'RAS', 'ST']:
            prog_name = {'GI': 'Genetic Innovation', 'RAS': 'Resilient Agrifood Systems', 'ST': 'Systems Transformation'}[prog]
            with st.expander(f"📋 {prog_name}"):
                prog_kpis = kpi_data[kpi_data['Program'] == prog]
                st.dataframe(prog_kpis, width='stretch', hide_index=True)
    
    with tab3:
        # KPIs by category
        cat_dist = kpi_data.groupby('Category').size().reset_index(name='Count')
        fig = px.pie(cat_dist, names='Category', values='Count', title='KPI Distribution by Category')
        st.plotly_chart(fig, width='stretch')
        
        # Detailed list
        for cat in kpi_data['Category'].unique():
            with st.expander(f"📊 {cat}"):
                cat_kpis = kpi_data[kpi_data['Category'] == cat]
                st.dataframe(cat_kpis, width='stretch', hide_index=True)
    
    st.markdown("---")
    
    # KPI Details
    st.subheader("🔍 KPI Details")
    
    col1, col2 = st.columns(2)
    
    with col1:
        selected_kpi = st.selectbox(
            "Select a KPI to view details",
            kpi_data['KPI Name'].tolist()
        )
        
        kpi_selected = kpi_data[kpi_data['KPI Name'] == selected_kpi].iloc[0]
        
        st.write(f"""
        **KPI Code:** {kpi_selected['KPI Code']}
        **Category:** {kpi_selected['Category']}
        **Program:** {kpi_selected['Program']}
        **Unit:** {kpi_selected['Unit']}
        **Target Value:** {kpi_selected['Target']}
        **Status:** {kpi_selected['Status']}
        """)
    
    with col2:
        st.info("""
        **Data Quality Guidelines**
        - All values should be numeric
        - Updates: Monthly or Quarterly
        - Validation: Required before submission
        - Ownership: Program Director
        """)
