import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from utils.data_manager import DataManager
from datetime import datetime, timedelta

def show():
    # Header
    st.markdown("""
    <div class="header-title">
        <h1>📊 IITA KPI Dashboard</h1>
        <p>Real-time Performance Management & Strategic Monitoring</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Date range filter
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        period = st.selectbox(
            "Select Period",
            ["Last 30 Days", "Last Quarter", "Last Year", "All Time"]
        )
    
    with col2:
        program = st.selectbox(
            "Filter by Program",
            ["All Programs", "Genetic Innovation", "Resilient Agrifood Systems", "Systems Transformation"]
        )
    
    with col3:
        service = st.selectbox(
            "Filter by Service Unit",
            ["All Units", "Finance", "HR", "Communications", "IT"]
        )
    
    with col4:
        refresh = st.button("🔄 Refresh Data", width='stretch')
    
    st.markdown("---")
    
    # === EXECUTIVE SUMMARY SECTION ===
    st.subheader("📈 Executive Summary")
    
    col1, col2, col3, col4, col5 = st.columns(5)
    
    with col1:
        st.metric(
            label="Total KPIs",
            value="45",
            delta="+3",
            delta_color="inverse"
        )
    
    with col2:
        st.metric(
            label="On Track",
            value="32",
            delta="+5%",
            delta_color="normal"
        )
    
    with col3:
        st.metric(
            label="At Risk",
            value="10",
            delta="-2%",
            delta_color="normal"
        )
    
    with col4:
        st.metric(
            label="Off Track",
            value="3",
            delta="+1",
            delta_color="inverse"
        )
    
    with col5:
        st.metric(
            label="Achievement Rate",
            value="85.5%",
            delta="+1.2%",
            delta_color="normal"
        )
    
    st.markdown("---")
    
    # === PERFORMANCE OVERVIEW ===
    st.subheader("🎯 Performance Overview")
    
    col1, col2 = st.columns([2, 1])
    
    with col1:
        # Overall Achievement Gauge
        fig = go.Figure(go.Indicator(
            mode="gauge+number+delta",
            value=85.5,
            domain={'x': [0, 1], 'y': [0, 1]},
            title={'text': "Overall Achievement Rate (%)"},
            delta={'reference': 80, 'suffix': "%"},
            gauge={
                'axis': {'range': [0, 100]},
                'bar': {'color': "#667eea"},
                'steps': [
                    {'range': [0, 50], 'color': "#ffcccb"},
                    {'range': [50, 80], 'color': "#ffe4b5"}
                ],
                'threshold': {
                    'line': {'color': "red", 'width': 4},
                    'thickness': 0.75,
                    'value': 90
                }
            }
        ))
        fig.update_layout(height=350)
        st.plotly_chart(fig, width='stretch')
    
    with col2:
        # Status Summary
        status_data = {
            'Status': ['On Track', 'At Risk', 'Off Track'],
            'Count': [32, 10, 3],
            'Percentage': [71.1, 22.2, 6.7]
        }
        df_status = pd.DataFrame(status_data)
        
        st.write("**KPI Status Distribution**")
        for idx, row in df_status.iterrows():
            if row['Status'] == 'On Track':
                color = "🟢"
            elif row['Status'] == 'At Risk':
                color = "🟡"
            else:
                color = "🔴"
            st.write(f"{color} **{row['Status']}**: {row['Count']} KPIs ({row['Percentage']:.1f}%)")
        
        # Summary box
        st.info("""
        **Key Insights**
        - 75% of programs meeting targets
        - 3 KPIs require immediate attention
        - Overall trajectory: Positive ↗
        """)
    
    st.markdown("---")
    
    # === PROGRAM PERFORMANCE SECTION ===
    st.subheader("🏆 Program Performance")
    
    # Program data
    programs_data = {
        'Program': ['Genetic Innovation', 'Resilient Agrifood Systems', 'Systems Transformation'],
        'Achievement': [88.5, 82.3, 85.1],
        'Target': [90, 85, 90],
        'KPIs': [15, 20, 10],
        'On Track': [14, 16, 8],
        'At Risk': [1, 3, 2],
        'Off Track': [0, 1, 0]
    }
    df_programs = pd.DataFrame(programs_data)
    
    # Bar chart
    fig = go.Figure(data=[
        go.Bar(name='Achievement', x=df_programs['Program'], y=df_programs['Achievement'], 
               marker_color='#667eea', showlegend=True),
        go.Bar(name='Target', x=df_programs['Program'], y=df_programs['Target'], 
               marker_color='#764ba2', showlegend=True)
    ])
    fig.update_layout(
        barmode='group',
        title='Program Performance vs Target',
        height=400,
        xaxis_title='Program',
        yaxis_title='Achievement (%)'
    )
    st.plotly_chart(fig, width='stretch')
    
    # Program details table
    st.write("**Program Details**")
    
    display_cols = ['Program', 'Achievement', 'Target', 'KPIs', 'On Track', 'At Risk', 'Off Track']
    
    col_config = {
        'Achievement': st.column_config.NumberColumn('Achievement (%)', format='%.1f'),
        'Target': st.column_config.NumberColumn('Target (%)', format='%.1f'),
        'KPIs': st.column_config.NumberColumn('Total KPIs'),
        'On Track': st.column_config.NumberColumn('On Track'),
        'At Risk': st.column_config.NumberColumn('At Risk'),
        'Off Track': st.column_config.NumberColumn('Off Track')
    }
    
    st.dataframe(
        df_programs[display_cols],
        width='stretch',
        hide_index=True,
        column_config=col_config
    )
    
    st.markdown("---")
    
    # === TOP PERFORMERS & AREAS OF CONCERN ===
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("⭐ Top 5 Performing KPIs")
        
        top_kpis = pd.DataFrame({
            'KPI': ['Crop Varieties Released', 'Publications', 'Training Events', 'Staff Productivity', 'Tech Adoption'],
            'Achievement': [125, 110, 105, 102, 98],
            'Status': ['✅ Exceeding', '✅ Exceeding', '✅ Exceeding', '✅ On Track', '✅ On Track']
        })
        
        st.dataframe(
            top_kpis,
            width='stretch',
            hide_index=True
        )
    
    with col2:
        st.subheader("⚠️ Areas Needing Attention")
        
        concern_kpis = pd.DataFrame({
            'KPI': ['Budget Utilization', 'System Uptime', 'Community Satisfaction', 'Farmer Reach', 'Yield Improvement'],
            'Achievement': [75, 78, 82, 80, 85],
            'Status': ['🔴 Critical', '🟡 At Risk', '🟡 At Risk', '🟡 At Risk', '🟢 Monitor']
        })
        
        st.dataframe(
            concern_kpis,
            width='stretch',
            hide_index=True
        )
    
    st.markdown("---")
    
    # === SERVICE UNIT PERFORMANCE ===
    st.subheader("🏢 Service Unit Performance")
    
    service_data = {
        'Service Unit': ['Finance', 'HR', 'Communications', 'IT'],
        'Efficiency Score': [88, 85, 82, 90],
        'Satisfaction': [4.2, 4.1, 3.8, 4.3],
        'Projects Completed': [12, 8, 15, 6]
    }
    df_services = pd.DataFrame(service_data)
    
    fig = px.bar(df_services, x='Service Unit', y='Efficiency Score', 
                 color='Efficiency Score', color_continuous_scale='Viridis',
                 title='Service Unit Efficiency Scores')
    st.plotly_chart(fig, width='stretch')
    
    st.dataframe(
        df_services,
        width='stretch',
        hide_index=True
    )
    
    st.markdown("---")
    
    # === RECENT UPDATES ===
    st.subheader("📝 Recent KPI Updates")
    
    recent_data = pd.DataFrame({
        'KPI Code': ['KPI001', 'KPI005', 'KPI012', 'KPI003', 'KPI008'],
        'KPI Name': ['Crop Varieties', 'Budget Utilization', 'System Uptime', 'Training Events', 'Yield Improvement'],
        'Current Value': [15, 82.5, 98.5, 24, 18.5],
        'Target': [12, 95, 99.9, 24, 20],
        'Variance': ['+25%', '-13.2%', '-1.4%', '100%', '-7.5%'],
        'Status': ['✅', '🟡', '🟡', '✅', '⚠️'],
        'Last Updated': ['2026-04-08', '2026-04-07', '2026-04-08', '2026-04-06', '2026-04-05']
    })
    
    st.dataframe(
        recent_data,
        width='stretch',
        hide_index=True
    )
    
    st.markdown("---")
    
    # === QUICK INSIGHTS ===
    st.subheader("💡 Quick Insights & Recommendations")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.success("""
        **✅ Strong Performance**
        
        - Genetic Innovation program trending upward
        - Publication targets exceeded by 10%
        - Staff productivity at excellent levels
        """)
    
    with col2:
        st.warning("""
        **⚠️ Areas to Monitor**
        
        - Budget utilization below 85%
        - System downtime trending up
        - Community satisfaction declining
        """)
    
    with col3:
        st.info("""
        **💡 Recommendations**
        
        - Increase budget allocation review
        - Schedule IT infrastructure assessment
        - Launch community engagement campaign
        """)

