import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from utils.data_manager import DataManager

def show():
    st.title("� Analytics & Performance Analysis")
    st.markdown("Deep dive into program and service performance metrics")
    
    dm = DataManager()
    
    # Filter section
    st.markdown("### 🔍 Filter Options")
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        programs = ["All", "Genetic Innovation", "Resilient Agrifood Systems", "Systems Transformation"]
        selected_program = st.selectbox("Program", programs)
    
    with col2:
        service_units = ["All", "Finance", "HR", "Communications", "IT"]
        selected_service = st.selectbox("Service Unit", service_units)
    
    with col3:
        period = st.selectbox(
            "Period",
            ["Last 30 Days", "Last Quarter", "Last 6 Months", "Last Year"]
        )
    
    with col4:
        st.empty()
    
    st.markdown("---")
    
    # Tabs - Enhanced with more tabs
    tab1, tab2, tab3, tab4, tab5 = st.tabs(["Overview", "Variance Analysis", "Distribution", "Comparison", "Detailed Report"])
    
    with tab1:
        st.subheader("📈 Performance Overview")
        
        col1, col2 = st.columns([2, 1])
        
        with col1:
            # Simulated data
            perf_data = pd.DataFrame({
                'KPI': ['Output 1', 'Output 2', 'Efficiency 1', 'Service 1', 'Impact 1', 'Financial 1'],
                'Target': [160, 200, 95, 100, 85, 90],
                'Actual': [150, 190, 92, 98, 80, 82],
                'Performance': [93.75, 95, 96.84, 98, 94.12, 91.11]
            })
            
            fig = px.bar(
                perf_data,
                x='KPI',
                y='Performance',
                color='Performance',
                color_continuous_scale='RdYlGn',
                range_color=[80, 100]
            )
            fig.add_hline(y=90, line_dash="dash", line_color="blue", annotation_text="Target: 90%")
            fig.update_layout(height=400, showlegend=False)
            st.plotly_chart(fig, width='stretch')
        
        with col1:
            # Performance table
            st.write("**KPI Performance Metrics**")
            st.dataframe(perf_data, width='stretch', hide_index=True)
        
        with col2:
            # Statistics
            st.write("**Quick Statistics**")
            st.metric("Average Performance", f"{perf_data['Performance'].mean():.1f}%")
            st.metric("Best KPI", f"{perf_data['Performance'].max():.1f}%")
            st.metric("Lowest KPI", f"{perf_data['Performance'].min():.1f}%")
            
            st.info("""
            **Performance Indicators**
            
            - 🟢 > 90%: Excellent
            - 🟡 80-90%: Good
            - 🔴 < 80%: Needs Attention
            """)
    
    with tab2:
        st.subheader("📊 Variance Analysis")
        
        var_data = pd.DataFrame({
            'KPI': ['KPI001', 'KPI002', 'KPI003', 'KPI004', 'KPI005', 'KPI006'],
            'Variance %': [-6.25, -5.0, -3.2, 2.5, -7.6, 1.8],
            'Category': ['Output', 'Output', 'Service', 'Impact', 'Financial', 'Efficiency']
        })
        
        fig = px.bar(
            var_data,
            x='KPI',
            y='Variance %',
            color='Variance %',
            color_continuous_scale=['red', 'yellow', 'green'],
            range_color=[-10, 5]
        )
        fig.add_hline(y=0, line_dash="dash", line_color="black")
        fig.update_layout(height=400, showlegend=False)
        st.plotly_chart(fig, width='stretch')
        
        st.write("**Variance Details**")
        st.dataframe(var_data, width='stretch', hide_index=True)
        
        # Analysis
        col1, col2 = st.columns(2)
        with col1:
            st.success(f"✅ {len(var_data[var_data['Variance %'] > 0])} KPIs exceeding targets")
        with col2:
            st.error(f"❌ {len(var_data[var_data['Variance %'] < 0])} KPIs below targets")
    
    with tab3:
        st.subheader("🎯 KPI Distribution & Status")
        
        col1, col2 = st.columns(2)
        
        with col1:
            # Status distribution
            status_dist = pd.DataFrame({
                'Status': ['On Target', 'Above Target', 'Below Target (Acceptable)', 'Below Target (Critical)'],
                'Count': [32, 3, 8, 2]
            })
            
            fig = px.pie(
                status_dist,
                values='Count',
                names='Status',
                color_discrete_map={
                    'On Target': '#2ecc71',
                    'Above Target': '#3498db',
                    'Below Target (Acceptable)': '#f39c12',
                    'Below Target (Critical)': '#e74c3c'
                }
            )
            fig.update_layout(height=400)
            st.plotly_chart(fig, width='stretch')
        
        with col2:
            # Category distribution
            cat_dist = pd.DataFrame({
                'Category': ['Output', 'Service', 'Efficiency', 'Impact', 'Financial'],
                'Count': [10, 8, 12, 10, 5]
            })
            
            fig = px.bar(
                cat_dist,
                x='Category',
                y='Count',
                color='Count',
                color_continuous_scale='Blues'
            )
            fig.update_layout(height=400, showlegend=False)
            st.plotly_chart(fig, width='stretch')
    
    with tab4:
        st.subheader("📊 Period Comparison")
        
        comp_data = pd.DataFrame({
            'Month': ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'],
            'Achievement %': [82, 84, 85, 86, 86, 85.5],
            'Target %': [85, 85, 85, 90, 90, 90]
        })
        
        fig = go.Figure()
        fig.add_trace(go.Scatter(
            x=comp_data['Month'],
            y=comp_data['Achievement %'],
            mode='lines+markers',
            name='Achievement',
            line=dict(color='#667eea', width=3),
            marker=dict(size=10)
        ))
        fig.add_trace(go.Scatter(
            x=comp_data['Month'],
            y=comp_data['Target %'],
            mode='lines+markers',
            name='Target',
            line=dict(color='#764ba2', width=2, dash='dash'),
            marker=dict(size=8)
        ))
        
        fig.update_layout(
            title='Achievement vs Target Over Time',
            xaxis_title='Month',
            yaxis_title='Percentage (%)',
            hovermode='x unified',
            height=400
        )
        st.plotly_chart(fig, width='stretch')
        
        # Trend analysis
        col1, col2, col3 = st.columns(3)
        with col1:
            st.info(f"📈 Trend: +{comp_data['Achievement %'].iloc[-1] - comp_data['Achievement %'].iloc[0]:.1f}pp")
        with col2:
            st.metric("Current", f"{comp_data['Achievement %'].iloc[-1]:.1f}%")
        with col3:
            st.metric("Average", f"{comp_data['Achievement %'].mean():.1f}%")
    
    with tab5:
        st.subheader("📋 Detailed Performance Report")
        
        # Generate report
        report_data = pd.DataFrame({
            'Program': ['GI', 'GI', 'RAS', 'RAS', 'ST', 'ST'],
            'KPI Category': ['Output', 'Impact', 'Output', 'Service', 'Efficiency', 'Financial'],
            'Target': [12, 85, 150, 100, 90, 95],
            'Actual': [15, 82, 145, 98, 92, 82],
            'Achievement %': [125, 96.5, 96.7, 98, 102.2, 86.3],
            'Status': ['✅ Exceeding', '✅ On Track', '✅ On Track', '✅ On Track', '✅ Exceeding', '🟡 At Risk']
        })
        
        st.dataframe(report_data, width='stretch', hide_index=True)
        
        # Export options
        col1, col2, col3 = st.columns(3)
        with col1:
            csv = report_data.to_csv(index=False)
            st.download_button(
                label="📥 Download as CSV",
                data=csv,
                file_name="analytics_report.csv",
                mime="text/csv"
            )
        with col2:
            st.button("📧 Email Report")
        with col3:
            st.button("🖨️ Print Report")

