import streamlit as st
import pandas as pd
from datetime import datetime
import os
from pathlib import Path

# Set page config
st.set_page_config(
    page_title="IITA KPI Dashboard",
    page_icon="🎯",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for better styling
st.markdown("""
<style>
    .main {
        padding-top: 1rem;
    }
    
    .metric-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        padding: 1.5rem;
        border-radius: 15px;
        text-align: center;
        box-shadow: 0 4px 15px rgba(0,0,0,0.1);
    }
    
    .metric-value {
        font-size: 2.5rem;
        font-weight: bold;
        margin: 0.5rem 0;
    }
    
    .metric-label {
        font-size: 0.9rem;
        opacity: 0.9;
    }
    
    .status-on-track {
        background-color: #d4edda;
        color: #155724;
        padding: 0.5rem 1rem;
        border-radius: 5px;
        font-weight: bold;
    }
    
    .status-at-risk {
        background-color: #fff3cd;
        color: #856404;
        padding: 0.5rem 1rem;
        border-radius: 5px;
        font-weight: bold;
    }
    
    .status-off-track {
        background-color: #f8d7da;
        color: #721c24;
        padding: 0.5rem 1rem;
        border-radius: 5px;
        font-weight: bold;
    }
    
    .header-title {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        padding: 2rem;
        border-radius: 10px;
        text-align: center;
        margin-bottom: 1.5rem;
    }
    
    .header-title h1 {
        margin: 0;
        font-size: 2.5rem;
    }
    
    .header-title p {
        margin: 0.5rem 0 0 0;
        opacity: 0.9;
    }
</style>
""", unsafe_allow_html=True)

# Initialize session state
if 'data' not in st.session_state:
    st.session_state.data = None

# Sidebar Navigation
with st.sidebar:
    st.title("🎯 IITA KPI Dashboard")
    st.markdown("---")
    
    page = st.radio(
        "📍 Navigation",
        ["Dashboard", "Upload Data", "Analytics", "Trends", "KPI Library", "Settings"]
    )
    
    st.markdown("---")
    
    # Summary info in sidebar
    st.subheader("📊 Quick Stats")
    col1, col2 = st.columns(2)
    with col1:
        st.metric("Total KPIs", 45)
    with col2:
        st.metric("On Track", 32)
    
    st.markdown("---")
    st.markdown("""
    ### About
    **IITA Programs & Services Dashboard**
    
    Real-time KPI monitoring aligned with IITA Strategy 2024–2030
    - 🔄 Programs: 3
    - 🏢 Service Units: 4
    - 📊 Metrics: 45+
    """)
    
    st.markdown("---")
    st.caption("Last updated: Today @ " + datetime.now().strftime("%H:%M"))

# Main content routing
if page == "Dashboard":
    from pages import dashboard
    dashboard.show()
elif page == "Upload Data":
    from pages import upload
    upload.show()
elif page == "Analytics":
    from pages import analytics
    analytics.show()
elif page == "Trends":
    from pages import trends
    trends.show()
elif page == "KPI Library":
    from pages import kpi_library
    kpi_library.show()
elif page == "Settings":
    from pages import settings
    settings.show()

# Footer
st.markdown("---")
st.markdown("""
<div style="text-align: center; color: #888; font-size: 0.85rem; padding: 1rem;">
    <b>IITA Programs & Services Dashboard</b> | Version 2.0 | April 2026
    <br><small>For support, contact: dashboard@iita.org</small>
</div>
""", unsafe_allow_html=True)

