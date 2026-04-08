import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from datetime import datetime, timedelta

def show():
    st.title("📊 Trend Analysis")
    st.markdown("Track KPI performance over time")
    
    # KPI Selection
    selected_kpi = st.selectbox(
        "Select KPI to Analyze",
        ["All KPIs", "Output Indicator 1", "Service Delivery 2", "Impact Metric 3", "Efficiency KPI 2"]
    )
    
    # Time range
    col1, col2 = st.columns(2)
    
    with col1:
        start_date = st.date_input(
            "Start Date",
            datetime.now() - timedelta(days=365)
        )
    
    with col2:
        end_date = st.date_input("End Date", datetime.now())
    
    st.markdown("---")
    
    # Generate sample trend data
    dates = pd.date_range(start=start_date, end=end_date, freq='D')
    trend_data = pd.DataFrame({
        'Date': dates,
        'Value': [85 + i*0.1 + (i%7)*0.5 for i in range(len(dates))],
        'Target': [90] * len(dates),
        'Moving Average (7d)': None
    })
    
    # Calculate moving average
    trend_data['Moving Average (7d)'] = trend_data['Value'].rolling(window=7).mean()
    
    # Plot
    fig = go.Figure()
    
    fig.add_trace(go.Scatter(
        x=trend_data['Date'],
        y=trend_data['Value'],
        mode='lines',
        name='Actual Value',
        line=dict(color='#667eea', width=2),
        fill='tozeroy',
        fillcolor='rgba(102, 126, 234, 0.1)'
    ))
    
    fig.add_trace(go.Scatter(
        x=trend_data['Date'],
        y=trend_data['Target'],
        mode='lines',
        name='Target',
        line=dict(color='#e74c3c', width=2, dash='dash')
    ))
    
    fig.add_trace(go.Scatter(
        x=trend_data['Date'],
        y=trend_data['Moving Average (7d)'],
        mode='lines',
        name='7-Day MA',
        line=dict(color='#2ecc71', width=2, dash='dot')
    ))
    
    fig.update_layout(
        title=f"Trend Analysis: {selected_kpi}",
        xaxis_title="Date",
        yaxis_title="Value",
        hovermode='x unified',
        height=500
    )
    
    st.plotly_chart(fig, width='stretch')
    
    st.markdown("---")
    
    # Statistics
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric("Current Value", f"{trend_data['Value'].iloc[-1]:.1f}")
    
    with col2:
        st.metric("Average", f"{trend_data['Value'].mean():.1f}")
    
    with col3:
        st.metric("Min", f"{trend_data['Value'].min():.1f}")
    
    with col4:
        st.metric("Max", f"{trend_data['Value'].max():.1f}")
    
    st.markdown("---")
    
    # Trend details
    st.subheader("Monthly Summary")
    
    monthly_data = trend_data.set_index('Date').resample('M').agg({
        'Value': ['mean', 'min', 'max'],
        'Target': 'first'
    }).round(2)
    
    st.dataframe(monthly_data, width=True)
    
    # Growth rate
    st.subheader("Period-over-Period Change")
    
    pct_change = ((trend_data['Value'].iloc[-1] - trend_data['Value'].iloc[0]) / trend_data['Value'].iloc[0] * 100)
    
    if pct_change >= 0:
        st.success(f"📈 Growth: {pct_change:+.2f}%")
    else:
        st.error(f"📉 Decline: {pct_change:+.2f}%")
