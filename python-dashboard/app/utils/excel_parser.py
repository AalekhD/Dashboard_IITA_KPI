import pandas as pd
import io
from datetime import datetime

def parse_excel_file(file):
    """
    Parse uploaded Excel file and validate structure
    
    Args:
        file: Streamlit uploaded file object
        
    Returns:
        DataFrame or None if parsing fails
    """
    try:
        if file.type == 'text/csv':
            df = pd.read_csv(file)
        else:
            df = pd.read_excel(file)
        
        # Validate required columns
        required_columns = ['kpi_code', 'period_date', 'value']
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            raise ValueError(f"Missing required columns: {', '.join(missing_columns)}")
        
        # Convert period_date to datetime
        if 'period_date' in df.columns:
            df['period_date'] = pd.to_datetime(df['period_date'], errors='coerce')
        
        # Convert value to numeric
        if 'value' in df.columns:
            df['value'] = pd.to_numeric(df['value'], errors='coerce')
        
        return df
    
    except Exception as e:
        print(f"Error parsing file: {str(e)}")
        return None
