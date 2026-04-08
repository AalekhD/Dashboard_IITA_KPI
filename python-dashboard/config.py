# IITA Dashboard Configuration
# This file contains environment variables and configuration

# Application Settings
APP_NAME = "IITA Program & Service Dashboard"
APP_VERSION = "0.1.0"

# Data Storage
DATA_DIR = "data"
UPLOADS_DIR = "uploads"
MAX_UPLOAD_SIZE = 10 * 1024 * 1024  # 10 MB

# Programs
PROGRAMS = {
    'GI': 'Genetic Innovation',
    'RAS': 'Resilient Agrifood Systems',
    'ST': 'Systems Transformation'
}

# Service Units
SERVICE_UNITS = {
    'FIN': 'Finance',
    'HR': 'Human Resources',
    'COM': 'Communications',
    'IT': 'Information Technology'
}

# KPI Categories
KPI_CATEGORIES = [
    'Output Indicators',
    'Service Delivery',
    'Efficiency Indicators',
    'Impact Metrics',
    'Financial Performance'
]

# Database (for future implementation)
# DATABASE_URL = "postgresql://user:password@localhost:5432/iita_dashboard"

# Streamlit Configuration
STREAMLIT_CONFIG = {
    'theme': {
        'primaryColor': '#667eea',
        'backgroundColor': '#ffffff',
        'secondaryBackgroundColor': '#f0f2f6',
        'textColor': '#262730',
        'font': 'sans serif'
    }
}
