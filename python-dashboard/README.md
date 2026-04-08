# IITA Program & Service Dashboard - Python/Streamlit Version

A lightweight, interactive KPI dashboard built with **Streamlit** and **Python** for real-time data visualization and performance monitoring.

## 🎯 Features

- **📊 Interactive Dashboard**: Real-time KPI visualization with Plotly charts
- **📤 Excel Upload**: Direct import of KPI data from Excel/CSV files
- **📈 Trend Analysis**: Historical trends with moving averages
- **📉 Performance Analytics**: Variance analysis and period comparisons
- **⚙️ Configuration Management**: Program, service unit, and KPI settings
- **🎨 Responsive Design**: Works on desktop and mobile devices

## 🚀 Quick Start

### Prerequisites
- Python 3.8+
- pip or conda

### Installation

1. **Create Virtual Environment**
```bash
cd python-dashboard
python -m venv venv

# Activate (Windows)
venv\Scripts\activate

# Activate (macOS/Linux)
source venv/bin/activate
```

2. **Install Dependencies**
```bash
pip install -r requirements.txt
```

3. **Run Dashboard**
```bash
streamlit run app/main.py
```

Dashboard opens at `http://localhost:8501`

## 📁 Project Structure

```
python-dashboard/
├── app/
│   ├── main.py                 # Main Streamlit app
│   ├── pages/
│   │   ├── dashboard.py        # Dashboard overview
│   │   ├── upload.py           # Data upload page
│   │   ├── analytics.py        # Analytics & insights
│   │   ├── trends.py           # Trend analysis
│   │   └── settings.py         # Configuration
│   ├── utils/
│   │   ├── data_manager.py     # Data storage logic
│   │   └── excel_parser.py     # Excel file parsing
│   └── models/
│       └── (database models - future implementation)
├── data/                       # Local data storage
├── uploads/                    # Uploaded files
├── requirements.txt            # Python dependencies
├── config.py                   # Configuration settings
└── README.md
```

## 📊 Pages

### 1. **Dashboard** 🏠
- KPI overview and key metrics
- Program performance comparison
- Recent KPI updates
- Status gauges and visualizations

### 2. **Upload Data** 📤
- Excel/CSV file upload
- Data validation
- Upload history
- Sample template download

### 3. **Analytics** 📈
- Performance analysis
- Variance analysis
- KPI distribution
- Period-over-period comparison

### 4. **Trends** 📊
- Historical trend visualization
- Moving averages
- Statistical summary
- Growth rate calculation

### 5. **Settings** ⚙️
- Database configuration
- Display preferences
- Program management
- KPI definitions

## 📋 Upload File Format

### Required Columns
- `kpi_code`: Unique KPI identifier (e.g., KPI001)
- `period_date`: Date in YYYY-MM-DD format
- `value`: Numeric value

### Optional Columns
- `program_code`: Program identifier (GI, RAS, ST)
- `service_unit_code`: Service unit identifier
- `target`: Target value
- `data_source`: Data source name

### Example
```
kpi_code,program_code,period_date,value,target
KPI001,GI,2025-12-31,150,160
KPI002,RAS,2025-12-31,85.5,90
KPI003,ST,2025-12-31,92,95
```

## 🔧 Configuration

Edit `config.py` to customize:
- Programs and Service Units
- KPI Categories
- Data directories
- App theme colors

## 💾 Data Storage

Currently uses **CSV files** for storage:
- `data/kpi_data.csv`: KPI records
- `data/upload_history.json`: Upload logs

**Future**: Easily extensible to PostgreSQL or other databases

## 🎨 Customization

### Change Theme Colors
Edit `config.py`:
```python
STREAMLIT_CONFIG = {
    'theme': {
        'primaryColor': '#667eea',
        'backgroundColor': '#ffffff',
        ...
    }
}
```

### Add New Pages
1. Create file in `app/pages/`
2. Add import and navigation in `app/main.py`
3. Implement `show()` function

## 📦 Dependencies

- **streamlit**: Web app framework
- **pandas**: Data manipulation
- **plotly**: Interactive charts
- **openpyxl**: Excel file handling
- **sqlalchemy**: ORM (for future DB integration)

## 🚀 Deployment

### Local Server
```bash
streamlit run app/main.py
```

### Streamlit Cloud
```bash
# Push to GitHub, then deploy at share.streamlit.io
```

### Docker
```bash
docker build -t kpi-dashboard .
docker run -p 8501:8501 kpi-dashboard
```

## 📝 Development

### Add Logging
```python
import logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)
```

### Add Caching
```python
@st.cache_data
def load_data():
    return pd.read_csv("data.csv")
```

## 🛣️ Roadmap

- [ ] PostgreSQL integration
- [ ] User authentication
- [ ] Advanced filtering
- [ ] Export reports (PDF, Excel)
- [ ] Real-time data sync
- [ ] Mobile app
- [ ] API backend

## 📞 Support

For issues or questions, refer to:
- [Streamlit Documentation](https://docs.streamlit.io)
- [Plotly Documentation](https://plotly.com/python)

## 📄 License

IITA Internal Use

---

**Created**: April 2025
**Status**: Active Development
