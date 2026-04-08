import pandas as pd
import json
from pathlib import Path
from datetime import datetime

class DataManager:
    """
    Manages data storage and retrieval
    Uses local CSV files for storage (can be extended to use database)
    """
    
    def __init__(self, data_dir="data"):
        self.data_dir = Path(data_dir)
        self.data_dir.mkdir(exist_ok=True)
        self.kpi_file = self.data_dir / "kpi_data.csv"
        self.history_file = self.data_dir / "upload_history.json"
    
    def save_kpi_data(self, df: pd.DataFrame) -> tuple[bool, str]:
        """
        Save KPI data to CSV file
        
        Args:
            df: DataFrame with KPI data
            
        Returns:
            Tuple of (success: bool, message: str)
        """
        try:
            # Add timestamp if not present
            if 'uploaded_at' not in df.columns:
                df['uploaded_at'] = datetime.now()
            
            # Append to existing file or create new
            if self.kpi_file.exists():
                existing_df = pd.read_csv(self.kpi_file)
                df = pd.concat([existing_df, df], ignore_index=True)
            
            df.to_csv(self.kpi_file, index=False)
            
            # Update upload history
            self._log_upload(len(df), len(df), 0)
            
            return True, f"Successfully uploaded {len(df)} records"
        
        except Exception as e:
            return False, f"Error saving data: {str(e)}"
    
    def get_kpi_data(self, filters=None) -> pd.DataFrame:
        """
        Retrieve KPI data with optional filters
        
        Args:
            filters: Dict with filter criteria
            
        Returns:
            DataFrame with KPI data
        """
        try:
            if not self.kpi_file.exists():
                return pd.DataFrame()
            
            df = pd.read_csv(self.kpi_file)
            
            if filters:
                if 'program' in filters:
                    df = df[df['program_code'] == filters['program']]
                if 'start_date' in filters:
                    df['period_date'] = pd.to_datetime(df['period_date'])
                    df = df[df['period_date'] >= filters['start_date']]
                if 'end_date' in filters:
                    df['period_date'] = pd.to_datetime(df['period_date'])
                    df = df[df['period_date'] <= filters['end_date']]
            
            return df
        
        except Exception as e:
            print(f"Error retrieving data: {str(e)}")
            return pd.DataFrame()
    
    def get_upload_history(self):
        """
        Get upload history
        
        Returns:
            DataFrame with upload history
        """
        try:
            if not self.history_file.exists():
                return pd.DataFrame()
            
            with open(self.history_file, 'r') as f:
                history = json.load(f)
            
            return pd.DataFrame(history)
        
        except Exception as e:
            print(f"Error retrieving history: {str(e)}")
            return pd.DataFrame()
    
    def _log_upload(self, total_records: int, successful: int, failed: int):
        """
        Log upload event
        
        Args:
            total_records: Total records processed
            successful: Successfully uploaded
            failed: Failed uploads
        """
        try:
            history = []
            if self.history_file.exists():
                with open(self.history_file, 'r') as f:
                    history = json.load(f)
            
            upload_record = {
                'timestamp': datetime.now().isoformat(),
                'total_records': total_records,
                'successful_records': successful,
                'failed_records': failed,
                'status': 'Success' if failed == 0 else 'Partial'
            }
            
            history.append(upload_record)
            
            with open(self.history_file, 'w') as f:
                json.dump(history, f, indent=2)
        
        except Exception as e:
            print(f"Error logging upload: {str(e)}")
