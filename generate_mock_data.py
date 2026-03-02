import pandas as pd
from datetime import datetime, timedelta

def create_mock_excel():
    now = datetime.now()
    
    data = {
        'Plate Number': ['ABC-1234', 'XYZ-5678', 'DEF-9012', 'MMM-1111', 'NNN-2222', 'OOO-3333'],
        'Vehicle Model': ['Toyota Vios', 'Honda Civic', 'Ford Everest', 'Mitsubishi Montero', 'Nissan Navara', 'Suzuki Ertiga'],
        'Expiration Date': [
            (now - timedelta(days=5)).strftime('%Y-%m-%d'), # Expired
            (now + timedelta(days=5)).strftime('%Y-%m-%d'), # Days Before Expiry
            (now + timedelta(days=20)).strftime('%Y-%m-%d'), # 2-Week Notice
            (now + timedelta(days=45)).strftime('%Y-%m-%d'), # Sufficient Time
            '', # Empty
            (now + timedelta(days=10)).strftime('%Y-%m-%d')  # Will be overridden by Registered status
        ],
        'Status Column': ['', '', '', '', '', 'REGISTERED']
    }
    
    df = pd.DataFrame(data)
    df.to_excel('VehicleMonitoring1.xlsx', index=False, engine='openpyxl')
    print("Mock Excel file 'VehicleMonitoring1.xlsx' created successfully.")

if __name__ == "__main__":
    create_mock_excel()
