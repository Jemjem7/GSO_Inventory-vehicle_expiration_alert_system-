import pandas as pd
from datetime import datetime, timedelta

# Create some mock business permits expiring soon
today = datetime.now()

data = {
    'Permit ID': ['BP-2024-001', 'BP-2024-002', 'BP-2024-003', 'BP-2024-004'],
    'Business Name': ['Acme Corp', 'Stark Industries', 'Wayne Enterprises', 'Daily Planet'],
    'Department': ['Zoning', 'Fire Safety', 'Health Inspection', 'Business LGU'],
    'Valid Until Date': [
        (today - timedelta(days=2)).strftime('%Y-%m-%d'), # Exists (Negative) -> EXPIRED
        (today + timedelta(days=5)).strftime('%Y-%m-%d'), # Exists (0-7 days) -> 1 WEEK
        (today + timedelta(days=15)).strftime('%Y-%m-%d'), # Exists (8-30) -> 1 MONTH
        (today + timedelta(days=90)).strftime('%Y-%m-%d')  # Exists (>60) -> SUFFICIENT
    ]
}

df = pd.DataFrame(data)

# Create an excel file with some empty rows at the top to simulate real user formats
with pd.ExcelWriter('Business Permits.xlsx', engine='openpyxl') as writer:
    df.to_excel(writer, sheet_name='Q1 2024 Permits', index=False, startrow=2)

print("Mock Business Permits.xlsx generated.")
