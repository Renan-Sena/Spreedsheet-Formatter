import pandas as pd
from datetime import datetime

# Create sample data
data = {
    "Date": [datetime(2025, 8, 12), datetime(2025, 8, 13), datetime(2025, 8, 14)],
    "Company 1 (US$)": [150.00, 90.00, 200.00],
    "Company 2 (US$)": [120.00, 130.00, 180.00],
}

df = pd.DataFrame(data)

# Save Excel file
df.to_excel("entrada.xlsx", index=False)
