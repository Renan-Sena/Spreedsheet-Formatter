import pandas as pd
from datetime import datetime

# Criar dados de exemplo
data = {
    "Data": [datetime(2025, 8, 12), datetime(2025, 8, 13), datetime(2025, 8, 14)],
    "Empresa 1 (R$)": [150.00, 90.00, 200.00],
    "Empresa 2 (R$)": [120.00, 130.00, 180.00],
}

df = pd.DataFrame(data)

# Salvar arquivo Excel
df.to_excel("entrada.xlsx", index=False)
