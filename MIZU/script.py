import pandas as pd
import sys
sys.path.append('backend')
from data_processor import load_data, calculate_stats

with open('MIZU/MIZU2026_EDA to process.xlsx', 'rb') as f:
    df = load_data(f.read(), 'MIZU2026_EDA to process.xlsx')

print('--- NEW STATS ---')
stats = calculate_stats(df)
print(pd.DataFrame(stats['stats']).to_string())

print('--- CORR ---')
print(df.corr().to_string())
