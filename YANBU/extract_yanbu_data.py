import pandas as pd
import numpy as np

# 1. Read Ranges
ranges_df = pd.read_excel('YANBU RANGES.xlsx')
print('--- RANGES ---')
print(ranges_df.to_string())

# 2. Read Correlations
corr_df = pd.read_csv('correlation_matrix YANBU.csv', index_col=0)
# Find top correlations (absolute value > 0.8, excluding 1.0)
corrs = []
for col in corr_df.columns:
    for row in corr_df.index:
        if col != row:
            val = corr_df.loc[row, col]
            if not np.isnan(val) and abs(val) > 0.75:
                corrs.append((col, row, val))

# Remove duplicates
unique_corrs = []
seen = set()
for c1, c2, val in corrs:
    pair = tuple(sorted([c1, c2]))
    if pair not in seen:
        seen.add(pair)
        unique_corrs.append((pair[0], pair[1], val))

unique_corrs.sort(key=lambda x: abs(x[2]), reverse=True)

print('\n--- TOP CORRELATIONS ---')
for c1, c2, val in unique_corrs[:30]:
    print(f"{c1} vs {c2}: {val:.3f}")
