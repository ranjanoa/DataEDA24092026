import pandas as pd
import numpy as np

corr_df = pd.read_csv('correlation_matrix YANBU.csv', index_col=0)

pairs_to_check = [
    ('Summation_of_Kiln_Feed_1&2', 'Oil_Flow'),
    ('Summation_of_Kiln_Feed_1&2', 'Kiln_Drives,Current_Average'),
    ('Oil_Flow', 'Kiln_Drives,Current_Average'),
    ('Oil_Flow', 'Kiln_Inlet_String_1,_Temp'),
    ('Oil_Flow', 'Kiln_IN_Gas_Ana,_O2_%'),
    ('Kiln_Drives,Current_Average', 'Kiln_Inlet_String_1,_Temp'),
    ('Oil_Flow', 'Gas_Anal_Bhind_PH1,_CO-PPM'),
    ('Summation_of_Kiln_Feed_1&2', 'Kiln_IN_Gas_Ana,_O2_%'),
    ('Secondary_Air_Temp', 'Kiln_Drives,Current_Average')
]

print('--- SELECTED CORRELATIONS ---')
for p1, p2 in pairs_to_check:
    try:
        val = corr_df.loc[p1, p2]
        print(f"{p1} vs {p2}: {val:.3f}")
    except KeyError as e:
        print(f"Missing key: {e}")
