import pandas as pd
from ksp import calculate_total_charts, fill_mission_payload_column, extract_program_name

df = pd.read_csv('Career-1-Brownsville - Sheet1.csv')
df = df.dropna(subset=['KSP Date'])
df = df[df['KSP Date'].str.strip() != '']
df = fill_mission_payload_column(df)

total = calculate_total_charts(df)
print(f'Calculated total: {total}')

# Check program count
df['Program'] = df['Mission / Payload'].apply(extract_program_name)
programs = df['Program'].unique()
programs = [p for p in programs if p != 'Other' and p.strip()]
print(f'Programs found: {len(programs)}')
print(f'Programs: {programs}')

# Check the breakdown
print('Base charts: 22')
print(f'Program analysis charts: {len(programs)} * 6 = {len(programs) * 6}')
print(f'Total should be: {22 + len(programs) * 6}')
