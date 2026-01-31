"""
Generate Sample EDR Alert Data for Testing SOC Report Generator
"""

import pandas as pd
from datetime import datetime, timedelta
import random

# Sample entities
entities = ['Workstation-A', 'Server-B', 'Workstation-C', 'Server-D', 'Endpoint-E']

# Alert categories
severities = ['Critical', 'High', 'Medium', 'Low']
statuses = ['closed', 'WIP']
efficiencies = ['True Positive', 'False Positive']

# Generate sample data
data = []

# Generate alerts over the past week
start_date = datetime.now() - timedelta(days=7)

for i in range(150):  # Generate 150 sample alerts
    entity = random.choice(entities)
    severity = random.choices(
        severities,
        weights=[10, 20, 40, 30],  # More medium/low alerts
        k=1
    )[0]
    
    # Critical and High alerts more likely to be WIP
    if severity in ['Critical', 'High']:
        status = random.choices(statuses, weights=[60, 40], k=1)[0]
    else:
        status = random.choices(statuses, weights=[85, 15], k=1)[0]
    
    # Critical alerts more likely to be True Positive
    if severity == 'Critical':
        efficiency = random.choices(efficiencies, weights=[80, 20], k=1)[0]
    elif severity == 'High':
        efficiency = random.choices(efficiencies, weights=[65, 35], k=1)[0]
    else:
        efficiency = random.choices(efficiencies, weights=[40, 60], k=1)[0]
    
    # Random date within the week
    alert_date = start_date + timedelta(
        days=random.randint(0, 6),
        hours=random.randint(0, 23),
        minutes=random.randint(0, 59)
    )
    
    data.append({
        'Entity': entity,
        'Alert Severity': severity,
        'Alert Status': status,
        'Alert Efficiency': efficiency,
        'Alert Trend': alert_date
    })

# Create DataFrame
df = pd.DataFrame(data)

# Sort by date
df = df.sort_values('Alert Trend').reset_index(drop=True)

# Save to Excel
output_file = 'sample_edr_alerts.xlsx'
df.to_excel(output_file, index=False)

print(f"Sample data generated: {output_file}")
print(f"Total alerts: {len(df)}")
print(f"\nSeverity distribution:")
print(df['Alert Severity'].value_counts())
print(f"\nStatus distribution:")
print(df['Alert Status'].value_counts())
print(f"\nEfficiency distribution:")
print(df['Alert Efficiency'].value_counts())
