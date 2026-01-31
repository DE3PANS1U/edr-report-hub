"""
Detailed analysis of combined.xlsx structure
"""
import pandas as pd

df = pd.read_excel('combined.xlsx')

print("=" * 60)
print("DETAILED DATA STRUCTURE ANALYSIS")
print("=" * 60)

print("\nComplete Data:")
print(df.to_string())

print("\n" + "=" * 60)
print("HIERARCHICAL STRUCTURE DETECTED")
print("=" * 60)

# The data appears to be hierarchical:
# Entity -> Status -> Severity -> Count

current_entity = None
current_status = None

for idx, row in df.iterrows():
    label = row['Row Labels']
    count = row['Count of Alert Severity']
    
    # Check if it's an entity (DIAL, ENT, HIAL, MRO)
    if label in ['DIAL', 'ENT', 'HIAL', 'MRO']:
        current_entity = label
        print(f"\nEntity: {label} -> Total: {count}")
    
    # Check if it's a status
    elif label in ['closed', 'WIP']:
        current_status = label
        print(f"  Status: {label} -> Count: {count}")
    
    # Check if it's a severity
    elif label in ['Critical', 'High', 'Medium', 'Low']:
        print(f"    Severity: {label} -> Count: {count}")
    
    # Grand Total
    elif label == 'Grand Total':
        print(f"\nGrand Total: {count}")

print("\n" + "=" * 60)
print("EXTRACTING STRUCTURED DATA")
print("=" * 60)

# Parse the hierarchical structure
entities_data = {}
current_entity = None
current_status = None

for idx, row in df.iterrows():
    label = row['Row Labels']
    count = row['Count of Alert Severity']
    
    if label in ['DIAL', 'ENT', 'HIAL', 'MRO']:
        current_entity = label
        entities_data[current_entity] = {
            'total': count,
            'closed': {},
            'WIP': {},
            'all_severities': {}
        }
    elif label in ['closed', 'WIP'] and current_entity:
        current_status = label
    elif label in ['Critical', 'High', 'Medium', 'Low'] and current_entity and current_status:
        entities_data[current_entity][current_status][label] = count
        # Also track overall severity counts
        if label not in entities_data[current_entity]['all_severities']:
            entities_data[current_entity]['all_severities'][label] = 0
        entities_data[current_entity]['all_severities'][label] += count

print("\nParsed Data Structure:")
import json
print(json.dumps(entities_data, indent=2))
