# Identified Mistakes in Generated Report

## Issue 1: Slide 8 - Entity Name Display
**Template**: Shows "HIAL- MRO" as a combined entity in row 13
**Generated**: Shows "HIAL" and "MRO" as separate entities in rows 6-10

**Template Table Row 13**:
```
['4', 'HIAL- MRO', '22', '19', '03', 'Critical', '02', '19', '00', ...]
```

**Generated Table Rows 6-10**:
```
Row 6: ['6', 'HIAL', '4', '3', '1', 'Critical', '2', ...]
Row 7: ['7', '', '', '', '', 'High', '2', ...]
Row 8: ['8', 'MRO', '22', '19', '3', 'Critical', '2', ...]
```

**Root Cause**: The template appears to merge HIAL and MRO into one row with combined name "HIAL- MRO" but MRO's data. Our script treats them as separate entities.

## Issue 2: Need to verify exact template row structure

Looking at template Slide 8 rows:
- Row 1: DIAL with all its severity levels
- Row 5: ENT (but shows as 'ENT ') with all its severity levels  
- Row 13: "HIAL- MRO" showing MRO data (22 total, Critical 2, etc.)

This suggests the template may be combining entities or has a specific display format.

## Issue 3: Verify if there are other formatting differences

Need to check:
1. Exact row count in template Slide 8
2. How entities are grouped
3. Whether "HIAL- MRO" is intentional or a template-specific quirk
4. Column header exact text match
5. Check if FP/TP columns have actual data in template

## Next Steps

1. Extract exact template Slide 8 table data
2. Compare row-by-row with generated
3. Identify all discrepancies
4. Update script to match exact format
