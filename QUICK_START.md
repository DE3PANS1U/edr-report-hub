# Quick Usage Guide - SOC Report Generator

## Installation (One-time setup)

```bash
pip install pandas python-pptx openpyxl
```

## Usage

### Option 1: Interactive Mode (Easiest)
```bash
python soc_report_generator.py
```
Then enter your Excel file path when prompted.

### Option 2: Python API
```python
from soc_report_generator import generate_soc_report

# Generate report from your Excel file
output = generate_soc_report('your_edr_data.xlsx')
print(f"Report saved as: {output}")
```

### Option 3: Test with Sample Data
```bash
# Step 1: Generate sample data
python generate_sample_data.py

# Step 2: Generate report
python soc_report_generator.py
# Enter: sample_edr_alerts.xlsx
```

## Required Excel Columns

| Column Name | Valid Values | Example |
|------------|--------------|---------|
| Entity | Any text | Workstation-A |
| Alert Severity | Critical, High, Medium, Low | High |
| Alert Status | closed, WIP | closed |
| Alert Efficiency | True Positive, False Positive | True Positive |
| Alert Trend | Date/DateTime | 2026-01-25 14:30:00 |

## Output

File: `EDR-Weekly-Incident-Report_YYYY-MM-DD.pptx`

**Contents:**
1. Title Slide
2. Alert Severity Distribution (Chart + Table)
3. Alert Status Overview (Chart + Table)
4. Alert Efficiency Analysis (Chart + Table)
5. Alert Trend Analysis (Line Chart - Editable)
6. Weekly Incident Summary (Detailed Table)

## Customization

To change colors, edit `COLORS` in `soc_report_generator.py`:
```python
COLORS = {
    'critical': RGBColor(220, 53, 69),  # Red
    'high': RGBColor(255, 133, 27),     # Orange
    # ... modify as needed
}
```

## Support

See [README.md](README.md) for detailed documentation.
