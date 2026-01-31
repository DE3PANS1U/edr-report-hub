# SOC Weekly Report Generator

A professional Python automation tool for generating SOC (Security Operations Center) weekly incident reports from EDR alert data.

## Features

✅ **Multiple Chart Types**
- Clustered bar charts for severity and status distribution
- Line charts with markers for trend analysis
- Professional summary tables

✅ **Comprehensive Analysis Slides**
1. **Alert Severity Distribution** - Entity-wise breakdown with color-coded severity levels
2. **Alert Status Overview** - Closed vs WIP analysis
3. **Alert Efficiency Analysis** - True Positive vs False Positive with accuracy metrics
4. **Alert Trend Analysis** - Daily trends with editable charts
5. **Weekly Incident Summary** - Detailed entity and severity breakdown

✅ **Professional Design**
- Corporate SOC color theme
- Clean, readable layouts
- Professional fonts and formatting
- All charts are fully editable in PowerPoint

✅ **Automated Workflow**
- Reads Excel files with EDR alert data
- Auto-generates date-stamped reports
- Color-coded severity levels (Critical=Red, High=Orange, Medium=Amber, Low=Green)

## Installation

1. Install required dependencies:
```bash
pip install -r requirements.txt
```

## Required Excel Format

Your Excel file must contain the following columns:

| Column Name | Description | Valid Values |
|------------|-------------|--------------|
| Entity | System/endpoint name | Any text |
| Alert Severity | Severity level | Critical, High, Medium, Low |
| Alert Status | Current status | closed, WIP |
| Alert Efficiency | Classification | True Positive, False Positive |
| Alert Trend | Alert timestamp | Date/datetime |

## Usage

### Method 1: Interactive Mode

Run the script and enter the Excel file path when prompted:

```bash
python soc_report_generator.py
```

### Method 2: Programmatic Usage

```python
from soc_report_generator import generate_soc_report

# Generate report
output_file = generate_soc_report('path/to/your/edr_alerts.xlsx')
print(f"Report generated: {output_file}")
```

### Method 3: Custom Output Directory

```python
from soc_report_generator import generate_soc_report

output_file = generate_soc_report(
    'path/to/edr_alerts.xlsx',
    output_dir='C:/Reports'
)
```

## Testing with Sample Data

Generate sample EDR alert data for testing:

```bash
python generate_sample_data.py
```

This creates `sample_edr_alerts.xlsx` with 150 sample alerts. Then generate the report:

```bash
python soc_report_generator.py
```

Enter `sample_edr_alerts.xlsx` when prompted.

## Output

The script generates a PowerPoint file named:
```
EDR-Weekly-Incident-Report_YYYY-MM-DD.pptx
```

## Color Scheme

The report uses a professional SOC color theme:

- **Critical**: Red (#DC3545)
- **High**: Orange (#FF851B)
- **Medium**: Amber (#FFC107)
- **Low**: Green (#28A745)
- **Closed**: Green (#28A745)
- **WIP**: Amber (#FFC107)
- **True Positive**: Red (#DC3545)
- **False Positive**: Green (#28A745)
- **Accent**: Blue (#007BFF)

## Slide Details

### Slide 1: Alert Severity Distribution
- Clustered bar chart showing severity distribution per entity
- Summary table with counts per severity level
- Total alerts count

### Slide 2: Alert Status Overview
- Bar chart comparing Closed vs WIP alerts per entity
- Status summary table

### Slide 3: Alert Efficiency Analysis
- True Positive vs False Positive comparison
- Accuracy percentage calculation
- Efficiency summary table

### Slide 4: Alert Trend Analysis (Editable)
- Line chart with markers showing daily trends
- Entity-wise trend lines
- Fully editable in PowerPoint

### Slide 5: Weekly Incident Summary
- Comprehensive table with:
  - Entity and severity breakdown
  - Incident counts
  - TP/FP counts
  - Category classification
  - Status remarks

## Customization

### Modify Colors

Edit the `COLORS` dictionary in `soc_report_generator.py`:

```python
COLORS = {
    'critical': RGBColor(220, 53, 69),
    'high': RGBColor(255, 133, 27),
    # ... add your custom colors
}
```

### Adjust Chart Sizes

Modify the `Inches()` parameters in the chart creation functions:

```python
create_clustered_bar_chart(
    slide, chart_data, title,
    Inches(0.5),  # left position
    Inches(1.2),  # top position
    Inches(9),    # width
    Inches(3.5),  # height
    colors
)
```

### Change Font Sizes

Update the `Pt()` values throughout the script:

```python
title_frame.paragraphs[0].font.size = Pt(32)  # Title size
chart.legend.font.size = Pt(11)                # Legend size
cell.text_frame.paragraphs[0].font.size = Pt(11)  # Table cell size
```

## Troubleshooting

**Issue**: Missing columns error
- **Solution**: Ensure your Excel file has all required columns: Entity, Alert Severity, Alert Status, Alert Efficiency, Alert Trend

**Issue**: Date parsing errors
- **Solution**: Verify that the Alert Trend column contains valid dates

**Issue**: Charts not appearing
- **Solution**: Ensure pandas and python-pptx are properly installed

**Issue**: Import errors
- **Solution**: Run `pip install -r requirements.txt`

## Requirements

- Python 3.7+
- pandas >= 2.0.0
- python-pptx >= 0.6.21
- openpyxl >= 3.1.0

## License

This is a professional automation tool for SOC reporting.

## Author

Senior Python Automation Engineer
SOC Reporting & Executive Dashboards Specialist
