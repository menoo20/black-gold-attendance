# Multi-Week Attendance Analysis System Usage Guide

## Quick Start

### 1. Analyzing Your First Week

```python
from multi_week_analyzer import MultiWeekAttendanceAnalyzer

# Initialize the analyzer
analyzer = MultiWeekAttendanceAnalyzer()

# Add your week
week_info = analyzer.add_week(
    week_id="week_7Sep-11Sep",          # Unique ID for this week
    start_date="7-Sep",                  # Start date
    end_date="11-Sep",                   # End date
    excel_file_path="path/to/your/attendance_sheet.xlsx",
    description="Second week of September 2025"
)

# Analyze the week
summary = analyzer.analyze_week("week_7Sep-11Sep", week_info["excel_file"])

# Create master dashboard
analyzer.create_master_dashboard()
```

### 2. Adding Multiple Weeks

```python
# Initialize analyzer
analyzer = MultiWeekAttendanceAnalyzer()

# Add multiple weeks
weeks_to_analyze = [
    {
        "week_id": "week_14Sep-18Sep",
        "start_date": "14-Sep",
        "end_date": "18-Sep",
        "excel_file": "attendance_week3.xlsx",
        "description": "Third week of September"
    },
    {
        "week_id": "week_21Sep-25Sep", 
        "start_date": "21-Sep",
        "end_date": "25-Sep",
        "excel_file": "attendance_week4.xlsx",
        "description": "Fourth week of September"
    }
]

# Process each week
for week_data in weeks_to_analyze:
    # Add week
    week_info = analyzer.add_week(**week_data)
    
    # Analyze
    summary = analyzer.analyze_week(week_data["week_id"], week_data["excel_file"])
    
    print(f"Week {week_data['week_id']} completed!")

# Update master dashboard
analyzer.create_master_dashboard()

# Save weeks index for persistence
analyzer.save_weeks_index()
```

## Understanding the System

### Directory Structure
```
attendance statistics/
├── master_dashboard.html          # Main selection interface
├── multi_week_analyzer.py         # The analyzer script
├── weeks_index.json               # Weeks metadata
└── weeks/                         # All weeks data
    ├── week_31Aug-4Sep/          # Individual week folder
    │   ├── dashboard_week_31Aug-4Sep.html
    │   ├── attendance_report_week_31Aug-4Sep.xlsx
    │   ├── group_distribution.png
    │   ├── overall_distribution.png
    │   └── data_week_31Aug-4Sep.json
    ├── week_7Sep-11Sep/          # Next week folder
    └── week_14Sep-18Sep/         # Another week folder
```

### What Gets Generated for Each Week

1. **Individual HTML Dashboard**: Interactive charts and statistics
2. **Excel Report**: Detailed attendance data with group breakdowns
3. **PNG Charts**: Distribution visualizations
4. **JSON Data**: Raw analysis results for data integration

### Master Dashboard Features

- **Week Selection**: Click any week card to view detailed analysis
- **Quick Stats**: Total students and attendance rate preview
- **Responsive Design**: Works on desktop and mobile
- **Professional Styling**: Power BI-inspired interface

## Advanced Usage

### Loading Existing Weeks Data

```python
# Load previously analyzed weeks
analyzer = MultiWeekAttendanceAnalyzer()
analyzer.load_weeks_index()  # Loads from weeks_index.json

# Now you can access all previously analyzed weeks
print("Available weeks:", list(analyzer.weeks_data.keys()))
```

### Custom Week Analysis

```python
# Analyze specific groups only
analyzer = MultiWeekAttendanceAnalyzer()

# Add week with specific groups to focus on
week_info = analyzer.add_week(
    week_id="week_special_analysis",
    start_date="1-Oct", 
    end_date="5-Oct",
    excel_file_path="october_week1.xlsx",
    description="October focus groups analysis"
)

# Analyze normally - the system handles group detection automatically
summary = analyzer.analyze_week("week_special_analysis", week_info["excel_file"])
```

## File Naming Conventions

### Recommended Week ID Format
- `week_DDMon-DDMon` (e.g., `week_31Aug-4Sep`)
- Use 3-letter month abbreviations: Jan, Feb, Mar, Apr, May, Jun, Jul, Aug, Sep, Oct, Nov, Dec

### Excel File Requirements
- Must follow the same structure as your original file
- Student names in Column B (index 1)
- Attendance data starting from Column D
- 4 sessions per day × 5 days = 20 attendance columns

## Troubleshooting

### Common Issues

1. **Excel File Not Found**
   - Ensure the file path is correct and accessible
   - Use absolute paths for reliability

2. **Empty Week Analysis**
   - Check that your Excel file has the correct sheet structure
   - Verify student names are in Column B

3. **Dashboard Not Loading Charts**
   - Ensure internet connection (uses CDN for Chart.js)
   - Check browser console for JavaScript errors

### Performance Tips

- Process weeks one at a time for large datasets
- Keep Excel files organized in a dedicated folder
- Regular cleanup of old analysis files if not needed

## Integration with Git

The system works well with version control:

```bash
# Add new week analysis
git add weeks/week_new/
git commit -m "Add analysis for week_new"

# Update master dashboard
git add master_dashboard.html weeks_index.json
git commit -m "Update master dashboard with new week"
```

## System Requirements

- Python 3.7+
- Required packages: pandas, matplotlib, seaborn, openpyxl, numpy
- Modern web browser for viewing dashboards
- Git (optional, for version control)
