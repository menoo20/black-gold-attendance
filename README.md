# Weekly Attendance Statistics Analysis

## Overview
This project analyzes weekly attendance data for multiple student groups for the week of **31 August - 4 September 2025**.

## Key Features
- **Smart Attendance Logic**: 3/4 or 4/4 sessions per day = full day attendance
- **Multi-Group Analysis**: Processes 20 different groups with 439 total students
- **Professional Visualizations**: Power BI-style charts and dashboards
- **Multiple Output Formats**: PNG charts, Excel reports, and interactive HTML dashboard

## Analysis Results (31 Aug - 4 Sep 2025)
- **Total Students**: 439 across 20 groups
- **Full Week Attendance**: 95 students (21.6%)
- **Partial Attendance**: 212 students (48.3%)
- **Never Attended**: 132 students (30.1%)
- **Overall Average**: 48.6% attendance rate

## Generated Files
### Visualizations
- `chart1_group_distribution.png` - Group attendance distribution
- `chart2_overall_pie.png` - Overall attendance breakdown
- `chart3_average_attendance.png` - Average performance by group
- `chart4_daily_pattern.png` - Daily attendance patterns

### Reports
- `weekly_attendance_results_31Aug-4Sep.xlsx` - Detailed Excel report
- `weekly_attendance_dashboard_31Aug-4Sep.html` - Interactive web dashboard

### Source Code
- `updated_analyzer.py` - Main analysis script

## Usage
```bash
python updated_analyzer.py
```

## Requirements
- Python 3.8+
- pandas
- matplotlib
- seaborn
- openpyxl
- numpy

## Data Source
- Excel file: `كشوفات الغياب الاسبوعي 31-8-2025(drive).xlsx`
- Contains attendance data with checkmarks/numeric values
- Structure: Student info in columns A-C, attendance data in subsequent columns

## Analysis Logic
- **Daily Attendance**: Student considered present if attended 3+ out of 4 sessions
- **Weekly Analysis**: 5 working days (Sunday to Thursday)
- **Categories**: Full week (5/5 days), Partial (1-4 days), Never attended (0 days)
