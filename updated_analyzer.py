import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from collections import defaultdict
import os

# Set up plotting style
plt.rcParams['font.size'] = 10
plt.rcParams['figure.figsize'] = (15, 10)
plt.rcParams['axes.grid'] = True
plt.rcParams['grid.alpha'] = 0.3

def analyze_attendance_updated():
    """
    Analyze attendance data from the updated Excel file structure
    """
    print("=== Weekly Attendance Statistics Analysis (31 Aug - 4 Sep 2025) ===")
    print("=== Analysis Logic: 3/4 sessions = full day attendance ===")
    
    # Path to the updated Excel file
    excel_file = r'f:\work\Black Gold\attendance statistics\كشوفات الغياب الاسبوعي 31-8-2025(drive).xlsx'
    
    if not os.path.exists(excel_file):
        print(f"Error: Excel file not found at {excel_file}")
        return
    
    # Read Excel file and get all sheets
    try:
        excel_data = pd.ExcelFile(excel_file)
        sheet_names = [sheet for sheet in excel_data.sheet_names if sheet != 'الورقة1']
        
        print(f"Found {len(sheet_names)} group sheets: {sheet_names}")
        
        # Initialize statistics containers
        all_students = []
        group_stats = {}
        
        for sheet_name in sheet_names:
            print(f"\nProcessing sheet: {sheet_name}")
            
            try:
                # Read the sheet
                df = pd.read_excel(excel_file, sheet_name=sheet_name, header=None)
                
                if df.empty:
                    print(f"  - Sheet {sheet_name} is empty, skipping")
                    continue
                
                # Find where student data starts (should be around row 3, index 3)
                student_data_start = 3  # Based on diagnostic output
                
                # Extract student information and attendance data
                students_in_group = []
                
                for row_idx in range(student_data_start, len(df)):
                    # Check if we have student data in this row
                    student_number = df.iloc[row_idx, 0]  # Column A
                    student_name = df.iloc[row_idx, 1]    # Column B  
                    student_id = df.iloc[row_idx, 2]      # Column C
                    
                    # Skip if no student name or if it's not a valid row
                    if pd.isna(student_name) or not isinstance(student_name, str) or len(str(student_name).strip()) < 3:
                        # Check if we've hit the end of student data by looking ahead
                        empty_count = 0
                        for check_idx in range(row_idx, min(row_idx + 3, len(df))):
                            check_name = df.iloc[check_idx, 1]
                            if pd.isna(check_name) or not isinstance(check_name, str) or len(str(check_name).strip()) < 3:
                                empty_count += 1
                        if empty_count >= 2:  # If 2+ consecutive invalid rows, stop
                            break
                        else:
                            continue  # Skip this row but keep checking
                    
                    # Extract attendance data for all sessions across the week
                    # Based on the structure: 4 sessions per day, 5 days
                    attendance_data = []
                    
                    # Days are in columns starting from D (index 3)
                    # We have 5 days, each with 4 sessions
                    for day_offset in range(5):  # Sunday to Thursday
                        day_sessions = []
                        for session in range(4):  # 4 sessions per day
                            col_idx = 3 + (day_offset * 4) + session
                            if col_idx < len(df.columns):
                                session_value = df.iloc[row_idx, col_idx]
                                # Convert to attendance status
                                if pd.notna(session_value):
                                    if session_value == 1.0 or session_value == 1 or session_value == True:
                                        day_sessions.append(1)
                                    else:
                                        day_sessions.append(0)
                                else:
                                    day_sessions.append(0)
                            else:
                                day_sessions.append(0)
                        attendance_data.append(day_sessions)
                    
                    # Calculate daily attendance (present if attended 3/4 or 4/4 sessions)
                    daily_attendance = []
                    for day_sessions in attendance_data:
                        sessions_attended = sum(day_sessions)
                        daily_attendance.append(1 if sessions_attended >= 3 else 0)
                    
                    total_days_attended = sum(daily_attendance)
                    attendance_percentage = (total_days_attended / 5) * 100
                    
                    student_info = {
                        'group': sheet_name,
                        'student_number': student_number if pd.notna(student_number) else 'N/A',
                        'name': str(student_name).strip(),
                        'student_id': student_id if pd.notna(student_id) else 'N/A',
                        'days_attended': total_days_attended,
                        'attendance_percentage': attendance_percentage,
                        'daily_attendance': daily_attendance,
                        'session_data': attendance_data,
                        'total_sessions': sum(sum(day) for day in attendance_data),
                        'possible_sessions': 20  # 4 sessions × 5 days
                    }
                    
                    all_students.append(student_info)
                    students_in_group.append(student_info)
                
                # Calculate group statistics
                if students_in_group:
                    group_attendance_rates = [s['attendance_percentage'] for s in students_in_group]
                    full_week_students = [s for s in students_in_group if s['days_attended'] == 5]
                    partial_students = [s for s in students_in_group if 0 < s['days_attended'] < 5]
                    never_attended = [s for s in students_in_group if s['days_attended'] == 0]
                    
                    group_stats[sheet_name] = {
                        'total_students': len(students_in_group),
                        'average_attendance': sum(group_attendance_rates) / len(group_attendance_rates),
                        'full_week_count': len(full_week_students),
                        'partial_count': len(partial_students),
                        'never_attended_count': len(never_attended),
                        'students': students_in_group
                    }
                    
                    print(f"  - Students found: {len(students_in_group)}")
                    print(f"  - Full week attendance: {len(full_week_students)}")
                    print(f"  - Partial attendance: {len(partial_students)}")
                    print(f"  - Never attended: {len(never_attended)}")
                    print(f"  - Average attendance: {group_stats[sheet_name]['average_attendance']:.1f}%")
                
            except Exception as e:
                print(f"  - Error processing sheet {sheet_name}: {str(e)}")
                continue
        
        # Generate overall statistics
        print(f"\n=== Overall Statistics ===")
        print(f"Total students across all groups: {len(all_students)}")
        
        if all_students:
            overall_full_week = sum(1 for s in all_students if s['days_attended'] == 5)
            overall_partial = sum(1 for s in all_students if 0 < s['days_attended'] < 5)
            overall_never = sum(1 for s in all_students if s['days_attended'] == 0)
            overall_avg = sum(s['attendance_percentage'] for s in all_students) / len(all_students)
            
            print(f"Full week attendance (5/5 days with 3+ sessions): {overall_full_week} students ({overall_full_week/len(all_students)*100:.1f}%)")
            print(f"Partial attendance (1-4 days): {overall_partial} students ({overall_partial/len(all_students)*100:.1f}%)")
            print(f"Never attended: {overall_never} students ({overall_never/len(all_students)*100:.1f}%)")
            print(f"Overall average attendance: {overall_avg:.1f}%")
            
            # Additional session-level statistics
            total_sessions_attended = sum(s['total_sessions'] for s in all_students)
            total_possible_sessions = len(all_students) * 20
            session_attendance_rate = (total_sessions_attended / total_possible_sessions) * 100
            print(f"Overall session attendance rate: {session_attendance_rate:.1f}% ({total_sessions_attended}/{total_possible_sessions} sessions)")
            
            # Show difference between old logic (4/4) vs new logic (3/4+)
            old_logic_full_days = 0
            for student in all_students:
                for day_sessions in student['session_data']:
                    if sum(day_sessions) == 4:
                        old_logic_full_days += 1
            
            new_logic_full_days = 0
            for student in all_students:
                for day_sessions in student['session_data']:
                    if sum(day_sessions) >= 3:
                        new_logic_full_days += 1
                        
            print(f"Comparison - Old logic (4/4 only): {old_logic_full_days} full days")
            print(f"Comparison - New logic (3/4+): {new_logic_full_days} full days (+{new_logic_full_days - old_logic_full_days} days)")
            
            # Create visualizations
            create_updated_visualizations(group_stats, all_students, overall_full_week, overall_partial, overall_never)
            
            # Create detailed Excel report
            create_updated_excel_report(group_stats, all_students)
            
            # Create HTML BI Dashboard
            create_updated_html_dashboard(group_stats, all_students, overall_full_week, overall_partial, overall_never)
        
        else:
            print("No student data found in any sheets.")
        
    except Exception as e:
        print(f"Error reading Excel file: {str(e)}")

def create_updated_visualizations(group_stats, all_students, total_full_week, total_partial, total_never):
    """Create large Power BI-style visualization charts"""
    
    # Set up Power BI-inspired styling
    plt.style.use('default')
    plt.rcParams.update({
        'font.size': 14,
        'font.weight': 'bold',
        'axes.titlesize': 18,
        'axes.labelsize': 14,
        'xtick.labelsize': 12,
        'ytick.labelsize': 12,
        'legend.fontsize': 12,
        'figure.titlesize': 20,
        'axes.grid': True,
        'grid.alpha': 0.3,
        'axes.spines.top': False,
        'axes.spines.right': False,
        'axes.spines.left': True,
        'axes.spines.bottom': True,
        'axes.axisbelow': True
    })
    
    # Power BI color palette
    colors = {
        'primary': '#1f77b4',
        'success': '#2ca02c', 
        'warning': '#ff7f0e',
        'danger': '#d62728',
        'info': '#17becf',
        'secondary': '#9467bd'
    }
    
    # Create 4 separate large figures
    
    # 1. CHART 1: Students by Group (Stacked Bar Chart) - FULL SCREEN
    fig1, ax1 = plt.subplots(figsize=(20, 12))
    
    groups = list(group_stats.keys())
    full_week_counts = [group_stats[g]['full_week_count'] for g in groups]
    partial_counts = [group_stats[g]['partial_count'] for g in groups] 
    never_counts = [group_stats[g]['never_attended_count'] for g in groups]
    
    # Create stacked bars
    bar_width = 0.8
    x_pos = range(len(groups))
    
    bars1 = ax1.bar(x_pos, full_week_counts, bar_width, 
                   label='Full Week (5/5 days)', color=colors['success'], alpha=0.9)
    bars2 = ax1.bar(x_pos, partial_counts, bar_width, bottom=full_week_counts,
                   label='Partial (1-4 days)', color=colors['warning'], alpha=0.9)
    bars3 = ax1.bar(x_pos, never_counts, bar_width, 
                   bottom=[f+p for f,p in zip(full_week_counts, partial_counts)],
                   label='Never Attended', color=colors['danger'], alpha=0.9)
    
    # Styling
    ax1.set_title('📊 Weekly Attendance Distribution by Group\n(31 Aug - 4 Sep 2025)', 
                 fontsize=24, fontweight='bold', pad=30)
    ax1.set_xlabel('Groups', fontsize=16, fontweight='bold')
    ax1.set_ylabel('Number of Students', fontsize=16, fontweight='bold')
    ax1.set_xticks(x_pos)
    ax1.set_xticklabels(groups, rotation=45, ha='right')
    ax1.legend(loc='upper right', frameon=True, shadow=True)
    ax1.grid(True, alpha=0.3, axis='y')
    
    # Add value labels on bars
    for i, (full, partial, never) in enumerate(zip(full_week_counts, partial_counts, never_counts)):
        total = full + partial + never
        if full > 0:
            ax1.text(i, full/2, str(full), ha='center', va='center', fontweight='bold', color='white')
        if partial > 0:
            ax1.text(i, full + partial/2, str(partial), ha='center', va='center', fontweight='bold', color='white')
        if never > 0:
            ax1.text(i, full + partial + never/2, str(never), ha='center', va='center', fontweight='bold', color='white')
        # Total at top
        ax1.text(i, total + 1, f'Total: {total}', ha='center', va='bottom', fontweight='bold', fontsize=10)
    
    plt.tight_layout()
    plt.savefig('chart1_group_distribution.png', dpi=300, bbox_inches='tight', facecolor='white')
    plt.close()
    
    # 2. CHART 2: Overall Distribution (Large Pie Chart) - FULL SCREEN
    fig2, ax2 = plt.subplots(figsize=(16, 12))
    
    sizes = [total_full_week, total_partial, total_never]
    labels = ['Full Week\nAttendance', 'Partial\nAttendance', 'Never\nAttended']
    colors_pie = [colors['success'], colors['warning'], colors['danger']]
    explode = (0.05, 0.05, 0.1)  # explode the slices
    
    wedges, texts, autotexts = ax2.pie(sizes, labels=labels, colors=colors_pie, autopct='%1.1f%%',
                                      startangle=90, explode=explode, shadow=True, textprops={'fontsize': 14})
    
    # Enhance text
    for autotext in autotexts:
        autotext.set_color('white')
        autotext.set_fontsize(16)
        autotext.set_fontweight('bold')
    
    for text in texts:
        text.set_fontsize(14)
        text.set_fontweight('bold')
    
    ax2.set_title('🥧 Weekly Attendance Distribution\n(31 Aug - 4 Sep 2025)\nTotal Students: {:,}'.format(len(all_students)), 
                 fontsize=24, fontweight='bold', pad=30)
    
    # Add statistics box
    stats_text = f"""
    Full Week: {total_full_week} students ({total_full_week/len(all_students)*100:.1f}%)
    Partial: {total_partial} students ({total_partial/len(all_students)*100:.1f}%)
    Never: {total_never} students ({total_never/len(all_students)*100:.1f}%)
    """
    ax2.text(1.2, 0.5, stats_text, transform=ax2.transAxes, fontsize=12,
             bbox=dict(boxstyle="round,pad=0.3", facecolor='lightgray', alpha=0.8))
    
    plt.tight_layout()
    plt.savefig('chart2_overall_pie.png', dpi=300, bbox_inches='tight', facecolor='white')
    plt.close()
    
    # 3. CHART 3: Average Attendance by Group - FULL SCREEN
    fig3, ax3 = plt.subplots(figsize=(20, 12))
    
    avg_attendance = [group_stats[g]['average_attendance'] for g in groups]
    
    # Create gradient bars
    bars = ax3.bar(groups, avg_attendance, color=[colors['primary'] if x >= 70 else colors['warning'] if x >= 50 else colors['danger'] for x in avg_attendance], 
                   alpha=0.8, edgecolor='black', linewidth=1)
    
    ax3.set_title('📈 Average Weekly Attendance by Group\n(31 Aug - 4 Sep 2025)', 
                 fontsize=24, fontweight='bold', pad=30)
    ax3.set_xlabel('Groups', fontsize=16, fontweight='bold')
    ax3.set_ylabel('Average Attendance (%)', fontsize=16, fontweight='bold')
    ax3.set_ylim(0, 100)
    ax3.tick_params(axis='x', rotation=45)
    ax3.grid(True, alpha=0.3, axis='y')
    
    # Add horizontal reference lines
    ax3.axhline(y=50, color='orange', linestyle='--', alpha=0.7, linewidth=2, label='50% Threshold')
    ax3.axhline(y=70, color='green', linestyle='--', alpha=0.7, linewidth=2, label='70% Target')
    ax3.legend()
    
    # Add value labels on bars
    for bar, value in zip(bars, avg_attendance):
        height = bar.get_height()
        ax3.text(bar.get_x() + bar.get_width()/2., height + 1,
                f'{value:.1f}%', ha='center', va='bottom', fontsize=12, fontweight='bold')
    
    plt.tight_layout()
    plt.savefig('chart3_average_attendance.png', dpi=300, bbox_inches='tight', facecolor='white')
    plt.close()
    
    # 4. CHART 4: Daily Attendance Pattern - FULL SCREEN
    fig4, ax4 = plt.subplots(figsize=(16, 10))
    
    days = ['Sun (31-Aug)', 'Mon (1-Sep)', 'Tue (2-Sep)', 'Wed (3-Sep)', 'Thu (4-Sep)']
    daily_counts = []
    
    for day_idx in range(5):
        day_attendance = sum(1 for student in all_students if student['daily_attendance'][day_idx] == 1)
        daily_counts.append(day_attendance)
    
    # Create area plot
    ax4.plot(days, daily_counts, marker='o', linewidth=4, markersize=12, 
             color=colors['primary'], markerfacecolor=colors['info'], markeredgecolor='white', markeredgewidth=3)
    ax4.fill_between(days, daily_counts, alpha=0.3, color=colors['primary'])
    
    ax4.set_title('📅 Daily Attendance Pattern\nWeek of 31 Aug - 4 Sep 2025', 
                 fontsize=24, fontweight='bold', pad=30)
    ax4.set_xlabel('Days of Week', fontsize=16, fontweight='bold')
    ax4.set_ylabel('Number of Students Present', fontsize=16, fontweight='bold')
    ax4.grid(True, alpha=0.3)
    
    # Add value labels
    for i, count in enumerate(daily_counts):
        ax4.text(i, count + 5, str(count), ha='center', va='bottom', 
                fontweight='bold', fontsize=14, 
                bbox=dict(boxstyle="round,pad=0.3", facecolor='white', alpha=0.8))
    
    # Add trend line
    import numpy as np
    x_numeric = np.array(range(len(days)))
    y_numeric = np.array(daily_counts)
    z = np.polyfit(x_numeric, y_numeric, 1)
    trend_line = np.poly1d(z)(x_numeric)
    
    # Calculate R-squared
    y_mean = np.mean(y_numeric)
    ss_tot = np.sum((y_numeric - y_mean) ** 2)
    ss_res = np.sum((y_numeric - trend_line) ** 2)
    r_squared = 1 - (ss_res / ss_tot)
    
    ax4.plot(days, trend_line, '--', color=colors['danger'], linewidth=2, alpha=0.7, label=f'Trend (R²={r_squared:.3f})')
    ax4.legend()
    
    plt.tight_layout()
    plt.savefig('chart4_daily_pattern.png', dpi=300, bbox_inches='tight', facecolor='white')
    plt.close()
    
    print("\n🎨 Power BI-Style Visualizations Created (31 Aug - 4 Sep 2025):")
    print("   📊 chart1_group_distribution.png - Large stacked bar chart")
    print("   🥧 chart2_overall_pie.png - Large pie chart with stats")
    print("   📈 chart3_average_attendance.png - Large bar chart with thresholds")
    print("   📅 chart4_daily_pattern.png - Large area chart with trend")
    
def create_updated_excel_report(group_stats, all_students):
    """Create detailed Excel report with updated data"""
    
    with pd.ExcelWriter('weekly_attendance_results_31Aug-4Sep.xlsx', engine='openpyxl') as writer:
        
        # Main summary sheet
        summary_data = []
        for group, stats in group_stats.items():
            summary_data.append({
                'Group': group,
                'Total Students': stats['total_students'],
                'Full Week (5/5)': stats['full_week_count'],
                'Partial (1-4)': stats['partial_count'],
                'Never Attended': stats['never_attended_count'],
                'Average Attendance %': round(stats['average_attendance'], 1)
            })
        
        summary_df = pd.DataFrame(summary_data)
        summary_df.to_excel(writer, sheet_name='Summary', index=False)
        
        # All students details
        students_data = []
        for student in all_students:
            # Calculate session details for each day
            day_details = []
            for day_idx, day_sessions in enumerate(student['session_data']):
                sessions_count = sum(day_sessions)
                day_details.append(f"{sessions_count}/4")
            
            students_data.append({
                'Group': student['group'],
                'Student Number': student['student_number'],
                'Student Name': student['name'],
                'Student ID': student['student_id'],
                'Days Attended': student['days_attended'],
                'Attendance %': round(student['attendance_percentage'], 1),
                'Total Sessions': f"{student['total_sessions']}/20",
                'Sun (31-Aug)': '✓' if student['daily_attendance'][0] else '✗',
                'Mon (1-Sep)': '✓' if student['daily_attendance'][1] else '✗',
                'Tue (2-Sep)': '✓' if student['daily_attendance'][2] else '✗',
                'Wed (3-Sep)': '✓' if student['daily_attendance'][3] else '✗',
                'Thu (4-Sep)': '✓' if student['daily_attendance'][4] else '✗',
                'Sun Sessions': day_details[0],
                'Mon Sessions': day_details[1],
                'Tue Sessions': day_details[2],
                'Wed Sessions': day_details[3],
                'Thu Sessions': day_details[4]
            })
        
        all_students_df = pd.DataFrame(students_data)
        all_students_df.to_excel(writer, sheet_name='All Students', index=False)
        
        # Full week students (5/5 days with 3+ sessions each day)
        full_week_students = [s for s in all_students if s['days_attended'] == 5]
        if full_week_students:
            full_week_data = []
            for student in full_week_students:
                full_week_data.append({
                    'Group': student['group'],
                    'Student Number': student['student_number'],
                    'Student Name': student['name'],
                    'Student ID': student['student_id']
                })
            full_week_df = pd.DataFrame(full_week_data)
            full_week_df.to_excel(writer, sheet_name='Full Week Students', index=False)
        
        # Never attended students
        never_students = [s for s in all_students if s['days_attended'] == 0]
        if never_students:
            never_data = []
            for student in never_students:
                never_data.append({
                    'Group': student['group'],
                    'Student Number': student['student_number'],
                    'Student Name': student['name'],
                    'Student ID': student['student_id']
                })
            never_df = pd.DataFrame(never_data)
            never_df.to_excel(writer, sheet_name='Never Attended', index=False)
        
    print("Weekly Excel report saved as: weekly_attendance_results_31Aug-4Sep.xlsx")

def create_updated_html_dashboard(group_stats, all_students, full_week, partial, never):
    """Create HTML BI Dashboard with updated data"""
    
    # Prepare data for charts
    groups = list(group_stats.keys())
    full_week_counts = [group_stats[g]['full_week_count'] for g in groups]
    partial_counts = [group_stats[g]['partial_count'] for g in groups]
    never_counts = [group_stats[g]['never_attended_count'] for g in groups]
    avg_attendance = [group_stats[g]['average_attendance'] for g in groups]
    
    # Daily attendance data with specific dates
    days = ['Sun (31-Aug)', 'Mon (1-Sep)', 'Tue (2-Sep)', 'Wed (3-Sep)', 'Thu (4-Sep)']
    daily_counts = []
    for day_idx in range(5):
        day_attendance = sum(1 for student in all_students if student['daily_attendance'][day_idx] == 1)
        daily_counts.append(day_attendance)
    
    html_content = f"""
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Weekly Attendance Analytics Dashboard (31 Aug - 4 Sep 2025)</title>
        <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
        <style>
            body {{
                font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
                margin: 0;
                padding: 20px;
                background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
                min-height: 100vh;
            }}
            
            .dashboard-container {{
                max-width: 1400px;
                margin: 0 auto;
                background: white;
                border-radius: 15px;
                box-shadow: 0 10px 30px rgba(0,0,0,0.2);
                padding: 30px;
            }}
            
            .header {{
                text-align: center;
                margin-bottom: 40px;
                color: #333;
            }}
            
            .header h1 {{
                font-size: 2.5rem;
                margin: 0;
                color: #2c3e50;
            }}
            
            .header p {{
                font-size: 1.1rem;
                color: #7f8c8d;
                margin: 10px 0;
            }}
            
            .kpi-grid {{
                display: grid;
                grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
                gap: 20px;
                margin-bottom: 40px;
            }}
            
            .kpi-card {{
                background: linear-gradient(135deg, #3498db, #2980b9);
                color: white;
                padding: 25px;
                border-radius: 10px;
                text-align: center;
                box-shadow: 0 5px 15px rgba(0,0,0,0.1);
                transition: transform 0.3s ease;
            }}
            
            .kpi-card:hover {{
                transform: translateY(-5px);
            }}
            
            .kpi-card.success {{
                background: linear-gradient(135deg, #27ae60, #229954);
            }}
            
            .kpi-card.warning {{
                background: linear-gradient(135deg, #f39c12, #e67e22);
            }}
            
            .kpi-card.danger {{
                background: linear-gradient(135deg, #e74c3c, #c0392b);
            }}
            
            .kpi-number {{
                font-size: 2.5rem;
                font-weight: bold;
                margin: 0;
            }}
            
            .kpi-label {{
                font-size: 0.9rem;
                margin: 5px 0 0 0;
                opacity: 0.9;
            }}
            
            .charts-grid {{
                display: grid;
                grid-template-columns: repeat(auto-fit, minmax(400px, 1fr));
                gap: 30px;
                margin-bottom: 40px;
            }}
            
            .chart-card {{
                background: white;
                border-radius: 10px;
                padding: 25px;
                box-shadow: 0 5px 15px rgba(0,0,0,0.1);
                border: 1px solid #ecf0f1;
            }}
            
            .chart-title {{
                font-size: 1.3rem;
                font-weight: bold;
                color: #2c3e50;
                margin-bottom: 20px;
                text-align: center;
            }}
            
            .chart-container {{
                position: relative;
                height: 300px;
            }}
            
            .insights {{
                background: #f8f9fa;
                border-radius: 10px;
                padding: 25px;
                margin-top: 30px;
                border-left: 5px solid #3498db;
            }}
            
            .insights h3 {{
                color: #2c3e50;
                margin-top: 0;
            }}
            
            .insight-item {{
                margin: 10px 0;
                padding: 10px;
                background: white;
                border-radius: 5px;
                border-left: 3px solid #3498db;
            }}
        </style>
    </head>
    <body>
        <div class="dashboard-container">
            <div class="header">
                <h1>📊 Weekly Attendance Analytics Dashboard</h1>
                <h2 style="color: #3498db; margin: 10px 0;">Week of 31 August - 4 September 2025</h2>
                <p>Comprehensive analysis with updated logic: 3/4 or 4/4 sessions = Full Day Attendance</p>
                <p><strong>Total Students Analyzed:</strong> {len(all_students)} students from {len(group_stats)} groups</p>
            </div>
            
            <div class="kpi-grid">
                <div class="kpi-card success">
                    <div class="kpi-number">{full_week}</div>
                    <div class="kpi-label">Full Week Attendees</div>
                </div>
                <div class="kpi-card warning">
                    <div class="kpi-number">{partial}</div>
                    <div class="kpi-label">Partial Attendees</div>
                </div>
                <div class="kpi-card danger">
                    <div class="kpi-number">{never}</div>
                    <div class="kpi-label">Never Attended</div>
                </div>
                <div class="kpi-card">
                    <div class="kpi-number">{sum(s['attendance_percentage'] for s in all_students) / len(all_students):.1f}%</div>
                    <div class="kpi-label">Overall Avg Attendance</div>
                </div>
            </div>
            
            <div class="charts-grid">
                <div class="chart-card">
                    <div class="chart-title">📊 Weekly Attendance Distribution by Group</div>
                    <div class="chart-container">
                        <canvas id="groupChart"></canvas>
                    </div>
                </div>
                
                <div class="chart-card">
                    <div class="chart-title">🥧 Weekly Attendance Overview</div>
                    <div class="chart-container">
                        <canvas id="pieChart"></canvas>
                    </div>
                </div>
                
                <div class="chart-card">
                    <div class="chart-title">📈 Average Weekly Performance by Group</div>
                    <div class="chart-container">
                        <canvas id="avgChart"></canvas>
                    </div>
                </div>
                
                <div class="chart-card">
                    <div class="chart-title">📅 Daily Attendance Pattern (31 Aug - 4 Sep)</div>
                    <div class="chart-container">
                        <canvas id="dailyChart"></canvas>
                    </div>
                </div>
            </div>
            
            <div class="insights">
                <h3>🔍 Key Insights</h3>
                <div class="insight-item">
                    <strong>Analysis Period:</strong> Week of 31 August - 4 September 2025 (5 working days)
                </div>
                <div class="insight-item">
                    <strong>Highest Performing Group:</strong> {max(group_stats.keys(), key=lambda g: group_stats[g]['average_attendance'])} 
                    ({group_stats[max(group_stats.keys(), key=lambda g: group_stats[g]['average_attendance'])]['average_attendance']:.1f}% avg attendance)
                </div>
                <div class="insight-item">
                    <strong>Best Day:</strong> {days[daily_counts.index(max(daily_counts))]} 
                    ({max(daily_counts)} students present)
                </div>
                <div class="insight-item">
                    <strong>Weekly Completion Rate:</strong> {full_week/len(all_students)*100:.1f}% of students attended all 5 days
                </div>
                <div class="insight-item">
                    <strong>Total Data Analyzed:</strong> {len(group_stats)} groups with {len(all_students)} students
                </div>
            </div>
        </div>

        <script>
            // Group Attendance Chart (Stacked Bar)
            const groupCtx = document.getElementById('groupChart').getContext('2d');
            new Chart(groupCtx, {{
                type: 'bar',
                data: {{
                    labels: {groups},
                    datasets: [
                        {{
                            label: 'Full Week (5/5)',
                            data: {full_week_counts},
                            backgroundColor: '#27ae60',
                            borderColor: '#229954',
                            borderWidth: 1
                        }},
                        {{
                            label: 'Partial (1-4)',
                            data: {partial_counts},
                            backgroundColor: '#f39c12',
                            borderColor: '#e67e22',
                            borderWidth: 1
                        }},
                        {{
                            label: 'Never Attended',
                            data: {never_counts},
                            backgroundColor: '#e74c3c',
                            borderColor: '#c0392b',
                            borderWidth: 1
                        }}
                    ]
                }},
                options: {{
                    responsive: true,
                    maintainAspectRatio: false,
                    scales: {{
                        x: {{
                            stacked: true,
                            ticks: {{
                                maxRotation: 45
                            }}
                        }},
                        y: {{
                            stacked: true,
                            beginAtZero: true
                        }}
                    }},
                    plugins: {{
                        legend: {{
                            position: 'top',
                        }}
                    }}
                }}
            }});

            // Overall Distribution Pie Chart
            const pieCtx = document.getElementById('pieChart').getContext('2d');
            new Chart(pieCtx, {{
                type: 'pie',
                data: {{
                    labels: ['Full Week', 'Partial', 'Never Attended'],
                    datasets: [{{
                        data: [{full_week}, {partial}, {never}],
                        backgroundColor: ['#27ae60', '#f39c12', '#e74c3c'],
                        borderColor: ['#229954', '#e67e22', '#c0392b'],
                        borderWidth: 2
                    }}]
                }},
                options: {{
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {{
                        legend: {{
                            position: 'bottom'
                        }}
                    }}
                }}
            }});

            // Average Attendance Chart
            const avgCtx = document.getElementById('avgChart').getContext('2d');
            new Chart(avgCtx, {{
                type: 'bar',
                data: {{
                    labels: {groups},
                    datasets: [{{
                        label: 'Average Attendance %',
                        data: {avg_attendance},
                        backgroundColor: '#3498db',
                        borderColor: '#2980b9',
                        borderWidth: 1
                    }}]
                }},
                options: {{
                    responsive: true,
                    maintainAspectRatio: false,
                    scales: {{
                        x: {{
                            ticks: {{
                                maxRotation: 45
                            }}
                        }},
                        y: {{
                            beginAtZero: true,
                            max: 100,
                            ticks: {{
                                callback: function(value) {{
                                    return value + '%';
                                }}
                            }}
                        }}
                    }},
                    plugins: {{
                        legend: {{
                            display: false
                        }}
                    }}
                }}
            }});

            // Daily Pattern Chart
            const dailyCtx = document.getElementById('dailyChart').getContext('2d');
            new Chart(dailyCtx, {{
                type: 'line',
                data: {{
                    labels: {days},
                    datasets: [{{
                        label: 'Students Present',
                        data: {daily_counts},
                        borderColor: '#9b59b6',
                        backgroundColor: 'rgba(155, 89, 182, 0.2)',
                        borderWidth: 3,
                        fill: true,
                        tension: 0.4,
                        pointRadius: 6,
                        pointHoverRadius: 8
                    }}]
                }},
                options: {{
                    responsive: true,
                    maintainAspectRatio: false,
                    scales: {{
                        y: {{
                            beginAtZero: true
                        }}
                    }},
                    plugins: {{
                        legend: {{
                            display: false
                        }}
                    }}
                }}
            }});
        </script>
    </body>
    </html>
    """
    
    with open('weekly_attendance_dashboard_31Aug-4Sep.html', 'w', encoding='utf-8') as f:
        f.write(html_content)
    
    print("Weekly HTML BI Dashboard saved as: weekly_attendance_dashboard_31Aug-4Sep.html")

if __name__ == "__main__":
    analyze_attendance_updated()
