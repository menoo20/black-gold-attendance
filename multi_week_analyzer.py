import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from collections import defaultdict
import os
import json
from datetime import datetime, timedelta
import numpy as np

class MultiWeekAttendanceAnalyzer:
    """
    Multi-week attendance analysis system that organizes data by weeks
    and provides a unified interface for week selection and analysis
    """
    
    def __init__(self, base_dir="weeks"):
        self.base_dir = base_dir
        self.weeks_data = {}
        
        # Ensure weeks directory exists
        if not os.path.exists(self.base_dir):
            os.makedirs(self.base_dir)
    
    def add_week(self, week_id, start_date, end_date, excel_file_path, description=""):
        """
        Add a new week to the analysis system
        
        Args:
            week_id (str): Unique identifier for the week (e.g., "week_31Aug-4Sep")
            start_date (str): Start date in format "DD-Mon" (e.g., "31-Aug")
            end_date (str): End date in format "DD-Mon" (e.g., "4-Sep")
            excel_file_path (str): Path to the Excel file for this week
            description (str): Optional description
        """
        week_info = {
            "week_id": week_id,
            "start_date": start_date,
            "end_date": end_date,
            "excel_file": excel_file_path,
            "description": description,
            "year": 2025,  # Default year
            "analysis_date": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }
        
        # Create week directory
        week_dir = os.path.join(self.base_dir, week_id)
        if not os.path.exists(week_dir):
            os.makedirs(week_dir)
        
        week_info["directory"] = week_dir
        self.weeks_data[week_id] = week_info
        
        return week_info
    
    def analyze_week(self, week_id, excel_file_path):
        """
        Analyze attendance for a specific week
        """
        if week_id not in self.weeks_data:
            raise ValueError(f"Week {week_id} not found. Please add it first.")
        
        week_info = self.weeks_data[week_id]
        week_dir = week_info["directory"]
        
        print(f"\\n=== Analyzing Week: {week_id} ({week_info['start_date']} - {week_info['end_date']} 2025) ===")
        
        # Initialize statistics containers
        all_students = []
        group_stats = {}
        
        try:
            # Read Excel file and get all sheets
            excel_data = pd.ExcelFile(excel_file_path)
            sheet_names = [sheet for sheet in excel_data.sheet_names if sheet != 'الورقة1']
            
            print(f"Found {len(sheet_names)} group sheets")
            
            for sheet_name in sheet_names:
                print(f"Processing sheet: {sheet_name}")
                
                try:
                    # Read the sheet
                    df = pd.read_excel(excel_file_path, sheet_name=sheet_name, header=None)
                    
                    if df.empty:
                        continue
                    
                    # Find where student data starts
                    student_data_start = 3
                    students_in_group = []
                    
                    for row_idx in range(student_data_start, len(df)):
                        student_number = df.iloc[row_idx, 0]  # Column A
                        student_name = df.iloc[row_idx, 1]    # Column B  
                        student_id = df.iloc[row_idx, 2]      # Column C
                        
                        # Skip if no student name
                        if pd.isna(student_name) or not isinstance(student_name, str) or len(str(student_name).strip()) < 3:
                            # Check if we've hit the end of student data
                            empty_count = 0
                            for check_idx in range(row_idx, min(row_idx + 3, len(df))):
                                check_name = df.iloc[check_idx, 1]
                                if pd.isna(check_name) or not isinstance(check_name, str) or len(str(check_name).strip()) < 3:
                                    empty_count += 1
                            if empty_count >= 2:
                                break
                            else:
                                continue
                        
                        # Extract attendance data for all sessions across the week
                        attendance_data = []
                        
                        # Days are in columns starting from D (index 3)
                        # 5 days, each with 4 sessions
                        for day_offset in range(5):
                            day_sessions = []
                            for session in range(4):
                                col_idx = 3 + (day_offset * 4) + session
                                if col_idx < len(df.columns):
                                    session_value = df.iloc[row_idx, col_idx]
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
                            'possible_sessions': 20
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
                        print(f"  - Full week: {len(full_week_students)}, Partial: {len(partial_students)}, Never: {len(never_attended)}")
                
                except Exception as e:
                    print(f"  - Error processing sheet {sheet_name}: {str(e)}")
                    continue
            
            # Generate overall statistics
            if all_students:
                overall_full_week = sum(1 for s in all_students if s['days_attended'] == 5)
                overall_partial = sum(1 for s in all_students if 0 < s['days_attended'] < 5)
                overall_never = sum(1 for s in all_students if s['days_attended'] == 0)
                overall_avg = sum(s['attendance_percentage'] for s in all_students) / len(all_students)
                
                week_summary = {
                    'total_students': len(all_students),
                    'full_week': overall_full_week,
                    'partial': overall_partial,
                    'never': overall_never,
                    'average_attendance': overall_avg,
                    'groups': len(group_stats)
                }
                
                # Save week summary
                self.weeks_data[week_id]['summary'] = week_summary
                
                print(f"\\n=== Week Summary ===")
                print(f"Total students: {len(all_students)}")
                print(f"Full week attendance: {overall_full_week} ({overall_full_week/len(all_students)*100:.1f}%)")
                print(f"Partial attendance: {overall_partial} ({overall_partial/len(all_students)*100:.1f}%)")
                print(f"Never attended: {overall_never} ({overall_never/len(all_students)*100:.1f}%)")
                print(f"Overall average: {overall_avg:.1f}%")
                
                # Create visualizations for this week
                self.create_week_visualizations(week_id, group_stats, all_students, overall_full_week, overall_partial, overall_never)
                
                # Create Excel report for this week
                self.create_week_excel_report(week_id, group_stats, all_students)
                
                # Create individual HTML dashboard for this week
                self.create_week_html_dashboard(week_id, group_stats, all_students, overall_full_week, overall_partial, overall_never)
                
                # Save analysis data
                self.save_week_data(week_id, group_stats, all_students)
                
                return week_summary
            
        except Exception as e:
            print(f"Error analyzing week {week_id}: {str(e)}")
            return None
    
    def create_week_visualizations(self, week_id, group_stats, all_students, full_week, partial, never):
        """Create visualizations for a specific week"""
        week_info = self.weeks_data[week_id]
        week_dir = week_info["directory"]
        
        # Set up styling
        plt.style.use('default')
        plt.rcParams.update({
            'font.size': 12,
            'font.weight': 'bold',
            'axes.titlesize': 16,
            'figure.titlesize': 18,
        })
        
        colors = {
            'success': '#2ca02c',
            'warning': '#ff7f0e', 
            'danger': '#d62728',
            'primary': '#1f77b4'
        }
        
        # 1. Group Distribution Chart
        fig, ax = plt.subplots(figsize=(14, 8))
        groups = list(group_stats.keys())
        full_counts = [group_stats[g]['full_week_count'] for g in groups]
        partial_counts = [group_stats[g]['partial_count'] for g in groups]
        never_counts = [group_stats[g]['never_attended_count'] for g in groups]
        
        x_pos = range(len(groups))
        ax.bar(x_pos, full_counts, label='Full Week', color=colors['success'], alpha=0.8)
        ax.bar(x_pos, partial_counts, bottom=full_counts, label='Partial', color=colors['warning'], alpha=0.8)
        ax.bar(x_pos, never_counts, bottom=[f+p for f,p in zip(full_counts, partial_counts)], 
               label='Never', color=colors['danger'], alpha=0.8)
        
        ax.set_title(f'Attendance Distribution by Group\\n{week_info["start_date"]} - {week_info["end_date"]} 2025')
        ax.set_xlabel('Groups')
        ax.set_ylabel('Number of Students')
        ax.set_xticks(x_pos)
        ax.set_xticklabels(groups, rotation=45, ha='right')
        ax.legend()
        ax.grid(True, alpha=0.3)
        
        plt.tight_layout()
        plt.savefig(os.path.join(week_dir, 'group_distribution.png'), dpi=300, bbox_inches='tight')
        plt.close()
        
        # 2. Overall Pie Chart
        fig, ax = plt.subplots(figsize=(10, 8))
        sizes = [full_week, partial, never]
        labels = ['Full Week', 'Partial', 'Never']
        colors_pie = [colors['success'], colors['warning'], colors['danger']]
        
        ax.pie(sizes, labels=labels, colors=colors_pie, autopct='%1.1f%%', startangle=90)
        ax.set_title(f'Overall Attendance Distribution\\n{week_info["start_date"]} - {week_info["end_date"]} 2025')
        
        plt.tight_layout()
        plt.savefig(os.path.join(week_dir, 'overall_distribution.png'), dpi=300, bbox_inches='tight')
        plt.close()
        
        print(f"Visualizations saved for {week_id}")
    
    def create_week_excel_report(self, week_id, group_stats, all_students):
        """Create Excel report for a specific week"""
        week_info = self.weeks_data[week_id]
        week_dir = week_info["directory"]
        
        filename = f"attendance_report_{week_id}.xlsx"
        filepath = os.path.join(week_dir, filename)
        
        with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
            # Summary sheet
            summary_data = []
            for group, stats in group_stats.items():
                summary_data.append({
                    'Group': group,
                    'Total Students': stats['total_students'],
                    'Full Week': stats['full_week_count'],
                    'Partial': stats['partial_count'],
                    'Never': stats['never_attended_count'],
                    'Avg Attendance %': round(stats['average_attendance'], 1)
                })
            
            summary_df = pd.DataFrame(summary_data)
            summary_df.to_excel(writer, sheet_name='Summary', index=False)
            
            # Detailed students data
            students_data = []
            for student in all_students:
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
                    f'{week_info["start_date"].split("-")[1]} {week_info["start_date"].split("-")[0]}': '✓' if student['daily_attendance'][0] else '✗',
                    'Day 2': '✓' if student['daily_attendance'][1] else '✗',
                    'Day 3': '✓' if student['daily_attendance'][2] else '✗',
                    'Day 4': '✓' if student['daily_attendance'][3] else '✗',
                    'Day 5': '✓' if student['daily_attendance'][4] else '✗',
                })
            
            students_df = pd.DataFrame(students_data)
            students_df.to_excel(writer, sheet_name='All Students', index=False)
        
        print(f"Excel report saved: {filename}")
    
    def create_week_html_dashboard(self, week_id, group_stats, all_students, full_week, partial, never):
        """Create individual HTML dashboard for a specific week"""
        week_info = self.weeks_data[week_id]
        week_dir = week_info["directory"]
        
        # Prepare data for charts
        groups = list(group_stats.keys())
        full_counts = [group_stats[g]['full_week_count'] for g in groups]
        partial_counts = [group_stats[g]['partial_count'] for g in groups]
        never_counts = [group_stats[g]['never_attended_count'] for g in groups]
        avg_attendance = [group_stats[g]['average_attendance'] for g in groups]
        
        html_content = f'''
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Attendance Dashboard - {week_info["start_date"]} to {week_info["end_date"]} 2025</title>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <style>
        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            margin: 0;
            padding: 20px;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
        }}
        .container {{
            max-width: 1200px;
            margin: 0 auto;
            background: white;
            border-radius: 15px;
            padding: 30px;
            box-shadow: 0 10px 30px rgba(0,0,0,0.2);
        }}
        .header {{
            text-align: center;
            margin-bottom: 30px;
        }}
        .header h1 {{
            color: #2c3e50;
            margin: 0;
        }}
        .week-info {{
            background: #3498db;
            color: white;
            padding: 15px;
            border-radius: 10px;
            margin: 20px 0;
            text-align: center;
        }}
        .stats-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 20px;
            margin-bottom: 30px;
        }}
        .stat-card {{
            background: linear-gradient(135deg, #3498db, #2980b9);
            color: white;
            padding: 20px;
            border-radius: 10px;
            text-align: center;
        }}
        .stat-number {{
            font-size: 2rem;
            font-weight: bold;
        }}
        .charts-container {{
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 30px;
            margin-top: 30px;
        }}
        .chart-card {{
            background: #f8f9fa;
            padding: 20px;
            border-radius: 10px;
            border: 1px solid #dee2e6;
        }}
        .chart-title {{
            text-align: center;
            font-weight: bold;
            margin-bottom: 15px;
            color: #2c3e50;
        }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>Weekly Attendance Dashboard</h1>
            <div class="week-info">
                <h2>Week of {week_info["start_date"]} - {week_info["end_date"]}, 2025</h2>
                <p>Analysis completed on {week_info.get("analysis_date", "N/A")}</p>
            </div>
        </div>
        
        <div class="stats-grid">
            <div class="stat-card">
                <div class="stat-number">{len(all_students)}</div>
                <div>Total Students</div>
            </div>
            <div class="stat-card">
                <div class="stat-number">{full_week}</div>
                <div>Full Week Attendance</div>
            </div>
            <div class="stat-card">
                <div class="stat-number">{partial}</div>
                <div>Partial Attendance</div>
            </div>
            <div class="stat-card">
                <div class="stat-number">{never}</div>
                <div>Never Attended</div>
            </div>
        </div>
        
        <div class="charts-container">
            <div class="chart-card">
                <div class="chart-title">Group Distribution</div>
                <canvas id="groupChart"></canvas>
            </div>
            <div class="chart-card">
                <div class="chart-title">Overall Distribution</div>
                <canvas id="pieChart"></canvas>
            </div>
        </div>
    </div>

    <script>
        // Group Chart
        const groupCtx = document.getElementById('groupChart').getContext('2d');
        new Chart(groupCtx, {{
            type: 'bar',
            data: {{
                labels: {groups},
                datasets: [
                    {{
                        label: 'Full Week',
                        data: {full_counts},
                        backgroundColor: '#2ca02c'
                    }},
                    {{
                        label: 'Partial',
                        data: {partial_counts},
                        backgroundColor: '#ff7f0e'
                    }},
                    {{
                        label: 'Never',
                        data: {never_counts},
                        backgroundColor: '#d62728'
                    }}
                ]
            }},
            options: {{
                responsive: true,
                scales: {{
                    x: {{ stacked: true }},
                    y: {{ stacked: true }}
                }}
            }}
        }});

        // Pie Chart
        const pieCtx = document.getElementById('pieChart').getContext('2d');
        new Chart(pieCtx, {{
            type: 'pie',
            data: {{
                labels: ['Full Week', 'Partial', 'Never'],
                datasets: [{{
                    data: [{full_week}, {partial}, {never}],
                    backgroundColor: ['#2ca02c', '#ff7f0e', '#d62728']
                }}]
            }},
            options: {{
                responsive: true
            }}
        }});
    </script>
</body>
</html>'''
        
        filename = f"dashboard_{week_id}.html"
        with open(os.path.join(week_dir, filename), 'w', encoding='utf-8') as f:
            f.write(html_content)
        
        print(f"Individual HTML dashboard saved: {filename}")
    
    def save_week_data(self, week_id, group_stats, all_students):
        """Save week analysis data as JSON"""
        week_info = self.weeks_data[week_id]
        week_dir = week_info["directory"]
        
        # Prepare serializable data
        data = {
            'week_info': week_info,
            'group_stats': {
                group: {
                    'total_students': stats['total_students'],
                    'average_attendance': stats['average_attendance'],
                    'full_week_count': stats['full_week_count'],
                    'partial_count': stats['partial_count'],
                    'never_attended_count': stats['never_attended_count']
                }
                for group, stats in group_stats.items()
            },
            'summary': week_info.get('summary', {})
        }
        
        filename = f"data_{week_id}.json"
        with open(os.path.join(week_dir, filename), 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
    
    def create_master_dashboard(self):
        """Create the master HTML dashboard with week selection"""
        
        # Load all weeks data
        weeks_list = []
        for week_id, week_info in self.weeks_data.items():
            if 'summary' in week_info:
                weeks_list.append({
                    'id': week_id,
                    'display_name': f"{week_info['start_date']} - {week_info['end_date']}, 2025",
                    'students': week_info['summary']['total_students'],
                    'attendance_rate': f"{week_info['summary']['average_attendance']:.1f}%",
                    'dashboard_url': f"weeks/{week_id}/dashboard_{week_id}.html"
                })
        
        # Build the weeks cards HTML
        weeks_cards_html = ""
        if weeks_list:
            weeks_cards_html = '<div class="weeks-grid">'
            for week in weeks_list:
                weeks_cards_html += f'''
            <a href="{week['dashboard_url']}" class="week-card">
                <div class="week-title">Week: {week['display_name']}</div>
                <div class="week-stats">
                    <div class="stat-item">
                        <div class="stat-number">{week['students']}</div>
                        <div class="stat-label">Total Students</div>
                    </div>
                    <div class="stat-item">
                        <div class="stat-number">{week['attendance_rate']}</div>
                        <div class="stat-label">Average Attendance</div>
                    </div>
                </div>
            </a>'''
            weeks_cards_html += '</div>'
        else:
            weeks_cards_html = '''
        <div class="no-weeks">
            <h3>No weeks analyzed yet</h3>
            <p>Use the MultiWeekAttendanceAnalyzer to add and analyze attendance data for different weeks.</p>
        </div>'''

        html_content = f'''
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Multi-Week Attendance Analysis System</title>
    <style>
        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            margin: 0;
            padding: 0;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
        }}
        
        .header {{
            background: rgba(255,255,255,0.1);
            backdrop-filter: blur(10px);
            padding: 30px;
            text-align: center;
            color: white;
        }}
        
        .header h1 {{
            margin: 0;
            font-size: 2.5rem;
            text-shadow: 2px 2px 4px rgba(0,0,0,0.3);
        }}
        
        .container {{
            max-width: 1200px;
            margin: 0 auto;
            padding: 30px;
        }}
        
        .weeks-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fill, minmax(350px, 1fr));
            gap: 25px;
            margin-top: 30px;
        }}
        
        .week-card {{
            background: white;
            border-radius: 15px;
            padding: 25px;
            box-shadow: 0 8px 25px rgba(0,0,0,0.15);
            transition: all 0.3s ease;
            cursor: pointer;
            text-decoration: none;
            color: inherit;
        }}
        
        .week-card:hover {{
            transform: translateY(-5px);
            box-shadow: 0 15px 40px rgba(0,0,0,0.25);
        }}
        
        .week-title {{
            font-size: 1.4rem;
            font-weight: bold;
            color: #2c3e50;
            margin-bottom: 15px;
            border-bottom: 2px solid #3498db;
            padding-bottom: 10px;
        }}
        
        .week-stats {{
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 15px;
            margin-top: 15px;
        }}
        
        .stat-item {{
            text-align: center;
            padding: 10px;
            background: #f8f9fa;
            border-radius: 8px;
        }}
        
        .stat-number {{
            font-size: 1.8rem;
            font-weight: bold;
            color: #3498db;
        }}
        
        .stat-label {{
            font-size: 0.9rem;
            color: #7f8c8d;
            margin-top: 5px;
        }}
        
        .no-weeks {{
            text-align: center;
            background: white;
            padding: 50px;
            border-radius: 15px;
            color: #7f8c8d;
            font-size: 1.2rem;
        }}
        
        .add-week-btn {{
            background: #3498db;
            color: white;
            padding: 15px 30px;
            border: none;
            border-radius: 25px;
            font-size: 1.1rem;
            cursor: pointer;
            margin-top: 20px;
            transition: background 0.3s ease;
        }}
        
        .add-week-btn:hover {{
            background: #2980b9;
        }}
        
        .footer {{
            text-align: center;
            color: rgba(255,255,255,0.8);
            padding: 30px;
            margin-top: 50px;
        }}
    </style>
</head>
<body>
    <div class="header">
        <h1>Multi-Week Attendance Analysis System</h1>
        <p>Select a week to view detailed attendance statistics and reports</p>
    </div>
    
    <div class="container">
        {weeks_cards_html}
    </div>
    
    <div class="footer">
        <p>Multi-Week Attendance Analysis System - Total Weeks: {len(weeks_list)}</p>
    </div>
</body>
</html>'''
        
        with open('master_dashboard.html', 'w', encoding='utf-8') as f:
            f.write(html_content)
        
        print(f"Master dashboard created with {len(weeks_list)} weeks")
        return 'master_dashboard.html'
    
    def save_weeks_index(self):
        """Save weeks index as JSON for persistence"""
        with open('weeks_index.json', 'w', encoding='utf-8') as f:
            json.dump(self.weeks_data, f, indent=2, ensure_ascii=False)
    
    def load_weeks_index(self):
        """Load weeks index from JSON"""
        if os.path.exists('weeks_index.json'):
            with open('weeks_index.json', 'r', encoding='utf-8') as f:
                self.weeks_data = json.load(f)


# Usage example and main execution
if __name__ == "__main__":
    # Initialize the multi-week system
    analyzer = MultiWeekAttendanceAnalyzer()
    
    # Add the current week (31 Aug - 4 Sep)
    week1_info = analyzer.add_week(
        week_id="week_31Aug-4Sep",
        start_date="31-Aug",
        end_date="4-Sep", 
        excel_file_path=r'Attendance sheets\كشوفات الغياب الاسبوعي 31-8-2025(drive).xlsx',
        description="First week of September 2025"
    )
    
    # Analyze the week
    summary = analyzer.analyze_week("week_31Aug-4Sep", week1_info["excel_file"])
    
    # Create master dashboard
    analyzer.create_master_dashboard()
    
    # Save the weeks index
    analyzer.save_weeks_index()
    
    print("\\nMulti-week attendance system initialized!")
    print("Master dashboard created: master_dashboard.html")
    print("Week data organized in: weeks/ directory")
    print("\\nTo add more weeks:")
    print("1. analyzer.add_week(week_id, start_date, end_date, excel_file_path)")
    print("2. analyzer.analyze_week(week_id, excel_file_path)")
    print("3. analyzer.create_master_dashboard()")
