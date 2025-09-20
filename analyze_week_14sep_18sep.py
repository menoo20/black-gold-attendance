from multi_week_analyzer import MultiWeekAttendanceAnalyzer

def analyze_new_week():
    """
    Analyze the new week (14-Sep to 18-Sep 2025)
    """
    print("🚀 Starting analysis for new week: 14-Sep to 18-Sep 2025")
    print("=" * 60)
    
    # Initialize the multi-week analyzer
    analyzer = MultiWeekAttendanceAnalyzer()
    
    # Load existing weeks data
    analyzer.load_weeks_index()
    print(f"📂 Loaded existing weeks: {list(analyzer.weeks_data.keys())}")
    
    # Add the new week
    week_info = analyzer.add_week(
        week_id="week_14Sep-18Sep",
        start_date="14-Sep",
        end_date="18-Sep", 
        excel_file_path=r'Attendance sheets\كشوفات الغياب الاسبوعي 14-09-2025(drive).xlsx',
        description="Third week of September 2025 - Continued monitoring"
    )
    
    print(f"✅ Added new week: {week_info['week_id']}")
    print(f"📅 Date range: {week_info['start_date']} - {week_info['end_date']}, 2025")
    print(f"📁 Directory: {week_info['directory']}")
    
    # Analyze the week
    print("\n🔍 Starting attendance analysis...")
    summary = analyzer.analyze_week("week_14Sep-18Sep", week_info["excel_file"])
    
    if summary:
        print("\n📊 ANALYSIS SUMMARY:")
        print(f"   Total Students: {summary['total_students']}")
        print(f"   Full Week: {summary['full_week']} ({summary['full_week']/summary['total_students']*100:.1f}%)")
        print(f"   Partial: {summary['partial']} ({summary['partial']/summary['total_students']*100:.1f}%)")
        print(f"   Never: {summary['never']} ({summary['never']/summary['total_students']*100:.1f}%)")
        print(f"   Average Attendance: {summary['average_attendance']:.1f}%")
        print(f"   Groups Analyzed: {summary['groups']}")
    
    # Update master dashboard
    print("\n🎨 Updating master dashboard...")
    master_file = analyzer.create_master_dashboard()
    print(f"✅ Master dashboard updated: {master_file}")
    
    # Save weeks index
    analyzer.save_weeks_index()
    print("💾 Weeks index saved")
    
    return summary

if __name__ == "__main__":
    summary = analyze_new_week()
    print("\n🎉 Analysis complete!")
    print("=" * 60)