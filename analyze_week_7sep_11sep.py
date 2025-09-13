from multi_week_analyzer import MultiWeekAttendanceAnalyzer

def analyze_new_week():
    """
    Analyze the new week (7-Sep to 11-Sep 2025)
    """
    print("🚀 Starting analysis for new week: 7-Sep to 11-Sep 2025")
    print("=" * 60)
    
    # Initialize the multi-week analyzer
    analyzer = MultiWeekAttendanceAnalyzer()
    
    # Load existing weeks data
    analyzer.load_weeks_index()
    print(f"📂 Loaded existing weeks: {list(analyzer.weeks_data.keys())}")
    
    # Add the new week
    week_info = analyzer.add_week(
        week_id="week_7Sep-11Sep",
        start_date="7-Sep",
        end_date="11-Sep", 
        excel_file_path=r'Attendance sheets\كشوفات الغياب الاسبوعي 07-09-2025(drive).xlsx',
        description="Second week of September 2025 - Post DEYE refinement"
    )
    
    print(f"✅ Added new week: {week_info['week_id']}")
    print(f"📅 Date range: {week_info['start_date']} - {week_info['end_date']}, 2025")
    print(f"📁 Directory: {week_info['directory']}")
    
    # Analyze the week
    print("\n🔍 Starting attendance analysis...")
    summary = analyzer.analyze_week("week_7Sep-11Sep", week_info["excel_file"])
    
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
    
    print("\n🎉 New week analysis completed successfully!")
    print("=" * 60)
    
    return analyzer, summary

if __name__ == "__main__":
    analyzer, summary = analyze_new_week()