import pandas as pd
import os

def check_new_week_structure(excel_file_path):
    """
    Check the structure of the new week's Excel file to compare with previous week
    """
    print(f"=== Analyzing structure of: {os.path.basename(excel_file_path)} ===")
    
    try:
        # Read Excel file and get all sheets
        excel_data = pd.ExcelFile(excel_file_path)
        sheet_names = [sheet for sheet in excel_data.sheet_names if sheet != 'الورقة1']
        
        print(f"\\nTotal sheets found: {len(sheet_names)}")
        print("\\nGroup sheets:")
        for i, sheet in enumerate(sheet_names, 1):
            print(f"{i:2d}. {sheet}")
        
        # Analyze each sheet for student count
        total_students = 0
        group_details = {}
        
        for sheet_name in sheet_names:
            try:
                # Read the sheet
                df = pd.read_excel(excel_file_path, sheet_name=sheet_name, header=None)
                
                if df.empty:
                    continue
                
                # Count students in this group
                student_count = 0
                student_data_start = 3
                
                for row_idx in range(student_data_start, len(df)):
                    student_name = df.iloc[row_idx, 1]  # Column B
                    
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
                    
                    student_count += 1
                
                group_details[sheet_name] = student_count
                total_students += student_count
                
                print(f"  {sheet_name}: {student_count} students")
                
            except Exception as e:
                print(f"  {sheet_name}: Error reading sheet - {str(e)}")
                continue
        
        print(f"\\n=== SUMMARY ===")
        print(f"Total Groups: {len(group_details)}")
        print(f"Total Students: {total_students}")
        
        # Compare with previous week (31 Aug - 4 Sep had 439 students across 20 groups)
        print(f"\\n=== COMPARISON WITH PREVIOUS WEEK ===")
        print(f"Previous week (31 Aug - 4 Sep): 20 groups, 439 students")
        print(f"Current week (7 Sep - ?): {len(group_details)} groups, {total_students} students")
        
        if len(group_details) != 20:
            print(f"⚠️  GROUP COUNT CHANGED: {len(group_details) - 20:+d} groups")
        else:
            print("✅ Group count unchanged")
            
        if total_students != 439:
            print(f"⚠️  STUDENT COUNT CHANGED: {total_students - 439:+d} students")
        else:
            print("✅ Student count unchanged")
        
        # Show any new or missing groups
        previous_groups = {
            'SAIPEM 8', 'SAIPEM 7', 'SAIPEM 6', 'SAIPEM 5', 'SAIPEM 3', 'SAIPEM 4', 
            'SAIPEM 1', 'Alfa 2', 'SAIPEM 2', 'Sin 4', 'DEYE', 'SAM 1', 'SAM 2', 
            'SAM 6', 'SAM 3', 'SAM 4', 'SAM 5', 'Diang', 'Dabal', 'Aman+Elc+Fahss'
        }
        
        current_groups = set(group_details.keys())
        
        new_groups = current_groups - previous_groups
        missing_groups = previous_groups - current_groups
        
        if new_groups:
            print(f"\\n🆕 NEW GROUPS ({len(new_groups)}):")
            for group in sorted(new_groups):
                print(f"  + {group} ({group_details[group]} students)")
        
        if missing_groups:
            print(f"\\n❌ MISSING GROUPS ({len(missing_groups)}):")
            for group in sorted(missing_groups):
                print(f"  - {group}")
        
        if not new_groups and not missing_groups:
            print("\\n✅ All groups match previous week")
        
        return {
            'total_groups': len(group_details),
            'total_students': total_students,
            'group_details': group_details,
            'new_groups': list(new_groups),
            'missing_groups': list(missing_groups)
        }
        
    except Exception as e:
        print(f"Error analyzing file: {str(e)}")
        return None

# Check the new week's structure
new_week_file = r'Attendance sheets\كشوفات الغياب الاسبوعي 07-09-2025(drive).xlsx'
structure_info = check_new_week_structure(new_week_file)