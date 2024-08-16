import pandas as pd
from datetime import datetime, timedelta, time

total_time_path = 'C:/Users/damod/Downloads/Total Time Card_20240709105647.xlsx'
total_time_sheet = '20240709'
ttl_df = pd.read_excel(total_time_path, sheet_name=total_time_sheet, header=2)

# Debugging lines to evaluate pd.read_excel
print(ttl_df.head())
print(ttl_df['Clock In'].head(3))
print(ttl_df['Clock Out'].head(3))

import pandas as pd
from datetime import datetime, timedelta

# Define shift times
shift_times = {
    'A': {'start': '07:00', 'end': '15:00'},
    'B': {'start': '15:00', 'end': '23:00'},
    'C': {'start': '23:00', 'end': '07:00'},
    'G': {'start': '09:00', 'end': '18:00'},
    'General': {'start': '08:00', 'end': '17:00'}
}

def parse_time(t):
    if pd.isna(t):
        return None
    try:
        return datetime.strptime(t, '%H:%M').time()
    except ValueError:
        return None

def is_within_shift(check_time, shift_start, shift_end):
    # Convert times to datetime for comparison
    today = datetime.today()
    shift_start_datetime = datetime.combine(today, shift_start)
    shift_end_datetime = datetime.combine(today, shift_end)

    if shift_start > shift_end:  # Overnight shift
        if check_time >= shift_start or check_time <= shift_end:
            return True
    else:
        if shift_start <= check_time <= shift_end:
            return True

    return False

def shift_evaluate(row):
    first_check_in = parse_time(row['Clock In'])
    last_check_out = parse_time(row['Clock Out'])

    for shift, times in shift_times.items():
        shift_start = parse_time(times['start'])
        shift_end = parse_time(times['end'])

        if first_check_in and is_within_shift(first_check_in, shift_start, shift_end):
            return shift

        if last_check_out and is_within_shift(last_check_out, shift_start, shift_end):
            return shift

    return None

def evaluate_ttl(df):
    # Evaluate clock_in and clock_out using XOR logic
    employee_df = ['Employee ID', 'First Name', 'Department', 'Weekday', 'Exception', 'Timetable', 'Duration', 'Check In', 'Check Out', 'Clock In', 'Clock Out']
    df['Clock In'] = df['Clock In'].astype(str).str.strip().replace('nan', pd.NA)
    df['Clock Out'] = df['Clock Out'].astype(str).str.strip().replace('nan', pd.NA)
    
    miss_df = df[(df['Clock In'].isna() & df['Clock Out'].notna()) | (df['Clock In'].notna() & df['Clock Out'].isna())]
    print(miss_df.head(3))  # Display clock_in/clock_out pairs with either/or value but not both
    
    # Create a DataFrame for mispunch records
    mispunch_df = miss_df[employee_df].copy()
    
    # Apply shift evaluation
    mispunch_df['Reason'] = mispunch_df.apply(lambda row: f"Worked Shift {shift_evaluate(row)}" if shift_evaluate(row) else 'No Clock In' if pd.isna(row['Clock In']) else 'No Clock Out', axis=1)
    print(mispunch_df.head(10))

    # Off Change Detection
    timetable = df['Timetable']
    off_days = df[timetable.astype(str).str.contains('O', na=False)]
    off_days = off_days[employee_df].copy()
    
    off_days['Checked In/Out'] = off_days.apply(lambda row: 'Off Day Change' if pd.notna(row['Clock In']) or pd.notna(row['Clock Out']) else 'No Change', axis=1)
    off_days = off_days[off_days['Checked In/Out'] == 'Off Day Change']
    print(off_days.head(10))  # Evaluate if boolean mask is correct
    
    # Save Results
    attendance_validation_sheet = 'C:/Users/damod/Documents/Validated_Attendance_15082024.xlsx'
    with pd.ExcelWriter(attendance_validation_sheet, engine='openpyxl') as writer:
        mispunch_df.to_excel(writer, sheet_name='Mispunched Entries', index=False)
        off_days.to_excel(writer, sheet_name='Off Change Detection', index=False)

evaluate_ttl(ttl_df)