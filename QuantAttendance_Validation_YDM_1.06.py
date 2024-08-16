import pandas as pd
from datetime import datetime

# Define paths and sheet names
total_time_path = 'C:/Users/damod/Downloads/Total Time Card_20240709105647.xlsx'
total_time_sheet = '20240709'
filo_path = 'C:/Users/damod/Downloads/First In Last Out_20240716141519.xlsx'
filo_sheet = '20240716'

# Load data
ttl_df = pd.read_excel(total_time_path, sheet_name=total_time_sheet, header=2)
filo_df = pd.read_excel(filo_path, sheet_name=filo_sheet, header=2)

# Print columns to debug
print("Columns in ttl_df:")
print(ttl_df.columns)
print("Columns in filo_df:")
print(filo_df.columns)

# Define shift times
shift_times = {
    'A': {'start': '07:00', 'end': '15:00'},
    'B': {'start': '15:00', 'end': '23:00'},
    'C': {'start': '23:00', 'end': '07:00'},
    'G': {'start': '09:00', 'end': '17:00'},
    'General': {'start': '09:00', 'end': '18:00'}
}

def parse_time(t):
    if pd.isna(t):
        return None
    try:
        return datetime.strptime(t, '%H:%M').time()
    except ValueError:
        return None

def is_within_shift(check_time, shift_start, shift_end):
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
    first_check_in = parse_time(row.get('First Check In'))
    last_check_out = parse_time(row.get('Last Check Out'))

    for shift, times in shift_times.items():
        shift_start = parse_time(times['start'])
        shift_end = parse_time(times['end'])

        if first_check_in and is_within_shift(first_check_in, shift_start, shift_end):
            return shift

        if last_check_out and is_within_shift(last_check_out, shift_start, shift_end):
            return shift

    return None

def evaluate_ttl(df):
    # Strip whitespace from column names
    df.columns = df.columns.str.strip()
    filo_df.columns = filo_df.columns.str.strip()

    # Evaluate clock_in and clock_out using XOR logic
    employee_df = ['Employee ID', 'First Name', 'Department', 'Date', 'Weekday', 'Exception', 'Timetable', 'Duration', 'Check In', 'Check Out', 'Clock In', 'Clock Out']
    df['Clock In'] = df['Clock In'].astype(str).str.strip().replace('nan', pd.NA)
    df['Clock Out'] = df['Clock Out'].astype(str).str.strip().replace('nan', pd.NA)

    miss_df = df[(df['Clock In'].isna() & df['Clock Out'].notna()) | (df['Clock In'].notna() & df['Clock Out'].isna())]
    print(miss_df.head(3))  # Display clock_in/clock_out pairs with either/or value but not both

    # Create a DataFrame for mispunch records
    mispunch_df = miss_df[employee_df].copy() 

    # Perform the merge
    mispunch_df = pd.merge(
        mispunch_df,
        filo_df[['Employee ID', 'Date', 'First Check In', 'Last Check Out']],
        on=['Employee ID', 'Date'],
        how='left'
)
    # Apply shift evaluation
    mispunch_df['Reason'] = mispunch_df.apply(
        lambda row: f"Worked Shift {shift_evaluate(row)}" 
                    if shift_evaluate(row) 
                    else 'No Clock In' if pd.isna(row['Clock In']) else 'No Clock Out', 
        axis=1
    )
    print(mispunch_df.head(10))

    # Off Change Detection
    timetable = df['Timetable']
    off_days = df[timetable.astype(str).str.contains('O', na=False)]
    off_days = off_days[employee_df].copy()

    # Perform the merge
    off_days = pd.merge(
        off_days,
        filo_df[['Employee ID', 'Date', 'First Check In', 'Last Check Out']],
        on=['Employee ID', 'Date'],
        how='left'
    )

    # Apply shift evaluation for off_days
    off_days['Comment'] = off_days.apply(
        lambda row: f"Worked Shift {shift_evaluate(row)}" 
                    if pd.notna(row['First Check In']) or pd.notna(row['Last Check Out']) 
                    else 'Remove',
        axis=1
    )
    off_days = off_days[off_days['Comment'] != 'Remove']
    off_days = off_days.drop(['Duration', 'Clock In', 'Clock Out'], axis=1)
    print(off_days.head(10))  # Evaluate if boolean masks are correct

    # Save Results
    attendance_validation_sheet = 'C:/Users/damod/Documents/Validated_Attendance_15082024.xlsx'
    with pd.ExcelWriter(attendance_validation_sheet, engine='openpyxl') as writer:
        mispunch_df.to_excel(writer, sheet_name='Mispunched Entries', index=False)
        off_days.to_excel(writer, sheet_name='Off Change Detection', index=False)

evaluate_ttl(ttl_df)
