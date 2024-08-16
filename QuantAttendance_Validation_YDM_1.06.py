import pandas as pd 
from datetime import datetime, timedelta, time 

total_time_path = 'C:/Users/damod/Downloads/Total Time Card_20240709105647.xlsx'
total_time_sheet = '20240709'
ttl_df = pd.read_excel(total_time_path, sheet_name=total_time_sheet, header = 2)

print(ttl_df.head()) #Debugg line to evaluate pd.read_excel, NaN values appear in some columns 
print(ttl_df['Clock In'].head(3)) #Debugg line to evaluate if function will perform as intended 
print(ttl_df['Clock Out'].head(3)) #Debugg line to evaluate if function will perform as intended 

def evaluate_ttl(df):
	#1 --> Evaluate clock_in and clock_out using XOR logic
	employee_df = ['Employee ID', 'First Name', 'Department', 'Weekday', 'Exception', 'Timetable', 'Duration', 'Check In', 'Check Out', 'Clock In', 'Clock Out']#Relevant employee data to be used in creating result excel sheets
	clock_in = ttl_df['Clock In'].astype(str).str.strip()
	clock_in.replace('nan', pd.NA, inplace=True)
	clock_out = ttl_df['Clock Out'].astype(str).str.strip()
	clock_out.replace('nan', pd.NA, inplace=True)
	miss_df = df[(clock_in.isna() & clock_out.notna()) | (clock_in.notna() & clock_out.isna())]
	print (miss_df.head(3))#Displays clock_in/clock_out pairs withtout either/or value but not both
	mispunch_df = miss_df[employee_df].copy()#Employee data
	mispunch_df['Reason'] = mispunch_df.apply(lambda row: 'No Clock In' if pd.isna(row['Clock In']) else 'No Clock Out', axis=1)
	
	print (mispunch_df.head(10))

	#2 --> Off Change Detection
	timetable = ttl_df['Timetable']
	off_days = df[timetable.astype(str).str.contains('O', na=False)]
	off_days = off_days[employee_df].copy()
	def check_in_out(row):
	    if pd.notna(row['Clock In']) or pd.notna(row['Clock Out']):
	        return 'Off Day Change'
	    else:
	        return 'No Change'
	off_days['Checked In/Out'] = off_days.apply(check_in_out, axis=1)
	off_days = off_days[off_days['Checked In/Out'] == 'Off Day Change']
	print (off_days.head(10))#Evaluate if boolean mask is correct 
	#3--> Results
	attendance_validation_sheet = 'C:/Users/damod/Documents/Validated_Attendance_15082024.xlsx'
	with pd.ExcelWriter(attendance_validation_sheet, engine='openpyxl') as writer: 
		mispunch_df.to_excel(writer, sheet_name='Mispunched Entries', index=False)
		off_days.to_excel(writer, sheet_name='Off Change Detection', index=False)
evaluate_ttl(ttl_df)