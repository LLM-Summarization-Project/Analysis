"""
คำนวณ Average ของแต่ละ whisper_temp และ reference_tool
แล้วเพิ่ม sheet ใหม่ใน evaluation_results.xlsx
"""

import pandas as pd

# อ่านข้อมูล
df = pd.read_excel(r'd:\final_project_\analysis\evaluation_results.xlsx')

print("Columns:", df.columns.tolist())
print("Unique whisper_temp:", df['whisper_temp'].unique())
print("Unique reference_tool:", df['reference_tool'].unique())

# คำนวณ average แยกตาม whisper_temp และ reference_tool
avg_by_temp_ref = df.groupby(['whisper_temp', 'reference_tool']).agg({
    'precision': 'mean',
    'recall': 'mean',
    'f1': 'mean',
    'cand_length': 'mean',
    'ref_length': 'mean'
}).round(4).reset_index()

# คำนวณ average แยกตาม whisper_temp อย่างเดียว
avg_by_temp = df.groupby(['whisper_temp']).agg({
    'precision': 'mean',
    'recall': 'mean',
    'f1': 'mean',
    'cand_length': 'mean',
    'ref_length': 'mean'
}).round(4).reset_index()

# คำนวณ average แยกตาม reference_tool อย่างเดียว
avg_by_ref = df.groupby(['reference_tool']).agg({
    'precision': 'mean',
    'recall': 'mean',
    'f1': 'mean',
    'cand_length': 'mean',
    'ref_length': 'mean'
}).round(4).reset_index()

print("\n=== Average by Temp & Reference ===")
print(avg_by_temp_ref)

print("\n=== Average by Temp ===")
print(avg_by_temp)

print("\n=== Average by Reference ===")
print(avg_by_ref)

# บันทึกลงไฟล์เดิม โดยเพิ่ม sheets ใหม่
with pd.ExcelWriter(r'd:\final_project_\analysis\evaluation_results.xlsx', 
                    mode='a', 
                    engine='openpyxl',
                    if_sheet_exists='replace') as writer:
    avg_by_temp_ref.to_excel(writer, sheet_name='avg_by_temp_ref', index=False)
    avg_by_temp.to_excel(writer, sheet_name='avg_by_temp', index=False)
    avg_by_ref.to_excel(writer, sheet_name='avg_by_ref', index=False)

print("\n✅ บันทึก 3 sheets ใหม่ลง evaluation_results.xlsx เรียบร้อย:")
print("   - avg_by_temp_ref")
print("   - avg_by_temp")
print("   - avg_by_ref")
