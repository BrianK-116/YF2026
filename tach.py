import pandas as pd
import json

df = pd.read_excel('DataTV.xlsx', sheet_name='ChamCong', header=None)
# Hàng 3 (index 2) là tiêu đề, dữ liệu từ hàng 4 (index 3)
students = {}
for i in range(3, len(df)):
    row = df.iloc[i]
    mssv = str(row[4]).strip() # Cột E: MSSV
    name = str(row[1]).strip() # Cột B: Họ và tên
    role = str(row[5]).strip() # Cột F: Hiện là
    if mssv and mssv != 'nan':
        students[mssv] = {"name": name, "role": role}

with open('students.json', 'w', encoding='utf-8') as f:
    json.dump(students, f, ensure_ascii=False, indent=2)

print(f"Đã xuất thành công {len(students)} sinh viên sang students.json!")