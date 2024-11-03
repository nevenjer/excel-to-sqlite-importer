# excel-to-sqlite-importer
# A Python-based tool to efficiently import data from Excel files into a SQLite database, automating data transfer and setup for further analysis.

import pandas as pd
import sqlite3

# โหลดข้อมูลจากไฟล์ Excel
file_path = r'D:Your_Path_File.xlsx'  # เปลี่ยนเป็นที่อยู่ไฟล์ Excel ของคุณ
df = pd.read_excel(file_path, header=None)

# เลือกข้อมูลคอลัมน์ B2 เป็นต้นไป (คอลัมน์ที่ 1) และ C2 เป็นต้นไป (คอลัมน์ที่ 2) จนถึง Q2 (คอลัมน์ที่ 16)
selected_data = df.iloc[1:, 1:17]  # เลือกแถวที่ 2 เป็นต้นไป (index 1) และคอลัมน์ B ถึง Q (index 1 ถึง 16)

# สร้างการเชื่อมต่อกับ SQLite3
conn = sqlite3.connect(r'D:Your_Path_File\AUTO_SQL.db')  # เปลี่ยนชื่อฐานข้อมูลตามต้องการ
cursor = conn.cursor()

# สร้างตาราง (ถ้ายังไม่มี)
cursor.execute('''
CREATE TABLE IF NOT EXISTS FREQ_BPM (
    no INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL,
    id_pos TEXT,
    ca TEXT,
    name TEXT,
    service TEXT,
    recipt TEXT,
    debt REAL,
    rounding_num REAL,
    total REAL,
    cash REAL,
    cheque REAL,
    pay_in_slip REAL,
    qr_payment REAL,
    date_time TEXT,
    status TEXT,
    num_bill INTEGER,
    num_clients INTEGER
)
''')

# บันทึกข้อมูลลงในฐานข้อมูล
for index, row in selected_data.iterrows():
    cursor.execute('''
    INSERT INTO FREQ_BPM (id_pos, ca, name, service, recipt, debt,
    rounding_num, total, cash, cheque, pay_in_slip, qr_payment, date_time, status, num_bill, num_clients)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
''', tuple(row))


# บันทึกการเปลี่ยนแปลงและปิดการเชื่อมต่อ
conn.commit()

print('Your data is uploaded completely to SQLite')

conn.close()
