import pyodbc

# Connection details
database_path = r'E:\\Assignments\\ITC LAB PROJECT\\AttendenceSystem.accdb' 
connection_str = f'DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={database_path};'

try:
    conn = pyodbc.connect(connection_str)
    cursor = conn.cursor()
    print("Connected to Access Database!")
except Exception as e:
    print("Error connecting to database:", e)
    exit()

def insert_student(name, student_class):
    query = """
    insert into Students (Name, Class)
    VALUES (?, ?)
    """
    cursor.execute(query, (name, student_class))
    conn.commit()
    print(f"Student '{name}' added successfully!")

    cursor.execute("SELECT @@IDENTITY AS LastID")
    result = cursor.fetchone()
    return result.LastID

def insert_attendance(sid, date, status):
    query = """
    insert into Attendance (SID, Date, Status)
    values (?, ?, ?)
    """
    cursor.execute(query, (sid, date, status))
    conn.commit()
    print(f"Attendance record for Student ID {sid} added successfully!")


def main():
    while True:
        print("\n--- Add Student and Attendance Records ---")
        name = input("Enter Student Name: ")
        student_class = input("Enter Class: ")
        date = input("Enter Attendance Date (YYYY-MM-DD): ")
        status = input("Enter Attendance Status (Present/Absent): ")

        sid = insert_student(name, student_class)
        insert_attendance(sid, date, status)

        more = input("Do you want to add another record? (yes/no): ").strip().lower()
        if more != 'yes':
            break

    print("All records have been added successfully!")

#Runprogram
if __name__ == "__main__":
    main()

cursor.close()
conn.close()
