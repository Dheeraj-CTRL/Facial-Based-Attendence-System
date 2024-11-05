import cv2
import openpyxl
import face_recognition
from datetime import datetime
face_cascade = cv2.CascadeClassifier(cv2.data.haarcascades + 'haarcascade_frontalface_default.xml')
camera_index = int(input("Enter 0 to use the laptop camera or 1 to use the classroom camera: "))
cap = cv2.VideoCapture(camera_index)
if not cap.isOpened():
    print("Error: Could not open the selected camera.")
    exit()
workbook = openpyxl.Workbook()
sheet = workbook.active
sheet.title = "Attendance"

sheet['A1'] = 'Date'
sheet['B1'] = 'Roll Number'
sheet['C1'] = 'In-Time'
sheet['D1'] = 'Out-Time'
sheet['E1'] = 'Subject'
sheet['F1'] = 'Faculty'
sheet['G1'] = 'Student Name'
sheet['H1'] = 'Status' 
students = [
    #('Roll Number', 'Name', r"Image Path of the students")
    ("20207156",'Mr X','')
]

timetable = {
    'Monday': [
       #('Class Starting Time', 'Ending Time', 'Course Code - Course Name', 'Faculty Name'), 
    ],
    'Tuesday': [
       #('Class Starting Time', 'Ending Time', 'Course Code - Course Name', 'Faculty Name'), 
    ],
    'Wednesday': [
        #('Class Starting Time', 'Ending Time', 'Course Code - Course Name', 'Faculty Name'), 
    ],
    'Thursday': [
        #('Class Starting Time', 'Ending Time', 'Course Code - Course Name', 'Faculty Name'), 
    ],
    'Friday': [
        #('Class Starting Time', 'Ending Time', 'Course Code - Course Name', 'Faculty Name'), 
        
    ]
}
attendance_record = {}
def load_students_from_list():
    loaded_students = []
    for roll_number, student_name, image_path in students:
        try:
            student_image = face_recognition.load_image_file(image_path)
            student_face_encoding = face_recognition.face_encodings(student_image)[0]
            loaded_students.append((roll_number, student_name, student_face_encoding))
        except Exception as e:
            print(f"Error processing {student_name}'s image: {e}")
    return loaded_students
loaded_students = load_students_from_list()
def get_current_class():
    now = datetime.now()
    current_day = now.strftime('%A')
    current_time = now.strftime('%H:%M')
    
    if current_day in timetable:
        for start_time, end_time, subject, faculty in timetable[current_day]:
            if start_time <= current_time <= end_time:
                return subject, faculty
    return None, None
def take_attendance(frame):
    subject, faculty = get_current_class()
    
    if subject:
        now = datetime.now()
        current_date = now.strftime("%Y-%m-%d")

        face_locations = face_recognition.face_locations(frame)
        face_encodings = face_recognition.face_encodings(frame, face_locations)
        
        for face_encoding in face_encodings:
            matches = face_recognition.compare_faces([student[2] for student in loaded_students], face_encoding)
            
            if True in matches:
                matched_index = matches.index(True)
                roll_number, student_name = loaded_students[matched_index][:2]
                
                if roll_number not in attendance_record:
                    attendance_record[roll_number] = {'in_time': now.strftime("%H:%M:%S")}
                    print(f"Marked in-time for {student_name} ({roll_number})")
                    
        return subject, faculty, current_date
    else:
        print("No class is currently active.")
        return None, None, None
def mark_out_time(subject, faculty, current_date):
    now = datetime.now()
    current_time = now.strftime("%H:%M:%S")
    
    all_students = [student[0] for student in students] 
    present_students = list(attendance_record.keys())  
    
    print("Present Students and their In/Out Times:")
    
    for roll_number in present_students:
        student_name = next((student[1] for student in loaded_students if student[0] == roll_number), "Unknown")
        sheet.append([current_date, roll_number, attendance_record[roll_number]['in_time'], current_time, subject, faculty, student_name, 'Present'])
        print(f"{student_name} ({roll_number}): In-Time: {attendance_record[roll_number]['in_time']}, Out-Time: {current_time}")
    
    
    for roll_number in all_students:
        if roll_number not in present_students:
            student_name = next((student[1] for student in loaded_students if student[0] == roll_number), "Unknown")
            sheet.append([current_date, roll_number, '', current_time, subject, faculty, student_name, 'Absent'])

    
    attendance_record.clear()
    
    workbook.save("Attendance.xlsx")
    print("Attendance has been recorded.")
while True:
    ret, frame = cap.read()
    if not ret:
        print("Failed to grab frame.")
        break

    frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)

    subject, faculty, current_date = take_attendance(frame)

    cv2.imshow('Video', frame)

    if cv2.waitKey(1) & 0xFF == ord('q'):
        mark_out_time(subject, faculty, current_date)  
        break

cap.release()
cv2.destroyAllWindows()
