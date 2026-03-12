import pandas as pd
from icalendar import Calendar, Event
from pprint import pp
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo

DEBUG_LEVEL = "Normal"

def get_resource_path(relative_path):
    """获取资源的绝对路径，兼容开发环境和打包后的环境"""

    import sys
    import os
    try:
        # PyInstaller创建临时文件夹，将路径存储在_MEIPASS中
        base_path = sys._MEIPASS
    except Exception:
        # 开发环境：使用当前目录
        base_path = os.path.abspath(".")
    
    return os.path.join(base_path, relative_path)

def get_all_exams():
    # Read Excel timetable
    source_table_path = get_resource_path("2026 JUNE CIE ExamTimetableComponentSpreadSheet.xlsx")
    source_table = pd.read_excel(source_table_path, sheet_name="ExamTimetableComponentSprea")

    exams = []
    for i in range(len(source_table)):
        event_dict = {}

        for row_name in source_table:
            event_dict[row_name] = source_table[row_name][i]

        event_dict["Subject Code"] = str(event_dict["Code"]).split("/")[0]
        
        try:
            exam_time = str(event_dict["Exam Time"]).split("-")
            exam_date = event_dict["Exam Date"].strftime("%Y-%-m-%-d")
            event_dict["Exam Start Time"] = datetime.strptime(f"{exam_date} {exam_time[0]}", "%Y-%m-%d %H:%M").replace(tzinfo=ZoneInfo('Asia/Shanghai'))
            event_dict["Exam End Time"] = datetime.strptime(f"{exam_date} {exam_time[1]}", "%Y-%m-%d %H:%M").replace(tzinfo=ZoneInfo('Asia/Shanghai'))
        except Exception as e:
            if DEBUG_LEVEL == "All":
                print("Unable to get the exam start time and end time:")
                pp(event_dict)
                print("="*40)
            continue

        exams.append(event_dict)
    
    return exams


def generate(student_name: str, student_exams: list[str] = []):

    all_exams = get_all_exams()
    cal = Calendar()

    for exam in all_exams:
        event = Event()

        if student_exams:
            if exam["Subject Code"] not in student_exams \
                and exam["Code"] not in student_exams:
                continue

        event.add(
            "summary",
            f"{exam["Code"]} {exam["Syllabus"]}"
        )
        event.add(
            "location",
            exam["Room"]
        )
        event.add(
            "dtstart",
            exam["Exam Start Time"]
        )
        event.add(
            "dtend",
            exam["Exam End Time"]
        )
        event.add('dtstamp', datetime.now()) 
        description = f"Component Title: {exam["Component Title"]}\n\nQualification: {exam["Qualification"]}"
        event.add(
            "description",
            description
        )
        event.add('valarm',
            {
                'trigger': timedelta(minutes=-15),  # 提前15分钟提醒
                'action': 'DISPLAY',
                'description': 'Enter the Exam Room'
            }
        )

        print(f"Add: {exam["Code"]} {exam["Syllabus"]}")
        cal.add_component(event)

    filename = f'{student_name}-Exams2026S_{datetime.now().timestamp()}.ics'
    with open(filename, 'wb') as f:
        f.write(cal.to_ical())

    return filename

if __name__ == "__main__":
    generate("Tom", [  # fill you name
        "9231/14",  # place you subject codes
        "9231/44",
        "9700/14",
        "9700/24",
        "9700/38",
    ])

    # generate("All")  # create a calender for all exams