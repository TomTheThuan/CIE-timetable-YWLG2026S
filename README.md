# CIE Exam Calendar

A simple tool to generate iCalendar (.ics) files for CIE exams from an Excel timetable.

## Features

- 📅 Generate calendar files from CIE exam timetable spreadsheet
- 🔍 Filter exams by subject codes
- ⏰ Automatic timezone handling (Asia/Shanghai)
- 🔔 15-minute reminders for each exam
- 🖥️ Simple GUI interface

## Requirements

- Python 3.8+
- Dependencies (installed via uv):
  - pandas
  - ttkbootstrap
  - icalendar
  - openpyxl

## Installation

```bash
# Install uv if not already installed
pip install uv

# Install dependencies
uv sync
```

## Usage

### GUI Mode
```bash
python gui.py
```

1. Enter student name
2. List subject codes (one per line)
3. Click "Generate"
4. Find your .ics file in the current directory

### Command Line Mode
```bash
python main.py
```
You need to modify the code for your own subjects.

## Input File Format

Place `2026 JUNE CIE ExamTimetableComponentSpreadSheet.xlsx` in the project root directory.

## Output

Generates .ics files named: `{StudentName}-Exams2026S_{timestamp}.ics`

## License

MIT