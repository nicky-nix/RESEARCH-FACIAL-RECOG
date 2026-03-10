@echo off
echo ============================================
echo Fast Attendance System - Installation
echo Sibutad National High School
echo ============================================
echo.
echo Installing required packages...
echo.

pip install opencv-python
pip install face-recognition
pip install numpy
pip install pandas
pip install openpyxl

echo.
echo ============================================
echo Installation complete!
echo ============================================
echo.
echo Next steps:
echo 1. Prepare student photos in 'student_photos' folder
echo 2. Edit 'students.json' with student information
echo 3. Run: python fast_attendance_system.py
echo.
pause
