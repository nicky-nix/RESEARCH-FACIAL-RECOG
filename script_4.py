
# Create a quick start guide
quickstart = """# 🚀 QUICK START GUIDE
## Fast Facial Recognition Attendance System

### ⚡ 5-Minute Setup

#### Step 1: Install Python Packages
Open Command Prompt and run:
```bash
pip install opencv-python face-recognition pandas openpyxl
```

#### Step 2: Prepare Student Photos
1. Create a folder named `student_photos`
2. Take clear photos of each student
3. Save photos as `{student_id}.jpg` (e.g., `2024-001.jpg`)
4. Edit `students.json` with your student information

#### Step 3: Register Students
```bash
python fast_attendance_system.py
```
- Choose option: **1**
- Enter folder path: `student_photos`
- Wait for registration to complete

#### Step 4: Run Attendance
```bash
python fast_attendance_system.py
```
- Choose option: **2**
- Camera will open automatically
- Students stand in front of camera
- Press **'s'** to save
- Press **'q'** to quit

#### Step 5: Check Excel File
- Open `Attendance_Records.xlsx`
- View professionally formatted attendance data

### 🎯 Key Controls
- **Option 1**: Register students (first time only)
- **Option 2**: Run attendance (daily use)
- **Option 3**: View reports
- **'s' key**: Save attendance during camera mode
- **'q' key**: Quit camera mode

### 📋 File Structure
```
your_folder/
├── fast_attendance_system.py    # Main program
├── students.json                 # Student list (EDIT THIS)
├── students_database.pkl         # Face database (auto-generated)
├── Attendance_Records.xlsx       # Output file (auto-generated)
└── student_photos/               # Student photos folder
    ├── 2024-001.jpg
    ├── 2024-002.jpg
    └── ...
```

### ✅ Success Indicators
- ✓ Green boxes around recognized faces
- ✓ Console shows "Attendance marked: [Name]"
- ✓ Excel file updates with new records
- ✓ No duplicate entries per day

### ❌ Common Issues & Fixes

**Problem**: "No face found in image"
- **Fix**: Ensure photos are clear and front-facing

**Problem**: Camera won't open
- **Fix**: Close other apps using camera (Zoom, Teams, etc.)

**Problem**: Face not recognized
- **Fix**: Improve lighting, student faces camera directly

**Problem**: dlib installation error
- **Fix**: Install Visual Studio Build Tools with C++ workload

### 📞 Need Help?
Check the full README.md for detailed troubleshooting!

---
**Tip**: For best results, register students in good lighting conditions!
"""

with open('QUICKSTART.md', 'w', encoding='utf-8') as f:
    f.write(quickstart)

print("✓ Quick start guide created: QUICKSTART.md")
