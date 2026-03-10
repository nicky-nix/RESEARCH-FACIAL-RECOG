

---

```markdown
# Fast Facial Recognition Attendance System (v5.5)
### Designed for Sibutad National High School

An automated, high-performance facial recognition system designed to streamline classroom attendance with professional Excel reporting and real-time status tracking.

---

## 🚀 Phase 1: Software Installation

Before running the code, you must prepare your environment with the following tools:

### 1. Install Python
* Download and install **Python 3.10 or newer** from [python.org](https://www.python.org/).
* **Important:** During installation, ensure you check the box that says **"Add Python to PATH."**

### 2. Install C++ Build Tools
The facial recognition library relies on `dlib`, which requires a C++ compiler:
* Download the **[Visual Studio Community](https://visualstudio.microsoft.com/downloads/)** installer.
* Select **"Desktop development with C++"** and install.
* Also, install CMake via terminal: `pip install cmake`.

### 3. Install Required Libraries
The system depends on several specialized libraries for computer vision and data management. Run the following command:
```bash
pip install opencv-python face_recognition numpy pandas openpyxl

```

---

## 📂 Phase 2: Initial Setup & File Preparation

### 1. The `students.json` File

The system uses a `students.json` file to link IDs to names and sections. If it doesn't exist, the system will generate a sample for you. The structure should be:

```json
{
    "students": [
        {"id": "2025-001", "name": "Juan Dela Cruz", "section": "ICT"},
        {"id": "2025-002", "name": "Maria Santos", "section": "Cookery"}
    ]
}

```

### 2. Preparing Student Photos

* Create a folder (e.g., `student_images`).
* Add clear photos of each student.
* **Naming Format:** The filename must begin with the **Student ID** (e.g., `2025-001.jpg`) to match the database records.

---

## 🛠 Phase 3: How to Use the Program

Run the script by executing:

```bash
python fast_attendance_system_v6_FINXED_FINAL.py

```

### Step 1: Register Students (Menu Option 1)

Use this option to "train" the system. Enter the path to your image folder, and the system will save face encodings to `students_database.pkl`.

### Step 2: Start Attendance (Menu Option 2)

1. **Select Subject:** Choose from predefined subjects like *Media and Info Literacy* or *Cookery*.
2. **Set Start Time:** Enter the class start time (HH:MM) to calculate "Late" status.
3. **Live Scanning:** The camera will open and track faces in real-time.
* **Green Box:** Recognized & Marked "Present".
* **Yellow Box:** Recognized & Marked "Late" (arrived after the threshold).
* **Verifying...:** The system is confirming the identity.



**Controls During Scanning:**

* `p`: Pause/Resume scanning.
* `s`: Save current progress to Excel immediately.
* `q`: Quit camera and finalize the session (automatically marks remaining students as "Absent").

### Step 3: Generate Reports (Menu Options 3 & 4)

Generate weekly or monthly summaries. The system calculates the total classes and the **Attendance Rate %** for every student.

---

## 📊 Phase 4: Accessing Records & Output

All files are automatically organized into folders created by the system:

### 📁 `Attendance_Records/`

Contains detailed daily Excel logs for each subject.

* **Naming:** `SubjectName_YYYY-MM-DD.xlsx`.
* **Formatting:** Status cells are bold and color-coded for clarity:

| Status | Color | Font | Description |
| --- | --- | --- | --- |
| **Present** | 🟢 Bright Green | White (Bold) | Arrived on time. |
| **Late** | 🟡 Bright Yellow | Black (Bold) | Arrived after 10-minute threshold. |
| **Absent** | 🔴 Bright Red | White (Bold) | Not detected during the session. |

### 📁 `Attendance_Summary/`

Contains Weekly and Monthly reports for administrative review.

---

## ⚠️ Phase 5: Technical Notes & Restrictions

* **Section Lock:** The system enforces section restrictions (e.g., only "ICT" students are allowed in "Computer System Servicing" sessions).
* **Lighting:** Ensure the room is well-lit. Avoid strong backlighting behind students.
* **Multiple Faces:** A warning is displayed if more than one face is detected to ensure accuracy.
* **Verification Time:** The system records exactly how many seconds the AI took to recognize each student.

```
