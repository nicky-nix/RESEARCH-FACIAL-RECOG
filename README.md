# RESEARCH-FACIAL-RECOG
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
Open your terminal or Command Prompt (CMD) and run:
```bash
pip install opencv-python face_recognition numpy pandas openpyxl
