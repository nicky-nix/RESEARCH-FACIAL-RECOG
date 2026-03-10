
# Create a requirements.txt file for dependencies
requirements = """# Fast Facial Recognition Attendance System
# Required Python packages

opencv-python==4.8.1.78
face-recognition==1.3.0
numpy==1.24.3
pandas==2.0.3
openpyxl==3.1.2
dlib==19.24.2

# Additional dependencies (auto-installed with face-recognition)
# - cmake
# - face_recognition_models
# - Pillow
# - scipy
"""

with open('requirements.txt', 'w') as f:
    f.write(requirements)

print("✓ Requirements file created: requirements.txt")
