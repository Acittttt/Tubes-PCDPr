# **Gesture Control for PowerPoint Presentations**
### Control your PowerPoint slides with hand gestures! This project allows you to navigate through a PowerPoint presentation using intuitive hand gestures captured via webcam, powered by MediaPipe and a sleek Streamlit interface.

# **Features**
## - Hand Gesture Navigation: Use simple hand gestures to move to the next/previous slide or close the presentation.
## - Streamlit Interface: Upload your PowerPoint file (.pptx) through an easy-to-use web interface.
## - Real-Time Webcam Feedback: See your hand landmarks in real-time for precise gesture control.
## - Optimized Performance: Reduced webcam resolution and frame rate control for smooth operation.

# **Prerequisites**
## - Operating System: Windows (required for PowerPoint COM automation via pywin32).
## - Microsoft PowerPoint: Installed on your system.
## - Webcam: Any standard webcam compatible with OpenCV/Python.
## - Python: Version 3.8 or higher.

# **Installation**
## 1. Clone the Repository:
### git clone https://github.com/Acittttt/Tubes-PCD.git 
### cd GestureControlIPPT

## 2. Create a Virtual Environment (optional but recommended):
### python -m venv venv
### source venv/bin/activate  # On Windows: venv\Scripts\activate

## 3. Install Dependencies:
### pip install -r requirements.txt

# **Usage**
## 1. Run the Application:
### streamlit run app.py
### This will open a web interface in your default browser (usually at http://localhost:8501).

## 2. Upload a PowerPoint File:
### - Click the file uploader to select a .pptx file.
### - The selected file name will be displayed.

## 3. Start Gesture Control:
### - Click the "Start Gesture Control" button.
### - A webcam window will open, and the PowerPoint slideshow will start.

# Supported Gestures
## Gesture
### Pinch - Thumb and index finger close together - Close the presentation
### Swipe Right - Thumb to the right of index finger - Go to the next slide
### Swipe Left - Thumb to the left of index finger - Go to the previous slide

# **Project Structure**

### project/
### ├── app.py                  # Streamlit web interface
### ├── gesture_control/        # Core logic for gesture control
### │   ├── __init__.py
### │   ├── powerpoint.py       # PowerPoint initialization and control
### │   ├── webcam.py           # Webcam setup and frame processing
### │   ├── gesture.py          # Hand gesture detection with MediaPipe
### ├── gesture_control.py      # Main script to orchestrate gesture control
### ├── requirements.txt        # Python dependencies
### ├── README.md               # Project documentation

# **Dependencies**
## - pywin32: For PowerPoint automation on Windows.
## - opencv-python: For webcam access and image processing.
## - mediapipe: For hand tracking and gesture detection.
## - numpy: For numerical computations.
## - streamlit: For the web-based user interface.

# **Troubleshooting**
## - Webcam not working: Ensure your webcam is connected and not used by another application.
## - PowerPoint errors: Verify that Microsoft PowerPoint is installed and the .pptx file is valid.
## - Laggy performance: Try lowering the webcam resolution in webcam.py (e.g., to 320x240).
## - Gestures not detected: Adjust the pinch threshold in gesture.py or ensure good lighting for webcam capture.

# **Created with ❤️ by [Annisa Dian - Rasyid Raafi - Rindi Indriani]**
