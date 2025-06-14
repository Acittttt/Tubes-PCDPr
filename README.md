# **Gesture Control for PowerPoint Presentations**
 Control your PowerPoint slides with head gestures! This project allows you to navigate through a PowerPoint presentation using intuitive head movements captured via webcam, powered by MediaPipe and a sleek Streamlit interface.

# **Features**
  - Head Gesture Navigation: Use simple head gestures to move to the next/previous slide or close the presentation.
  - Streamlit Interface: Upload your PowerPoint file (.pptx) through an easy-to-use web interface.
  - Real-Time Webcam Feedback: See your head pose tracking in real-time for precise gesture control.
  - Triple Tilt Detection: Advanced gesture recognition for closing presentations with three consecutive head tilts.
  - Optimized Performance: High-resolution webcam support with frame rate control for smooth operation.

# **Prerequisites**
  - Operating System: Windows (required for PowerPoint COM automation via pywin32).
  - Microsoft PowerPoint: Installed on your system.
  - Webcam: Any standard webcam compatible with OpenCV/Python.
  - Python: Version 3.8 or higher.

# **Installation**
### 1. Clone the Repository:
  - git clone https://github.com/Acittttt/Tubes-PCD.git 
  - cd gesture_control

### 2. Create a Virtual Environment (optional but recommended):
  - python -m venv venv
  - source venv/bin/activate  # On Windows: venv\Scripts\activate

### 3. Install Dependencies:
  - pip install -r requirements.txt

# **Usage**
### 1. Run the Application:
  - streamlit run app.py OR python -m streamlit run app.py
  - This will open a web interface in your default browser (usually at http://localhost:8501).

### 2. Upload a PowerPoint File:
  - Click the file uploader to select a .pptx file.
  - The selected file name will be displayed.

### 3. Start Gesture Control:
  - Click the "Start Gesture Control" button.
  - A webcam window will open, and the PowerPoint slideshow will start.

# **Supported Gestures**
  - Tilt Right - Next slide - Tilt your head to the right (≥15°)
  - Tilt Left - Previous slide - Tilt your head to the left (≥15°)
  - Triple Tilt - Close presentation - Tilt your head 3 times in the same direction within 3 seconds (≥20°)

# **Gesture Details**
  - Navigation Tilts: Require at least 15-degree head tilt with 0.8-second cooldown between gestures
  - Triple Tilt: Requires three consecutive tilts of at least 20 degrees in the same direction within 3 seconds
  - Real-time Feedback: The application shows your current head tilt angle and triple tilt progress

# **Project Structure**
 #### project/
 #### ├── app.py                  # Streamlit web interface
 #### ├── gesture_control/        # Core logic for gesture control
 #### │   ├── __init__.py
 #### │   ├── powerpoint.py       # PowerPoint initialization and control
 #### │   ├── webcam.py           # Webcam setup and frame processing
 #### │   ├── gesture.py          # Head gesture detection with MediaPipe Face Mesh
 #### ├── gesture_control.py      # Main script to orchestrate gesture control
 #### ├── requirements.txt        # Python dependencies
 #### ├── README.md               # Project documentation

# **Dependencies**
  - pywin32: For PowerPoint automation on Windows.
  - opencv-python: For webcam access and image processing.
  - mediapipe: For hand tracking and gesture detection.
  - numpy: For numerical computations.
  - streamlit: For the web-based user interface.

# **How It Wroks**
  - Face Detection: MediaPipe Face Mesh detects facial landmarks in real-time
  - Head Pose Estimation: Calculates head orientation using key facial points (eyes, nose, chin)
  - Gesture Recognition:
    - Analyzes head roll angle for tilt detection
    - Tracks sequence of tilts for triple tilt pattern
    - Implements cooldown periods to prevent false positives
  - PowerPoint Control: Uses Windows COM automation to control slideshow navigation

# **Tips for Best Results**
  - 💡 Lighting: Ensure good, even lighting on your face
  - 📹 Positioning: Keep your face centered in the camera frame
  - 🎯 Movement: Make clear, deliberate head movements
  - ⏱️ Timing: Wait for gesture cooldown periods between actions
  - 🎚️ Angle: Navigate with 15° tilts, close with 20° triple tilts

# **Troubleshooting**
  - Webcam not working: Ensure your webcam is connected and not used by another application.
  - PowerPoint errors: Verify that Microsoft PowerPoint is installed and the .pptx file is valid.
  - Laggy performance: Try lowering the webcam resolution in webcam.py (e.g., to 640x480).
  - Gestures not detected:
    - Check lighting conditions
    - Ensure face is clearly visible and centered
    - Verify head tilt angles meet minimum thresholds
    - Make sure movements are deliberate and not too fast
  - Triple tilt not working:
    - Ensure all three tilts are in the same direction
    - Complete the sequence within 3 seconds
    - Use more pronounced head movements (≥20°)

# **Technical Features**
  - Advanced Head Pose Calculation: Uses multiple facial landmarks for accurate head orientation
  - Gesture Sequence Tracking: Sophisticated pattern recognition for triple tilt detection
  - Adaptive Thresholds: Different sensitivity levels for navigation vs. exit gestures
  - Real-time Visual Feedback: Live display of head pose angles and gesture progress
  - Error Handling: Robust error recovery and resource cleanup

# **Created with ❤️ by [Annisa Dian - Rasyid Raafi - Rindi Indriani]**
