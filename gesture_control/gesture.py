import mediapipe as mp
import cv2
import math
import time
import numpy as np

# Global variables for gesture timing
last_tilt_time = 0
tilt_cooldown = 1.5  # 1.5 second cooldown between tilts
nod_sequence = []
last_nod_time = 0
nod_timeout = 2.0  # 2 seconds timeout for nod sequence

def initialize_face_mesh():
    """Initialize MediaPipe Face Mesh for head tracking."""
    mp_face_mesh = mp.solutions.face_mesh
    face_mesh = mp_face_mesh.FaceMesh(
        max_num_faces=1,
        refine_landmarks=True,
        min_detection_confidence=0.5,
        min_tracking_confidence=0.5
    )
    return mp_face_mesh, face_mesh

def safe_slideshow_control(powerpoint, action):
    """Safely control slideshow with error handling."""
    try:
        # Check if slideshow is active
        if powerpoint.SlideShowWindows.Count == 0:
            print("Warning: No active slideshow window found")
            return False
            
        slideshow = powerpoint.SlideShowWindows(1)
        
        if action == "next":
            slideshow.View.Next()
            print("Next slide")
        elif action == "previous":
            slideshow.View.Previous()
            print("Previous slide")
        elif action == "exit":
            slideshow.View.Exit()
            print("Slideshow exited")
            
        return True
        
    except Exception as e:
        print(f"Error controlling slideshow: {e}")
        return False

def calculate_head_pose(landmarks, image_size):
    """Calculate head pose from face landmarks."""
    # Key facial landmarks for head pose estimation
    nose_tip = landmarks[1]  # Nose tip
    chin = landmarks[18]     # Chin
    left_eye_corner = landmarks[33]   # Left eye outer corner
    right_eye_corner = landmarks[263] # Right eye outer corner
    left_mouth = landmarks[61]        # Left mouth corner
    right_mouth = landmarks[291]      # Right mouth corner
    
    # Convert normalized coordinates to pixel coordinates
    h, w = image_size
    nose_tip = (int(nose_tip.x * w), int(nose_tip.y * h))
    chin = (int(chin.x * w), int(chin.y * h))
    left_eye = (int(left_eye_corner.x * w), int(left_eye_corner.y * h))
    right_eye = (int(right_eye_corner.x * w), int(right_eye_corner.y * h))
    left_mouth = (int(left_mouth.x * w), int(left_mouth.y * h))
    right_mouth = (int(right_mouth.x * w), int(right_mouth.y * h))
    
    # Calculate head tilt (roll) - based on eye line angle
    eye_center_x = (left_eye[0] + right_eye[0]) / 2
    eye_center_y = (left_eye[1] + right_eye[1]) / 2
    
    # Calculate angle of eye line
    dx = right_eye[0] - left_eye[0]
    dy = right_eye[1] - left_eye[1]
    roll_angle = math.degrees(math.atan2(dy, dx))
    
    # Calculate head pitch (nod) - based on nose-chin vertical distance
    face_height = abs(chin[1] - nose_tip[1])
    
    # Calculate head yaw (turn) - based on nose position relative to eye center
    nose_offset = nose_tip[0] - eye_center_x
    face_width = abs(right_eye[0] - left_eye[0])
    yaw_ratio = nose_offset / face_width if face_width > 0 else 0
    
    return {
        'roll': roll_angle,
        'pitch': face_height,
        'yaw': yaw_ratio,
        'nose_tip': nose_tip,
        'chin': chin,
        'left_eye': left_eye,
        'right_eye': right_eye,
        'eye_center': (int(eye_center_x), int(eye_center_y))
    }

def detect_head_gestures(head_pose, current_time):
    """Detect head gestures based on head pose."""
    global last_tilt_time, nod_sequence, last_nod_time
    
    roll = head_pose['roll']
    pitch = head_pose['pitch']
    
    gesture_detected = None
    
    # Head tilt detection (roll)
    tilt_threshold = 15  # degrees
    if current_time - last_tilt_time > tilt_cooldown:
        if roll > tilt_threshold:
            gesture_detected = "tilt_right"
            last_tilt_time = current_time
        elif roll < -tilt_threshold:
            gesture_detected = "tilt_left"
            last_tilt_time = current_time
    
    # Head nod detection (pitch changes)
    # Simple nod detection based on face height changes
    if len(nod_sequence) == 0:
        nod_sequence.append(pitch)
        last_nod_time = current_time
    else:
        # Check if enough time has passed since last nod measurement
        if current_time - last_nod_time > 0.3:  # Check every 300ms
            pitch_change = abs(pitch - nod_sequence[-1])
            
            if pitch_change > 15:  # Significant head movement
                nod_sequence.append(pitch)
                last_nod_time = current_time
                
                # Check for nod pattern (up-down-up or down-up-down)
                if len(nod_sequence) >= 3:
                    # Look for alternating pattern in last 3 measurements
                    recent_sequence = nod_sequence[-3:]
                    if ((recent_sequence[0] < recent_sequence[1] > recent_sequence[2]) or
                        (recent_sequence[0] > recent_sequence[1] < recent_sequence[2])):
                        gesture_detected = "double_nod"
                        nod_sequence = []  # Reset sequence
    
    # Reset nod sequence if too much time has passed
    if current_time - last_nod_time > nod_timeout:
        nod_sequence = []
    
    return gesture_detected

def process_gestures(frame, face_mesh, mp_drawing, mp_face_mesh, powerpoint):
    """Process head gestures and control PowerPoint."""
    rgb_frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
    results = face_mesh.process(rgb_frame)
    head_detected = False
    gesture_detected = None
    current_time = time.time()

    if results.multi_face_landmarks:
        for face_landmarks in results.multi_face_landmarks:
            # Draw face mesh (optional - can be commented out for cleaner display)
            mp_drawing.draw_landmarks(
                frame,
                face_landmarks,
                mp_face_mesh.FACEMESH_CONTOURS,
                landmark_drawing_spec=None,
                connection_drawing_spec=mp_drawing.DrawingSpec(color=(0, 255, 0), thickness=1)
            )
            
            # Calculate head pose
            head_pose = calculate_head_pose(face_landmarks.landmark, frame.shape[:2])
            head_detected = True
            
            # Draw head pose indicators
            nose_tip = head_pose['nose_tip']
            eye_center = head_pose['eye_center']
            
            # Draw nose tip
            cv2.circle(frame, nose_tip, 5, (0, 0, 255), -1)
            
            # Draw eye center
            cv2.circle(frame, eye_center, 3, (255, 0, 0), -1)
            
            # Draw head orientation line
            cv2.line(frame, eye_center, nose_tip, (255, 255, 0), 2)
            
            # Detect gestures
            gesture_detected = detect_head_gestures(head_pose, current_time)
            
            # Execute gesture commands
            if gesture_detected == "tilt_right":
                if safe_slideshow_control(powerpoint, "next"):
                    cv2.putText(frame, "HEAD TILT RIGHT - NEXT SLIDE", (10, 60), 
                               cv2.FONT_HERSHEY_SIMPLEX, 1, (0, 255, 0), 2)
                    return frame, True, False, 1.0
                    
            elif gesture_detected == "tilt_left":
                if safe_slideshow_control(powerpoint, "previous"):
                    cv2.putText(frame, "HEAD TILT LEFT - PREVIOUS SLIDE", (10, 60), 
                               cv2.FONT_HERSHEY_SIMPLEX, 1, (255, 0, 0), 2)
                    return frame, True, False, 1.0
                    
            elif gesture_detected == "double_nod":
                if safe_slideshow_control(powerpoint, "exit"):
                    cv2.putText(frame, "DOUBLE NOD - CLOSING PRESENTATION", (10, 60), 
                               cv2.FONT_HERSHEY_SIMPLEX, 1, (0, 0, 255), 2)
                    return frame, True, True, 1.0  # Exit presentation
            
            # Display head pose information
            roll_angle = head_pose['roll']
            cv2.putText(frame, f"Head Tilt: {roll_angle:.1f}°", (10, 30), 
                       cv2.FONT_HERSHEY_SIMPLEX, 0.7, (255, 255, 255), 2)
            
            # Display nod sequence length
            nod_count = len(nod_sequence)
            cv2.putText(frame, f"Nod Sequence: {nod_count}", (10, frame.shape[0] - 160), 
                       cv2.FONT_HERSHEY_SIMPLEX, 0.6, (255, 255, 255), 2)

    # Display instructions
    instructions = [
        "ESC: Exit",
        "Double Nod: Close presentation", 
        "Tilt Right: Next slide",
        "Tilt Left: Previous slide",
        "Keep your head visible in frame"
    ]
    
    for i, instruction in enumerate(instructions):
        cv2.putText(frame, instruction, (10, frame.shape[0] - 130 + i * 25), 
                   cv2.FONT_HERSHEY_SIMPLEX, 0.5, (255, 255, 255), 1)

    # Status display
    if head_detected:
        cv2.putText(frame, "HEAD DETECTED", (10, frame.shape[0] - 30), 
                   cv2.FONT_HERSHEY_SIMPLEX, 0.7, (0, 255, 0), 2)
    else:
        cv2.putText(frame, "NO HEAD DETECTED", (10, frame.shape[0] - 30), 
                   cv2.FONT_HERSHEY_SIMPLEX, 0.7, (0, 0, 255), 2)

    return frame, head_detected, gesture_detected == "double_nod", 0.0