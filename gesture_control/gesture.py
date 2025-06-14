import mediapipe as mp
import cv2
import math
import time
import numpy as np

# Global variables for gesture timing
last_tilt_time = 0
tilt_cooldown = 0.8  # Reduced cooldown for triple tilt
nod_sequence = []
last_nod_time = 0
nod_timeout = 2.0

# New variables for triple tilt detection
triple_tilt_sequence = []
last_triple_tilt_time = 0
triple_tilt_timeout = 3.0  # 3 seconds to complete triple tilt
triple_tilt_threshold = 20  # degrees, slightly higher threshold

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
    nose_tip = landmarks[1]
    chin = landmarks[18]
    left_eye_corner = landmarks[33]
    right_eye_corner = landmarks[263]
    left_mouth = landmarks[61]
    right_mouth = landmarks[291]
    
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

def detect_triple_tilt(roll_angle, current_time):
    """Detect triple head tilt gesture for closing presentation."""
    global triple_tilt_sequence, last_triple_tilt_time
    
    # Reset sequence if timeout
    if current_time - last_triple_tilt_time > triple_tilt_timeout:
        triple_tilt_sequence = []
    
    # Detect significant tilt
    if abs(roll_angle) > triple_tilt_threshold:
        tilt_direction = "right" if roll_angle > 0 else "left"
        
        # Check if enough time has passed since last tilt detection (prevent duplicate detection)
        if len(triple_tilt_sequence) == 0 or current_time - last_triple_tilt_time > 0.5:
            triple_tilt_sequence.append({
                'direction': tilt_direction,
                'angle': roll_angle,
                'time': current_time
            })
            last_triple_tilt_time = current_time
            
            print(f"Triple tilt progress: {len(triple_tilt_sequence)}/3 - Direction: {tilt_direction}")
            
            # Check if we have 3 tilts in the same direction
            if len(triple_tilt_sequence) >= 3:
                # Verify all tilts are in the same direction
                recent_tilts = triple_tilt_sequence[-3:]
                directions = [tilt['direction'] for tilt in recent_tilts]
                
                if all(direction == directions[0] for direction in directions):
                    # Check timing - all tilts should be within timeout period
                    time_span = recent_tilts[-1]['time'] - recent_tilts[0]['time']
                    if time_span <= triple_tilt_timeout:
                        print(f"TRIPLE TILT DETECTED! Direction: {directions[0]}")
                        triple_tilt_sequence = []  # Reset
                        return True
                
                # If we have more than 3 tilts but they don't match pattern, keep only recent ones
                if len(triple_tilt_sequence) > 3:
                    triple_tilt_sequence = triple_tilt_sequence[-2:]
    
    return False

def detect_head_gestures(head_pose, current_time):
    """Detect head gestures based on head pose."""
    global last_tilt_time
    
    roll = head_pose['roll']
    gesture_detected = None
    
    # Check for triple tilt first (for closing presentation)
    if detect_triple_tilt(roll, current_time):
        return "triple_tilt"
    
    # Regular head tilt detection for navigation (with cooldown to prevent conflict)
    tilt_threshold = 15  # degrees
    if current_time - last_tilt_time > tilt_cooldown:
        if roll > tilt_threshold:
            gesture_detected = "tilt_right"
            last_tilt_time = current_time
        elif roll < -tilt_threshold:
            gesture_detected = "tilt_left"
            last_tilt_time = current_time
    
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
            # Draw face mesh (optional)
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
            
            # Draw nose tip and eye center
            cv2.circle(frame, nose_tip, 5, (0, 0, 255), -1)
            cv2.circle(frame, eye_center, 3, (255, 0, 0), -1)
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
                    
            elif gesture_detected == "triple_tilt":
                if safe_slideshow_control(powerpoint, "exit"):
                    cv2.putText(frame, "TRIPLE TILT - CLOSING PRESENTATION", (10, 60), 
                               cv2.FONT_HERSHEY_SIMPLEX, 1, (0, 0, 255), 2)
                    return frame, True, True, 2.0  # Longer delay for exit
            
            # Display head pose information
            roll_angle = head_pose['roll']
            cv2.putText(frame, f"Head Tilt: {roll_angle:.1f}°", (10, 30), 
                       cv2.FONT_HERSHEY_SIMPLEX, 0.7, (255, 255, 255), 2)
            
            # Display triple tilt progress
            triple_tilt_count = len(triple_tilt_sequence)
            cv2.putText(frame, f"Triple Tilt Progress: {triple_tilt_count}/3", (10, frame.shape[0] - 160), 
                       cv2.FONT_HERSHEY_SIMPLEX, 0.6, (255, 255, 0), 2)

    # Display instructions
    instructions = [
        "ESC: Exit",
        "Triple Tilt (same direction): Close presentation", 
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

    return frame, head_detected, gesture_detected == "triple_tilt", 0.0