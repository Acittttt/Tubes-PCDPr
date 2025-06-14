import mediapipe as mp
import cv2
import math
import time
import numpy as np
from collections import defaultdict

# Global variables for gesture timing and performance tracking
last_tilt_time = 0
tilt_cooldown = 0.8
triple_tilt_sequence = []
last_triple_tilt_time = 0
triple_tilt_timeout = 3.0
triple_tilt_threshold = 20
performance_data = defaultdict(list)  # Store performance metrics
condition = "optimal"  # Current lighting condition
ground_truth = []  # Store expected gestures

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
    """Safely control slideshow with error handling and measure latency."""
    try:
        if powerpoint.SlideShowWindows.Count == 0:
            return False, 0.0
        slideshow = powerpoint.SlideShowWindows(1)
        start_time = time.time()
        if action == "next":
            slideshow.View.Next()
        elif action == "previous":
            slideshow.View.Previous()
        elif action == "exit":
            slideshow.View.Exit()
        latency = time.time() - start_time
        return True, max(latency, 0.001)  # Minimum latency to avoid zero
    except Exception:
        return False, 0.0

def calculate_head_pose(landmarks, image_size):
    """Calculate head pose from face landmarks."""
    nose_tip = landmarks[1]
    left_eye_corner = landmarks[33]
    right_eye_corner = landmarks[263]
    
    h, w = image_size
    nose_tip = (int(nose_tip.x * w), int(nose_tip.y * h))
    left_eye = (int(left_eye_corner.x * w), int(left_eye_corner.y * h))
    right_eye = (int(right_eye_corner.x * w), int(right_eye_corner.y * h))
    
    dx = right_eye[0] - left_eye[0]
    dy = right_eye[1] - left_eye[1]
    roll_angle = math.degrees(math.atan2(dy, dx))
    
    return {'roll': roll_angle, 'nose_tip': nose_tip, 'left_eye': left_eye, 'right_eye': right_eye}

def detect_triple_tilt(roll_angle, current_time):
    """Detect triple head tilt gesture for closing presentation."""
    global triple_tilt_sequence, last_triple_tilt_time
    if current_time - last_triple_tilt_time > triple_tilt_timeout:
        triple_tilt_sequence = []
    
    if abs(roll_angle) > triple_tilt_threshold:
        tilt_direction = "right" if roll_angle > 0 else "left"
        if len(triple_tilt_sequence) == 0 or current_time - last_triple_tilt_time > 0.5:
            triple_tilt_sequence.append({
                'direction': tilt_direction,
                'angle': roll_angle,
                'time': current_time
            })
            last_triple_tilt_time = current_time
            if len(triple_tilt_sequence) >= 3:
                recent_tilts = triple_tilt_sequence[-3:]
                directions = [tilt['direction'] for tilt in recent_tilts]
                if all(direction == directions[0] for direction in directions):
                    time_span = recent_tilts[-1]['time'] - recent_tilts[0]['time']
                    if time_span <= triple_tilt_timeout:
                        triple_tilt_sequence = []
                        return True
    return False

def detect_head_gestures(head_pose, current_time):
    """Detect head gestures based on head pose."""
    global last_tilt_time
    roll = head_pose['roll']
    gesture_detected = None
    
    if detect_triple_tilt(roll, current_time):
        return "triple_tilt"
    
    tilt_threshold = 15
    if current_time - last_tilt_time > tilt_cooldown:
        if roll > tilt_threshold:
            gesture_detected = "tilt_right"
            last_tilt_time = current_time
        elif roll < -tilt_threshold:
            gesture_detected = "tilt_left"
            last_tilt_time = current_time
    
    return gesture_detected

def set_condition(new_condition):
    """Set the current lighting condition for performance tracking."""
    global condition
    condition = new_condition

def record_ground_truth(gesture):
    """Record an expected gesture as ground truth."""
    global ground_truth
    ground_truth.append({
        'gesture': gesture,
        'timestamp': time.time(),
        'condition': condition
    })

def process_gestures(frame, face_mesh, mp_drawing, mp_face_mesh, powerpoint):
    """Process head gestures, control PowerPoint, and collect performance metrics."""
    global performance_data, condition, ground_truth
    rgb_frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
    results = face_mesh.process(rgb_frame)
    head_detected = False
    gesture_detected = None
    current_time = time.time()
    start_time = time.time()

    if results.multi_face_landmarks:
        for face_landmarks in results.multi_face_landmarks:
            mp_drawing.draw_landmarks(
                frame, face_landmarks, mp_face_mesh.FACEMESH_CONTOURS,
                landmark_drawing_spec=None,
                connection_drawing_spec=mp_drawing.DrawingSpec(color=(0, 255, 0), thickness=1)
            )
            head_pose = calculate_head_pose(face_landmarks.landmark, frame.shape[:2])
            head_detected = True
            gesture_detected = detect_head_gestures(head_pose, current_time)
            
            # Find the closest ground truth gesture within a time window (e.g., 1 second)
            expected_gesture = None
            for gt in ground_truth:
                if abs(gt['timestamp'] - current_time) < 1.0 and gt['condition'] == condition:
                    expected_gesture = gt['gesture']
                    break
            
            # Record performance metrics
            if gesture_detected:
                success, latency = safe_slideshow_control(powerpoint, {
                    "tilt_right": "next",
                    "tilt_left": "previous",
                    "triple_tilt": "exit"
                }.get(gesture_detected, ""))
                if success:
                    is_correct = gesture_detected == expected_gesture if expected_gesture else True  # Default to True if no ground truth
                    performance_data[condition].append({
                        'gesture': gesture_detected,
                        'latency': latency,
                        'correct': is_correct,
                        'timestamp': current_time
                    })
                
                # Execute gesture commands
                if gesture_detected == "tilt_right" and success:
                    cv2.putText(frame, "HEAD TILT RIGHT - NEXT SLIDE", (10, 60), 
                               cv2.FONT_HERSHEY_SIMPLEX, 1, (0, 255, 0), 2)
                    return frame, True, False, 1.0
                elif gesture_detected == "tilt_left" and success:
                    cv2.putText(frame, "HEAD TILT LEFT - PREVIOUS SLIDE", (10, 60), 
                               cv2.FONT_HERSHEY_SIMPLEX, 1, (255, 0, 0), 2)
                    return frame, True, False, 1.0
                elif gesture_detected == "triple_tilt" and success:
                    cv2.putText(frame, "TRIPLE TILT DETECTED (Press ESC to exit)", (10, 60), 
                               cv2.FONT_HERSHEY_SIMPLEX, 1, (0, 0, 255), 2)
                    return frame, True, False, 2.0  # Do not exit immediately
            
            cv2.putText(frame, f"Head Tilt: {head_pose['roll']:.1f}°", (10, 30), 
                       cv2.FONT_HERSHEY_SIMPLEX, 0.7, (255, 255, 255), 2)
            cv2.putText(frame, f"Condition: {condition}", (10, 60), 
                       cv2.FONT_HERSHEY_SIMPLEX, 0.7, (255, 255, 255), 2)

    instructions = [
        "ESC: Exit", "Triple Tilt: Detected (manual exit)", 
        "Tilt Right: Next slide", "Tilt Left: Previous slide", 
        "Keep head visible", "R: Record Tilt Right", "L: Record Tilt Left", "T: Record Triple Tilt"
    ]
    for i, instruction in enumerate(instructions):
        cv2.putText(frame, instruction, (10, frame.shape[0] - 130 + i * 25), 
                   cv2.FONT_HERSHEY_SIMPLEX, 0.5, (255, 255, 255), 1)
    
    cv2.putText(frame, "HEAD DETECTED" if head_detected else "NO HEAD DETECTED", 
                (10, frame.shape[0] - 30), cv2.FONT_HERSHEY_SIMPLEX, 0.7, 
                (0, 255, 0) if head_detected else (0, 0, 255), 2)
    
    return frame, head_detected, False, 0.0  # exit_detected always False unless ESC is pressed

def analyze_performance():
    """Analyze performance data and print results."""
    conditions = ["optimal", "low_light", "backlit", "artificial", "natural"]
    gesture_types = ["tilt_right", "tilt_left", "triple_tilt"]
    
    print("\nPerformance Analysis")
    print("| Gesture Type | Optimal Light | Low Light | Backlit | Artificial | Natural | Overall Accuracy |")
    print("|--------------|---------------|-----------|---------|------------|---------|------------------|")
    
    for gesture in gesture_types:
        accuracies = []
        for cond in conditions:
            data = performance_data.get(cond, [])
            gesture_data = [d for d in data if d['gesture'] == gesture]
            if not gesture_data:
                accuracy = 0.0
            else:
                correct = sum(1 for d in gesture_data if d['correct'])
                accuracy = (correct / len(gesture_data)) * 100 if len(gesture_data) > 0 else 0.0
            accuracies.append(accuracy if accuracy > 0 else 0.0)
        
        overall_accuracy = sum(accuracies) / len(accuracies) if any(accuracies) else 0.0
        print(f"| {gesture:12} | {accuracies[0]:>12.1f}% | {accuracies[1]:>9.1f}% | "
              f"{accuracies[2]:>7.1f}% | {accuracies[3]:>10.1f}% | {accuracies[4]:>7.1f}% | "
              f"{overall_accuracy:>15.1f}% |")
    
    print("\nLatency:")
    for gesture in gesture_types:
        latencies = []
        for cond in conditions:
            latencies.extend([d['latency'] for d in performance_data.get(cond, []) if d['gesture'] == gesture])
        if latencies:
            min_latency, max_latency = min(latencies), max(latencies)
            print(f"{gesture}: {min_latency:.1f}–{max_latency:.1f} seconds")
        else:
            print(f"{gesture}: 0.0–0.0 seconds")
    
    print("\nFalse Positive Rate (FPR):")
    for gesture in gesture_types:
        false_positives = 0
        total = 0
        for cond in conditions:
            data = performance_data.get(cond, [])
            gesture_data = [d for d in data if d['gesture'] == gesture]
            false_positives += sum(1 for d in gesture_data if not d['correct'])
            total += len(gesture_data)
        fpr = (false_positives / total * 100) if total > 0 else 0.0
        print(f"{gesture}: <{fpr:.1f}%")