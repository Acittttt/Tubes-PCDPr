import mediapipe as mp
import cv2
import math
import time

# Global variables for gesture timing
last_swipe_time = 0
swipe_cooldown = 1.0  # 1 second cooldown between swipes

def initialize_hands():
    """Initialize MediaPipe Hands for hand tracking."""
    mp_hands = mp.solutions.hands
    hands = mp_hands.Hands(
        min_detection_confidence=0.7,
        min_tracking_confidence=0.5,
        max_num_hands=1
    )
    return mp_hands, hands

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

def process_gestures(frame, hands, mp_drawing, mp_hands, powerpoint):
    """Process hand gestures and control PowerPoint."""
    global last_swipe_time
    
    rgb_frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
    results = hands.process(rgb_frame)
    hand_detected = False
    pinch_detected = False
    current_time = time.time()

    if results.multi_hand_landmarks:
        for hand_landmarks in results.multi_hand_landmarks:
            # Draw hand landmarks
            mp_drawing.draw_landmarks(
                frame, 
                hand_landmarks, 
                mp_hands.HAND_CONNECTIONS,
                mp_drawing.DrawingSpec(color=(0, 0, 255), thickness=2, circle_radius=2),
                mp_drawing.DrawingSpec(color=(0, 255, 0), thickness=2)
            )

            # Get landmark positions
            thumb_tip = hand_landmarks.landmark[mp_hands.HandLandmark.THUMB_TIP]
            index_tip = hand_landmarks.landmark[mp_hands.HandLandmark.INDEX_FINGER_TIP]
            middle_tip = hand_landmarks.landmark[mp_hands.HandLandmark.MIDDLE_FINGER_TIP]
            ring_tip = hand_landmarks.landmark[mp_hands.HandLandmark.RING_FINGER_TIP]
            pinky_tip = hand_landmarks.landmark[mp_hands.HandLandmark.PINKY_TIP]

            # Pinch Gesture Detection (thumb and index finger close together)
            thumb_index_distance = math.sqrt(
                (thumb_tip.x - index_tip.x) ** 2 +
                (thumb_tip.y - index_tip.y) ** 2 +
                (thumb_tip.z - index_tip.z) ** 2
            )

            pinch_threshold = 0.04
            if thumb_index_distance < pinch_threshold:
                pinch_detected = True
                cv2.putText(frame, "PINCH DETECTED - CLOSING", (10, 30), 
                           cv2.FONT_HERSHEY_SIMPLEX, 1, (0, 0, 255), 2)

            # Swipe Gestures (only if cooldown period has passed)
            if current_time - last_swipe_time > swipe_cooldown:
                # Check if hand is in a good position for swiping
                # All fingertips should be roughly at the same height (extended hand)
                finger_tips = [thumb_tip, index_tip, middle_tip, ring_tip, pinky_tip]
                avg_y = sum(tip.y for tip in finger_tips) / len(finger_tips)
                
                # Check if fingers are extended (y positions are similar)
                fingers_extended = all(abs(tip.y - avg_y) < 0.1 for tip in finger_tips)
                
                if fingers_extended:
                    # Swipe Right (Next slide) - hand moves from left to right
                    if thumb_tip.x < index_tip.x and index_tip.x < middle_tip.x:
                        if safe_slideshow_control(powerpoint, "next"):
                            last_swipe_time = current_time
                            cv2.putText(frame, "SWIPE RIGHT - NEXT SLIDE", (10, 60), 
                                       cv2.FONT_HERSHEY_SIMPLEX, 1, (0, 255, 0), 2)
                            return frame, True, False, 1.0  # Swipe detected, delay 1s
                    
                    # Swipe Left (Previous slide) - hand moves from right to left  
                    elif thumb_tip.x > index_tip.x and index_tip.x > middle_tip.x:
                        if safe_slideshow_control(powerpoint, "previous"):
                            last_swipe_time = current_time
                            cv2.putText(frame, "SWIPE LEFT - PREVIOUS SLIDE", (10, 60), 
                                       cv2.FONT_HERSHEY_SIMPLEX, 1, (255, 0, 0), 2)
                            return frame, True, False, 1.0  # Swipe detected, delay 1s

            hand_detected = True
            
            # Display gesture status
            if hand_detected and not pinch_detected:
                cv2.putText(frame, "HAND DETECTED", (10, frame.shape[0] - 30), 
                           cv2.FONT_HERSHEY_SIMPLEX, 0.7, (255, 255, 0), 2)

    # Display instructions
    instructions = [
        "ESC: Exit",
        "Pinch: Close presentation", 
        "Swipe Right: Next slide",
        "Swipe Left: Previous slide"
    ]
    
    for i, instruction in enumerate(instructions):
        cv2.putText(frame, instruction, (10, frame.shape[0] - 120 + i * 25), 
                   cv2.FONT_HERSHEY_SIMPLEX, 0.5, (255, 255, 255), 1)

    return frame, hand_detected, pinch_detected, 0.0