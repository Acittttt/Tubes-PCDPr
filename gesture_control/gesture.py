import mediapipe as mp
import cv2
import math

def initialize_hands():
    """Initialize MediaPipe Hands for hand tracking."""
    mp_hands = mp.solutions.hands
    hands = mp_hands.Hands(
        min_detection_confidence=0.7,
        min_tracking_confidence=0.5,
        max_num_hands=1
    )
    return mp_hands, hands

def process_gestures(frame, hands, mp_drawing, mp_hands, powerpoint):
    """Process hand gestures and control PowerPoint."""
    rgb_frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
    results = hands.process(rgb_frame)
    hand_detected = False
    pinch_detected = False

    if results.multi_hand_landmarks:
        for hand_landmarks in results.multi_hand_landmarks:
            mp_drawing.draw_landmarks(frame, hand_landmarks, mp_hands.HAND_CONNECTIONS)

            # Pinch Gesture Detection
            thumb_tip = hand_landmarks.landmark[mp_hands.HandLandmark.THUMB_TIP]
            index_tip = hand_landmarks.landmark[mp_hands.HandLandmark.INDEX_FINGER_TIP]

            thumb_index_distance = math.sqrt(
                (thumb_tip.x - index_tip.x) ** 2 +
                (thumb_tip.y - index_tip.y) ** 2 +
                (thumb_tip.z - index_tip.z) ** 2
            )

            pinch_threshold = 0.03
            if thumb_index_distance < pinch_threshold:
                pinch_detected = True

            # Swipe Gestures
            thumb_tip_x = thumb_tip.x
            index_tip_x = index_tip.x
            if thumb_tip_x > index_tip_x:
                powerpoint.SlideShowWindows(1).View.Next()
                return frame, True, False, 1.0  # Swipe detected, delay 1s
            elif index_tip_x > thumb_tip_x:
                powerpoint.SlideShowWindows(1).View.Previous()
                return frame, True, False, 1.0  # Swipe detected, delay 1s

            hand_detected = True

    return frame, hand_detected, pinch_detected, 0.0