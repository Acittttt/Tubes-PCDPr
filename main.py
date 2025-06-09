import ctypes
import win32com.client
import win32gui
import win32con
import time
import mediapipe as mp
import cv2
import math
import sys

# Minimize the command prompt window
kernel32 = ctypes.WinDLL('kernel32')
user32 = ctypes.WinDLL('user32')
hWnd = kernel32.GetConsoleWindow()
if hWnd:
    user32.ShowWindow(hWnd, 6)  # 6 = Minimize window

# Initialize PowerPoint application
try:
    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    pptx_path = sys.argv[1] if len(sys.argv) > 1 else r'D:\PCD\Tubes\Kelompok3_Ragam_dan_Kesantunan_Bahasa.pptx'
    presentation = powerpoint.Presentations.Open(pptx_path)
    presentation.SlideShowSettings.Run()
    time.sleep(3)  # Tunggu 3 detik untuk memastikan jendela slideshow muncul
except Exception as e:
    print(f"Error opening PowerPoint: {e}")
    exit(1)

# Bring PowerPoint window to the foreground
hwnd = powerpoint.SlideShowWindows(1).HWND  # Dapatkan handle jendela slideshow
if hwnd:
    try:
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        win32gui.SetForegroundWindow(hwnd)
    except Exception as e:
        print(f"Error setting PowerPoint window to foreground: {e}")
else:
    print("Error: Could not get PowerPoint slideshow window handle.")
    exit(1)

# Initialize MediaPipe for hand tracking
mp_hands = mp.solutions.hands
hands = mp_hands.Hands(min_detection_confidence=0.7, min_tracking_confidence=0.5)
mp_drawing = mp.solutions.drawing_utils

# Open a webcam feed
cap = cv2.VideoCapture(0)
if not cap.isOpened():
    print("Error: Could not open webcam.")
    exit(1)

try:
    # Flag to track whether to close PowerPoint
    close_presentation = False

    while True:
        ret, frame = cap.read()
        if not ret:
            print("Error: Failed to capture image.")
            break

        frame = cv2.flip(frame, 1)
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

                print(f"Thumb-Index Distance: {thumb_index_distance}")

                pinch_threshold = 0.03
                if thumb_index_distance < pinch_threshold:
                    print("Pinch Gesture Detected")
                    pinch_detected = True

                thumb_tip_x = hand_landmarks.landmark[mp_hands.HandLandmark.THUMB_TIP].x
                index_tip_x = hand_landmarks.landmark[mp_hands.HandLandmark.INDEX_FINGER_TIP].x
                if thumb_tip_x > index_tip_x:
                    print("Swipe Right Gesture Detected: Next Slide")
                    powerpoint.SlideShowWindows(1).View.Next()
                    time.sleep(1)
                elif index_tip_x > thumb_tip_x:
                    print("Swipe Left Gesture Detected: Previous Slide")
                    powerpoint.SlideShowWindows(1).View.Previous()
                    time.sleep(1)

                hand_detected = True

        cv2.imshow('Hand Gesture Control', frame)

        if cv2.waitKey(1) & 0xFF == 27:  # Esc key to exit
            break

        if pinch_detected:
            print("Closing PowerPoint presentation.")
            close_presentation = True
            break

    if close_presentation or not hand_detected:
        print("Exiting program.")
except Exception as e:
    print(f"Error during execution: {e}")

finally:
    cap.release()
    cv2.destroyAllWindows()
    if 'presentation' in locals():
        presentation.Close()
    if 'powerpoint' in locals():
        powerpoint.Quit()
    if 'hands' in locals():
        hands.close()