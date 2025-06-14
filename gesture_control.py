import sys
import time
import cv2
from gesture_control.powerpoint import minimize_console, initialize_powerpoint, bring_to_foreground, close_powerpoint
from gesture_control.webcam import initialize_webcam, read_frame, release_webcam
from gesture_control.gesture import initialize_hands, process_gestures

def main():
    # Minimize console
    minimize_console()

    # Initialize PowerPoint
    pptx_path = sys.argv[1] if len(sys.argv) > 1 else None
    powerpoint, presentation = initialize_powerpoint(pptx_path)
    bring_to_foreground(powerpoint)

    # Initialize webcam and MediaPipe
    cap = initialize_webcam()
    mp_hands, hands = initialize_hands()
    mp_drawing = mp.solutions.drawing_utils

    # Frame rate control
    target_fps = 30
    frame_time = 1.0 / target_fps

    try:
        while True:
            start_time = time.time()

            frame = read_frame(cap)
            if frame is None:
                break

            # Process gestures
            frame, hand_detected, pinch_detected, delay = process_gestures(
                frame, hands, mp_drawing, mp_hands, powerpoint
            )

            cv2.imshow('Hand Gesture Control', frame)

            if cv2.waitKey(1) & 0xFF == 27:  # Esc key to exit
                break

            if pinch_detected:
                print("Closing PowerPoint presentation.")
                break

            # Control frame rate
            elapsed_time = time.time() - start_time
            sleep_time = max(0, frame_time - elapsed_time - delay)
            time.sleep(sleep_time)

    except Exception as e:
        print(f"Error during execution: {e}")

    finally:
        release_webcam(cap)
        cv2.destroyAllWindows()
        close_powerpoint(powerpoint, presentation)
        hands.close()

if __name__ == "__main__":
    main()