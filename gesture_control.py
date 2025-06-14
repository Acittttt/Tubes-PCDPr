import sys
import time
import cv2
import os

# Add the current directory to Python path for imports
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

try:
    from gesture_control.powerpoint import minimize_console, initialize_powerpoint, bring_to_foreground, close_powerpoint, check_slideshow_active
    from gesture_control.webcam import initialize_webcam, read_frame, release_webcam
    from gesture_control.gesture import initialize_hands, process_gestures
    import mediapipe as mp
except ImportError as e:
    print(f"Import Error: {e}")
    print("Make sure all required files are in the gesture_control/ directory")
    sys.exit(1)

def main():
    print("Starting gesture control application...")
    
    # Get PowerPoint file path from arguments
    if len(sys.argv) < 2:
        print("Error: No PowerPoint file path provided.")
        print("Usage: python gesture_control.py <path_to_pptx_file>")
        sys.exit(1)
    
    pptx_path = sys.argv[1]
    print(f"PowerPoint file path: {pptx_path}")
    
    # Check if file exists
    if not os.path.exists(pptx_path):
        print(f"Error: PowerPoint file does not exist: {pptx_path}")
        sys.exit(1)
    
    # Minimize console
    try:
        minimize_console()
    except Exception as e:
        print(f"Warning: Could not minimize console: {e}")

    # Initialize PowerPoint
    try:
        powerpoint, presentation = initialize_powerpoint(pptx_path)
        # Try to bring to foreground, but don't fail if it doesn't work
        foreground_success = bring_to_foreground(powerpoint)
        if not foreground_success:
            print("Warning: Could not bring PowerPoint to foreground, but continuing...")
    except Exception as e:
        print(f"Error initializing PowerPoint: {e}")
        sys.exit(1)

    # Initialize webcam and MediaPipe
    try:
        cap = initialize_webcam()
        mp_hands, hands = initialize_hands()
        mp_drawing = mp.solutions.drawing_utils
        print("Webcam and MediaPipe initialized successfully")
    except Exception as e:
        print(f"Error initializing webcam/MediaPipe: {e}")
        # Clean up PowerPoint before exiting
        try:
            close_powerpoint(powerpoint, presentation)
        except:
            pass
        sys.exit(1)

    # Frame rate control
    target_fps = 30
    frame_time = 1.0 / target_fps

    print("Starting gesture detection loop...")
    print("Use ESC key to exit or pinch gesture to close")

    try:
        while True:
            start_time = time.time()

            frame = read_frame(cap)
            if frame is None:
                print("Failed to read frame from webcam")
                break

            # Process gestures
            try:
                frame, hand_detected, pinch_detected, delay = process_gestures(
                    frame, hands, mp_drawing, mp_hands, powerpoint
                )
            except Exception as e:
                print(f"Error processing gestures: {e}")
                # Continue the loop instead of breaking
                continue

            # Show the frame
            cv2.imshow('Hand Gesture Control', frame)

            # Check for exit conditions
            key = cv2.waitKey(1) & 0xFF
            if pinch_detected:
                print("Pinch gesture detected. Closing PowerPoint presentation.")
                break

            # Control frame rate
            elapsed_time = time.time() - start_time
            sleep_time = max(0, frame_time - elapsed_time - delay)
            if sleep_time > 0:
                time.sleep(sleep_time)

    except KeyboardInterrupt:
        print("Interrupted by user")
    except Exception as e:
        print(f"Error during execution: {e}")

    finally:
        print("Cleaning up...")
        try:
            release_webcam(cap)
            cv2.destroyAllWindows()
            close_powerpoint(powerpoint, presentation)
            hands.close()
            print("Cleanup completed successfully")
        except Exception as e:
            print(f"Error during cleanup: {e}")

if __name__ == "__main__":
    main()