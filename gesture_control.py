import sys
import time
import cv2
import os

sys.path.append(os.path.dirname(os.path.abspath(__file__)))

try:
    from gesture_control.powerpoint import minimize_console, initialize_powerpoint, bring_to_foreground, close_powerpoint, check_slideshow_active
    from gesture_control.webcam import initialize_webcam, read_frame, release_webcam
    from gesture_control.gesture import initialize_face_mesh, process_gestures, analyze_performance, set_condition, record_ground_truth
    import mediapipe as mp
except ImportError as e:
    print(f"Import Error: {e}")
    print("Make sure all required files are in the gesture_control/ directory")
    sys.exit(1)

def main():
    print("Starting head gesture control application...")
    
    if len(sys.argv) < 2:
        print("Error: No PowerPoint file path provided.")
        print("Usage: python gesture_control.py <path_to_pptx_file>")
        sys.exit(1)
    
    pptx_path = sys.argv[1]
    print(f"PowerPoint file path: {pptx_path}")
    
    if not os.path.exists(pptx_path):
        print(f"Error: PowerPoint file does not exist: {pptx_path}")
        sys.exit(1)
    
    try:
        minimize_console()
    except Exception as e:
        print(f"Warning: Could not minimize console: {e}")

    try:
        powerpoint, presentation = initialize_powerpoint(pptx_path)
        foreground_success = bring_to_foreground(powerpoint)
        if not foreground_success:
            print("Warning: Could not bring PowerPoint to foreground, but continuing...")
    except Exception as e:
        print(f"Error initializing PowerPoint: {e}")
        sys.exit(1)

    try:
        cap = initialize_webcam(width=1280, height=720)
        mp_face_mesh, face_mesh = initialize_face_mesh()
        mp_drawing = mp.solutions.drawing_utils
        print("Webcam and MediaPipe Face Mesh initialized successfully")
    except Exception as e:
        print(f"Error initializing webcam/MediaPipe: {e}")
        try:
            if 'powerpoint' in locals():
                close_powerpoint(powerpoint, presentation)
        except:
            pass
        sys.exit(1)

    target_fps = 30
    frame_time = 1.0 / target_fps

    print("Starting head gesture detection loop...")
    print("Head gesture controls:")
    print("- Tilt head RIGHT: Next slide")
    print("- Tilt head LEFT: Previous slide") 
    print("- Triple TILT (same direction): Detected (Press ESC to exit)")
    print("- ESC key: Exit application")
    print("Record ground truth: Press R (Tilt Right), L (Tilt Left), T (Triple Tilt) when performing the gesture!")
    print("Make sure your face is clearly visible in the camera!")

    conditions = ["optimal", "low_light", "backlit", "artificial", "natural"]
    condition_idx = 0

    try:
        while True:
            start_time = time.time()

            frame = read_frame(cap)
            if frame is None:
                print("Failed to read frame from webcam")
                break

            current_condition = conditions[condition_idx // 200 % len(conditions)]
            set_condition(current_condition)
            
            try:
                frame, head_detected, exit_detected, delay = process_gestures(
                    frame, face_mesh, mp_drawing, mp.solutions.face_mesh, powerpoint
                )
            except Exception as e:
                print(f"Error processing gestures: {e}")
                continue

            cv2.imshow('Head Gesture Control for PowerPoint', frame)

            key = cv2.waitKey(1) & 0xFF
            if key == 27:  # ESC
                print("ESC key pressed. Exiting...")
                break
            elif key == ord('r'):  # Record Tilt Right
                record_ground_truth("tilt_right")
                print("Recorded ground truth: Tilt Right")
            elif key == ord('l'):  # Record Tilt Left
                record_ground_truth("tilt_left")
                print("Recorded ground truth: Tilt Left")
            elif key == ord('t'):  # Record Triple Tilt
                record_ground_truth("triple_tilt")
                print("Recorded ground truth: Triple Tilt")
                
            if exit_detected:
                print("Triple tilt gesture detected. Closing PowerPoint presentation.")
                break

            condition_idx += 1
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
            analyze_performance()  # Print performance analysis
            if cap:
                release_webcam(cap)
            if 'cv2' in globals():
                cv2.destroyAllWindows()
            if 'powerpoint' in locals() and 'presentation' in locals():
                close_powerpoint(powerpoint, presentation)
            if 'face_mesh' in locals():
                face_mesh.close()
            print("Cleanup completed successfully")
        except Exception as e:
            print(f"Error during cleanup: {e}")

if __name__ == "__main__":
    main()