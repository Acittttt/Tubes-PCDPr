import cv2

def initialize_webcam(width=640, height=480, fps=30):
    """Initialize webcam with specified resolution and frame rate."""
    cap = cv2.VideoCapture(0)
    cap.set(cv2.CAP_PROP_FRAME_WIDTH, width)
    cap.set(cv2.CAP_PROP_FRAME_HEIGHT, height)
    cap.set(cv2.CAP_PROP_FPS, fps)
    if not cap.isOpened():
        print("Error: Could not open webcam.")
        raise RuntimeError("Webcam initialization failed.")
    return cap

def read_frame(cap):
    """Read and preprocess a frame from the webcam."""
    ret, frame = cap.read()
    if not ret:
        print("Error: Failed to capture image.")
        return None
    frame = cv2.flip(frame, 1)
    return frame

def release_webcam(cap):
    """Release the webcam resource."""
    if cap:
        cap.release()