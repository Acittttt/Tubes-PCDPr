import cv2

def initialize_webcam(width=1280, height=720, fps=30):
    """Initialize webcam with specified resolution and frame rate for optimal face detection."""
    print(f"Initializing webcam with resolution: {width}x{height} at {fps} FPS")
    
    cap = cv2.VideoCapture(0)
    
    # Set camera properties
    cap.set(cv2.CAP_PROP_FRAME_WIDTH, width)
    cap.set(cv2.CAP_PROP_FRAME_HEIGHT, height)
    cap.set(cv2.CAP_PROP_FPS, fps)
    
    # Additional settings for better face detection
    cap.set(cv2.CAP_PROP_AUTO_EXPOSURE, 0.25)  # Reduce auto exposure for more stable lighting
    cap.set(cv2.CAP_PROP_BRIGHTNESS, 0.5)     # Adjust brightness
    cap.set(cv2.CAP_PROP_CONTRAST, 0.5)       # Adjust contrast
    
    if not cap.isOpened():
        print("Error: Could not open webcam.")
        raise RuntimeError("Webcam initialization failed.")
    
    # Verify actual resolution
    actual_width = int(cap.get(cv2.CAP_PROP_FRAME_WIDTH))
    actual_height = int(cap.get(cv2.CAP_PROP_FRAME_HEIGHT))
    actual_fps = cap.get(cv2.CAP_PROP_FPS)
    
    print(f"Webcam initialized with actual resolution: {actual_width}x{actual_height} at {actual_fps:.1f} FPS")
    
    return cap

def read_frame(cap):
    """Read and preprocess a frame from the webcam."""
    ret, frame = cap.read()
    if not ret:
        print("Error: Failed to capture image.")
        return None
    
    # Flip frame horizontally (mirror effect)
    frame = cv2.flip(frame, 1)
    
    # Optional: Apply some preprocessing for better face detection
    # Enhance contrast slightly
    frame = cv2.convertScaleAbs(frame, alpha=1.1, beta=10)
    
    return frame

def release_webcam(cap):
    """Release the webcam resource."""
    if cap:
        print("Releasing webcam...")
        cap.release()
        print("Webcam released successfully")