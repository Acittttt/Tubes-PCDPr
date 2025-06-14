import streamlit as st
import subprocess
import sys
import os
import tempfile
import time
import shutil

def run_gesture_control(pptx_path):
    """Run gesture control script with the given PowerPoint file path."""
    try:
        print(f"Running gesture_control.py with file: {pptx_path}")
        
        # Use shell=True for Windows compatibility
        result = subprocess.run(
            [sys.executable, "gesture_control.py", pptx_path],
            check=True,
            capture_output=True,
            text=True,
            shell=True,
            timeout=300  # 5 minute timeout
        )
        
        st.success("Head gesture control completed successfully!")
        if result.stdout:
            st.write("Output from gesture_control.py:")
            st.code(result.stdout)
        print("Head gesture control process completed.")
        
    except subprocess.TimeoutExpired:
        st.warning("Head gesture control timed out after 5 minutes")
        print("Head gesture control process timed out")
        
    except subprocess.CalledProcessError as e:
        error_msg = f"Failed to run gesture_control.py (Exit code: {e.returncode})"
        st.error(error_msg)
        
        if e.stderr:
            st.write("Error output:")
            st.code(e.stderr)
            print(f"STDERR: {e.stderr}")
            
        if e.stdout:
            st.write("Standard output:")
            st.code(e.stdout)
            print(f"STDOUT: {e.stdout}")
            
        print(error_msg)
        raise
        
    except Exception as e:
        error_msg = f"Unexpected error: {e}"
        st.error(error_msg)
        print(error_msg)
        raise

def safe_delete_file(file_path, max_attempts=10, delay=2):
    """Safely delete a file with multiple retry attempts."""
    if not os.path.exists(file_path):
        print(f"File {file_path} does not exist, nothing to delete.")
        return True
        
    print(f"Attempting to delete file: {file_path}")
    
    # First, try to close any handles to the file
    try:
        # Force garbage collection
        import gc
        gc.collect()
        time.sleep(1)
    except:
        pass
    
    for attempt in range(max_attempts):
        try:
            # Try different methods to delete the file
            if attempt < 5:
                os.unlink(file_path)
            else:
                # Use shutil.rmtree for stubborn files
                if os.path.isdir(file_path):
                    shutil.rmtree(file_path)
                else:
                    os.chmod(file_path, 0o777)  # Change permissions
                    os.unlink(file_path)
                    
            print(f"File {file_path} deleted successfully on attempt {attempt + 1}.")
            return True
            
        except PermissionError as e:
            print(f"Attempt {attempt + 1} failed: {e}")
            if attempt < max_attempts - 1:
                print(f"Waiting {delay} seconds before retry...")
                time.sleep(delay)
                delay *= 1.5  # Exponential backoff
                
        except Exception as e:
            print(f"Unexpected error while deleting file: {e}")
            if attempt < max_attempts - 1:
                time.sleep(delay)
            else:
                return False
    
    print(f"Failed to delete file {file_path} after {max_attempts} attempts.")
    print("File will be left in temp directory and cleaned up by system later.")
    return False

def main():
    st.title("🧠 Head Gesture Control for PowerPoint Presentation")
    st.markdown("Upload your PowerPoint file and control it with head movements!")

    st.header("📁 Select PowerPoint File")
    uploaded_file = st.file_uploader(
        "Choose a PowerPoint file", 
        type=["pptx", "ppt"],
        help="Upload your PowerPoint presentation file to control with head gestures"
    )

    if uploaded_file is not None:
        # Create a more persistent temporary file
        temp_dir = tempfile.gettempdir()
        safe_filename = "".join(c for c in uploaded_file.name if c.isalnum() or c in (' ', '.', '_')).rstrip()
        pptx_path = os.path.join(temp_dir, f"head_gesture_control_{int(time.time())}_{safe_filename}")
        
        try:
            # Save the uploaded file
            with open(pptx_path, "wb") as tmp_file:
                tmp_file.write(uploaded_file.read())
            
            st.success(f"✅ File uploaded: {uploaded_file.name}")
            st.info(f"📍 Temporary file location: {pptx_path}")

            # Instructions
            st.header("🎮 How to Use Head Gesture Control")
            
            # Create columns for better layout
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown("""
                **Head Gesture Controls:**
                - 🔄 **Tilt Head Right**: Next slide
                - 🔄 **Tilt Head Left**: Previous slide  
                - 🔄🔄🔄 **Triple Tilt**: Close presentation
                - **ESC Key**: Exit application
                """)
            
            with col2:
                st.markdown("""
                **Tips for Best Results:**
                - 💡 Ensure good lighting on your face
                - 📹 Keep your face centered in the camera
                - 🎯 Make clear, deliberate head movements
                - ⏱️ Wait 0.8 seconds between gestures
                """)

            st.markdown("""
            ---
            **Step-by-Step Instructions:**
            1. Click "Start Head Gesture Control" below
            2. PowerPoint will open automatically in slideshow mode
            3. Position yourself in front of the camera with good lighting
            4. Use head movements to control the presentation:
               - **Tilt your head to the RIGHT** → Next slide
               - **Tilt your head to the LEFT** → Previous slide
               - **Tilt your head 3 times in the SAME direction quickly** → Close presentation
            5. The application will show your head position and detected gestures
            """)
            
            # Warning box
            st.warning("""
            ⚠️ **Important:** 
            - Make sure your webcam is working properly
            - Ensure your face is clearly visible and well-lit
            - Head gestures need to be deliberate (tilt at least 20 degrees for triple tilt, 15 degrees for navigation)
            - For triple tilt: Make three clear head tilts in the same direction within 3 seconds
            - Triple tilt requires more pronounced head movement than regular navigation tilts
            """)

            if st.button("🚀 Start Head Gesture Control", type="primary"):
                if os.path.exists(pptx_path):
                    with st.spinner("Starting head gesture control... Please wait"):
                        try:
                            run_gesture_control(pptx_path)
                        except Exception as e:
                            st.error(f"❌ Error during head gesture control: {e}")
                        finally:
                            # Clean up the temporary file
                            st.info("🧹 Cleaning up temporary files...")
                            safe_delete_file(pptx_path)
                else:
                    st.error("❌ Temporary file not found. Please re-upload your file.")
                    
        except Exception as e:
            st.error(f"❌ Error handling file: {e}")
            # Try to clean up if file was created
            if 'pptx_path' in locals():
                safe_delete_file(pptx_path)
            
    else:
        st.info("👆 Please upload a PowerPoint file to get started.")
        
        # Add demo section
        st.header("🎭 How Head Gestures Work")
        
        demo_col1, demo_col2, demo_col3 = st.columns(3)
        
        with demo_col1:
            st.markdown("""
            **Tilt Right** ➡️
            
            Tilt your head to the right (your right shoulder moves down) to advance to the next slide.
            """)
        
        with demo_col2:
            st.markdown("""
            **Tilt Left** ⬅️
            
            Tilt your head to the left (your left shoulder moves down) to go to the previous slide.
            """)
            
        with demo_col3:
            st.markdown("""
            **Triple Tilt** 🔄🔄🔄
            
            Tilt your head three times quickly in the same direction (left or right) to close the presentation.
            """)
        
    # Footer
    st.markdown("---")
    st.markdown("""
    **System Requirements:** 
    - Working webcam with good resolution (720p or higher recommended)
    - PowerPoint installed on your system
    - Good lighting conditions for face detection
    - MediaPipe and OpenCV libraries installed
    """)

if __name__ == "__main__":
    main()