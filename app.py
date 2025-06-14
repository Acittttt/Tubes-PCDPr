import streamlit as st
import subprocess
import sys
import os
import tempfile

def run_gesture_control(pptx_path):
    try:
        result = subprocess.run(
            [sys.executable, "gesture_control.py", pptx_path],
            check=True,
            capture_output=True,
            text=True
        )
        st.success("Gesture control started successfully!")
        st.write("Output from gesture_control.py:")
        st.code(result.stdout)
    except subprocess.CalledProcessError as e:
        error_msg = f"Failed to run gesture_control.py: {e}\nOutput: {e.stderr}"
        st.error(error_msg)

def main():
    st.title("Gesture Control for Presentation")

    st.header("Select PowerPoint File")
    uploaded_file = st.file_uploader("Choose a PowerPoint file", type=["pptx"])

    if uploaded_file is not None:
        # Save the uploaded file to a temporary location
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as tmp_file:
            tmp_file.write(uploaded_file.read())
            pptx_path = tmp_file.name

        st.write(f"Selected file: {uploaded_file.name}")

        if st.button("Start Gesture Control"):
            if pptx_path:
                run_gesture_control(pptx_path)
                # Clean up the temporary file after execution
                os.unlink(pptx_path)
            else:
                st.error("No file selected. Please upload a PowerPoint file.")
    else:
        st.write("No file selected.")

if __name__ == "__main__":
    main()