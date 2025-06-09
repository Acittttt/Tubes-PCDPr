import tkinter as tk
from tkinter import filedialog, messagebox
import subprocess
import sys

class GestureControlApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Gesture Control for Presentation")
        self.pptx_path = ""

        tk.Label(root, text="Select PowerPoint File:").pack(pady=10)
        tk.Button(root, text="Browse", command=self.browse_file).pack()
        self.path_label = tk.Label(root, text="No file selected")
        self.path_label.pack(pady=10)
        tk.Button(root, text="Start Gesture Control", command=self.start_control).pack(pady=10)

    def browse_file(self):
        self.pptx_path = filedialog.askopenfilename(filetypes=[("PowerPoint files", "*.pptx")])
        self.path_label.config(text=self.pptx_path or "No file selected")

    def start_control(self):
        if self.pptx_path:
            try:
                result = subprocess.run([sys.executable, "main.py", self.pptx_path], 
                                      check=True, 
                                      capture_output=True, 
                                      text=True)
                print("Output from main.py:", result.stdout)  # Debugging
            except subprocess.CalledProcessError as e:
                error_msg = f"Failed to run main.py: {e}\nOutput: {e.stderr}"
                print(error_msg)  # Debugging
                messagebox.showerror("Error", error_msg)
        else:
            messagebox.showerror("Error", "Please select a PowerPoint file.")

if __name__ == "__main__":
    root = tk.Tk()
    app = GestureControlApp(root)
    root.geometry("400x200")
    root.mainloop()