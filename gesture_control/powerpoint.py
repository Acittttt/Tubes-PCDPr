import ctypes
import win32com.client
import win32gui
import win32con
import time
import sys

def minimize_console():
    """Minimize the command prompt window."""
    kernel32 = ctypes.WinDLL('kernel32')
    user32 = ctypes.WinDLL('user32')
    hWnd = kernel32.GetConsoleWindow()
    if hWnd:
        user32.ShowWindow(hWnd, 6)  # 6 = Minimize window

def initialize_powerpoint(pptx_path):
    """Initialize PowerPoint and open the presentation."""
    try:
        if not pptx_path:
            print("Error: No PowerPoint file path provided.")
            sys.exit(1)
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        presentation = powerpoint.Presentations.Open(pptx_path)
        presentation.SlideShowSettings.Run()
        time.sleep(3)  # Wait for slideshow window to appear
        return powerpoint, presentation
    except Exception as e:
        print(f"Error opening PowerPoint: {e}")
        sys.exit(1)

def bring_to_foreground(powerpoint):
    """Bring the PowerPoint slideshow window to the foreground."""
    try:
        hwnd = powerpoint.SlideShowWindows(1).HWND
        if hwnd:
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            win32gui.SetForegroundWindow(hwnd)
        else:
            print("Error: Could not get PowerPoint slideshow window handle.")
            sys.exit(1)
    except Exception as e:
        print(f"Error setting PowerPoint window to foreground: {e}")
        sys.exit(1)

def close_powerpoint(powerpoint, presentation):
    """Close the PowerPoint presentation and application."""
    if presentation:
        presentation.Close()
    if powerpoint:
        powerpoint.Quit()