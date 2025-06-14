import ctypes
import win32com.client
import win32gui
import win32con
import time
import sys

def minimize_console():
    """Minimize the command prompt window."""
    print("Minimizing console window...")
    try:
        kernel32 = ctypes.WinDLL('kernel32')
        user32 = ctypes.WinDLL('user32')
        hWnd = kernel32.GetConsoleWindow()
        if hWnd:
            user32.ShowWindow(hWnd, 6)  # 6 = Minimize window
            print("Console window minimized successfully.")
        else:
            print("Console window handle not found.")
    except Exception as e:
        print(f"Warning: Could not minimize console window: {e}")

def initialize_powerpoint(pptx_path):
    """Initialize PowerPoint and open the presentation."""
    print(f"Initializing PowerPoint with file: {pptx_path}")
    try:
        if not pptx_path:
            print("Error: No PowerPoint file path provided.")
            sys.exit(1)
            
        # Check if file exists
        import os
        if not os.path.exists(pptx_path):
            print(f"Error: PowerPoint file does not exist: {pptx_path}")
            sys.exit(1)
            
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        powerpoint.Visible = 1  # Make PowerPoint visible
        
        presentation = powerpoint.Presentations.Open(pptx_path)
        
        # Start slideshow
        slideshow = presentation.SlideShowSettings.Run()
        
        # Wait for slideshow window to appear
        time.sleep(3)
        
        print("PowerPoint presentation opened and slideshow started.")
        return powerpoint, presentation
        
    except Exception as e:
        print(f"Error opening PowerPoint: {e}")
        sys.exit(1)

def bring_to_foreground(powerpoint):
    """Bring the PowerPoint slideshow window to the foreground."""
    print("Bringing PowerPoint window to foreground...")
    try:
        # Method 1: Try to get slideshow window handle
        try:
            slideshow_window = powerpoint.SlideShowWindows(1)
            hwnd = int(slideshow_window.HWND)  # Convert to integer
            
            if hwnd:
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                win32gui.SetForegroundWindow(hwnd)
                win32gui.BringWindowToTop(hwnd)
                print("PowerPoint slideshow window brought to foreground (Method 1).")
                return True
        except Exception as e1:
            print(f"Method 1 failed: {e1}")
        
        # Method 2: Find PowerPoint window by class name
        try:
            def enum_window_callback(hwnd, windows):
                if win32gui.IsWindowVisible(hwnd):
                    window_text = win32gui.GetWindowText(hwnd)
                    class_name = win32gui.GetClassName(hwnd)
                    if "PowerPoint" in window_text or "screenClass" in class_name:
                        windows.append(hwnd)
                return True
            
            windows = []
            win32gui.EnumWindows(enum_window_callback, windows)
            
            if windows:
                hwnd = windows[0]
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                win32gui.SetForegroundWindow(hwnd)
                win32gui.BringWindowToTop(hwnd)
                print("PowerPoint window brought to foreground (Method 2).")
                return True
        except Exception as e2:
            print(f"Method 2 failed: {e2}")
        
        # Method 3: Use PowerPoint application window
        try:
            app_hwnd = powerpoint.HWND
            if app_hwnd:
                hwnd = int(app_hwnd)
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                win32gui.SetForegroundWindow(hwnd)
                print("PowerPoint application window brought to foreground (Method 3).")
                return True
        except Exception as e3:
            print(f"Method 3 failed: {e3}")
        
        print("Warning: Could not bring PowerPoint window to foreground, but continuing...")
        return False
        
    except Exception as e:
        print(f"Error setting PowerPoint window to foreground: {e}")
        print("Continuing without setting foreground...")
        return False

def close_powerpoint(powerpoint, presentation):
    """Close the PowerPoint presentation and application."""
    print("Closing PowerPoint...")
    try:
        # Close slideshow first
        try:
            if powerpoint.SlideShowWindows.Count > 0:
                powerpoint.SlideShowWindows(1).View.Exit()
                print("Slideshow closed.")
        except Exception as e:
            print(f"Could not close slideshow: {e}")
        
        # Close presentation
        if presentation:
            try:
                presentation.Close()
                print("PowerPoint presentation closed.")
            except Exception as e:
                print(f"Could not close presentation: {e}")
        
        # Quit PowerPoint application
        if powerpoint:
            try:
                powerpoint.Quit()
                print("PowerPoint application quit.")
            except Exception as e:
                print(f"Could not quit PowerPoint: {e}")
            
        # Clean up COM objects
        try:
            import pythoncom
            pythoncom.CoUninitialize()
            print("COM objects cleaned up.")
        except Exception as e:
            print(f"Could not clean up COM objects: {e}")
            
    except Exception as e:
        print(f"Error closing PowerPoint: {e}")

def check_slideshow_active(powerpoint):
    """Check if slideshow is currently active."""
    try:
        return powerpoint.SlideShowWindows.Count > 0
    except:
        return False