import tkinter as tk
import sys
import os
import logging
import platform
from platform_utils import PlatformUtils

# Import the compatible version for all platforms
# It will handle platform-specific behavior internally
from ezcad_automation_compat import EZCADAutomationApp

def setup_exception_logging():
    """Setup global exception handler to log unhandled exceptions"""
    def handle_exception(exc_type, exc_value, exc_traceback):
        """Log unhandled exceptions"""
        if issubclass(exc_type, KeyboardInterrupt):
            # Let keyboard interrupts pass through
            sys.__excepthook__(exc_type, exc_value, exc_traceback)
            return
        
        # Log the exception
        logger = logging.getLogger('EZCADAutomation')
        logger.critical("Unhandled exception", exc_info=(exc_type, exc_value, exc_traceback))
        
        # Show error message to user
        from tkinter import messagebox
        messagebox.showerror("Unexpected Error", 
                             f"An unexpected error occurred: {exc_value}\n\n"
                             "Please check the log file for details.")
    
    # Set the exception hook
    sys.excepthook = handle_exception

def check_requirements():
    """Check for required libraries and show warning if missing"""
    try:
        import pandas
        import watchdog
        if PlatformUtils.is_windows():
            import pywinauto
        import psutil
    except ImportError as e:
        from tkinter import messagebox
        req_packages = "pandas watchdog psutil"
        if PlatformUtils.is_windows():
            req_packages += " pywinauto"
            
        messagebox.showwarning("Missing Dependency", 
                               f"Required library is missing: {e}\n\n"
                               "Please install the required packages:\n"
                               f"pip install {req_packages}")
        return False
    return True

def main():
    """Main entry point for the application"""
    # Set up exception logging
    setup_exception_logging()
    
    # Check for required dependencies
    if not check_requirements():
        return
    
    # Create the main window
    root = tk.Tk()
    root.title("EZCAD2 Automation")
    
    # Set icon (optional)
    try:
        root.iconbitmap("icon.ico")
    except:
        pass  # No icon available, continue without
    
    # Create the application
    app = EZCADAutomationApp(root)
    
    # Set up the window close handler
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    
    # Start the main loop
    root.mainloop()

if __name__ == "__main__":
    main()
