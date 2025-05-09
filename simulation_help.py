import tkinter as tk
from tkinter import ttk, scrolledtext
from platform_utils import PlatformUtils

class SimulationHelpWindow:
    """Shows information about simulation mode in non-Windows environments"""
    
    def __init__(self, parent=None):
        """Initialize the help window
        
        Args:
            parent: Parent window or None to create a standalone window
        """
        if parent:
            self.window = tk.Toplevel(parent)
        else:
            self.window = tk.Tk()
            
        self.window.title("EZCAD2 Automation - Simulation Mode Information")
        self.window.geometry("700x500")
        
        self._create_widgets()
    
    def _create_widgets(self):
        """Create the window widgets"""
        # Main frame
        main_frame = ttk.Frame(self.window, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Title
        title_label = ttk.Label(main_frame, text="EZCAD2 Automation - Simulation Mode", 
                               font=("Arial", 16, "bold"))
        title_label.pack(pady=10)
        
        # Platform info
        platform_text = f"Current Platform: {PlatformUtils.get_platform_info()}"
        platform_label = ttk.Label(main_frame, text=platform_text, font=("Arial", 10))
        platform_label.pack(pady=5)
        
        # Information text
        info_text = self._get_info_text()
        info_area = scrolledtext.ScrolledText(main_frame, wrap=tk.WORD, height=15)
        info_area.pack(fill=tk.BOTH, expand=True, pady=10)
        info_area.insert(tk.END, info_text)
        info_area.config(state=tk.DISABLED)
        
        # Feature support frame
        support_frame = ttk.LabelFrame(main_frame, text="Feature Support in Simulation Mode")
        support_frame.pack(fill=tk.X, pady=10)
        
        # Feature grid
        features = [
            ("Excel File Loading", "✓ Full Support"),
            ("Excel File Processing", "✓ Full Support"),
            ("Directory Watching", "✓ Full Support"),
            ("Job Queue Management", "✓ Full Support"),
            ("EZCAD2 Launch", "⚠ Simulated"),
            ("EZCAD2 Commands (RED/MARK)", "⚠ Simulated"),
            ("EZD File Handling", "⚠ Simulated"),
        ]
        
        for i, (feature, status) in enumerate(features):
            ttk.Label(support_frame, text=feature).grid(row=i, column=0, sticky="w", padx=5, pady=2)
            ttk.Label(support_frame, text=status).grid(row=i, column=1, sticky="w", padx=5, pady=2)
        
        # Close button
        ttk.Button(main_frame, text="Close", command=self.window.destroy).pack(pady=10)
    
    def _get_info_text(self):
        """Get the information text"""
        return (
            "This application is running in SIMULATION MODE because EZCAD2 is a Windows-only application "
            "that requires specific Windows components to operate.\n\n"
            
            "In simulation mode:\n"
            "• All Windows-specific operations related to EZCAD2 control are simulated\n"
            "• The application will show what would happen, without actually controlling EZCAD2\n"
            "• Excel file loading and processing functions work normally\n"
            "• Directory monitoring and job queue management work normally\n\n"
            
            "This mode is useful for:\n"
            "• Testing and development on non-Windows platforms\n"
            "• Configuring directory monitoring and processing settings\n"
            "• Working with Excel files for import/export\n"
            "• Setting up and saving configuration profiles\n\n"
            
            "For full functionality including actual EZCAD2 control, please run this application on a Windows system "
            "with EZCAD2 installed."
        )
    
    def show(self):
        """Show the window"""
        if isinstance(self.window, tk.Tk):
            self.window.mainloop()
        # For Toplevel windows, they're shown automatically

class SimulationHelper:
    """Helper class for simulating EZCAD operations in non-Windows environments"""

    @staticmethod
    def simulate_ezcad_start():
        """Simulate starting EZCAD"""
        return True

    @staticmethod
    def simulate_command(command):
        """Simulate sending a command to EZCAD"""
        return True

    @staticmethod
    def simulate_close():
        """Simulate closing EZCAD"""
        return True

def show_simulation_info(parent=None):
    """Show simulation mode information"""
    if not PlatformUtils.is_windows():
        SimulationHelpWindow(parent).show()