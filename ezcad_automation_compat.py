import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import os
import threading
import queue
import time
from datetime import datetime
import json
import platform
from platform_utils import PlatformUtils

# Import our modules
from logger import LoggerSetup, LogPanel
from config_manager_compat import ConfigManager
from excel_handler import ExcelHandler
from ezcad_controller_compat import EZCADController
from processor import Processor
from queue_manager import QueueManager
from watcher import DirectoryWatcher
from simulation_help import show_simulation_info

class EZCADAutomationApp:
    """The main EZCAD Automation application with cross-platform compatibility"""
    
    def __init__(self, root):
        """Initialize the application"""
        self.root = root
        root.title("EZCAD2 Automation" + (" (Simulation Mode)" if not PlatformUtils.is_windows() else ""))
        root.geometry("900x700")
        
        # Initialize logger
        self.logger_setup = LoggerSetup()
        self.logger = self.logger_setup.get_logger()
        self.log_queue = self.logger_setup.get_queue()
        
        # Log platform information
        self.logger.info(f"Platform: {platform.system()} {platform.release()}")
        if not PlatformUtils.is_windows():
            self.logger.warning("Running in non-Windows environment. EZCAD automation will run in simulation mode.")
        
        # Initialize config
        self.config = ConfigManager()
        
        # Initialize components
        self.excel_handler = ExcelHandler(self.logger)
        self.ezcad_controller = EZCADController(self.config, self.logger)
        self.processor = Processor(self.excel_handler, self.ezcad_controller, self.config, self.logger)
        self.queue_manager = QueueManager(self.processor, self.config, self.logger)
        self.directory_watcher = DirectoryWatcher(self.config, self.queue_manager.file_queue, self.logger)
        
        # Show simulation help on non-Windows platforms
        if not PlatformUtils.is_windows():
            # Add a short delay to let the GUI initialize first
            self.root.after(1000, lambda: show_simulation_info(self.root))
        
        # Setup the UI
        self._create_ui()
        
        # Refresh UI with loaded config
        self._refresh_from_config()
        
        # Start processing if auto-start is enabled
        if self.config.getboolean('Settings', 'auto_start', fallback=False):
            self._start_automation()
    
    def _create_ui(self):
        """Create the application UI"""
        # Create a notebook (tabbed interface)
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Create main tab
        main_tab = ttk.Frame(notebook)
        notebook.add(main_tab, text="Main")
        self._create_main_tab(main_tab)
        
        # Create monitoring tab
        monitor_tab = ttk.Frame(notebook)
        notebook.add(monitor_tab, text="Monitoring")
        self._create_monitor_tab(monitor_tab)
        
        # Create jobs tab
        jobs_tab = ttk.Frame(notebook)
        notebook.add(jobs_tab, text="Jobs")
        self._create_jobs_tab(jobs_tab)
        
        # Create settings tab
        settings_tab = ttk.Frame(notebook)
        notebook.add(settings_tab, text="Settings")
        self._create_settings_tab(settings_tab)
        
        # Create log tab
        log_tab = ttk.Frame(notebook)
        notebook.add(log_tab, text="Logs")
        self._create_log_tab(log_tab)
        
        # Status bar
        status_frame = ttk.Frame(self.root)
        status_frame.pack(fill=tk.X, padx=10, pady=5)
        
        self.status_var = tk.StringVar(value="Ready")
        status_label = ttk.Label(status_frame, textvariable=self.status_var)
        status_label.pack(side=tk.LEFT)
        
        # Clock label on right of status bar
        self.clock_var = tk.StringVar()
        clock_label = ttk.Label(status_frame, textvariable=self.clock_var)
        clock_label.pack(side=tk.RIGHT)
        
        # Start the clock update
        self._update_clock()
    
    def _create_main_tab(self, parent):
        """Create the main tab content"""
        # EZCAD2.exe selection
        path_frame = ttk.LabelFrame(parent, text="Paths")
        path_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # EZCAD2.exe
        ttk.Label(path_frame, text="EZCAD2.exe:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.ezcad_exe_var = tk.StringVar()
        ttk.Entry(path_frame, textvariable=self.ezcad_exe_var, width=60).grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(path_frame, text="Browse", command=self._select_ezcad_exe).grid(row=0, column=2, padx=5, pady=5)
        
        # Excel file
        ttk.Label(path_frame, text="Excel File:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.excel_path_var = tk.StringVar()
        ttk.Entry(path_frame, textvariable=self.excel_path_var, width=60).grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(path_frame, text="Browse", command=self._select_excel).grid(row=1, column=2, padx=5, pady=5)
        
        # EZD file
        ttk.Label(path_frame, text="EZD File:").grid(row=2, column=0, sticky="w", padx=5, pady=5)
        self.ezd_path_var = tk.StringVar()
        ttk.Entry(path_frame, textvariable=self.ezd_path_var, width=60).grid(row=2, column=1, padx=5, pady=5)
        ttk.Button(path_frame, text="Browse", command=self._select_ezd).grid(row=2, column=2, padx=5, pady=5)
        
        # Excel preview
        preview_frame = ttk.LabelFrame(parent, text="Excel Preview")
        preview_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        self.preview_text = scrolledtext.ScrolledText(preview_frame, height=10)
        self.preview_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.preview_text.config(state=tk.DISABLED)
        
        # Control buttons
        control_frame = ttk.Frame(parent)
        control_frame.pack(fill=tk.X, padx=10, pady=10)
        
        self.start_button = ttk.Button(control_frame, text="Start EZCAD", command=self._run_ezcad)
        self.start_button.pack(side=tk.LEFT, padx=5)
        
        self.red_button = ttk.Button(control_frame, text="Red", command=lambda: self._send_command('red'), state=tk.DISABLED)
        self.red_button.pack(side=tk.LEFT, padx=5)
        
        self.mark_button = ttk.Button(control_frame, text="Mark", command=lambda: self._send_command('mark'), state=tk.DISABLED)
        self.mark_button.pack(side=tk.LEFT, padx=5)
        
        self.process_button = ttk.Button(control_frame, text="Process Excel", command=self._process_excel)
        self.process_button.pack(side=tk.LEFT, padx=5)
        
        self.select_window_button = ttk.Button(control_frame, text="Select EZCAD Window", command=self._select_ezcad_window)
        self.select_window_button.pack(side=tk.LEFT, padx=5)
    
    def _create_monitor_tab(self, parent):
        """Create the monitoring tab content"""
        # Watch directory frame
        watch_frame = ttk.LabelFrame(parent, text="Directory Monitoring")
        watch_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # Watch directory
        ttk.Label(watch_frame, text="Watch Directory:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.watch_dir_var = tk.StringVar()
        ttk.Entry(watch_frame, textvariable=self.watch_dir_var, width=60).grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(watch_frame, text="Browse", command=self._select_watch_dir).grid(row=0, column=2, padx=5, pady=5)
        
        # Monitoring options
        options_frame = ttk.Frame(watch_frame)
        options_frame.grid(row=1, column=0, columnspan=3, sticky="w", padx=5, pady=5)
        
        self.monitor_enabled_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(options_frame, text="Enable Monitoring", variable=self.monitor_enabled_var).pack(side=tk.LEFT, padx=5)
        
        self.recursive_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(options_frame, text="Include Subdirectories", variable=self.recursive_var).pack(side=tk.LEFT, padx=5)
        
        self.auto_trigger_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(options_frame, text="Auto-Trigger Processing", variable=self.auto_trigger_var).pack(side=tk.LEFT, padx=5)
        
        # Control buttons
        monitor_control_frame = ttk.Frame(watch_frame)
        monitor_control_frame.grid(row=2, column=0, columnspan=3, sticky="w", padx=5, pady=5)
        
        self.start_monitoring_button = ttk.Button(monitor_control_frame, text="Start Monitoring", command=self._start_monitoring)
        self.start_monitoring_button.pack(side=tk.LEFT, padx=5)
        
        self.stop_monitoring_button = ttk.Button(monitor_control_frame, text="Stop Monitoring", command=self._stop_monitoring, state=tk.DISABLED)
        self.stop_monitoring_button.pack(side=tk.LEFT, padx=5)
        
        # File patterns
        patterns_frame = ttk.LabelFrame(parent, text="File Patterns")
        patterns_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Label(patterns_frame, text="Excel Files:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.excel_pattern_var = tk.StringVar(value="*.xls;*.xlsx")
        ttk.Entry(patterns_frame, textvariable=self.excel_pattern_var).grid(row=0, column=1, sticky="ew", padx=5, pady=5)
        
        ttk.Label(patterns_frame, text="EZD Files:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.ezd_pattern_var = tk.StringVar(value="*.ezd")
        ttk.Entry(patterns_frame, textvariable=self.ezd_pattern_var).grid(row=1, column=1, sticky="ew", padx=5, pady=5)
        
        patterns_frame.columnconfigure(1, weight=1)
        
        # File events list
        events_frame = ttk.LabelFrame(parent, text="Recent File Events")
        events_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        self.events_text = scrolledtext.ScrolledText(events_frame, height=10)
        self.events_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.events_text.config(state=tk.DISABLED)
    
    def _create_jobs_tab(self, parent):
        """Create the jobs tab content"""
        # Job queue controls
        controls_frame = ttk.Frame(parent)
        controls_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Button(controls_frame, text="Start Processing", command=self._start_processing).pack(side=tk.LEFT, padx=5)
        ttk.Button(controls_frame, text="Stop Processing", command=self._stop_processing).pack(side=tk.LEFT, padx=5)
        ttk.Button(controls_frame, text="Clear Completed", command=self._clear_completed_jobs).pack(side=tk.LEFT, padx=5)
        ttk.Button(controls_frame, text="Refresh", command=self._refresh_job_list).pack(side=tk.LEFT, padx=5)
        
        # Job list
        job_frame = ttk.Frame(parent)
        job_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Create treeview for jobs
        columns = ("id", "file", "type", "status", "added", "duration")
        self.job_tree = ttk.Treeview(job_frame, columns=columns, show="headings")
        
        # Configure columns
        self.job_tree.heading("id", text="Job ID")
        self.job_tree.heading("file", text="File")
        self.job_tree.heading("type", text="Type")
        self.job_tree.heading("status", text="Status")
        self.job_tree.heading("added", text="Added")
        self.job_tree.heading("duration", text="Duration")
        
        self.job_tree.column("id", width=100)
        self.job_tree.column("file", width=250)
        self.job_tree.column("type", width=80)
        self.job_tree.column("status", width=100)
        self.job_tree.column("added", width=150)
        self.job_tree.column("duration", width=100)
        
        # Add scrollbar
        job_scroll = ttk.Scrollbar(job_frame, orient=tk.VERTICAL, command=self.job_tree.yview)
        self.job_tree.configure(yscrollcommand=job_scroll.set)
        
        # Pack the treeview and scrollbar
        self.job_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        job_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Bind right-click for context menu
        self.job_tree.bind("<Button-3>", self._show_job_context_menu)
        
        # Bind double-click for job details
        self.job_tree.bind("<Double-1>", self._show_job_details)
    
    def _create_settings_tab(self, parent):
        """Create the settings tab content"""
        # General settings
        general_frame = ttk.LabelFrame(parent, text="General Settings")
        general_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # Auto-start
        self.auto_start_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(general_frame, text="Auto-Start Processing on Launch", variable=self.auto_start_var).grid(row=0, column=0, sticky="w", padx=5, pady=5)
        
        # Minimize on start
        self.minimize_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(general_frame, text="Minimize EZCAD After Launch", variable=self.minimize_var).grid(row=1, column=0, sticky="w", padx=5, pady=5)
        
        # Multiple instances
        self.multiple_instances_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(general_frame, text="Allow Multiple EZCAD Instances", variable=self.multiple_instances_var).grid(row=2, column=0, sticky="w", padx=5, pady=5)
        
        # Batch processing
        self.batch_process_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(general_frame, text="Batch Processing", variable=self.batch_process_var).grid(row=3, column=0, sticky="w", padx=5, pady=5)
        
        # Concurrency
        concurrency_frame = ttk.Frame(general_frame)
        concurrency_frame.grid(row=4, column=0, sticky="w", padx=5, pady=5)
        
        ttk.Label(concurrency_frame, text="Max Concurrent Processes:").pack(side=tk.LEFT)
        
        self.max_concurrent_var = tk.StringVar(value="1")
        concurrency_spin = ttk.Spinbox(concurrency_frame, from_=1, to=5, width=5, textvariable=self.max_concurrent_var)
        concurrency_spin.pack(side=tk.LEFT, padx=5)
        
        # Simulation mode indicator (for non-Windows platforms)
        if not PlatformUtils.is_windows():
            sim_frame = ttk.LabelFrame(parent, text="Simulation Mode")
            sim_frame.pack(fill=tk.X, padx=10, pady=10)
            
            ttk.Label(sim_frame, text="Running in simulation mode. EZCAD operations will be simulated.", 
                    foreground="blue").pack(padx=10, pady=5)
            
            ttk.Button(sim_frame, text="Show Simulation Info", 
                      command=lambda: show_simulation_info(self.root)).pack(pady=5)
        
        # Profiles
        profile_frame = ttk.LabelFrame(parent, text="Configuration Profiles")
        profile_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # Profile controls
        profile_controls = ttk.Frame(profile_frame)
        profile_controls.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(profile_controls, text="Profile Name:").pack(side=tk.LEFT, padx=5)
        
        self.profile_name_var = tk.StringVar()
        ttk.Entry(profile_controls, textvariable=self.profile_name_var, width=20).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(profile_controls, text="Save Profile", command=self._save_profile).pack(side=tk.LEFT, padx=5)
        ttk.Button(profile_controls, text="Load Profile", command=self._show_load_profile).pack(side=tk.LEFT, padx=5)
        
        # Save/Apply buttons
        button_frame = ttk.Frame(parent)
        button_frame.pack(fill=tk.X, padx=10, pady=(20, 10))
        
        ttk.Button(button_frame, text="Apply Settings", command=self._apply_settings).pack(side=tk.RIGHT, padx=5)
        ttk.Button(button_frame, text="Reset to Default", command=self._reset_settings).pack(side=tk.RIGHT, padx=5)
    
    def _create_log_tab(self, parent):
        """Create the log tab content"""
        # Create the log panel
        log_frame = ttk.Frame(parent)
        log_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        self.log_panel = LogPanel(log_frame, self.log_queue)
        
        # Log controls
        controls_frame = ttk.Frame(parent)
        controls_frame.pack(fill=tk.X, padx=10, pady=5)
        
        ttk.Button(controls_frame, text="Clear Log", command=self.log_panel.clear).pack(side=tk.LEFT, padx=5)
        ttk.Button(controls_frame, text="Open Log Directory", command=self._open_log_directory).pack(side=tk.LEFT, padx=5)
    
    def _update_clock(self):
        """Update the clock in the status bar"""
        self.clock_var.set(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        self.root.after(1000, self._update_clock)
    
    def _refresh_from_config(self):
        """Refresh UI elements from the config"""
        # Paths
        self.ezcad_exe_var.set(self.config.get('Paths', 'ezcad_exe', fallback=''))
        self.watch_dir_var.set(self.config.get('Monitoring', 'watch_directory', fallback=''))
        
        # Settings
        self.auto_start_var.set(self.config.getboolean('Settings', 'auto_start', fallback=False))
        self.minimize_var.set(self.config.getboolean('Settings', 'minimize_on_start', fallback=False))
        self.multiple_instances_var.set(self.config.getboolean('Settings', 'multiple_instances', fallback=False))
        self.batch_process_var.set(self.config.getboolean('Settings', 'batch_process', fallback=False))
        self.max_concurrent_var.set(str(self.config.getint('Settings', 'max_concurrent_processes', fallback=1)))
        
        # Monitoring
        self.monitor_enabled_var.set(self.config.getboolean('Monitoring', 'enabled', fallback=False))
        self.recursive_var.set(self.config.getboolean('Monitoring', 'recursive', fallback=False))
        self.auto_trigger_var.set(self.config.getboolean('Settings', 'auto_trigger', fallback=False))
        
        # File patterns
        self.excel_pattern_var.set(self.config.get('Settings', 'file_pattern_excel', fallback='*.xls;*.xlsx'))
        self.ezd_pattern_var.set(self.config.get('Settings', 'file_pattern_ezd', fallback='*.ezd'))
    
    def _apply_settings(self):
        """Apply settings from UI to config"""
        # Paths
        self.config.set('Paths', 'ezcad_exe', self.ezcad_exe_var.get())
        self.config.set('Monitoring', 'watch_directory', self.watch_dir_var.get())
        
        # Settings
        self.config.set('Settings', 'auto_start', str(self.auto_start_var.get()))
        self.config.set('Settings', 'minimize_on_start', str(self.minimize_var.get()))
        self.config.set('Settings', 'multiple_instances', str(self.multiple_instances_var.get()))
        self.config.set('Settings', 'batch_process', str(self.batch_process_var.get()))
        self.config.set('Settings', 'max_concurrent_processes', self.max_concurrent_var.get())
        
        # Monitoring
        self.config.set('Monitoring', 'enabled', str(self.monitor_enabled_var.get()))
        self.config.set('Monitoring', 'recursive', str(self.recursive_var.get()))
        self.config.set('Settings', 'auto_trigger', str(self.auto_trigger_var.get()))
        
        # File patterns
        self.config.set('Settings', 'file_pattern_excel', self.excel_pattern_var.get())
        self.config.set('Settings', 'file_pattern_ezd', self.ezd_pattern_var.get())
        
        # Save config
        self.config.save_config()
        
        self.logger.info("Settings applied and saved")
        messagebox.showinfo("Settings", "Settings have been applied and saved.")
    
    def _reset_settings(self):
        """Reset settings to default"""
        if messagebox.askyesno("Reset Settings", "Are you sure you want to reset all settings to default?"):
            self.config._create_default_config()
            self._refresh_from_config()
            self.logger.info("Settings reset to default")
            messagebox.showinfo("Settings", "Settings have been reset to default.")
    
    def _select_ezcad_exe(self):
        """Browse for EZCAD2.exe"""
        file_types = [("EZCAD2 Executable", "EZCAD2.exe"), ("EXE files", "*.exe"), ("All files", "*.*")]
        file_path = filedialog.askopenfilename(title="Select EZCAD2.exe", filetypes=file_types)
        
        if file_path:
            self.ezcad_exe_var.set(file_path)
            self.config.set('Paths', 'ezcad_exe', file_path)
            self.config.save_config()
            self.logger.info(f"Selected EZCAD executable: {file_path}")
    
    def _select_excel(self):
        """Browse for Excel file"""
        file_types = [("Excel files", "*.xls;*.xlsx"), ("All files", "*.*")]
        file_path = filedialog.askopenfilename(title="Select Excel File", filetypes=file_types)
        
        if file_path:
            self.excel_path_var.set(file_path)
            self.config.set('Paths', 'last_excel_file', file_path)
            last_dir = os.path.dirname(file_path)
            self.config.set('Paths', 'last_excel_dir', last_dir)
            self.config.save_config()
            
            # Load and preview
            df = self.excel_handler.load_excel(file_path)
            if df is not None:
                preview = self.excel_handler.get_preview()
                
                self.preview_text.config(state=tk.NORMAL)
                self.preview_text.delete("1.0", tk.END)
                self.preview_text.insert(tk.END, preview)
                self.preview_text.config(state=tk.DISABLED)
    
    def _select_ezd(self):
        """Browse for EZD file"""
        file_types = [("EZD files", "*.ezd"), ("All files", "*.*")]
        file_path = filedialog.askopenfilename(title="Select EZD File", filetypes=file_types)
        
        if file_path:
            self.ezd_path_var.set(file_path)
            self.config.set('Paths', 'last_ezd_file', file_path)
            last_dir = os.path.dirname(file_path)
            self.config.set('Paths', 'last_ezd_dir', last_dir)
            self.config.save_config()
            self.logger.info(f"Selected EZD file: {file_path}")
    
    def _select_watch_dir(self):
        """Browse for directory to watch"""
        dir_path = filedialog.askdirectory(title="Select Directory to Watch")
        
        if dir_path:
            self.watch_dir_var.set(dir_path)
            self.config.set('Monitoring', 'watch_directory', dir_path)
            self.config.save_config()
            self.logger.info(f"Selected watch directory: {dir_path}")
    
    def _run_ezcad(self):
        """Run EZCAD with the selected EZD file"""
        ezd_file = self.ezd_path_var.get()
        if not ezd_file:
            messagebox.showwarning("Missing File", "Please select an EZD file first.")
            return
        
        # Start EZCAD in a separate thread to keep UI responsive
        threading.Thread(target=self._run_ezcad_thread, args=(ezd_file,), daemon=True).start()
    
    def _run_ezcad_thread(self, ezd_file):
        """Thread function to run EZCAD"""
        try:
            self.status_var.set("Starting EZCAD...")
            window_id = self.ezcad_controller.start_ezcad(ezd_file)
            
            if window_id:
                self.logger.info(f"EZCAD started with window ID: {window_id}")
                self.status_var.set("EZCAD running")
                
                # Set the window ID to use for commands
                self.current_window_id = window_id
                
                # Enable command buttons
                self.root.after(0, lambda: self.red_button.config(state=tk.NORMAL))
                self.root.after(0, lambda: self.mark_button.config(state=tk.NORMAL))
            else:
                self.logger.error("Failed to start EZCAD")
                self.root.after(0, lambda: messagebox.showerror("EZCAD Error", "Failed to start EZCAD."))
                self.status_var.set("Error starting EZCAD")
                
        except Exception as e:
            self.logger.error(f"Error in EZCAD thread: {str(e)}")
            self.root.after(0, lambda: messagebox.showerror("EZCAD Error", f"Error: {str(e)}"))
            self.status_var.set("Error")
    
    def _send_command(self, command):
        """Send a command to the active EZCAD window"""
        if hasattr(self, 'current_window_id'):
            success = self.ezcad_controller.send_command(self.current_window_id, command)
            if success:
                self.logger.info(f"Sent {command} command")
                self.status_var.set(f"Sent {command} command")
            else:
                self.logger.error(f"Failed to send {command} command")
                messagebox.showerror("Command Error", f"Failed to send {command} command.")
        else:
            self.logger.warning("No active EZCAD window")
            messagebox.showwarning("EZCAD", "No active EZCAD window. Please start EZCAD first.")
    
    def _select_ezcad_window(self):
        """Manually select EZCAD window"""
        messagebox.showinfo("Select Window", "Please ensure EZCAD is running, then click OK to select the window.")
        
        # This would normally use platform-specific code to select a window
        # In our compatible version, we'll just simulate it
        if not PlatformUtils.is_windows():
            self.logger.info("SIMULATION: Selecting EZCAD window")
            self.current_window_id = f"sim_ezcad_{int(time.time() * 1000)}"
            
            with self.ezcad_controller.lock:
                self.ezcad_controller.instances[self.current_window_id] = {
                    'app': None,
                    'window': None,
                    'ezd_file': self.ezd_path_var.get(),
                    'start_time': time.time(),
                    'simulated': True
                }
            
            self.red_button.config(state=tk.NORMAL)
            self.mark_button.config(state=tk.NORMAL)
            messagebox.showinfo("Window Selected", "EZCAD window selected (simulation).")
            self.status_var.set("EZCAD window selected")
            return
            
        # On Windows, use the real implementation
        # [Windows implementation would go here]
    
    def _process_excel(self):
        """Process the selected Excel file"""
        excel_file = self.excel_path_var.get()
        if not excel_file:
            messagebox.showwarning("Missing File", "Please select an Excel file first.")
            return
        
        self.logger.info(f"Processing Excel file: {excel_file}")
        self.status_var.set("Processing Excel...")
        
        # Process in a separate thread to keep UI responsive
        threading.Thread(target=self._process_excel_thread, args=(excel_file,), daemon=True).start()
    
    def _process_excel_thread(self, excel_file):
        """Thread function to process Excel"""
        try:
            results = self.processor.process_excel(excel_file)
            
            # Update UI from the main thread
            self.root.after(0, lambda: self._show_excel_results(results))
            
        except Exception as e:
            self.logger.error(f"Error processing Excel: {str(e)}")
            self.root.after(0, lambda: messagebox.showerror("Processing Error", f"Error: {str(e)}"))
            self.root.after(0, lambda: self.status_var.set("Error processing Excel"))
    
    def _show_excel_results(self, results):
        """Show Excel processing results"""
        if 'error' in results:
            messagebox.showerror("Processing Error", f"Error: {results['error']}")
            self.status_var.set("Error")
            return
        
        # Format and show results
        msg = f"Processed {results.get('total_rows', 0)} rows\n"
        if 'batches' in results:
            msg += f"Successfully processed {results.get('successful_rows', 0)} rows\n"
            msg += f"Failed to process {results.get('failed_rows', 0)} rows\n"
            msg += f"Processed in {len(results['batches'])} batches"
        
        messagebox.showinfo("Processing Complete", msg)
        self.status_var.set("Processing complete")
    
    def _start_monitoring(self):
        """Start directory monitoring"""
        directory = self.watch_dir_var.get()
        if not directory or not os.path.isdir(directory):
            messagebox.showwarning("Invalid Directory", "Please select a valid directory to watch.")
            return
        
        # Update config from UI before starting
        self.config.set('Monitoring', 'enabled', 'true')
        self.config.set('Monitoring', 'watch_directory', directory)
        self.config.set('Monitoring', 'recursive', str(self.recursive_var.get()))
        self.config.set('Settings', 'file_pattern_excel', self.excel_pattern_var.get())
        self.config.set('Settings', 'file_pattern_ezd', self.ezd_pattern_var.get())
        self.config.save_config()
        
        # Start monitoring
        if self.directory_watcher.start_watching():
            self.logger.info(f"Started monitoring directory: {directory}")
            self.status_var.set("Monitoring active")
            
            # Update button states
            self.start_monitoring_button.config(state=tk.DISABLED)
            self.stop_monitoring_button.config(state=tk.NORMAL)
            
            # Add event message
            self._add_event(f"Started monitoring directory: {directory}")
        else:
            messagebox.showerror("Monitoring Error", "Failed to start directory monitoring.")
    
    def _stop_monitoring(self):
        """Stop directory monitoring"""
        self.directory_watcher.stop_watching()
        self.config.set('Monitoring', 'enabled', 'false')
        self.config.save_config()
        
        self.logger.info("Stopped directory monitoring")
        self.status_var.set("Monitoring stopped")
        
        # Update button states
        self.start_monitoring_button.config(state=tk.NORMAL)
        self.stop_monitoring_button.config(state=tk.DISABLED)
        
        # Add event message
        self._add_event("Stopped monitoring")
    
    def _add_event(self, message):
        """Add an event message to the events text"""
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        event_text = f"{timestamp} - {message}\n"
        
        self.events_text.config(state=tk.NORMAL)
        self.events_text.insert(tk.END, event_text)
        self.events_text.see(tk.END)
        self.events_text.config(state=tk.DISABLED)
    
    def _start_processing(self):
        """Start job processing"""
        self.queue_manager.start_processing()
        self.logger.info("Started job processing")
        self.status_var.set("Processing active")
        self._refresh_job_list()
        
        messagebox.showinfo("Processing", "Job processing started.")
    
    def _stop_processing(self):
        """Stop job processing"""
        self.queue_manager.stop_processing()
        self.processor.request_stop()
        self.logger.info("Stopped job processing")
        self.status_var.set("Processing stopped")
        self._refresh_job_list()
        
        messagebox.showinfo("Processing", "Job processing stopped.")
    
    def _refresh_job_list(self):
        """Refresh the job list display"""
        # Clear existing items
        for item in self.job_tree.get_children():
            self.job_tree.delete(item)
        
        # Get all jobs
        jobs = self.queue_manager.get_all_jobs()
        
        # Add each job to the treeview
        for job in jobs:
            # Calculate duration
            duration = ""
            if job.start_time and job.end_time:
                dur_seconds = (job.end_time - job.start_time).total_seconds()
                duration = f"{dur_seconds:.1f}s"
            elif job.start_time:
                dur_seconds = (datetime.now() - job.start_time).total_seconds()
                duration = f"{dur_seconds:.1f}s (running)"
            
            # Format time
            added_time = job.added_time.strftime("%Y-%m-%d %H:%M:%S")
            
            # Add to treeview
            self.job_tree.insert("", "end", values=(
                job.id,
                os.path.basename(job.file_path),
                job.job_type,
                job.status,
                added_time,
                duration
            ))
        
        self.logger.debug(f"Refreshed job list with {len(jobs)} jobs")
    
    def _show_job_context_menu(self, event):
        """Show context menu for job tree items"""
        # Get the item under cursor
        item = self.job_tree.identify_row(event.y)
        if not item:
            return
        
        # Select the item
        self.job_tree.selection_set(item)
        
        # Get the job ID
        job_id = self.job_tree.item(item, "values")[0]
        
        # Create a popup menu
        popup = tk.Menu(self.root, tearoff=0)
        popup.add_command(label="View Details", command=lambda: self._show_job_details_by_id(job_id))
        popup.add_command(label="Cancel Job", command=lambda: self._cancel_job(job_id))
        
        # Display the popup
        try:
            popup.tk_popup(event.x_root, event.y_root, 0)
        finally:
            popup.grab_release()
    
    def _show_job_details(self, event):
        """Show job details when an item is double-clicked"""
        item = self.job_tree.focus()
        if not item:
            return
            
        job_id = self.job_tree.item(item, "values")[0]
        self._show_job_details_by_id(job_id)
    
    def _show_job_details_by_id(self, job_id):
        """Show job details by job ID"""
        job = self.queue_manager.get_job(job_id)
        if not job:
            messagebox.showwarning("Job Not Found", f"Job {job_id} not found.")
            return
        
        # Create detail window
        detail_win = tk.Toplevel(self.root)
        detail_win.title(f"Job Details: {job_id}")
        detail_win.geometry("500x400")
        
        # Details frame
        frame = ttk.Frame(detail_win, padding=(10, 10))
        frame.pack(fill=tk.BOTH, expand=True)
        
        # Basic info
        ttk.Label(frame, text="Job ID:").grid(row=0, column=0, sticky="w", padx=5, pady=2)
        ttk.Label(frame, text=job.id).grid(row=0, column=1, sticky="w", padx=5, pady=2)
        
        ttk.Label(frame, text="File:").grid(row=1, column=0, sticky="w", padx=5, pady=2)
        ttk.Label(frame, text=job.file_path).grid(row=1, column=1, sticky="w", padx=5, pady=2)
        
        ttk.Label(frame, text="Type:").grid(row=2, column=0, sticky="w", padx=5, pady=2)
        ttk.Label(frame, text=job.job_type).grid(row=2, column=1, sticky="w", padx=5, pady=2)
        
        ttk.Label(frame, text="Status:").grid(row=3, column=0, sticky="w", padx=5, pady=2)
        ttk.Label(frame, text=job.status).grid(row=3, column=1, sticky="w", padx=5, pady=2)
        
        ttk.Label(frame, text="Added:").grid(row=4, column=0, sticky="w", padx=5, pady=2)
        ttk.Label(frame, text=job.added_time.strftime("%Y-%m-%d %H:%M:%S")).grid(row=4, column=1, sticky="w", padx=5, pady=2)
        
        if job.start_time:
            ttk.Label(frame, text="Started:").grid(row=5, column=0, sticky="w", padx=5, pady=2)
            ttk.Label(frame, text=job.start_time.strftime("%Y-%m-%d %H:%M:%S")).grid(row=5, column=1, sticky="w", padx=5, pady=2)
        
        if job.end_time:
            ttk.Label(frame, text="Completed:").grid(row=6, column=0, sticky="w", padx=5, pady=2)
            ttk.Label(frame, text=job.end_time.strftime("%Y-%m-%d %H:%M:%S")).grid(row=6, column=1, sticky="w", padx=5, pady=2)
            
            duration = (job.end_time - job.start_time).total_seconds()
            ttk.Label(frame, text="Duration:").grid(row=7, column=0, sticky="w", padx=5, pady=2)
            ttk.Label(frame, text=f"{duration:.2f} seconds").grid(row=7, column=1, sticky="w", padx=5, pady=2)
        
        # Result details
        ttk.Label(frame, text="Results:").grid(row=8, column=0, sticky="nw", padx=5, pady=5)
        
        result_text = scrolledtext.ScrolledText(frame, height=10, width=50)
        result_text.grid(row=9, column=0, columnspan=2, sticky="nsew", padx=5, pady=5)
        
        if job.error:
            result_text.insert(tk.END, f"ERROR: {job.error}\n")
        elif job.result:
            result_text.insert(tk.END, json.dumps(job.result, indent=2))
        else:
            result_text.insert(tk.END, "No results available.")
            
        result_text.config(state=tk.DISABLED)
        
        # Make result area expandable
        frame.columnconfigure(1, weight=1)
        frame.rowconfigure(9, weight=1)
        
        # Close button
        ttk.Button(detail_win, text="Close", command=detail_win.destroy).pack(pady=10)
    
    def _cancel_job(self, job_id):
        """Cancel a job by ID"""
        if self.queue_manager.cancel_job(job_id):
            self.logger.info(f"Canceled job {job_id}")
            self._refresh_job_list()
        else:
            messagebox.showwarning("Cancel Failed", f"Could not cancel job {job_id}. It may already be running or completed.")
    
    def _clear_completed_jobs(self):
        """Clear completed jobs from the queue"""
        count = self.queue_manager.clear_completed_jobs()
        self.logger.info(f"Cleared {count} completed jobs")
        self._refresh_job_list()
        
        messagebox.showinfo("Jobs Cleared", f"Cleared {count} completed jobs.")
    
    def _save_profile(self):
        """Save current settings as a profile"""
        profile_name = self.profile_name_var.get().strip()
        if not profile_name:
            messagebox.showwarning("Missing Name", "Please enter a profile name.")
            return
        
        try:
            # Apply current settings first
            self._apply_settings()
            
            # Save the profile
            profile_file = self.config.save_profile(profile_name)
            self.logger.info(f"Saved profile '{profile_name}' to {profile_file}")
            messagebox.showinfo("Profile Saved", f"Profile '{profile_name}' has been saved.")
            
        except Exception as e:
            self.logger.error(f"Error saving profile: {str(e)}")
            messagebox.showerror("Save Error", f"Error saving profile: {str(e)}")
    
    def _show_load_profile(self):
        """Show dialog to select a profile to load"""
        profiles = self.config.list_profiles()
        if not profiles:
            messagebox.showinfo("No Profiles", "No saved profiles found.")
            return
        
        # Create profile selection window
        profile_win = tk.Toplevel(self.root)
        profile_win.title("Load Profile")
        profile_win.geometry("300x300")
        profile_win.transient(self.root)
        profile_win.grab_set()
        
        ttk.Label(profile_win, text="Select a profile to load:").pack(padx=10, pady=5)
        
        # Create a listbox with profiles
        listbox = tk.Listbox(profile_win, width=40, height=10)
        listbox.pack(padx=10, pady=5, fill=tk.BOTH, expand=True)
        
        for profile in profiles:
            listbox.insert(tk.END, profile)
        
        # Buttons
        button_frame = ttk.Frame(profile_win)
        button_frame.pack(padx=10, pady=10, fill=tk.X)
        
        ttk.Button(
            button_frame, 
            text="Load", 
            command=lambda: self._load_selected_profile(listbox, profile_win)
        ).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            button_frame, 
            text="Cancel", 
            command=profile_win.destroy
        ).pack(side=tk.RIGHT, padx=5)
    
    def _load_selected_profile(self, listbox, window):
        """Load the selected profile"""
        selection = listbox.curselection()
        if not selection:
            messagebox.showwarning("No Selection", "Please select a profile to load.")
            return
        
        profile_name = listbox.get(selection[0])
        
        try:
            self.config.load_profile(profile_name)
            self.logger.info(f"Loaded profile: {profile_name}")
            
            # Refresh UI with new config
            self._refresh_from_config()
            
            window.destroy()
            messagebox.showinfo("Profile Loaded", f"Profile '{profile_name}' has been loaded.")
            
        except Exception as e:
            self.logger.error(f"Error loading profile: {str(e)}")
            messagebox.showerror("Load Error", f"Error loading profile: {str(e)}")
    
    def _open_log_directory(self):
        """Open the log directory in file explorer"""
        log_dir = os.path.join(os.getcwd(), "logs")
        if not os.path.exists(log_dir):
            os.makedirs(log_dir)
        
        # For cross-platform compatibility
        if PlatformUtils.is_windows():
            # Use subprocess instead of os.startfile for better compatibility
            import subprocess
            subprocess.Popen(['explorer', log_dir])
        else:
            self.logger.info(f"Log directory: {log_dir}")
            messagebox.showinfo("Log Directory", f"The log directory is located at:\n{log_dir}")
    
    def _start_automation(self):
        """Start all automated components"""
        try:
            # Start monitoring if enabled
            if self.config.getboolean('Monitoring', 'enabled', fallback=False):
                self.directory_watcher.start_watching()
                self.start_monitoring_button.config(state=tk.DISABLED)
                self.stop_monitoring_button.config(state=tk.NORMAL)
            
            # Start processing
            self.queue_manager.start_processing()
            
            self.logger.info("Started automation")
            self.status_var.set("Automation active")
            
        except Exception as e:
            self.logger.error(f"Error starting automation: {str(e)}")
            messagebox.showerror("Automation Error", f"Error starting automation: {str(e)}")
    
    def on_closing(self):
        """Handle window closing"""
        try:
            # Stop all background threads
            if hasattr(self, 'directory_watcher'):
                self.directory_watcher.stop_watching()
            
            if hasattr(self, 'queue_manager'):
                self.queue_manager.stop_processing()
            
            if hasattr(self, 'ezcad_controller'):
                self.ezcad_controller.close_all_ezcad()
            
            # Save config
            if hasattr(self, 'config'):
                self.config.save_config()
            
            # Stop log panel
            if hasattr(self, 'log_panel'):
                self.log_panel.stop()
            
            self.logger.info("Application closed")
            
        except Exception as e:
            print(f"Error during shutdown: {str(e)}")
            
        finally:
            # Close the window
            self.root.destroy()