import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import os
import threading
import queue
import time
from datetime import datetime
import json

# Import our modules
from logger import LoggerSetup, LogPanel
from config_manager import ConfigManager
from excel_handler import ExcelHandler
from ezcad_controller import EZCADController
from processor import Processor
from queue_manager import QueueManager
from watcher import DirectoryWatcher

class EZCADAutomationApp:
    """The main EZCAD Automation application"""

    def __init__(self, root):
        """Initialize the application"""
        self.root = root
        root.title("EZCAD2 Automation")
        root.geometry("900x700")

        # Initialize logger
        self.logger_setup = LoggerSetup()
        self.logger = self.logger_setup.get_logger()
        self.log_queue = self.logger_setup.get_queue()

        # Initialize config
        self.config = ConfigManager()

        # Initialize components
        self.excel_handler = ExcelHandler(self.logger)
        self.ezcad_controller = EZCADController(self.config, self.logger)
        self.processor = Processor(self.excel_handler, self.ezcad_controller, self.config, self.logger)
        self.queue_manager = QueueManager(self.processor, self.config, self.logger)
        self.directory_watcher = DirectoryWatcher(self.config, self.queue_manager.file_queue, self.logger)

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
        self.ezcad_exe_var.set(self.config.get('Paths', 'ezcad_exe', ''))
        self.watch_dir_var.set(self.config.get('Monitoring', 'watch_directory', ''))

        # Settings
        self.auto_start_var.set(self.config.getboolean('Settings', 'auto_start', False))
        self.minimize_var.set(self.config.getboolean('Settings', 'minimize_on_start', False))
        self.monitor_enabled_var.set(self.config.getboolean('Monitoring', 'enabled', False))
        self.recursive_var.set(self.config.getboolean('Monitoring', 'recursive', False))
        self.auto_trigger_var.set(self.config.getboolean('Settings', 'auto_trigger', False))
        self.batch_process_var.set(self.config.getboolean('Settings', 'batch_process', False))
        self.multiple_instances_var.set(self.config.getboolean('Settings', 'multiple_instances', False))

        self.max_concurrent_var.set(str(self.config.getint('Settings', 'max_concurrent_processes', 1)))

        # File patterns
        self.excel_pattern_var.set(self.config.get('Settings', 'file_pattern_excel', '*.xls;*.xlsx'))
        self.ezd_pattern_var.set(self.config.get('Settings', 'file_pattern_ezd', '*.ezd'))

    def _apply_settings(self):
        """Apply settings from UI to config"""
        # Paths
        self.config.set('Paths', 'ezcad_exe', self.ezcad_exe_var.get())
        self.config.set('Monitoring', 'watch_directory', self.watch_dir_var.get())

        # Settings
        self.config.set('Settings', 'auto_start', str(self.auto_start_var.get()))
        self.config.set('Settings', 'minimize_on_start', str(self.minimize_var.get()))
        self.config.set('Monitoring', 'enabled', str(self.monitor_enabled_var.get()))
        self.config.set('Monitoring', 'recursive', str(self.recursive_var.get()))
        self.config.set('Settings', 'auto_trigger', str(self.auto_trigger_var.get()))
        self.config.set('Settings', 'batch_process', str(self.batch_process_var.get()))
        self.config.set('Settings', 'multiple_instances', str(self.multiple_instances_var.get()))

        self.config.set('Settings', 'max_concurrent_processes', self.max_concurrent_var.get())

        # File patterns
        self.config.set('Settings', 'file_pattern_excel', self.excel_pattern_var.get())
        self.config.set('Settings', 'file_pattern_ezd', self.ezd_pattern_var.get())

        # Save the config
        self.config.save_config()

        self.logger.info("Settings applied and saved")
        messagebox.showinfo("Settings", "Settings applied and saved")

    def _reset_settings(self):
        """Reset settings to default"""
        if messagebox.askyesno("Reset Settings", "Are you sure you want to reset all settings to default?"):
            # Create a new default config
            self.config._create_default_config()
            # Refresh UI
            self._refresh_from_config()
            self.logger.info("Settings reset to default")

    def _select_ezcad_exe(self):
        """Browse for EZCAD2.exe"""
        file_path = filedialog.askopenfilename(
            title="Select EZCAD2.exe", 
            filetypes=[("EZCAD2", "EZCAD2.exe")]
        )
        if file_path:
            self.ezcad_exe_var.set(file_path)
            self.config.set('Paths', 'ezcad_exe', file_path)
            self.config.save_config()
            self.logger.info(f"EZCAD2.exe path set: {file_path}")

    def _select_excel(self):
        """Browse for Excel file"""
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xls;*.xlsx")]
        )
        if file_path:
            self.excel_path_var.set(file_path)
            self.config.set('Paths', 'last_excel_file', file_path)
            self.config.set('Paths', 'last_excel_dir', os.path.dirname(file_path))
            self.config.save_config()

            # Load and preview Excel
            df = self.excel_handler.load_excel(file_path)
            if df is not None:
                preview = self.excel_handler.get_preview()

                self.preview_text.config(state=tk.NORMAL)
                self.preview_text.delete(1.0, tk.END)
                self.preview_text.insert(tk.END, preview)
                self.preview_text.config(state=tk.DISABLED)

    def _select_ezd(self):
        """Browse for EZD file"""
        file_path = filedialog.askopenfilename(
            title="Select EZD File",
            filetypes=[("EZD files", "*.ezd")]
        )
        if file_path:
            self.ezd_path_var.set(file_path)
            self.config.set('Paths', 'last_ezd_file', file_path)
            self.config.set('Paths', 'last_ezd_dir', os.path.dirname(file_path))
            self.config.save_config()
            self.logger.info(f"EZD file selected: {file_path}")
            # EZD dosyası seçilince butonları aktif yap
            self.red_button.config(state=tk.NORMAL)
            self.mark_button.config(state=tk.NORMAL)

    def _select_watch_dir(self):
        """Browse for directory to watch"""
        dir_path = filedialog.askdirectory(title="Select Directory to Watch")
        if dir_path:
            self.watch_dir_var.set(dir_path)
            self.config.set('Monitoring', 'watch_directory', dir_path)
            self.config.save_config()
            self.logger.info(f"Watch directory set: {dir_path}")

    def _run_ezcad(self):
        """Run EZCAD with the selected EZD file"""
        ezd_file = self.ezd_path_var.get()
        if not ezd_file:
            messagebox.showerror("Error", "Please select an EZD file first")
            return

        self.logger.info(f"Starting EZCAD with file: {ezd_file}")

        # Start EZCAD in a separate thread
        threading.Thread(target=self._run_ezcad_thread, args=(ezd_file,), daemon=True).start()

    def _run_ezcad_thread(self, ezd_file):
        """Thread function to run EZCAD"""
        try:
            window_id = self.ezcad_controller.start_ezcad(ezd_file)

            if window_id:
                # Enable the control buttons
                self.root.after(0, lambda: self.red_button.config(state=tk.NORMAL))
                self.root.after(0, lambda: self.mark_button.config(state=tk.NORMAL))

                # Store window ID for later use
                self.current_window_id = window_id

                self.logger.info(f"EZCAD started successfully with window ID: {window_id}")
                self.root.after(0, lambda: self.status_var.set(f"EZCAD running - Window ID: {window_id}"))
            else:
                self.logger.error("Failed to start EZCAD")
                self.root.after(0, lambda: messagebox.showerror("Error", "Failed to start EZCAD"))

        except Exception as e:
            self.logger.error(f"Error starting EZCAD: {str(e)}")
            self.root.after(0, lambda: messagebox.showerror("Error", f"Error starting EZCAD: {str(e)}"))

def _send_command(self, command):
    """Send a command to the active EZCAD window"""
    if not hasattr(self, 'current_window_id'):
        messagebox.showerror("Error", "No active EZCAD window")
        return

    try:
        # Show status
        self.status_var.set(f"Sending {command} command...")
        self.root.update_idletasks()

        # Send command through the controller
        success = self.ezcad_controller.send_command(self.current_window_id, command)

        if success:
            self.logger.info(f"Sent {command.upper()} command to EZCAD")
            self.status_var.set(f"{command.upper()} command sent successfully")

            # Automatically handle any post-command tasks
            if command.lower() == 'mark':
                self.logger.info("Mark operation completed")
                self.status_var.set("Mark operation completed")
            elif command.lower() == 'red':
                self.logger.info("Red laser operation completed")
                self.status_var.set("Red laser operation completed")
        else:
            raise Exception("Command failed to send")

    except Exception as e:
        self.logger.error(f"Failed to send {command.upper()} command: {str(e)}")
        messagebox.showerror("Error", f"Failed to send {command.upper()} command")
        self.status_var.set("Command failed")

    def _process_excel(self):
        """Process the selected Excel file"""
        excel_file = self.excel_path_var.get()
        if not excel_file:
            messagebox.showerror("Error", "Please select an Excel file first")
            return

        self.logger.info(f"Processing Excel file: {excel_file}")
        try:
            result = self.processor.process_file(excel_file)
            messagebox.showinfo("Success", f"Excel file processed successfully.\nRows processed: {result.get('rows_processed', 0)}")
        except Exception as e:
            self.logger.error(f"Failed to process Excel file: {str(e)}")
            messagebox.showerror("Error", f"Failed to process Excel file: {str(e)}")

    def _select_ezcad_window(self):
        """Manually select EZCAD window"""
        try:
            ezd_file = self.ezd_path_var.get()
            if not ezd_file:
                raise ValueError("No EZD file selected")

            instances = self.ezcad_controller.get_active_instances()
            if not instances:
                raise ValueError("No active EZCAD instances found")

            # If there's only one instance, use it
            if len(instances) == 1:
                self.current_window_id = list(instances.keys())[0]
                self.red_button.config(state=tk.NORMAL)
                self.mark_button.config(state=tk.NORMAL)
                self.logger.info(f"Selected EZCAD window: {self.current_window_id}")
                messagebox.showinfo("Success", "EZCAD window selected")
                return

            # If multiple instances, let user select
            window_list = "\n".join([f"{wid}: {info.get('ezd_file', 'Unknown')}" for wid, info in instances.items()])
            messagebox.showinfo("Multiple Windows", f"Multiple EZCAD windows found:\n{window_list}\n\nPlease use one of these IDs in the next prompt.")

            window_id = simpledialog.askstring("Select Window", "Enter window ID:")
            if not window_id or window_id not in instances:
                raise ValueError("Invalid window ID")

            self.current_window_id = window_id
            self.red_button.config(state=tk.NORMAL)
            self.mark_button.config(state=tk.NORMAL)
            self.logger.info(f"Selected EZCAD window: {self.current_window_id}")
            messagebox.showinfo("Success", "EZCAD window selected")

        except Exception as e:
            self.logger.error(f"Error selecting EZCAD window: {str(e)}")
            messagebox.showerror("Error", f"Error selecting EZCAD window: {str(e)}")

    def _start_monitoring(self):
        """Start directory monitoring"""
        # Apply current settings
        self.config.set('Monitoring', 'enabled', 'true')
        self.config.set('Monitoring', 'watch_directory', self.watch_dir_var.get())
        self.config.set('Monitoring', 'recursive', str(self.recursive_var.get()))
        self.config.set('Settings', 'auto_trigger', str(self.auto_trigger_var.get()))
        self.config.set('Settings', 'file_pattern_excel', self.excel_pattern_var.get())
        self.config.set('Settings', 'file_pattern_ezd', self.ezd_pattern_var.get())
        self.config.save_config()

        # Start the directory watcher
        if self.directory_watcher.start_watching():
            self.start_monitoring_button.config(state=tk.DISABLED)
            self.stop_monitoring_button.config(state=tk.NORMAL)
            self.status_var.set("Monitoring active")

            # Start job processing if not already running
            if not self.queue_manager.should_run:
                self.queue_manager.start_processing()

            # Add log to events text
            self._add_event("Monitoring started: " + self.watch_dir_var.get())
        else:
            messagebox.showerror("Error", "Failed to start monitoring")

    def _stop_monitoring(self):
        """Stop directory monitoring"""
        self.directory_watcher.stop_watching()
        self.config.set('Monitoring', 'enabled', 'false')
        self.config.save_config()

        self.start_monitoring_button.config(state=tk.NORMAL)
        self.stop_monitoring_button.config(state=tk.DISABLED)
        self.status_var.set("Monitoring stopped")

        # Add log to events text
        self._add_event("Monitoring stopped")

    def _add_event(self, message):
        """Add an event message to the events text"""
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        self.events_text.config(state=tk.NORMAL)
        self.events_text.insert(tk.END, f"{timestamp}: {message}\n")
        self.events_text.see(tk.END)
        self.events_text.config(state=tk.DISABLED)

    def _start_processing(self):
        """Start job processing"""
        if not self.queue_manager.should_run:
            self.queue_manager.start_processing()
            self.logger.info("Job processing started")
            self.status_var.set("Job processing active")
        else:
            self.logger.info("Job processing already active")

    def _stop_processing(self):
        """Stop job processing"""
        if self.queue_manager.should_run:
            self.queue_manager.stop_processing()
            self.logger.info("Job processing stopped")
            self.status_var.set("Job processing stopped")

    def _refresh_job_list(self):
        """Refresh the job list display"""
        # Clear current items
        for item in self.job_tree.get_children():
            self.job_tree.delete(item)

        # Get all jobs and add to the treeview
        for job in self.queue_manager.get_all_jobs():
            # Calculate duration if applicable
            duration = ""
            if job.start_time and job.end_time:
                duration = str(job.end_time - job.start_time).split('.')[0]  # Remove microseconds
            elif job.start_time:
                duration = str(datetime.now() - job.start_time).split('.')[0]

            # Format the added time
            added_time = job.added_time.strftime("%Y-%m-%d %H:%M:%S")

            # Get the file name only, not the full path
            file_name = os.path.basename(job.file_path)

            self.job_tree.insert("", "end", values=(
                job.id,
                file_name,
                job.job_type,
                job.status,
                added_time,
                duration
            ))

    def _show_job_context_menu(self, event):
        """Show context menu for job tree items"""
        # Select row under mouse
        iid = self.job_tree.identify_row(event.y)
        if iid:
            # Select the item
            self.job_tree.selection_set(iid)

            # Get the job ID
            job_id = self.job_tree.item(iid, "values")[0]

            # Create popup menu
            menu = tk.Menu(self.root, tearoff=0)
            menu.add_command(label="View Details", command=lambda: self._show_job_details_by_id(job_id))
            menu.add_command(label="Cancel Job", command=lambda: self._cancel_job(job_id))

            # Display the menu
            menu.post(event.x_root, event.y_root)

    def _show_job_details(self, event):
        """Show job details when an item is double-clicked"""
        iid = self.job_tree.focus()
        if iid:
            job_id = self.job_tree.item(iid, "values")[0]
            self._show_job_details_by_id(job_id)

    def _show_job_details_by_id(self, job_id):
        """Show job details by job ID"""
        job = self.queue_manager.get_job(job_id)
        if not job:
            messagebox.showerror("Error", f"Job not found: {job_id}")
            return

        # Create a new dialog window
        details_window = tk.Toplevel(self.root)
        details_window.title(f"Job Details: {job_id}")
        details_window.geometry("600x400")

        # Add job information
        info_frame = ttk.LabelFrame(details_window, text="Job Information")
        info_frame.pack(fill=tk.X, padx=10, pady=10)

        ttk.Label(info_frame, text=f"ID: {job.id}").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        ttk.Label(info_frame, text=f"Type: {job.job_type}").grid(row=0, column=1, sticky="w", padx=5, pady=5)
        ttk.Label(info_frame, text=f"Status: {job.status}").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        ttk.Label(info_frame, text=f"Priority: {job.priority}").grid(row=1, column=1, sticky="w", padx=5, pady=5)

        ttk.Label(info_frame, text=f"Added: {job.added_time}").grid(row=2, column=0, sticky="w", padx=5, pady=5)

        # Add timing information if available
        if job.start_time:
            ttk.Label(info_frame, text=f"Started: {job.start_time}").grid(row=3, column=0, sticky="w", padx=5, pady=5)
        if job.end_time:
            ttk.Label(info_frame, text=f"Completed: {job.end_time}").grid(row=3, column=1, sticky="w", padx=5, pady=5)
            duration = job.end_time - job.start_time
            ttk.Label(info_frame, text=f"Duration: {duration}").grid(row=4, column=0, sticky="w", padx=5, pady=5)

        # Add file path
        path_frame = ttk.LabelFrame(details_window, text="File Path")
        path_frame.pack(fill=tk.X, padx=10, pady=10)

        path_entry = ttk.Entry(path_frame, width=80)
        path_entry.pack(padx=5, pady=5, fill=tk.X)
        path_entry.insert(0, job.file_path)
        path_entry.config(state="readonly")

        # Add results/error information
        result_frame = ttk.LabelFrame(details_window, text="Results")
        result_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        result_text = scrolledtext.ScrolledText(result_frame)
        result_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        if job.error:
            result_text.insert(tk.END, f"ERROR:\n{job.error}")
            result_text.tag_configure("error", foreground="red")
            result_text.tag_add("error", "1.0", "end")
        elif job.result:
            if isinstance(job.result, dict):
                result_text.insert(tk.END, json.dumps(job.result, indent=2))
            else:
                result_text.insert(tk.END, str(job.result))
        else:
            result_text.insert(tk.END, "No results available")

        result_text.config(state=tk.DISABLED)

        # Close button
        ttk.Button(details_window, text="Close", command=details_window.destroy).pack(pady=10)

    def _cancel_job(self, job_id):
        """Cancel a job by ID"""
        job = self.queue_manager.get_job(job_id)
        if job and job.status == "PENDING":
            if self.queue_manager.cancel_job(job_id):
                self.logger.info(f"Canceled job: {job_id}")
                self._refresh_job_list()
            else:
                messagebox.showerror("Error", f"Failed to cancel job: {job_id}")
        else:
            messagebox.showinfo("Info", "Job cannot be canceled (not in pending state)")

    def _clear_completed_jobs(self):
        """Clear completed jobs from the queue"""
        count = self.queue_manager.clear_completed_jobs()
        self.logger.info(f"Cleared {count} completed jobs")
        self._refresh_job_list()

    def _save_profile(self):
        """Save current settings as a profile"""
        profile_name = self.profile_name_var.get().strip()
        if not profile_name:
            messagebox.showerror("Error", "Please enter a profile name")
            return

        try:
            # Apply current settings first
            self._apply_settings()

            # Save profile
            profile_file = self.config.save_profile(profile_name)
            self.logger.info(f"Saved profile: {profile_name} to {profile_file}")
            messagebox.showinfo("Profile Saved", f"Profile '{profile_name}' saved successfully")

        except Exception as e:
            self.logger.error(f"Error saving profile: {str(e)}")
            messagebox.showerror("Error", f"Failed to save profile: {str(e)}")

    def _show_load_profile(self):
        """Show dialog to select a profile to load"""
        profiles = self.config.list_profiles()
        if not profiles:
            messagebox.showinfo("No Profiles", "No saved profiles found")
            return

        # Create a dialog window
        profile_window = tk.Toplevel(self.root)
        profile_window.title("Load Profile")
        profile_window.geometry("300x400")

        # Profile list
        ttk.Label(profile_window, text="Select a profile to load:").pack(padx=10, pady=10)

        profile_listbox = tk.Listbox(profile_window)
        profile_listbox.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        for profile in profiles:
            profile_listbox.insert(tk.END, profile)

        # Buttons
        button_frame = ttk.Frame(profile_window)
        button_frame.pack(fill=tk.X, padx=10, pady=10)

        ttk.Button(button_frame, text="Load", command=lambda: self._load_selected_profile(profile_listbox, profile_window)).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Cancel", command=profile_window.destroy).pack(side=tk.RIGHT, padx=5)

    def _load_selected_profile(self, listbox, window):
        """Load the selected profile"""
        selection = listbox.curselection()
        if not selection:
            messagebox.showerror("Error", "Please select a profile")
            return

        profile_name = listbox.get(selection[0])

        try:
            self.config.load_profile(profile_name)
            self.logger.info(f"Loaded profile: {profile_name}")

            # Refresh UI with loaded settings
            self._refresh_from_config()

            window.destroy()
            messagebox.showinfo("Profile Loaded", f"Profile '{profile_name}' loaded successfully")

        except Exception as e:
            self.logger.error(f"Error loading profile: {str(e)}")
            messagebox.showerror("Error", f"Failed to load profile: {str(e)}")

    def _open_log_directory(self):
        """Open the log directory in file explorer"""
        log_dir = os.path.join(os.getcwd(), "logs")
        if os.path.exists(log_dir):
            if os.name == 'nt':  # Windows
                os.startfile(log_dir)
            elif os.name == 'posix':  # macOS, Linux
                import subprocess
                subprocess.Popen(['xdg-open', log_dir])
        else:
            messagebox.showerror("Error", "Log directory not found")

    def _start_automation(self):
        """Start all automated components"""
        # Start job processing
        self.queue_manager.start_processing()

        # Start directory monitoring if enabled
        if self.config.getboolean('Monitoring', 'enabled', False):
            self.directory_watcher.start_watching()
            self.start_monitoring_button.config(state=tk.DISABLED)
            self.stop_monitoring_button.config(state=tk.NORMAL)

        self.logger.info("Automation started")
        self.status_var.set("Automation active")

    def on_closing(self):
        """Handle window closing"""
        if messagebox.askyesno("Exit", "Are you sure you want to exit?"):
            # Stop monitoring and processing
            self.directory_watcher.stop_watching()
            self.queue_manager.stop_processing()

            # Close any open EZCAD instances
            self.ezcad_controller.close_all_ezcad()

            # Stop the log panel
            self.log_panel.stop()

            # Destroy the window
            self.root.destroy()