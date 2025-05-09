import logging
import os
from logging.handlers import RotatingFileHandler
import tkinter as tk
from tkinter import scrolledtext
import threading
import queue
import datetime

class LoggerSetup:
    """Setup logging for the application with GUI integration and file output"""
    
    def __init__(self):
        """Initialize the logger"""
        self.log_dir = "logs"
        self.log_queue = queue.Queue()
        self._setup_log_directory()
        self._configure_logger()
        
    def _setup_log_directory(self):
        """Create log directory if it doesn't exist"""
        if not os.path.exists(self.log_dir):
            os.makedirs(self.log_dir)
            
    def _configure_logger(self):
        """Configure the logger with appropriate handlers and formatters"""
        self.logger = logging.getLogger('EZCADAutomation')
        self.logger.setLevel(logging.DEBUG)
        
        # Clear any existing handlers
        if self.logger.handlers:
            self.logger.handlers.clear()
        
        # Create a formatter
        formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
        
        # File handler for all logs (rotating file handler to manage size)
        log_file = os.path.join(self.log_dir, f"ezcad_automation_{datetime.datetime.now().strftime('%Y%m%d')}.log")
        file_handler = RotatingFileHandler(log_file, maxBytes=5*1024*1024, backupCount=5)
        file_handler.setLevel(logging.DEBUG)
        file_handler.setFormatter(formatter)
        self.logger.addHandler(file_handler)
        
        # Queue handler for GUI integration
        queue_handler = QueueHandler(self.log_queue)
        queue_handler.setLevel(logging.INFO)
        queue_handler.setFormatter(formatter)
        self.logger.addHandler(queue_handler)
    
    def get_logger(self):
        """Return the configured logger"""
        return self.logger
    
    def get_queue(self):
        """Return the log queue for GUI consumption"""
        return self.log_queue


class QueueHandler(logging.Handler):
    """Handler to route logging messages to a queue"""
    
    def __init__(self, log_queue):
        """Initialize with the queue to send logs to"""
        super().__init__()
        self.log_queue = log_queue
        
    def emit(self, record):
        """Put the log record into the queue"""
        self.log_queue.put(record)


class LogPanel:
    """A log panel that can be embedded in a Tkinter GUI"""
    
    def __init__(self, parent_frame, log_queue):
        """Initialize the log panel within a parent frame"""
        self.log_queue = log_queue
        self.parent_frame = parent_frame
        
        # Create a scrolled text widget for displaying logs
        self.log_widget = scrolledtext.ScrolledText(parent_frame, height=10, width=80)
        self.log_widget.grid(row=0, column=0, sticky="nsew")
        self.log_widget.config(state=tk.DISABLED)
        
        # Configure text tag for different log levels
        self.log_widget.tag_config('INFO', foreground="black")
        self.log_widget.tag_config('DEBUG', foreground="gray")
        self.log_widget.tag_config('WARNING', foreground="orange")
        self.log_widget.tag_config('ERROR', foreground="red")
        self.log_widget.tag_config('CRITICAL', foreground="red", background="yellow")
        
        # Start the queue consumer thread
        self.running = True
        self.queue_thread = threading.Thread(target=self._process_log_queue)
        self.queue_thread.daemon = True
        self.queue_thread.start()
    
    def _process_log_queue(self):
        """Process log records from the queue and display them in the widget"""
        while self.running:
            try:
                # Get a log record from the queue with a timeout
                record = self.log_queue.get(block=True, timeout=0.2)
                self._display_log(record)
            except queue.Empty:
                # If queue is empty, just continue the loop
                continue
            except Exception as e:
                # If any other error occurs, print it and continue
                print(f"Error processing log queue: {e}")
    
    def _display_log(self, record):
        """Display a log record in the widget"""
        msg = self.format_log_record(record)
        
        # Enable widget to insert text, then disable it again to make it read-only
        self.log_widget.config(state=tk.NORMAL)
        self.log_widget.insert(tk.END, msg + '\n', record.levelname)
        self.log_widget.see(tk.END)  # Scroll to end
        self.log_widget.config(state=tk.DISABLED)
    
    def format_log_record(self, record):
        """Format a log record for display"""
        return f"{record.asctime} - {record.levelname} - {record.message}"
    
    def clear(self):
        """Clear all logs from the display"""
        self.log_widget.config(state=tk.NORMAL)
        self.log_widget.delete(1.0, tk.END)
        self.log_widget.config(state=tk.DISABLED)
    
    def stop(self):
        """Stop the log processing thread"""
        self.running = False
