import os
import time
import threading
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import logging
import queue

class FileChangeHandler(FileSystemEventHandler):
    """Handler for file system events"""
    
    def __init__(self, file_patterns, callback, logger=None):
        """
        Initialize the handler with file patterns to watch
        
        Args:
            file_patterns (list): List of file extensions to watch (e.g., ['.xls', '.xlsx'])
            callback (function): Function to call when a matching file changes
            logger: Logger instance
        """
        self.file_patterns = file_patterns
        self.callback = callback
        self.logger = logger or logging.getLogger('EZCADAutomation')
        self.last_events = {}  # Track last event times to handle duplicates
        
    def on_created(self, event):
        """Called when a file or directory is created"""
        if not event.is_directory and self._is_valid_file(event.src_path):
            self._handle_event('created', event.src_path)
    
    def on_modified(self, event):
        """Called when a file or directory is modified"""
        if not event.is_directory and self._is_valid_file(event.src_path):
            self._handle_event('modified', event.src_path)
    
    def _is_valid_file(self, path):
        """Check if the file has a monitored extension"""
        _, ext = os.path.splitext(path)
        return any(path.endswith(pattern) for pattern in self.file_patterns)
    
    def _handle_event(self, event_type, path):
        """Handle file event with debounce to prevent duplicate events"""
        current_time = time.time()
        
        # Skip if we've seen this file recently (within 2 seconds)
        if path in self.last_events and current_time - self.last_events[path] < 2:
            return
        
        # Update last seen time
        self.last_events[path] = current_time
        
        # Log and call the callback
        self.logger.info(f"File {event_type}: {path}")
        self.callback(path, event_type)


class DirectoryWatcher:
    """Watches directories for file changes and triggers processing"""
    
    def __init__(self, config_manager, file_queue, logger=None):
        """
        Initialize the directory watcher
        
        Args:
            config_manager: Configuration manager instance
            file_queue: Queue to add files for processing
            logger: Logger instance
        """
        self.config_manager = config_manager
        self.file_queue = file_queue
        self.logger = logger or logging.getLogger('EZCADAutomation')
        self.observer = None
        self.is_running = False
    
    def start_watching(self):
        """Start watching for file changes"""
        if self.is_running:
            self.logger.warning("Directory watcher is already running")
            return
        
        # Get configuration
        directory = self.config_manager.get('Monitoring', 'watch_directory')
        if not directory or not os.path.isdir(directory):
            self.logger.error(f"Invalid watch directory: {directory}")
            return False
        
        recursive = self.config_manager.getboolean('Monitoring', 'recursive', fallback=False)
        
        # Get file patterns
        excel_patterns = self._parse_patterns(
            self.config_manager.get('Settings', 'file_pattern_excel', fallback="*.xls;*.xlsx")
        )
        ezd_patterns = self._parse_patterns(
            self.config_manager.get('Settings', 'file_pattern_ezd', fallback="*.ezd")
        )
        file_patterns = excel_patterns + ezd_patterns
        
        # Create event handler and observer
        event_handler = FileChangeHandler(file_patterns, self._file_callback, self.logger)
        self.observer = Observer()
        self.observer.schedule(event_handler, directory, recursive=recursive)
        
        # Start the observer
        self.observer.start()
        self.is_running = True
        self.logger.info(f"Started watching directory: {directory} (recursive: {recursive})")
        self.logger.debug(f"Watching for file patterns: {file_patterns}")
        
        return True
    
    def stop_watching(self):
        """Stop watching for file changes"""
        if self.observer and self.is_running:
            self.observer.stop()
            self.observer.join()
            self.is_running = False
            self.logger.info("Directory watcher stopped")
    
    def _file_callback(self, file_path, event_type):
        """Callback for file system events"""
        try:
            if self.config_manager.getboolean('Monitoring', 'enabled', fallback=False):
                self.logger.info(f"Queuing file for processing: {file_path}")
                self.file_queue.put((file_path, event_type))
        except Exception as e:
            self.logger.error(f"Error in file change callback: {str(e)}")
    
    def _parse_patterns(self, pattern_string):
        """Parse file pattern string into list of extensions"""
        if not pattern_string:
            return []
            
        patterns = []
        for p in pattern_string.split(';'):
            p = p.strip()
            if p.startswith('*.'):
                patterns.append(p[1:])  # Remove the * from *.ext
            else:
                patterns.append(p)
        return patterns
