import time
import logging
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

class DirectoryWatcher:
    """Watch directory for file changes"""

    def __init__(self, config, file_queue, logger=None):
        self.config = config
        self.file_queue = file_queue
        self.logger = logger or logging.getLogger('EZCADAutomation')
        self.observer = None

    def start_watching(self):
        """Start watching the configured directory"""
        try:
            watch_dir = self.config.get('Monitoring', 'watch_directory')
            if not watch_dir:
                self.logger.error("Watch directory not configured")
                return False

            event_handler = FileChangeHandler(self.config, self.file_queue, self.logger)
            self.observer = Observer()
            self.observer.schedule(event_handler, watch_dir, recursive=self.config.getboolean('Monitoring', 'recursive'))
            self.observer.start()

            self.logger.info(f"Started watching directory: {watch_dir}")
            return True

        except Exception as e:
            self.logger.error(f"Error starting directory watcher: {str(e)}")
            return False

    def stop_watching(self):
        """Stop watching for changes"""
        if self.observer:
            self.observer.stop()
            self.observer.join()
            self.logger.info("Stopped watching directory")

class FileChangeHandler(FileSystemEventHandler):
    """Handle file system events"""

    def __init__(self, config, file_queue, logger):
        self.config = config
        self.file_queue = file_queue
        self.logger = logger

    def on_created(self, event):
        """Handle file creation events"""
        if event.is_directory:
            return

        file_path = event.src_path
        if self._is_valid_file(file_path):
            self.logger.info(f"New file detected: {file_path}")
            # Add to processing queue if auto-trigger is enabled
            if self.config.getboolean('Settings', 'auto_trigger'):
                self.file_queue.put((0, file_path))

    def _is_valid_file(self, file_path):
        """Check if the file matches configured patterns"""
        excel_patterns = self.config.get('Settings', 'file_pattern_excel').split(';')
        ezd_patterns = self.config.get('Settings', 'file_pattern_ezd').split(';')

        file_path = file_path.lower()
        return any(file_path.endswith(pat.strip('*').lower()) for pat in excel_patterns + ezd_patterns)