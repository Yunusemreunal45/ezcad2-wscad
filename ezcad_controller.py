import os
import time
import psutil
import logging
from pywinauto.application import Application
import subprocess
import threading

class EZCADController:
    """Control EZCAD2 application instances"""
    
    def __init__(self, config_manager, logger=None):
        """
        Initialize the EZCAD controller
        
        Args:
            config_manager: Configuration manager
            logger: Logger instance
        """
        self.config_manager = config_manager
        self.logger = logger or logging.getLogger('EZCADAutomation')
        self.instances = {}  # Store active EZCAD instances: window_id -> info
        self.lock = threading.Lock()  # For thread safety
    
    def start_ezcad(self, ezd_file=None):
        """
        Start a new EZCAD2 instance
        
        Args:
            ezd_file: Optional EZD file to open
            
        Returns:
            window_id: ID for the EZCAD window or None if failed
        """
        try:
            exe_path = self.config_manager.get('Paths', 'ezcad_exe')
            if not exe_path or not os.path.isfile(exe_path):
                self.logger.error("EZCAD2.exe path not set or invalid")
                return None
            
            # Check for existing EZCAD instances if we don't want multiple
            if not self.config_manager.getboolean('Settings', 'multiple_instances', fallback=False):
                for proc in psutil.process_iter(['pid', 'name']):
                    if 'EZCAD2.exe' in proc.info['name']:
                        self.logger.warning("EZCAD2 already running - new instance not started")
                        return None
            
            # Build command line
            cmd = [exe_path]
            if ezd_file:
                if not os.path.isfile(ezd_file):
                    self.logger.error(f"EZD file not found: {ezd_file}")
                    return None
                cmd.append(ezd_file)
            
            # Start EZCAD
            self.logger.info(f"Starting EZCAD2: {' '.join(cmd)}")
            os.spawnv(os.P_NOWAIT, exe_path, cmd)
            
            # Handle license agreement dialog
            window_id = None
            for _ in range(15):
                time.sleep(1)
                try:
                    # Try to find and accept license dialog more aggressively
                    app = Application(backend="win32").connect(title_re=".*License.*|.*Agreement.*|.*Terms.*|.*EZCAD.*")
                    for window in app.windows():
                        try:
                            if any(text in window.window_text() for text in ["License", "Agreement", "Terms"]):
                                window.set_focus()
                                # Try multiple ways to accept
                                for button in window.children():
                                    if any(text in button.window_text().lower() for text in ["agree", "accept", "ok", "yes", "kabul"]):
                                        button.click()
                                        time.sleep(0.2)
                                # Press Enter for good measure
                                window.type_keys('{ENTER}')
                                time.sleep(0.2)
                        except Exception as e:
                            self.logger.debug(f"Window handling error: {str(e)}")
                    
                    # Try to connect to the EZCAD window
                    if ezd_file:
                        ezd_name = os.path.basename(ezd_file).replace(".", "\.")
                        app = Application().connect(title_re=f".*{ezd_name}.*")
                    else:
                        app = Application().connect(title_re=f".*EZCAD2.*")
                    
                    window = app.top_window()
                    window_id = f"ezcad_{int(time.time() * 1000)}"
                    
                    with self.lock:
                        self.instances[window_id] = {
                            'app': app,
                            'window': window,
                            'ezd_file': ezd_file,
                            'start_time': time.time()
                        }
                    
                    self.logger.info(f"EZCAD2 started with window ID: {window_id}")
                    break
                    
                except Exception as e:
                    self.logger.debug(f"Waiting for EZCAD window: {str(e)}")
            
            if not window_id:
                self.logger.error("Failed to connect to EZCAD window after launch")
            
            return window_id
            
        except Exception as e:
            self.logger.error(f"Error starting EZCAD: {str(e)}")
            return None
    
    def close_ezcad(self, window_id):
        """
        Close a specific EZCAD instance
        
        Args:
            window_id: Window ID from start_ezcad
            
        Returns:
            success: Whether the instance was closed
        """
        try:
            with self.lock:
                instance = self.instances.get(window_id)
                if not instance:
                    self.logger.warning(f"EZCAD instance not found: {window_id}")
                    return False
                
                # Try to close properly
                try:
                    window = instance['window']
                    window.close()
                    time.sleep(0.5)
                    
                    # Handle any save dialogs
                    try:
                        save_dialog = instance['app'].top_window()
                        save_dialog.type_keys("{ESC}")
                    except:
                        pass
                    
                    # Remove from tracked instances
                    del self.instances[window_id]
                    self.logger.info(f"Closed EZCAD instance: {window_id}")
                    return True
                    
                except Exception as e:
                    self.logger.error(f"Error closing EZCAD instance {window_id}: {str(e)}")
                    return False
        
        except Exception as e:
            self.logger.error(f"Error in close_ezcad: {str(e)}")
            return False
    
    def close_all_ezcad(self):
        """Close all tracked EZCAD instances"""
        with self.lock:
            window_ids = list(self.instances.keys())
        
        closed_count = 0
        for window_id in window_ids:
            if self.close_ezcad(window_id):
                closed_count += 1
        
        self.logger.info(f"Closed {closed_count} EZCAD instances")
        return closed_count
    
    def send_command(self, window_id, command):
        """
        Send a command to EZCAD instance
        
        Args:
            window_id: Window ID
            command: Command to send ('red', 'mark', etc.)
            
        Returns:
            success: Whether command was sent
        """
        try:
            with self.lock:
                instance = self.instances.get(window_id)
                if not instance:
                    self.logger.warning(f"EZCAD instance not found: {window_id}")
                    return False
                
                window = instance['window']
                
                if command.lower() == 'red':
                    window.type_keys("{F1}")
                    self.logger.info(f"Sent RED command to window {window_id}")
                elif command.lower() == 'mark':
                    window.type_keys("{F2}")
                    self.logger.info(f"Sent MARK command to window {window_id}")
                # Add more commands as needed
                else:
                    self.logger.warning(f"Unknown command: {command}")
                    return False
                
                return True
                
        except Exception as e:
            self.logger.error(f"Error sending command {command} to window {window_id}: {str(e)}")
            return False
    
    def get_active_instances(self):
        """Get all active EZCAD instances"""
        with self.lock:
            return self.instances.copy()
