import os
import time
import logging
import threading
import platform
from platform_utils import PlatformUtils, WinCompatMock

# Import Windows-specific libraries only if on Windows
if PlatformUtils.is_windows():
    import psutil
    from pywinauto.application import Application
else:
    # Mock classes for non-Windows environments
    class MockPsUtil:
        @staticmethod
        def process_iter(attrs=None):
            return []
    
    class MockApp:
        def connect(self, **kwargs):
            return self
        
        def top_window(self):
            return MockWindow()
    
    class MockWindow:
        def child_window(self, **kwargs):
            return self
        
        def exists(self):
            return True
        
        def set_focus(self):
            pass
        
        def click(self):
            pass
        
        def minimize(self):
            pass
        
        def type_keys(self, keys, set_foreground=True):
            pass
        
        def close(self):
            pass
    
    # Create mock instances
    psutil = MockPsUtil()
    
    def Application(backend=None):
        return MockApp()

class EZCADController:
    """Control EZCAD2 application instances with cross-platform compatibility"""
    
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
        
        # Create mock handler for non-Windows platforms
        if not PlatformUtils.is_windows():
            self.win_mock = WinCompatMock(self.logger)
    
    def start_ezcad(self, ezd_file=None):
        """
        Start a new EZCAD2 instance
        
        Args:
            ezd_file: Optional EZD file to open
            
        Returns:
            window_id: ID for the EZCAD window or None if failed
        """
        exe_path = self.config_manager.get('Paths', 'ezcad_exe')
        self.logger.info(f"EZCAD exe_path: {exe_path}")
        self.logger.info(f"EZD file: {ezd_file}")
        # In non-Windows environment, simulate the behavior
        if not PlatformUtils.is_windows():
            window_id = f"sim_ezcad_{int(time.time() * 1000)}"
            self.logger.info(f"SIMULATION: Starting EZCAD with EZD file: {ezd_file}")
            
            with self.lock:
                self.instances[window_id] = {
                    'app': None,
                    'window': None,
                    'ezd_file': ezd_file,
                    'start_time': time.time(),
                    'simulated': True
                }
            
            return window_id
        
        # On Windows, use the actual implementation
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
                    # Try to find and accept license dialog
                    app = Application(backend="win32").connect(title_re=".*License.*|.*Agreement.*|.*Terms.*")
                    license_win = app.top_window()
                    agree_button = license_win.child_window(title_re=".*agree.*|.*Kabul.*|.*Annehmen.*|.*同意.*", control_type="Button")
                    if agree_button.exists():
                        license_win.set_focus()
                        agree_button.click()
                        time.sleep(0.5)
                        license_win.minimize()
                        time.sleep(0.2)
                        license_win.type_keys('{ENTER}', set_foreground=False)
                    
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
                            'start_time': time.time(),
                            'simulated': False
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
                
                # For simulated instances
                if not PlatformUtils.is_windows() or instance.get('simulated', False):
                    self.logger.info(f"SIMULATION: Closing EZCAD instance: {window_id}")
                    del self.instances[window_id]
                    return True
                
                # For real Windows instances
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
                
                # For simulated instances
                if not PlatformUtils.is_windows() or instance.get('simulated', False):
                    self.logger.info(f"SIMULATION: Sent {command} command to EZCAD window {window_id}")
                    return True
                
                # For real Windows instances    
                window = instance['window']
                
                if command.lower() == 'red':
                    window.set_focus()  # Pencereyi odakla
                    window.type_keys("{F1}", set_foreground=True)
                    self.logger.info(f"Sent RED command to window {window_id}")
                elif command.lower() == 'mark':
                    window.set_focus()  # Pencereyi odakla
                    window.type_keys("{F2}", set_foreground=True)
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