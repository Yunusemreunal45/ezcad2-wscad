import os
import time
import psutil
import logging
import subprocess
import threading
from platform_utils import PlatformUtils, IS_WINDOWS

if IS_WINDOWS:
    from pywinauto.application import Application

class EZCADController:
    """Control EZCAD2 application instances"""

    def __init__(self, config_manager, logger=None):
        """Initialize the EZCAD controller"""
        self.config_manager = config_manager
        self.logger = logger or logging.getLogger('EZCADAutomation')
        self.instances = {}  # Store active EZCAD instances: window_id -> info
        self.lock = threading.Lock()  # For thread safety

        if not IS_WINDOWS:
            self.logger.warning("Running in non-Windows environment. EZCAD features will be simulated.")

    def start_ezcad(self, ezd_file=None):
        """Start a new EZCAD2 instance"""
        try:
            exe_path = self.config_manager.get('Paths', 'ezcad_exe')
            if not exe_path or not os.path.isfile(exe_path):
                self.logger.error("EZCAD2.exe path not set or invalid")
                return None

            if not self.config_manager.getboolean('Settings', 'multiple_instances', fallback=False):
                for proc in psutil.process_iter(['pid', 'name']):
                    if 'EZCAD2.exe' in proc.info['name']:
                        self.logger.warning("EZCAD2 already running - new instance not started")
                        return None

            cmd = [exe_path]
            if ezd_file:
                if not os.path.isfile(ezd_file):
                    self.logger.error(f"EZD file not found: {ezd_file}")
                    return None
                cmd.append(ezd_file)

            try:
                process = os.spawnv(os.P_NOWAIT, exe_path, cmd)
                if process:
                    self.logger.info("EZCAD2 process started successfully")
            except Exception as e:
                self.logger.error(f"Failed to start EZCAD2 process: {str(e)}")
                return None

            window_id = None
            # First try to find and handle the I Agree dialog
            for _ in range(5):  # Try for 5 seconds
                time.sleep(1)
                try:
                    agree_app = Application().connect(title_re=".*I Agree.*", timeout=1)
                    agree_window = agree_app.top_window()
                    if agree_window.is_visible():
                        # Try multiple ways to close the dialog
                        try:
                            agree_window.type_keys("{ENTER}")
                        except:
                            try:
                                agree_window.type_keys(" ")  # Space key
                            except:
                                try:
                                    agree_button = agree_window.child_window(title="I Agree")
                                    agree_button.click()
                                except:
                                    pass
                        time.sleep(0.5)
                        break
                except:
                    continue

            # Then look for the main EZCAD window
            for _ in range(10):
                time.sleep(1)
                try:

                    if ezd_file:
                        ezd_name = os.path.basename(ezd_file).replace(".", "\.")
                        app = Application().connect(title_re=f".*{ezd_name}.*")
                    else:
                        app = Application().connect(title_re=f".*EZCAD2.*")

                    window = app.top_window()
                    window.set_focus()
                    time.sleep(0.5)
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

    def send_command(self, window_id, command):
        """Send a command to EZCAD instance"""
        try:
            with self.lock:
                instance = self.instances.get(window_id)
                if not instance:
                    self.logger.warning(f"EZCAD instance not found: {window_id}")
                    return False

                window = instance['window']
                app = instance['app']

                # Verify window is still valid
                try:
                    window.wait('visible', timeout=5)
                except Exception as e:
                    self.logger.error(f"Window validation failed: {str(e)}")
                    return False

                # Check if EZCAD2 is still running
                if not app.is_process_running():
                    self.logger.error("EZCAD2 process is not running")
                    return False

                # Ensure window is visible and active
                try:
                    if not window.is_visible():
                        window.restore()
                        time.sleep(1.0)

                    # Try multiple times to set focus
                    max_attempts = 3
                    for attempt in range(max_attempts):
                        window.set_focus()
                        time.sleep(0.5)

                        if window.is_active():
                            break

                        if attempt == max_attempts - 1:
                            self.logger.error("Failed to activate EZCAD window after multiple attempts")
                            return False

                        time.sleep(1.0)

                    command = command.lower()
                    if command == 'red':
                        window.set_focus()
                        time.sleep(0.5)
                        window.type_keys("{F1}")
                        self.logger.info(f"Sent RED command to window {window_id}")
                        time.sleep(1.0)  # Wait for command to take effect
                        return True

                    elif command == 'mark':
                        window.set_focus()
                        time.sleep(0.5)
                        window.type_keys("{F2}")
                        self.logger.info(f"Sent MARK command to window {window_id}")
                        time.sleep(1.0)  # Wait for command to take effect
                        return True

                    else:
                        self.logger.warning(f"Unknown command: {command}")
                        return False

                except Exception as e:
                    self.logger.error(f"Error sending command {command} to window {window_id}: {str(e)}")
                    return False

        except Exception as e:
            self.logger.error(f"Error sending command {command} to window {window_id}: {str(e)}")
            return False

    def close_ezcad(self, window_id):
        """Close a specific EZCAD instance"""
        try:
            with self.lock:
                instance = self.instances.get(window_id)
                if not instance:
                    self.logger.warning(f"EZCAD instance not found: {window_id}")
                    return False

                try:
                    window = instance['window']
                    window.close()
                    time.sleep(0.5)

                    try:
                        save_dialog = instance['app'].top_window()
                        save_dialog.type_keys("{ESC}")
                    except:
                        pass

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

    def get_active_instances(self):
        """Get all active EZCAD instances"""
        with self.lock:
            return self.instances.copy()