import platform
import logging
import os
import sys

# Global platform detection
IS_WINDOWS = platform.system() == 'Windows'

class PlatformUtils:
    """Utility class to handle platform-specific operations"""
    
    @staticmethod
    def is_windows():
        """Check if running on Windows"""
        return IS_WINDOWS
    
    @staticmethod
    def get_logger():
        """Get a logger for platform utilities"""
        return logging.getLogger('EZCADAutomation')
    
    @staticmethod
    def get_platform_info():
        """Get detailed platform information"""
        info = {
            'system': platform.system(),
            'release': platform.release(),
            'version': platform.version(),
            'machine': platform.machine(),
            'processor': platform.processor(),
            'python_version': platform.python_version(),
            'architecture': platform.architecture()[0],
        }
        
        return f"{info['system']} {info['release']} ({info['architecture']}), Python {info['python_version']}"

class WinCompatMock:
    """Mock class for Windows-specific functions in non-Windows environments"""
    
    def __init__(self, logger=None):
        """Initialize with optional logger"""
        self.logger = logger or PlatformUtils.get_logger()
        if not IS_WINDOWS:
            self.logger.warning("Running in non-Windows environment. EZCAD automation will run in simulation mode.")
    
    def mock_function(self, function_name, *args, **kwargs):
        """Generic mock function that logs the call and returns a success result"""
        self.logger.info(f"SIMULATION: Called {function_name} with args={args} kwargs={kwargs}")
        return True