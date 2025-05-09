import platform
import sys

IS_WINDOWS = platform.system().lower() == 'windows'
IS_LINUX = platform.system().lower() == 'linux'
IS_MAC = platform.system().lower() == 'darwin'

class PlatformUtils:
    @staticmethod
    def is_windows():
        return IS_WINDOWS

    @staticmethod
    def is_linux():
        return IS_LINUX

    @staticmethod
    def is_mac():
        return IS_MAC

    @staticmethod
    def get_system_info():
        return {
            'system': platform.system(),
            'release': platform.release(),
            'version': platform.version(),
            'machine': platform.machine(),
            'processor': platform.processor(),
            'python_version': sys.version
        }