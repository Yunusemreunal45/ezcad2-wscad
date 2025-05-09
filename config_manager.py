import os
import configparser
import json
from datetime import datetime

class ConfigManager:
    """Manage application configuration with save/load functionality"""
    
    def __init__(self, config_file="ezcad_config.ini", profiles_dir="profiles"):
        """Initialize the config manager with the config file path"""
        self.config_file = config_file
        self.profiles_dir = profiles_dir
        self.config = configparser.ConfigParser()
        
        # Ensure profiles directory exists
        if not os.path.exists(self.profiles_dir):
            os.makedirs(self.profiles_dir)
        
        # Load configuration or create default
        self.load_config()
    
    def load_config(self):
        """Load configuration from file or create with defaults"""
        if os.path.exists(self.config_file):
            self.config.read(self.config_file)
        else:
            self._create_default_config()
    
    def _create_default_config(self):
        """Create a default configuration file"""
        # Paths section
        self.config["Paths"] = {
            "ezcad_exe": "",
            "last_excel_dir": "",
            "last_ezd_dir": "",
            "watch_directory": ""
        }
        
        # Settings section
        self.config["Settings"] = {
            "auto_start": "false",
            "minimize_on_start": "false",
            "auto_trigger": "false",
            "batch_process": "false",
            "max_concurrent_processes": "1",
            "file_pattern_excel": "*.xls;*.xlsx",
            "file_pattern_ezd": "*.ezd"
        }
        
        # Monitoring section
        self.config["Monitoring"] = {
            "enabled": "false",
            "interval_seconds": "5",
            "recursive": "false"
        }
        
        # Save default configuration
        self.save_config()
    
    def save_config(self):
        """Save configuration to file"""
        with open(self.config_file, "w") as f:
            self.config.write(f)
    
    def get(self, section, key, fallback=None):
        """Get a configuration value"""
        return self.config.get(section, key, fallback=fallback)
    
    def getboolean(self, section, key, fallback=False):
        """Get a boolean configuration value"""
        return self.config.getboolean(section, key, fallback=fallback)
    
    def getint(self, section, key, fallback=0):
        """Get an integer configuration value"""
        return self.config.getint(section, key, fallback=fallback)
    
    def set(self, section, key, value):
        """Set a configuration value"""
        if not self.config.has_section(section):
            self.config.add_section(section)
        self.config.set(section, key, str(value))
    
    def save_profile(self, profile_name):
        """Save current configuration as a named profile"""
        if not profile_name:
            raise ValueError("Profile name cannot be empty")
        
        # Convert config to dictionary
        config_dict = {}
        for section in self.config.sections():
            config_dict[section] = dict(self.config[section])
        
        # Add metadata
        config_dict["Metadata"] = {
            "created": datetime.now().isoformat(),
            "profile_name": profile_name
        }
        
        # Save to file
        profile_file = os.path.join(self.profiles_dir, f"{profile_name}.json")
        with open(profile_file, "w") as f:
            json.dump(config_dict, f, indent=4)
        
        return profile_file
    
    def load_profile(self, profile_name):
        """Load a named profile"""
        profile_file = os.path.join(self.profiles_dir, f"{profile_name}.json")
        
        if not os.path.exists(profile_file):
            raise FileNotFoundError(f"Profile '{profile_name}' not found")
        
        # Load profile from file
        with open(profile_file, "r") as f:
            config_dict = json.load(f)
        
        # Create new config parser and populate it
        new_config = configparser.ConfigParser()
        for section, options in config_dict.items():
            if section != "Metadata":  # Skip metadata section
                if not new_config.has_section(section):
                    new_config.add_section(section)
                for key, value in options.items():
                    new_config.set(section, key, str(value))
        
        # Replace current config with loaded one
        self.config = new_config
        self.save_config()
    
    def list_profiles(self):
        """List all available profiles"""
        profiles = []
        for file in os.listdir(self.profiles_dir):
            if file.endswith(".json"):
                profile_name = os.path.splitext(file)[0]
                profiles.append(profile_name)
        return profiles
