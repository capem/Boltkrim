from typing import Dict, List, Callable, Optional
from json import load as json_load, dump as json_dump
from os import path

class ConfigManager:
    """Manages configuration settings for the application with file persistence and change notifications."""
    
    def __init__(self, config_file: str = 'config.json') -> None:
        """Initialize the config manager with a default or specified config file."""
        self.config_file: str = config_file
        self.presets_file: str = 'presets.json'  # File to store preset configurations
        self.default_config: Dict[str, str] = {
            'source_folder': '',
            'processed_folder': '',
            'excel_file': '',
            'excel_sheet': '',
            'filter1_column': '',  
            'filter2_column': '',
            'filter3_column': '',  # Added for third filter
            'output_template': '{processed_folder}/{filter1|str.upper} - {filter2|str.upper}.pdf'
        }
        self.config: Dict[str, str] = self.default_config.copy()
        self.presets: Dict[str, Dict[str, str]] = {}  # Store preset configurations
        self.change_callbacks: List[Callable[[], None]] = []
        
        # Load both config and presets
        self.load_config()
        self.load_presets()
        
    def add_change_callback(self, callback: Callable[[], None]) -> None:
        """Add a callback to be called when config changes.
        
        Args:
            callback: A function taking no arguments and returning nothing
        """
        if callback not in self.change_callbacks:
            self.change_callbacks.append(callback)
            
    def remove_change_callback(self, callback: Callable[[], None]) -> None:
        """Remove a previously added callback.
        
        Args:
            callback: The callback function to remove
        """
        if callback in self.change_callbacks:
            self.change_callbacks.remove(callback)
            
    def _notify_callbacks(self) -> None:
        """Notify all registered callbacks about config changes."""
        for callback in self.change_callbacks:
            try:
                callback()
            except Exception as e:
                print(f"Error in config change callback: {str(e)}")
        
    def load_config(self) -> None:
        """Load configuration from file.
        
        If the file doesn't exist or there's an error, keeps default values.
        """
        try:
            if path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    loaded_config: Dict[str, str] = json_load(f)
                    # Update config with loaded values, keeping defaults for missing keys
                    self.config.update(loaded_config)
        except Exception as e:
            print(f"Error loading config: {str(e)}")
            # Keep default values if loading fails
            self.config = self.default_config.copy()
            
    def save_config(self) -> None:
        """Save current configuration to file.
        
        Creates the file if it doesn't exist, overwrites if it does.
        """
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json_dump(self.config, f, indent=4)
        except Exception as e:
            print(f"Error saving config: {str(e)}")
            
    def update_config(self, new_values: Dict[str, str]) -> None:
        """Update configuration with new values.
        
        Args:
            new_values: Dictionary of new configuration values to update
        """
        self.config.update(new_values)
        self.save_config()
        self._notify_callbacks()
        
    def get_config(self) -> Dict[str, str]:
        """Get current configuration.
        
        Returns:
            A copy of the current configuration dictionary
        """
        return self.config.copy()
        
    def reset_config(self) -> None:
        """Reset configuration to defaults."""
        self.config = self.default_config.copy()
        self.save_config()
        self._notify_callbacks()
        
    def load_presets(self) -> None:
        """Load preset configurations from file."""
        try:
            if path.exists(self.presets_file):
                with open(self.presets_file, 'r', encoding='utf-8') as f:
                    self.presets = json_load(f)
        except Exception as e:
            print(f"Error loading presets: {str(e)}")
            self.presets = {}
            
    def save_presets(self) -> None:
        """Save preset configurations to file."""
        try:
            with open(self.presets_file, 'w', encoding='utf-8') as f:
                json_dump(self.presets, f, indent=4)
        except Exception as e:
            print(f"Error saving presets: {str(e)}")
            
    def get_preset_names(self) -> List[str]:
        """Get list of available preset names.
        
        Returns:
            List of preset configuration names
        """
        return list(self.presets.keys())
        
    def get_preset(self, preset_name: str) -> Optional[Dict[str, str]]:
        """Get a specific preset configuration.
        
        Args:
            preset_name: Name of the preset to retrieve
            
        Returns:
            Preset configuration dictionary or None if not found
        """
        return self.presets.get(preset_name)
        
    def save_preset(self, preset_name: str, config: Dict[str, str]) -> None:
        """Save a new preset configuration.
        
        Args:
            preset_name: Name for the preset configuration
            config: Configuration dictionary to save as preset
        """
        self.presets[preset_name] = config.copy()
        self.save_presets()
        
    def delete_preset(self, preset_name: str) -> None:
        """Delete a preset configuration.
        
        Args:
            preset_name: Name of the preset to delete
        """
        if preset_name in self.presets:
            del self.presets[preset_name]
            self.save_presets()
