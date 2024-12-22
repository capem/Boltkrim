import json
import os

class ConfigManager:
    def __init__(self, config_file='config.json'):
        self.config_file = config_file
        self.default_config = {
            'source_folder': '',
            'processed_folder': '',
            'excel_file': '',
            'excel_sheet': '',
            'filter1_column': '',  
            'filter2_column': '',
            'output_template': '{processed_folder}/{filter1|str.upper} - {filter2|str.upper}.pdf'
        }
        self.config = self.default_config.copy()
        self.change_callbacks = []
        
    def add_change_callback(self, callback):
        """Add a callback to be called when config changes."""
        if callback not in self.change_callbacks:
            self.change_callbacks.append(callback)
            
    def remove_change_callback(self, callback):
        """Remove a previously added callback."""
        if callback in self.change_callbacks:
            self.change_callbacks.remove(callback)
            
    def _notify_callbacks(self):
        """Notify all registered callbacks about config changes."""
        for callback in self.change_callbacks:
            try:
                callback()
            except Exception as e:
                print(f"Error in config change callback: {str(e)}")
        
    def load_config(self):
        """Load configuration from file."""
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r') as f:
                    loaded_config = json.load(f)
                    # Update config with loaded values, keeping defaults for missing keys
                    self.config.update(loaded_config)
        except Exception as e:
            print(f"Error loading config: {str(e)}")
            # Keep default values if loading fails
            self.config = self.default_config.copy()
            
    def save_config(self):
        """Save current configuration to file."""
        try:
            with open(self.config_file, 'w') as f:
                json.dump(self.config, f, indent=4)
        except Exception as e:
            print(f"Error saving config: {str(e)}")
            
    def update_config(self, new_values):
        """Update configuration with new values."""
        self.config.update(new_values)
        self.save_config()
        self._notify_callbacks()
        
    def get_config(self):
        """Get current configuration."""
        return self.config.copy()
        
    def reset_config(self):
        """Reset configuration to defaults."""
        self.config = self.default_config.copy()
        self.save_config()
        self._notify_callbacks()
