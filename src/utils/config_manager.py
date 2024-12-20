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
            'filter2_column': ''   
        }
        self.config = self.default_config.copy()
        
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
        
    def get_config(self):
        """Get current configuration."""
        return self.config.copy()
        
    def reset_config(self):
        """Reset configuration to defaults."""
        self.config = self.default_config.copy()
        self.save_config()
