"""Configuration handling for the DienstplanSplitter printer service."""

import json
import os
from pathlib import Path
from typing import Dict, Any, Optional

DEFAULT_CONFIG = {
    "printer_name": "Dienstplan Splitter",
    "output_dir": str(Path.home() / "Documents" / "Dienstplan Splitter"),
    "port_name": "DS_PORT:",  # Virtual printer port name
    "driver_name": "Microsoft Print To PDF",  # Default driver until we implement our own
    "log_level": "INFO",
    "auto_start": True,
}

class PrinterConfig:
    """Handles configuration loading, saving, and access."""
    
    def __init__(self, config_dir: Optional[Path] = None):
        """Initialize configuration.
        
        Args:
            config_dir: Directory to store configuration. Defaults to user's app data directory.
        """
        if config_dir is None:
            config_dir = Path(os.getenv('APPDATA', str(Path.home() / '.config'))) / "DienstplanSplitter"
            
        self.config_dir = config_dir
        self.config_file = self.config_dir / "printer_config.json"
        self.config: Dict[str, Any] = {}
        self._load_config()
        
    def _load_config(self) -> None:
        """Load configuration from file or create default if not exists."""
        try:
            self.config_dir.mkdir(parents=True, exist_ok=True)
            if self.config_file.exists():
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    stored_config = json.load(f)
                    # Merge with defaults to ensure all required keys exist
                    self.config = {**DEFAULT_CONFIG, **stored_config}
            else:
                self.config = DEFAULT_CONFIG.copy()
                self._save_config()
        except Exception as e:
            print(f"Error loading config: {e}. Using defaults.")
            self.config = DEFAULT_CONFIG.copy()
            
    def _save_config(self) -> None:
        """Save current configuration to file."""
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, indent=4)
        except Exception as e:
            print(f"Error saving config: {e}")
            
    def get(self, key: str, default: Any = None) -> Any:
        """Get a configuration value.
        
        Args:
            key: Configuration key to retrieve
            default: Default value if key doesn't exist
            
        Returns:
            The configuration value or default
        """
        return self.config.get(key, default)
        
    def set(self, key: str, value: Any) -> None:
        """Set a configuration value and save to file.
        
        Args:
            key: Configuration key to set
            value: Value to set
        """
        self.config[key] = value
        self._save_config()
        
    def get_output_dir(self) -> Path:
        """Get the configured output directory as a Path object.
        
        Returns:
            Path object for the output directory
        """
        return Path(self.get('output_dir'))
        
    @property
    def printer_name(self) -> str:
        """Get the configured printer name.
        
        Returns:
            The printer name as it should appear in Windows
        """
        return self.get('printer_name') 