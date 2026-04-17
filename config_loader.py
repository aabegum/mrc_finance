"""
config_loader.py
================
Configuration file loader for MRC Business Review generator.

Loads a single consolidated YAML configuration file and provides a unified
interface to access all settings without hardcoding values in the script.

Usage:
    from config_loader import ConfigLoader
    
    config = ConfigLoader()
    
    # Access theme colors
    primary_color = config.get('THEMES.premium.Colors.Primary.Hex')
    
    # Get active theme info
    theme_info = config.get_active_theme_template()
    
    # Access theme-specific settings
    title_size = config.get('THEMES.premium.Typography.Sizes.Title_Main')
"""

import os
import yaml
import json
from pathlib import Path
from typing import Any, Dict, Optional


class ConfigLoader:
    """
    Loads and provides access to a consolidated configuration file.
    
    The single config.yaml file contains all theme and app settings:
    - excel_mapping.yaml: Excel data source mappings
    - design.yaml: Colors, fonts, visual styling
    - layouts.yaml: Slide element positioning
    - chart_params.yaml: Chart rendering parameters
    """
    
    def __init__(self, config_dir: Optional[str] = None):
        """
        Initialize ConfigLoader.
        
        Args:
            config_dir: Path to config directory. Defaults to './config'
        """
        if config_dir is None:
            # Look for config directory relative to this script
            script_dir = os.path.dirname(os.path.abspath(__file__))
            config_dir = os.path.join(script_dir, 'config')
        
        self.config_dir = config_dir
        self._configs: Dict[str, Dict[str, Any]] = {}
        
        # Load all configuration files
        self._load_all_configs()
    
    def _load_all_configs(self) -> None:
        """Load the consolidated config.yaml file."""
        filepath = os.path.join(self.config_dir, 'config.yaml')
        if os.path.exists(filepath):
            self._load_config_file(filepath, 'config.yaml')
        else:
            print(f"ERROR: Consolidated config file not found: {filepath}")
            raise FileNotFoundError(f"Config file required: {filepath}")
    
    def _load_config_file(self, filepath: str, filename: str) -> None:
        """
        Load a single YAML configuration file.
        
        Args:
            filepath: Full path to config file
            filename: Config file name (e.g., 'config.yaml')
        """
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                config = yaml.safe_load(f)
                if config:
                    # For consolidated config.yaml, store top-level keys directly
                    if filename == 'config.yaml':
                        # Merge all top-level keys from consolidated config
                        for key, value in config.items():
                            self._configs[key] = value
                    else:
                        # For individual files, use filename without extension as key
                        key = filename.replace('.yaml', '')
                        self._configs[key] = config

                    print(f"[OK] Loaded config: {filename}")
        except yaml.YAMLError as e:
            print(f"ERROR: Failed to parse {filename}: {e}")
        except IOError as e:
            print(f"ERROR: Failed to read {filename}: {e}")
    
    def get(self, path: str, default: Any = None) -> Any:
        """
        Get a configuration value using dot notation.
        
        Args:
            path: Dot-separated path, e.g., 'design.COLORS.Primary.Hex'
            default: Default value if path not found
        
        Returns:
            Configuration value or default if not found
        
        Examples:
            config.get('design.COLORS.Primary.Hex')
            → '#00B4D8'
            
            config.get('layouts.CONTENT_SLIDE.Chart.Width_Wide')
            → 12.0
            
            config.get('excel_mapping.SUMMARY_TL.NS_kTL.Contract_Row')
            → 15
            
            config.get('nonexistent.path', default='fallback')
            → 'fallback'
        """
        parts = path.split('.')
        
        # First part is config file name
        if not parts:
            return default
        
        config_name = parts[0]
        if config_name not in self._configs:
            print(f"WARNING: Config section not found: {config_name}")
            return default
        
        # Navigate through nested dict
        value = self._configs[config_name]
        
        for part in parts[1:]:
            if isinstance(value, dict) and part in value:
                value = value[part]
            else:
                return default
        
        return value
    
    def get_all(self, section: str) -> Dict[str, Any]:
        """
        Get entire configuration section.
        
        Args:
            section: Section name, e.g., 'design.COLORS'
        
        Returns:
            Dictionary or default {} if not found
        
        Example:
            colors = config.get_all('design.COLORS')
            → {'Primary': {...}, 'Dark': {...}, ...}
        """
        value = self.get(section, default={})
        return value if isinstance(value, dict) else {}
    
    def list_configs(self) -> None:
        """Print available configuration files and top-level keys."""
        print("\n" + "="*60)
        print("LOADED CONFIGURATIONS")
        print("="*60)
        
        for config_name, config_data in self._configs.items():
            print(f"\n📄 {config_name}.yaml")
            if isinstance(config_data, dict):
                for key in config_data.keys():
                    print(f"   ├─ {key}")
    
    def export_to_dict(self) -> Dict[str, Dict[str, Any]]:
        """
        Export all configurations as single dictionary.
        
        Returns:
            Full configuration dictionary
        """
        return self._configs.copy()
    
    def reload(self) -> None:
        """Reload all configuration files."""
        self._configs.clear()
        self._load_all_configs()
    
    def get_active_theme_template(self) -> dict:
        """
        Get the template path and design config for the active theme.
        
        Reads active theme directly from APP.Active_Theme in config.yaml and returns
        the corresponding template path.
        
        Returns:
            Dict with keys:
            - 'theme': active theme name (e.g., 'default' or 'premium')
            - 'template': path to template PPTX file
            - 'design_config': path to design config YAML file (consolidated)
            - 'template_config': path to template config YAML file (consolidated)
        
        Example:
            >>> loader.get_active_theme_template()
            {
                'theme': 'premium',
                'template': 'template/MRC_Business_Review_Template_Premium.pptx',
                'design_config': 'config/config.yaml',
                'template_config': 'config/config.yaml'
            }
        """
        # Read active theme directly from config.yaml
        active_theme = self.get('APP.Active_Theme', 'default')
        
        # Get theme configuration from consolidated config
        try:
            template_path = self.get(f'THEMES.{active_theme}.Template.File')
            
            return {
                'theme': active_theme,
                'template': template_path,
                'design_config': 'config/config.yaml',
                'template_config': 'config/config.yaml'
            }
        except Exception as e:
            # Fallback if theme not found in config
            print(f"Warning: Could not load theme '{active_theme}' from config: {e}")
            print("Using classic template as fallback.")
            return {
                'theme': 'default',
                'template': 'template/MRC_Business_Review_Template.pptx',
                'design_config': 'config/config.yaml',
                'template_config': 'config/config.yaml'
            }


# ─────────────────────────────────────────────────────
# Color helper functions
# ─────────────────────────────────────────────────────

def hex_to_rgb(hex_color: str) -> tuple:
    """
    Convert hex color to RGB tuple.
    
    Args:
        hex_color: Color in hex format, e.g., '#00B4D8'
    
    Returns:
        Tuple of (R, G, B) values 0-255
    
    Example:
        hex_to_rgb('#00B4D8')
        → (0, 180, 216)
    """
    hex_color = hex_color.lstrip('#')
    return tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))


def rgb_to_hex(r: int, g: int, b: int) -> str:
    """
    Convert RGB to hex color.
    
    Args:
        r, g, b: RGB values 0-255
    
    Returns:
        Color in hex format, e.g., '#00B4D8'
    
    Example:
        rgb_to_hex(0, 180, 216)
        → '#00B4D8'
    """
    return f'#{r:02x}{g:02x}{b:02x}'.upper()


# ─────────────────────────────────────────────────────
# Dimension helper functions
# ─────────────────────────────────────────────────────

def inches_to_pixels(inches: float, dpi: int = 150) -> int:
    """
    Convert inches to pixels.
    
    Args:
        inches: Dimension in inches
        dpi: Dots per inch (default 150)
    
    Returns:
        Dimension in pixels
    
    Example:
        inches_to_pixels(1.0, 150)
        → 150
    """
    return int(inches * dpi)


def pixels_to_inches(pixels: int, dpi: int = 150) -> float:
    """
    Convert pixels to inches.
    
    Args:
        pixels: Dimension in pixels
        dpi: Dots per inch (default 150)
    
    Returns:
        Dimension in inches
    """
    return pixels / dpi


# ─────────────────────────────────────────────────────
# Usage example
# ─────────────────────────────────────────────────────

if __name__ == '__main__':
    # Example usage
    print("Initializing ConfigLoader...")
    config = ConfigLoader()
    
    # List all loaded configs
    config.list_configs()
    
    # Access specific values
    print("\n" + "="*60)
    print("EXAMPLE CONFIGURATION ACCESS")
    print("="*60)
    
    examples = [
        'design.COLORS.Primary.Hex',
        'layouts.CONTENT_SLIDE.Chart.Width_Wide',
        'excel_mapping.SUMMARY_TL.NS_kTL.Contract_Row',
        'chart_params.STACKED_CHART.Label.Threshold_Percent',
    ]
    
    for example_path in examples:
        value = config.get(example_path)
        print(f"\nconfig.get('{example_path}')")
        print(f"  → {value}")
    
    # Convert colors
    print("\n" + "="*60)
    print("COLOR CONVERSION EXAMPLES")
    print("="*60)
    print(f"\nhex_to_rgb('#00B4D8') → {hex_to_rgb('#00B4D8')}")
    print(f"rgb_to_hex(0, 180, 216) → {rgb_to_hex(0, 180, 216)}")
