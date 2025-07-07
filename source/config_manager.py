import os
import yaml
from typing import Tuple, Dict, Any, Optional


class ConfigManager:
    
    def __init__(self, config_path: str):
        self.config_path = config_path
        self.config = None
    
    def load_config(self) -> Tuple[Optional[Dict[str, Any]], Optional[Exception]]:
        try:
            with open(self.config_path, 'r', encoding='utf-8') as file:
                self.config = yaml.safe_load(file)
            return self.config, None
        except Exception as e:
            return None, e
    
    def create_template(self) -> None:
        config_template = {
            'quotation': {
                'input_path': 'sample/inputs',
                'input_sheet': '2025',
                'area1': {
                    'process': True,
                    'suffix': '1'
                },
                'area2': {
                    'process': True,
                    'suffix': '2'
                }
            },
            'report': {
                'output_path': 'sample/outputs',
                'template_path': 'templates',
                'area1': {
                    'template_file': 'TEMPLATE_AREA_1.XLSX',
                    '20feet': {
                        'report_sheet': '운임견적(20피트)',
                        'bestprices_sheet': 'AREA 1 - 20FT'
                    },
                    '40feet': {
                        'report_sheet': '운임견적(40피트)',
                        'bestprices_sheet': 'AREA 1 - 40HC'
                    }
                },
                'area2': {
                    'template_file': 'TEMPLATE_AREA_2.XLSX',
                    '20feet': {
                        'report_sheet': '운임견적(20피트)',
                        'bestprices_sheet': 'AREA 2 - 20FT'
                    },
                    '40feet': {
                        'report_sheet': '운임견적(40피트)',
                        'bestprices_sheet': 'AREA 2 - 40HC'
                    }
                }
            }
        }
        
        try:
            with open(self.config_path, 'w', encoding='utf-8') as file:
                yaml.dump(config_template, file, default_flow_style=False, allow_unicode=True)
        except Exception as e:
            raise Exception(f"Failed to create configuration template: {e}")
    
    def get_config(self) -> Optional[Dict[str, Any]]:
        return self.config
    
    def validate_config(self) -> Tuple[bool, Optional[str]]:
        if not self.config:
            return False, "Configuration not loaded"
        
        required_keys = ['quotation', 'report']
        for key in required_keys:
            if key not in self.config:
                return False, f"Missing required configuration section: {key}"
        
        quotation_config = self.config['quotation']
        required_quotation_keys = ['input_path', 'input_sheet', 'area1', 'area2']
        for key in required_quotation_keys:
            if key not in quotation_config:
                return False, f"Missing required quotation configuration: {key}"
        
        report_config = self.config['report']
        required_report_keys = ['output_path', 'template_path', 'area1', 'area2']
        for key in required_report_keys:
            if key not in report_config:
                return False, f"Missing required report configuration: {key}"
        
        for area in ['area1', 'area2']:
            area_config = report_config[area]
            if 'template_file' not in area_config:
                return False, f"Missing template_file in {area} configuration"
        
        return True, None 