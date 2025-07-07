import os
import sys
import logging
import shutil
from datetime import datetime

from config_manager import ConfigManager
from progress_tracker import ProgressTracker
from data_processor import DataProcessor
from report_generator import ReportGenerator


class QuotationApp:
    
    def __init__(self, config_path: str = 'config.yaml'):
        self.config_manager = ConfigManager(config_path)
        self.progress_tracker = ProgressTracker()
        self.logger = self._setup_logging()
        
        config, err = self.config_manager.load_config()
        if err:
            self.config_manager.create_template()
            self.logger.error(f"{datetime.now()}: File configuration '{config_path}' not found, template file has been generated!")
            sys.exit()
        
        self.config = config
        self.report_generator = ReportGenerator(config)
    
    def _setup_logging(self) -> logging.Logger:
        logging.basicConfig(filename='quotation.log', level=logging.INFO)
        return logging.getLogger(__name__)
    
    def _generate_report_filename(self, area: str) -> str:
        now = datetime.now()
        month = now.strftime("%B").upper()
        year = now.strftime("%y")
        area_num = area.replace('area', '')
        return f"QUOTATION_{month}_{year}_AREA_{area_num}.XLSX"
    
    def _copy_template_to_output(self, area: str) -> str:
        template_path = self.config['report']['template_path']
        output_path = self.config['report']['output_path']
        template_file = self.config['report'][area]['template_file']
        
        source_template = os.path.join(template_path, template_file)
        report_filename = self._generate_report_filename(area)
        destination_file = os.path.join(output_path, report_filename)
        
        try:
            shutil.copy2(source_template, destination_file)
            self.logger.info(f"{datetime.now()}: Template copied: {template_file} -> {report_filename}")
            return destination_file
        except Exception as e:
            self.logger.error(f"{datetime.now()}: Failed to copy template {template_file}: {str(e)}")
            raise e
    
    def _validate_environment(self) -> bool:
        input_path = self.config['quotation']['input_path']
        output_path = self.config['report']['output_path']
        template_path = self.config['report']['template_path']
        
        if not os.path.exists(input_path):
            self.logger.error(f"{datetime.now()}: Input file path '{input_path}' not found")
            return False
        
        if not os.path.exists(output_path):
            self.logger.error(f"{datetime.now()}: Output file path '{output_path}' not found")
            return False
        
        if not os.path.exists(template_path):
            self.logger.error(f"{datetime.now()}: Template file path '{template_path}' not found")
            return False
        
        area1 = self.config['quotation']['area1'].get('process', False)
        area2 = self.config['quotation']['area2'].get('process', False)
        
        if area1:
            area1_template = os.path.join(template_path, self.config['report']['area1']['template_file'])
            if not os.path.exists(area1_template):
                self.logger.error(f"{datetime.now()}: Template file '{area1_template}' not found")
                return False
        
        if area2:
            area2_template = os.path.join(template_path, self.config['report']['area2']['template_file'])
            if not os.path.exists(area2_template):
                self.logger.error(f"{datetime.now()}: Template file '{area2_template}' not found")
                return False
        
        return True

    def _determine_file_area(self, filename: str, area1_suffix: str, area2_suffix: str) -> str:
        filename_lower = filename.lower()
        if (filename_lower.endswith(area1_suffix + '.xls') or 
            filename.endswith(area1_suffix + '.xlsx') or 
            filename.endswith(area1_suffix + '.xlsb')):
            return 'area1'
        elif (filename_lower.endswith(area2_suffix + '.xls') or 
              filename.endswith(area2_suffix + '.xlsx') or 
              filename.endswith(area2_suffix + '.xlsb')):
            return 'area2'
        return None

    def _get_area_config(self, area: str) -> dict:
        area_config = self.config['report'][area]
        
        return {
            '20feet_sheet': area_config['20feet']['report_sheet'],
            '40feet_sheet': area_config['40feet']['report_sheet']
        }

    def _process_single_file_with_template(self, input_file: str, input_path: str, area: str, report_file: str) -> dict:
        try:
            area_config = self._get_area_config(area)
            
            input_file_path = os.path.join(input_path, input_file)
            data_processor = DataProcessor(input_file_path, self.config)
            
            forwarder_data = data_processor.Run(
                report_file, 
                area_config['20feet_sheet'], 
                area_config['40feet_sheet']
            )

            return forwarder_data
            
        except Exception as e:
            self.logger.error(f"{datetime.now()}: Process quotation file {input_file} error: {str(e)}")
            return None

    def process_quotations(self) -> bool:
        if not self._validate_environment():
            return False
        
        input_path = self.config['quotation']['input_path']
        area1 = self.config['quotation']['area1'].get('process', False)
        area2 = self.config['quotation']['area2'].get('process', False)
        area1_suffix = self.config['quotation']['area1']['suffix']
        area2_suffix = self.config['quotation']['area2']['suffix']

        input_files = os.listdir(input_path)
        if not input_files:
            self.logger.info(f"{datetime.now()}: Input file path '{input_path}' is empty")
            return True

        processed = 0
        total = len(input_files)
        failed_files = []
        
        area1_partners = {'20ft': [], '40ft': []}
        area2_partners = {'20ft': [], '40ft': []}
        area1_report_file = None
        area2_report_file = None
        
        for input_file in input_files:
            self.progress_tracker.update_progress(processed, total)
            processed += 1

            area = self._determine_file_area(input_file, area1_suffix, area2_suffix)
            
            if area == 'area1' and area1:
                if not area1_report_file:
                    area1_report_file = self._copy_template_to_output('area1')
                
                partner_data = self._process_single_file_with_template(input_file, input_path, 'area1', area1_report_file)
                if partner_data:
                    area1_partners['20ft'].append(partner_data['20ft'])
                    area1_partners['40ft'].append(partner_data['40ft'])
                else:
                    failed_files.append(input_file)
            elif area == 'area2' and area2:
                if not area2_report_file:
                    area2_report_file = self._copy_template_to_output('area2')
                
                partner_data = self._process_single_file_with_template(input_file, input_path, 'area2', area2_report_file)
                if partner_data:
                    area2_partners['20ft'].append(partner_data['20ft'])
                    area2_partners['40ft'].append(partner_data['40ft'])
                else:
                    failed_files.append(input_file)
            else:
                continue    
        
        self.progress_tracker.update_progress(processed, total)
        
        if failed_files:
            self.logger.warning(f"{datetime.now()}: Failed to process {len(failed_files)} files: {failed_files}")
        
        try:
            if area1 and area1_partners['20ft'] and area1_report_file:
                self.report_generator.write_all_partners_data(area1_partners, area1_report_file)
            
            if area2 and area2_partners['20ft'] and area2_report_file:
                self.report_generator.write_all_partners_data(area2_partners, area2_report_file)
                
        except Exception as e:
            self.logger.error(f"{datetime.now()}: Write partner data error: {str(e)}")
            return False
        
        return len(failed_files) == 0

    def generate_best_prices(self) -> bool:
        processed = 0
        
        try:
            for area in ["area1", "area2"]:
                report_filename = self._generate_report_filename(area)
                output_path = self.config['report']['output_path']
                report_file = os.path.join(output_path, report_filename)
                
                if os.path.exists(report_file):
                    area_config = self.config['report'][area]
                    
                    self.progress_tracker.update_progress(processed, 2)
                    processed += 1

                    self.report_generator.generate_and_write_best_prices(report_file, area_config)
            
            self.progress_tracker.update_progress(processed, 2)
            
        except Exception as e:
            self.logger.error(f"{datetime.now()}: Build best price list error: {str(e)}")
            return False
        
        return True

    def run(self) -> None:
        if self.process_quotations():
            print("\r\nQuotation has been processed successfully, continue to select best price list!")
        else:
            print("\r\nQuotation has been processed fail, please check error in log file!")
            sys.exit()

        if self.generate_best_prices():
            print("\r\nBest price list has been processed successfully, please check report!")
        else:
            print("\r\nBest price list has been processed fail, please check error in log file!")
            sys.exit() 