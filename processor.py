import os
import logging
import time
import threading
from platform_utils import PlatformUtils

class Processor:
    """Process Excel and EZD files with EZCAD2"""
    
    def __init__(self, excel_handler, ezcad_controller, config_manager, logger=None):
        """
        Initialize the processor
        
        Args:
            excel_handler: Excel handler instance
            ezcad_controller: EZCAD controller instance
            config_manager: Configuration manager instance
            logger: Logger instance
        """
        self.excel_handler = excel_handler
        self.ezcad_controller = ezcad_controller
        self.config_manager = config_manager
        self.logger = logger or logging.getLogger('EZCADAutomation')
        self.stop_requested = False
    
    def process_excel(self, excel_file):
        """
        Process an Excel file
        
        Args:
            excel_file: Path to Excel file
            
        Returns:
            dict: Results of processing
        """
        try:
            self.logger.info(f"Processing Excel file: {excel_file}")
            
            # Load the Excel file
            df = self.excel_handler.load_excel(excel_file)
            if df is None:
                raise ValueError(f"Failed to load Excel file: {excel_file}")
            
            # Check if batch processing is enabled
            batch_process = self.config_manager.getboolean('Settings', 'batch_process', fallback=False)
            
            if batch_process:
                # Process in batches
                batch_size = self.config_manager.getint('Settings', 'batch_size', fallback=10)
                batches = self.excel_handler.get_batch_data(batch_size)
                
                results = {
                    'total_rows': len(df),
                    'processed_rows': 0,
                    'successful_rows': 0,
                    'failed_rows': 0,
                    'batches': []
                }
                
                # Process each batch
                for i, batch in enumerate(batches):
                    if self.stop_requested:
                        self.logger.info("Processing stopped by user")
                        break
                        
                    self.logger.info(f"Processing batch {i+1}/{len(batches)}")
                    
                    batch_result = self._process_batch(batch, i)
                    results['batches'].append(batch_result)
                    
                    results['processed_rows'] += batch_result['rows_processed']
                    results['successful_rows'] += batch_result['rows_successful']
                    results['failed_rows'] += batch_result['rows_failed']
                
            else:
                # Process entire file
                results = self._process_entire_file(df)
            
            return results
            
        except Exception as e:
            self.logger.error(f"Error processing Excel file {excel_file}: {str(e)}")
            return {'error': str(e)}
    
    def _process_batch(self, batch, batch_index):
        """Process a batch of rows from Excel"""
        result = {
            'batch_index': batch_index,
            'rows_processed': len(batch),
            'rows_successful': 0,
            'rows_failed': 0,
            'errors': []
        }
        
        # Find the EZD file to use
        ezd_file = self.config_manager.get('Paths', 'last_ezd_file')
        if not ezd_file or not os.path.isfile(ezd_file):
            err = "No EZD file configured for processing"
            self.logger.error(err)
            result['errors'].append(err)
            result['rows_failed'] = len(batch)
            return result
        
        # Start EZCAD with the EZD file
        window_id = self.ezcad_controller.start_ezcad(ezd_file)
        if not window_id:
            err = "Failed to start EZCAD"
            self.logger.error(err)
            result['errors'].append(err)
            result['rows_failed'] = len(batch)
            return result
        
        # Process each row in the batch
        try:
            for i, row in batch.iterrows():
                # TODO: Implement actual processing logic for each row
                # This would involve sending data to EZCAD, triggering marking, etc.
                
                # For now we'll just simulate success
                time.sleep(0.5)  # Simulate processing time
                
                if i % 3 == 0:  # Simulate some failures for testing
                    result['rows_failed'] += 1
                    result['errors'].append(f"Simulated error for row {i}")
                else:
                    result['rows_successful'] += 1
                    # Send mark command for successful rows
                    self.ezcad_controller.send_command(window_id, 'mark')
        
        finally:
            # Always close EZCAD when done
            self.ezcad_controller.close_ezcad(window_id)
        
        return result
    
    def _process_entire_file(self, df):
        """Process the entire Excel file at once"""
        result = {
            'total_rows': len(df),
            'processed': True,
            'errors': []
        }
        
        # Find the EZD file to use
        ezd_file = self.config_manager.get('Paths', 'last_ezd_file')
        if not ezd_file or not os.path.isfile(ezd_file):
            err = "No EZD file configured for processing"
            self.logger.error(err)
            result['errors'].append(err)
            result['processed'] = False
            return result
        
        # Start EZCAD with the EZD file
        window_id = self.ezcad_controller.start_ezcad(ezd_file)
        if not window_id:
            err = "Failed to start EZCAD"
            self.logger.error(err)
            result['errors'].append(err)
            result['processed'] = False
            return result
        
        # TODO: Implement actual processing logic for the whole file
        # This depends on how you want to integrate with EZCAD
        
        # For now, just simulate processing and send a mark command
        time.sleep(2)  # Simulate processing time
        self.ezcad_controller.send_command(window_id, 'mark')
        
        # Close EZCAD when done
        self.ezcad_controller.close_ezcad(window_id)
        
        return result
    
    def process_ezd(self, ezd_file):
        """
        Process an EZD file
        
        Args:
            ezd_file: Path to EZD file
            
        Returns:
            dict: Results of processing
        """
        try:
            self.logger.info(f"Processing EZD file: {ezd_file}")
            
            # Just open the EZD file in EZCAD
            window_id = self.ezcad_controller.start_ezcad(ezd_file)
            
            if window_id:
                # Update last used EZD file in config
                self.config_manager.set('Paths', 'last_ezd_file', ezd_file)
                self.config_manager.save_config()
                
                return {
                    'success': True,
                    'window_id': window_id,
                    'file': ezd_file
                }
            else:
                return {
                    'success': False,
                    'error': 'Failed to open EZD file'
                }
                
        except Exception as e:
            self.logger.error(f"Error processing EZD file {ezd_file}: {str(e)}")
            return {'error': str(e)}
    
    def request_stop(self):
        """Request stopping ongoing processing"""
        self.stop_requested = True
        self.logger.info("Stop requested for processor")
    
    def reset_stop_flag(self):
        """Reset the stop flag"""
        self.stop_requested = False
