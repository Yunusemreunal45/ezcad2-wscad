import logging
import time

class Processor:
    """Process files with EZCAD"""
    
    def __init__(self, excel_handler, ezcad_controller, config, logger=None):
        self.excel_handler = excel_handler
        self.ezcad_controller = ezcad_controller
        self.config = config
        self.logger = logger or logging.getLogger('EZCADAutomation')

    def process_file(self, file_path):
        """Process a file based on its type"""
        try:
            if file_path.lower().endswith(('.xls', '.xlsx')):
                return self._process_excel(file_path)
            elif file_path.lower().endswith('.ezd'):
                return self._process_ezd(file_path)
            else:
                raise ValueError(f"Unsupported file type: {file_path}")
        except Exception as e:
            self.logger.error(f"Error processing file {file_path}: {str(e)}")
            raise

    def _process_excel(self, excel_file):
        """Process an Excel file"""
        df = self.excel_handler.load_excel(excel_file)
        if df is None:
            raise Exception("Failed to load Excel file")

        # Process the Excel data here
        result = {
            "rows_processed": len(df),
            "columns": list(df.columns)
        }

        return result

    def _process_ezd(self, ezd_file):
        """Process an EZD file"""
        window_id = self.ezcad_controller.start_ezcad(ezd_file)
        if not window_id:
            raise Exception("Failed to start EZCAD")

        try:
            time.sleep(1)
            self.ezcad_controller.send_command(window_id, "red")
            time.sleep(0.5)
            self.ezcad_controller.send_command(window_id, "mark")
            time.sleep(1)
            # Şablonu tekrar yükle (isteğe bağlı)
            # self.ezcad_controller.send_command(window_id, "load_template")  # Eğer böyle bir komutunuz varsa
            result = {
                "window_id": window_id,
                "status": "completed"
            }
            return result
        finally:
            self.ezcad_controller.close_ezcad(window_id)