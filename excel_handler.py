import pandas as pd
import os
import logging
from datetime import datetime

class ExcelHandler:
    """Handle Excel file operations including reading, validation, and data extraction"""
    
    def __init__(self, logger=None):
        """Initialize with optional logger"""
        self.logger = logger or logging.getLogger('EZCADAutomation')
        self.current_file = None
        self.current_data = None
    
    def load_excel(self, file_path):
        """Load an Excel file and return the DataFrame"""
        try:
            self.logger.info(f"Loading Excel file: {file_path}")
            
            # Check if file exists
            if not os.path.exists(file_path):
                self.logger.error(f"Excel file not found: {file_path}")
                return None
            
            # Use appropriate engine based on file extension
            engine = 'xlrd' if file_path.lower().endswith('.xls') else None
            
            # Load the file
            df = pd.read_excel(file_path, engine=engine)
            
            self.current_file = file_path
            self.current_data = df
            
            self.logger.info(f"Excel file loaded successfully: {file_path}")
            self.logger.debug(f"Excel dimensions: {df.shape[0]} rows, {df.shape[1]} columns")
            
            return df
        
        except Exception as e:
            self.logger.error(f"Failed to load Excel file {file_path}: {str(e)}")
            return None
    
    def get_preview(self, max_rows=5, max_cols=10):
        """Get a string preview of the current DataFrame"""
        if self.current_data is None:
            return "No Excel file loaded"
        
        try:
            preview_rows = []
            # Get column headers
            headers = []
            for col_idx in range(min(max_cols, self.current_data.shape[1])):
                col_name = str(self.current_data.columns[col_idx])
                headers.append(col_name[:8].ljust(8))
            preview_rows.append(" ".join(headers))
            
            # Get data rows
            for row_idx in range(min(max_rows, self.current_data.shape[0])):
                row_cells = []
                for col_idx in range(min(max_cols, self.current_data.shape[1])):
                    try:
                        val = str(self.current_data.iat[row_idx, col_idx])
                    except:
                        val = ""
                    row_cells.append(val[:8].ljust(8))
                preview_rows.append(" ".join(row_cells))
            
            return "\n".join(preview_rows)
        
        except Exception as e:
            self.logger.error(f"Error generating Excel preview: {str(e)}")
            return f"Error in preview: {str(e)}"
    
    def validate_excel(self, required_columns=None):
        """Validate if the Excel has the required structure"""
        if self.current_data is None:
            self.logger.error("No Excel file loaded for validation")
            return False
        
        # If no specific columns are required, consider valid
        if not required_columns:
            return True
            
        # Check for required columns
        missing_columns = []
        for col in required_columns:
            if col not in self.current_data.columns:
                missing_columns.append(col)
        
        if missing_columns:
            self.logger.warning(f"Missing required columns: {', '.join(missing_columns)}")
            return False
        
        return True
    
    def get_batch_data(self, batch_size=10):
        """Get data in batches for processing"""
        if self.current_data is None:
            self.logger.error("No Excel file loaded for batch processing")
            return []
        
        total_rows = self.current_data.shape[0]
        batches = []
        
        for start_idx in range(0, total_rows, batch_size):
            end_idx = min(start_idx + batch_size, total_rows)
            batch = self.current_data.iloc[start_idx:end_idx]
            batches.append(batch)
        
        self.logger.info(f"Split data into {len(batches)} batches")
        return batches
    
    def save_processed_status(self, processed_rows, status_col='Processed', timestamp_col='Processed_Time'):
        """Save processing status back to the Excel file"""
        if self.current_data is None or self.current_file is None:
            self.logger.error("No Excel file loaded for saving status")
            return False
        
        try:
            # Create backup of original file
            backup_file = f"{os.path.splitext(self.current_file)[0]}_backup_{datetime.now().strftime('%Y%m%d%H%M%S')}.xlsx"
            self.current_data.to_excel(backup_file, index=False)
            self.logger.info(f"Created backup of Excel file: {backup_file}")
            
            # Add status columns if they don't exist
            if status_col not in self.current_data.columns:
                self.current_data[status_col] = ''
            if timestamp_col not in self.current_data.columns:
                self.current_data[timestamp_col] = ''
            
            # Update status for processed rows
            timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            for row_idx in processed_rows:
                if 0 <= row_idx < self.current_data.shape[0]:
                    self.current_data.at[row_idx, status_col] = 'Processed'
                    self.current_data.at[row_idx, timestamp_col] = timestamp
            
            # Save updated file
            self.current_data.to_excel(self.current_file, index=False)
            self.logger.info(f"Updated processing status in Excel file: {self.current_file}")
            
            return True
            
        except Exception as e:
            self.logger.error(f"Failed to save processing status: {str(e)}")
            return False
    
    def get_column_names(self):
        """Get the column names of the current DataFrame"""
        if self.current_data is None:
            return []
        return list(self.current_data.columns)