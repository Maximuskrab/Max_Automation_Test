"""Core Excel appending logic."""

import os
from typing import List
from pathlib import Path
import pandas as pd
from io import BytesIO


class ExcelAppender:
    """Handles discovering, reading, and appending Excel files."""
    
    def __init__(self):
        """Initialize the ExcelAppender."""
        self.discovered_files = []
        self.combined_dataframe = None
    
    def discover_files(self, folder_path: str) -> tuple[List[str], str]:
        """
        Discover all Excel files in the specified folder.
        
        Args:
            folder_path: Path to the folder containing Excel files
            
        Returns:
            Tuple of (list of file paths, status message)
        """
        try:
            # Validate folder exists
            if not os.path.isdir(folder_path):
                return [], f"Error: Folder '{folder_path}' does not exist."
            
            # Search for Excel files
            excel_extensions = ('.xlsx', '.xls')
            files = []
            
            for file in os.listdir(folder_path):
                if file.lower().endswith(excel_extensions):
                    full_path = os.path.join(folder_path, file)
                    if os.path.isfile(full_path):
                        files.append(full_path)
            
            self.discovered_files = sorted(files)
            
            if not files:
                return [], f"No Excel files found in '{folder_path}'."
            
            status = f"Found {len(files)} Excel file(s)."
            return files, status
            
        except Exception as e:
            return [], f"Error discovering files: {str(e)}"
    
    def append_files(self, file_list: List[str]) -> tuple[pd.DataFrame, str]:
        """
        Read and concatenate all Excel files.
        Simply appends all data without header checks.
        
        Args:
            file_list: List of file paths to append
            
        Returns:
            Tuple of (combined DataFrame, status message)
        """
        try:
            if not file_list:
                return None, "Error: No files to append."
            
            dataframes = []
            total_rows = 0
            
            for file_path in file_list:
                try:
                    # Read all sheets from the Excel file
                    excel_file = pd.ExcelFile(file_path)
                    
                    # Read and append each sheet (no headers)
                    for sheet in excel_file.sheet_names:
                        df = pd.read_excel(file_path, sheet_name=sheet, header=None)
                        if not df.empty:
                            dataframes.append(df)
                            total_rows += len(df)
                        
                except Exception as e:
                    return None, f"Error reading {os.path.basename(file_path)}: {str(e)}"
            
            if not dataframes:
                return None, "No data found in Excel files."
            
            # Concatenate all data
            combined = pd.concat(dataframes, ignore_index=True)
            self.combined_dataframe = combined
            
            status = f"✅ Successfully appended all files! Total rows: {len(combined)}"
            return combined, status
            
        except Exception as e:
            return None, f"Error: {str(e)}"
    
    def get_output_filename(self, file_list: List[str]) -> str:
        """
        Generate output filename by combining input file names.
        
        Args:
            file_list: List of file paths
            
        Returns:
            Combined filename (e.g., "book1_book2.xlsx")
        """
        file_names = [os.path.splitext(os.path.basename(f))[0] for f in file_list]
        combined_name = "_".join(file_names)
        return f"{combined_name}.xlsx"
    
    def save_to_bytes(self, dataframe: pd.DataFrame) -> tuple[bytes, str]:
        """
        Save the combined dataframe to bytes (in-memory).
        
        Args:
            dataframe: The combined DataFrame to save
            
        Returns:
            Tuple of (file bytes, status message)
        """
        try:
            if dataframe is None or dataframe.empty:
                return None, "Error: No data to save."
            
            # Save to BytesIO (in-memory)
            buffer = BytesIO()
            dataframe.to_excel(buffer, index=False, header=False, engine='openpyxl')
            buffer.seek(0)
            
            status = f"✅ File ready for download!"
            return buffer.getvalue(), status
            
        except Exception as e:
            return None, f"Error saving file: {str(e)}"
