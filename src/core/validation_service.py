"""
Validation Service for ASTM Test Calculator

This module handles all validation logic including path validation, file format validation,
and data integrity checks. Extracted from main_window.py to separate validation logic 
from UI logic.
"""

import os
import pandas as pd
from pathlib import Path
from typing import List, Tuple, Optional, Dict, Any
import traceback


class ValidationError(Exception):
    """Custom exception for validation errors"""
    pass


class ValidationService:
    """Service class for handling all validation operations"""
    
    def __init__(self, debug_mode: bool = False):
        self.debug_mode = debug_mode
    
    def validate_all_paths(self, test_plan_path: str, slm_data_d_path: str, 
                          slm_data_e_path: str, report_output_path: str) -> bool:
        """
        Validate all required paths for the application
        
        Args:
            test_plan_path: Path to test plan file
            slm_data_d_path: Path to SLM data directory (meter 1)
            slm_data_e_path: Path to SLM data directory (meter 2)
            report_output_path: Path to output directory
            
        Returns:
            True if all paths are valid
            
        Raises:
            ValidationError: If any path validation fails
        """
        try:
            # Check that all paths are provided
            if not all([test_plan_path, slm_data_d_path, slm_data_e_path, report_output_path]):
                raise ValidationError("All paths must be provided")
            
            # Validate test plan file
            self.validate_test_plan_file(test_plan_path)
            
            # Validate SLM data directories
            self.validate_slm_directory(slm_data_d_path, "SLM Data Meter 1")
            self.validate_slm_directory(slm_data_e_path, "SLM Data Meter 2")
            
            # Validate/create output directory
            self.validate_output_directory(report_output_path)
            
            if self.debug_mode:
                print("\nAll path validations passed:")
                print(f"Test plan: {test_plan_path}")
                print(f"SLM Data D: {slm_data_d_path}")
                print(f"SLM Data E: {slm_data_e_path}")
                print(f"Output: {report_output_path}")
            
            return True
            
        except Exception as e:
            if self.debug_mode:
                print(f"Path validation failed: {str(e)}")
            raise ValidationError(f"Path validation failed: {str(e)}")
    
    def validate_test_plan_file(self, test_plan_path: str) -> bool:
        """
        Validate test plan file exists and is readable
        
        Args:
            test_plan_path: Path to test plan file
            
        Returns:
            True if file is valid
            
        Raises:
            ValidationError: If file validation fails
        """
        try:
            if not os.path.exists(test_plan_path):
                raise ValidationError(f"Test plan file not found: {test_plan_path}")
            
            if not os.path.isfile(test_plan_path):
                raise ValidationError(f"Test plan path must be a file: {test_plan_path}")
            
            # Check file extension
            file_ext = os.path.splitext(test_plan_path)[1].lower()
            if file_ext not in ['.csv', '.xlsx', '.xls']:
                raise ValidationError(f"Test plan file must be CSV or Excel format: {test_plan_path}")
            
            # Try to read the file to ensure it's not corrupted
            try:
                if file_ext == '.csv':
                    pd.read_csv(test_plan_path, nrows=1)
                else:
                    pd.read_excel(test_plan_path, nrows=1)
            except Exception as e:
                if "File is not a zip file" in str(e) or "appears to be corrupted" in str(e):
                    raise ValidationError(f"Excel file appears to be corrupted or has wrong extension: {test_plan_path}")
                else:
                    raise ValidationError(f"Cannot read test plan file: {str(e)}")
            
            if self.debug_mode:
                print(f"Test plan file validation passed: {test_plan_path}")
            
            return True
            
        except ValidationError:
            raise
        except Exception as e:
            raise ValidationError(f"Test plan file validation failed: {str(e)}")
    
    def validate_slm_directory(self, directory_path: str, directory_name: str) -> bool:
        """
        Validate SLM data directory exists and contains data files
        
        Args:
            directory_path: Path to SLM data directory
            directory_name: Human-readable name for error messages
            
        Returns:
            True if directory is valid
            
        Raises:
            ValidationError: If directory validation fails
        """
        try:
            if not os.path.exists(directory_path):
                raise ValidationError(f"{directory_name} directory not found: {directory_path}")
            
            if not os.path.isdir(directory_path):
                raise ValidationError(f"{directory_name} path must be a directory: {directory_path}")
            
            # Check if directory contains any Excel files
            excel_files = [f for f in os.listdir(directory_path) 
                          if f.lower().endswith(('.xlsx', '.xls'))]
            
            if not excel_files:
                raise ValidationError(f"{directory_name} directory contains no Excel files: {directory_path}")
            
            if self.debug_mode:
                print(f"{directory_name} directory validation passed: {directory_path} ({len(excel_files)} Excel files found)")
            
            return True
            
        except ValidationError:
            raise
        except Exception as e:
            raise ValidationError(f"{directory_name} directory validation failed: {str(e)}")
    
    def validate_output_directory(self, output_path: str) -> bool:
        """
        Validate output directory exists or can be created
        
        Args:
            output_path: Path to output directory
            
        Returns:
            True if directory is valid or was created
            
        Raises:
            ValidationError: If directory validation fails
        """
        try:
            if not os.path.exists(output_path):
                try:
                    os.makedirs(output_path, exist_ok=True)
                    if self.debug_mode:
                        print(f"Created output directory: {output_path}")
                except PermissionError:
                    raise ValidationError(f"Cannot create output directory: {output_path}. Please check permissions.")
                except Exception as e:
                    raise ValidationError(f"Failed to create output directory: {str(e)}")
            
            if not os.path.isdir(output_path):
                raise ValidationError(f"Output path must be a directory: {output_path}")
            
            # Test write permissions
            try:
                test_file = os.path.join(output_path, ".test_write_permissions")
                with open(test_file, 'w') as f:
                    f.write("test")
                os.remove(test_file)
            except Exception as e:
                raise ValidationError(f"No write permissions for output directory: {output_path}")
            
            if self.debug_mode:
                print(f"Output directory validation passed: {output_path}")
            
            return True
            
        except ValidationError:
            raise
        except Exception as e:
            raise ValidationError(f"Output directory validation failed: {str(e)}")
    
    def validate_frequency_data(self, freq_data: Dict[str, Any], test_type: str) -> bool:
        """
        Validate frequency data has correct structure and length
        
        Args:
            freq_data: Dictionary containing frequency data
            test_type: Type of test (AIIC, ASTC, NIC)
            
        Returns:
            True if data is valid
            
        Raises:
            ValidationError: If data validation fails
        """
        try:
            expected_length = 17  # Standard number of frequency points
            
            # Check required keys based on test type
            if test_type == "AIIC":
                required_keys = ['source', 'avg_pos', 'background', 'rt']
            elif test_type == "ASTC":
                required_keys = ['source', 'receive', 'background', 'rt']
            elif test_type == "NIC":
                required_keys = ['source', 'receive', 'background', 'rt']
            else:
                raise ValidationError(f"Unknown test type: {test_type}")
            
            # Check all required keys are present
            missing_keys = [key for key in required_keys if key not in freq_data]
            if missing_keys:
                raise ValidationError(f"Missing required frequency data keys for {test_type}: {missing_keys}")
            
            # Check data lengths
            for key in required_keys:
                data = freq_data[key]
                if hasattr(data, '__len__') and len(data) != expected_length:
                    raise ValidationError(f"{test_type} {key} data has incorrect length. Expected {expected_length}, got {len(data)}")
            
            if self.debug_mode:
                print(f"Frequency data validation passed for {test_type}")
                for key in required_keys:
                    data = freq_data[key]
                    print(f"  {key}: length {len(data) if hasattr(data, '__len__') else 'N/A'}")
            
            return True
            
        except ValidationError:
            raise
        except Exception as e:
            raise ValidationError(f"Frequency data validation failed: {str(e)}")
    
    def validate_file_format(self, file_path: str, expected_formats: List[str]) -> bool:
        """
        Validate file has expected format
        
        Args:
            file_path: Path to file
            expected_formats: List of expected file extensions (e.g., ['.csv', '.xlsx'])
            
        Returns:
            True if file format is valid
            
        Raises:
            ValidationError: If file format validation fails
        """
        try:
            if not os.path.exists(file_path):
                raise ValidationError(f"File not found: {file_path}")
            
            file_ext = os.path.splitext(file_path)[1].lower()
            
            if file_ext not in expected_formats:
                raise ValidationError(f"File must have one of these extensions: {expected_formats}. Got: {file_ext}")
            
            if self.debug_mode:
                print(f"File format validation passed: {file_path} ({file_ext})")
            
            return True
            
        except ValidationError:
            raise
        except Exception as e:
            raise ValidationError(f"File format validation failed: {str(e)}")
    
    def suggest_error_fixes(self, error_message: str) -> str:
        """
        Provide helpful suggestions based on error message
        
        Args:
            error_message: The error message to analyze
            
        Returns:
            Enhanced error message with suggestions
        """
        suggestions = []
        
        if "File is not a zip file" in error_message or "appears to be corrupted" in error_message:
            suggestions.extend([
                "This file appears to be a CSV file with an incorrect .xlsx extension",
                "Open the file in Excel/LibreOffice and save it in a new format (CSV or XLSX)",
                "Check if the file is password protected",
                "Try using a different Excel file"
            ])
        
        elif "No matching files found" in error_message:
            suggestions.extend([
                "Verify SLM data file naming format",
                "Check paths to SLM data directories", 
                "Ensure data files are in the expected format (.xlsx)"
            ])
        
        elif "not found" in error_message.lower():
            suggestions.extend([
                "Check that all paths are correct and files/directories exist",
                "Verify file/directory names are spelled correctly",
                "Ensure you have read permissions for the specified paths"
            ])
        
        elif "permission" in error_message.lower():
            suggestions.extend([
                "Check file/directory permissions",
                "Try running with administrator privileges",
                "Ensure the files are not open in other applications"
            ])
        
        if suggestions:
            return error_message + "\n\nSuggestions:\n" + "\n".join(f"- {s}" for s in suggestions)
        else:
            return error_message