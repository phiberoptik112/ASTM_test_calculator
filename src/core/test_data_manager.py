from typing import Dict, Optional, List, Union
import pandas as pd
from pathlib import Path
import os
from os import listdir
from os.path import isfile, join

from src.core.data_processor import (
    TestType,
    TestData,
    RoomProperties,
    AIICTestData,
    ASTCTestData,
    NICTestData,
    DTCtestData,
    SLMData
)
from src.core.test_processor import TestProcessor

class TestDataManager:
    def __init__(self, debug_mode: bool = True):
        self.debug_mode = debug_mode
        self.test_processor = TestProcessor(debug_mode=debug_mode)
        self.test_data_collection: Dict[str, Dict[TestType, Dict]] = {}
        self.test_plan: Optional[pd.DataFrame] = None
        
        # Initialize paths
        self.test_plan_path = None
        self.slm_data_d_path = None
        self.slm_data_e_path = None
        self.report_output_path = None
        
        # Initialize data file lists
        self.D_datafiles = []
        self.E_datafiles = []
        
        # Initialize slm_data_paths dictionary
        self.slm_data_paths = {
            'meter_1': '',
            'meter_2': ''
        }

    def set_data_paths(self, test_plan_path: str, meter_d_path: str, meter_e_path: str, output_path: str) -> None:
        """Set all necessary file paths and load data files"""
        try:
            print("\n=== Setting Data Paths ===")
            print(f"Test plan: {test_plan_path}")
            print(f"Meter D: {meter_d_path}")
            print(f"Meter E: {meter_e_path}")
            print(f"Output: {output_path}")
            print(f"Current working directory: {os.getcwd()}")
            print(f"Test plan directory: {os.path.dirname(test_plan_path)}")
            print(f"Meter D directory: {os.path.dirname(meter_d_path)}")
            print(f"Meter E directory: {os.path.dirname(meter_e_path)}")
            
            # Verify paths exist
            for path, desc in [
                (test_plan_path, "test plan"),
                (meter_d_path, "meter D data"),
                (meter_e_path, "meter E data"),
                (output_path, "output")
            ]:
                if not os.path.exists(path):
                    raise ValueError(f"Invalid {desc} path: {path}")
                print(f"\nChecking {desc} directory contents:")
                try:
                    print(os.listdir(os.path.dirname(path)))
                except Exception as e:
                    print(f"Error listing directory: {str(e)}")
            
            # Verify test plan file format
            test_plan_ext = os.path.splitext(test_plan_path)[1].lower()
            if test_plan_ext not in ['.csv', '.xlsx', '.xls']:
                raise ValueError(f"Unsupported test plan file format: {test_plan_ext}")
            
            # Set paths
            self.test_plan_path = test_plan_path
            self.slm_data_d_path = meter_d_path
            self.slm_data_e_path = meter_e_path
            self.report_output_path = output_path
            
            # Update slm_data_paths dictionary
            self.slm_data_paths = {
                'meter_1': meter_d_path,
                'meter_2': meter_e_path
            }
            
            # Load data files lists
            try:
                self.D_datafiles = [
                    f for f in listdir(meter_d_path) 
                    if isfile(join(meter_d_path, f)) and f.endswith('.xlsx')
                ]
                self.E_datafiles = [
                    f for f in listdir(meter_e_path) 
                    if isfile(join(meter_e_path, f)) and f.endswith('.xlsx')
                ]
                
                print(f'\nFound {len(self.D_datafiles)} D meter files:')
                for f in self.D_datafiles:
                    print(f"  - {f}")
                print(f'\nFound {len(self.E_datafiles)} E meter files:')
                for f in self.E_datafiles:
                        print(f"  - {f}")
                    
            except Exception as e:
                raise ValueError(f"Error accessing data files: {str(e)}")
                
        except Exception as e:
            raise ValueError(f"Error setting data paths: {str(e)}")

    def load_test_plan(self) -> None:
        """Load test plan from Excel or CSV file"""
        print(f"Loading test plan from {self.test_plan_path}")
        try:
            # First, check if the file exists
            if not os.path.exists(self.test_plan_path):
                raise FileNotFoundError(f"Test plan file not found: {self.test_plan_path}")
            
            # Determine file type from extension
            file_ext = os.path.splitext(self.test_plan_path)[1].lower()
            
            print(f"\nAttempting to load test plan:")
            print(f"File path: {self.test_plan_path}")
            print(f"File extension: {file_ext}")
            
            # Try loading the file based on extension
            if file_ext == '.csv':
                print("Loading as CSV file")
                # Use pandas to read CSV with proper handling of quoted values
                self.test_plan = pd.read_csv(
                    self.test_plan_path,
                    encoding='utf-8',
                    quoting=1,  # QUOTE_ALL
                    skipinitialspace=True,  # Skip leading whitespace
                    on_bad_lines='warn'  # Warn about problematic lines
                )
                
                # Clean up column names
                self.test_plan.columns = self.test_plan.columns.str.strip()
                

                print("CSV load result:")
                print(f"Shape: {self.test_plan.shape}")
                print("Columns:", self.test_plan.columns.tolist())
                print("\nTestplan loaded:")
                print(self.test_plan)
            elif file_ext in ['.xlsx', '.xls']:
                # Try different Excel engines in order
                engines = ['openpyxl', 'xlrd']
                last_error = None
                
                for engine in engines:
                    try:
                        print(f"Attempting to read Excel file with engine: {engine}")
                        if engine == 'xlrd' and file_ext == '.xlsx':
                            print("Warning: Using xlrd with .xlsx file - limited functionality")
                        
                        # First, try to get sheet names
                        if engine == 'openpyxl':
                            xl = pd.ExcelFile(self.test_plan_path, engine=engine)
                            sheet_names = xl.sheet_names

                            print(f"Available sheets: {sheet_names}")
                            
                            # Try to read first sheet or sheet named 'Test Plan' if it exists
                            sheet_to_use = 'Test Plan' if 'Test Plan' in sheet_names else sheet_names[0]
                            print(f"Reading sheet: {sheet_to_use}")
                            
                            self.test_plan = pd.read_excel(
                                self.test_plan_path,
                                sheet_name=sheet_to_use,
                                engine=engine
                            )
                        else:
                            self.test_plan = pd.read_excel(self.test_plan_path, engine=engine)
                        
                        
                        print(f"Excel load result with {engine}:")
                        print(f"Shape: {self.test_plan.shape}")
                        print("Columns:", self.test_plan.columns.tolist())
                        if not self.test_plan.empty:
                            print("\nFirst row:")
                            print(self.test_plan.iloc[0])
                        
                        if not self.test_plan.empty:
                            print(f"Successfully loaded with {engine} engine")
                            break
                        else:
                            print(f"Warning: {engine} engine loaded empty DataFrame")
                            
                    except Exception as e:
                        last_error = e
                        print(f"Failed with {engine} engine: {str(e)}")
                        
                        # If the error suggests the file might be CSV, try that as a fallback
                        if "File is not a zip file" in str(e) or "Unsupported format" in str(e):
                            try:
                                print("Excel engines failed. Trying to load as CSV...")
                                self.test_plan = pd.read_csv(
                                    self.test_plan_path,
                                    encoding='utf-8',
                                    quoting=1,  # QUOTE_ALL
                                    skipinitialspace=True,  # Skip leading whitespace
                                    on_bad_lines='warn'  # Warn about problematic lines
                                )
                                if not self.test_plan.empty:
                                    print("Successfully loaded as CSV")
                                    break
                                else:
                                    print("Warning: CSV fallback loaded empty DataFrame")
                            except Exception as csv_e:
                                print(f"CSV fallback failed: {str(csv_e)}")
                                continue
                
                if self.test_plan is None or self.test_plan.empty:
                    if "File is not a zip file" in str(last_error):
                        raise ValueError(
                            f"The Excel file appears to be corrupted or is actually a CSV file with an .xlsx extension. "
                            f"Please check if the file can be opened in Excel/LibreOffice and save it with the correct extension. "
                            f"Try renaming the file to .csv if you believe it's actually a CSV file. "
                            f"Original error: {str(last_error)}"
                        )
                    else:
                        raise ValueError(f"Failed to read Excel file with all available engines. Last error: {str(last_error)}")
            else:
                raise ValueError(f"Unsupported file format: {file_ext}. Please use .xlsx, .xls, or .csv files.")
            
            # Validate loaded data
            if self.test_plan is None:
                raise ValueError("Failed to load test plan - data is None")
            
            if self.test_plan.empty:
                raise ValueError("Loaded test plan is empty")
            
            # Ensure required columns are present
            required_columns = [
                'Test_Label', 'AIIC', 'ASTC', 'NIC', 'DTC'
            ]
            missing_columns = [col for col in required_columns if col not in self.test_plan.columns]
            if missing_columns:
                print("\nMissing required columns:")
                print(f"Required: {required_columns}")
                print(f"Found: {self.test_plan.columns.tolist()}")
                raise ValueError(f"Test plan missing required columns: {missing_columns}")
            
            # Validate test types are properly formatted
            test_type_columns = ['AIIC', 'ASTC', 'NIC', 'DTC']
            for col in test_type_columns:
                if col in self.test_plan.columns:
                    # Show original values before conversion
                    print(f"\nValues in {col} before conversion:")
                    print(self.test_plan[col].value_counts())
                    
                    # Convert any non-numeric values to 0
                    self.test_plan[col] = pd.to_numeric(self.test_plan[col], errors='coerce').fillna(0).astype(int)
                    
                    # Ensure values are only 0 or 1
                    invalid_values = self.test_plan[~self.test_plan[col].isin([0, 1])]
                    if not invalid_values.empty:
                        print(f"Warning: Found invalid values in {col} column. Converting to 0.")
                        self.test_plan.loc[invalid_values.index, col] = 0
                    
                    # Show converted values
                    print(f"Values in {col} after conversion:")
                    print(self.test_plan[col].value_counts())
            
            print(f"\nSuccessfully loaded test plan:")
            print(f"Number of rows: {len(self.test_plan)}")
            print(f"Columns: {self.test_plan.columns.tolist()}")
            print("\nTest types summary:")
            for col in test_type_columns:
                    if col in self.test_plan.columns:
                        num_enabled = (self.test_plan[col] == 1).sum()
                        print(f"{col}: {num_enabled} tests enabled")
            print("\nFirst few rows of test plan:")
            print(self.test_plan.head())
                
        except Exception as e:
            print("\nError details:")
            print(f"Error type: {type(e)}")
            print(f"Error message: {str(e)}")
            import traceback
            traceback.print_exc()
            raise ValueError(f"Error loading test plan: {str(e)}")

    def process_test_data(self) -> None:
        """Process all tests in the test plan"""
        print("now in process_test_data")
        print(f"test_plan: {self.test_plan}")
        if self.test_plan is None:
            raise ValueError("No test plan loaded")


        print("\n=== Processing Test Data ===")
        print(f"Number of tests to process: {len(self.test_plan)}")
        print("Test plan columns:", self.test_plan.columns.tolist())
        print("\nFirst few rows of test plan:")
        print(self.test_plan.head())
        print("resetting test data collection")
        # Reset test data collection before processing
        self.test_data_collection = {}
        print(f"test_data_collection: {self.test_data_collection}")
        processed_tests = 0
        failed_tests = []
        print("now iterating over the test plan")
        for index, curr_test in self.test_plan.iterrows():
            try:
                # Validate test label
                test_label = curr_test.get('Test_Label')
                print(f"curr_test: {curr_test}")
                print(f"test_label: {test_label}")
                if pd.isna(test_label) or not test_label:
                    raise ValueError(f"Row {index + 1}: Missing or invalid Test Label")

                print(f"\nProcessing test {test_label} (row {index + 1})")
                print("Test types enabled:")
                for test_type in ['AIIC', 'ASTC', 'NIC', 'DTC']:
                        value = curr_test.get(test_type, 0)
                        print(f"{test_type}: {value} ({type(value)})")

                # Create room properties
                print("creating room properties")
                try:
                    room_props = self._create_room_properties(curr_test)

                    print("Successfully created room properties")
                except Exception as e:
                    raise ValueError(f"Error creating room properties: {str(e)}")

                # Process each test type
                print("now processing each test type")
                print(f"curr_test: {curr_test}")
                curr_test_data = {}
                test_types_processed = []

                for test_type, column in {
                    TestType.AIIC: 'AIIC',
                    TestType.ASTC: 'ASTC',
                    TestType.NIC: 'NIC',
                    TestType.DTC: 'DTC'
                }.items():
                    # Ensure the column exists and has a valid value
                    if column not in curr_test:
                    
                        print(f"Warning: {column} column not found in test plan")
                        continue

                    try:
                        # Convert to numeric and handle NaN values
                        value = curr_test[column]
                        if isinstance(value, (int, float)):
                            test_enabled = value == 1
                        else:
                            test_enabled = pd.to_numeric(value, errors='coerce').fillna(0) == 1
                        print(f"Test type {test_type.value} enabled: {test_enabled}")
                    except (ValueError, TypeError) as e:
                        print(f"Warning: Invalid value in {column} column: {curr_test[column]}")
                        print(f"Warning: Invalid value in {column} column: {curr_test[column]}")
                        print(f"Error: {str(e)}")
                        test_enabled = False

                    if test_enabled:
                        print(f"test_enabled: {test_enabled}")
                        try:

                            print(f"\nProcessing {test_type.value} test")
                            print("Required SRS files:")
                            for col in ['Source Room SRS File', 'Receive Room SRS File', 
                                        'Background Room SRS File', 'RT Room SRS File']:
                                print(f"{col}: {curr_test.get(col)}")

                            test_data = self.load_test_data(
                                curr_test=curr_test,
                                test_type=test_type,
                                room_props=room_props
                            )

                            if test_data:
                                curr_test_data[test_type] = {
                                    'room_properties': room_props,
                                    'test_data': test_data
                                }
                                test_types_processed.append(test_type.value)

                                print(f"Successfully processed {test_type.value} test")
                                print(f"Test data type: {type(test_data)}")
                        except Exception as e:
                            print(f"Error processing {test_type.value} test: {str(e)}")
                            import traceback
                            traceback.print_exc()
                            failed_tests.append((test_label, test_type.value, str(e)))

                if curr_test_data:
                    self.test_data_collection[test_label] = curr_test_data
                    processed_tests += 1

                    print(f"\nSuccessfully added test data for {test_label}")
                    print(f"Test types processed: {test_types_processed}")
                else:
                    print(f"\nWarning: No test data processed for {test_label}")

            except Exception as e:
                error_msg = f"Error processing test row {index + 1}"
                if 'Test_Label' in curr_test:
                    error_msg += f" ({curr_test['Test_Label']})"
                error_msg += f": {str(e)}"
                
                print(f"\nERROR: {error_msg}")
                print("Test row data:")
                print(curr_test)
                import traceback
                traceback.print_exc()
                
                failed_tests.append((curr_test.get('Test_Label', f'Row {index + 1}'), 'ALL', str(e)))

        # Final status report

        print("\n=== Test Processing Summary ===")
        print(f"Total tests in plan: {len(self.test_plan)}")
        print(f"Successfully processed: {processed_tests}")
        print(f"Tests with failures: {len(failed_tests)}")
            
        if failed_tests:
            print("\nFailed tests:")
            for test_label, test_type, error in failed_tests:
                print(f"- {test_label} ({test_type}): {error}")
        
        print("\nTest data collection:")
        print(f"Number of tests: {len(self.test_data_collection)}")
        for test_label, test_data in self.test_data_collection.items():
            print(f"\n{test_label}:")
            print(f"Test types: {list(test_data.keys())}")

        if not self.test_data_collection:
            raise ValueError("No test data was successfully processed")

    def get_test_data(self, test_label: str, test_type: TestType) -> Optional[Dict]:
        """Retrieve test data for specific test and type"""
        try:
            return self.test_data_collection[test_label][test_type]
        except KeyError:
            return None

    def get_all_test_labels(self) -> List[str]:
        """Get list of all test labels"""
        return list(self.test_data_collection.keys())

    def get_test_types_for_label(self, test_label: str) -> List[TestType]:
        """Get list of test types available for a test label"""
        try:
            return list(self.test_data_collection[test_label].keys())
        except KeyError:
            return []

    def _create_room_properties(self, curr_test: pd.Series) -> RoomProperties:
        """Create room properties from test data"""
        print("now in _create_room_properties")
        try:

            print("\nCreating room properties")
            print("Available columns:", curr_test.index.tolist())
            print("Test data:")
            print(curr_test)

            # Extract room properties with fallback values
            site_name = curr_test.get('Site_Name', '')
            client_name = curr_test.get('Client_Name', '')
            source_room = curr_test.get('Source Room', '')
            receive_room = curr_test.get('Receiving Room', '')
            test_date = curr_test.get('Test Date', '')
            report_date = curr_test.get('Report Date', '')
            project_name = curr_test.get('Project Name', '')
            test_label = curr_test.get('Test_Label', '')
            source_vol = float(curr_test.get('source room vol', 0))
            receive_vol = float(curr_test.get('receive room vol', 0))
            partition_area = float(curr_test.get('partition area', 0))
            partition_dim = curr_test.get('partition dim', '')
            source_room_finish = curr_test.get('source room finish', '')
            receive_room_finish = curr_test.get('receive room finish', '')
            srs_floor = curr_test.get('srs floor descrip.', '')
            srs_walls = curr_test.get('srs walls descrip.', '')
            srs_ceiling = curr_test.get('srs ceiling descrip.', '')
            rec_floor = curr_test.get('rec floor descrip.', '')
            rec_walls = curr_test.get('rec walls descrip.', '')
            rec_ceiling = curr_test.get('rec ceiling descrip.', '')
            tested_assembly = curr_test.get('tested assembly', '')
            expected_performance = curr_test.get('expected performance', '')
            annex_2_used = bool(curr_test.get('Annex 2 used?', False))
            test_assembly_type = curr_test.get('Test assembly Type', '')

            print("\nExtracted room properties:")
            print(f"Site Name: {site_name}")
            print(f"Client Name: {client_name}")
            print(f"Source Room: {source_room}")
            print(f"Receive Room: {receive_room}")
            print(f"Test Date: {test_date}")
            print(f"Report Date: {report_date}")
            print(f"Project Name: {project_name}")
            print(f"Test Label: {test_label}")
            print(f"Source Volume: {source_vol}")
            print(f"Receive Volume: {receive_vol}")
            print(f"Partition Area: {partition_area}")
            print(f"Partition Dim: {partition_dim}")
            print(f"Source Room Finish: {source_room_finish}")
            print(f"Receive Room Finish: {receive_room_finish}")
            print(f"Source Floor: {srs_floor}")
            print(f"Source Walls: {srs_walls}")
            print(f"Source Ceiling: {srs_ceiling}")
            print(f"Receive Floor: {rec_floor}")
            print(f"Receive Walls: {rec_walls}")
            print(f"Receive Ceiling: {rec_ceiling}")
            print(f"Tested Assembly: {tested_assembly}")
            print(f"Expected Performance: {expected_performance}")
            print(f"Annex 2 Used: {annex_2_used}")
            print(f"Test Assembly Type: {test_assembly_type}")

            # Create room properties object with correct parameter names
            room_props = RoomProperties(
                site_name=site_name,
                client_name=client_name,
                source_room=source_room,
                receive_room=receive_room,
                test_date=test_date,
                report_date=report_date,
                project_name=project_name,
                test_label=test_label,
                source_vol=source_vol,
                receive_vol=receive_vol,
                partition_area=partition_area,
                partition_dim=partition_dim,
                source_room_finish=source_room_finish,
                source_room_name=source_room,
                receive_room_finish=receive_room_finish,
                receive_room_name=receive_room,
                srs_floor=srs_floor,
                srs_walls=srs_walls,
                srs_ceiling=srs_ceiling,
                rec_floor=rec_floor,
                rec_walls=rec_walls,
                rec_ceiling=rec_ceiling,
                tested_assembly=tested_assembly,
                expected_performance=expected_performance,
                annex_2_used=annex_2_used,
                test_assembly_type=test_assembly_type
            )

            print("\nSuccessfully created RoomProperties object")
            print(f"Site Name: {room_props.site_name}")
            print(f"Client Name: {room_props.client_name}")
            print(f"Source Room: {room_props.source_room}")
            print(f"Receive Room: {room_props.receive_room}")
            print(f"Test Date: {room_props.test_date}")
            print(f"Report Date: {room_props.report_date}")
            print(f"Project Name: {room_props.project_name}")
            print(f"Test Label: {room_props.test_label}")
            print(f"Source Volume: {room_props.source_vol}")
            print(f"Receive Volume: {room_props.receive_vol}")
            print(f"Partition Area: {room_props.partition_area}")
            print(f"Partition Dim: {room_props.partition_dim}")
            print(f"Source Room Finish: {room_props.source_room_finish}")
            print(f"Source Room Name: {room_props.source_room_name}")
            print(f"Receive Room Finish: {room_props.receive_room_finish}")
            print(f"Receive Room Name: {room_props.receive_room_name}")
            print(f"Source Floor: {room_props.srs_floor}")
            print(f"Source Walls: {room_props.srs_walls}")
            print(f"Source Ceiling: {room_props.srs_ceiling}")
            print(f"Receive Floor: {room_props.rec_floor}")
            print(f"Receive Walls: {room_props.rec_walls}")
            print(f"Receive Ceiling: {room_props.rec_ceiling}")
            print(f"Tested Assembly: {room_props.tested_assembly}")
            print(f"Expected Performance: {room_props.expected_performance}")
            print(f"Annex 2 Used: {room_props.annex_2_used}")
            print(f"Test Assembly Type: {room_props.test_assembly_type}")

            print("-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-")
            print("returning room_props")
            print(" ....is this the right thing to do here?")
            return room_props

        except Exception as e:
        
            print(f"\nError creating room properties: {str(e)}")
            print("Test data:")
            print(curr_test)
            import traceback
            traceback.print_exc()
        raise ValueError(f"Error creating room properties: {str(e)}")

    def get_sorted_tests(self) -> List[Dict]:
        """
        Get all tests sorted by test label with their associated data
        
        Returns:
            List of dictionaries containing test information:
            [
                {
                    'label': str,
                    'types': List[TestType],
                    'room_properties': RoomProperties,
                    'results': Dict[TestType, float]
                },
                ...
            ]
        """
        sorted_tests = []
        
        for test_label in sorted(self.test_data_collection.keys()):
            test_data = self.test_data_collection[test_label]
            test_types = list(test_data.keys())
            
            # Get room properties from first test type (should be same for all types)
            room_props = test_data[test_types[0]]['room_properties']
            
            # Collect results for each test type
            results = {}
            for test_type in test_types:
                test_instance = test_data[test_type]['test_data']
                if test_type == TestType.AIIC:
                    results[test_type] = test_instance.aiic_value if hasattr(test_instance, 'aiic_value') else None
                elif test_type == TestType.ASTC:
                    results[test_type] = test_instance.astc_value if hasattr(test_instance, 'astc_value') else None
                elif test_type == TestType.NIC:
                    results[test_type] = test_instance.nic_value if hasattr(test_instance, 'nic_value') else None
                elif test_type == TestType.DTC:
                    results[test_type] = test_instance.dtc_value if hasattr(test_instance, 'dtc_value') else None
            
            sorted_tests.append({
                'label': test_label,
                'types': test_types,
                'room_properties': room_props,
                'results': results
            })
            
            print(f"Processed test {test_label} with types {test_types}")
            print(f"Results: {results}")
        
        return sorted_tests

    def load_test_data(self, curr_test: pd.Series, test_type: TestType, room_props: RoomProperties) -> Union[AIICTestData, ASTCTestData, NICTestData, DTCtestData]:
        """Load and format raw test data based on test type"""
        # First verify this test type is enabled for the current test
        test_type_columns = {
            TestType.AIIC: 'AIIC',
            TestType.ASTC: 'ASTC',
            TestType.NIC: 'NIC',
            TestType.DTC: 'DTC'
        }
        
        test_column = test_type_columns.get(test_type)
        value = curr_test[test_column]
        if test_column not in curr_test or (isinstance(value, (int, float)) and value != 1) or (not isinstance(value, (int, float)) and pd.to_numeric(value, errors='coerce').fillna(0) != 1):
            print(f'Test type {test_type.value} is not enabled for test {curr_test["Test_Label"]}')
            raise ValueError(f"Test type {test_type.value} is not enabled for this test")

        print(f'\nLoading base data for test: {curr_test["Test_Label"]} ({test_type.value})')
        print("Available columns:", curr_test.index.tolist())

        try:
            # Clean up column names in curr_test
            curr_test.index = curr_test.index.str.strip()
            
            # Map new column names to old expected names
            column_mappings = {
                'Source': 'Source Room SRS File',
                'Receive': 'Receive Room SRS File',
                'BNL': 'Background Room SRS File',
                'RT': 'RT Room SRS File',
                'Position1': 'Position 1 SRS File',
                'Position2': 'Position 2 SRS File',
                'Position3': 'Position 3 SRS File',
                'Position4': 'Position 4 SRS File',
                'SourceTap': 'Tapper SRS File',
                'Carpet': 'Carpet SRS File'
            }
            
            # Helper function to get value using new or old column name
            def get_column_value(old_name, new_name):
                if new_name in curr_test and pd.notna(curr_test[new_name]) and curr_test[new_name]:
                    print(f"Using new column name: {new_name} = {curr_test[new_name]}")
                    return curr_test[new_name]
                elif old_name in curr_test and pd.notna(curr_test[old_name]) and curr_test[old_name]:
                    print(f"Using old column name: {old_name} = {curr_test[old_name]}")
                    return curr_test[old_name]
                print(f"Warning: No value found for {old_name} or {new_name}")
                return None
            
            # Load base data
            print('\nLoading base data')
            base_data = {
                'srs_data': self._raw_slm_datapull(
                    get_column_value('Source', 'Source Room SRS File'), '-831_Data.'),
                'recive_data': self._raw_slm_datapull(
                    get_column_value('Receive', 'Receive Room SRS File'), '-831_Data.'),
                'bkgrnd_data': self._raw_slm_datapull(
                    get_column_value('BNL', 'Background Room SRS File'), '-831_Data.'),
                'rt': self._raw_slm_datapull(
                    get_column_value('RT', 'RT Room SRS File'), '-RT_Data.')
            }

            # Verify base data
            self._verify_dataframes(base_data)

            # Handle specific test types
            if test_type == TestType.AIIC:
                return self._create_aiic_test(curr_test, room_props, base_data)
            elif test_type == TestType.DTC:
                return self._create_dtc_test(curr_test, room_props, base_data)
            elif test_type == TestType.ASTC:
                return ASTCTestData(room_properties=room_props, test_data=base_data)
            elif test_type == TestType.NIC:
                return NICTestData(room_properties=room_props, test_data=base_data)
            else:
                raise ValueError(f"Unsupported test type: {test_type}")

        except Exception as e:
            print(f"Error loading data: {str(e)}")
            print(f"Current test data: {curr_test}")
            import traceback
            traceback.print_exc()
            raise

    def _create_aiic_test(self, curr_test: pd.Series, room_props: RoomProperties, base_data: Dict) -> AIICTestData:
        """Create AIIC test data instance with additional validation"""
        try:
            print("\nCreating AIIC test data")
            print("Required columns:", ['Position1', 'Position2', 'Position3', 'Position4', 'SourceTap', 'Carpet'])
            print("Available columns:", curr_test.index.tolist())
            
            # Check for both old and new column names
            # Define helper function if not already defined in this scope
            def get_column_value(old_name, new_name):
                if new_name in curr_test and pd.notna(curr_test[new_name]) and curr_test[new_name]:
                    return curr_test[new_name]
                elif old_name in curr_test and pd.notna(curr_test[old_name]) and curr_test[old_name]:
                    return curr_test[old_name]
                return None
            
            # Verify required columns exist, checking both old and new names
            old_cols = ['Position1', 'Position2', 'Position3', 'Position4', 'SourceTap', 'Carpet']
            new_cols = ['Position 1 SRS File', 'Position 2 SRS File', 'Position 3 SRS File', 
                       'Position 4 SRS File', 'Tapper SRS File', 'Carpet SRS File']
            
            missing_data = []
            for i, (old, new) in enumerate(zip(old_cols, new_cols)):
                if get_column_value(old, new) is None:
                    missing_data.append(f"{old}/{new}")
            
            if missing_data:
                raise ValueError(f"Missing required AIIC data: {missing_data}")

            aiic_data = base_data.copy()
            additional_data = {
                'AIIC_pos1': self._raw_slm_datapull(get_column_value('Position1', 'Position 1 SRS File'), '-831_Data.'),
                'AIIC_pos2': self._raw_slm_datapull(get_column_value('Position2', 'Position 2 SRS File'), '-831_Data.'),
                'AIIC_pos3': self._raw_slm_datapull(get_column_value('Position3', 'Position 3 SRS File'), '-831_Data.'),
                'AIIC_pos4': self._raw_slm_datapull(get_column_value('Position4', 'Position 4 SRS File'), '-831_Data.'),
                'AIIC_source': self._raw_slm_datapull(get_column_value('SourceTap', 'Tapper SRS File'), '-831_Data.'),
                'AIIC_carpet': self._raw_slm_datapull(get_column_value('Carpet', 'Carpet SRS File'), '-831_Data.')
            }

            # Verify all data was loaded successfully
            for key, data in additional_data.items():
                if data is None or (hasattr(data, 'empty') and data.empty):
                    raise ValueError(f"Failed to load {key} data")

            aiic_data.update(additional_data)
            
            print("AIIC data loaded successfully")
            print(f"Number of data points: {len(additional_data)}")
            
            return AIICTestData(room_properties=room_props, test_data=aiic_data)
            
        except Exception as e:
            print(f"Error creating AIIC test: {str(e)}")
            import traceback
            traceback.print_exc()
            raise

    def _create_dtc_test(self, curr_test: pd.Series, room_props: RoomProperties, base_data: Dict) -> DTCtestData:
        """Create DTC test data instance"""
        # Define helper function if not already defined in this scope
        def get_column_value(old_name, new_name):
            if new_name in curr_test and pd.notna(curr_test[new_name]) and curr_test[new_name]:
                return curr_test[new_name]
            elif old_name in curr_test and pd.notna(curr_test[old_name]) and curr_test[old_name]:
                return curr_test[old_name]
            return None
            
        dtc_data = base_data.copy()
        additional_data = {
            'srs_door_open': self._raw_slm_datapull(
                get_column_value('Source_Door_Open', 'Source Door Open SRS File'), '-831_Data.'),
            'srs_door_closed': self._raw_slm_datapull(
                get_column_value('Source_Door_Closed', 'Source Door Closed SRS File'), '-831_Data.'),
            'recive_door_open': self._raw_slm_datapull(
                get_column_value('Receive_Door_Open', 'Receive Door Open SRS File'), '-831_Data.'),
            'recive_door_closed': self._raw_slm_datapull(
                get_column_value('Receive_Door_Closed', 'Receive Door Closed SRS File'), '-831_Data.')
        }
        self._verify_dataframes(additional_data)
        dtc_data.update(additional_data)
        return DTCtestData(room_properties=room_props, test_data=dtc_data)

    @staticmethod
    def format_slm_data(df: pd.DataFrame) -> pd.DataFrame:
        """
        Format raw SLM data from OBA sheet into properly structured DataFrame
        
        Args:
            df: Raw DataFrame from OBA sheet
        
        Returns:
            Formatted DataFrame with proper column labels and frequency data
        """
        try:
            # Extract frequency rows (first few rows contain 1/3 octave data)
            freq_data = df.iloc[0:12]  # Adjust range if needed
            
            # Get frequency values from first row
            frequencies = freq_data.iloc[0].dropna().tolist()[1:]  # Skip first column
            
            # Create dictionary to store different measurements
            measurements = {
                'Overall 1/3 Spectra': freq_data.iloc[2].dropna().tolist()[1:],
                'Max 1/3 Spectra': freq_data.iloc[3].dropna().tolist()[1:],
                'Min 1/3 Spectra': freq_data.iloc[4].dropna().tolist()[1:]
            }
            
            # Create formatted DataFrame
            formatted_df = pd.DataFrame(measurements, index=frequencies)
            
            # Set frequency as column name instead of index
            formatted_df.reset_index(inplace=True)
            formatted_df.rename(columns={'index': 'Frequency (Hz)'}, inplace=True)
            
            # Convert frequency values to numeric
            formatted_df['Frequency (Hz)'] = pd.to_numeric(formatted_df['Frequency (Hz)'])
            
            # Convert measurement values to numeric
            for col in formatted_df.columns[1:]:
                formatted_df[col] = pd.to_numeric(formatted_df[col], errors='coerce')
            
            return formatted_df
        
        except Exception as e:
            raise ValueError(f"Error formatting SLM data: {str(e)}")

    def _raw_slm_datapull(self, find_datafile: str, datatype: str) -> SLMData:
        """Pull data from SLM files"""
        try:
            # Clean up any spaces in datafile name
            if isinstance(find_datafile, str):
                find_datafile = find_datafile.strip()
            
            print(f"\nLooking for data file: {find_datafile}")
            print(f"Data type: {datatype}")
            print(f"SLM data D path: {self.slm_data_d_path}")
            print(f"SLM data E path: {self.slm_data_e_path}")
            
            # Get meter identifier and number
            if not find_datafile or len(find_datafile) < 2:
                raise ValueError(f"Invalid file identifier: {find_datafile}")
            
            meter_id = find_datafile[0].upper()
            file_number = find_datafile[1:].zfill(3)  # e.g., '034' or '282'
            
            print(f"Meter ID: {meter_id}")
            print(f"File number: {file_number}")
            
            # Define paths for different meter types
            raw_testpaths = {
                'A': self.slm_data_d_path,
                'D': self.slm_data_d_path,
                'E': self.slm_data_e_path
            }

            if meter_id not in raw_testpaths:
                raise ValueError(f"Unknown meter identifier: {meter_id}")

            path = raw_testpaths[meter_id]
            datafiles = self.D_datafiles if meter_id in ['A', 'D'] else self.E_datafiles

            print(f"\nSearching in path: {path}")
            print(f"Number of files: {len(datafiles)}")
            print(f"Looking for file number: {file_number}")
            print(f"Data type: {datatype}")

            # Clean up datatype string for matching
            datatype_clean = datatype.replace('.', '').replace('-', '')  # Remove dots and hyphens

            # Look for files that match the actual pattern
            matching_files = []
            for filename in datafiles:
                # Check if both the data type and file number are in the filename
                if (datatype_clean in filename.replace('-', '') and  # Handle data type
                    f".{file_number}.xlsx" in filename):            # Handle file number
                    matching_files.append(filename)
                    print(f"Found matching file: {filename}")

            if not matching_files:
                print("\nNo matches found. Available files:")
                for f in datafiles:
                    print(f"  {f}")
                    print(f"\nWas looking for:")
                    print(f"  - Data type: {datatype_clean}")
                    print(f"  - File number: {file_number}")
                raise ValueError(f"No matching files found for number {file_number} with type {datatype_clean}")

            # Use the most recent file if multiple matches exist
            selected_file = sorted(matching_files)[-1]
            full_path = os.path.join(path, selected_file)

            print(f"\nSelected file: {selected_file}")
            print(f"Full path: {full_path}")

            try:
                if 'RT_Data' in datatype:
                    print(f"Reading Summary sheet for RT data")
                    df = pd.read_excel(
                        full_path, 
                        sheet_name='Summary',
                        engine='openpyxl',
                        header=None  # Important: Don't try to auto-detect headers
                    )
                    measurement_type = 'RT_Data'
                else:
                    print(f"Reading OBA sheet for measurement data")
                    df = pd.read_excel(full_path, sheet_name='OBA', engine='openpyxl')
                    measurement_type = '831_Data'
                
                print("\nRaw DataFrame Structure:")
                print(f"Shape: {df.shape}")
                print(f"Columns: {df.columns.tolist()}")
                print("\nFirst few rows:")
                print(df.head())
                print("\nSearching for 'Frequency (Hz)' in first 10 rows:")
                for idx, row in df.iloc[:10].iterrows():
                    print(f"Row {idx}:", row.tolist())
            
                if df.empty:
                    raise ValueError(f"Empty DataFrame loaded from {full_path}")
                
                # Create SLMData object with file path
                slm_data = SLMData(raw_data=df, measurement_type=measurement_type)
                slm_data.file_path = full_path  # Add this line to store the file path
                
                print("\nSLMData object created:")
                print(f"Measurement type: {measurement_type}")
                print(f"Frequency bands: {len(slm_data.frequency_bands)}")
                print("Raw data shape:", slm_data.raw_data.shape)
                print("First few rows of processed data:")
                print(slm_data.raw_data.head())
                print(f"\nFile path stored: {slm_data.file_path}")
                    
                return slm_data
                
            except Exception as e:
                print(f"\nError reading file:")
                print(f"Exception: {str(e)}")
                print(f"File: {selected_file}")
                print(f"Path attempted: {full_path}")
                raise

        except Exception as e:
            print(f"Error in _raw_slm_datapull: {str(e)}")
            raise

    def _verify_dataframes(self, data_dict: Dict[str, Union[pd.DataFrame, SLMData]]) -> None:
        """Verify and align data objects in dictionary"""
        print("\n=== Starting Data Verification ===")
        # print("Recieve data: ", self.recieve_data)
        # print("SRS data: ", self.srs_data)
        # print("Background data: ", self.bkgrnd_data)
        # print("RT data: ", self.rt)
        print(f"Data dictionary: {data_dict}")

        # Required keys
        required_keys = ['srs_data', 'recive_data', 'bkgrnd_data', 'rt']
        
        # Required frequency bands (1/3 octave)
        required_freq_third = [100, 125, 160, 200, 250, 315, 400, 500, 630, 800, 1000, 1250, 1600, 2000, 2500, 3150, 4000]
        
        # Required frequency bands (1/1 octave)
        required_freq_first = [125, 250, 500, 1000, 2000, 4000]
        
        # Tolerance for frequency matching (in Hz)
        tolerance = 0.1
        
        def is_freq_match(freq1: float, freq2: float) -> bool:
            """Check if two frequencies match within tolerance"""
            return abs(freq1 - freq2) <= tolerance * freq1
        
        def get_freq_column(df: pd.DataFrame) -> pd.Series:
            """Extract frequency column from DataFrame, handling different formats"""
            print("\nAttempting to extract frequency column")
            print(f"DataFrame shape: {df.shape}")
            print(f"Columns: {df.columns.tolist()}")
            
            # Try to find frequency column
            freq_col = None
            
            # Check if DataFrame has a frequency column
            if 'Frequency (Hz)' in df.columns:
                print("Found 'Frequency (Hz)' column")
                freq_col = df['Frequency (Hz)']
            elif '1/1 Octave' in df.columns:
                print("Found '1/1 Octave' format")
                # For 1/1 octave format, first row contains frequencies
                freq_col = pd.to_numeric(df.iloc[0, 1:], errors='coerce')
                print(f"Extracted frequencies: {freq_col.tolist()}")
            elif '1/3 Octave' in df.columns:
                print("Found '1/3 Octave' format")
                # For 1/3 octave format, look for frequency row
                for idx, row in df.iterrows():
                    if 'Frequency (Hz)' in str(row.iloc[0]):
                        freq_col = pd.to_numeric(row.iloc[1:], errors='coerce')
                        print(f"Found frequencies at row {idx}")
                        break
            
            if freq_col is None:
                print("WARNING: Could not find frequency column")
                print("First few rows of data:")
                print(df.head())
                raise ValueError("Could not find frequency column")
                
            freq_col = pd.Series(freq_col.dropna())
            print(f"Final frequency range: {min(freq_col)}-{max(freq_col)} Hz")
            print(f"Number of frequency bands: {len(freq_col)}")
            return freq_col
        
        # Check all required keys are present
        print("\nChecking required keys...")
        for key in required_keys:
            if key not in data_dict:
                raise ValueError(f"Missing required key: {key}")
            if data_dict[key] is None:
                raise ValueError(f"Data for {key} is None")
        
        # First pass: Process measurement data (non-RT)
        processed_dfs = {}
        for key, data in data_dict.items():
            if key == 'rt':
                continue
                
            print(f"\n=== Processing {key} ===")
            
            # Convert SLMData to DataFrame if needed
            if isinstance(data, SLMData):
                print("Converting SLMData to DataFrame")
                df = data.processed_data
            else:
                df = data
                
            if df.empty:
                raise ValueError(f"Empty DataFrame for {key}")
                
            try:
                # Get frequency column
                freq_col = get_freq_column(df)
                print(f"Found {len(freq_col)} frequency bands")
                print(f"Frequency range: {min(freq_col)}-{max(freq_col)} Hz")
                
                # Determine if it's 1/1 or 1/3 octave format
                is_third_octave = any(is_freq_match(freq, 160) for freq in freq_col)
                format_type = "1/3 octave" if is_third_octave else "1/1 octave"
                required_freqs = required_freq_third if is_third_octave else required_freq_first
                print(f"\nDetected {format_type} format")
                
                # Verify required frequencies are present
                missing_freqs = []
                for req_freq in required_freqs:
                    if not any(is_freq_match(freq, req_freq) for freq in freq_col):
                        missing_freqs.append(req_freq)
                
                if missing_freqs:
                    print(f"\nWARNING: Missing frequencies in {key}:")
                    print(f"Missing: {missing_freqs}")
                    print(f"Available: {freq_col.tolist()}")
                    print("\nFirst few rows of data:")
                    print(df.head())
                    raise ValueError(f"Missing required frequencies in {key}: {missing_freqs}")
                
                processed_dfs[key] = df
                    
            except Exception as e:
                print(f"\nError processing {key}:")
                print(f"Error type: {type(e).__name__}")
                print(f"Error message: {str(e)}")
                print("\nDataFrame info:")
                print(df.info())
                print("\nFirst few rows:")
                print(df.head())
                raise ValueError(f"Could not verify frequency data in {key}: {str(e)}")
        
        # Second pass: Process RT data
        print("\n=== Processing RT Data ===")
        try:
            rt_data = data_dict['rt']
            if isinstance(rt_data, SLMData):
                rt_df = rt_data.raw_data
            else:
                rt_df = rt_data
            
            print("RT data shape:", rt_df.shape)
            print("RT data columns:", rt_df.columns.tolist())
            
            # Find RT data headers
            header_row_idx = None
            for idx, row in rt_df.iterrows():
                if 'T30 (ms)' in row.values and 'Frequency (Hz)' in row.values:
                    header_row_idx = int(idx)
                    freq_col_idx = int(row.tolist().index('Frequency (Hz)'))
                    rt_col_idx = int(row.tolist().index('T30 (ms)'))
                    break
            
            if header_row_idx is None:
                print("WARNING: Could not find RT data headers")
                print("First few rows of RT data:")
                print(rt_df.head())
                raise ValueError("Could not find RT data headers")
            
            print(f"Found RT headers at row {header_row_idx}")
            print(f"Frequency column: {freq_col_idx}")
            print(f"RT column: {rt_col_idx}")
            
            # Extract and clean RT data
            data_start_idx = header_row_idx + 1
            freq_data = rt_df.iloc[data_start_idx:, freq_col_idx]
            rt_values = rt_df.iloc[data_start_idx:, rt_col_idx]
            
            # Clean and convert
            freq_data = freq_data.astype(str).str.replace('Hz', '').str.strip()
            freq_data = pd.to_numeric(freq_data, errors='coerce')
            rt_values = pd.to_numeric(rt_values, errors='coerce')
            
            # Remove invalid data
            valid_mask = ~(freq_data.isna() | rt_values.isna())
            freq_data = freq_data[valid_mask]
            rt_values = rt_values[valid_mask]
            
            print("\nRT Data Summary:")
            print(f"Number of valid measurements: {len(freq_data)}")
            print(f"Frequency range: {freq_data.min():.1f} - {freq_data.max():.1f} Hz")
            print(f"RT values range: {rt_values.min():.3f} - {rt_values.max():.3f} ms")
            
            # Create cleaned RT DataFrame
            rt_df_clean = pd.DataFrame({
                'Frequency (Hz)': freq_data,
                'RT60': rt_values
            })
            
            # Update RT data in dictionary
            if isinstance(data_dict['rt'], SLMData):
                data_dict['rt'].frequencies = freq_data.values
                data_dict['rt'].raw_data = rt_df_clean
            else:
                data_dict['rt'] = rt_df_clean
            
        except Exception as e:
            print("\nError processing RT data:")
            print(f"Error type: {type(e).__name__}")
            print(f"Error message: {str(e)}")
            raise ValueError(f"Failed to process RT data: {str(e)}")
        
        # Third pass: Align frequency bands across all measurement data
        print("\n=== Aligning Frequency Bands ===")
        try:
            # Find common frequencies across measurement data
            common_freqs = None
            for df in processed_dfs.values():
                curr_freqs = pd.to_numeric(df['Frequency (Hz)'], errors='coerce')
                curr_freqs = curr_freqs.dropna()
                
                if common_freqs is None:
                    common_freqs = set(curr_freqs)
                else:
                    # Match frequencies within tolerance
                    matched_freqs = set()
                    for freq1 in common_freqs:
                        for freq2 in curr_freqs:
                            if is_freq_match(freq1, freq2):
                                matched_freqs.add(freq1)
                                break
                    common_freqs = matched_freqs
            
            common_freqs = sorted(list(common_freqs))
            print(f"\nFound {len(common_freqs)} common frequency bands:")
            print(f"Range: {min(common_freqs):.1f} - {max(common_freqs):.1f} Hz")
            print("Bands:", [f"{f:.1f}" for f in common_freqs])
            
            # Align each measurement DataFrame to common frequencies
            for key, df in processed_dfs.items():
                print(f"\nAligning {key}...")
                print(f"Original shape: {df.shape}")
                
                # Create frequency matching mask
                mask = pd.to_numeric(df['Frequency (Hz)'], errors='coerce').apply(
                    lambda x: any(is_freq_match(x, f) for f in common_freqs)
                )
                
                # Apply mask and reset index
                aligned_df = df[mask].copy()
                aligned_df.reset_index(drop=True, inplace=True)
                
                print(f"Aligned shape: {aligned_df.shape}")
                
                # Update data in dictionary
                if isinstance(data_dict[key], SLMData):
                    data_dict[key].raw_data = aligned_df
                else:
                    data_dict[key] = aligned_df
            
        except Exception as e:
            print("\nError aligning frequency bands:")
            print(f"Error type: {type(e).__name__}")
            print(f"Error message: {str(e)}")
            raise ValueError(f"Failed to align frequency bands: {str(e)}")
        
        print("\n=== Data Verification Complete ===")

    def get_test_collection(self) -> Dict[str, Dict[TestType, Dict]]:
        """Return the processed test data collection"""
        if not self.test_data_collection:
            raise ValueError("No test data loaded")
        return self.test_data_collection 