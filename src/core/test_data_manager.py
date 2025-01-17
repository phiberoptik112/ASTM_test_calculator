from typing import Dict, Optional, List, Union
import pandas as pd
from pathlib import Path
import os
from os import listdir
from os.path import isfile, join

from data_processor import (
    TestType,
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
            if self.debug_mode:
                print("\n=== Setting Data Paths ===")
                print(f"Test plan: {test_plan_path}")
                print(f"Meter D: {meter_d_path}")
                print(f"Meter E: {meter_e_path}")
                print(f"Output: {output_path}")
            
            # Verify paths exist
            for path, desc in [
                (test_plan_path, "test plan"),
                (meter_d_path, "meter D data"),
                (meter_e_path, "meter E data"),
                (output_path, "output")
            ]:
                if not os.path.exists(path):
                    raise ValueError(f"Invalid {desc} path: {path}")
            
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
                
                if self.debug_mode:
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
        """Load test plan from Excel file"""
        print(f"Loading test plan from {self.test_plan_path}")
        try:
            self.test_plan = pd.read_excel(self.test_plan_path)
            if self.debug_mode:
                print(f"Loaded test plan with {len(self.test_plan)} rows")
                print("Columns:", self.test_plan.columns.tolist())
        except Exception as e:
            raise ValueError(f"Error loading test plan: {str(e)}")

    def process_test_data(self) -> None:
        """Process all tests in the test plan"""
        if self.test_plan is None:
            raise ValueError("No test plan loaded")

        if self.debug_mode:
            print("\n=== Processing Test Data ===")
            print(f"Number of tests to process: {len(self.test_plan)}")

        for index, curr_test in self.test_plan.iterrows():
            try:
                test_label = curr_test['Test_Label']
                if self.debug_mode:
                    print(f"\nProcessing test {test_label} (row {index})")
                    print(f"Test types enabled:")
                    print(f"AIIC: {curr_test.get('AIIC', 0)}")
                    print(f"ASTC: {curr_test.get('ASTC', 0)}")
                    print(f"NIC: {curr_test.get('NIC', 0)}")
                    print(f"DTC: {curr_test.get('DTC', 0)}")

                # Create room properties
                room_props = self._create_room_properties(curr_test)
                
                # Process each test type
                curr_test_data = {}
                for test_type, column in {
                    TestType.AIIC: 'AIIC',
                    TestType.ASTC: 'ASTC',
                    TestType.NIC: 'NIC',
                    TestType.DTC: 'DTC'
                }.items():
                    if curr_test[column] == 1:
                        try:
                            if self.debug_mode:
                                print(f"\nProcessing {test_type.value} test")
                                
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
                                if self.debug_mode:
                                    print(f"Successfully processed {test_type.value} test")
                                    print(f"Test data type: {type(test_data)}")
                                    print(f"Test data attributes: {dir(test_data)}")
                        except Exception as e:
                            if self.debug_mode:
                                print(f"Error processing {test_type.value} test: {str(e)}")
                                import traceback
                                traceback.print_exc()

                if curr_test_data:
                    self.test_data_collection[test_label] = curr_test_data
                    if self.debug_mode:
                        print(f"\nAdded test data for {test_label}")
                        print(f"Test types: {list(curr_test_data.keys())}")
                        print(f"Data structure: {curr_test_data}")

            except Exception as e:
                if self.debug_mode:
                    print(f"Error processing test row {index}: {str(e)}")
                print(f"Test label: {curr_test.get('Test_Label', 'Unknown')}")
                import traceback
                traceback.print_exc()

        if self.debug_mode:
            print(f"\nProcessed {len(self.test_data_collection)} tests")
        
        # Show test overview window
        if hasattr(self, 'parent_window'):
            from src.gui.components.test_overview import TestOverviewWindow
            TestOverviewWindow(self.parent_window, self)

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

    def _create_room_properties(self, test_row: pd.Series) -> RoomProperties:
        """Create RoomProperties instance from test row"""
        return RoomProperties(
            site_name=test_row['Site_Name'],
            client_name=test_row['Client_Name'],
            source_room_name=test_row['Source Room'],
            source_room=test_row['Source Room'],
            receive_room_name=test_row['Receiving Room'],
            receive_room=test_row['Receiving Room'],
            test_date=test_row['Test Date'],
            report_date=test_row['Report Date'],
            project_name=test_row['Project Name'],
            test_label=test_row['Test_Label'],
            source_vol=test_row['source room vol'],
            receive_vol=test_row['receive room vol'],
            partition_area=test_row['partition area'],
            partition_dim=test_row['partition dim'],
            source_room_finish=test_row['source room finish'],
            receive_room_finish=test_row['receive room finish'],
            srs_floor=test_row['srs floor descrip.'],
            srs_walls=test_row['srs walls descrip.'],
            srs_ceiling=test_row['srs ceiling descrip.'],
            rec_floor=test_row['rec floor descrip.'],
            rec_walls=test_row['rec walls descrip.'],
            rec_ceiling=test_row['rec ceiling descrip.'],
            tested_assembly=test_row['tested assembly'],
            expected_performance=test_row['expected performance'],
            test_assembly_type=test_row['Test assembly Type'],
            annex_2_used=bool(test_row['Annex 2 used?'])
        )

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
            
            if self.debug_mode:
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
        if test_column not in curr_test or curr_test[test_column] != 1:
            if self.debug_mode:
                print(f'Test type {test_type.value} is not enabled for test {curr_test["Test_Label"]}')
            raise ValueError(f"Test type {test_type.value} is not enabled for this test")

        if self.debug_mode:
            print(f'Loading base data for test: {curr_test["Test_Label"]} ({test_type.value})')

        try:
            # Clean up column names in curr_test
            curr_test.index = curr_test.index.str.strip()
            
            # Load base data
            if self.debug_mode:
                print('Loading base data')
            base_data = {
                'srs_data': self._raw_slm_datapull(curr_test['Source'], '-831_Data.'),
                'recive_data': self._raw_slm_datapull(curr_test['Receive'], '-831_Data.'),
                'bkgrnd_data': self._raw_slm_datapull(curr_test['BNL'], '-831_Data.'),
                'rt': self._raw_slm_datapull(curr_test['RT'], '-RT_Data.')
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
            if self.debug_mode:
                print(f"Error loading data: {str(e)}")
                print(f"Current test data: {curr_test}")
            raise

    def _create_aiic_test(self, curr_test: pd.Series, room_props: RoomProperties, base_data: Dict) -> AIICTestData:
        """Create AIIC test data instance with additional validation"""
        try:
            if self.debug_mode:
                print("\nCreating AIIC test data")
                print("Required columns:", ['Position1', 'Position2', 'Position3', 'Position4', 'SourceTap', 'Carpet'])
                print("Available columns:", curr_test.index.tolist())
            
            # Verify required columns exist
            required_cols = ['Position1', 'Position2', 'Position3', 'Position4', 'SourceTap', 'Carpet']
            missing_cols = [col for col in required_cols if col not in curr_test]
            if missing_cols:
                raise ValueError(f"Missing required AIIC columns: {missing_cols}")

            aiic_data = base_data.copy()
            additional_data = {
                'AIIC_pos1': self._raw_slm_datapull(curr_test['Position1'], '-831_Data.'),
                'AIIC_pos2': self._raw_slm_datapull(curr_test['Position2'], '-831_Data.'),
                'AIIC_pos3': self._raw_slm_datapull(curr_test['Position3'], '-831_Data.'),
                'AIIC_pos4': self._raw_slm_datapull(curr_test['Position4'], '-831_Data.'),
                'AIIC_source': self._raw_slm_datapull(curr_test['SourceTap'], '-831_Data.'),
                'AIIC_carpet': self._raw_slm_datapull(curr_test['Carpet'], '-831_Data.')
            }

            # Verify all data was loaded successfully
            for key, data in additional_data.items():
                if data is None or (hasattr(data, 'empty') and data.empty):
                    raise ValueError(f"Failed to load {key} data")

            aiic_data.update(additional_data)
            
            if self.debug_mode:
                print("AIIC data loaded successfully")
                print(f"Number of data points: {len(additional_data)}")
            
            return AIICTestData(room_properties=room_props, test_data=aiic_data)
            
        except Exception as e:
            if self.debug_mode:
                print(f"Error creating AIIC test: {str(e)}")
                import traceback
                traceback.print_exc()
            raise

    def _create_dtc_test(self, curr_test: pd.Series, room_props: RoomProperties, base_data: Dict) -> DTCtestData:
        """Create DTC test data instance"""
        dtc_data = base_data.copy()
        additional_data = {
            'srs_door_open': self._raw_slm_datapull(curr_test['Source_Door_Open'], '-831_Data.'),
            'srs_door_closed': self._raw_slm_datapull(curr_test['Source_Door_Closed'], '-831_Data.'),
            'recive_door_open': self._raw_slm_datapull(curr_test['Receive_Door_Open'], '-831_Data.'),
            'recive_door_closed': self._raw_slm_datapull(curr_test['Receive_Door_Closed'], '-831_Data.')
        }
        self._verify_dataframes(additional_data)
        dtc_data.update(additional_data)
        return DTCtestData(room_properties=room_props, test_data=dtc_data)

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
        """Pull data from SLM files and return as SLMData object"""
        if self.debug_mode:
            print("\n=== RAW_SLM_DATAPULL ===")
            print(f"Looking for file identifier: {find_datafile}")
            print(f"Data type: {datatype}")

        # Clean up input file identifier
        if isinstance(find_datafile, str):
            find_datafile = find_datafile.strip()  # Remove any trailing spaces
        else:
            raise ValueError(f"Invalid file identifier type: {type(find_datafile)}")

        # Get meter identifier and number
        if not find_datafile or len(find_datafile) < 2:
            raise ValueError(f"Invalid file identifier: {find_datafile}")
        
        meter_id = find_datafile[0].upper()
        file_number = find_datafile[1:].zfill(3)  # e.g., '034' or '282'
        
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

        if self.debug_mode:
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
                if self.debug_mode:
                    print(f"Found matching file: {filename}")

        if not matching_files:
            if self.debug_mode:
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

        if self.debug_mode:
            print(f"\nSelected file: {selected_file}")
            print(f"Full path: {full_path}")

        try:
            if 'RT_Data' in datatype:
                print(f"Reading Summary sheet for RT data")
                df = pd.read_excel(full_path, sheet_name='Summary')
                measurement_type = 'RT_Data'
            else:
                print(f"Reading OBA sheet for measurement data")
                df = pd.read_excel(full_path, sheet_name='OBA')
                measurement_type = '831_Data'
            
            if self.debug_mode:
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
            
            if self.debug_mode:
                print("\nSLMData object created:")
                print(f"Measurement type: {measurement_type}")
                print(f"Frequency bands: {len(slm_data.frequency_bands)}")
                print("Raw data shape:", slm_data.raw_data.shape)
                print("First few rows of processed data:")
                print(slm_data.raw_data.head())
                print(f"\nFile path stored: {slm_data.file_path}")
                
            return slm_data
            
        except Exception as e:
            if self.debug_mode:
                print(f"\nError reading file:")
                print(f"Exception: {str(e)}")
                print(f"File: {selected_file}")
                print(f"Path attempted: {full_path}")
            raise

    def _verify_dataframes(self, data_dict: Dict[str, Union[pd.DataFrame, SLMData]]) -> None:
        """
        Verify and align data objects in the dictionary, ensuring required frequency coverage.
        """
        if not data_dict:
            raise ValueError("Empty data dictionary provided")
        
        required_keys = {'srs_data', 'recive_data', 'bkgrnd_data', 'rt'}
        missing_keys = required_keys - set(data_dict.keys())
        if missing_keys:
            raise ValueError(f"Missing required data keys: {missing_keys}")

        # Define required frequency range and tolerance
        REQUIRED_MIN_FREQ = 63.0    # Hz
        REQUIRED_MAX_FREQ = 4000.0  # Hz
        EXPECTED_BINS = 19          # Number of 1/3 octave bands
        FREQ_TOLERANCE = 0.1        # Tolerance for float comparison
        
        # Define required 1/3 octave bands with exact values
        REQUIRED_FREQS = [
            63.0, 80.0,  # Stored as floats in Excel
            100, 125, 160, 200, 250, 315, 400, 500,
            630, 800, 1000, 1250, 1600, 2000, 2500, 3150, 4000
        ]

        def is_freq_match(freq1: float, freq2: float) -> bool:
            """Compare two frequencies within tolerance"""
            return abs(freq1 - freq2) <= FREQ_TOLERANCE

        # PHASE 1: Process and structure all DataFrames
        processed_dfs = {}
        for key, data in data_dict.items():
            if data is None:
                raise ValueError(f"Missing data for {key}")
            
            # Extract DataFrame from SLMData if needed
            df = data.raw_data if isinstance(data, SLMData) else data
            
            if df.empty:
                raise ValueError(f"Empty DataFrame for {key}")
            
            if self.debug_mode:
                print(f"\nProcessing {key}:")
                print(f"Original shape: {df.shape}")
                print(f"Original columns: {df.columns.tolist()}")
                
            # Special handling for RT data
            if key == 'rt':
                try:
                    # Find the row containing 'T30 (ms)' and 'Frequency (Hz)'
                    header_row_idx = None
                    for idx, row in df.iterrows():
                        if 'T30 (ms)' in row.values and 'Frequency (Hz)' in row.values:
                            header_row_idx = int(idx)
                            freq_col_idx = int(row.tolist().index('Frequency (Hz)'))
                            rt_col_idx = int(row.tolist().index('T30 (ms)'))
                            break
                    
                    if header_row_idx is None:
                        raise ValueError("Could not find RT data headers")
                        
                    # Extract data starting from the row after headers
                    data_start_idx = header_row_idx + 1
                    
                    freq_data = df.iloc[data_start_idx:, freq_col_idx]
                    rt_values = df.iloc[data_start_idx:, rt_col_idx]
                    
                    # Clean and convert frequency values
                    freq_data = freq_data.astype(str).str.replace('Hz', '').str.strip()
                    freq_data = pd.to_numeric(freq_data, errors='coerce')
                    rt_values = pd.to_numeric(rt_values, errors='coerce')
                    
                    # Remove any rows where conversion failed
                    valid_mask = ~(freq_data.isna() | rt_values.isna())
                    freq_data = freq_data[valid_mask]
                    rt_values = rt_values[valid_mask]
                    
                    if self.debug_mode:
                        print("\nExtracted data:")
                        print(f"Frequency column index: {freq_col_idx}")
                        print(f"RT column index: {rt_col_idx}")
                        print(f"Number of valid rows: {len(freq_data)}")
                        print(f"Frequency range: {freq_data.min():.1f} - {freq_data.max():.1f} Hz")
                    
                    # Create new DataFrame with cleaned data
                    df = pd.DataFrame({
                        'Frequency (Hz)': freq_data,
                        'RT60': rt_values
                    })
                    
                    # Store the frequency data in the original object
                    if isinstance(data_dict['rt'], SLMData):
                        data_dict['rt'].frequencies = freq_data.values
                        data_dict['rt'].raw_data = df
                    
                    if self.debug_mode:
                        print("\nProcessed RT data:")
                        print(f"Final shape: {df.shape}")
                        print(f"Frequencies: {freq_data.tolist()}")
                        print(f"RT values: {rt_values.tolist()}")
                    
                except Exception as e:
                    raise ValueError(f"Error processing RT data: {str(e)}\nDataFrame info:\n{df.info()}")
            else:
                # Handle 831_Data files
                if self.debug_mode:
                    print(f"\nProcessing {key} DataFrame:")
                    print("Initial shape:", df.shape)
                    print("Initial columns:", df.columns.tolist())

                # Find the header row containing 'Frequency (Hz)' in the 1/3 octave section
                header_row_idx = None
                found_1_3_octave = False

                if self.debug_mode:
                    print("\nSearching for frequency data:")
                    print(f"DataFrame shape: {df.shape}")

                for idx, row in df.iterrows():
                    # Look for the '1/3 Octave' marker first
                    if '1/3 Octave' in str(row.iloc[0]):
                        found_1_3_octave = True
                        if self.debug_mode:
                            print(f"Found '1/3 Octave' marker at row {idx}")
                        continue
                    
                    # After finding '1/3 Octave', look for the frequency header
                    if found_1_3_octave and 'Frequency (Hz)' in row.values:
                        header_row_idx = int(idx)  # Ensure integer
                        freq_col_idx = int(row.tolist().index('Frequency (Hz)'))  # Find column index
                        if self.debug_mode:
                            print(f"Found 'Frequency (Hz)' at row {header_row_idx}, column {freq_col_idx}")
                            print("Header row values:", row.tolist())
                        break

                if header_row_idx is None:
                    if self.debug_mode:
                        print(f"Could not find frequency header in {key}")
                        print("First few rows:")
                        print(df.head())
                    raise ValueError(f"Could not find frequency header in {key}")

                try:
                    # Extract data starting from the header row
                    df = df.iloc[header_row_idx:].copy()
                    df.reset_index(drop=True, inplace=True)  # Reset index after slicing
                    
                    # Use the header row as column names and remove it from the data
                    df.columns = df.iloc[0]
                    df = df.iloc[1:].reset_index(drop=True)
                    
                    # Find the rows containing our required data types
                    required_rows = ['Overall 1/3 Spectra', 'Max 1/3 Spectra', 'Min 1/3 Spectra']
                    data_rows = {}
                    
                    for idx, row in df.iterrows():
                        first_val = str(row.iloc[0]).strip()
                        if first_val in required_rows:
                            data_rows[first_val] = row.iloc[1:].astype(float)
                    
                    # Check if we found all required rows
                    missing_rows = set(required_rows) - set(data_rows.keys())
                    if missing_rows:
                        if self.debug_mode:
                            print(f"Missing required rows: {missing_rows}")
                            print("Available first column values:")
                            print(df.iloc[:, 0].tolist())
                        raise ValueError(f"Missing required data rows in {key}: {missing_rows}")
                    
                    # Create new DataFrame with proper structure
                    freq_values = [float(col) for col in df.columns[1:] if pd.notna(col)]
                    new_df = pd.DataFrame({
                        'Frequency (Hz)': freq_values,
                        'Overall 1/3 Spectra': data_rows['Overall 1/3 Spectra'],
                        'Max 1/3 Spectra': data_rows['Max 1/3 Spectra'],
                        'Min 1/3 Spectra': data_rows['Min 1/3 Spectra']
                    })
                    
                    # Remove any rows with NaN values
                    new_df = new_df.dropna()
                    
                    if self.debug_mode:
                        print("\nProcessed DataFrame:")
                        print("Final shape:", new_df.shape)
                        print("Final columns:", new_df.columns.tolist())
                        print("Frequency range:", f"{new_df['Frequency (Hz)'].min():.1f} - {new_df['Frequency (Hz)'].max():.1f} Hz")
                        print("Number of frequency bands:", len(new_df))
                    
                    df = new_df

                except Exception as e:
                    if self.debug_mode:
                        print(f"\nError processing data:")
                        print(f"Header row index: {header_row_idx}")
                        print(f"Frequency column index: {freq_col_idx}")
                        print(f"DataFrame shape: {df.shape}")
                        print("Exception:", str(e))
                    raise

            if self.debug_mode:
                print(f"Processed shape: {df.shape}")
                print(f"Processed columns: {df.columns.tolist()}")
                if 'Frequency (Hz)' in df.columns:
                    freqs = pd.to_numeric(df['Frequency (Hz)'], errors='coerce')
                    valid_freqs = freqs.dropna()
                    if not valid_freqs.empty:
                        print(f"Frequency range: {valid_freqs.min():.1f} - {valid_freqs.max():.1f} Hz")
                        print(f"Number of frequency bands: {len(valid_freqs)}")
        
            processed_dfs[key] = df

        # Update the original data dictionary with processed DataFrames
        for key, df in processed_dfs.items():
            if isinstance(data_dict[key], SLMData):
                data_dict[key].raw_data = df
            else:
                data_dict[key] = df

        # PHASE 2: Validate and align frequencies
        measurement_dfs = {k: v for k, v in processed_dfs.items() if k != 'rt'}
        
        # Validate frequency ranges for all datasets
        for key, df in processed_dfs.items():
            if 'Frequency (Hz)' not in df.columns:
                raise ValueError(f"{key}: Missing Frequency (Hz) column")
            
            freqs = pd.to_numeric(df['Frequency (Hz)'], errors='coerce')
            freqs = freqs.dropna()
            
            if freqs.empty:
                raise ValueError(f"{key}: No valid frequency values found")
            
            # For measurement data, check required frequency range
            if key != 'rt':
                min_freq = freqs.min()
                max_freq = freqs.max()
                
                if min_freq > (REQUIRED_MIN_FREQ + FREQ_TOLERANCE):
                    raise ValueError(
                        f"{key}: Missing low frequency data. Minimum frequency is {min_freq}Hz "
                        f"but {REQUIRED_MIN_FREQ}Hz is required"
                    )
                
                if max_freq < (REQUIRED_MAX_FREQ - FREQ_TOLERANCE):
                    raise ValueError(
                        f"{key}: Missing high frequency data. Maximum frequency is {max_freq}Hz "
                        f"but {REQUIRED_MAX_FREQ}Hz is required"
                    )

        # Find common frequency bands across measurement DataFrames
        common_freqs = None
        for df in measurement_dfs.values():
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
        
        # Validate common frequency bands
        common_required_freqs = [
            f for f in common_freqs 
            if REQUIRED_MIN_FREQ - FREQ_TOLERANCE <= f <= REQUIRED_MAX_FREQ + FREQ_TOLERANCE
        ]
        
        if len(common_required_freqs) != EXPECTED_BINS:
            raise ValueError(
                f"Insufficient common frequency bands in required range. "
                f"Expected {EXPECTED_BINS} bands, found {len(common_required_freqs)}. "
                f"Available bands: {[f'{f:g}Hz' for f in common_required_freqs]}"
            )
        
        if self.debug_mode:
            print(f"\nFound {len(common_freqs)} common frequency bands")
            print(f"Range: {min(common_freqs):g} - {max(common_freqs):g} Hz")
            print("Required frequency bands present:", [f'{f:g}Hz' for f in common_required_freqs])

        # Align measurement DataFrames to common frequency bands
        aligned_dfs = {}
        for key, df in measurement_dfs.items():
            # Match frequencies within tolerance
            mask = pd.to_numeric(df['Frequency (Hz)'], errors='coerce').apply(
                lambda x: any(is_freq_match(x, f) for f in common_freqs)
            )
            aligned_df = df[mask].copy()
            aligned_df.reset_index(drop=True, inplace=True)
            aligned_dfs[key] = aligned_df
            
            if self.debug_mode:
                print(f"\nAligned {key}:")
                print(f"Original shape: {df.shape}")
                print(f"Aligned shape: {aligned_df.shape}")

        # Update the original data objects with aligned data
        for key, data in data_dict.items():
            if key != 'rt':
                if isinstance(data, SLMData):
                    data.raw_data = aligned_dfs[key]
                else:
                    data_dict[key] = aligned_dfs[key]

        if self.debug_mode:
            print("\nData verification and alignment complete")