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
                    print(f"Test data: {curr_test}")

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
                                
                            # Use the class's own load_test_data method instead of test_processor
                            test_data = self.load_test_data(
                                curr_test=curr_test,
                                test_type=test_type,
                                room_props=room_props
                            )
                            
                            curr_test_data[test_type] = {
                                'room_properties': room_props,
                                'test_data': test_data
                            }
                            
                            if self.debug_mode:
                                print(f"Successfully processed {test_type.value} test")
                                
                        except Exception as e:
                            if self.debug_mode:
                                print(f"Error processing {test_type.value} test: {str(e)}")
                                import traceback
                                traceback.print_exc()

                if curr_test_data:
                    self.test_data_collection[test_label] = curr_test_data
                    if self.debug_mode:
                        print(f"Added test data for {test_label} with types: {list(curr_test_data.keys())}")

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
        """Create AIIC test data instance"""
        aiic_data = base_data.copy()
        additional_data = {
            'AIIC_pos1': self._raw_slm_datapull(curr_test['Position1'], '-831_Data.'),
            'AIIC_pos2': self._raw_slm_datapull(curr_test['Position2'], '-831_Data.'),
            'AIIC_pos3': self._raw_slm_datapull(curr_test['Position3'], '-831_Data.'),
            'AIIC_pos4': self._raw_slm_datapull(curr_test['Position4'], '-831_Data.'),
            'AIIC_source': self._raw_slm_datapull(curr_test['SourceTap'], '-831_Data.'),
            'AIIC_carpet': self._raw_slm_datapull(curr_test['Carpet'], '-831_Data.')
        }
        self._verify_dataframes(additional_data)
        aiic_data.update(additional_data)
        return AIICTestData(room_properties=room_props, test_data=aiic_data)

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
        ### need to put in format_SLM data function???
        if self.debug_mode:
            print("\n=== RAW_SLM_DATAPULL ===")
            print(f"Looking for file identifier: {find_datafile}")
            print(f"Data type: {datatype}")

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

        # Look for files that match the actual pattern:
        # 831_*-*-831_Data.034.xlsx or 831_*-*-RT_Data.034.xlsx
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
            
            if df.empty:
                raise ValueError(f"Empty DataFrame loaded from {full_path}")
            
            # Create SLMData object
            slm_data = SLMData(raw_data=df, measurement_type=measurement_type)
            
            if self.debug_mode:
                print(f"Successfully loaded data with shape: {df.shape}")
                print("Measurement type:", measurement_type)
                print("Frequency bands:", len(slm_data.frequency_bands))
                
            return slm_data
            
        except Exception as e:
            if self.debug_mode:
                print(f"\nError reading file:")
                print(f"Exception: {str(e)}")
                print(f"File: {selected_file}")
                print(f"Path attempted: {full_path}")
            raise