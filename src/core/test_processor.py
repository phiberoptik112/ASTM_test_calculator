from typing import Dict, Union, Optional, List
import pandas as pd
from os.path import join
from pathlib import Path

from data_processor import (
    TestType,
    RoomProperties,
    AIICTestData,
    ASTCTestData,
    NICTestData,
    DTCtestData,
    format_SLMdata,
    extract_sound_levels,
    calculate_onethird_Logavg
)

class TestProcessor:
    def __init__(self, debug_mode: bool = False):
        self.debug_mode = debug_mode
        self.test_data_collection = None  # Initialize as None

    def set_test_collection(self, collection):
        """Set the test data collection reference"""
        self.test_data_collection = collection

    def load_test_data(
        self, 
        curr_test: pd.Series, 
        test_type: TestType, 
        room_props: RoomProperties,
        slm_data_paths: Dict[str, str]
    ) -> Union[AIICTestData, ASTCTestData, NICTestData, DTCtestData]:
        """
        Load and format raw test data based on test type
        
        Args:
            curr_test: Current test row from test plan
            test_type: Type of test to process
            room_props: Room properties for the test
            slm_data_paths: Dictionary containing paths to SLM data files
                {
                    'meter_1': path to first meter data,
                    'meter_2': path to second meter data
                }
        """
        try:
            # Verify test type is enabled
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

            # Load base data common to all test types
            base_data = self._load_base_data(curr_test, slm_data_paths)

            # Create appropriate test data instance based on type
            if test_type == TestType.AIIC:
                return self._create_aiic_test(curr_test, room_props, base_data, slm_data_paths)
            elif test_type == TestType.ASTC:
                return self._create_astc_test(curr_test, room_props, base_data)
            elif test_type == TestType.NIC:
                return self._create_nic_test(curr_test, room_props, base_data)
            elif test_type == TestType.DTC:
                return self._create_dtc_test(curr_test, room_props, base_data, slm_data_paths)
            else:
                raise ValueError(f"Unsupported test type: {test_type}")

        except Exception as e:
            if self.debug_mode:
                print(f"Error loading data: {str(e)}")
                print(f"Current test data: {curr_test}")
            raise

    def _load_base_data(self, curr_test: pd.Series, slm_data_paths: Dict[str, str]) -> Dict[str, pd.DataFrame]:
        """Load common base data required for all test types"""
        if self.debug_mode:
            print('Loading base data')
            
        base_data = {
            'source_data': self._load_slm_data(curr_test['Source'], slm_data_paths['meter_1'], '-831_Data.'),
            'receive_data': self._load_slm_data(self._get_receive_field(curr_test), slm_data_paths['meter_2'], '-831_Data.'),
            'background_data': self._load_slm_data(curr_test['BNL'], slm_data_paths['meter_2'], '-831_Data.'),
            'rt_data': self._load_slm_data(curr_test['RT'], slm_data_paths['meter_2'], '-RT_Data.')
        }

        # Verify data
        self._verify_dataframes(base_data)
        return base_data

    def _create_aiic_test(self, curr_test: pd.Series, room_props: RoomProperties, 
                         base_data: Dict[str, pd.DataFrame], slm_data_paths: Dict[str, str]) -> AIICTestData:
        """Create AIIC test data instance"""
        if self.debug_mode:
            print('Creating AIICTestData instance')
            
        aiic_data = base_data.copy()
        additional_data = {
            'AIIC_pos1': self._load_slm_data(curr_test['Position1'], slm_data_paths['meter_2'], '-831_Data.'),
            'AIIC_pos2': self._load_slm_data(curr_test['Position2'], slm_data_paths['meter_2'], '-831_Data.'),
            'AIIC_pos3': self._load_slm_data(curr_test['Position3'], slm_data_paths['meter_2'], '-831_Data.'),
            'AIIC_pos4': self._load_slm_data(curr_test['Position4'], slm_data_paths['meter_2'], '-831_Data.'),
            'AIIC_source': self._load_slm_data(curr_test['SourceTap'], slm_data_paths['meter_1'], '-831_Data.'),
            'AIIC_carpet': self._load_slm_data(curr_test['Carpet'], slm_data_paths['meter_2'], '-831_Data.')
        }
        
        self._verify_dataframes(additional_data)
        aiic_data.update(additional_data)
        return AIICTestData(room_properties=room_props, test_data=aiic_data)

    def _create_dtc_test(self, curr_test: pd.Series, room_props: RoomProperties, 
                        base_data: Dict[str, pd.DataFrame], slm_data_paths: Dict[str, str]) -> DTCtestData:
        """Create DTC test data instance"""
        if self.debug_mode:
            print('Creating DTCTestData instance')
            
        dtc_data = base_data.copy()
        additional_data = {
            'srs_door_open': self._load_slm_data(curr_test['Source_Door_Open'], slm_data_paths['meter_1'], '-831_Data.'),
            'srs_door_closed': self._load_slm_data(curr_test['Source_Door_Closed'], slm_data_paths['meter_1'], '-831_Data.'),
            'receive_door_open': self._load_slm_data(curr_test['Receive_Door_Open'], slm_data_paths['meter_2'], '-831_Data.'),
            'receive_door_closed': self._load_slm_data(curr_test['Receive_Door_Closed'], slm_data_paths['meter_2'], '-831_Data.')
        }
        
        self._verify_dataframes(additional_data)
        dtc_data.update(additional_data)
        return DTCtestData(room_properties=room_props, test_data=dtc_data)

    def _create_astc_test(self, curr_test: pd.Series, room_props: RoomProperties, 
                         base_data: Dict[str, pd.DataFrame]) -> ASTCTestData:
        """Create ASTC test data instance"""
        if self.debug_mode:
            print('Creating ASTCTestData instance')
        return ASTCTestData(room_properties=room_props, test_data=base_data)

    def _create_nic_test(self, curr_test: pd.Series, room_props: RoomProperties, 
                        base_data: Dict[str, pd.DataFrame]) -> NICTestData:
        """Create NIC test data instance"""
        if self.debug_mode:
            print('Creating NICTestData instance')
        return NICTestData(room_properties=room_props, test_data=base_data)

    @staticmethod
    def _load_slm_data(file_id: str, base_path: str, suffix: str) -> pd.DataFrame:
        """Load SLM data from Excel file"""
        file_path = Path(base_path) / f"{file_id}{suffix}xlsx"
        return pd.read_excel(file_path, engine='openpyxl')

    @staticmethod
    def _verify_dataframes(data_dict: Dict[str, pd.DataFrame]) -> None:
        """Verify all DataFrames in dictionary are valid"""
        for key, df in data_dict.items():
            if df is None or df.empty:
                raise ValueError(f"Invalid or empty DataFrame for {key}")

    @staticmethod
    def _get_receive_field(curr_test: pd.Series) -> str:
        """Handle both spellings of 'Receive'"""
        if 'Receive' in curr_test:
            return curr_test['Receive']
        return curr_test['Recieve']  # Fall back to misspelled version 

    def store_calculated_values(self, test_label: str, test_type: TestType, calculated_values: dict) -> None:
        """Store calculated test values in the test data object for later use in reporting
        
        Args:
            test_label: The test identifier
            test_type: Type of test (AIIC, ASTC, NIC)
            calculated_values: Dictionary containing calculated results
        """
        try:
            if self.debug_mode:
                print(f"\nStoring calculated values for {test_label} ({test_type.value})")
                print(f"Values to store: {list(calculated_values.keys())}")
            
            if self.test_data_collection is None:
                raise ValueError("Test data collection not initialized")
                
            # Get the test data object
            test_data = self.test_data_collection[test_label][test_type]['test_data']
            
            # Store calculated values based on test type
            if test_type == TestType.AIIC:
                test_data.calculated_values = {
                    'NR_val': calculated_values['NR_val'],
                    'AIIC_recieve_corr': calculated_values['AIIC_recieve_corr'],
                    'AIIC_Normalized_recieve': calculated_values['AIIC_Normalized_recieve'],
                    'positions': calculated_values['positions'],
                    'AIIC_contour_val': calculated_values['AIIC_contour_val'],
                    'AIIC_contour_result': calculated_values['AIIC_contour_result'],
                    'room_vol': calculated_values['room_vol'],
                    'sabines': calculated_values['sabines']
                }
                
            elif test_type == TestType.ASTC:
                test_data.calculated_values = {
                    'NR_val': calculated_values['NR_val'],
                    'ATL_val': calculated_values['ATL_val'],
                    'ASTC_final_val': calculated_values['ASTC_final_val'],
                    'ASTC_contour_val': calculated_values['ASTC_contour_val'],
                    'ASTC_recieve_corr': calculated_values['ASTC_recieve_corr'],
                    'sabines': calculated_values['sabines'],
                    'room_vol': calculated_values['room_vol']
                }
                
            elif test_type == TestType.NIC:
                test_data.calculated_values = {
                    'NR_val': calculated_values['NR_val'],
                    'NIC_contour_val': calculated_values['NIC_contour_val'],
                    'NIC_final_val': calculated_values['NIC_final_val'],
                    'NIC_recieve_corr': calculated_values['NIC_recieve_corr'],
                    'sabines': calculated_values['sabines'],
                    'room_vol': calculated_values['room_vol']
                }
                
            if self.debug_mode:
                print(f"Successfully stored calculated values")
                print(f"Stored values: {test_data.calculated_values}")
                
        except Exception as e:
            print(f"Error storing calculated values: {str(e)}")
            raise 