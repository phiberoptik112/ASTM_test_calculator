"""
Calculation Service for ASTM Test Calculator

This module handles all test-specific calculations for AIIC, ASTC, and NIC tests.
Extracted from main_window.py to separate business logic from UI logic.
"""

import traceback
import numpy as np
from typing import Dict, Optional, Any, List
from src.core.data_processor import (
    TestType,
    TestData,
    RoomProperties,
    AIICTestData,
    ASTCTestData,
    NICTestData,
    DTCtestData,
    calc_NR_new,
    calc_AIIC_val_claude,
    calc_atl_val,
    calc_astc_val
)


class CalculationService:
    """Service class for handling all test calculations"""
    
    def __init__(self, debug_mode: bool = False):
        self.debug_mode = debug_mode
    
    def calculate_test_values(self, test_type: TestType, test_obj: TestData, freq_data: Dict[str, Any]) -> Optional[Dict[str, Any]]:
        """
        Unified interface for calculating test values based on test type
        
        Args:
            test_type: The type of test (AIIC, ASTC, NIC, DTC)
            test_obj: The test data object
            freq_data: Processed frequency data dictionary
            
        Returns:
            Dictionary of calculated values or None if calculation fails
        """
        try:
            if test_type == TestType.AIIC:
                return self._calculate_aiic_values(test_obj, freq_data)
            elif test_type == TestType.ASTC:
                return self._calculate_astc_values(test_obj, freq_data)
            elif test_type == TestType.NIC:
                return self._calculate_nic_values(test_obj, freq_data)
            elif test_type == TestType.DTC:
                return self._calculate_dtc_values(test_obj, freq_data)
            else:
                print(f"Unknown test type: {test_type}")
                return None
                
        except Exception as e:
            print(f"Error in unified calculation for {test_type}: {str(e)}")
            traceback.print_exc()
            return None
    
    def _calculate_aiic_values(self, test_obj: AIICTestData, freq_data: Dict[str, Any]) -> Optional[Dict[str, Any]]:
        """Calculate AIIC values"""
        try:
            if self.debug_mode:
                print("\nCalculating AIIC values:")
                print("-=-=-=-=-=-=-=-=-=-raw data:-=-=-=--=-=-=-=-=-=")
                print(f"Source: {freq_data['source']}")
                print(f"Avg Pos: {freq_data['avg_pos']}")
                print(f"Background: {freq_data['background']}")
                print(f"Room Props: {freq_data['room_props']}")
                print(f"RT: {freq_data['rt']}")
                print("-=-=-=-=-=-=-=-=-=-raw data:-=-=-=--=-=-=-=-=-=")
            
            # Calculate NR using the averaged position data
            NR_val, _, sabines, AIIC_recieve_corr, ASTC_recieve_corr, AIIC_Normalized_recieve = calc_NR_new(
                freq_data['source'],
                freq_data['avg_pos'],
                None,  # No receive data for AIIC
                freq_data['background'],
                float(freq_data['room_props'].receive_vol),
                freq_data['rt']
            )
            
            if self.debug_mode:
                print(f"AIIC_recieve_corr: {AIIC_recieve_corr}")
                print(f"AIIC_Normalized_recieve: {AIIC_Normalized_recieve}")
                print(f"NR_val: {NR_val}")
                print(f"Calculated sabines value: {sabines}")

            if AIIC_Normalized_recieve is None:
                print("Error: AIIC Normalized receive calculation failed")
                return None

            AIIC_contour_val, AIIC_contour_result = calc_AIIC_val_claude(AIIC_Normalized_recieve)
            
            if self.debug_mode:
                print(f"AIIC_contour_val: {AIIC_contour_val}")
                print(f"AIIC_contour_result: {AIIC_contour_result}")
            
            # Process exceptions
            AIIC_Exceptions = []
            AIIC_exceptions_backcheck = []
            rec_roomvol = float(freq_data['room_props'].receive_vol)

            # Calculate exceptions using stored values
            for val in sabines:
                val_float = float(val)
                AIIC_Exceptions.append('1' if val_float > 2 * rec_roomvol**(2/3) else '0')

            # Background check exceptions
            if self.debug_mode:
                print(f"AIIC_Normalized_recieve: {AIIC_Normalized_recieve}")
                print(f"onethird_bkgrd: {freq_data['background']}")
            
            # Calculate background difference using correct column (overall level)
            background_diff = AIIC_Normalized_recieve - freq_data['background']
            
            # Create exceptions list based on difference
            AIIC_exceptions_backcheck = ['0' if diff > 5 else '1' for diff in background_diff]
            
            if self.debug_mode:
                print(f"AIIC_exceptions_backcheck: {AIIC_exceptions_backcheck}")

            # Create return dictionary
            calculated_values = {
                'NR_val': NR_val,
                'sabines': sabines.copy() if hasattr(sabines, 'copy') else sabines,
                'AIIC_recieve_corr': AIIC_recieve_corr,
                'AIIC_Normalized_recieve': AIIC_Normalized_recieve,
                'AIIC_Exceptions': AIIC_Exceptions,
                'AIIC_exceptions_backcheck': AIIC_exceptions_backcheck,
                'positions': freq_data['positions'],
                'AIIC_contour_val': AIIC_contour_val,
                'AIIC_contour_result': AIIC_contour_result,
                'room_vol': float(freq_data['room_props'].receive_vol)
            }
            
            # Verify the dictionary contains sabines
            if self.debug_mode:
                print("\nVerifying calculated values before return:")
                for key, value in calculated_values.items():
                    print(f"  {key}: {type(value)}")
                    if hasattr(value, 'shape'):
                        print(f"    shape: {value.shape}")
                        if key == 'sabines':
                            print(f"    values: {value}")
            
            return calculated_values
            
        except Exception as e:
            print(f"Error calculating AIIC values: {str(e)}")
            traceback.print_exc()
            return None

    def _calculate_astc_values(self, test_obj: ASTCTestData, freq_data: Dict[str, Any]) -> Optional[Dict[str, Any]]:
        """Calculate ASTC values"""
        try:
            if self.debug_mode:
                print("\nCalculating ASTC values:")
                print("Input data shapes:")
                print(f"Source data: {freq_data['source'].shape if hasattr(freq_data['source'], 'shape') else len(freq_data['source'])}")
                print(f"Receive data: {freq_data['receive'].shape if hasattr(freq_data['receive'], 'shape') else len(freq_data['receive'])}")
                print(f"Background data: {freq_data['background'].shape if hasattr(freq_data['background'], 'shape') else len(freq_data['background'])}")
                print(f"RT data: {freq_data['rt'].shape if hasattr(freq_data['rt'], 'shape') else len(freq_data['rt'])}")
            
            # Ensure all data has correct length before calculation
            expected_length = 17  # We expect 17 frequency points
            if any(len(data) != expected_length for data in [freq_data['source'], freq_data['receive'], 
                                                        freq_data['background'], freq_data['rt']]):
                print("\nData length mismatch detected:")
                print(f"Source length: {len(freq_data['source'])}")
                print(f"Receive length: {len(freq_data['receive'])}")
                print(f"Background length: {len(freq_data['background'])}")
                print(f"RT length: {len(freq_data['rt'])}")
                raise ValueError(f"All frequency data must have length {expected_length}")
                
            NR_val, _, sabines, _, ASTC_recieve_corr, _ = calc_NR_new(
                freq_data['source'],
                None,
                freq_data['receive'],  # No receive data for AIIC
                freq_data['background'],
                float(freq_data['room_props'].receive_vol),
                freq_data['rt']
            )
            
            ATL_val, sabines = calc_atl_val(
                freq_data['source'],
                freq_data['receive'],
                freq_data['background'],
                freq_data['rt'],
                float(freq_data['room_props'].partition_area),
                float(freq_data['room_props'].receive_vol)
            )
            
            if ATL_val is None:
                print("Error: ATL calculation failed")
                return None
            
            if self.debug_mode:
                print(f"ATL values shape: {ATL_val.shape if hasattr(ATL_val, 'shape') else len(ATL_val)}")
                print(f"ATL values: {ATL_val}")
            
            # Calculate ASTC and contour
            ASTC_final_val = calc_astc_val(ATL_val)
            STCCurve = [-16, -13, -10, -7, -4, -1, 0, 1, 2, 3, 4, 4, 4, 4, 4, 4, 4]
            ASTC_contour_val = [val + ASTC_final_val for val in STCCurve]
            
            if self.debug_mode:
                print(f"ASTC final value: {ASTC_final_val}")
            
            return {
                'ATL_val': ATL_val,
                'NR_val': NR_val,
                'ASTC_final_val': ASTC_final_val,
                'ASTC_contour_val': ASTC_contour_val,
                'ASTC_recieve_corr': ASTC_recieve_corr,
                'sabines': sabines,
                'room_vol': float(freq_data['room_props'].receive_vol)
            }
            
        except Exception as e:
            print(f"Error calculating ASTC values: {str(e)}")
            traceback.print_exc()
            return None

    def _calculate_nic_values(self, test_obj: NICTestData, freq_data: Dict[str, Any]) -> Optional[Dict[str, Any]]:
        """Calculate NIC values"""
        try:
            if self.debug_mode:
                print("\nCalculating NIC values:")
                print("Input data validation:")
                for key in ['source', 'receive', 'background', 'rt']:
                    print(f"{key} data length: {len(freq_data[key])}")
                    print(f"{key} values: {freq_data[key]}")
            
            room_vol = test_obj.room_properties.receive_vol
            if self.debug_mode:
                print(f"Room volume: {room_vol}")
            
            # Calculate NIC using the processed data
            NR_val, NIC_contour_val, sabines, _, NIC_recieve_corr, _ = calc_NR_new(
                freq_data['source'],
                None,  # No AIIC data
                freq_data['receive'],
                freq_data['background'],
                room_vol,
                freq_data['rt']
            )
            
            NIC_final_val = calc_astc_val(NR_val)
            
            if self.debug_mode:
                print(f"NIC final value: {NIC_final_val}")
            
            if NR_val is not None:
                if self.debug_mode:
                    print(f"NR values shape: {NR_val.shape if hasattr(NR_val, 'shape') else len(NR_val)}")
                    print(f"NIC contour value: {NIC_contour_val}")
                    print(f"NIC final value: {NIC_final_val}")
                    print(f"Sabines: {sabines}")
                
                return {
                    'NR_val': NR_val,
                    'NIC_contour_val': NIC_contour_val,
                    'NIC_final_val': NIC_final_val,
                    'NIC_recieve_corr': NIC_recieve_corr,
                    'sabines': sabines,
                    'room_vol': room_vol
                }
            
            return None
            
        except Exception as e:
            print(f"Error calculating NIC values: {str(e)}")
            traceback.print_exc()
            return None

    def _calculate_dtc_values(self, test_obj: DTCtestData, freq_data: Dict[str, Any]) -> Optional[Dict[str, Any]]:
        """Calculate DTC values - placeholder for future implementation"""
        try:
            print("DTC calculation not yet implemented")
            return None
            
        except Exception as e:
            print(f"Error calculating DTC values: {str(e)}")
            traceback.print_exc()
            return None