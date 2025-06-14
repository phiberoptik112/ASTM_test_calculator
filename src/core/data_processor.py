import pandas as pd
import numpy as np
import logging
# import yaml
import os 
from os import listdir, walk
from dataclasses import dataclass
from typing import List, Tuple, Dict, Union
from enum import Enum
from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import BaseDocTemplate, PageTemplate, Frame, Paragraph, Table, TableStyle, Spacer, PageBreak, KeepInFrame, Image
from reportlab.lib.units import inch
from reportlab.lib.utils import ImageReader
from io import BytesIO
import matplotlib
matplotlib.use('Agg')  # Set backend to non-interactive
import matplotlib.pyplot as plt
from matplotlib import ticker
import matplotlib.ticker as ticker
from typing import List
import tempfile
import traceback

@dataclass
class RoomProperties:
    site_name: str
    client_name: str
    source_room: str
    receive_room: str
    test_date: str
    report_date: str
    project_name: str
    test_label: str
    source_vol: float
    receive_vol: float
    partition_area: float
    partition_dim: str
    source_room_finish: str
    source_room_name: str
    receive_room_finish: str
    receive_room_name: str
    srs_floor: str
    srs_walls: str
    srs_ceiling: str
    rec_floor: str
    rec_walls: str
    rec_ceiling: str
    tested_assembly: str
    expected_performance: str
    annex_2_used: bool
    test_assembly_type: str

    @classmethod
    def from_dict(cls, data: dict) -> 'RoomProperties':
        """Create RoomProperties from dictionary"""
        return cls(**data)

# each test data class has all the differnet test types in it
@dataclass
class TestData:
    def __init__(self, room_properties: RoomProperties):
        self.room_properties = room_properties
    
    def get_basic_data(self):
        return {
            'srs_data': self.srs_data,
            'recive_data': self.recive_data,
            'bkgrnd_data': self.bkgrnd_data,
            'rt': self.rt
        }

class AIICTestData(TestData):
    def __init__(self, room_properties: RoomProperties, test_data: dict):
        super().__init__(room_properties)
        self.srs_data = test_data['srs_data']
        self.recive_data = test_data['recive_data']
        self.bkgrnd_data = test_data['bkgrnd_data']
        self.rt = test_data['rt']
        self.pos1 = test_data['AIIC_pos1']
        self.pos2 = test_data['AIIC_pos2']
        self.pos3 = test_data['AIIC_pos3']
        self.pos4 = test_data['AIIC_pos4']
        self.source = test_data['AIIC_source']
        self.carpet = test_data['AIIC_carpet']
        
class DTCtestData(TestData):
    def __init__(self, room_properties: RoomProperties, test_data: dict):
        super().__init__(room_properties)
        self.srs_data = test_data['srs_data']
        self.recive_data = test_data['recive_data']
        self.bkgrnd_data = test_data['bkgrnd_data']
        self.rt = test_data['rt']
        self.srs_door_open = test_data['srs_door_open']
        self.srs_door_closed = test_data['srs_door_closed']
        self.recive_door_open = test_data['recive_door_open']
        self.recive_door_closed = test_data['recive_door_closed']

class ASTCTestData(TestData):
    def __init__(self, room_properties: RoomProperties, test_data: dict):
        super().__init__(room_properties)
        self.srs_data = test_data['srs_data']
        self.recive_data = test_data['recive_data']
        self.bkgrnd_data = test_data['bkgrnd_data']
        self.rt = test_data['rt']

class NICTestData(TestData):
    def __init__(self, room_properties: RoomProperties, test_data: dict):
        super().__init__(room_properties)
        self.srs_data = test_data['srs_data']
        self.recive_data = test_data['recive_data']
        self.bkgrnd_data = test_data['bkgrnd_data']
        self.rt = test_data['rt']

class TestType(Enum):
    AIIC = "AIIC"
    ASTC = "ASTC"
    NIC = "NIC"
    DTC = "DTC"

######### SLM import - hardcoded - gives me the willies, but it works for now.##########

def format_SLMdata(srs_data):
    ## verify that srs_data iloc[7] is correct- will have a label as 1/3 octave
    srs_thirdoct = srs_data.iloc[7] # hardcoded to SLM export data format
    # srssrs_thirdoct = srs_thirdoct[13:31] # previous code, only goes from 125 to 3150
    srs_thirdoct = srs_thirdoct[13:32] # 100hz to 4000hz

    # verify step to check for SPL info? 
    return srs_thirdoct

def extract_sound_levels(slm_data: pd.DataFrame, 
                        freq_bands: list = [100, 125, 160, 200, 250, 315, 400, 500, 630, 800, 
                                          1000, 1250, 1600, 2000, 2500, 3150, 4000]) -> pd.Series:
    """
    Extracts sound pressure level measurements for 1/3 octave bands from SLM export file.
    
    Args:
        slm_data: DataFrame containing SLM measurement data
        freq_bands: List of frequency bands in Hz to validate against
        
    Returns:
        pd.Series: Sound pressure levels (dB) for each frequency band
        
    Raises:
        ValueError: If data format is incorrect or SPL data cannot be validated
    """
    try:
        # Find the overall spectrum row (typically contains "Overall 1/3 Spectra" or similar)
        spectrum_row_idx = slm_data.apply(lambda x: x.astype(str).str.contains('Overall 1/3 Spectra')).any(axis=1).idxmax()
        
        # Get frequency labels (one row above spectrum data)
        freq_labels = pd.to_numeric(slm_data.iloc[spectrum_row_idx - 1, 13:32], errors='coerce')
        
        # Validate frequency bands
        if not all(f in freq_labels.values for f in freq_bands):
            raise ValueError(f"Expected frequency bands not found.\nExpected: {freq_bands}\nFound: {freq_labels.values}")
        
        # Extract SPL values
        spl_data = pd.to_numeric(slm_data.iloc[spectrum_row_idx, 13:31], errors='coerce')
        
        # Validate SPL data
        if spl_data.isnull().any():
            raise ValueError("Invalid or missing SPL values found")
        
        if not (0 <= spl_data).all() and (spl_data <= 140).all():
            raise ValueError("SPL values outside expected range (0-140 dB)")
            
        return spl_data[freq_labels.index[freq_labels.isin(freq_bands)]]
        
    except Exception as e:
        raise ValueError(f"Failed to extract SPL data: {str(e)}\n"
                        f"Please verify SLM data format is correct.") from e
### useage: 
# try:
#     spectrum_data = extract_third_octave_bands(slm_data)
#     print("Successfully extracted frequency bands:", spectrum_data)
# except ValueError as e:
#     print(f"Error processing SLM data: {e}")


def calculate_onethird_Logavg(average_pos):
    # Convert list of position arrays to numpy array for easier manipulation
    pos_array = np.array(average_pos)
    # Transpose to get frequency bins as first dimension
    pos_array = pos_array.T
    
    onethird_rec_Total = []
    # Now iterate over frequency bins
    for freqbin in pos_array:
        total = 0
        count = 0
        for val in freqbin:
            if not pd.isnull(val):
                total += 10**(val/10)
                count += 1
        if count > 0:
            average = total / count
            onethird_rec_Total.append(10 * np.log10(average))
        else:
            onethird_rec_Total.append(np.nan)
            
    onethird_rec_Total = np.round(onethird_rec_Total, 1)
    return onethird_rec_Total


def calc_NR_new(srs_overalloct, AIIC_rec_overalloct, ASTC_rec_overalloct, bkgrnd_overalloct, recieve_roomvol, rt_thirty):
    # Initialize return values
    NR_val = None
    NIC_final_val = None
    AIIC_recieve_corr = []
    ASTC_recieve_corr = []
    AIIC_Normalized_recieve = None
    NIC_start = 16
    
    try:
        # Determine test type and expected length based on input
        if AIIC_rec_overalloct is not None:
            test_type = "AIIC"
            expected_length = 16  # 100-3150 Hz
            freq_range = "100-3150 Hz"
            STCCurve = [-16, -13, -10, -7, -4, -1, 0, 1, 2, 3, 4, 4, 4, 4, 4, 4]
        else:
            test_type = "ASTC/NIC"
            expected_length = 17  # 100-4000 Hz
            freq_range = "100-4000 Hz"
            STCCurve = [-16, -13, -10, -7, -4, -1, 0, 1, 2, 3, 4, 4, 4, 4, 4, 4, 4]

        print(f"\n=== Processing {test_type} Test Data ===")
        print(f"Expected: {expected_length} values ({freq_range})")
        
        # Debug input data types and shapes
        print("\nInput Data Analysis:")
        print(f"Source data type: {type(srs_overalloct)}")
        print(f"Source length: {len(srs_overalloct) if hasattr(srs_overalloct, '__len__') else 'no length'}")
        if isinstance(srs_overalloct, pd.Series):
            print(f"Source index: {srs_overalloct.index.tolist()}")
        
        print(f"\nBackground data type: {type(bkgrnd_overalloct)}")
        print(f"Background length: {len(bkgrnd_overalloct) if hasattr(bkgrnd_overalloct, '__len__') else 'no length'}")
        if isinstance(bkgrnd_overalloct, pd.Series):
            print(f"Background index: {bkgrnd_overalloct.index.tolist()}")
        
        print(f"\nRT thirty data type: {type(rt_thirty)}")
        print(f"RT thirty length: {len(rt_thirty) if hasattr(rt_thirty, '__len__') else 'no length'}")
        if isinstance(rt_thirty, pd.Series):
            print(f"RT thirty index: {rt_thirty.index.tolist()}")
        
        if AIIC_rec_overalloct is not None:
            print(f"\nAIIC receive data type: {type(AIIC_rec_overalloct)}")
            print(f"AIIC receive length: {len(AIIC_rec_overalloct) if hasattr(AIIC_rec_overalloct, '__len__') else 'no length'}")
        
        if ASTC_rec_overalloct is not None:
            print(f"\nASTC receive data type: {type(ASTC_rec_overalloct)}")
            print(f"ASTC receive length: {len(ASTC_rec_overalloct) if hasattr(ASTC_rec_overalloct, '__len__') else 'no length'}")

        # Convert inputs to numpy arrays and truncate RT data if needed
        bkgrnd_overalloct = pd.to_numeric(bkgrnd_overalloct).to_numpy() if isinstance(bkgrnd_overalloct, pd.Series) else np.array(bkgrnd_overalloct)
        rt_thirty = rt_thirty.to_numpy() if isinstance(rt_thirty, pd.Series) else np.array(rt_thirty)
        srs_overalloct = pd.to_numeric(srs_overalloct).to_numpy() if isinstance(srs_overalloct, pd.Series) else np.array(srs_overalloct)

        # Truncate RT data to match expected length
        if len(rt_thirty) > expected_length:
            rt_thirty = rt_thirty[:expected_length]
            print(f"\nTruncated RT data to {expected_length} points")

        # Verify array lengths after truncation
        input_lengths = {
            'Source': len(srs_overalloct) if hasattr(srs_overalloct, '__len__') else None,
            'Background': len(bkgrnd_overalloct) if hasattr(bkgrnd_overalloct, '__len__') else None,
            'RT thirty': len(rt_thirty) if hasattr(rt_thirty, '__len__') else None,
            'AIIC receive': len(AIIC_rec_overalloct) if AIIC_rec_overalloct is not None and hasattr(AIIC_rec_overalloct, '__len__') else None,
            'ASTC receive': len(ASTC_rec_overalloct) if ASTC_rec_overalloct is not None and hasattr(ASTC_rec_overalloct, '__len__') else None
        }

        # Check for length mismatches
        length_issues = {name: length for name, length in input_lengths.items() 
                        if length is not None and length != expected_length}
        
        if length_issues:
            error_msg = f"\nLength mismatch detected for {test_type} test:\n"
            error_msg += f"Expected length: {expected_length}\n"
            error_msg += "Actual lengths:\n"
            print(f"srs_overalloct: {srs_overalloct}")
            print(f"AIIC_rec_overalloct: {AIIC_rec_overalloct}")
            print(f"ASTC_rec_overalloct: {ASTC_rec_overalloct}")
            print(f"bkgrnd_overalloct: {bkgrnd_overalloct}")
            print(f"rt_thirty: {rt_thirty}")
            for name, length in length_issues.items():
                error_msg += f"- {name}: {length}\n"
            raise ValueError(error_msg)

        # Calculate sabines
        sabines = 0.049 * recieve_roomvol/rt_thirty

        print("\nInput shapes:")
        print(f"Background: {bkgrnd_overalloct.shape}")
        print(f"RT thirty: {rt_thirty.shape}")
        print(f"Source: {srs_overalloct.shape}")

        # Process AIIC data if provided
        if AIIC_rec_overalloct is not None:
            try:
                # Convert to numpy arrays and ensure float64 dtype
                AIIC_rec_overalloct = np.array(AIIC_rec_overalloct, dtype=np.float64)
                bkgrnd_overalloct_aiic = np.array(bkgrnd_overalloct, dtype=np.float64)
                
                # Verify shapes match
                if AIIC_rec_overalloct.shape != bkgrnd_overalloct_aiic.shape:
                    raise ValueError(f"Shape mismatch: AIIC={AIIC_rec_overalloct.shape}, background={bkgrnd_overalloct_aiic.shape}")
                
                print("\nDebug - Arrays before calculation:")
                print(f"AIIC data: {AIIC_rec_overalloct}")
                print(f"Background data: {bkgrnd_overalloct_aiic}")
                print(f"Sabines data: {sabines}")
                
                # Calculate difference
                AIIC_recieve_vsBkgrnd = AIIC_rec_overalloct - bkgrnd_overalloct_aiic
                
                # Initialize output array
                AIIC_recieve_corr = np.zeros_like(AIIC_rec_overalloct)
                
                # Apply corrections using numpy where
                mask_lt_5 = AIIC_recieve_vsBkgrnd < 5
                
                # Apply the two different calculations based on the mask
                AIIC_recieve_corr = np.where(
                    mask_lt_5,
                    AIIC_rec_overalloct - 2,  # When difference < 5
                    10 * np.log10(np.abs(np.power(10, AIIC_rec_overalloct/10) - np.power(10, bkgrnd_overalloct_aiic/10)))  # When difference >= 5
                )
                
                # Round the results
                AIIC_recieve_corr = np.round(AIIC_recieve_corr, decimals=1)
                
                print("\nDebug - After correction:")
                print(f"AIIC_recieve_corr: {AIIC_recieve_corr}")
                
                # Ensure sabines is a numpy array of float64 and rounded to 1 decimal place
                sabines = np.array(sabines, dtype=np.float64).round(0)
                
                # Check for invalid values in sabines
                if np.any(sabines <= 0):
                    raise ValueError("Sabines contains zero or negative values")
                
                # Calculate normalized receive
                scaling = 10 * np.log10(108/sabines)
                AIIC_Normalized_recieve = AIIC_recieve_corr - scaling
                AIIC_Normalized_recieve = np.round(AIIC_Normalized_recieve, decimals=1)
                
                print("\nDebug - Final normalized values:")
                print(f"Scaling: {scaling}")
                print(f"AIIC_Normalized_recieve: {AIIC_Normalized_recieve}")
                print("sabines: ",sabines)
                
            except Exception as e:
                print(f"\nDetailed error in AIIC processing:")
                print(f"Error type: {type(e).__name__}")
                print(f"Error message: {str(e)}")
                print(f"AIIC_rec_overalloct shape: {getattr(AIIC_rec_overalloct, 'shape', 'unknown')}")
                print(f"bkgrnd_overalloct shape: {getattr(bkgrnd_overalloct, 'shape', 'unknown')}")
                print(f"sabines shape: {getattr(sabines, 'shape', 'unknown')}")
                AIIC_recieve_corr = np.array([], dtype=np.float64)
                AIIC_Normalized_recieve = None

        # Process ASTC data if provided
        if ASTC_rec_overalloct is not None:
            try:
                # Convert to numpy array and ensure float64 type
                ASTC_rec_overalloct = pd.to_numeric(ASTC_rec_overalloct).to_numpy(dtype=np.float64) if isinstance(ASTC_rec_overalloct, pd.Series) else np.array(ASTC_rec_overalloct, dtype=np.float64)
                
                # Verify ASTC array length
                if len(ASTC_rec_overalloct) != expected_length:
                    raise ValueError(f"ASTC receive array must have length {expected_length}")

                ASTC_recieve_vsBkgrnd = ASTC_rec_overalloct - bkgrnd_overalloct
                
                ASTC_recieve_corr = []
                for i, val in enumerate(ASTC_recieve_vsBkgrnd):
                    if val < 5:
                        ASTC_recieve_corr.append(float(ASTC_rec_overalloct[i])-2)
                    elif val < 10:
                        input_astc = 10**(float(ASTC_rec_overalloct[i])/10)
                        background = 10**(float(bkgrnd_overalloct[i])/10)
                        input_vs_bkgrnd = input_astc - background
                        if input_vs_bkgrnd < 0:
                            input_vs_bkgrnd = abs(input_vs_bkgrnd)
                        ASTC_recieve_corr.append(10*np.log10(input_vs_bkgrnd))
                    else:
                        ASTC_recieve_corr.append(float(ASTC_rec_overalloct[i]))
                
                # Convert to numpy array before rounding
                ASTC_recieve_corr = np.array(ASTC_recieve_corr, dtype=np.float64)
                
                # Calculate NR and NIC values
                NR_val = np.array(srs_overalloct, dtype=np.float64) - ASTC_recieve_corr
                NR_val = np.round(NR_val, decimals=1)  # Use numpy round instead of pd.to_numeric
                
                # NIC curve calculation
                NIC_val_list = NR_val.copy()
                NIC_final_val = calculate_nic_curve(NIC_val_list, STCCurve, NIC_start)
            except Exception as e:
                print(f"Error in ASTC processing: {str(e)}")
                print(f"ASTC_recieve_corr type: {type(ASTC_recieve_corr)}")
                print(f"NR_val type: {type(NR_val) if 'NR_val' in locals() else 'Not created'}")
                ASTC_recieve_corr = np.array([], dtype=np.float64)
                NR_val = None
                NIC_final_val = None

        print("\nFinal shapes:")
        print(f"NR_val: {getattr(NR_val, 'shape', 'None')}")
        print(f"AIIC_recieve_corr: {AIIC_recieve_corr.shape if len(AIIC_recieve_corr) > 0 else 'Empty'}")
        print(f"ASTC_recieve_corr: {ASTC_recieve_corr.shape if len(ASTC_recieve_corr) > 0 else 'Empty'}")
        print(f"AIIC_Normalized_recieve: {getattr(AIIC_Normalized_recieve, 'shape', 'None')}")

    except Exception as e:
        print(f"Error in calc_NR_new: {str(e)}")
        raise

    return NR_val, NIC_final_val, sabines, AIIC_recieve_corr, ASTC_recieve_corr, AIIC_Normalized_recieve

def calculate_nic_curve(NIC_val_list, STCCurve, NIC_start):
    """Helper function to calculate NIC curve"""
    diff_negative = 0
    new_sum = 0
    while (diff_negative <= 8 and new_sum <= 32):
        New_curve = [val + NIC_start for val in STCCurve]
        NIC_curve = New_curve - NIC_val_list
        NIC_curve = np.round(NIC_curve)
        diff_negative = np.max(NIC_curve)
        
        pos_diffs = [np.round(val) if val > 0 else 0 for val in NIC_curve]
        new_sum = np.sum(pos_diffs)
        
        if new_sum > 32 or diff_negative > 8:
            return NIC_start - 1
            
        NIC_start += 1
        if NIC_start > 80:
            return 80
    
    return NIC_start - 1

### this code revised 7/24/24 - functional and produces accurate ATL values

def calc_atl_val(srs_overalloct: pd.Series, rec_overalloct: pd.Series, 
                 bkgrnd_overalloct: pd.Series, rt_thirty: pd.Series, 
                 partition_area: float, receive_roomvol: float) -> pd.Series:
    """Calculate ATL values"""
    try:
        print("\nATL Calculation Input Validation:")
        print(f"Source data shape: {srs_overalloct.shape if hasattr(srs_overalloct, 'shape') else len(srs_overalloct)}")
        print(f"Receive data shape: {rec_overalloct.shape if hasattr(rec_overalloct, 'shape') else len(rec_overalloct)}")
        print(f"Background data shape: {bkgrnd_overalloct.shape if hasattr(bkgrnd_overalloct, 'shape') else len(bkgrnd_overalloct)}")
        print(f"RT data shape: {rt_thirty.shape if hasattr(rt_thirty, 'shape') else len(rt_thirty)}")
        
        ASTC_vollimit = 883
        if receive_roomvol > ASTC_vollimit:
            print('Using NIC calc, room volume too large')
        # constant = np.int32(20.047*np.sqrt(273.15+20))
        # intermed = 30/rt_thirty ## why did i do this? not right....sabines calc is off
        # thisval = np.int32(recieve_roomvol*intermed)
        # sabines =thisval/constant

        # RT value is right, why is this not working? - RT value MUST NOT BE ROUNDED TO 2 DECIMALS
        # print('recieve roomvol: ',recieve_roomvol)
        if isinstance(rt_thirty, pd.DataFrame):
            rt_thirty = rt_thirty.values
        print('rt_thirty: ',rt_thirty)
        print('receive roomvol: ',receive_roomvol)
        sabines = 0.049*receive_roomvol/rt_thirty  # this produces accurate sabines values
        ## sabines gets this calc, but is still not accurate to spreadhseet reference, even with the right rt_thirty value
        if isinstance(bkgrnd_overalloct, pd.DataFrame):
            bkgrnd_overalloct = bkgrnd_overalloct.values
        print('bkgrnd_overalloct: ',bkgrnd_overalloct)
        sabines = np.int32(sabines) 
        sabines = np.round(sabines)
        # print('sabines: ',sabines)
        
        recieve_corr = list()
        # print('recieve: ',rec_overalloct)
        recieve_vsBkgrnd = rec_overalloct - bkgrnd_overalloct

        if isinstance(recieve_vsBkgrnd, pd.DataFrame):
            recieve_vsBkgrnd = recieve_vsBkgrnd.values

        if isinstance(rec_overalloct, pd.DataFrame):
            rec_overalloct = rec_overalloct.values
        print('recieve vs  background: ',recieve_vsBkgrnd)
        # print('recieve roomvol: ',recieve_roomvol)
        #### something wrong with this loop #### 
        for i, val in enumerate(recieve_vsBkgrnd):
            if val < 5:
                print('recieve vs background: ',val)
                recieve_corr.append(rec_overalloct[i]-2)
                print('less than 5, appending: ',recieve_corr[i])
            elif val < 10:
                print('recieve vs background: ',val)
                recieve_corr.append(10*np.log10((10**(rec_overalloct[i]/10))-(10**(bkgrnd_overalloct[i]/10))))
                print('less than 10, appending: ',recieve_corr[i])
            else:
                print('recieve vs background: ',val)
                recieve_corr.append(rec_overalloct[i])
                print('greater than 10, appending: ',recieve_corr[i])
            
        
        # print('recieve correction: ',recieve_corr)
        if isinstance(srs_overalloct, pd.DataFrame):
            recieve_corr = recieve_corr.values
        if isinstance(srs_overalloct, pd.DataFrame):
            srs_overalloct = srs_overalloct.values
        if isinstance(sabines, pd.DataFrame):
            sabines = sabines.values
        print('srs overalloct: ',srs_overalloct)
        print('recieve correction: ',recieve_corr)
        print('sabines: ',sabines)

        # ATL val initalized to 0
        # ATL_val = []
        print('=-=-=-=-=-=-=-ATL val initalized to 0-=-=-=-=-=-=-=-')
        ATL_val = np.zeros(len(srs_overalloct))
        # for i, val in enumerate(srs_overalloct):
        #     ATL_val.append(srs_overalloct[i]-recieve_corr[i]+10*(np.log10(partition_area/sabines.iloc[i])))

        # Ensure all inputs are numpy arrays
        srs_overalloct = np.array(srs_overalloct)
        recieve_corr = np.array(recieve_corr)
        sabines = np.array(sabines)
        

        # Ensure all arrays have same length before calculation
        min_length = min(len(srs_overalloct), len(recieve_corr), len(sabines))
        srs_overalloct = srs_overalloct[:min_length]
        recieve_corr = recieve_corr[:min_length] 
        sabines = sabines[:min_length]
        
        # Vectorized calculation
        # print('=-=-=-=-=-=-=-Vectorized calculation-=-=-=-=-=-=-=-')
        ATL_val = srs_overalloct - recieve_corr + 10 * np.log10(partition_area / sabines)
        # Convert ATL_val to numpy array if not already
        # ATL_val = np.array(ATL_val)
        # ATL_val = srs_overalloct - recieve_corr+10*(np.log(parition_area/sabines)) 
        print('-=-=-Rounding ATL val-=-=-=')
        ATL_val = np.round(ATL_val,1)
        print('ATL val: ',ATL_val)
        return ATL_val, sabines
    except Exception as e:
        print(f"Error in calc_atl_val: {str(e)}")
        raise

def calc_AIIC_val_claude(Normalized_recieve_IIC, verbose=True):
    pos_diffs = list()
    diff_negative_min = 0
    AIIC_start = 94
    AIIC_contour_val = 16
    IIC_contour = list()
    AIIC_curve = list()
    new_sum = 0
    diff_negative_max = 0
    # IIC_curve = [2,2,2,2,2,2,1,0,-1,-2,-3,-6,-9,-12,-15,-18]
    # shorening curve for properly representing the frq bands used, not using 63hz, 4 and 5khz
    IIC_curve = [2,2,2,2,2,2,1,0,-1,-2,-3,-6,-9,-12,-15,-18]
    max_iterations = 100  # Maximum number of iterations to prevent infinite loop
    iteration_count = 0
    # initial application of the IIC curve to the first AIIC start value 
    for vals in IIC_curve:
        IIC_contour.append(vals+AIIC_start)
    Normalized_recieve_IIC = pd.to_numeric(Normalized_recieve_IIC, errors='coerce')
    Normalized_recieve_IIC = np.array(Normalized_recieve_IIC)
    Normalized_recieve_IIC = np.round(Normalized_recieve_IIC)
    # Normalized_recieve_IIC = Normalized_recieve_IIC[1:14] # values from 125-3150 (not including 4khz)
    print('Normalized recieve ANISPL: ', Normalized_recieve_IIC)
    # print('length of normalized recieve: ',len(Normalized_recieve_IIC))
    # print('length of IIC contour: ',len(IIC_contour))
    # Contour_curve_result = IIC_contour - Normalized_recieve_IIC
    # print('IIC_contour: ',IIC_contour)
    Contour_curve_result =  Normalized_recieve_IIC - IIC_contour
    Contour_curve_result = np.round(Contour_curve_result,1)
    
    print('Contour curve: ', Contour_curve_result)
    # Debug 315Hz bin (index 4)
    # print("\nDebugging 315Hz in AIIC calculation:")
    # print(f"315Hz Normalized input: {Normalized_recieve_IIC[4]}")
    # print(f"315Hz IIC contour value: {IIC_contour[4]}")
    
    while (diff_negative_max <= 8 and new_sum <= 32 and iteration_count < max_iterations):
        if verbose:
            print(f"\nIteration {iteration_count}:")
            print(f"AIIC_contour_val: {AIIC_contour_val}")
            print(f"Max difference: {diff_negative_max}")
            print(f"Sum of positive differences: {new_sum}")
            print(f"Current contour: {IIC_contour}")
            # print(f"315Hz Difference: {diff_negative[4]}")
            # print(f"315Hz Positive diff: {pos_diffs[4]}")
        
        diff_negative = Normalized_recieve_IIC - IIC_contour
        diff_negative_max = np.max(diff_negative)
        diff_negative = pd.to_numeric(diff_negative, errors='coerce')
        diff_negative = np.array(diff_negative)
        
        pos_diffs = [np.round(val,1) if val > 0 else 0 for val in diff_negative]
        new_sum = np.sum(pos_diffs)
        
        if new_sum > 32 or diff_negative_max > 8:
            print('Difference condition met! AIIC value: ', AIIC_contour_val-1)
            print('AIIC result curve: ', Contour_curve_result)
            return AIIC_contour_val-1, Contour_curve_result
        else:
            AIIC_start -= 1
            AIIC_contour_val += 1
            IIC_contour = [vals + AIIC_start for vals in IIC_curve]
            
            Contour_curve_result =  Normalized_recieve_IIC - IIC_contour
            # Contour_curve_result =  IIC_contour - Normalized_recieve_IIC
            iteration_count += 1

    if iteration_count == max_iterations:
        print(f"Maximum iterations ({max_iterations}) reached without meeting conditions.")
    else:
        print("Loop completed without meeting conditions. Returning last calculated values.")
    print(f"Loop exited. Final values:")
    print(f"  iterations: {iteration_count}")
    print('Contour curve (IIC curve minus ANISPL): ', Contour_curve_result)


    return AIIC_contour_val, Contour_curve_result

def calc_astc_val(atl_val: pd.Series) -> float:
    pos_diffs = list()
    diff_negative = 0
    diff_positive = 0 
    ASTC_start = 16
    New_curve = list()
    new_sum = 0
    # print('__________CALCULATING ASTC__________')
    # Debug input shape
    print(f"\nInput ATL_val shape: {atl_val.shape}")
    # Since ATL values only go from 125 to 4k, we need to match array lengths
    # ATL_val is length 16, so we need to adjust STCCurve to match  
    print('ATL val: ', atl_val)
    STCCurve = [-16, -13, -10, -7, -4, -1, 0, 1, 2, 3, 4, 4, 4, 4, 4]  # Length 15
    
    # Convert ATL_val to numpy array if it isn't already
    if isinstance(atl_val, pd.Series):
        atl_val = atl_val.to_numpy()
    
    # Ensure we're using the right slice of ATL values
    # we need to truncate the ATL_val to 16 values from 125-4000
    ATL_val_STC = atl_val[1:16]  # must use values from 125-4000, not 100-4000

    print('ATL val STC: ', ATL_val_STC)
    print('shape of ATL val STC: ', ATL_val_STC.shape)
    while (diff_negative <= 8 and new_sum <= 32):
        print('ASTC fit test value: ', ASTC_start)
        
        # Create new curve with matching length
        New_curve = np.array([val + ASTC_start for val in STCCurve])
        
        # Verify shapes before subtraction
        # print(f"New_curve shape: {New_curve.shape}")
        # print(f"ATL_val_STC shape: {ATL_val_STC.shape}")
        
        # Calculate ASTC curve
        ASTC_curve = New_curve - ATL_val_STC
        ASTC_curve = np.round(ASTC_curve) ## may need a better rounding method
        
        # print('ASTC curve: ', ASTC_curve)
        diff_negative = np.max(ASTC_curve)
        # print('Max, single diff: ', diff_negative)

        # Calculate positive differences
        pos_diffs = [np.round(val) if val > 0 else 0 for val in ASTC_curve]
        new_sum = np.sum(pos_diffs)
        # print('Sum Positive diffs: ', new_sum)
        
        if new_sum > 32 or diff_negative > 8:
            print('Curve too high! ASTC fit: ', ASTC_start-1) 
            return ASTC_start-1
            
        pos_diffs = []
        New_curve = []
        ASTC_start = ASTC_start + 1
        
        if ASTC_start > 80: 
            break
    print('ASTC final value: ', ASTC_start-1)
    return ASTC_start-1

def plot_curves(frequencies: List[float], y_label: str, ref_curve: np.ndarray, 
                field_curve: np.ndarray, ref_label: str, field_label: str) -> ImageReader:
    """Creates a plot and returns it in a format compatible with ReportLab.
    
    Args:
        frequencies: List of frequency values for x-axis (length 15)
        y_label: Label for y-axis
        ref_curve: Reference curve data (length 15)
        field_curve: Field measurement curve data (length 15)
        ref_label: Label for reference curve
        field_label: Label for field curve

    Returns:
        str: Path to the saved plot image
    """
    # Verify input shapes
    print(f"Plot input shapes:")
    print(f"frequencies: {len(frequencies)}")
    print(f"ref_curve: {ref_curve}")
    print(f"field_curve: {field_curve}")
    
    # Create figure
    plt.figure(figsize=(10, 6))
    
    # Plot curves
    plt.plot(frequencies, ref_curve, label=ref_label, color='red')
    plt.plot(frequencies, field_curve, label=field_label, color='black', 
             marker='s', linestyle='--')
    
    # Configure axes and labels
    plt.xlabel('Frequency')
    plt.ylabel(y_label)
    plt.grid(True)
    plt.tick_params(axis='x', rotation=45)
    plt.xticks(frequencies)
    plt.xscale('log')
    plt.gca().get_xaxis().set_major_formatter(plt.ScalarFormatter())
    plt.gca().xaxis.set_major_locator(ticker.FixedLocator(frequencies))
    plt.legend()
    
    # Create temp file with .png extension
    temp_dir = tempfile.gettempdir()
    temp_path = os.path.join(temp_dir, 'plot.png')
    
    plt.savefig(temp_path, format='png', bbox_inches='tight', dpi=300)
    plt.close()
    
    return temp_path

# Additional utility functions
def sanitize_filepath(filepath: str) -> str:
    filepath = filepath.replace('T:', '//DLA-04/Shared/')
    filepath = filepath.replace('\\', '/')
    return filepath

@dataclass
class SLMData:
    """Class to handle Sound Level Meter data with both raw and processed formats"""
    raw_data: pd.DataFrame
    measurement_type: str  # '831_Data' or 'RT_Data'
    
    def __post_init__(self):
        """Process raw data after initialization"""
        if self.measurement_type == 'RT_Data':
            self._validate_rt_data()
            self._process_rt_data()
        else:
            self._process_oba_data()
            self._validate_processed_data()
    
    def _validate_rt_data(self):
        """Validate RT data format"""
        if self.raw_data.empty:
            raise ValueError("Empty DataFrame provided")
    
    def _validate_processed_data(self):
        """Validate processed data has required columns"""
        expected_cols = ['Frequency (Hz)', 'Overall 1/3 Spectra', 'Max 1/3 Spectra', 'Min 1/3 Spectra']
        if not all(col in self.processed_data.columns for col in expected_cols):
            raise ValueError(f"Processing failed to create required columns: {expected_cols}")
    
    def _process_rt_data(self):
        """Process RT data"""
        try:
            print("\nProcessing RT data...")
            # First try the hardcoded approach that works with the original dataset
            try:
                print("Attempting hardcoded approach...")
                self.rt_thirty = np.array(self.raw_data['Unnamed: 10'][24:41]/1000, dtype=np.float64).round(3)
                self.frequency_bands = np.array([100, 125, 160, 200, 250, 315, 400, 500, 630, 
                                              800, 1000, 1250, 1600, 2000, 2500, 3150, 4000])
                self.overall_levels = self.rt_thirty
                print("Hardcoded approach successful")
                return
            except Exception as e:
                print(f"Hardcoded approach failed, trying dynamic header detection: {str(e)}")
                
            # If hardcoded approach fails, try dynamic header detection
            header_row_idx = None
            for idx, row in self.raw_data.iterrows():
                # Convert row values to strings and clean them
                row_values = [str(val).strip() if pd.notna(val) else '' for val in row.values]
                if 'T30 (ms)' in row_values and 'Frequency (Hz)' in row_values:
                    header_row_idx = int(idx)
                    freq_col_idx = row_values.index('Frequency (Hz)')
                    rt_col_idx = row_values.index('T30 (ms)')
                    print(f"Found headers at row {header_row_idx}")
                    print(f"Frequency column index: {freq_col_idx}")
                    print(f"RT column index: {rt_col_idx}")
                    break
                    
            if header_row_idx is None:
                raise ValueError("Could not find RT data headers")
                
            # Extract data starting from the row after headers
            data_start_idx = header_row_idx + 1
            freq_data = self.raw_data.iloc[data_start_idx:, freq_col_idx]
            rt_values = self.raw_data.iloc[data_start_idx:, rt_col_idx]
            
            print("Raw data extracted:")
            print(f"Frequency data: {freq_data.head()}")
            print(f"RT values: {rt_values.head()}")
            
            # Clean and convert frequency values
            freq_data = freq_data.astype(str).str.replace('Hz', '').str.strip()
            freq_data = pd.to_numeric(freq_data, errors='coerce')
            rt_values = pd.to_numeric(rt_values, errors='coerce')/1000  # Convert from ms to seconds
            
            # Remove any rows where conversion failed
            valid_mask = ~(freq_data.isna() | rt_values.isna())
            self.frequency_bands = freq_data[valid_mask].values
            self.rt_thirty = rt_values[valid_mask].values
            self.overall_levels = self.rt_thirty
            
            print("\nProcessed RT data:")
            print(f"Number of valid frequency bands: {len(self.frequency_bands)}")
            print(f"Frequency range: {self.frequency_bands.min():.1f} - {self.frequency_bands.max():.1f} Hz")
            print(f"RT values range: {self.rt_thirty.min():.3f} - {self.rt_thirty.max():.3f} seconds")
            
        except Exception as e:
            print("\nError details:")
            print(f"Raw data shape: {self.raw_data.shape}")
            print(f"Raw data columns: {self.raw_data.columns.tolist()}")
            print("First few rows of raw data:")
            print(self.raw_data.head())
            raise ValueError(f"Error processing RT data: {str(e)}")
    
    def _process_oba_data(self):
        """Process OBA sheet data into structured format"""
        try:
            # Check if raw_data is empty
            if self.raw_data.empty:
                print("Warning: Empty DataFrame provided to _process_oba_data")
                # Initialize with empty data but valid structure
                self.processed_data = pd.DataFrame({
                    'Frequency (Hz)': [],
                    'Overall 1/3 Spectra': [],
                    'Max 1/3 Spectra': [],
                    'Min 1/3 Spectra': []
                })
                self.frequency_bands = np.array([])
                self.overall_levels = np.array([])
                return
                
            # First, find the 1/3 Octave section
            third_oct_idx = None
            for idx, row in self.raw_data.iterrows():
                if isinstance(row.iloc[0], str) and '1/3 Octave' in row.iloc[0]:
                    third_oct_idx = idx
                    break
                
            if third_oct_idx is None:
                # If no 1/3 Octave section, try using 1/1 Octave section
                if len(self.raw_data) >= 2:
                    freq_row = self.raw_data.iloc[0]  # First row contains frequencies
                    data_row = self.raw_data.iloc[1]  # Second row contains Overall spectra
                else:
                    print("Warning: DataFrame has insufficient rows for processing")
                    # Initialize with empty data but valid structure
                    self.processed_data = pd.DataFrame({
                        'Frequency (Hz)': [],
                        'Overall 1/3 Spectra': [],
                        'Max 1/3 Spectra': [],
                        'Min 1/3 Spectra': []
                    })
                    self.frequency_bands = np.array([])
                    self.overall_levels = np.array([])
                    return
            else:
                # Use 1/3 Octave section
                if len(self.raw_data) > third_oct_idx + 2:
                    freq_row = self.raw_data.iloc[third_oct_idx + 1]
                    data_row = self.raw_data.iloc[third_oct_idx + 2]
                else:
                    print("Warning: DataFrame has insufficient rows after 1/3 Octave section")
                    # Initialize with empty data but valid structure
                    self.processed_data = pd.DataFrame({
                        'Frequency (Hz)': [],
                        'Overall 1/3 Spectra': [],
                        'Max 1/3 Spectra': [],
                        'Min 1/3 Spectra': []
                    })
                    self.frequency_bands = np.array([])
                    self.overall_levels = np.array([])
                    return

            # Extract frequencies and values
            frequencies = []
            values = []
            
            # Iterate through columns to collect data
            for col in self.raw_data.columns:
                try:
                    freq = pd.to_numeric(freq_row[col], errors='coerce')
                    val = pd.to_numeric(data_row[col], errors='coerce')
                    
                    if pd.notna(freq) and pd.notna(val):
                        frequencies.append(freq)
                        values.append(val)
                except (IndexError, KeyError) as e:
                    print(f"Warning: Error processing column {col}: {str(e)}")
                    continue

            # Create processed DataFrame
            self.processed_data = pd.DataFrame({
                'Frequency (Hz)': frequencies,
                'Overall 1/3 Spectra': values,
                'Max 1/3 Spectra': values,  # Using same values for now
                'Min 1/3 Spectra': values   # Using same values for now
            })
            
            # Store frequency bands and levels
            self.frequency_bands = np.array(frequencies)
            self.overall_levels = np.array(values)
            
            print("\nProcessed OBA Data:")
            print(f"Found {len(values)} frequency bands")
            if len(values) > 0:
                print(f"Frequency range: {min(frequencies)} - {max(frequencies)} Hz")
            print("Data shape:", self.processed_data.shape)
            print("\nFirst few rows:")
            print(self.processed_data.head())
            
        except Exception as e:
            print("\nDebug information:")
            print("Raw data shape:", self.raw_data.shape)
            print("First few rows of raw data:")
            print(self.raw_data.head())
            print("\nDetailed error:")
            print(f"Error type: {type(e).__name__}")
            print(f"Error message: {str(e)}")
            # Instead of raising an error, initialize with empty data
            self.processed_data = pd.DataFrame({
                'Frequency (Hz)': [],
                'Overall 1/3 Spectra': [],
                'Max 1/3 Spectra': [],
                'Min 1/3 Spectra': []
            })
            self.frequency_bands = np.array([])
            self.overall_levels = np.array([])
            print("Initialized with empty data due to error")
    
    def get_levels(self, freq_range: tuple = None) -> np.ndarray:
        """Get overall levels, optionally filtered by frequency range"""
        if freq_range is None:
            return self.overall_levels
            
        mask = (self.frequency_bands >= freq_range[0]) & (self.frequency_bands <= freq_range[1])
        return self.overall_levels[mask]
    
    def get_rt_thirty(self, freq_range: tuple = None) -> np.ndarray:
        """Get RT30 values if this is RT data"""
        if self.measurement_type != 'RT_Data':
            raise ValueError("RT30 values only available for RT_Data type")
            
        if freq_range is None:
            return self.rt_thirty
            
        mask = (self.frequency_bands >= freq_range[0]) & (self.frequency_bands <= freq_range[1])
        return self.rt_thirty[mask]
    
    def plot_data(self) -> None:
        """Plot the measurement data"""
        plt.figure(figsize=(10, 6))
        plt.semilogx(self.frequency_bands, self.overall_levels, 'b-o')
        plt.grid(True)
        plt.xlabel('Frequency (Hz)')
        plt.ylabel('Level (dB)' if self.measurement_type == '831_Data' else 'RT30 (s)')
        plt.title(f'SLM {self.measurement_type} Measurements')

    def process_rt_data(self, df):
        """Process RT data and store both RT values and frequencies"""
        try:
            # Find header row
            header_row_idx = None
            for idx, row in df.iterrows():
                if 'T30 (ms)' in row.values and 'Frequency (Hz)' in row.values:
                    header_row_idx = int(idx)
                    freq_col_idx = int(row.tolist().index('Frequency (Hz)'))
                    rt_col_idx = int(row.tolist().index('T30 (ms)'))
                    break

            if header_row_idx is None:
                raise ValueError("Could not find RT data headers")

            # Extract and clean data
            data_start_idx = header_row_idx + 1
            freq_data = df.iloc[data_start_idx:, freq_col_idx]
            rt_values = df.iloc[data_start_idx:, rt_col_idx]

            # Clean and convert
            freq_data = freq_data.astype(str).str.replace('Hz', '').str.strip()
            freq_data = pd.to_numeric(freq_data, errors='coerce')
            rt_values = pd.to_numeric(rt_values, errors='coerce')

            # Store both values
            self.frequencies = freq_data.values
            self.rt_thirty = rt_values.values

            # Store in raw_data as well
            self.raw_data = pd.DataFrame({
                'Frequency (Hz)': freq_data,
                'RT60': rt_values
            })

        except Exception as e:
            raise ValueError(f"Error processing RT data: {str(e)}")



