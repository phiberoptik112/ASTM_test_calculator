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
import matplotlib.pyplot as plt
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
from typing import List


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
### maybe change to something thats a better troubleshooting effort later ###
##  ask GPT later about how best to do this, maybe a config file or something? 
def format_SLMdata(srs_data):
    ## verify that srs_data iloc[7] is correct- will have a label as 1/3 octave
    srs_thirdoct = srs_data.iloc[7] # hardcoded to SLM export data format
    srs_thirdoct = srs_thirdoct[13:31] # select only the frequency bands of interest
    # verify step to check for SPL info? 
    return srs_thirdoct

def extract_sound_levels(slm_data: pd.DataFrame, 
                        freq_bands: list = [125, 160, 200, 250, 315, 400, 500, 630, 800, 
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
        freq_labels = pd.to_numeric(slm_data.iloc[spectrum_row_idx - 1, 13:31], errors='coerce')
        
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
    if isinstance(average_pos, pd.DataFrame):
        average_pos = average_pos.values
    onethird_rec_Total = []
    for i in range(len(average_pos)):
        freqbin = average_pos[i]
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



### NEW REWORK OF NR CALC ### # functional as of 7/25/24 ### 
def calc_nr_new(srs_overalloct: pd.Series, rec_overalloct: pd.Series, 
                bkgrnd_overalloct: pd.Series, rt_thirty: pd.Series, 
                receive_roomvol: float, 
                test_type: TestType) -> Tuple[pd.Series, pd.Series, pd.Series, pd.Series]:
    # Implement the NR calculation logic
    NIC_vollimit = 150  # cu. ft.
    if receive_roomvol > NIC_vollimit:
        print('Using NIC calc, room volume too large')
    # sabines = 0.049*(recieve_roomvol/rt_thirty)  # this produces accurate sabines values
    recieve_corr = list()
    rec_overalloct = pd.to_numeric(rec_overalloct)
    bkgrnd_overalloct = pd.to_numeric(bkgrnd_overalloct)

    recieve_vsBkgrnd = rec_overalloct - bkgrnd_overalloct
    # print('recieve: ',rec_overalloct)
    # print('background: ',bkgrnd_overalloct)
    recieve_vsBkgrnd = np.round(recieve_vsBkgrnd,1)
    
    print('rec vs background:',recieve_vsBkgrnd)
    if test_type == 'AIIC':
        for i, val in enumerate(recieve_vsBkgrnd):
            if val < 5:
                recieve_corr.append(rec_overalloct[i]-2)
            else:
                recieve_corr.append(10*np.log10(10**(rec_overalloct[i]/10)-10**(bkgrnd_overalloct[i]/10)))   
    elif test_type == 'ASTC':
        for i, val in enumerate(recieve_vsBkgrnd):
            # print('val:', val)
            # print('count: ', i)
            if val < 5:
                recieve_corr.append(rec_overalloct.iloc[i]-2)
            elif val < 10:
                recieve_corr.append(10*np.log10(10**(rec_overalloct[i]/10)-10**(bkgrnd_overalloct[i]/10)))
            else:
                recieve_corr.append(rec_overalloct[i])
        # print('-=-=-=-=-')
        # print('recieve_corr: ',recieve_corr)
    recieve_corr = np.round(recieve_corr,1)
    print('corrected recieve ISPL: ',recieve_corr)
    NR_val = srs_overalloct - recieve_corr

    # Normalized_recieve = recieve_corr / srs_overalloct
    sabines = pd.to_numeric(sabines, errors='coerce')
    sabines = np.round(sabines)
    if isinstance(sabines, pd.DataFrame):
        sabines = sabines.values

    Normalized_recieve = list()
    Normalized_recieve = recieve_corr-10*(np.log10(108/sabines))
    Normalized_recieve = np.round(Normalized_recieve)
    print('Normalized_recieve: ',Normalized_recieve)
    return NR_val, sabines,recieve_corr, Normalized_recieve

### this code revised 7/24/24 - functional and produces accurate ATL values

def calc_atl_val(srs_overalloct: pd.Series, rec_overalloct: pd.Series, 
                 bkgrnd_overalloct: pd.Series, rt_thirty: pd.Series, 
                 partition_area: float, receive_roomvol: float) -> pd.Series:
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
    # print('rt_thirty: ',rt_thirty)
    
    sabines = 0.049*receive_roomvol/rt_thirty  # this produces accurate sabines values

    if isinstance(bkgrnd_overalloct, pd.DataFrame):
        bkgrnd_overalloct = bkgrnd_overalloct.values
    # print('bkgrnd_overalloct: ',bkgrnd_overalloct)
    # sabines = np.int32(sabines) ## something not right with this calc
    sabines = np.round(sabines)
    # print('sabines: ',sabines)
    
    recieve_corr = list()
    # print('recieve: ',rec_overalloct)
    recieve_vsBkgrnd = rec_overalloct - bkgrnd_overalloct

    if isinstance(recieve_vsBkgrnd, pd.DataFrame):
        recieve_vsBkgrnd = recieve_vsBkgrnd.values

    if isinstance(rec_overalloct, pd.DataFrame):
        rec_overalloct = rec_overalloct.values
    # print('recieve vs  background: ',recieve_vsBkgrnd)
    # print('recieve roomvol: ',recieve_roomvol)
    #### something wrong with this loop #### 
    for i, val in enumerate(recieve_vsBkgrnd):
        if val < 5:
            # print('recieve vs background: ',val)
            recieve_corr.append(rec_overalloct[i]-2)
            # print('less than 5, appending: ',recieve_corr[i])
        elif val < 10:
            # print('recieve vs background: ',val)
            recieve_corr.append(10*np.log10((10**(rec_overalloct[i]/10))-(10**(bkgrnd_overalloct[i]/10))))
            # print('less than 10, appending: ',recieve_corr[i])
        else:
            # print('recieve vs background: ',val)
            recieve_corr.append(rec_overalloct[i])
            # print('greater than 10, appending: ',recieve_corr[i])
            
    
    # print('recieve correction: ',recieve_corr)
    if isinstance(srs_overalloct, pd.DataFrame):
        recieve_corr = recieve_corr.values
    if isinstance(srs_overalloct, pd.DataFrame):
        srs_overalloct = srs_overalloct.values
    if isinstance(sabines, pd.DataFrame):
        sabines = sabines.values
    # print('srs overalloct: ',srs_overalloct)
    # print('recieve correction: ',recieve_corr)
    # print('sabines: ',sabines)
    ATL_val = []
    for i, val in enumerate(srs_overalloct):
        ATL_val.append(srs_overalloct[i]-recieve_corr[i]+10*(np.log10(partition_area/sabines.iloc[i])))
    # ATL_val = srs_overalloct - recieve_corr+10*(np.log(parition_area/sabines)) 
    ATL_val = np.round(ATL_val,1)
    print('ATL val: ',ATL_val)
    return ATL_val
def calc_AIIC_val_claude(Normalized_recieve_IIC, verbose=True):
    pos_diffs = list()
    diff_negative_min = 0
    AIIC_start = 94
    AIIC_contour_val = 16
    IIC_contour = list()
    AIIC_curve = list()
    new_sum = 0
    diff_negative_max = 0
    IIC_curve = [2,2,2,2,2,2,1,0,-1,-2,-3,-6,-9,-12,-15,-18]
    max_iterations = 100  # Maximum number of iterations to prevent infinite loop
    iteration_count = 0
    # initial application of the IIC curve to the first AIIC start value 
    for vals in IIC_curve:
        IIC_contour.append(vals+AIIC_start)
    Normalized_recieve_IIC = pd.to_numeric(Normalized_recieve_IIC, errors='coerce')
    Normalized_recieve_IIC = np.array(Normalized_recieve_IIC)
    Normalized_recieve_IIC = np.round(Normalized_recieve_IIC,1)
    Normalized_recieve_IIC = Normalized_recieve_IIC[1:17]
    print('Normalized recieve ANISPL: ', Normalized_recieve_IIC)
    # Contour_curve_result = IIC_contour - Normalized_recieve_IIC
    Contour_curve_result =  Normalized_recieve_IIC - IIC_contour
    Contour_curve_result = np.round(Contour_curve_result,1)
    
    # print('Contour curve: ', Contour_curve_result)
    while (diff_negative_max < 8 and new_sum < 32 and iteration_count < max_iterations):
        if verbose:
            print(f"Iteration {iteration_count}:")
            print(f"  AIIC_contour_val: {AIIC_contour_val}")
            print(f"  diff_negative_max: {diff_negative_max}")
            print(f"  new_sum: {new_sum}")
        print('Inside loop, current AIIC contour: ', AIIC_contour_val)
        # print('Contour curve (IIC curve minus ANISPL): ', Contour_curve_result)
        
        diff_negative = Normalized_recieve_IIC - IIC_contour
        # print('diff negative: ', diff_negative)
        diff_negative_max = np.max(diff_negative)
        diff_negative = pd.to_numeric(diff_negative, errors='coerce')
        diff_negative = np.array(diff_negative)
        # print('Max, single diff: ', diff_negative_max)
        pos_diffs = [np.round(val,1) if val > 0 else 0 for val in diff_negative]
        # print('positive diffs: ', pos_diffs)
        new_sum = np.sum(pos_diffs)
        # ipdb.set_trace()   ### debugging 
        
        # print('Sum Positive diffs: ', new_sum)
        # print('Evaluating sums and differences vs 32, 8: ', new_sum, diff_negative_max)
        
        if new_sum > 32 or diff_negative_max > 8:
            print('Difference condition met! AIIC value: ', AIIC_contour_val)
            print('AIIC result curve: ', Contour_curve_result)
            return AIIC_contour_val, Contour_curve_result
        else:
            print('difference condition not met, subtracting 1 from AIIC start and recalculating the IIC contour')
            AIIC_start -= 1
            print('new AIIC start: ', AIIC_start)
            AIIC_contour_val += 1
            print('AIIC contour value: ', AIIC_contour_val)
            IIC_contour = [vals + AIIC_start for vals in IIC_curve]
            print('IIC contour: ', IIC_contour)
            
            Contour_curve_result =  Normalized_recieve_IIC - IIC_contour
            print('Contour curve result: ', Contour_curve_result)
            
            iteration_count += 1

    if iteration_count == max_iterations:
        print(f"Maximum iterations ({max_iterations}) reached without meeting conditions.")
    else:
        print("Loop completed without meeting conditions. Returning last calculated values.")
    print(f"Loop exited. Final values:")
    print(f"  diff_negative_max: {diff_negative_max}")
    print(f"  new_sum: {new_sum}")
    print(f"  iterations: {iteration_count}")
    print('Contour curve (IIC curve minus ANISPL): ', Contour_curve_result)
    print('IIC contour value: ', AIIC_contour_val)


    return AIIC_contour_val, Contour_curve_result

def calc_aiic_val(Normalized_recieve_IIC: pd.Series) -> Tuple[float, pd.Series]:
    pos_diffs = list()
    diff_negative_min = 0
    AIIC_start = 94
    AIIC_contour_val = 16
    IIC_contour = list()
    AIIC_curve= list()
    new_sum = 0
    diff_negative_max = 0
    IIC_curve = [2,2,2,2,2,2,1,0,-1,-2,-3,-6,-9,-12,-15,-18]

    # initial application of the IIC curve to the first AIIC start value 
    for vals in IIC_curve:
        IIC_contour.append(vals+AIIC_start)
    Normalized_recieve_IIC = np.round(Normalized_recieve_IIC,1)
    Normalized_recieve_IIC = Normalized_recieve_IIC[1:17]
    Contour_curve_result = IIC_contour - Normalized_recieve_IIC
    Contour_curve_result = np.round(Contour_curve_result)
    # print('Normalized recieve ANISPL: ', Normalized_recieve_IIC)
    
    print('Contour curve: ',Contour_curve_result)

    while (diff_negative_max < 8 and new_sum < 32):
        print('Inside loop, current AIIC contour: ', AIIC_contour_val)
        print('Contour curve (IIC curve minus ANISPL): ',Contour_curve_result)
        
        diff_negative =  Normalized_recieve_IIC-IIC_contour
        print('diff negative: ', diff_negative)

        diff_negative_max =  np.max(diff_negative)
        diff_negative = pd.to_numeric(diff_negative, errors='coerce')
        diff_negative = np.array(diff_negative)

        print('Max, single diff: ', diff_negative_max)
        for val in diff_negative:
            if val > 0:
                pos_diffs.append(np.round(val))
            else:
                pos_diffs.append(0)
        print('positive diffs: ',pos_diffs)
        new_sum = np.sum(pos_diffs)
        print('Sum Positive diffs: ', new_sum)
        print('Evaluating sums and differences vs 32, 8: ', new_sum, diff_negative_max)
        # need to debug this return statement placement 
        if new_sum > 32 or diff_negative_max > 8:
            print('Difference condition met! AIIC value: ', AIIC_contour_val) # 
            print('AIIC result curve: ', Contour_curve_result)
            return AIIC_contour_val, Contour_curve_result
        # condition not met, resetting arrays
        pos_diffs = []
        IIC_contour = []
        print('difference condition not met, subtracting 1 from AIIC start and recalculating the IIC contour')

        AIIC_start -= 1
        print('new AIIC start: ', AIIC_start)
        AIIC_contour_val += 1
        # print('AIIC start: ', AIIC_start)
        print('AIIC contour value: ', AIIC_contour_val)
        for vals in IIC_curve:
            print('vals: ', vals)
            IIC_contour.append(vals+AIIC_start)
        
        Contour_curve_result = IIC_contour - Normalized_recieve_IIC
        print('Contour curve iterate: ',Contour_curve_result)
    return aiic_val, contour_curve

def calc_astc_val(atl_val: pd.Series) -> float:
    pos_diffs = list()
    diff_negative=0
    diff_positive=0 
    ASTC_start = 16
    New_curve =list()
    new_sum = 0
    ## Since ATL values only go from 125 to 4k, remove the end values from the curve
    ATL_val_STC = atl_val[1:17]
    STCCurve = [-16, -13, -10, -7, -4, -1, 0, 1, 2, 3, 4, 4, 4, 4, 4, 4]
    while (diff_negative <= 8 and new_sum <= 32):
        # print('starting loop')
        print('ASTC fit test value: ', ASTC_start)
        for vals in STCCurve:
            New_curve.append(vals+ASTC_start)
     
        ASTC_curve = New_curve - ATL_val_STC
        
        ASTC_curve = np.round(ASTC_curve)
        print('ASTC curve: ',ASTC_curve)
        diff_negative = np.max(ASTC_curve)

        print('Max, single diff: ', diff_negative)

        for val in ASTC_curve:
            if val > 0:
                pos_diffs.append(np.round(val))
            else:
                pos_diffs.append(0)
        # print(pos_diffs)
        new_sum = np.sum(pos_diffs)
        print('Sum Positive diffs: ', new_sum)
        
        if new_sum > 32 or diff_negative > 8:
            print('Curve too high! ASTC fit: ', ASTC_start-1) 
            return ASTC_start-1
        pos_diffs = []
        New_curve = []
        ASTC_start = ASTC_start + 1
        
        if ASTC_start >80: break

def plot_curves(frequencies: List[float], y_label: str, ref_curve: pd.Series, 
                field_curve: pd.Series, ref_label: str, field_label: str) -> ImageReader:
    """Creates a plot and returns it in a format compatible with ReportLab.
    
    Args:
        frequencies: List of frequency values for x-axis
        y_label: Label for y-axis
        ref_curve: Reference curve data
        field_curve: Field measurement curve data
        ref_label: Label for reference curve
        field_label: Label for field curve
        
    Returns:
        ImageReader: ReportLab-compatible image object
    """
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
    
    # Save plot to bytes buffer instead of file
    img_buffer = BytesIO()
    plt.savefig(img_buffer, format='png', bbox_inches='tight', dpi=300)
    img_buffer.seek(0)
    
    # Clear the current figure to free memory
    plt.close()
    
    # Create ReportLab ImageReader object
    return ImageReader(img_buffer)

def process_single_test(test_plan_entry: pd.Series, slm_data_paths: Dict[str, Path], 
                        output_folder: Path) -> Path:
    room_props = RoomProperties(**test_plan_entry.to_dict())
    
    srs_data = load_slm_data(slm_data_paths['source'], 'OBA')
    rec_data = load_slm_data(slm_data_paths['receive'], 'OBA')
    bkgrnd_data = load_slm_data(slm_data_paths['background'], 'OBA')
    rt_data = load_slm_data(slm_data_paths['rt'], 'Summary')
    
    test_data = TestData(srs_data, rec_data, bkgrnd_data, rt_data, room_props)
    
    test_type = TestType(test_plan_entry['TestType'])
    
    report_path = create_report(test_data, output_folder, test_type)
    
    return report_path

# Additional utility functions
def sanitize_filepath(filepath: str) -> str:
    filepath = filepath.replace('T:', '//DLA-04/Shared/')
    filepath = filepath.replace('\\', '/')
    return filepath



