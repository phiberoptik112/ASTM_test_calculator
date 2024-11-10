import pandas as pd
import numpy as np
import logging
# import yaml
import os 
from os import listdir, walk
from dataclasses import dataclass
from typing import List, Tuple, Dict
from enum import Enum
from pathlib import Path
import matplotlib.pyplot as plt
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import BaseDocTemplate, PageTemplate, Frame, Paragraph, Table, TableStyle, Spacer, PageBreak, KeepInFrame, Image
from reportlab.lib.units import inch

## having gone through the PDF edit, this room properties will need to be revised 
# im not even sure if moving to a class variable from a dataframe is best, 
# i think it will work for now.  
    # room_properties = pd.DataFrame(
    #     {
    #         "Site Name": curr_test['Site_Name'],
    #         "Client Name": curr_test['Client_Name'],
    #         "Source Room Name": curr_test['Source_Room'],
    #         "Recieve Room Name": curr_test['Receiving_Room'],
    #         "Testdate": curr_test['Test_Date'],
    #         "ReportDate": curr_test['Report_Date'],
    #         "Project Name": curr_test['Project_Name'],
    #         "Test number": curr_test['Test_Label'],
    #         "Source Vol" : curr_test['source_room_vol'],
    #         "Recieve Vol": curr_test['receive_room_vol'],
    #         "Partition area": curr_test['partition_area'],
    #         "Partition dim.": curr_test['partition_dim'],
    #         "Source room Finish" : curr_test['source_room_finish'],
    #         "Recieve room Finish": curr_test['receive_room_finish'],
    #         "Srs Floor Descrip.": curr_test['srs_floor'],
    #         "Srs Ceiling Descrip.": curr_test['srs_ceiling'],
    #         "Srs Walls Descrip.": curr_test['srs_Walls'],
    #         "Rec Floor Descrip.": curr_test['rec_floor'],
    #         "Rec Ceiling Descrip.": curr_test['rec_ceiling'],
    #         "Rec Walls Descrip.": curr_test['rec_Wall'],          
    #         "Tested Assembly": curr_test['tested_assembly'],
    #         "Expected Performance": curr_test['expected_performance'],
    #         "Annex 2 used?": curr_test['Annex_2_used?'],
    #         "Test assem. type": curr_test['Test_assembly_Type'],
    #         "AIIC_test": curr_test['AIIC_test'],
    #         "NIC_test": curr_test['NIC_test'],
    #         "ASTC_test": curr_test['ASTC_test'],
    #         "DTC_test": curr_test['DTC_test'],
    #         "NIC reporting Note": NICreporting_Note
    #     },
    #     index=[0]
    # )
        ### 
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
    receive_room_finish: str
    tested_assembly: str
    expected_performance: str
    annex_2_used: bool
    test_assembly_type: str

# each test data class has all the differnet test types in it
@dataclass
class TestData:
    single_ASTCtest_data: pd.DataFrame
    single_AIICtest_data: pd.DataFrame
    single_NICtest_data: pd.DataFrame
    single_DTCtest_data: pd.DataFrame
    # room_properties: RoomProperties  ## removing this to ReportData class

class TestType(Enum):
    AIIC = "AIIC"
    ASTC = "ASTC"
    NIC = "NIC"
    DTC = "DTC"
@dataclass
class ReportData:
    room_properties: RoomProperties
    test_data: TestData
    test_type: TestType
#####
# class TestData:
#     def __init__(self, room_properties: RoomProperties, ...):
#         self.room_properties = room_properties
#         # other initializations...

# class ReportData:
#     def __init__(self, test_data: TestData, ...):
#         self.test_data = test_data  # Contains room_properties via test_data
#         # other initializations...
        
#     def get_room_properties(self) -> RoomProperties:
#         return self.test_data.room_properties
### SLM import - hardcoded - gives me the willies, but it works for now.
### maybe change to something thats a better troubleshooting effort later ###
##  ask GPT later about how best to do this, maybe a config file or something? 
def format_SLMdata(srs_data):
    ## verify that srs_data iloc[7] is correct- will have a label as 1/3 octave
    srs_thirdoct = srs_data.iloc[7] # hardcoded to SLM export data format
    srs_thirdoct = srs_thirdoct[13:31] # select only the frequency bands of interest
    # verify step to check for SPL info? 
    return srs_thirdoct

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

def load_test_plan(curr_test: pd.DataFrame) -> pd.DataFrame:
    # Load the test plan data from the Excel file

    # do i need to run the format_slm_data function here?

    if int(curr_test['source room vol']) >= 5300 or int(curr_test['receive room vol']) >= 5300:
        NICreporting_Note = 'The receiver and/or source room had a volume exceeding 150 m3 (5,300 cu. ft.), and the absorption of the receiver and/or source room was greater than the maximum allowed per E336-16, Paragraph 9.4.1.2.'
    elif int(curr_test['source room vol']) <= 833 or int(curr_test['receive room vol']) <= 833:
        NICreporting_Note = 'The receiver and/or source room has a volume less than the minimum volume requirement of 25 m3 (883 cu. ft.).'
    else:
        NICreporting_Note = '---'
    
    # load a test properties dataframe
    # rewrite to the TestType enum
    test_types = {
        TestType.AIIC: bool(curr_test['AIIC_test'].iloc[0]),  # Convert to boolean
        TestType.NIC: bool(curr_test['NIC_test'].iloc[0]),
        TestType.ASTC: bool(curr_test['ASTC_test'].iloc[0]),
        TestType.DTC: bool(curr_test['DTC_test'].iloc[0])
    }
    # test_properties = pd.DataFrame(
    #     {
    #         "AIIC_test": curr_test['AIIC_test'],
    #         "NIC_test": curr_test['NIC_test'],
    #         "ASTC_test": curr_test['ASTC_test'],
    #         "DTC_test": curr_test['DTC_test']
    #     },
    #     index=[0]
    # )
    room_properties = RoomProperties(
        site_name=curr_test['Site_Name'],
        client_name=curr_test['Client_Name'],
        source_room=curr_test['Source_Room'],
        receive_room=curr_test['Receiving_Room'],
        test_date=curr_test['Test_Date'],
        report_date=curr_test['Report_Date'],
        project_name=curr_test['Project_Name'],
        test_label=curr_test['Test_Label'],
        source_vol=curr_test['source room vol'],
        receive_vol=curr_test['receive room vol'],
        partition_area=curr_test['partition area'],
        partition_dim=curr_test['partition dim.'],
        source_room_finish=curr_test['source room finish'],
        receive_room_finish=curr_test['receive room finish'],
        srs_floor=curr_test['srs floor descrip.'],
        srs_ceiling=curr_test['srs ceiling descrip.'],
        srs_walls=curr_test['srs walls descrip.'],
        rec_floor=curr_test['rec floor descrip.'],
        rec_ceiling=curr_test['rec ceiling descrip.'],
        rec_wall=curr_test['rec walls descrip.'],   
    )

    if curr_test['AIIC_test'] == 1:
        print("AIIC testing enabled, copying data...")
### IIC variables  #### When extending to IIC data
        find_source = curr_test['Source']
        find_rec = curr_test['Recieve '] #trailing whitespace? be sure to verify this is consistent in the excel file
        find_BNL = curr_test['BNL']
        find_RT = curr_test['RT']
        find_posOne = curr_test['Position1']
        find_posTwo = curr_test['Position2']
        find_posThree = curr_test['Position3']
        find_posFour = curr_test['Position4']
        find_poscarpet = curr_test['Carpet']
        find_Tapsrs = curr_test['SourceTap']
        srs_data = RAW_SLM_datapull(self,find_source,'-831_Data.')
        recive_data = RAW_SLM_datapull(self,find_rec,'-831_Data.')
        bkgrnd_data = RAW_SLM_datapull(self,find_BNL,'-831_Data.')

        AIIC_pos1 = RAW_SLM_datapull(self,find_posOne,'-831_Data.')
        AIIC_pos2 = RAW_SLM_datapull(self,find_posTwo,'-831_Data.')
        AIIC_pos3 = RAW_SLM_datapull(self,find_posThree,'-831_Data.')
        AIIC_pos4 = RAW_SLM_datapull(self,find_posFour,'-831_Data.')
        AIIC_carpet = RAW_SLM_datapull(self,find_poscarpet,'-831_Data.')
        AIIC_source = RAW_SLM_datapull(self,find_Tapsrs,'-831_Data.')

        rt = RAW_SLM_datapull(self,find_RT,'-RT_Data.')
        rt_thirty = rt['Unnamed: 10'][24:42]/1000 ## need to validate that this works
        if isinstance(rt, pd.DataFrame):
            rt_thirty = rt.values
        rt_thirty = pd.to_numeric(rt_thirty, errors='coerce')
        rt_thirty = np.round(rt_thirty,3)
        ##### May need to put in the average of all the positions, log averaged together #####
        # not sure where ive got it calculated in, it may be in my IIC function itself, 
        # but it might be best to just average it here and then pass through the total log avg. 
        average_pos = []
        for i in range(1, 5):
            pos_input = f'AIIC_pos{i}'
            pos_data = format_SLMdata(single_AIICtest_data[pos_input]) # need to change to get the raw data 
            average_pos.append(pos_data)

        average_pos = pd.concat(average_pos, axis=1)
        onethird_rec_Total = calculate_onethird_Logavg(average_pos)
        print('tap total:', onethird_rec_Total)

        single_AIICtest_data = {
            'srs_data': pd.DataFrame(srs_data),
            'recive_data': pd.DataFrame(recive_data),
            'bkgrnd_data': pd.DataFrame(bkgrnd_data),
            'rt': pd.DataFrame(rt),
            'AIIC_pos1': pd.DataFrame(AIIC_pos1),
            'AIIC_pos2': pd.DataFrame(AIIC_pos2),
            'AIIC_pos3': pd.DataFrame(AIIC_pos3),
            'AIIC_pos4': pd.DataFrame(AIIC_pos4),
            'AIIC_source': pd.DataFrame(AIIC_source),
            'AIIC_carpet': pd.DataFrame(AIIC_carpet),
        }

    else:
        single_AIICtest_data = None
        # return single_AIICtest_data
    
    if curr_test['ASTC_test'] == 1:
        srs_data = RAW_SLM_datapull(self,curr_test['Source'],'-831_Data.')
        recive_data = RAW_SLM_datapull(self, curr_test['Recieve '],'-831_Data.')
        bkgrnd_data = RAW_SLM_datapull(self,curr_test['BNL'],'-831_Data.')
        rt = RAW_SLM_datapull(self,curr_test['RT'],'-RT_Data.')
        ## formating the RT data further 
        rt_thirty = rt['Summary']['Unnamed: 10'][24:42]/1000
        if isinstance(rt, pd.DataFrame):
            rt_thirty = rt.values
        rt_thirty = pd.to_numeric(rt_thirty, errors='coerce')
        rt_thirty = np.round(rt_thirty,3)

        single_ASTCtest_data = {
            'srs_data': pd.DataFrame(srs_data),
            'recive_data': pd.DataFrame(recive_data),
            'bkgrnd_data': pd.DataFrame(bkgrnd_data),
            'rt': pd.DataFrame(rt_thirty),
            }
        
    else:
        single_ASTCtest_data = None
        # return single_ASTCtest_data
        
    if curr_test['NIC_test'] == 1:
        srs_data = RAW_SLM_datapull(self,curr_test['Source'],'-831_Data.')
        recive_data = RAW_SLM_datapull(self, curr_test['Recieve '],'-831_Data.')
        bkgrnd_data = RAW_SLM_datapull(self,curr_test['BNL'],'-831_Data.')
        rt = RAW_SLM_datapull(self,curr_test['RT'],'-RT_Data.')

        if isinstance(rt, pd.DataFrame):
            rt_thirty = rt.values
        rt_thirty = pd.to_numeric(rt_thirty, errors='coerce')
        rt_thirty = np.round(rt_thirty,3)

        single_NICtest_data = {
            'srs_data': pd.DataFrame(srs_data),
            'recive_data': pd.DataFrame(recive_data),
            'bkgrnd_data': pd.DataFrame(bkgrnd_data),
            'rt': pd.DataFrame(rt_thirty),
            }
        # return single_NICtest_data
    else:
        single_NICtest_data = None
    
    if curr_test['DTC_test'] == 1:
        print("DTC testing enabled, copying data...")
        ## move this into a calc_DTC_data function

        srs_data = RAW_SLM_datapull(self,curr_test['Source'],'-831_Data.')
        recive_data = RAW_SLM_datapull(self, curr_test['Recieve '],'-831_Data.')
        bkgrnd_data = RAW_SLM_datapull(self,curr_test['BNL'],'-831_Data.')
        rt = RAW_SLM_datapull(self,curr_test['RT'],'-RT_Data.')
        srs_door_open = RAW_SLM_datapull(self,curr_test['Source_Door_Open'],'-831_Data.')
        srs_door_closed = RAW_SLM_datapull(self,curr_test['Source_Door_Closed'],'-831_Data.')
        recive_door_open = RAW_SLM_datapull(self, curr_test['Recieve_Door_Open '],'-831_Data.')
        recive_door_closed = RAW_SLM_datapull(self, curr_test['Recieve_Door_Closed '],'-831_Data.')

        if isinstance(rt, pd.DataFrame): # converting RT data to numeric for calcs 
            rt_thirty = rt.values
        rt_thirty = pd.to_numeric(rt_thirty, errors='coerce')
        rt_thirty = np.round(rt_thirty,3)

        single_DTCtest_data = {
            'srs_data': pd.DataFrame(srs_data),
            'recive_data': pd.DataFrame(recive_data),
            'bkgrnd_data': pd.DataFrame(bkgrnd_data),
            'rt': pd.DataFrame(rt_thirty),
            'srs_door_open': pd.DataFrame(srs_door_open),
            'srs_door_closed': pd.DataFrame(srs_door_closed),
            'recive_door_open': pd.DataFrame(recive_door_open),
            'recive_door_closed': pd.DataFrame(recive_door_closed),
            }
    else:
        single_DTCtest_data = None

    #### Put all dataframes into the data structure ####
    test_data = TestData(
        room_properties=room_properties,
        single_AIICtest_data=single_AIICtest_data,
        single_ASTCtest_data=single_ASTCtest_data,
        single_NICtest_data=single_NICtest_data,
        single_DTCtest_data=single_DTCtest_data
    )
    # Create ReportData object with collected data structures
    report_data = ReportData(
        test_data=test_data,
        test_type=test_types
    )
    return report_data

def RAW_SLM_datapull(self, find_datafile, datatype):
    # pass datatype as '-831_Data.' or '-RT_Data.' to pull the correct data
    raw_testpaths = {
        'D': self.slm_data_d_path,
        'E': self.slm_data_e_path,
        # 'A': self.slm_data_a_path
    }
    datafiles = {}
    for key, path in raw_testpaths.items():
        datafiles[key] = [f for f in listdir(path) if isfile(join(path, f))]

    if find_datafile[0] in datafiles:
        datafile_num = datatype + find_datafile[1:] + '.xlsx'
        slm_found = [x for x in datafiles[find_datafile[0]] if datafile_num in x]
        slm_found[0] = raw_testpaths[find_datafile[0]] + slm_found[0]  # If this line errors, the test file is mislabeled or doesn't exist 

    print(slm_found[0])
    if datatype == '-831_Data.':
        srs_data = pd.read_excel(slm_found[0], sheet_name='OBA')
    elif datatype == '-RT_Data.': ## may change to -RT60Pink. for RT60 data
        srs_data = pd.read_excel(slm_found[0], sheet_name='Summary')  # data must be in Summary tab for RT meas.
    return srs_data

### NEW REWORK OF NR CALC ### # functional as of 7/25/24 ### 
def calc_nr_new(srs_overalloct: pd.Series, rec_overalloct: pd.Series, 
                bkgrnd_overalloct: pd.Series, rt_thirty: pd.Series, 
                receive_roomvol: float, nic_vollimit: float, 
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
        ATL_val.append(srs_overalloct[i]-recieve_corr[i]+10*(np.log10(parition_area/sabines.iloc[i])))
    # ATL_val = srs_overalloct - recieve_corr+10*(np.log(parition_area/sabines)) 
    ATL_val = np.round(ATL_val,1)
    # print('ATL val: ',ATL_val)
    return ATL_val

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
    ATL_val_STC = ATL_val[1:17]
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
                field_curve: pd.Series, ref_label: str, field_label: str) -> plt.Figure:
    # pass labels for both curves, depending on AIIC or ASTC 
    # AIIC will be Nrec_ANISPL, ASTC will be ATL_val
    plt.figure(figsize=(10, 6))
    plt.plot(frequencies, ref_curve, label=ref_label, color='red')
    plt.plot(frequencies, field_curve, label=field_label, color='black', marker='s', linestyle='--')
    plt.xlabel('Frequency')
    plt.ylabel(y_label)
    plt.grid(True)
    plt.tick_params(axis='x', rotation=45)
    plt.xticks(frequencies)
    plt.xscale('log')
    plt.gca().get_xaxis().set_major_formatter(plt.ScalarFormatter())  # Format x-ticks as scalars
    plt.gca().xaxis.set_major_locator(ticker.FixedLocator(frequencies))  # Force all x-ticks to display
    # plt.title('Reference vs Measured')
    plt.legend()
    # plt.show()
    fig = plt.get_figure()
    fig.savefig('plot.png')
    plot_fig = Image('plot.png')
    return fig

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

## this function below is a generated function, not sure if im going to include it or just use the GUI interface to iterate through all the tests by default. 
# i guess it makes the single test function much more clean, since i can just troublshoot that single test process if i want to do the entire thing...
## ok try to work this in. add it to the list. 
# def process_all_tests(test_plan_path: Path, slm_data_folder: Path, output_folder: Path) -> List[Path]:
#     test_plan = load_test_plan(test_plan_path)
#     report_paths = []
    
#     for _, test_entry in test_plan.iterrows():
#         slm_data_paths = {
#             'source': slm_data_folder / f"{test_entry['SourceFile']}.xlsx",
#             'receive': slm_data_folder / f"{test_entry['ReceiveFile']}.xlsx",
#             'background': slm_data_folder / f"{test_entry['BackgroundFile']}.xlsx",
#             'rt': slm_data_folder / f"{test_entry['RTFile']}.xlsx",
#         }
        
#         report_path = process_single_test(test_entry, slm_data_paths, output_folder)
#         report_paths.append(report_path)
    
#     return report_paths

# Constants (you might want to move these to a separate config file)
FREQUENCIES = [125, 160, 200, 250, 315, 400, 500, 630, 800, 1000, 1250, 1600, 2000, 2500, 3150, 4000]
NIC_VOLLIMIT = 883  # cu. ft.

# Additional utility functions
def sanitize_filepath(filepath: str) -> str:
    filepath = filepath.replace('T:', '//DLA-04/Shared/')
    filepath = filepath.replace('\\', '/')
    return filepath

def calculate_ASTM_results(test_data: TestData, test_type: TestType):