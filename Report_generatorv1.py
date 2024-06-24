# Report generator v1
from os import listdir, walk
from os.path import isfile, join
import pandas as pd
from pandas import ExcelWriter
import numpy as np
from openpyxl import load_workbook
from openpyxl import Workbook
from openpyxl import cell
from openpyxl.utils import get_column_letter
import xlsxwriter
import time
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
from IPython.display import display, Image
from matplotlib import ticker
import os

def sanitize_filepath(filepath):
    ##"""Sanitize a file path by replacing forward slashes with backslashes."""
    filepath = filepath.replace('T:', '//DLA-04/Shared/')
    filepath = filepath.replace('\\','/')
    # need to add a line to append a / at the end of the filename
    return filepath

#### ADDITION TO MAKE IT METER LETTER AGNOSTIC - rolled change in to RAW_SLM_datapull- just need to add the folders and the function will search through all of them for the correct file.

# writing meter data to report file function definition
# def write_testdata(self,find_datafile, reportfile, newsheetname):
#     rawDtestpath = self.slm_data_d_path
#     rawEtestpath = self.slm_data_e_path
#     rawAtestpath = self.slm_data_e_path
#     rawReportpath = self.report_output_folder_path
#     D_datafiles = [f for f in listdir(rawDtestpath) if isfile(join(rawDtestpath,f))]
#     E_datafiles = [f for f in listdir(rawEtestpath) if isfile(join(rawEtestpath,f))]
#     E_datafiles = [f for f in listdir(rawEtestpath) if isfile(join(rawEtestpath,f))]

#     if find_datafile[0] =='A': ## REPLACE WITH METER 1 - NEED TO DEFINE IN TESTPLAN ##
#         datafile_num = find_datafile[1:]
#         datafile_num = '-831_Data.'+datafile_num+'.xlsx'
#         slm_found = [x for x in A_datafiles if datafile_num in x]
#         slm_found[0] = rawAtestpath+slm_found[0]# If this line errors, the test file is mislabled or doesn't exist 
#         # print(srs_slm_found)
#     elif find_datafile[0] == 'E':
#         datafile_num = find_datafile[1:]
#         datafile_num = '-831_Data.'+datafile_num+'.xlsx'
#         slm_found = [x for x in E_datafiles if datafile_num in x]
#         slm_found[0] = rawEtestpath+slm_found[0]# If this line errors, the test file is mislabled or doesn't exist 

#     print(slm_found[0])

#     srs_data = pd.read_excel(slm_found[0],sheet_name='OBA') # data must be in OBA tab
#     with ExcelWriter(
#     rawReportpath+reportfile,
#     mode="a",
#     engine="openpyxl",
#     if_sheet_exists="replace",
#     ) as writer:
#         srs_data.to_excel(writer, sheet_name=newsheetname) #writes to report file
#     time.sleep(1)

# this raw datapaths method relies on only 2 paths being passed from the GUI - will we need more paths? seems like testing is always limited to 2 meters...
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
    # srs_data = pd.read_excel(slm_found[0], sheet_name='OBA')  # data must be in OBA tab
    # potentially need a write to excel here...similar to previous function
        srs_data = pd.read_excel(slm_found[0],sheet_name='OBA') # data must be in OBA tab
        with ExcelWriter(
        rawReportpath+reportfile, # should be in self.report 
        mode="a",
        engine="openpyxl",
        if_sheet_exists="replace",
        ) as writer:
            srs_data.to_excel(writer, sheet_name=newsheetname) #writes to report file
    return srs_data

def format_SLMdata(self, srs_data):
    srs_thirdoct = srs_data.iloc[7] # hardcoded to SLM export data format
    srs_thirdoct = srs_thirdoct[14:30] # select only the frequency bands of interest
    return srs_thirdoct

def calc_NR_new(srs_overalloct, rec_overalloct, bkgrnd_overalloct, rt_thirty, recieve_roomvol, NIC_vollimit,testtype):
    NIC_vollimit = 150  # cu. ft.
    if recieve_roomvol > NIC_vollimit:
        print('Using NIC calc, room volume too large')
    #constant = np.int32(20.047 * np.sqrt(273.15 + 20)) #  wut is this...
    # intermed = 30 / rt_thirty  # was i compensating for RT calcs?
   
    sabines = 0.049*(recieve_roomvol/rt_thirty)  # this produces accurate sabines values

    # thisval = np.int32(recieve_roomvol * intermed)
    # sabines = thisval / constant
    # sabines = np.round(sabines*(0.921))        
    recieve_corr = list()
    recieve_vsBkgrnd = rec_overalloct - bkgrnd_overalloct
    print('rec vs background:',recieve_vsBkgrnd)
    if testtype == 'AIIC':
        for i, val in enumerate(recieve_vsBkgrnd):
            if val < 5:
                recieve_corr.append(rec_overalloct.iloc[i]-2)
            else:
                recieve_corr.append(10*np.log10(10**(rec_overalloct.iloc[i]/10)-10**(bkgrnd_overalloct.iloc[i]/10)))   
    elif testtype == 'ASTC':
        for i, val in enumerate(recieve_vsBkgrnd):
            print('val:', val)
            print('count: ', i)
            if val < 5:
                recieve_corr.append(rec_overalloct.iloc[i]-2)
            elif val < 10:
                recieve_corr.append(10*np.log10(10**(rec_overalloct.iloc[i]/10)-10**(bkgrnd_overalloct.iloc[i]/10)))
            else:
                recieve_corr.append(rec_overalloct.iloc[i])
        # print('-=-=-=-=-')
        # print('recieve_corr: ',recieve_corr)
    recieve_corr = np.round(recieve_corr,1)
    NR_val = srs_overalloct - recieve_corr
    # Normalized_recieve = recieve_corr / srs_overalloct
    sabines = pd.to_numeric(sabines, errors='coerce')
    Normalized_recieve = recieve_corr-10*(np.log10(108/sabines))
    return NR_val, sabines,recieve_corr, Normalized_recieve

# #### database has raw OBA datasheet, needs to be cleaned for plotting
# OBAdatasheet = 'OBA'
# RTsummarysheet = 'Summary'
# freqbands = ['63','125','250','500','1000','2000','4000','8000']
# srs_OBAdata = pd.read_excel(srs_slm_file[0],OBAdatasheet)
# recive_OBAdata = pd.read_excel(receive_slm_file[0],OBAdatasheet)
# bkgrd_OBAdata = pd.read_excel(bkgrnd_slm_file[0],OBAdatasheet)
# rt = pd.read_excel(rt_slm_file[0],RTsummarysheet)

# transposed_srsOBAdata = srs_OBAdata.transpose()
# srs_OBAdata = srs_OBAdata.dropna()

# # Get the first row of variables as the new column names
# new_column_names = transposed_srsOBAdata.iloc[0]

# # Rename the columns of transposed_srsOBAdata
# transposed_srsOBAdata = transposed_srsOBAdata.rename(columns=new_column_names)

# # Remove the first row (variable labels)
# transposed_srsOBAdata = transposed_srsOBAdata[1:]
# onethird_srs = srs_OBAdata[6:10]
# onethird_rec = recive_OBAdata[6:10]
# onethird_bkgrd = bkgrd_OBAdata[6:10]

# OLD function, will not be used, just reference.
# def write_RTtestdata(find_datafile, reportfile,newsheetname):
#     rawDtestpath = self.slm_data_d_path
#     rawEtestpath = self.slm_data_e_path
#     # rawAtestpath = self.slm_data_a_path
#     rawReportpath = self.report_output_folder_path
#     # A_datafiles = [f for f in listdir(rawDtestpath) if isfile(join(rawDtestpath,f))]
#     D_datafiles = [f for f in listdir(rawDtestpath) if isfile(join(rawDtestpath,f))]
#     E_datafiles = [f for f in listdir(rawEtestpath) if isfile(join(rawEtestpath,f))]
#     if find_datafile[0] =='A':
#         datafile_num = find_datafile[1:]
#         datafile_num = '-RT_Data.'+datafile_num+'.xlsx'
#         slm_found = [x for x in A_datafiles if datafile_num in x]
#         slm_found[0] = rawAtestpath+slm_found[0]# If this line errors, the test file is mislabled or doesn't exist 
#         # print(srs_slm_found)
#     elif find_datafile[0] == 'E':
#         datafile_num = find_datafile[1:]
#         datafile_num = '-RT_Data.'+datafile_num+'.xlsx'
#         slm_found = [x for x in E_datafiles if datafile_num in x]
#         slm_found[0] = rawEtestpath+slm_found[0] # If this line errors, the test file is mislabled or doesn't exist 

#     print(slm_found[0])

#     srs_data = pd.read_excel(slm_found[0],sheet_name='Summary')# data must be in Summary tab for RT meas.
#     # could reduce this function by also passing the sheet to be read into the args. 
#     # transfer to either master Pandas database or SQL database

#     with ExcelWriter(
#     rawReportpath+reportfile,
#     mode="a",
#     engine="openpyxl",
#     if_sheet_exists="replace",
#     ) as writer:
#         srs_data.to_excel(writer, sheet_name=newsheetname) 
#     time.sleep(1)
#     # excel.Quit()
#     return srs_data

### DATA CALC FUNCTIONS ###

# ATL calc, reworked and funtional - use Recieve room recording, not the tapper level for this calc. this does not take the tapper into account
def calc_ATL_val(srs_overalloct,rec_overalloct,bkgrnd_overalloct,parition_area,recieve_roomvol,sabines):
    ASTC_vollimit = 883
    if recieve_roomvol > ASTC_vollimit:
        print('Using NIC calc, room volume too large')
    recieve_corr = list()
    bkgrnd_overalloct = pd.to_numeric(bkgrnd_overalloct, errors='coerce')
    bkgrnd_overalloct = np.array(bkgrnd_overalloct)
    rec_overalloct = pd.to_numeric(rec_overalloct, errors='coerce')
    rec_overalloct = np.array(rec_overalloct)
    
    recieve_vsBkgrnd = rec_overalloct - bkgrnd_overalloct
    print('recieve vs background:',recieve_vsBkgrnd)
    recieve_vsBkgrnd = pd.to_numeric(recieve_vsBkgrnd, errors='coerce')
    recieve_vsBkgrnd = np.array(recieve_vsBkgrnd)

    for i, val in enumerate(recieve_vsBkgrnd):
        if val < 5:
            recieve_corr.append(rec_overalloct[i]-2)
        elif val < 10:
            recieve_corr.append(10*np.log10(10**(rec_overalloct[i]/10)-10**(bkgrnd_overalloct[i]/10)))
        else:
            recieve_corr.append(rec_overalloct[i])
    recieve_corr = np.round(recieve_corr,1)
    print('corrected recieve: ',recieve_corr)
    
    sabines = pd.to_numeric(sabines, errors='coerce')
    #convert sabines to array
    sabines = np.array(sabines)
    print('sabines: ',sabines)
    srs_overalloct = pd.to_numeric(srs_overalloct, errors='coerce')
    srs_overalloct = np.array(srs_overalloct)
    print('srs_overalloct: ',srs_overalloct)
    ATL_val = srs_overalloct - recieve_corr+10*(np.log10(parition_area/sabines))
    return ATL_val

## functional. need to verify with excel doc calcs
def calc_AIIC_val(Normalized_recieve_IIC):
    pos_diffs = list()
    AIIC_start = 94
    AIIC_contour_val = 16
    IIC_contour = list()
    new_sum = 0
    diff_negative_max = 0
    IIC_curve = [2,2,2,2,2,2,1,0,-1,-2,-3,-6,-9,-12,-15,-18]
    # initial application of the IIC curve to the first AIIC start value 
    for vals in IIC_curve:
        IIC_contour.append(vals+AIIC_start)
    Contour_curve_result = IIC_contour - Normalized_recieve_IIC
    print('Normalized recieve ANISPL: ', Normalized_recieve_IIC)

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
        if new_sum > 32 or diff_negative_max > 8:
            print('Difference condition met! AIIC value: ', AIIC_contour_val) # 
            print('AIIC result curve: ', Contour_curve_result)
            return AIIC_contour_val, Contour_curve_result
        # condition not met, resetting arrays
        pos_diffs = []
        IIC_contour = []
        print('difference condition not met, subtracting 1 from AIIC start and recalculating the IIC contour')
        AIIC_start = AIIC_start - 1
        AIIC_contour_val = AIIC_contour_val + 1
        print('AIIC start: ', AIIC_start)
        print('AIIC contour value: ', AIIC_contour_val)
        for vals in IIC_curve:
            IIC_contour.append(vals+AIIC_start)
        Contour_curve_result = IIC_contour - Normalized_recieve_IIC
        if AIIC_start <10: break

# need to validate
def calc_ASTC_val(ATL_val):
    pos_diffs = list()
    diff_negative=0
    diff_positive=0 
    ASTC_start = 20
    New_curve =list()
    new_sum = 0
    STCCurve = [-16, -13, -10, -7, -4, -1, 0, 1, 2, 3, 4, 4, 4, 4, 4, 4]
    while (diff_negative < 8 and new_sum < 32):
        # print('starting loop')
        print('ASTC fit test value: ', ASTC_start)
        for vals in STCCurve:
            New_curve.append(vals+ASTC_start)
        ASTC_curve = New_curve - ATL_val
        # print('ASTC curve: ',ASTC_curve)

        diff_negative =  np.max(ASTC_curve - ATL_val)
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
            break
        pos_diffs = []
        New_curve = []
        ASTC_start = ASTC_start + 1
        
        
        if ASTC_start >80: break

# need to validate
def plot_curves(frequencies,Y_label,Ref_curve, Field_curve,Ref_label, Field_label):
    # pass labels for both curves, depending on AIIC or ASTC 
    # AIIC will be Nrec_ANISPL, ASTC will be ATL_val
    plt.figure(figsize=(10, 6))
    plt.plot(frequencies, Ref_curve, label=Ref_label, color='red')
    plt.plot(frequencies, Field_curve, label=Field_label, color='black', marker='s', linestyle='--')
    plt.xlabel('Frequency')
    plt.ylabel(Y_label)
    plt.grid(True)
    plt.tick_params(axis='x', rotation=45)
    plt.xticks(frequencies)
    plt.xscale('log')
    plt.gca().get_xaxis().set_major_formatter(plt.ScalarFormatter())  # Format x-ticks as scalars
    plt.gca().xaxis.set_major_locator(ticker.FixedLocator(frequencies))  # Force all x-ticks to display
    # plt.title('Reference vs Measured')
    plt.legend()
    plt.show()
    fig = plt.get_figure()
    fig.savefig('plot.png')
    plot_fig = Image('plot.png')
    return plot_fig
# need to write a big overall wrapper function here that will go through the testplan, depending on what test type, pull the data from each excel file, write the data to the database variables for that test and then use those database variables to plot the data in the report.
# for each testplan entry, first write the room properties database: 
# Organize the variables into a dictionary


# _=-=-=-=--=-=-=-_+_=-=-=-=-=-=-_+_=-=-=-=-=-= Text lookups for report print -=-=-=-=-=-=-=- 

stockNIC_note = ["The receiver and/or source room had a volume exceeding 150 m3 (5,300 cu. ft.), and the absorption of the receiver and/or source room was greater than the maximum allowed per E336-16, Paragraph 9.4.1.2.",
                 "The receiver and/or source room was not an enclosed space.",
                 "The receiver and/or source room has a volume less than the minimum volume requirement of 25 m3 (883 cu. ft.).",
                 "The receiver and/or source room has one or more dimensions less than the minimum requirement of 2.3 m (7.5 ft.)."]

ISR_ony_report = "The receiver room had a volume less than the minimum volume requirement of 40 m3."
stockISR_notes = ["The receiver and/or source room had a volume exceeding 150 m3 (5,300 cu. ft.), and the absorption of the receiver room was greater than the maximum allowed per E1007-16, Paragraph 10.3.1 and 10.4.5.",
"The receiver and/or source room was not an enclosed space.", 
"The receiver and/or source room has a volume less than the minimum volume requirement of 40 m3 (1413 cu. ft.).",
"The receiver and/or source room has one or more dimensions less than the minimum requirement of 2.3 m (7.5 ft.)."]

#standards 
standards_text = (("ASTC Test Procedure ASTM E336-16",	"Standard Test Method for Measurement of Airborne Sound Attenuation between Rooms in Buildings"),("STC Calculation	ASTM E413-16",	"Classification for Rating Sound Insulation"),("AIIC Test Procedure	ASTM E1007-14",	"Standard Test Method for Field Measurement of Tapping Machine Impact Sound Transmission Through Floor-Ceiling Assemblies and Associated Support Structure"),
("IIC Calculation	ASTM E989-06(2012)",	"Standard Classification for Determination of Impact Insulation Class (IIC)"),
("RT60 Test Procedure	ASTM E2235-04(2012)",	"Standard Test Method for Determination of Decay Rates for Use in Sound Insulation Test Methods"))

# refer to single standards like this : standards_text[0][0]

##statement of conformance 
#Testing was conducted in general accordance with ASTM E1007-14, with all exceptions noted below. 
#All requirements for measuring and reporting Absorption Normalized Impact Sound Pressure Level (ANISPL) and Apparent Impact Insulation Class (AIIC) were met.								
# code: =CONCATENATE("Testing was conducted in general accordance with ",AIIC or ASTC or NIC", with all exceptions noted below. All requirements for measuring and reporting Apparent Transmission Loss (ATL) and Apparent Sound Transmission Class (ASTC) were met.")


# test environment 
#text: 
# The source room was 2nd Floor Bed 1. The space was finished, unfurnished. The floor was Carpet. The ceiling was gyp. The walls were gyp. All doors and windows were closed during the testing period. The source room had a volume of approximately 1176 cu. ft.								
#code:
# =CONCATENATE('SLM Data'!B20,'SLM Data'!$C$20, 'SLM Data'!B21,'SLM Data'!$C$21,'SLM Data'!$B$22, 'SLM Data'!$C$22, 'SLM Data'!$B$23, 'SLM Data'!$C$23, 'SLM Data'!$B$24,'SLM Data'!$C$24,". ",'SLM Data'!$C$25,'SLM Data'!$C$26," The source room had a volume of approximately ",'SLM Data'!C16," cu. ft.")

# test procedure 
test_procedure_pg = 'Determination of space-average sound pressure levels was performed via the manually scanned microphones techique, described in ' + standards_text[0][0] + ', Paragraph 11.4.3.3.'+ 'The source room was selected in accordance with ASTM E336-11 Paragraph 9.2.5, which states that "If a corridor must be used as one of the spaces for measurement of ATL or FTL, it shall be used as the source space."'
# code:
# =CONCATENATE("The test was performaned in general accordance with ",AIIC or ASTC or NIC,". Determination of Space-Average Levels performed via the manually scanned microphones techique, described in ",'SLM Data'!C59,", Paragraph 11.4.2.2.")
# The test was performaned in general accordance with ASTM E1007-14. Determination of Space-Average Levels performed via the manually scanned microphones techique, described in ASTM E1007-14, Paragraph 11.4.2.2.								

flanking_text = "Flanking transmission was not evaluated."

# To evaluate room absorption, 1 microphone was used to measure 4 decays at 4 locations around the receiving room for a total of 16 measurements, per ASTM E2235-04(2012).								

RT_text = "To evaluate room absorption, 1 microphone was used to measure 4 decays at 4 locations around the receiving room for a total of 16 measurements, per"+standards_text[4][0]

# ASTC and NIC final result and blurb

# ASTC result, concat with this: 
ASTC_results_ATverbage = ' was calculated. The ASTC rating is based on Apparent Transmission Loss (ATL), and includes the effects of noise flanking. The ASTC reference contour is shown on the next page, and has been “fit” to the Apparent Transmission Loss values, in accordance with the procedure of'

#NIC results, concat with this:
NIC_results_NRverbage = ' was calculated. The NIC rating is based on Noise Reduction (NR), and includes the effects of noise flanking. The NIC reference contour is shown on the next page, and has been “fit” to the Apparent Transmission Loss values, in accordance with the procedure of'
# after results
results_blurb = 'The results stated in this report represent only the specific construction and acoustical conditions present at the time of the test. Measurements performed in accordance with this test method on nominally identical constructions and acoustical conditions may produce different results.'

#test instrumentation
# table with SLM serial, micpreamp, mic, calibrator, speaker, noise gen.
# LOGIC NEEDED: 
#  import very simply spreadsheet with this information preloaded, and pull from the spreadsheet
#  or select from a menu of all the SLMs, calibrators, tapping machine, speaker.
#  
# Example text Entry Box for the fifth path - must be inside a kivy app and build application:
# class FileLoaderApp(App):
#     def build(self):

        # self.fifth_text_input = TextInput(multiline=False, hint_text='File Path 5')
        # self.fifth_text_input.bind(on_text_validate=self.on_text_validate)
        # layout.add_widget(self.fifth_text_input)


Equip_type_list = [ "Sound Level Meter 1",
"Microphone Pre-Amp:",
"Microphone:",
"Calibrator:",
"Sound Level Meter 2",
"Microphone Pre-Amp:",
"Microphone:",
"Calibrator:",
"Amplified Loudspeakers",
"Noise Generator:"
]

Manuf_list = ["Larson Davis",
"Larson Davis",
"Larson Davis",
"Larson Davis",
"Larson Davis",
"Larson Davis",
"Larson Davis",
"Larson Davis",
"QSC",
"NTi Audio"
]

Model_numlist = ["831",
"PRM831",
"377B20",
"CAL200",
"831",
"PRM831",
"377B20",
"CAL200",
"K10",
"MR-PRO"
]

Serial_numList = ["3784",
"051188",	
"301698",	
"2775671",	
"4328",	
"046469",	
"168830",	
"5955",	
"GAA530909",	
"0162"
]

Last_NISTcal_list = ["[9/19/2022",
"9/19/2022",
"9/16/2022",
"9/19/2022",
"10/24/2022",
"10/24/2022",
"10/20/2022",
"10/26/2022",
"N/A",
"N/A"
]
LastLocalcalLIst = ["Apr 2024",
"Apr 2024",
"Apr 2024",
"N/A",
"Apr 2024",
"Apr 2024",
"Apr 2024",
"N/A",
"N/A",
"N/A"
]
test_instrumentation = pd.DataFrame(
    {
        "Equipment Type": Equip_type_list,
        "Manufacturer": Manuf_list,
        "Model Number": Model_numlist,
        "Serial Number": Serial_numList,
        "Last NIST Tracable Calibration": Last_NISTcal_list,
        "Last Local Calibration" : LastLocalcalLIst
    },
        # index=[0]
)    

## STATEMENT OF TEST RESULTS 
statement_test_results_text =' STATEMENT OF TEST RESULTS: '

###################################
####_+#_+_+_#+_#+_#+_####_######################################
## do i put the GUI here YES - ASTC GUI proto.py works - just need to add in the reportlab formatting and get going.
# wrap the following into a funciton - take in the testplan entry, and return the reportlab formatted pdf output path

def create_report(self,curr_test, single_test_dataframe, test_type):
        # all the code below
    # Kaulu by gentry testing ## EXAMPLE DATA CREATE LOOP FOR EACH TESTPLAN ENTRY
    testplan_path ='//DLA-04/Shared/KAILUA PROJECTS/2024/24-004 Kaulu by Gentry ASTC - AIIC testing/Documents/TestPlan_Kaulu_ASTM_testingv1.xlsx'
    # need to modify for current project number
    # test_list = pd.read_excel(testplan_path)
    # testnums = test_list['Test Label']
    ### Pass over the testplan entry database 

    #### database has raw OBA datasheet, needs to be cleaned for plotting
    ## this is done inside the report generation section


    # OBAdatasheet = 'OBA'
    # RTsummarysheet = 'Summary'
    # freqbands = ['63','125','250','500','1000','2000','4000','8000']

    # srs_OBAdata = pd.read_excel(single_test_dataframe['srs_data'],OBAdatasheet)
    # recive_OBAdata = pd.read_excel(single_test_dataframe['recive_data'],OBAdatasheet)
    # bkgrd_OBAdata = pd.read_excel(single_test_dataframe['bkgrnd_data'],OBAdatasheet)
    # rt = pd.read_excel(single_test_dataframe['rt'],RTsummarysheet)

    # transposed_srsOBAdata = srs_OBAdata.transpose()
    # srs_OBAdata = srs_OBAdata.dropna()

    # # Get the first row of variables as the new column names
    # new_column_names = transposed_srsOBAdata.iloc[0]

    # # Rename the columns of transposed_srsOBAdata
    # transposed_srsOBAdata = transposed_srsOBAdata.rename(columns=new_column_names)

    # # Remove the first row (variable labels)
    # transposed_srsOBAdata = transposed_srsOBAdata[1:]
    # onethird_srs = srs_OBAdata[6:10]
    # onethird_rec = recive_OBAdata[6:10]
    # onethird_bkgrd = bkgrd_OBAdata[6:10]


##### _+_+#__#+_#+_+_+_####################### REPORT GENERATION CODE ##############################
    import matplotlib.pyplot as plt
    import matplotlib.ticker as ticker
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import letter
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.platypus import BaseDocTemplate, PageTemplate, Frame, Paragraph, Table, TableStyle, Spacer,PageBreak, KeepInFrame, Image
    from reportlab.pdfgen import canvas

    from reportlab.lib.units import inch

    ######## =--=-=-=--= 
    ### custom margins
    # Specify custom margins (in points, where 1 inch equals 72 points)
    left_margin = right_margin = 0.75 * 72  # 0.75 inches
    top_margin = 0.25 * 72 # 0.25 inch
    bottom_margin = 1 * 72  # 1 inch
    header_height = 2 * inch
    footer_height = 0.5 * inch
    main_content_height = letter[1] - top_margin - bottom_margin - header_height - footer_height

    # # Create a document with custom page templates
    # doc = BaseDocTemplate("report_with_header.pdf", pagesize=letter)

    #### -=-=--=-= 

    # Document Setup 
    # document name and page size
    #  Example ouptut file name: '24-006 AIIC Test Report_1.1.1.pdf'
    #  format: project_name + test_type + test_num + '.pdf'
    doc_name = f'{single_test_dataframe['room_properties']['Project_Name'][0]} {test_type} Test Report_{single_test_dataframe['room_properties']['Test_Label'][0]}.pdf'

    doc = BaseDocTemplate(doc_name, pagesize=letter,
                        leftMargin=left_margin, rightMargin=right_margin,
                        topMargin=top_margin, bottomMargin=bottom_margin)

    # Define Frames for the header, main content, and footer
    header_frame = Frame(left_margin, letter[1] - top_margin - header_height, letter[0] - 2 * left_margin, header_height, id='header')
    main_frame = Frame(left_margin, bottom_margin + footer_height, letter[0] - 2 * left_margin, letter[1] - top_margin - header_height - footer_height - bottom_margin, id='main')
    footer_frame = Frame(left_margin, bottom_margin, letter[0] - 2 * left_margin, footer_height, id='footer')

    # Create styles
    styles = getSampleStyleSheet()
    custom_title_style = styles['Heading1']


    #     canvas.saveState()
    #     canvas.setFont('Helvetica', 10)
    #     canvas.drawCentredString(letter[0] / 2, letter[1] - top_margin - header_height / 2, "Field Impact Sound Transmission Test Report")
    #     canvas.drawCentredString(letter[0] / 2, bottom_margin + footer_height / 2, f"Page {doc.page} of 4")
    #     canvas.restoreState()
    # Define header elements
    def header_elements():
        elements = []
        if test_type == 'AIIC':
            elements.append(Paragraph("<b>Field Impact Sound Transmission Test Report</b>", custom_title_style))
            elements.append(Paragraph("<b>Apparent Impact Insulation Class (AIIC)</b>", custom_title_style))
        elif test_type == 'ASTC':
            elements.append(Paragraph("<b>Field Sound Transmission Test Report</b>", custom_title_style))
            elements.append(Paragraph("<b>Apparent Sound Transmission Class (ASTC)</b>", custom_title_style))
        elif test_type == 'NIC':
            elements.append(Paragraph("<b>Field Sound Transmission Test Report</b>", custom_title_style))
            elements.append(Paragraph("<b>Noise Isolation Class (NIC)</b>", custom_title_style))
        
        # elements.append(Paragraph("<b>Field Impact Sound Transmission Test Report</b>", custom_title_style))
        # elements.append(Paragraph("<b>Apparent Impact Insulation Class (AIIC)</b>", custom_title_style))
        elements.append(Spacer(1, 10))
        leftside_data = [
            ["Report Date:", single_test_dataframe['room_properties']['Report_Date'][0]],
            ['Test Date:', single_test_dataframe['room_properties']['Test_Date'][0]],
            ['DLAA Test No', single_test_dataframe['room_properties']['Test_Label'][0]]
        ]
        rightside_data = [
            ["Source Room:", single_test_dataframe['room_properties']['Source_Room'][0]],
            ["Receiver Room:", single_test_dataframe['room_properties']['Receiving_Room'][0]],
            ["Test Assembly:", single_test_dataframe['room_properties']['tested_assembly'][0]]
        ]

        table_left = Table(leftside_data)
        table_right = Table(rightside_data)
        table_left.setStyle(TableStyle([('GRID', (0, 0), (-1, -1), 1, colors.white)]))
        table_right.setStyle(TableStyle([('GRID', (0, 0), (-1, -1), 1, colors.white)]))

        table_combined_lr = Table([[table_left, table_right]], colWidths=[doc.width / 2.0] * 2)
        elements.append(KeepInFrame(maxWidth=doc.width, maxHeight=header_height, content=[table_combined_lr], hAlign='LEFT'))
        elements.append(Spacer(1, 10))
        elements.append(Paragraph('Test site: ' + sitename, styles['Normal']))
        elements.append(Spacer(1, 5))
        elements.append(Paragraph('Client: ' + client_Name, styles['Normal']))
        return elements

    # Define a function to draw the header and footer
    def header_footer(canvas, doc):
        canvas.saveState()

        # Build the header
        header_frame._leftPadding = header_frame._rightPadding = 0
        header_story = header_elements()
        header_frame.addFromList(header_story, canvas)

        # Footer
        canvas.setFont('Helvetica', 10)
        footer_text = f"Page {doc.page}"
        canvas.drawCentredString(letter[0] / 2, bottom_margin + footer_height / 2, footer_text)

        canvas.restoreState()

    # Create a page template with header and footer
    page_template = PageTemplate(id='Standard', frames=[main_frame, header_frame, footer_frame], onPage=header_footer)
    doc.addPageTemplates([page_template])
    # Main document content (Add your main document elements here)
    main_elements = []

    ## -=-==-= Heading 'STANDARDS'  # -=-=-=-=-=--=-=-=-=-=--=-=-
    styleHeading = ParagraphStyle('heading', parent=styles['Normal'], spaceAfter=10)
    main_elements.append(Spacer(1, 10))  # Adds some space 
    main_elements.append(Paragraph('<u>STANDARDS:</u>', styleHeading))

    # Standards table data
    if test_type == 'AIIC':
        standards_data = [
            ['ASTM E1007-14', Paragraph('Standard Test Method for Field Measurement of Tapping Machine Impact Sound Transmission Through Floor-Ceiling Assemblies and Associated Support Structure',styles['Normal'])],
            ['ASTM E413-16', Paragraph('Standard Classification for Rating Sound Insulation',styles['Normal'])],
            ['ASTM E1007-14', Paragraph('Standard Test Method for Field Measurement of Tapping Machine Impact Sound Transmission Through Floor-Ceiling Assemblies and Associated Support Structure',styles['Normal'])],
            ['ASTM E989-06(2012)', Paragraph('Standard Classification for Determination of Impact Insulation Class (IIC)',styles['Normal'])],
            ['ASTM E2235-04(2012)', Paragraph('Standard Test Method for Determination of Decay Rates for Use in Sound Insulation Test Methods',styles['Normal'])]
        ]
    elif test_type == 'ASTC' or test_type == 'NIC':
        standards_data = [
            ['ASTM E336-16', Paragraph('Standard Test Method for Measurement of Airborne Sound Attenuation between Rooms in Buildings',styles['Normal'])],
            ['ASTM E413-16', Paragraph('Classification for Rating Sound Insulation',styles['Normal'])],
            ['ASTM E2235-04(2012)', Paragraph('Standard Test Method for Determination of Decay Rates for Use in Sound Insulation Test Methods',styles['Normal'])]
        ]
    # standards_data = [
    #     ['ASTM E1007-14', Paragraph('Standard Test Method for Field Measurement of Tapping Machine Impact Sound Transmission Through Floor-Ceiling Assemblies and Associated Support Structure',styles['Normal'])],
    #     ['ASTM E989-06(2012)', Paragraph('Standard Classification for Determination of Impact Insulation Class (IIC)',styles['Normal'])],
    #     ['ASTM E2235-04(2012)', Paragraph('Standard Test Method for Determination of Decay Rates for Use in Sound Insulation Test Methods',styles['Normal'])]
    # ]

    # Create the table
    standards_table = Table(standards_data, hAlign='LEFT')
    standards_table.setStyle(TableStyle([
        ('INNERGRID', (0, 0), (-1, -1), 0.25, colors.white),
        ('BOX', (0, 0), (-1, -1), 0.25, colors.white),
        ('LEFTPADDING', (0, 0), (-1, -1), 6),
        ('RIGHTPADDING', (0, 0), (-1, -1), 6),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('ALIGN',(0,0), (-1,-1),'LEFT')
    ]))

    # Add the table to the elements list
    main_elements.append(standards_table)
    ## these elements are in the single test dataframe, room_properties
    ### 
    # Heading 'TEST ENVIRONMENT'
    main_elements.append(Paragraph("<u>TEST ENVIRONMENT:</u>", styleHeading))
    main_elements.append(Paragraph('The source room was '+single_test_dataframe['room_properties']['Source_Room'][0]+'. The space was'+single_test_dataframe['room_properties']['source_room_finish'][0]+'. The floor was '+single_test_dataframe['room_properties']['srs_floor'][0]+'. The ceiling was '+single_test_dataframe['room_properties']['srs_ceiling'][0]+". The walls were"+single_test_dataframe['room_properties']['srs_walls'][0]+". All doors and windows were closed during the testing period. The source room had a volume of approximately "+single_test_dataframe['room_properties']['source_room_vol'][0]+"cu. ft."))
    main_elements.append(Spacer(1, 10))  # Adds some space 
    ### Recieve room paragraph
    main_elements.append(Paragraph('The receiver room was '+single_test_dataframe['room_properties']['Receiving_Room'][0]+'. The space was'+single_test_dataframe['room_properties']['receiver_room_finish'][0]+'. The floor was '+single_test_dataframe['room_properties']['rec_floor'][0]+'. The ceiling was '+single_test_dataframe['room_properties']['rec_ceiling'][0]+". The walls were"+single_test_dataframe['room_properties']['rec_Wall'][0]+". All doors and windows were closed during the testing period. The source room had a volume of approximately "+single_test_dataframe['room_properties']['receive_room_vol'][0]+"cu. ft."))
    main_elements.append(Spacer(1, 10))  # Adds some space 
    main_elements.append(Paragraph('The test assembly measured approximately '+single_test_dataframe['room_properties']['partition_dim'][0]+", and had an area of approximately "+single_test_dataframe['room_properties']['partition_area'][0]+"sq. ft."))
    main_elements.append(Spacer(1, 10))  # Adds some space 
    # Heading 'TEST ENVIRONMENT'
    main_elements.append(Paragraph("<u>TEST ASSEMBLY:</u>", styleHeading))
    main_elements.append(Spacer(1, 10))  # Adds some space 
    main_elements.append(Paragraph("The tested assembly was the"+single_test_dataframe['room_properties']['Test_Assembly_Type'][0]+"The assembly was not field verified, and was based on information provided by the client and drawings for the project. The client advised that no slab treatment or self-leveling was applied. Results may vary if slab treatment or self-leveling or any adhesive is used in other installations."))
    # ##### END OF FIRST PAGE TEXT  - ########
        
    main_elements.append(PageBreak())
    ## 2nd page text : equipment table and test procedure 
    # test procedure 
    main_elements.append(Paragraph("<u>TEST PROCEDURE:</u>", styleHeading))
    main_elements.append(Paragraph('Determination of space-average sound pressure levels was performed via the manually scanned microphones techique, described in ' + standards_data[0][0] + ', Paragraph 11.4.3.3.'+ "The source room was selected in accordance with ASTM E336-11 Paragraph 9.2.5, which states that 'If a corridor must be used as one of the spaces for measurement of ATL or FTL, it shall be used as the source space.'"))
    main_elements.append(Spacer(1,10))
    main_elements.append(Paragraph("Flanking transmission was not evaluated."))
    main_elements.append(Paragraph("To evaluate room absorption, 1 microphone was used to measure 4 decays at 4 locations around the receiving room for a total of 16 measurements, per"+standards_data[2][0]))
    main_elements.append(Paragraph("<u>TEST INSTRUMENTATION:</u>", styleHeading))

    ## this should shift between meters - predefined, selected with GUI and lookup tabled
    # # will need to create a dynamic table 
    if test_type == 'ASTC':
        test_instrumentation_table = [
            ["Equipment Type","Manufacturer","Model Number","Serial Number",Paragraph("Last NIST Traceable Calibration"),Paragraph("Last Local Calibration")],
            ["Sound Level Meter 1", "Larson Davis","831","4328","10/24/2022","Apr 2024"],
            ["Microphone Pre-Amp:","Larson Davis","PRM831","046469","10/24/2022","Apr 2024"],
            ["Microphone:","Larson Davis","377B20","168830","10/20/2022","Apr 2024"],
            ["Calibrator:","Larson Davis","CAL200","5955","10/26/2022","N/A"],
            ["Sound Level Meter 2","Larson Davis","831","4328","10/24/2022","Apr 2024",],
            ["Microphone Pre-Amp:","Larson Davis","PRM831","046469","10/24/2022","Apr 2024"],
            ["Microphone:","Larson Davis","377B20","168830","10/20/2022","Apr 2024"],
            ["Calibrator:","Larson Davis","CAL200","5955","10/26/2022","N/A"],
            ["Amplified Loudspeaker","QSC","K10","GAA530909","N/A","N/A"]
        ]
    elif test_type == 'AIIC':
        ### AIIC table includes tapper 
        test_instrumentation_table = [["Equipment Type","Manufacturer","Model Number","Serial Number",Paragraph("Last NIST Traceable Calibration"),Paragraph("Last Local Calibration")],
        ["Tapping Machine:","Norsonics","CAL200","2775671","9/19/2022","N/A"],
        ["Sound Level Meter","Larson Davis","831","4328","10/24/2022","4/4/2024"],
        ["Microphone Pre-Amp","Larson Davis","PRM831","046469","10/24/2022","4/4/2024"],
        ["Microphone","Larson Davis","377B20","168830","10/20/2022","4/4/2024"],
        ["Calibrator","Larson Davis","CAL200","5955","10/26/2022","N/A"],
        ["Amplified Loudspeaker","QSC","K10","GAA530909","N/A","N/A"],
        ["Noise Generator","NTi Audio","MR-PRO","0162","N/A","N/A"]
        ]

    # Create the table - will change ASTC vs AIIC -=-= insert logic here

    test_instrumentation_table = Table(test_instrumentation_table, hAlign='LEFT') ## hardcoded, change to table variable for selected test
    test_instrumentation_table.setStyle(TableStyle([
        ('INNERGRID', (0, 0), (-1, -1), 0.25, colors.white),
        ('BOX', (0, 0), (-1, -1), 0.25, colors.white),
        ('LEFTPADDING', (0, 0), (-1, -1), 6),
        ('RIGHTPADDING', (0, 0), (-1, -1), 6),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('ALIGN',(0,0), (-1,-1),'LEFT')
    ]))

    # Add the table to the elements list
    main_elements.append(test_instrumentation_table)
    #### END OF SECOND PAGE TXT -=-=-=-=--=-=-=-=-=-=-=-=-=-=-=-=-=-
    main_elements.append(PageBreak())
    main_elements.append(Paragraph("<u>STATEMENT OF TEST RESULTS:</u>", styleHeading))
    #### Main calculation table section --- also split into AIIC, ASTC, NIC report results. 
    # function returning a database table to display frequency, level, OBA of source, reciever, RT60, NR, ATL, Exceptions
    ######### CALCS FOR PDF TABLES AND PLOTS WILL GO IN THIS SCRIPT ###########
    # freqThirdoct = single_test_dataframe['srs_data'].iloc[6]
    # frequencies = freqThirdoct[14:30]
    frequencies =[125,160,200,250,315,400,500,630,800,1000,1250,1600,2000,2500,3150,4000]
    # freqThirdoct.iloc[6]
    if test_type == 'AIIC':
        #dataframe for AIIC is in single_test_dataframe
        onethird_srs = format_SLMdata(single_test_dataframe['AIIC_source']) 
        average_pos = []
        # get average of 4 tapper positions for recieve total OBA
        for i in range(1, 5):
            pos_input = f'AIIC_pos{i}'
            pos_data = format_SLMdata(single_test_dataframe[pos_input])
            average_pos.append(pos_data)

        onethird_rec_Total = sum(average_pos) / len(average_pos)
        # this needs to be an average of the 4 tapper positions, stored in a dataframe of the average of the 4 dataframes octave band results. 


        onethird_bkgrd = format_SLMdata(single_test_dataframe['bkgrnd_data'])
        rt_thirty = single_test_dataframe['rt']['Unnamed: 10'][25:41]/1000

        calc_NR, sabines, corrected_recieve,Nrec_ANISPL = calc_NR_new(onethird_srs, onethird_rec_Total, onethird_bkgrd, rt_thirty,room_properties['Recieve Vol'][0],NIC_vollimit=883,testtype='AIIC')
        
        # ATL_val = calc_ATL_val(onethird_srs, onethird_rec, onethird_bkgrd,rt_thirty,room_properties['Partition area'][0],room_properties['Recieve Vol'][0])
        AIIC_contour_val, Contour_curve_result = calc_AIIC_val(Nrec_ANISPL)

        IIC_curve = [2,2,2,2,2,2,1,0,-1,-2,-3,-6,-9,-12,-15,-18]
        IIC_contour_final = list()
        # initial application of the IIC curve to the first AIIC start value 
        for vals in IIC_curve:
            IIC_contour_final.append(vals+(110-AIIC_contour_val))
        # print(IIC_contour_final)
        #### Contour_final is the AIIC contour that needs to be plotted vs the ANISPL curve- we have everything to plot the graphs and the results table  #####
        Ref_label = f'AIIC {AIIC_contour_val} Contour'
        Field_IIC_label = 'Absorption Normalized Impact Sound Pressure Level, ANISPL (dB)'

        Test_result_table = pd.DataFrame(
            {
                "Frequency": frequencies,
                "Absorption Normalized Impact Sound Pressure Level, ANISPL (dB)	": Nrec_ANISPL,
                "Average Receiver Background Level": onethird_bkgrd,
                "Average RT60 (Seconds)": rt_thirty,
                "Exceptions noted to ASTM E1007-14": AIIC_Exceptions
            }
        )

    elif test_type == 'ASTC':
        onethird_rec = format_SLMdata(single_test_dataframe['recive_data'])
        onethird_srs = format_SLMdata(single_test_dataframe['srs_data'])
        onethird_bkgrd = format_SLMdata(single_test_dataframe['bkgrnd_data'])
        rt_thirty = single_test_dataframe['rt']['Unnamed: 10'][25:41]/1000

        ATL_val,corrected_STC_recieve = calc_ATL_val(onethird_srs, onethird_rec, onethird_bkgrd,room_properties['Partition area'][0],room_properties['Recieve Vol'][0],sabines)

        calc_NR, sabines, corrected_recieve,Nrec_ANISPL = calc_NR_new(onethird_srs, onethird_rec, onethird_bkgrd, rt_thirty,room_properties['Recieve Vol'][0],NIC_vollimit=883,testtype='ASTC')

        Test_result_table = pd.DataFrame(
            {
                "Frequency": frequencies,
                "L1, Average Source Room Level (dB)": onethird_srs,
                "L2, Average Corrected Receiver Room Level (dB)":corrected_STC_recieve,
                "Average Receiver Background Level (dB)": onethird_bkgrd,
                "Average RT60 (Seconds)": rt_thirty,
                "Noise Reduction, NR (dB)": calc_NR,
                "Apparent Transmission Loss, ATL (dB)": ATL_val,
                "Exceptions": ASTC_Exceptions
            }
        )
    elif test_type == 'NIC':
        Test_result_table = pd.DataFrame(
            {
                "Frequency": freqbands,
                "Source OBA": onethird_srs,
                "Reciever OBA": onethird_rec,
                "Background OBA": onethird_bkgrd,
                "NR": rt['NR'],
                "Exceptions": NIC_exceptions
            }
        )
    
    Test_result_table = Table(Test_result_table, hAlign='LEFT') ## hardcoded, change to table variable for selected test
    Test_result_table.setStyle(TableStyle([
        ('INNERGRID', (0, 0), (-1, -1), 0.25, colors.white),
        ('BOX', (0, 0), (-1, -1), 0.25, colors.white),
        ('LEFTPADDING', (0, 0), (-1, -1), 6),
        ('RIGHTPADDING', (0, 0), (-1, -1), 6),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('ALIGN',(0,0), (-1,-1),'LEFT')
    ]))
    main_elements.append(Test_result_table)
    # test appended statements for exceptions 
    # test AIIC/ASTC result large text box 
    # take relevant variables for AIIC and ASTC tests and calculate the results, then insert them into the text box.
    if test_type == 'AIIC':
        main_elements.append(Paragraph("The Apparent Impact Insulation Class (AIIC) was calculated. The AIIC rating is based on Apparent Transmission Loss (ATL), and includes the effects of noise flanking. The AIIC reference contour is shown on the next page, and has been “fit” to the Apparent Transmission Loss values, in accordance with the procedure of "+standards_data[0][0]))
    elif test_type == 'ASTC':
        main_elements.append(Paragraph("The Apparent Sound Transmission Class (ASTC) was calculated. The ASTC rating is based on Apparent Transmission Loss (ATL), and includes the effects of noise flanking. The ASTC reference contour is shown on the next page, and has been “fit” to the Apparent Transmission Loss values, in accordance with the procedure of "+standards_data[0][0]))
    elif test_type == 'NIC':
        main_elements.append(Paragraph("The Noise Isolation Class (NIC) was calculated. The NIC rating is based on Noise Reduction (NR), and includes the effects of noise flanking. The NIC reference contour is shown on the next page, and has been “fit” to the Apparent Transmission Loss values, in accordance with the procedure of "+standards_data[0][0]))
    main_elements.append(Paragraph("The results stated in this report represent only the specific construction and acoustical conditions present at the time of the test. Measurements performed in accordance with this test method on nominally identical constructions and acoustical conditions may produce different results."))

    main_elements.append(PageBreak())
    ####### 3rd page - ASTC/NIC reference contour plot

    # need proper formatting for this plot.
    if test_type == 'AIIC':
       IIC_yAxis = 'Sound Pressure Level (dB)'
       plot_img = plot_curves(frequencies,IIC_yAxis, IIC_contour_final,Nrec_ANISPL,Ref_label, Field_IIC_label)
       main_elements.append(plot_img)

    elif test_type == 'ASTC':
        ASTCyAxis = 'Transmission Loss (dB)'
        plot_img = plot_curves(frequencies, ASTCyAxis, ASTC_contour_final,Nrec_ANISPL,Ref_label, Field_IIC_label)
        main_elements.append(plot_img)
    elif test_type == 'NIC':
        plot_title = 'NIC Reference Contour'
        plt.plot(ATL_curve, freqbands)
        plt.xlabel('Apparent Transmission Loss (dB)')
        plt.ylabel('Frequency (Hz)')
        plt.title('NIC Reference Contour')
        plt.grid()
        plt.show()


    # Output a file string for the PDF made up of test number and test type

    # Build the document
    doc.build(main_elements)

