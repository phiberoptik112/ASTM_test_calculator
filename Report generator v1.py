# Report generator v1

import pandas as pd
from pandas import ExcelWriter
import numpy as np
from openpyxl import load_workbook
from openpyxl import Workbook
from openpyxl import cell
from openpyxl.utils import get_column_letter
import xlsxwriter

def sanitize_filepath(filepath):
    ##"""Sanitize a file path by replacing forward slashes with backslashes."""
    filepath = filepath.replace('T:', '//DLA-04/Shared/')
    filepath = filepath.replace('\\','/')
    # need to add a line to append a / at the end of the filename
    return filepath

#### ADDITION TO MAKE IT METER LETTER AGNOSTIC - METER 1 and 2 #### STILL NEED TO DO THIS

# writing meter data to report file function definition
def write_testdata(self,find_datafile, reportfile, newsheetname):
    rawDtestpath = self.slm_data_d_path
    rawEtestpath = self.slm_data_e_path
    rawAtestpath = self.slm_data_e_path
    rawReportpath = self.report_output_folder_path
    D_datafiles = [f for f in listdir(rawDtestpath) if isfile(join(rawDtestpath,f))]
    E_datafiles = [f for f in listdir(rawEtestpath) if isfile(join(rawEtestpath,f))]
    E_datafiles = [f for f in listdir(rawEtestpath) if isfile(join(rawEtestpath,f))]

    if find_datafile[0] =='A': ## REPLACE WITH METER 1 - NEED TO DEFINE IN TESTPLAN ##
        datafile_num = find_datafile[1:]
        datafile_num = '-831_Data.'+datafile_num+'.xlsx'
        slm_found = [x for x in A_datafiles if datafile_num in x]
        slm_found[0] = rawAtestpath+slm_found[0]# If this line errors, the test file is mislabled or doesn't exist 
        # print(srs_slm_found)
    elif find_datafile[0] == 'E':
        datafile_num = find_datafile[1:]
        datafile_num = '-831_Data.'+datafile_num+'.xlsx'
        slm_found = [x for x in E_datafiles if datafile_num in x]
        slm_found[0] = rawEtestpath+slm_found[0]# If this line errors, the test file is mislabled or doesn't exist 

    print(slm_found[0])

    srs_data = pd.read_excel(slm_found[0],sheet_name='OBA') # data must be in OBA tab
    with ExcelWriter(
    rawReportpath+reportfile,
    mode="a",
    engine="openpyxl",
    if_sheet_exists="replace",
    ) as writer:
        srs_data.to_excel(writer, sheet_name=newsheetname) #writes to report file
    time.sleep(1)
    # excel.Quit()

def write_RTtestdata(find_datafile, reportfile,newsheetname):
    rawDtestpath = self.slm_data_d_path
    rawEtestpath = self.slm_data_e_path
    rawReportpath = self.report_output_folder_path
    D_datafiles = [f for f in listdir(rawDtestpath) if isfile(join(rawDtestpath,f))]
    E_datafiles = [f for f in listdir(rawEtestpath) if isfile(join(rawEtestpath,f))]
    if find_datafile[0] =='A':
        datafile_num = find_datafile[1:]
        datafile_num = '-RT_Data.'+datafile_num+'.xlsx'
        slm_found = [x for x in A_datafiles if datafile_num in x]
        slm_found[0] = rawAtestpath+slm_found[0]# If this line errors, the test file is mislabled or doesn't exist 
        # print(srs_slm_found)
    elif find_datafile[0] == 'E':
        datafile_num = find_datafile[1:]
        datafile_num = '-RT_Data.'+datafile_num+'.xlsx'
        slm_found = [x for x in E_datafiles if datafile_num in x]
        slm_found[0] = rawEtestpath+slm_found[0] # If this line errors, the test file is mislabled or doesn't exist 

    print(slm_found[0])

    srs_data = pd.read_excel(slm_found[0],sheet_name='Summary')# data must be in Summary tab for RT meas.
    # could reduce this function by also passing the sheet to be read into the args. 
    with ExcelWriter(
    rawReportpath+reportfile,
    mode="a",
    engine="openpyxl",
    if_sheet_exists="replace",
    ) as writer:
        srs_data.to_excel(writer, sheet_name=newsheetname) 
    time.sleep(1)
    # excel.Quit()

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

## do i put the GUI here YES - ASTC GUI proto.py works - just need to add in the reportlab formatting and get going.

from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.platypus import BaseDocTemplate, PageTemplate, Frame, Paragraph, Table, TableStyle, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

source_room_name = "2nd Floor Bed 3"
rec_roomName = "1st Floor Bed 3"
sitename = "Ka'ulu by Gentry"
client_Name = "Gentry Builders, LLC"
reportdate = "4-24-24"
testdate = "4-3-24"
testnum = '1.1.1' # this will be the pulled var from the testplan
rec_vol = "1441"
source_vol = "3643"
tested_assem = "partition"
partition_area = "1444"
partition_dim = "12x14"
source_rm_floor = 'LVT'
source_rm_walls = 'gyp'
source_rm_finish = 'unfinished'
source_rm_ceiling = 'gyp'
receiver_rm_floor = 'LVT'
receiver_rm_walls = 'gyp'
receiver_rm_finish = 'unfinished'
receiver_rm_ceiling = 'gyp'
test_assem_type = 'Floor-ceiling' ## will change AIIC vs ASTC
receiver_room_name = '1st floor great room/kitchen'
### custom margins
# Specify custom margins (in points, where 1 inch equals 72 points)
left_margin = 0.75 * 72  # 0.75 inches
right_margin = 0.75 * 72  # 0.75 inches
top_margin = 1 * 72  # 1 inch
bottom_margin = 1 * 72  # 1 inch


# Function to create PDF report
def create_pdf_report(file_name, left_data, right_data, client_data):
    doc = SimpleDocTemplate(file_name, pagesize=letter,
    leftMargin=left_margin,
    rightMargin=right_margin,
    topMargin=top_margin,
    bottomMargin=bottom_margin)
    elements = []

    # Create a table for the left data
    table_left = Table(left_data)
    table_left.setStyle(TableStyle([('GRID', (0, 0), (-1, -1), 1, colors.white)]))

    # Create a table for the right data
    table_right = Table(right_data)
    table_right.setStyle(TableStyle([('GRID', (0, 0), (-1, -1), 1, colors.white)]))

    # Create a table for the client data
    client_table = Table(client_data)
    client_table.setStyle(TableStyle([('GRID', (0, 0), (-1, -1), 1, colors.white)]))

    # Combine left and right tables into a single table for side by side display
    table_combined_lr = Table([[table_left, table_right]])
    # tables will be centered unless their style has been defined via column widths
    elements.append(table_combined_lr)
    elements.append(Spacer(1, 10))  # Adds some space 
    # elements.append(client_table)
    elements.append(Paragraph('Test site:'+sitename))
    elements.append(Paragraph('Client:'+client_Name))

    styles = getSampleStyleSheet()
    styleNormal = styles['Normal']
    styleHeading = ParagraphStyle('heading', parent=styles['Normal'], spaceAfter=10)
    elements.append(Spacer(1, 10))  # Adds some space 
    ## -=-==-= Heading 'STANDARDS'  # -=-=-=-=-=--=-=-=-=-=--=-=-
    elements.append(Paragraph('STANDARDS:', styleHeading))

    # Standards table data
    standards_data = [
        ['ASTM E1007-14', Paragraph('Standard Test Method for Field Measurement of Tapping Machine Impact Sound Transmission Through Floor-Ceiling Assemblies and Associated Support Structure',styles['Normal'])],
        ['ASTM E989-06(2012)', Paragraph('Standard Classification for Determination of Impact Insulation Class (IIC)',styles['Normal'])],
        ['ASTM E2235-04(2012)', Paragraph('Standard Test Method for Determination of Decay Rates for Use in Sound Insulation Test Methods',styles['Normal'])]
    ]
    # Define column widths
    col_widths = [doc.width/2.0, doc.width/2.0]  # for example, divide the available width into half

    # Create the table
    standards_table = Table(standards_data, colWidths=col_widths)
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
    elements.append(standards_table)

    # Heading 'TEST ENVIRONMENT'
    elements.append(Paragraph('TEST ENVIRONMENT:', styleHeading))
    elements.append(Paragraph('The source room was '+source_room_name+'. The space was'+source_rm_finish+'. The floor was '+source_rm_floor+'. The ceiling was '+source_rm_ceiling+". The walls were"+source_rm_walls+". All doors and windows were closed during the testing period. The source room had a volume of approximately"+source_vol+"cu. ft."))
    elements.append(Spacer(1, 10))  # Adds some space 
    ### Recieve room paragraph
    elements.append(Paragraph('The receiver room was '+receiver_room_name+'. The space was'+receiver_rm_finish+'. The floor was '+receiver_rm_floor+'. The ceiling was '+receiver_rm_ceiling+". The walls were"+receiver_rm_walls+". All doors and windows were closed during the testing period. The source room had a volume of approximately"+rec_vol+"cu. ft."))
    elements.append(Spacer(1, 10))  # Adds some space 
    elements.append(Paragraph('The test assembly measured approximately '+partition_dim+", and had an area of approximately "+partition_area+"sq. ft."))
    elements.append(Spacer(1, 10))  # Adds some space 
    # Heading 'TEST ENVIRONMENT'
    elements.append(Paragraph('TEST ASSEMBLY:', styleHeading))
    elements.append(Spacer(1, 10))  # Adds some space 
    elements.append(Paragraph("The tested assembly was the"+test_assem_type+"The assembly was not field verified, and was based on information provided by the client and drawings for the project. The client advised that no slab treatment or self-leveling was applied. Results may vary if slab treatment or self-leveling or any adhesive is used in other installations."))
    # END OF FIRST PAGE TEXT  - FOOTER TO COME ##-=-=-=-=-=-==-     


    # Build the document
    doc.build(elements)

# Data for the tables (example data)
leftside_Data = [
    ["Report Date:", "4-24-24"],
    ['Test Date:', "4-3-24"],
    ['DLAA Test No', '1.1.1']
]

rightside_Data = [
    ["Source Room:", "2nd Floor Bed 3, Volume: 3643 cu. ft."],
    ["Receiver Room:", "1st Floor Bed 3, Volume: 1441 cu. ft."],
    ["Test Assembly:", "Floor-ceiling, Area "+partition_area+"sq. ft."]
]

clienttable = [
    ['Test site:', "Ka'ulu by Gentry"],
    ["Client name:", "Gentry Builders, LLC"]
]
# variable names 
# rightside_Data = [["Source Room:", source_room_name+source_vol],
# ["Receiver Room:", rec_roomName+rec_vol],
# ["Test Assembly:", tested_assem]]
# leftside_Data = [
#     ["Report Date:", reportdate],
#     ['Test Date:', testdate],
#     ['DLAA Test No', testnum]
# ]

# clienttable = [
#     ['Test site:', sitename],
#     ["Client name:", client_Name]
# ]
# # Generate the PDF report
create_pdf_report("example_report.pdf", leftside_Data, rightside_Data, clienttable)