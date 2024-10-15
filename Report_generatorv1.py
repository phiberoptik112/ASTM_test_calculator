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
import time
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
from IPython.display import display, Image
from matplotlib import ticker
import os

# config file with constants
from config import *
from data_processor import *

def sanitize_filepath(filepath):
    ##"""Sanitize a file path by replacing forward slashes with backslashes."""
    filepath = filepath.replace('T:', '//DLA-04/Shared/')
    filepath = filepath.replace('\\','/')
    # need to add a line to append a / at the end of the filename
    return filepath

from reportlab.lib import colors
from reportlab.platypus import Table, TableStyle
def create_tables(data, col_width=None, row_heights=None, style=None):
    table = Table(data, colWidths=col_width, rowHeights=row_heights)

    if style is None:
        style = TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 14),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('TEXTWRAP', (0, 1), (-1, -1)),
            ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
            ('ALIGN', (0, 1), (-1, -1), 'LEFT'),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 12),
            ('BOTTOMPADDING', (0, 1), (-1, -1), 12),
            ('TOPPADDING', (0, 1), (-1, -1), 12),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ])
    table.setStyle(style)
    return table

# #### database has raw OBA datasheet, needs to be cleaned for plotting

### DATA CALC FUNCTIONS ###

# ATL calc, reworked and funtional - use Recieve room recording, not the tapper level for this calc. this does not take the tapper into account and is just used for ASTC 
# def calc_ATL_val(srs_overalloct,rec_overalloct,bkgrnd_overalloct,parition_area,recieve_roomvol,sabines):
#     ASTC_vollimit = 883
#     if recieve_roomvol > ASTC_vollimit:
#         print('Using NIC calc, room volume too large')
#     recieve_corr = list()
#     bkgrnd_overalloct = pd.to_numeric(bkgrnd_overalloct, errors='coerce')
#     bkgrnd_overalloct = np.array(bkgrnd_overalloct)
#     rec_overalloct = pd.to_numeric(rec_overalloct, errors='coerce')
#     rec_overalloct = np.array(rec_overalloct)
    
#     recieve_vsBkgrnd = rec_overalloct - bkgrnd_overalloct
#     print('recieve vs background:',recieve_vsBkgrnd)
#     recieve_vsBkgrnd = pd.to_numeric(recieve_vsBkgrnd, errors='coerce')
#     recieve_vsBkgrnd = np.array(recieve_vsBkgrnd)

#     for i, val in enumerate(recieve_vsBkgrnd):
#         if val < 5:
#             recieve_corr.append(rec_overalloct[i]-2)
#         elif val < 10:
#             recieve_corr.append(10*np.log10(10**(rec_overalloct[i]/10)-10**(bkgrnd_overalloct[i]/10)))
#         else:
#             recieve_corr.append(rec_overalloct[i])
#     recieve_corr = np.round(recieve_corr,1)
#     print('corrected recieve: ',recieve_corr)
    
#     sabines = pd.to_numeric(sabines, errors='coerce')
#     #convert sabines to array
#     sabines = np.array(sabines)
#     print('sabines: ',sabines)
#     srs_overalloct = pd.to_numeric(srs_overalloct, errors='coerce')
#     srs_overalloct = np.array(srs_overalloct)
#     print('srs_overalloct: ',srs_overalloct)
#     ATL_val = srs_overalloct - recieve_corr+10*(np.log10(parition_area/sabines))
#     return ATL_val

### this code revised 7/24/24 - functional and produces accurate ATL values
#  moving to data_processor.py

# def calc_ATL_val(srs_overalloct,rec_overalloct,bkgrnd_overalloct,rt_thirty,parition_area,recieve_roomvol):
#     ASTC_vollimit = 883
#     if recieve_roomvol > ASTC_vollimit:
#         print('Using NIC calc, room volume too large')
#     # constant = np.int32(20.047*np.sqrt(273.15+20))
#     # intermed = 30/rt_thirty ## why did i do this? not right....sabines calc is off
#     # thisval = np.int32(recieve_roomvol*intermed)
#     # sabines =thisval/constant

#     # RT value is right, why is this not working?
#     # print('recieve roomvol: ',recieve_roomvol)
#     if isinstance(rt_thirty, pd.DataFrame):
#         rt_thirty = rt_thirty.values
#     # print('rt_thirty: ',rt_thirty)
    
#     sabines = 0.049*recieve_roomvol/rt_thirty  # this produces accurate sabines values

#     if isinstance(bkgrnd_overalloct, pd.DataFrame):
#         bkgrnd_overalloct = bkgrnd_overalloct.values
#     # print('bkgrnd_overalloct: ',bkgrnd_overalloct)
#     # sabines = np.int32(sabines) ## something not right with this calc
#     sabines = np.round(sabines)
#     # print('sabines: ',sabines)
    
#     recieve_corr = list()
#     # print('recieve: ',rec_overalloct)
#     recieve_vsBkgrnd = rec_overalloct - bkgrnd_overalloct

#     if isinstance(recieve_vsBkgrnd, pd.DataFrame):
#         recieve_vsBkgrnd = recieve_vsBkgrnd.values

#     if isinstance(rec_overalloct, pd.DataFrame):
#         rec_overalloct = rec_overalloct.values
#     # print('recieve vs  background: ',recieve_vsBkgrnd)
#     # print('recieve roomvol: ',recieve_roomvol)
#     #### something wrong with this loop #### 
#     for i, val in enumerate(recieve_vsBkgrnd):
#         if val < 5:
#             # print('recieve vs background: ',val)
#             recieve_corr.append(rec_overalloct[i]-2)
#             # print('less than 5, appending: ',recieve_corr[i])
#         elif val < 10:
#             # print('recieve vs background: ',val)
#             recieve_corr.append(10*np.log10((10**(rec_overalloct[i]/10))-(10**(bkgrnd_overalloct[i]/10))))
#             # print('less than 10, appending: ',recieve_corr[i])
#         else:
#             # print('recieve vs background: ',val)
#             recieve_corr.append(rec_overalloct[i])
#             # print('greater than 10, appending: ',recieve_corr[i])
            
    
#     # print('recieve correction: ',recieve_corr)
#     if isinstance(srs_overalloct, pd.DataFrame):
#         recieve_corr = recieve_corr.values
#     if isinstance(srs_overalloct, pd.DataFrame):
#         srs_overalloct = srs_overalloct.values
#     if isinstance(sabines, pd.DataFrame):
#         sabines = sabines.values
#     # print('srs overalloct: ',srs_overalloct)
#     # print('recieve correction: ',recieve_corr)
#     # print('sabines: ',sabines)
#     ATL_val = []
#     for i, val in enumerate(srs_overalloct):
#         ATL_val.append(srs_overalloct[i]-recieve_corr[i]+10*(np.log10(parition_area/sabines.iloc[i])))
#     # ATL_val = srs_overalloct - recieve_corr+10*(np.log(parition_area/sabines)) 
#     ATL_val = np.round(ATL_val,1)
#     # print('ATL val: ',ATL_val)
#     return ATL_val

## functional. need to verify with excel doc calcs
# # moving to data_processor.py
# def calc_AIIC_val(Normalized_recieve_IIC):
#     pos_diffs = list()
#     AIIC_start = 94
#     AIIC_contour_val = 16
#     IIC_contour = list()
#     new_sum = 0
#     diff_negative_max = 0
#     IIC_curve = [2,2,2,2,2,2,1,0,-1,-2,-3,-6,-9,-12,-15,-18]
#     # initial application of the IIC curve to the first AIIC start value 
#     for vals in IIC_curve:
#         IIC_contour.append(vals+AIIC_start)
#     Contour_curve_result = IIC_contour - Normalized_recieve_IIC
#     print('Normalized recieve ANISPL: ', Normalized_recieve_IIC)

#     while (diff_negative_max < 8 and new_sum < 32):
#         print('Inside loop, current AIIC contour: ', AIIC_contour_val)
#         print('Contour curve (IIC curve minus ANISPL): ',Contour_curve_result)
        
#         diff_negative =  Normalized_recieve_IIC-IIC_contour
#         print('diff negative: ', diff_negative)

#         diff_negative_max =  np.max(diff_negative)
#         diff_negative = pd.to_numeric(diff_negative, errors='coerce')
#         diff_negative = np.array(diff_negative)

#         print('Max, single diff: ', diff_negative_max)
#         for val in diff_negative:
#             if val > 0:
#                 pos_diffs.append(np.round(val))
#             else:
#                 pos_diffs.append(0)
#         print('positive diffs: ',pos_diffs)
#         new_sum = np.sum(pos_diffs)
#         print('Sum Positive diffs: ', new_sum)
#         print('Evaluating sums and differences vs 32, 8: ', new_sum, diff_negative_max)
#         if new_sum > 32 or diff_negative_max > 8:
#             print('Difference condition met! AIIC value: ', AIIC_contour_val) # 
#             print('AIIC result curve: ', Contour_curve_result)
#             return AIIC_contour_val, Contour_curve_result
#         # condition not met, resetting arrays
#         pos_diffs = []
#         IIC_contour = []
#         print('difference condition not met, subtracting 1 from AIIC start and recalculating the IIC contour')
#         AIIC_start = AIIC_start - 1
#         AIIC_contour_val = AIIC_contour_val + 1
#         print('AIIC start: ', AIIC_start)
#         print('AIIC contour value: ', AIIC_contour_val)
#         for vals in IIC_curve:
#             IIC_contour.append(vals+AIIC_start)
#         Contour_curve_result = IIC_contour - Normalized_recieve_IIC
#         if AIIC_start <10: break


# need to write a big overall wrapper function here that will go through the testplan, depending on what test type, pull the data from each excel file, write the data to the database variables for that test and then use those database variables to plot the data in the report.
# for each testplan entry, first write the room properties database: 
# Organize the variables into a dictionary

#     recieve_vsBkgrnd = np.round(recieve_vsBkgrnd,1)
    
#     print('rec vs background:',recieve_vsBkgrnd)
#     if testtype == 'AIIC':
#         for i, val in enumerate(recieve_vsBkgrnd):
#             if val < 5:
#                 recieve_corr.append(rec_overalloct[i]-2)
#             else:
#                 recieve_corr.append(10*np.log10(10**(rec_overalloct[i]/10)-10**(bkgrnd_overalloct[i]/10)))   
#     elif testtype == 'ASTC':
#         for i, val in enumerate(recieve_vsBkgrnd):
#             # print('val:', val)
#             # print('count: ', i)
#             if val < 5:
#                 recieve_corr.append(rec_overalloct.iloc[i]-2)
#             elif val < 10:
#                 recieve_corr.append(10*np.log10(10**(rec_overalloct[i]/10)-10**(bkgrnd_overalloct[i]/10)))
#             else:
#                 recieve_corr.append(rec_overalloct[i])
#         # print('-=-=-=-=-')
#         # print('recieve_corr: ',recieve_cor
###################################
####_+#_+_+_#+_#+_#+_####_######################################
## do i put the GUI here YES - ASTC GUI proto.py works - just need to add in the reportlab formatting and get going.
# wrap the following into a funciton - take in the testplan entry, and return the reportlab formatted pdf output path
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import BaseDocTemplate, PageTemplate, Frame, Paragraph, Table, TableStyle, Spacer,PageBreak, KeepInFrame, Image
from reportlab.pdfgen import canvas

from reportlab.lib.units import inch

# def create_report(self,curr_test, single_test_dataframe, reportOutputfolder,test_type):
        # all the code below
    # Kaulu by gentry testing ## EXAMPLE DATA CREATE LOOP FOR EACH TESTPLAN ENTRY

##### _+_+#__#+_#+_+_+_####################### REPORT GENERATION CODE ##############################

    ######## =--=-=-=--= 
    ### custom margins
    # Specify custom margins (in points, where 1 inch equals 72 points)
    # left_margin = right_margin = 0.75 * 72  # 0.75 inches
    # top_margin = 0.25 * 72 # 0.25 inch
    # bottom_margin = 1 * 72  # 1 inch
    # header_height = 2 * inch
    # footer_height = 0.5 * inch
    # main_content_height = letter[1] - top_margin - bottom_margin - header_height - footer_height

    # # Create a document with custom page templates
    # doc = BaseDocTemplate("report_with_header.pdf", pagesize=letter)

    #### -=-=--=-= 
def header_elements(self, single_test_dataframe, test_type):
    # Document Setup 
    # document name and page size
    #  Example ouptut file name: '24-006 AIIC Test Report_1.1.1.pdf'
    #  format: project_name + test_type + test_num + '.pdf'
    # doc_name = f"{single_test_dataframe['room_properties']['Project_Name'][0]} {test_type} Test Report_{single_test_dataframe['room_properties']['Test_Label'][0]}.pdf"
    doc_name = f"{single_test_dataframe['room_properties']['Project_Name'][0]} {test_type} Test Report_{single_test_dataframe['room_properties']['Test_Label'][0]}.pdf"

    
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

    # Define header elements

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
    elements.append(Paragraph('Test site: ' + single_test_dataframe['room_properties']['Site_Name'][0], styles['Normal']))
    elements.append(Spacer(1, 5))
    elements.append(Paragraph('Client: ' + single_test_dataframe['room_properties']['Client_Name'][0], styles['Normal']))
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
    # page_template = PageTemplate(id='Standard', frames=[main_frame, header_frame, footer_frame], onPage=header_footer)
    # doc.addPageTemplates([page_template])
    # # Main document content (Add your main document elements here)
    # main_elements = []
def create_standards_section(test_type, single_test_dataframe, main_elements):
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
    return main_elements,styleHeading

def create_test_environment_section(test_type,single_test_dataframe, main_elements):
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

def create_calc_table_section(self, single_test_dataframe, test_type):
    main_elements.append(Paragraph("<u>STATEMENT OF TEST RESULTS:</u>", styleHeading))
    #### Main calculation table section --- also split into AIIC, ASTC, NIC report results. 
    # function returning a database table to display frequency, level, OBA of source, reciever, RT60, NR, ATL, Exceptions
    ######### CALCS FOR PDF TABLES AND PLOTS WILL GO IN THIS SCRIPT ###########
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

        calc_NR, sabines, corrected_recieve,Nrec_ANISPL = calc_NR_new(onethird_srs, onethird_rec_Total, onethird_bkgrd, rt_thirty,single_test_dataframe['room_properties']['receive_room_vol'][0],NIC_vollimit=883,testtype='AIIC')
        
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
        ## obtain SLM data from overall dataframe
        onethird_rec = format_SLMdata(single_test_dataframe['recive_data'])
        onethird_srs = format_SLMdata(single_test_dataframe['srs_data'])
        onethird_bkgrd = format_SLMdata(single_test_dataframe['bkgrnd_data'])
        rt_thirty = single_test_dataframe['rt']['Unnamed: 10'][25:41]/1000
        # Calculation of ATL
        ATL_val,corrected_STC_recieve = calc_ATL_val(onethird_srs, onethird_rec, onethird_bkgrd,single_test_dataframe['room_properties']['partition_area'][0],single_test_dataframe['room_properties']['receive_room_vol'][0],sabines)
        # Calculation of NR
        calc_NR, sabines, corrected_recieve,Nrec_ANISPL = calc_NR_new(onethird_srs, onethird_rec, onethird_bkgrd, rt_thirty,single_test_dataframe['room_properties']['receive_room_vol'][0],NIC_vollimit=883,testtype='ASTC')
        # creating reference curve for ASTC graph
        STCCurve = [-16, -13, -10, -7, -4, -1, 0, 1, 2, 3, 4, 4, 4, 4, 4, 4]
        ASTC_contour_final = list()
        for vals in STCCurve:
            ASTC_contour_final.append(vals+(ATL_val))
        
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
                "Frequency": frequencies,
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
def create_plot_page_section(test_type, main_elements, frequencies, IIC_contour_final, ASTC_contour_final, Nrec_ANISPL, Ref_label, Field_IIC_label):
    main_elements.append(PageBreak())
    ####### 3rd page - ASTC/NIC reference contour plot

    # need proper formatting for this plot.
    if test_type == 'AIIC':
        IIC_yAxis = 'Sound Pressure Level (dB)'
        AIIC_plot_img = plot_curves(frequencies,IIC_yAxis, IIC_contour_final,Nrec_ANISPL,Ref_label, Field_IIC_label)
        plot_image = Image(AIIC_plot_img, 6*inch, 4*inch)
        main_elements.append(plot_image)
        main_elements.append(Spacer(1, 10))
    # create a text box with the AIIC single value result, contained in 

    elif test_type == 'ASTC':
        ASTCyAxis = 'Transmission Loss (dB)'
        ASTC_plot_img = plot_curves(frequencies, ASTCyAxis, ASTC_contour_final,Nrec_ANISPL,Ref_label, Field_IIC_label)
        plot_img = Image(ASTC_plot_img, 6*inch, 4*inch)
        main_elements.append(plot_img)
        main_elements.append(Spacer(1, 10))


    elif test_type == 'NIC':
        ASTCyAxis = 'Transmission Loss (dB)'
        NIC_plot_img = plot_curves(frequencies, ASTCyAxis, ASTC_contour_final,Nrec_ANISPL,Ref_label, Field_IIC_label)
        plot_img = Image(NIC_plot_img, 6*inch, 4*inch)
        main_elements.append(plot_img)
        main_elements.append(Spacer(1, 10))


        return main_elements

#  not sure if this function will go here, but it will be the main function that is called to create the report. 
## still need to sync up with the changes that may come from working through how best to iterate through each test and all its reports

# def create_report(test_data: TestData, report_output_folder: str, test_type: str):
#     page_template = PageTemplate(id='Standard', frames=[main_frame, header_frame, footer_frame], onPage=header_footer)
#     doc.addPageTemplates([page_template])
#     main_elements = []

#     doc = BaseDocTemplate(...)
#     elements = []
#     elements.extend(create_header(test_type, test_data))
#     elements.extend(create_standards_section(test_type))
#     elements.extend(create_test_environment_section(test_data))
#     # ... other sections ...
#     # doc.build(elements)
#         # Output a file string for the PDF made up of test number and test type
#     output_file = f"{reportOutputfolder}/Report_{curr_test}_{test_type}.pdf"

#         # Build the document
#     doc.build(main_elements)

#         # Save the document as a PDF
#     doc.save(output_file)
