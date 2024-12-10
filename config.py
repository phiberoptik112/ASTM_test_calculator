# Config file for standards and stock notes for the report generator
import pandas as pd
from reportlab.platypus import Paragraph, Spacer, Table, TableStyle, KeepInFrame, PageBreak 
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

test_instrumentation_table = [["Equipment Type","Manufacturer","Model Number","Serial Number",Paragraph("Last NIST Traceable Calibration"),Paragraph("Last Local Calibration")],
["Tapping Machine:","Norsonics","CAL200","2775671","9/19/2022","N/A"],
["Sound Level Meter","Larson Davis","831","4328","10/24/2022","4/4/2024"],
["Microphone Pre-Amp","Larson Davis","PRM831","046469","10/24/2022","4/4/2024"],
["Microphone","Larson Davis","377B20","168830","10/20/2022","4/4/2024"],
["Calibrator","Larson Davis","CAL200","5955","10/26/2022","N/A"],
["Amplified Loudspeaker","QSC","K10","GAA530909","N/A","N/A"],
["Noise Generator","NTi Audio","MR-PRO","0162","N/A","N/A"]
]

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

## equipment list # may change to something that can be edited in the GUI

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

## STATEMENT OF TEST RESULTS 
statement_test_results_text =' STATEMENT OF TEST RESULTS: '
FREQUENCIES = [125, 160, 200, 250, 315, 400, 500, 630, 800, 1000, 1250, 1600, 2000, 2500, 3150, 4000]
NIC_VOLLIMIT = 883  # cu. ft.
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch

#### reportlab constants #####
left_margin = right_margin = 0.75 * 72  # 0.75 inches
top_margin = 0.25 * 72 # 0.25 inch
bottom_margin = 1 * 72  # 1 inch
header_height = 2 * inch
footer_height = 0.5 * inch
main_content_height = letter[1] - top_margin - bottom_margin - header_height - footer_height

# _=-=-=-=--=-=-=-_+_=-=-=-=-=-=-_+_=-=-=-=-=-= Text lookups for report print -=-=-=-=-=-=-=- 

# stockNIC_note = ["The receiver and/or source room had a volume exceeding 150 m3 (5,300 cu. ft.), and the absorption of the receiver and/or source room was greater than the maximum allowed per E336-16, Paragraph 9.4.1.2.",
#                  "The receiver and/or source room was not an enclosed space.",
#                  "The receiver and/or source room has a volume less than the minimum volume requirement of 25 m3 (883 cu. ft.).",
#                  "The receiver and/or source room has one or more dimensions less than the minimum requirement of 2.3 m (7.5 ft.)."]

# ISR_ony_report = "The receiver room had a volume less than the minimum volume requirement of 40 m3."
# stockISR_notes = ["The receiver and/or source room had a volume exceeding 150 m3 (5,300 cu. ft.), and the absorption of the receiver room was greater than the maximum allowed per E1007-16, Paragraph 10.3.1 and 10.4.5.",
# "The receiver and/or source room was not an enclosed space.", 
# "The receiver and/or source room has a volume less than the minimum volume requirement of 40 m3 (1413 cu. ft.).",
# "The receiver and/or source room has one or more dimensions less than the minimum requirement of 2.3 m (7.5 ft.)."]

# #standards 
# standards_text = (("ASTC Test Procedure ASTM E336-16",	"Standard Test Method for Measurement of Airborne Sound Attenuation between Rooms in Buildings"),("STC Calculation	ASTM E413-16",	"Classification for Rating Sound Insulation"),("AIIC Test Procedure	ASTM E1007-14",	"Standard Test Method for Field Measurement of Tapping Machine Impact Sound Transmission Through Floor-Ceiling Assemblies and Associated Support Structure"),
# ("IIC Calculation	ASTM E989-06(2012)",	"Standard Classification for Determination of Impact Insulation Class (IIC)"),
# ("RT60 Test Procedure	ASTM E2235-04(2012)",	"Standard Test Method for Determination of Decay Rates for Use in Sound Insulation Test Methods"))

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

# # test procedure 
# test_procedure_pg = 'Determination of space-average sound pressure levels was performed via the manually scanned microphones techique, described in ' + standards_text[0][0] + ', Paragraph 11.4.3.3.'+ 'The source room was selected in accordance with ASTM E336-11 Paragraph 9.2.5, which states that "If a corridor must be used as one of the spaces for measurement of ATL or FTL, it shall be used as the source space."'
# # code:
# # =CONCATENATE("The test was performaned in general accordance with ",AIIC or ASTC or NIC,". Determination of Space-Average Levels performed via the manually scanned microphones techique, described in ",'SLM Data'!C59,", Paragraph 11.4.2.2.")
# # The test was performaned in general accordance with ASTM E1007-14. Determination of Space-Average Levels performed via the manually scanned microphones techique, described in ASTM E1007-14, Paragraph 11.4.2.2.								

# flanking_text = "Flanking transmission was not evaluated."

# # To evaluate room absorption, 1 microphone was used to measure 4 decays at 4 locations around the receiving room for a total of 16 measurements, per ASTM E2235-04(2012).								

# RT_text = "To evaluate room absorption, 1 microphone was used to measure 4 decays at 4 locations around the receiving room for a total of 16 measurements, per"+standards_text[4][0]

# # ASTC and NIC final result and blurb

# # ASTC result, concat with this: 
# ASTC_results_ATverbage = ' was calculated. The ASTC rating is based on Apparent Transmission Loss (ATL), and includes the effects of noise flanking. The ASTC reference contour is shown on the next page, and has been “fit” to the Apparent Transmission Loss values, in accordance with the procedure of'

# #NIC results, concat with this:
# NIC_results_NRverbage = ' was calculated. The NIC rating is based on Noise Reduction (NR), and includes the effects of noise flanking. The NIC reference contour is shown on the next page, and has been “fit” to the Apparent Transmission Loss values, in accordance with the procedure of'
# # after results
# results_blurb = 'The results stated in this report represent only the specific construction and acoustical conditions present at the time of the test. Measurements performed in accordance with this test method on nominally identical constructions and acoustical conditions may produce different results.'

# #test instrumentation
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


# ## STATEMENT OF TEST RESULTS 
# statement_test_results_text =' STATEMENT OF TEST RESULTS: '

