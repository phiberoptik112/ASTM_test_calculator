from reportlab.lib.units import inch
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.platypus import Paragraph, Spacer, Table, TableStyle, KeepInFrame, PageBreak
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.platypus import BaseDocTemplate, Frame, PageTemplate


from config import *
from data_processor import *

class BaseTestReport:
    def __init__(self, curr_test, test_data, reportOutputfolder):
        self.curr_test = curr_test
        self.test_data = test_data
        self.reportOutputfolder = reportOutputfolder
        self.doc = None
        self.styles = getSampleStyleSheet()
        self.custom_title_style = self.styles['Heading1']
        
        # Define margins and heights
        self.left_margin = inch
        self.right_margin = inch
        self.top_margin = inch
        self.bottom_margin = inch
        self.header_height = 2 * inch
        self.footer_height = 0.5 * inch

    def setup_document(self):
        doc_name = f"{self.reportOutputfolder}/{self.get_doc_name()}"
        self.doc = BaseDocTemplate(
            doc_name, 
            pagesize=letter,
            leftMargin=self.left_margin, 
            rightMargin=self.right_margin,
            topMargin=self.top_margin, 
            bottomMargin=self.bottom_margin
        )

        # Define frames
        self.header_frame = Frame(
            self.left_margin, 
            letter[1] - self.top_margin - self.header_height, 
            letter[0] - 2 * self.left_margin, 
            self.header_height, 
            id='header'
        )
        
        self.main_frame = Frame(
            self.left_margin, 
            self.bottom_margin + self.footer_height, 
            letter[0] - 2 * self.left_margin, 
            letter[1] - self.top_margin - self.header_height - self.footer_height - self.bottom_margin, 
            id='main'
        )
        
        self.footer_frame = Frame(
            self.left_margin, 
            self.bottom_margin, 
            letter[0] - 2 * self.left_margin, 
            self.footer_height, 
            id='footer'
        )

        # Create page template
        page_template = PageTemplate(
            id='Standard', 
            frames=[self.main_frame, self.header_frame, self.footer_frame], 
            onPage=self.header_footer
        )
        self.doc.addPageTemplates([page_template])
        return self.doc

    def header_elements(self):
        elements = []
        props = self.test_data.room_properties
        elements.append(Paragraph(self.get_report_title(), self.custom_title_style))
        elements.append(Spacer(1, 10))
        
        leftside_data = [
            ["Report Date:", props['ReportDate'][0]],
            ['Test Date:', props['Testdate'][0]],
            ['DLAA Test No', props['Test number'][0]]
        ]
        rightside_data = [
            ["Source Room:", props['Source Room Name'][0]],
            ["Receiver Room:", props['Recieve Room Name'][0]],
            ["Test Assembly:", props['Tested Assembly'][0]]
        ]

        table_left = Table(leftside_data)
        table_right = Table(rightside_data)
        table_left.setStyle(TableStyle([('GRID', (0, 0), (-1, -1), 1, colors.white)]))
        table_right.setStyle(TableStyle([('GRID', (0, 0), (-1, -1), 1, colors.white)]))

        table_combined_lr = Table([[table_left, table_right]], colWidths=[self.doc.width / 2.0] * 2)
        elements.append(KeepInFrame(maxWidth=self.doc.width, maxHeight=self.header_height, content=[table_combined_lr], hAlign='LEFT'))
        elements.append(Spacer(1, 10))
        elements.append(Paragraph('Test site: ' + props['Site_Name'][0], self.styles['Normal']))
        elements.append(Spacer(1, 5))
        elements.append(Paragraph('Client: ' + props['Client_Name'][0], self.styles['Normal']))
        return elements

    def header_footer(self, canvas, doc):
        canvas.saveState()
        
        # Build header
        self.header_frame._leftPadding = self.header_frame._rightPadding = 0
        header_story = self.header_elements()
        self.header_frame.addFromList(header_story, canvas)
        
        # Build footer
        canvas.setFont('Helvetica', 10)
        footer_text = f"Page {doc.page}"
        canvas.drawCentredString(
            letter[0] / 2, 
            self.bottom_margin + self.footer_height / 2, 
            footer_text
        )
        
        canvas.restoreState()


    def get_standards_data(self):
        raise NotImplementedError

    def get_test_procedure(self):
        raise NotImplementedError

    def get_test_instrumentation(self):
        common_equipment = [
            ['Sound Level Meter','Larson Davis','831','4328','10/24/2022','4/4/2024'],
            ['Microphone Pre-Amp','Larson Davis','PRM831','046469','10/24/2022','4/4/2024'],
            ['Microphone','Larson Davis','377B20','168830','10/20/2022','4/4/2024'],
            ['Calibrator','Larson Davis','CAL200','5955','10/26/2022','N/A'],
            ['Amplified Loudspeaker','QSC','K10','GAA530909','N/A','N/A'],
            ['Noise Generator','NTi Audio','MR-PRO','0162','N/A','N/A']
        ]
        return common_equipment

    def create_results_table(self, Test_result_table, styleHeading):
        main_elements = []
        main_elements.append(Paragraph("<u>STATEMENT OF TEST RESULTS:</u>", styleHeading))
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
        return Test_result_table, main_elements

    def create_plot(self):
        raise NotImplementedError

    def get_report_title(self):
        raise NotImplementedError
    
    @classmethod
    def create_report(cls, curr_test, test_data, reportOutputfolder, test_type):
        """Factory method to create and build the appropriate test report
        
        Args:
            curr_test: Current test information
            test_data: TestData instance containing all test measurements and properties
            reportOutputfolder: Output directory for the report
            test_type: Type of test (AIIC, ASTC, NIC, or DTC)
            
        Returns:
            BaseTestReport: The generated report instance
        """
        # Create the appropriate test report object based on test_type
        report_classes = {
            TestType.AIIC: AIICTestReport,
            TestType.ASTC: ASTCTestReport,
            TestType.NIC: NICTestReport,
            TestType.DTC: DTCTestReport
        }
        
        report_class = report_classes.get(test_type)
        if not report_class:
            raise ValueError(f"Unsupported test type: {test_type}")
            
        # Create report instance
        report = report_class(curr_test, test_data, reportOutputfolder)
        
        # Setup document
        doc = report.setup_document()

        # Generate content
        main_elements = []
        main_elements.extend(report.create_first_page())
        main_elements.extend(report.create_second_page())
        main_elements.extend(report.create_third_page())
        main_elements.extend(report.create_fourth_page())

        # Build and save document
        doc.build(main_elements)
        print(f"Report saved as: {report.get_doc_name()}")
        
        return report

    def create_first_page(self):
        main_elements = []
        ### STANDARDS ###
        styleHeading = ParagraphStyle('heading', parent=self.styles['Normal'], spaceAfter=10)
        main_elements.append(Paragraph('<u>STANDARDS:</u>', styleHeading))
        standards_table = Table(self.get_standards_data(), hAlign='LEFT')
        # ... (set table style)
        main_elements.append(standards_table)

        # statement of conformance
        main_elements.append(Paragraph("<u>STATEMENT OF CONFORMANCE:</u>", styleHeading))
        main_elements.append(Paragraph(self.get_statement_of_conformance()))

        main_elements.append(Paragraph(self.get_test_environment()))

        main_elements.append(self.get_test_assembly())
        return main_elements

    def create_second_page(self): 
        main_elements = []
        
        # Test Procedure
        main_elements.append(Paragraph("<u>TEST PROCEDURE:</u>", self.custom_title_style))
        main_elements.append(Paragraph(self.get_test_procedure()))
        main_elements.append(Spacer(1,10))
        
        # Test Instrumentation and Calibration
        main_elements.append(Paragraph("<u>TEST INSTRUMENTATION:</u>", self.custom_title_style))
        
        test_instrumentation_table = self.get_test_instrumentation()
        test_instrumentation_table = Table(test_instrumentation_table, hAlign='LEFT')
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
        
        main_elements.append(test_instrumentation_table)
        main_elements.append(PageBreak())
        
        return main_elements

    def create_third_page(self):
        main_elements = []
        main_elements.append(Paragraph("<u>STATEMENT OF TEST RESULTS:</u>", self.custom_title_style))
        Test_result_table = self.get_test_results()
        Test_result_table = Table(Test_result_table, hAlign='LEFT')
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

        main_elements.append(self.get_testres_table_notes())
        ### need to figure out a single text box that displays the IIC test result single number 
        test_type_labels = {
            TestType.AIIC: "AIIC",
            TestType.ASTC: "ASTC",
            TestType.NIC: "NIC",
            TestType.DTC: "DTC"
        }
        test_label = test_type_labels.get(self.curr_test.test_type, "Unknown")
        main_elements.append(Paragraph(f"<u>{test_label}:</u>", self.custom_title_style))
        ### need to debug - append the test calc result here
        main_elements.append(Spacer(1,10))
        main_elements.append(self.get_test_results_paragraph())
        main_elements.append(PageBreak())
        return main_elements

    def create_fourth_page(self):
        main_elements = []
        main_elements.append(self.create_plot())
        test_type_labels = {
            TestType.AIIC: "AIIC",
            TestType.ASTC: "ASTC",
            TestType.NIC: "NIC",
            TestType.DTC: "DTC"
        }
        test_label = test_type_labels.get(self.curr_test.test_type, "Unknown")
        main_elements.append(Paragraph(f"<u>{test_label}:</u>", self.custom_title_style))
        # #### sane here -single number result from the test calc
        main_elements.append(Spacer(1,10))
        main_elements.append(self.get_signatures())
        # main_elements.append(PageBreak())
        return main_elements

class AIICTestReport(BaseTestReport):
    def get_doc_name(self):
        props = self.test_data.room_properties
        return f"{props['Project_Name'][0]} AIIC Test Report_{props['Test_Label'][0]}.pdf"

    def get_standards_data(self):
        return [
            ['ASTM E1007-14', 'Standard Test Method for Field Measurement of Tapping Machine Impact Sound Transmission Through Floor-Ceiling Assemblies and Associated Support Structure'],
            ['ASTM E413-16', 'Standard Classification for Rating Sound Insulation'],
            ['ASTM E2235-04(2012)', 'Standard Test Method for Determination of Decay Rates for Use in Sound Insulation Test Methods'],
            ['ASTM E989-06(2012)', 'Standard Classification for Determination of Impact Insulation Class (IIC)']
        ]
    def get_statement_of_conformace(self):
        return "Testing was conducted in accordance with ASTM E1007-14, ASTM E413-16, ASTM E2235-04(2012), and ASTM E989-06(2012), with exceptions noted below. All requrements for measuring abd reporting Absorption Normalized Impact Sound Pressure Level (ANISPL) and Apparent Impact Insulation Class (AIIC) were met."
    
    def get_test_environment(self,styleHeading,):
        main_elements = []
        props = self.test_data.room_properties
        main_elements.append(Paragraph("<u>TEST ENVIRONMENT:</u>", styleHeading))
        main_elements.append(Paragraph('The source room was '+props['Source_Room'][0]+'. The space was'+props['source_room_finish'][0]+'. The floor was '+props['srs_floor'][0]+'. The ceiling was '+props['srs_ceiling'][0]+". The walls were"+props['srs_walls'][0]+". All doors and windows were closed during the testing period. The source room had a volume of approximately "+props['source_room_vol'][0]+"cu. ft."))
        main_elements.append(Spacer(1, 10))  # Adds some space 
        ### Recieve room paragraph
        main_elements.append(Paragraph('The receiver room was '+props['Receiving_Room'][0]+'. The space was'+props['receiver_room_finish'][0]+'. The floor was '+props['rec_floor'][0]+'. The ceiling was '+props['rec_ceiling'][0]+". The walls were"+props['rec_Wall'][0]+". All doors and windows were closed during the testing period. The source room had a volume of approximately "+props['receive_room_vol'][0]+"cu. ft."))
        main_elements.append(Spacer(1, 10))  # Adds some space 
        main_elements.append(Paragraph('The test assembly measured approximately '+props['partition_dim'][0]+", and had an area of approximately "+props['partition_area'][0]+"sq. ft."))
        main_elements.append(Spacer(1, 10))  # Adds some space 
        # Heading 'TEST ENVIRONMENT'
        main_elements.append(Paragraph("<u>TEST ASSEMBLY:</u>", styleHeading))
        main_elements.append(Spacer(1, 10))  # Adds some space 
        main_elements.append(Paragraph("The tested assembly was the"+props['Test_Assembly_Type'][0]+"The assembly was not field verified, and was based on information provided by the client and drawings for the project. The client advised that no slab treatment or self-leveling was applied. Results may vary if slab treatment or self-leveling or any adhesive is used in other installations."))
        # ##### END OF FIRST PAGE TEXT  - ########
        #   main_elements.append(PageBreak())
        return main_elements
    

    def get_test_instrumentation(self):
        equipment = super().get_test_instrumentation()
        # Add AIIC-specific equipment
        aiic_equipment = [
            ["Tapping Machine:", "Norsonics", "CAL200", "2775671", "9/19/2022", "N/A"],
        ]
        return equipment + aiic_equipment
    
    def get_test_results(self):
                #dataframe for AIIC is in single_test_dataframe
        onethird_srs = format_SLMdata(self.test_data.srs_data) 
        average_pos = []
        main_elements = []
        # get average of 4 tapper positions for recieve total OBA
        for i in range(1, 5):
            pos_input = f'AIIC_pos{i}'
            pos_data = format_SLMdata(self.test_data.recive_data[pos_input]) ### need to verify if this is working correctly, should be pulling from all positions
            average_pos.append(pos_data)

        onethird_rec_Total = sum(average_pos) / len(average_pos)
        # this needs to be an average of the 4 tapper positions, stored in a dataframe of the average of the 4 dataframes octave band results. 


        onethird_bkgrd = format_SLMdata(single_test_dataframe['bkgrnd_data'])
        rt_thirty = single_test_dataframe['rt']['Unnamed: 10'][25:41]/1000

        calc_NR, sabines, corrected_recieve,Nrec_ANISPL = calc_nr_new(onethird_srs, onethird_rec_Total, onethird_bkgrd, rt_thirty,single_test_dataframe['room_properties']['receive_room_vol'][0],NIC_vollimit=883,testtype='AIIC')
        
        # ATL_val = calc_ATL_val(onethird_srs, onethird_rec, onethird_bkgrd,rt_thirty,room_properties['Partition area'][0],room_properties['Recieve Vol'][0])
        AIIC_contour_val, Contour_curve_result = calc_aiic_val(Nrec_ANISPL)

        IIC_curve = [2,2,2,2,2,2,1,0,-1,-2,-3,-6,-9,-12,-15,-18]
        IIC_contour_final = list()
        # initial application of the IIC curve to the first AIIC start value 
        for vals in IIC_curve:
            IIC_contour_final.append(vals+(110-AIIC_contour_val))
        # print(IIC_contour_final)
        #### Contour_final is the AIIC contour that needs to be plotted vs the ANISPL curve- we have everything to plot the graphs and the results table  #####
        Ref_label = f'AIIC {AIIC_contour_val} Contour'
        Field_IIC_label = 'Absorption Normalized Impact Sound Pressure Level, ANISPL (dB)'
        frequencies =[125,160,200,250,315,400,500,630,800,1000,1250,1600,2000,2500,3150,4000]
        Test_result_table = pd.DataFrame(
            {
                "Frequency": frequencies,
                "Absorption Normalized Impact Sound Pressure Level, ANISPL (dB)	": Nrec_ANISPL,
                "Average Receiver Background Level": onethird_bkgrd,
                "Average RT60 (Seconds)": rt_thirty,
                "Exceptions noted to ASTM E1007-14": AIIC_Exceptions
            }
        )
        main_elements.append(Paragraph("The Apparent Impact Insulation Class (AIIC) was calculated. The AIIC rating is based on Apparent Transmission Loss (ATL), and includes the effects of noise flanking. The AIIC reference contour is shown on the next page, and has been “fit” to the Apparent Transmission Loss values, in accordance with the procedure of "+standards_data[0][0]))

        return main_elements, Test_result_table

    def get_testres_table_notes(self):
        return "*This test does fully conform to the requir"
    def get_test_results_paragraph(self): 
        return (
            f"The Apparent Impact Insulation Class (AIIC) of {self.AIIC_contour_val} "
            f"and an Impact Sound Rating (ISR) of {self.ISR_val} was calculated. "
            "The AIIC rating is based on Apparent Transmission Loss (ATL), and includes "
            "the effects of noise flanking. The AIIC reference contour is shown on the "
            "next page, and has been fit to the Apparent Transmission Loss values, in "
            f"accordance with the procedure of {self.standards_data[0][0]}"
        )
    # Implement other methods specific to AIIC

class ASTCTestReport(BaseTestReport):
    # Implement ASTC-specific methods
    def get_standards_data(self):
        return [
            ['ASTM E336-20', 'Standard Test Method for Measurement of Airborne Sound Attenuation between Rooms in Buildings'],
            ['ASTM E413-16', 'Standard Classification for Rating Sound Insulation'],
            ['ASTM E2235-04(2012)', 'Standard Test Method for Determination of Decay Rates for Use in Sound Insulation Test Methods']
        ]
    def get_test_instrumentation(self):
        equipment = super().get_test_instrumentation()
        # Add ASTC-specific equipment
        astc_equipment = [
            ["Amplified Loudspeaker", "QSC", "K10", "GAA530909", "N/A", "N/A"],
            ["Noise Generator", "NTi Audio", "MR-PRO", "0162", "N/A", "N/A"],
        ]
        return equipment + astc_equipment
    def get_statement_of_conformace(self):
        return "Testing was conducted in accordance with ASTM E336-20, ASTM E413-16, and ASTM E2235-04(2012), with exceptions noted below. All requrements for measuring abd reporting Airborne Sound Attenuation between Rooms in Buildings (ATL) and Apparent Sound Transmission Class (ASTC) were met."
    def get_test_results(self):
        # Format the raw data
        onethird_rec = format_SLMdata(self.test_data.recive_data)
        onethird_srs = format_SLMdata(self.test_data.srs_data)
        onethird_bkgrd = format_SLMdata(self.test_data.bkgrnd_data)
        rt_thirty = self.test_data.rt['Unnamed: 10'][25:41]/1000

        # Get room properties
        partition_area = self.test_data.room_properties.partition_area
        receive_vol = self.test_data.room_properties.receive_vol

        # Calculate ATL and NR
        ATL_val, corrected_STC_recieve = calc_atl_val(
            onethird_srs, 
            onethird_rec, 
            onethird_bkgrd,
            rt_thirty,
            partition_area,
            receive_vol
        )

        calc_NR, sabines, corrected_recieve, Nrec_ANISPL = calc_nr_new(
            onethird_srs, 
            onethird_rec, 
            onethird_bkgrd, 
            rt_thirty,
            receive_vol,
            NIC_vollimit=883,
            testtype='ASTC'
        )

        # Create ASTC reference curve
        STCCurve = [-16, -13, -10, -7, -4, -1, 0, 1, 2, 3, 4, 4, 4, 4, 4, 4]
        ASTC_contour_final = [val + ATL_val for val in STCCurve]

        # Create results table
        Test_result_table = pd.DataFrame({
            "Frequency": frequencies,
            "L1, Average Source Room Level (dB)": onethird_srs,
            "L2, Average Corrected Receiver Room Level (dB)": corrected_STC_recieve,
            "Average Receiver Background Level (dB)": onethird_bkgrd,
            "Average RT60 (Seconds)": rt_thirty,
            "Noise Reduction, NR (dB)": calc_NR,
            "Apparent Transmission Loss, ATL (dB)": ATL_val,
            "Exceptions": ASTC_Exceptions
        })
        return ATL_val, Test_result_table
    
    def get_results_plot(self):
        plot_title = 'ASTC Reference Contour'
        plt.plot(ATL_curve, freqbands)
        plt.xlabel('Apparent Transmission Loss (dB)')
        plt.ylabel('Frequency (Hz)')
        plt.title('ASTC Reference Contour')
        plt.grid()

class NICTestReport(BaseTestReport):
    def get_doc_name(self):
        return f"{self.single_test_dataframe['room_properties']['Project_Name'][0]} NIC Test Report_{self.single_test_dataframe['room_properties']['Test_Label'][0]}.pdf"

    def get_standards_data(self):
        return [
            ['ASTM E336-20', 'Standard Test Method for Measurement of Airborne Sound Attenuation between Rooms in Buildings'],
            ['ASTM E413-16', 'Standard Classification for Rating Sound Insulation'],
            ['ASTM E2235-04(2012)', 'Standard Test Method for Determination of Decay Rates for Use in Sound Insulation Test Methods']
        ]

    # Implement other methods as needed
    def get_test_procedure(self):
        pass

    def get_test_instrumentation(self):
        pass

    def get_test_results(self, single_test_dataframe):
        ## obtain SLM data from overall dataframe
        ## need to convert all of this to use the dataclasses and the data_processor.py functions 
        onethird_rec = format_SLMdata(single_test_dataframe['recive_data'])
        onethird_srs = format_SLMdata(single_test_dataframe['srs_data'])
        onethird_bkgrd = format_SLMdata(single_test_dataframe['bkgrnd_data'])
        rt_thirty = single_test_dataframe['rt']['Unnamed: 10'][25:41]/1000
        # Calculation of ATL
        ATL_val,corrected_STC_recieve = calc_atl_val(onethird_srs, onethird_rec, onethird_bkgrd,single_test_dataframe['room_properties']['partition_area'][0],single_test_dataframe['room_properties']['receive_room_vol'][0],sabines)
        # Calculation of NR
        calc_NR, sabines, corrected_recieve,Nrec_ANISPL = calc_nr_new(onethird_srs, onethird_rec, onethird_bkgrd, rt_thirty,single_test_dataframe['room_properties']['receive_room_vol'][0],NIC_vollimit=883,testtype='NIC')
        # creating reference curve for ASTC graph
        STCCurve = [-16, -13, -10, -7, -4, -1, 0, 1, 2, 3, 4, 4, 4, 4, 4, 4]
        ASTC_contour_final = list()
        for vals in STCCurve:
            ASTC_contour_final.append(vals+(ATL_val))
        
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
        return Test_result_table

    def create_plot(self):
        pass

    def get_report_title(self):
        return "Noise Isolation Class (NIC) Test Report"

class DTCTestReport(BaseTestReport):
    pass
    # Implement DTC-specific methods