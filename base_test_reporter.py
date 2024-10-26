from reportlab.lib.units import inch
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import Paragraph, Spacer, Table, TableStyle, KeepInFrame, PageBreak
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.platypus import BaseDocTemplate, Frame, PageTemplate

from createreport_debug import getSampleStyleSheet
from config import *

class BaseTestReport:
    def __init__(self, curr_test, single_test_dataframe, reportOutputfolder):
        self.curr_test = curr_test
        self.single_test_dataframe = single_test_dataframe
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
        elements.append(Paragraph(self.get_report_title(), self.custom_title_style))
        elements.append(Spacer(1, 10))
        
        leftside_data = [
            ["Report Date:", self.single_test_dataframe['room_properties']['ReportDate'][0]],
            ['Test Date:', self.single_test_dataframe['room_properties']['Testdate'][0]],
            ['DLAA Test No', self.single_test_dataframe['room_properties']['Test number'][0]]
        ]
        rightside_data = [
            ["Source Room:", self.single_test_dataframe['room_properties']['Source Room Name'][0]],
            ["Receiver Room:", self.single_test_dataframe['room_properties']['Recieve Room Name'][0]],
            ["Test Assembly:", self.single_test_dataframe['room_properties']['Tested Assembly'][0]]
        ]

        table_left = Table(leftside_data)
        table_right = Table(rightside_data)
        table_left.setStyle(TableStyle([('GRID', (0, 0), (-1, -1), 1, colors.white)]))
        table_right.setStyle(TableStyle([('GRID', (0, 0), (-1, -1), 1, colors.white)]))

        table_combined_lr = Table([[table_left, table_right]], colWidths=[self.doc.width / 2.0] * 2)
        elements.append(KeepInFrame(maxWidth=self.doc.width, maxHeight=self.header_height, content=[table_combined_lr], hAlign='LEFT'))
        elements.append(Spacer(1, 10))
        elements.append(Paragraph('Test site: ' + self.single_test_dataframe['room_properties']['Site_Name'][0], self.styles['Normal']))
        elements.append(Spacer(1, 5))
        elements.append(Paragraph('Client: ' + self.single_test_dataframe['room_properties']['Client_Name'][0], self.styles['Normal']))
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

    def create_results_table(self):
        raise NotImplementedError

    def create_plot(self):
        raise NotImplementedError

    def get_report_title(self):
        raise NotImplementedError
    
class AIICTestReport(BaseTestReport):
    def get_doc_name(self):
        return f"{self.single_test_dataframe['room_properties']['Project_Name'][0]} AIIC Test Report_{self.single_test_dataframe['room_properties']['Test_Label'][0]}.pdf"

    def get_standards_data(self):
        return [
            ['ASTM E1007-14', 'Standard Test Method for Field Measurement of Tapping Machine Impact Sound Transmission Through Floor-Ceiling Assemblies and Associated Support Structure'],
            ['ASTM E413-16', 'Standard Classification for Rating Sound Insulation'],
            ['ASTM E2235-04(2012)', 'Standard Test Method for Determination of Decay Rates for Use in Sound Insulation Test Methods'],
            ['ASTM E989-06(2012)', 'Standard Classification for Determination of Impact Insulation Class (IIC)']
        ]
    def get_test_instrumentation(self):
        equipment = super().get_test_instrumentation()
        # Add AIIC-specific equipment
        aiic_equipment = [
            ["Tapping Machine:", "Norsonics", "CAL200", "2775671", "9/19/2022", "N/A"],
        ]
        return equipment + aiic_equipment

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

    def create_results_table(self):
        pass

    def create_plot(self):
        pass

    def get_report_title(self):
        return "Noise Isolation Class (NIC) Test Report"

class DTCTestReport(BaseTestReport):
    pass
    # Implement DTC-specific methods


def create_report(curr_test, single_test_dataframe, reportOutputfolder, test_type):
    # Create the appropriate test report object based on test_type
    if test_type == 'AIIC':
        report = AIICTestReport(curr_test, single_test_dataframe, reportOutputfolder)
    elif test_type == 'ASTC':
        report = ASTCTestReport(curr_test, single_test_dataframe, reportOutputfolder)
    elif test_type == 'NIC':
        report = NICTestReport(curr_test, single_test_dataframe, reportOutputfolder)
    elif test_type == 'DTC':
        report = DTCTestReport(curr_test, single_test_dataframe, reportOutputfolder)
    else:
        raise ValueError(f"Unsupported test type: {test_type}")

    # Document setup
    doc = BaseDocTemplate(f"{reportOutputfolder}/{report.get_doc_name()}", pagesize=letter,
                          leftMargin=left_margin, rightMargin=right_margin,
                          topMargin=top_margin, bottomMargin=bottom_margin)

    # ... (rest of the document setup code)

    # Generate content
    main_elements = []
    main_elements.extend(create_first_page(report, styles))
    main_elements.extend(create_second_page(report, styles))
    main_elements.extend(create_third_page(report, styles))

    # Build and save document
    doc.build(main_elements)
    print(f"Report saved as: {report.get_doc_name()}")

def create_first_page(report, styles):
    main_elements = []
    styleHeading = ParagraphStyle('heading', parent=styles['Normal'], spaceAfter=10)
    
    main_elements.append(Paragraph('<u>STANDARDS:</u>', styleHeading))
    standards_table = Table(report.get_standards_data(), hAlign='LEFT')
    # ... (set table style)
    main_elements.append(standards_table)

    # ... (add other elements using report methods)

    return main_elements

def create_second_page(report, styleHeading):
    main_elements = []
    
    # Test Procedure
    main_elements.append(Paragraph("<u>TEST PROCEDURE:</u>", styleHeading))
    main_elements.append(Paragraph('Determination of space-average sound pressure levels was performed via the manually scanned microphones techique, described in ' + standards_data[0][0] + ', Paragraph 11.4.3.3.'+ "The source room was selected in accordance with ASTM E336-11 Paragraph 9.2.5, which states that 'If a corridor must be used as one of the spaces for measurement of ATL or FTL, it shall be used as the source space.'"))
    main_elements.append(Spacer(1,10))
    main_elements.append(Paragraph("Flanking transmission was not evaluated."))
    main_elements.append(Paragraph("To evaluate room absorption, 1 microphone was used to measure 4 decays at 4 locations around the receiving room for a total of 16 measurements, per"+standards_data[2][0]))
    
    # Test Instrumentation
    main_elements.append(Paragraph("<u>TEST INSTRUMENTATION:</u>", styleHeading))
    
    test_instrumentation_table = report.get_test_instrumentation()

    
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
# Similarly update create_second_page and create_third_page

