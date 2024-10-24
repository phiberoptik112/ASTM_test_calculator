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
        raise NotImplementedError

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
            # ... other AIIC-specific standards
        ]

    # Implement other methods specific to AIIC

class ASTCTestReport(BaseTestReport):
    # Implement ASTC-specific methods
    pass

class NICTestReport(BaseTestReport):
    def get_doc_name(self):
        return f"{self.single_test_dataframe['room_properties']['Project_Name'][0]} NIC Test Report_{self.single_test_dataframe['room_properties']['Test_Label'][0]}.pdf"

    def get_standards_data(self):
        return [
            ['ASTM E336-20', 'Standard Test Method for Measurement of Airborne Sound Attenuation between Rooms in Buildings'],
            ['ASTM E413-16', 'Standard Classification for Rating Sound Insulation'],
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

# Similarly update create_second_page and create_third_page

