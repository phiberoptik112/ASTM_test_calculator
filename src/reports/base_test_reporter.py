from reportlab.lib.units import inch
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.platypus import Paragraph, Spacer, Table, TableStyle, KeepInFrame, PageBreak
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.platypus import BaseDocTemplate, Frame, PageTemplate
from reportlab.platypus import Image
from reportlab.lib.enums import TA_CENTER
from src.config import standards_text, test_instrumentation_table, test_procedure_pg, ISR_ony_report, stockNIC_note, stockISR_notes, FREQUENCIES
from src.core.data_processor import *

class BaseTestReport:
    def __init__(self, test_data, reportOutputfolder, test_type):
        # print(f"Debug: BaseTestReport initialization:")
        # print(f"- Test data type: {type(test_data)}")
        # print(f"- Test data attributes: {dir(test_data)}")
        # print(f"- Has room_properties: {hasattr(test_data, 'room_properties')}")
        self.test_data = test_data
        self.reportOutputfolder = reportOutputfolder
        self.test_type = test_type
        self.calculated_values = getattr(test_data, 'calculated_values', {})
        print(f"\nInitializing {test_type} report with calculated values:")
        print(f"Available keys: {self.calculated_values.keys()}")
        
        # Initialize common values from calculated_values
        self.room_vol = self.calculated_values.get('room_vol')
        self.sabines = self.calculated_values.get('sabines')
        self.NR_val = self.calculated_values.get('NR_val')

        self.doc = None
        self.styles = getSampleStyleSheet()
        self.custom_title_style = self.styles['Heading1']
        # Define margins and heights
        self.left_margin = 0.75 * 72
        self.right_margin = 0.75 * 72
        self.top_margin = 0.25 * 72 # 0.25 inch
        self.bottom_margin = 1 * 72  # 1 inch
        self.header_height = 2 * inch
        self.footer_height = 0.5 * inch
        self.main_content_height = letter[1] - self.top_margin - self.bottom_margin - self.header_height - self.footer_height

    def validate_calculated_values(self):
        """Validate that required calculated values are present for the specific test type"""
        if not hasattr(self, 'calculated_values'):
            raise ValueError("No calculated values found")
            
        required_base_values = ['room_vol']
        missing_base = [key for key in required_base_values 
                       if key not in self.calculated_values]
        
        if missing_base:
            raise ValueError(f"Missing required base values: {missing_base}")
            
        # Test type specific validation is handled in child classes
        if hasattr(self, 'validate_test_specific_values'):
            self.validate_test_specific_values()

    def setup_document(self):
        """Setup the document template with proper frames and margins"""
        try:
            print('-=-=-=-=- INSIDE DOCUMENT SETUP =-=-=-=-=-=-')
            # print(f"Test Type: {self.test_type}")
            # print(f"Output Folder: {self.reportOutputfolder}")
            # print(f"Room Properties: {vars(self.test_data.room_properties)}")  # Print all properties

            # Create output path
            output_path = Path(self.reportOutputfolder) / self.get_doc_name()
            # print(f'Output path: {output_path}')
            self.doc = BaseDocTemplate(
                str(output_path),  # Convert Path to string
                pagesize=letter,
                leftMargin=self.left_margin, 
                rightMargin=self.right_margin,
                topMargin=self.top_margin, 
                bottomMargin=self.bottom_margin
            )
            # print('Base Document Template created')
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
            
        except Exception as e:
            raise ReportGenerationError(f"Failed to setup document: {str(e)}")

    def header_elements(self):
        elements = []
        print('Building header elements')
        props = vars(self.test_data.room_properties)
        elements.extend(self.get_report_title())
        elements.append(Spacer(1, 5))
        print('Building left side data')
        
        # Convert all data to strings and wrap in Paragraphs where needed
        leftside_data = [
            ["Report Date:", Paragraph(str(props['report_date']), self.styles['Normal'])],
            ['Test Date:', Paragraph(str(props['test_date']), self.styles['Normal'])],
            ['DLAA Test No', Paragraph(str(props['test_label']), self.styles['Normal'])],
            ['Test Site', Paragraph(str(props['site_name']), self.styles['Normal'])],
            ['Client', Paragraph(str(props['client_name']), self.styles['Normal'])]
        ]
        
        print('Building right side data')
        rightside_data = [
            ["Source Room:", Paragraph(str(props['source_room']), self.styles['Normal'])],
            ["Receiver Room:", Paragraph(str(props['receive_room']), self.styles['Normal'])],
            ["Test Assembly:", Paragraph(str(props['tested_assembly']), self.styles['Normal'])]
        ]

        # Create tables with proper styling
        table_left = Table(leftside_data, colWidths=[100, 200])  # Adjust widths as needed
        table_right = Table(rightside_data, colWidths=[100, 200])  # Adjust widths as needed
        
        table_left.setStyle(TableStyle([
            ('GRID', (0, 0), (-1, -1), 1, colors.white),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),
            ('FONTNAME', (1, 0), (1, -1), 'Helvetica')
        ]))
        
        table_right.setStyle(TableStyle([
            ('GRID', (0, 0), (-1, -1), 1, colors.white),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),
            ('FONTNAME', (1, 0), (1, -1), 'Helvetica')
        ]))

        # Combine tables
        table_combined_lr = Table(
            [[table_left, table_right]], 
            colWidths=[self.doc.width/2.0]*2,
            hAlign='LEFT'
        )
        
        # Add combined table to elements
        elements.append(KeepInFrame(
            maxWidth=self.doc.width, 
            maxHeight=self.header_height, 
            content=[table_combined_lr], 
            hAlign='LEFT'
        ))
        
        elements.append(Spacer(1, 10))
        return elements

    def header_footer(self, canvas, doc):
        canvas.saveState()
        print('Building header and footer')
        
        # Build header
        canvas.setFont('Helvetica', 10)
        self.header_frame._leftPadding = self.header_frame._rightPadding = 0
        header_story = self.header_elements()
        self.header_frame.addFromList(header_story, canvas)
        
        # Build footer
        canvas.setFont('Helvetica', 10)
        footer_text = f"This page alone is not a complete report. Page {doc.page} of 4"
        canvas.drawCentredString(
            letter[0] / 2, 
            self.bottom_margin + self.footer_height / 2, 
            footer_text
        )
        
        # Draw address block in footer on right side
        canvas.setFont('Helvetica', 8)  # Smaller font for address
        address_lines = [
            "970 N. Kalaheo Ave",
            "Suite A311",
            "Kailua, HI 96734",
            "www.dlaa.com",
            "808-254-3318x222"
        ]
        # Calculate position for right-aligned text block
        line_height = 12  # Adjust spacing between lines
        total_height = line_height * len(address_lines)
        start_y = self.bottom_margin + self.footer_height - total_height
        
        # Draw each line of the address
        for i, line in enumerate(address_lines):
            text_width = canvas.stringWidth(line)
            x = letter[0] - self.right_margin - text_width - 10  # Add 10pt padding from right
            y = start_y + (i * line_height)
            canvas.drawString(x, y, line)
        
        canvas.restoreState()


    def get_standards_data(self):
        """Get standards data based on test type"""
        print('-=-=-=-=-=-=-= Getting standards table data -=-=-=-=-=-=-=')
        # Test-specific standards
        standards_by_type = {
        TestType.AIIC: [
            ['ASTM E1007-14', Paragraph('Standard Test Method for Field Measurement of Tapping Machine Impact Sound Transmission Through Floor-Ceiling Assemblies and Associated Support Structure',self.styles['Normal'])],
            ['ASTM E413-16', Paragraph('Standard Classification for Rating Sound Insulation',self.styles['Normal'])],
            ['ASTM E1007-14', Paragraph('Standard Test Method for Field Measurement of Tapping Machine Impact Sound Transmission Through Floor-Ceiling Assemblies and Associated Support Structure',self.styles['Normal'])],
            ['ASTM E989-06(2012)', Paragraph('Standard Classification for Determination of Impact Insulation Class (IIC)',self.styles['Normal'])],
            ['ASTM E2235-04(2012)', Paragraph('Standard Test Method for Determination of Decay Rates for Use in Sound Insulation Test Methods',self.styles['Normal'])]
        ],
        TestType.ASTC: [
            ['ASTM E336-16', Paragraph('Standard Test Method for Measurement of Airborne Sound Attenuation between Rooms in Buildings',self.styles['Normal'])],
            ['ASTM E413-16', Paragraph('Standard Classification for Rating Sound Insulation',self.styles['Normal'])],
            ['ASTM E2235-04(2012)', Paragraph('Standard Test Method for Determination of Decay Rates for Use in Sound Insulation Test Methods',self.styles['Normal'])]
        ],
        TestType.NIC: [
            ['ASTM E336-16', Paragraph('Standard Test Method for Measurement of Airborne Sound Attenuation between Rooms in Buildings',self.styles['Normal'])],
            ['ASTM E413-16', Paragraph('Standard Classification for Rating Sound Insulation',self.styles['Normal'])],
            ['ASTM E2235-04(2012)', Paragraph('Standard Test Method for Determination of Decay Rates for Use in Sound Insulation Test Methods',self.styles['Normal'])]
        ],
        TestType.DTC: [
            ['ASTM E336-20', 'Standard Test Method for Measurement of Airborne Sound Attenuation between Rooms in Buildings']
            
            ]
        }
        print('>>>>>>>>> Standards table data retrieved <<<<<<<<<<')
        try:
            return standards_by_type[self.test_type]
        except KeyError:
            raise ValueError(f"Unsupported test type: {self.test_type}")

    def get_test_procedure(self):
        main_elements = []
        print('-=-=-=-==- Getting test procedure -=-=-=-=-')
        main_elements.append(Paragraph("<u>TEST PROCEDURE:</u>", self.styleHeading))
        main_elements.append(Paragraph('Determination of space-average sound pressure levels was performed via the manually scanned microphones techique, described in ' + standards_text[0][0] + ', Paragraph 11.4.3.3.'+ "The source room was selected in accordance with ASTM E336-11 Paragraph 9.2.5, which states that 'If a corridor must be used as one of the spaces for measurement of ATL or FTL, it shall be used as the source space.'"))
        main_elements.append(Spacer(1,10))
        main_elements.append(Paragraph("Flanking transmission was not evaluated."))
        main_elements.append(Paragraph("To evaluate room absorption, 1 microphone was used to measure 4 decays at 4 locations around the receiving room for a total of 16 measurements, per "+standards_text[2][0]))
        print('>>>>>>>>> Test procedure retrieved <<<<<<<<<<')
        return main_elements
        
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

    def create_results_table(self, Test_result_table):
        main_elements = []
        main_elements.append(Paragraph("<u>STATEMENT OF TEST RESULTS:</u>", self.styleHeading))
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
        return main_elements


    def get_report_title(self, style=None):
        print('Getting report title for specific test')
        # style = style or self.custom_title_style
        style = self.styleHeading
        title_by_type = {
            TestType.AIIC: [
                Paragraph("<b>Field Impact Sound Transmission Test Report</b>", style),
                Paragraph("<b>Apparent Impact Insulation Class (AIIC)</b>", style)
            ],
            TestType.ASTC: [
                Paragraph("<b>Field Sound Transmission Test Report</b>", style),
                Paragraph("<b>Apparent Sound Transmission Class (ASTC)</b>", style)
            ],
            TestType.NIC: [
                Paragraph("<b>Field Sound Transmission Test Report</b>", style),
                Paragraph("<b>Noise Isolation Class (NIC)</b>", style)
            ]
        }
        try:
            return title_by_type[self.test_type]
        except KeyError:
            raise ValueError(f"Unsupported test type: {self.test_type}")
    
    def nic_reporting_note(self):
        props = vars(self.test_data.room_properties)
        if int(props['source room vol']) >= 5300 or int(props['receive room vol']) >= 5300:
            NICreporting_Note = 'The receiver and/or source room had a volume exceeding 150 m3 (5,300 cu. ft.), and the absorption of the receiver and/or source room was greater than the maximum allowed per E336-16, Paragraph 9.4.1.2.'
        elif int(props['source room vol']) <= 833 or int(props['receive room vol']) <= 833:
            NICreporting_Note = 'The receiver and/or source room has a volume less than the minimum volume requirement of 25 m3 (883 cu. ft.).'
        else:
            NICreporting_Note = '---'
        return NICreporting_Note
    
    def get_doc_name(self):
        # print('Getting doc name')
        props = vars(self.test_data.room_properties)  # Convert to dictionary
        return f'{props["project_name"]} {self.test_type.value} Test Report_{props["test_label"]}.pdf'

    def save_report(self):
        self.doc.save()
    def create_test_environment_section(self):
        try:
            props = vars(self.test_data.room_properties)
            main_elements = []
            
            # Debug print
            # print('Debug: Room properties:')
            # for key, value in props.items():
            #     print(f"{key}: {type(value)} = {value}")

            # Helper function to safely handle string conversion
            def safe_str(value):
                """Convert any value to an appropriate string representation"""
                if isinstance(value, list):
                    return ' '.join(map(str, value))
                elif isinstance(value, float):
                    return f"{value:.1f}"  # Format float to 1 decimal place
                elif isinstance(value, (int, float)):
                    return f"{value:,}"    # Add thousand separators
                return str(value)

            # Create paragraphs with safer string handling
            source_room_desc = (
                f"The source room was {safe_str(props['source_room_name'])}. "
                f"The space was {safe_str(props['source_room_finish'])}. "
                f"The floor was {safe_str(props['srs_floor'])}. "
                f"The ceiling was {safe_str(props['srs_ceiling'])}. "
                f"The walls were {safe_str(props['srs_walls'])}. "
                "All doors and windows were closed during the testing period. "
                f"The source room had a volume of approximately {safe_str(props['source_vol'])} cu. ft."
            )

            receive_room_desc = (
                f"The receiver room was {safe_str(props['receive_room_name'])}. "
                f"The space was {safe_str(props['receive_room_finish'])}. "
                f"The floor was {safe_str(props['rec_floor'])}. "
                f"The ceiling was {safe_str(props['rec_ceiling'])}. "
                f"The walls were {safe_str(props['rec_walls'])}. "
                "All doors and windows were closed during the testing period. "
                f"The source room had a volume of approximately {safe_str(props['receive_vol'])} cu. ft."
            )

            assembly_desc = (
                f"The test assembly measured approximately {safe_str(props['partition_dim'])}, "
                f"and had an area of approximately {safe_str(props['partition_area'])} sq. ft."
            )

            test_assembly_desc = (
                f"The tested assembly was the {safe_str(props['test_assembly_type'])} "
                "The assembly was not field verified, and was based on information provided by the client "
                "and drawings for the project. The client advised that no slab treatment or self-leveling "
                "was applied. Results may vary if slab treatment or self-leveling or any adhesive is used "
                "in other installations."
            )

            # Build the document elements
            main_elements.extend([
                Paragraph("<u>TEST ENVIRONMENT:</u>", self.styleHeading),
                Paragraph(source_room_desc),
                Spacer(1, 10),
                Paragraph(receive_room_desc),
                Spacer(1, 10),
                Paragraph(assembly_desc),
                Spacer(1, 10),
                Paragraph("<u>TEST ASSEMBLY:</u>", self.styleHeading),
                Spacer(1, 10),
                Paragraph(test_assembly_desc)
            ])

            print('-=-=-=-=-=-=-= Test environment section created -=-=-=-=-=-=-=')
            return main_elements

        except Exception as e:
            print(f"Error in create_test_environment_section: {str(e)}")
            print(f"Properties causing error: {props}")
            raise
    
    def get_signatures(self):
        print('Getting signatures')
        main_elements = []
        main_elements.append(Paragraph('Test Conducted By:', self.styles['Normal']))
        jp_sig_image = Image('./images/JPsignature.png')
        jp_sig_image.drawHeight = 50
        jp_sig_image.drawWidth = 100
        jp_sig_image.hAlign = 'LEFT'
        main_elements.append(jp_sig_image)
        main_elements.append(Paragraph('Jake Pfitsch, Project Consultant', self.styles['Normal']))

        zw_sig_image = Image('./images/ZWsignature.png')
        zw_sig_image.drawHeight = 40
        zw_sig_image.drawWidth = 130
        zw_sig_image.hAlign = 'LEFT'
        main_elements.append(zw_sig_image)
        main_elements.append(Paragraph('Zane Wright, Project Consultant', self.styles['Normal']))
        return main_elements
    
    @classmethod
    def create_report(cls, test_data, output_folder: Path, test_type):
        """Factory method to create and build the appropriate test report"""
        # Validate output directory
        output_dir = Path(output_folder)
        if not output_dir.exists():
            output_dir.mkdir(parents=True, exist_ok=True)
        
        report_classes = {
            TestType.AIIC: AIICTestReport,
            TestType.ASTC: ASTCTestReport,
            TestType.NIC: NICTestReport,
            TestType.DTC: DTCTestReport
        }
        
        report_class = report_classes.get(test_type)
        if not report_class:
            raise ValueError(f"Unsupported test type: {test_type}")
        
        try:
            print('--=-=-=-=-= Creating Report Class and setting up document -=-=-=-=-=-=-')
            report = report_class(test_data=test_data, 
                                reportOutputfolder=output_dir, 
                                test_type=test_type)
            
            # Setup document
            doc = report.setup_document()
            print('Document setup complete')
            
            # Generate content with error handling
            main_elements = []
            try:
                # Ensure all elements are proper Flowable objects
                for page_method in [report.create_first_page, 
                                  report.create_second_page,
                                  report.create_third_page,
                                  report.create_fourth_page]:
                    # print(f'Creating {page_method.__name__}')
                    page_elements = page_method()
                    
                    # Validate elements
                    if not isinstance(page_elements, (list, tuple)):
                        raise ReportGenerationError(f"{page_method.__name__} did not return a list of elements")
                    
                    # Filter out any non-Flowable elements
                    from reportlab.platypus.flowables import Flowable
                    valid_elements = []
                    for element in page_elements:
                        if isinstance(element, (Flowable, str)):
                            if isinstance(element, str):
                                # Convert strings to Paragraphs
                                element = Paragraph(element, report.styles['Normal'])
                            valid_elements.append(element)
                        else:
                            print(f"Warning: Skipping invalid element type {type(element)} in {page_method.__name__}")
                    
                    main_elements.extend(valid_elements)
                    main_elements.append(PageBreak())
                    
            except Exception as e:
                print(f"Error generating report pages: {str(e)}")
                raise ReportGenerationError(f"Error in page generation: {str(e)}")
            
            try:
                # Build and save document
                output_path = output_dir / report.get_doc_name()
                # print(f"Building document with {len(main_elements)} elements")
                
                # Remove the last PageBreak if it exists
                if main_elements and isinstance(main_elements[-1], PageBreak):
                    main_elements.pop()
                    
                doc.build(main_elements)
                
                if not output_path.exists():
                    raise ReportGenerationError(f"Failed to save report to {output_path}")
                print(f"Report saved successfully to: {output_path}")
                return True
                
            except Exception as e:
                print(f"Error during document build: {str(e)}")
                print(f"Number of elements: {len(main_elements)}")
                print(f"Types of elements: {[type(elem) for elem in main_elements[:5]]}...")  # Print first 5 element types
                raise ReportGenerationError(f"Document build failed: {str(e)}")
                
        except Exception as e:
            print(f"Failed to create report: {str(e)}")
            raise ReportGenerationError(f"Report generation failed: {str(e)}")


    def create_first_page(self):
        """Creates the first page of the report with standards and conformance info.
        
        Returns:
            list: List of report elements for the first page
            
        Raises:
            ValueError: If required data is missing or invalid
            ReportGenerationError: If there's an error creating report elements
        """
        try:
            main_elements = []
            # print('Creating standards section')
            # Standards section
            self.styleHeading = ParagraphStyle('heading', parent=self.styles['Normal'], spaceAfter=10)

            main_elements.append(Paragraph('<u>STANDARDS:</u>', self.styleHeading))
            
            # Get and validate standards data
            standards_data = self.get_standards_data()
            if not standards_data:
                raise ValueError("Standards data is missing")
                
            standards_table = Table(standards_data, hAlign='LEFT')
            standards_table.setStyle(TableStyle([
                ('INNERGRID', (0, 0), (-1, -1), 0.25, colors.white),
                ('BOX', (0, 0), (-1, -1), 0.25, colors.white),
                ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                ('ALIGN', (0, 0), (-1, -1), 'LEFT')
            ]))
            main_elements.append(standards_table)
            # Statement of conformance
            main_elements.append(Paragraph("<u>STATEMENT OF CONFORMANCE:</u>", self.styleHeading))
            conformance_statement = self.get_statement_of_conformance()
            if not conformance_statement:
                raise ValueError("Conformance statement is missing")
            main_elements.append(Paragraph(conformance_statement))
            main_elements.append(Spacer(1, 10))
            # Test environment
            test_env = self.create_test_environment_section()
            print('>>>>>>>>> appending test environment <<<<<<<<<<')
            if not test_env:
                raise ValueError("Test environment information is missing")
            
            main_elements.extend(test_env)

            print('>>>>>>>>> Returning first page elements <<<<<<<<<<')
            return main_elements

        except AttributeError as e:
            raise ValueError(f"Missing required attribute: {str(e)}")
        except Exception as e:
            raise ReportGenerationError(f"Error generating first page: {str(e)}")

    def create_second_page(self):
        """Creates the second page with test procedure and instrumentation.
        
        Returns:
            list: List of report elements for the second page
        """
        try:
            main_elements = []
            
            # Test Procedure
            print('Creating test procedure')
            procedure = self.get_test_procedure()
            if not procedure:
                raise ValueError("Test procedure is missing")
                
            main_elements.extend(procedure)
            main_elements.append(Spacer(1, 10))
            
            # Test Instrumentation
            print('Creating test instrumentation')
            main_elements.append(Paragraph("<u>TEST INSTRUMENTATION:</u>", self.styleHeading))
            
            instrumentation = self.get_test_instrumentation()

            # Create header row with rotated text
            header_row = ["Equipment Type","Manufacturer","Model Number","Serial Number","Last NIST Traceable Calibration","Last Local Calibration"]
            rotated_style = ParagraphStyle(
                'RotatedHeader',
                fontName='Helvetica-Bold',
                fontSize=8,
                alignment=TA_CENTER,
                textColor=colors.black,
                leading=10
            )
            rotated_headers = [
                Paragraph(f'<rotate>{text}</rotate>', rotated_style)
                for text in header_row
            ]
            
            # Combine headers with data
            table_data = [rotated_headers] + instrumentation

            if not table_data:
                raise ValueError("Test instrumentation data is missing")
                
            test_instrumentation_table = Table(
                table_data, 
                rowHeights=[65] + [20]*(len(instrumentation)), 
                hAlign='LEFT'
            )

            test_instrumentation_table.setStyle(TableStyle([
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 8),  # Smaller font for header
                ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
                ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                ('FONTSIZE', (0, 1), (-1, -1), 8),  # Smaller font for data
                ('INNERGRID', (0, 0), (-1, -1), 0.25, colors.black),
                ('BOX', (0, 0), (-1, -1), 0.25, colors.black),
                ('LEFTPADDING', (0, 0), (-1, -1), 6),
                ('RIGHTPADDING', (0, 0), (-1, -1), 6),
                ('TOPPADDING', (0, 0), (-1, -1), 6),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
                ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                ('ALIGN', (0, 0), (-1, -1), 'LEFT')
            ]))
            
            main_elements.append(test_instrumentation_table)
            return main_elements
            
        except Exception as e:
            raise ReportGenerationError(f"Error generating second page: {str(e)}")

    def create_third_page(self):
        """Creates the third page with test results table and summary."""
        try:
            main_elements = []
            
            # Add test results heading
            main_elements.append(Paragraph("<u>STATEMENT OF TEST RESULTS:</u>", self.styleHeading))
            
            # Get and format test results table
            try:
                test_results_elements = self.get_test_results()
                print(f"Number of test result elements: {len(test_results_elements)}")
                main_elements.extend(test_results_elements)
                
            except Exception as e:
                print(f"Error details: {str(e)}")
                print(f"Test results type: {type(test_results_elements) if 'test_results_elements' in locals() else 'Not created'}")
                raise ReportGenerationError(f"Error creating test results table: {str(e)}")

            # Add single number result text box
            try:
                print('Adding single number result text box')
                result_paragraph = self.get_single_number_result()
                if result_paragraph:
                    main_elements.append(Spacer(1, 5))  # Add space before the box
                    main_elements.append(result_paragraph)
                    main_elements.append(Spacer(1, 5))  # Add space after the box
            except Exception as e:
                print(f"Error adding single number result: {str(e)}")

            # Add table notes
            try:
                print('=-=-=-=-=-=-= Appending test results table notes -=-=-=-=-=-=-=-')
                main_elements.append(Spacer(1,10))
                main_elements.append(Paragraph(self.get_test_results_paragraph()))
                main_elements.append(self.get_test_results_table_notes())
            except Exception as e:
                raise ReportGenerationError(f"Error adding table notes: {str(e)}")

            return main_elements
            
        except Exception as e:
            print(f"Full error in create_third_page: {str(e)}")
            raise ReportGenerationError(f"Error generating third page: {str(e)}")
        
    def create_fourth_page(self):
        main_elements = []
        print('=-=-=-=-=-=-= Creating results plot -=-=-=-=-=-=-=-')
        ## i have to get the different elements of each test type to create the plot

        main_elements.extend(self.get_results_plot())

        ## need to append a large text box with teh appropriate single number result from the test calc
        # #### sane here -single number result from the test calc
        main_elements.append(Spacer(1,10))
        main_elements.append(self.get_single_number_result())
        main_elements.append(Spacer(1,10))
        main_elements.extend(self.get_signatures())
        
        return main_elements

    def get_single_number_result(self):
        """Get the single number result text box based on test type"""
        try:
            # Get test result value based on test type
            result_value = None
            if hasattr(self.test_data, 'single_number_result'):
                result_value = self.test_data.single_number_result
            
            # Format text based on test type
            result_text = {
                TestType.AIIC: f"AIIC: {result_value}",
                TestType.ASTC: f"ASTC: {result_value}",
                TestType.NIC: f"NIC: {result_value}",
                TestType.DTC: f"DTC: {result_value}"
            }.get(self.test_type)

            if not result_text:
                raise ValueError(f"Unsupported test type: {self.test_type}")

            # Create paragraph style for large text
            large_text_style = ParagraphStyle(
                'LargeText',
                parent=self.styles['Normal'],
                fontSize=18,
                alignment=1,  # Right alignment
                spaceAfter=5,
                spaceBefore=5,
                borderWidth=3,
                borderColor=colors.black,
                borderPadding=7,
                leftIndent=200,
                rightIndent=200,
                leading=24,
                valighn='MIDDLE'
            )

            return Paragraph(result_text, large_text_style)

        except Exception as e:
            print(f"Error creating single number result: {str(e)}")
            return None

class AIICTestReport(BaseTestReport):
    def __init__(self, test_data, reportOutputfolder, test_type):
        super().__init__(test_data, reportOutputfolder, test_type)
        
        # Initialize AIIC-specific calculated values with defaults
        self.calculated_values = getattr(test_data, 'calculated_values', {})
        
        # Store all required values with proper error handling
        try:
            self.AIIC_recieve_corr = self.calculated_values.get('AIIC_recieve_corr')
            self.AIIC_Normalized_recieve = self.calculated_values.get('AIIC_Normalized_recieve')
            self.AIIC_contour_val = self.calculated_values.get('AIIC_contour_val')
            self.Contour_curve_result = self.calculated_values.get('AIIC_contour_result')
            self.sabines = self.calculated_values.get('sabines')  # Add this line
            self.NR_val = self.calculated_values.get('NR_val')    # Add this line
            
            # Validate required values
            required_values = ['AIIC_recieve_corr', 'AIIC_Normalized_recieve', 
                             'AIIC_contour_val', 'AIIC_contour_result', 'sabines']
            missing_values = [key for key in required_values 
                            if self.calculated_values.get(key) is None]
            
            if missing_values:
                print(f"Warning: Missing required values: {missing_values}")
                print(f"Available values: {self.calculated_values.keys()}")
        
        except Exception as e:
            print(f"Error initializing AIIC report: {str(e)}")
            print(f"Available calculated values: {self.calculated_values.keys()}")
            raise
        
        # Debug logging - fixed to handle numpy arrays
        print(f"AIIC Report Initialization:")
        print(f"- Calculated values keys: {self.calculated_values.keys()}")
        print(f"- AIIC normalized values present: {self.AIIC_Normalized_recieve is not None}")
        print(f"- AIIC normalized values shape: {self.AIIC_Normalized_recieve.shape if self.AIIC_Normalized_recieve is not None else 'None'}")

    def get_doc_name(self):
        props = vars(self.test_data.room_properties)
        return f"{props['project_name']} AIIC Test Report_{props['test_label']}.pdf"

    # def get_standards_data(self):
    #     return [
    #         ['ASTM E1007-14', 'Standard Test Method for Field Measurement of Tapping Machine Impact Sound Transmission Through Floor-Ceiling Assemblies and Associated Support Structure'],
    #         ['ASTM E413-16', 'Standard Classification for Rating Sound Insulation'],
    #         ['ASTM E2235-04(2012)', 'Standard Test Method for Determination of Decay Rates for Use in Sound Insulation Test Methods'],
    #         ['ASTM E989-06(2012)', 'Standard Classification for Determination of Impact Insulation Class (IIC)']
    #     ]
    def get_statement_of_conformance(self):
        return "Testing was conducted in accordance with ASTM E1007-14, ASTM E413-16, ASTM E2235-04(2012), and ASTM E989-06(2012), with exceptions noted below. All requrements for measuring abd reporting Absorption Normalized Impact Sound Pressure Level (ANISPL) and Apparent Impact Insulation Class (AIIC) were met."
    
    def get_test_environment(self):
        main_elements = []
        props = vars(self.test_data.room_properties)
        print('-=-=-=-=-=-=-= Getting AIIC test environment-=-=-=-=-=-=-=-=-')
        main_elements.append(Paragraph("<u>TEST ENVIRONMENT:</u>", self.styleHeading))
        main_elements.append(Paragraph('The source room was '+props['source_room_name']+'. The space was'+props['source_room_finish']+'. The floor was '+props['srs_floor']+'. The ceiling was '+props['srs_ceiling']+". The walls were"+props['srs_walls']+". All doors and windows were closed during the testing period. The source room had a volume of approximately "+props['source_vol']+"cu. ft."))
        ### Recieve room paragraph
        main_elements.append(Paragraph('The receiver room was '+props['receive_room_name']+'. The space was'+props['receive_room_finish']+'. The floor was '+props['rec_floor']+'. The ceiling was '+props['rec_ceiling']+". The walls were"+props['rec_walls']+". All doors and windows were closed during the testing period. The source room had a volume of approximately "+str(props['receive_vol'])+"cu. ft."))
        main_elements.append(Paragraph('The test assembly measured approximately '+props['partition_dim']+", and had an area of approximately "+str(props['partition_area'])+"sq. ft."))
        main_elements.append(Spacer(1, 5))  # Adds some space 
        # Heading 'TEST ASSEMBLY'
        main_elements.append(Paragraph("<u>TEST ASSEMBLY:</u>", self.styleHeading))
        main_elements.append(Paragraph("The tested assembly was the "+props['test_assembly_type']+". The assembly was not field verified, and was based on information provided by the client and drawings for the project. The client advised that no slab treatment or self-leveling was applied. Results may vary if slab treatment or self-leveling or any adhesive is used in other installations."))
        return main_elements

    def get_test_instrumentation(self):
        equipment = super().get_test_instrumentation()
        # Add AIIC-specific equipment - only adding tapping machine since other equipment is in base list
        aiic_equipment = [
            ["Tapping Machine", "Norsonics", "CAL200", "2775671", "9/19/2022", "N/A"],
        ]
        return equipment + aiic_equipment
    
    def get_test_results(self):
        print('-=-=-=-=-=-=-= Getting AIIC test results-=-=-=-=-=-=-=-=-')
        props = vars(self.test_data.room_properties)
        main_elements = []
        frequencies = [100, 125, 160, 200, 250, 315, 400, 500, 630, 800, 1000, 1250, 1600, 2000, 2500, 3150]
        try:
            # Validate required values are present and have correct shapes
            required_arrays = {
                'AIIC_recieve_corr': self.AIIC_recieve_corr,
                'AIIC_Normalized_recieve': self.AIIC_Normalized_recieve,
                'AIIC_contour_val': self.AIIC_contour_val
            }
            
            missing_arrays = [key for key, val in required_arrays.items() if val is None]
            if missing_arrays:
                raise ValueError(f"Missing required arrays: {', '.join(missing_arrays)}")
            
            # Debug print shapes
            print("\nArray shapes for AIIC test results:")
            for name, arr in required_arrays.items():
                if hasattr(arr, 'shape'):
                    print(f"- {name}: {arr.shape}")
                else:
                    print(f"- {name}: {type(arr)}")
            
            # Continue with existing table creation code...

            rt_thirty = self.test_data.rt[:16]
            rt_thirty = np.array(rt_thirty, dtype=np.float64).round(3)

            # Load background and RT data
            freq_indices = slice(13, 29)
            self.onethird_bkgrd = self.test_data.bkgrnd_data[freq_indices]
            self.onethird_bkgrd = np.array(self.onethird_bkgrd, dtype=np.float64).round(1)
            print(f"onethird_bkgrd: {self.onethird_bkgrd}")
            # Use stored values instead of recalculating
            self.NR_val = self.calculated_values['NR_val']
            self.sabines = self.calculated_values['sabines']
            self.AIIC_recieve_corr = self.calculated_values['AIIC_recieve_corr']
            self.ASTC_recieve_corr = self.calculated_values.get('ASTC_recieve_corr')
            self.AIIC_Normalized_recieve = self.calculated_values['AIIC_Normalized_recieve']
            self.AIIC_contour_val = self.calculated_values['AIIC_contour_val']
            self.Contour_curve_result = self.calculated_values['AIIC_contour_result']

            # still need to add these, but not used in current report format
            self.ISR_contour_val = self.calculated_values.get('ISR_contour_val')
            self.ISR_contour_result = self.calculated_values.get('ISR_contour_result')
            
            ##### Process exceptions #### moving this to the debug plot in main_window.py
            self.AIIC_Exceptions = []
            self.AIIC_exceptions_backcheck = []
            rec_roomvol = float(props['receive_vol'])

            # Calculate exceptions using stored values
            for val in self.sabines:
                val_float = float(val)
                self.AIIC_Exceptions.append('1' if val_float > 2 * rec_roomvol**(2/3) else '0')

            # Background check exceptions
            print(f"AIIC_Normalized_recieve: {self.AIIC_Normalized_recieve}")
            print(f"onethird_bkgrd: {self.onethird_bkgrd}")
            # Calculate background difference using correct column (overall level)
            background_diff = self.AIIC_Normalized_recieve - self.onethird_bkgrd[:,1]
            
            # Create exceptions list based on difference
            self.AIIC_exceptions_backcheck = ['0' if diff > 5 else '1' for diff in background_diff]
            #########################################################
            
            self.test_data.single_number_result = self.AIIC_contour_val

            # Continue with existing table creation code...
            if self.AIIC_Normalized_recieve is not None and self.onethird_bkgrd is not None and rt_thirty is not None and self.AIIC_Exceptions is not None:
                print(f"\nArray lengths before table creation:")
                print(f"frequencies: {len(frequencies)}")
                print(f"ANISPL: {len(self.AIIC_Normalized_recieve)}")
                print(f"Background: {len(self.onethird_bkgrd)}")
                print(f"RT30: {len(rt_thirty)}")
                print(f"Absorption Exceptions: {len(self.AIIC_Exceptions)}")
                print(f"Backgrnd check Exceptions: {len(self.AIIC_exceptions_backcheck)}")
                table_data = [
                    ['Frequency (Hz)',
                        'Absorption Normalized Impact Sound Pressure Level, ANISPL (dB)',
                        'Average Receiver Background Level (dB)',
                        'Average RT60 (seconds)',
                        'Exceptions noted to ASTM E1007-14',
                        'Backgrnd check Exceptions'
                        ]
                ]
                print("\nTable data types:")
                print(f"ANISPL: {type(self.AIIC_Normalized_recieve)}")
                print(f"Background: {type(self.onethird_bkgrd)}")
                print(f"RT30: {type(rt_thirty)}")
                print(f"Exceptions: {type(self.AIIC_Exceptions)}")
                print(f"Backgrnd check Exceptions: {type(self.AIIC_exceptions_backcheck)}")

                # Correct the loop to iterate over frequencies
                for i in range(len(frequencies)):
                    try:
                        anispl_val = float(self.AIIC_Normalized_recieve[i])
                        bkg_val = float(self.onethird_bkgrd[i,1])
                        rt_val = float(rt_thirty[i])
                        exceptions_val = self.AIIC_Exceptions[i]
                        backgrnd_check_val = self.AIIC_exceptions_backcheck[i]
                        
                        row = [
                            str(frequencies[i]),
                            f"{anispl_val:.1f}",
                            f"{bkg_val:.1f}",
                            f"{rt_val:.3f}",
                            str(exceptions_val),
                            str(backgrnd_check_val)
                        ]
                        print(f"Created row {i}: {row}")
                        table_data.append(row)
                        
                    except Exception as e:
                        print(f"Error creating row {i}: {str(e)}")
                        print(f"Values at index {i}:")
                        print(f"  ANISPL: {self.AIIC_Normalized_recieve[i] if i < len(self.AIIC_Normalized_recieve) else 'index error'}")
                        print(f"  Background: {self.onethird_bkgrd[i] if i < len(self.onethird_bkgrd) else 'index error'}")
                        print(f"  RT30: {rt_thirty[i] if i < len(rt_thirty) else 'index error'}")
                        continue

                if len(table_data) > 1:

                    # First, create shorter header texts that will be more readable when rotated
                    header_row = [
                        'Frequency\n(Hz)',
                        'Absorption\nNormalized\nImpact Sound\nPressure Level,\nANISPL (dB)',
                        'Average\nReceiver\nBackground\nLevel (dB)',
                        'Average\nRT60\n(seconds)',
                        'Exceptions\nnoted to\nASTM\nE1007-14',
                        'Backgrnd check\nExceptions'
                    ]
                    
                    # Create paragraph style for rotated headers
                    rotated_style = ParagraphStyle(
                        'RotatedHeader',
                        fontName='Helvetica-Bold',
                        fontSize=8,  # Slightly smaller font for headers
                        alignment=TA_CENTER,
                        textColor=colors.black,
                        leading=10  # Controls line spacing for multi-line text
                    )
                    
                    # Create rotated header paragraphs
                    rotated_headers = [
                        Paragraph(f'<rotate>{text}</rotate>', rotated_style)
                        for text in header_row
                    ]
                    table_data[0] = rotated_headers

                    # Create and style the table with adjusted dimensions
                    Test_result_table = Table(
                        table_data, 
                        colWidths=[65, 85, 65, 45, 45, 45],  # Narrower columns
                        rowHeights=[60] + [10]*(len(table_data)-1),  # Shorter rows overall
                        hAlign='LEFT'
                    )
                    
                    Test_result_table.setStyle(TableStyle([
                        # Header row styling
                        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                        ('FONTSIZE', (0, 0), (-1, 0), 8),  # Smaller font for header
                        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
                        
                        # Data rows styling
                        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                        ('FONTSIZE', (0, 1), (-1, -1), 8),  # Smaller font for data
                        
                        # General table styling
                        ('INNERGRID', (0, 0), (-1, -1), 0.25, colors.black),
                        ('BOX', (0, 0), (-1, -1), 0.25, colors.black),
                        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                        ('TOPPADDING', (0, 1), (-1, -1), 2),  # Reduced padding for data rows
                        ('BOTTOMPADDING', (0, 1), (-1, -1), 2),
                        ('LEFTPADDING', (0, 0), (-1, -1), 4),
                        ('RIGHTPADDING', (0, 0), (-1, -1), 4),
                    ]))
                    
                    main_elements.append(Test_result_table)
                    main_elements.append(Spacer(1, 10))  # Reduced spacer after table
                    print(f"Table created with {len(table_data)} rows")
                else:
                    print("Warning: No data rows were created for the table")
        except Exception as e:
            print(f"Error in AIIC test results: {str(e)}")
            print(f"Available calculated values: {self.calculated_values.keys()}")
            raise

        return main_elements

    def get_test_results_table_notes(self):
        return "The results stated in this report represent only the specific construction and acoustical conditions present at the time of the test. Measurements performed in accordance with this test method on nominally identical constructions and acoustical conditions may produce different results."
    def get_test_results_paragraph(self): 
        return (
            f"The Apparent Impact Insulation Class (AIIC) of {self.AIIC_contour_val} was calculated. The AIIC rating is based on Absorption Normalized Impact Sound Pressure Level (ANISPL), and includes the effects of noise flanking. The AIIC reference contour is shown on the next page, and has been fit to the Absorption Normalized Impact Sound Pressure Level values, in accordance with the procedure of "+standards_text[0][0]
        )
    def get_results_plot(self):
        main_elements = []

        IIC_curve = [2,2,2,2,2,2,1,0,-1,-2,-3,-6,-9,-12,-15,-18]
        IIC_contour_final = list()
        # initial application of the IIC curve to the first AIIC start value 
        for vals in IIC_curve:
            IIC_contour_final.append(vals+(110-self.AIIC_contour_val))
        # Create ASTC contour
        # IIC_contour_final = [val + self.ASTC_final_val for val in IIC_curve]
        
        # Define the target frequency range (125Hz to 3150Hz) - removed 4000Hz
        # need to add in 100hz
        freq_series = pd.Series([100, 125, 160, 200, 250, 315, 400, 500, 630, 800, 
                               1000, 1250, 1600, 2000, 2500, 3150])
        
        print(f'table_freqs: {freq_series.tolist()}')
        print(f'IIC_contour_final: {IIC_contour_final}')
        
        Ref_label = f'AIIC {self.AIIC_contour_val} Contour'
        
        AIIC_yAxis = 'Sound Pressure Level (dB)'
        Field_AIIC_label = 'Absorption Normalized Impact Sound Pressure Level, AIIC (dB)'
        
        # Ensure all arrays are numpy arrays of the same length
        ASTC_plot_img = plot_curves(
            frequencies=freq_series.tolist(),
            y_label=AIIC_yAxis,
            ref_curve=np.array(IIC_contour_final),
            field_curve=np.array(self.AIIC_Normalized_recieve),
            ref_label=Ref_label,
            field_label=Field_AIIC_label
        )
        # Create a flowable image that ReportLab can handle
        img = Image(ASTC_plot_img)
        img.drawHeight = 300
        img.drawWidth = 500
        main_elements.append(img)
        return main_elements

class ASTCTestReport(BaseTestReport):
    def __init__(self, test_data, reportOutputfolder, test_type):
        super().__init__(test_data, reportOutputfolder, test_type)
        
        # Initialize ASTC-specific calculated values with defaults
        self.calculated_values = getattr(test_data, 'calculated_values', {})
        
        try:
            # Store all required values
            self.ATL_val = self.calculated_values.get('ATL_val')
            self.ASTC_final_val = self.calculated_values.get('ASTC_final_val')
            self.ASTC_contour_val = self.calculated_values.get('ASTC_contour_val')
            self.sabines = self.calculated_values.get('sabines')
            self.NR_val = self.calculated_values.get('NR_val')  # Add this line
            
            # Validate required values
            required_values = ['ATL_val', 'ASTC_final_val', 'ASTC_contour_val', 
                             'sabines', 'NR_val']
            missing_values = [key for key in required_values 
                            if self.calculated_values.get(key) is None]
            
            if missing_values:
                print(f"Warning: Missing required values: {missing_values}")
                print(f"Available values: {self.calculated_values.keys()}")
                
        except Exception as e:
            print(f"Error initializing ASTC report: {str(e)}")
            print(f"Available calculated values: {self.calculated_values.keys()}")
            raise

    def get_doc_name(self):
        props = vars(self.test_data.room_properties)
        return f"{props['project_name']} ASTC Test Report_{props['test_label']}.pdf"

    def get_test_instrumentation(self):
        equipment = super().get_test_instrumentation()
        # Add ASTC-specific equipment - removing duplicates that are already in base list
        astc_equipment = []  # Empty since all equipment is already in base list
        return equipment + astc_equipment

    def get_statement_of_conformance(self):
        return "Testing was conducted in accordance with ASTM E336-20, ASTM E413-16, and ASTM E2235-04(2012), with exceptions noted below. All requrements for measuring abd reporting Airborne Sound Attenuation between Rooms in Buildings (ATL) and Apparent Sound Transmission Class (ASTC) were met."

    def get_test_results(self):
        try:
            print('-=-=-=-=-=-=-= Getting ASTC test results-=-=-=-=-=-=-=-=-')
            props = vars(self.test_data.room_properties)
            main_elements = []
                        # Define frequencies including 4kHz
            frequencies = [100, 125, 160, 200, 250, 315, 400, 500, 630, 800, 1000, 1250, 1600, 2000, 2500, 3150, 4000]
            # Get stored calculated values
            calculated_values = getattr(self.test_data, 'calculated_values', None)
            if not calculated_values:
                raise ValueError("No calculated values found - please run calculation and storage first")
            
            # Use stored values instead of recalculating
            self.ATL_val = calculated_values['ATL_val']
            self.NR_val = calculated_values['NR_val']
            self.sabines = calculated_values['sabines']
            self.ASTC_recieve_corr = calculated_values['ASTC_recieve_corr']
            self.ASTC_final_val = calculated_values['ASTC_final_val']
            self.ASTC_contour_val = calculated_values['ASTC_contour_val']
            
            # Still need to load raw data for the report
            freq_indices = slice(12, 29)  # Include 100, 4kHz
            onethird_srs = self.test_data.srs_data[freq_indices]
            print(f"ASTC onethird_srs: {onethird_srs}")

            onethird_srs = np.array(onethird_srs, dtype=np.float64).round(1)
            
            onethird_rec = self.test_data.recive_data[freq_indices]
            print(f"ASTC onethird_rec: {onethird_rec}")
            onethird_rec = np.array(onethird_rec, dtype=np.float64).round(1)
            
            onethird_bkgrd = self.test_data.bkgrnd_data[freq_indices]
            onethird_bkgrd = np.array(onethird_bkgrd, dtype=np.float64).round(1)
            
            rt_thirty = self.test_data.rt[:17]
            rt_thirty = np.array(rt_thirty, dtype=np.float64).round(3)

            # Calculate background check exceptions
            self.ASTC_exceptions_backcheck = []
            for i in range(len(onethird_rec)):
                background_diff = onethird_rec[i,1] - onethird_bkgrd[i,1]
                self.ASTC_exceptions_backcheck.append('0' if background_diff > 5 else '1')

            # Set single number result
            self.test_data.single_number_result = self.ASTC_final_val



            # Verify data for table creation
            if self.NR_val is not None and self.ASTC_recieve_corr is not None:
                print(f"\nArray lengths before table creation:")
                print(f"frequencies: {len(frequencies)}")
                print(f"NR_val: {len(self.NR_val)}")
                print(f"Background: {len(onethird_bkgrd[:,1])}")
                print(f"Source room level: {len(onethird_srs[:,1])}")
                print(f"RT30: {len(rt_thirty)}")
                print(f"Average corrected receiver room level: {len(self.ASTC_recieve_corr)}")
                print(f"ASTC exceptions backcheck: {len(self.ASTC_exceptions_backcheck)}")

                # Rest of the existing table creation code remains unchanged...
                table_data = [
                    ['Frequency (Hz)',
                     'L1, Average Source Room Level (dB)',
                     'L2, Average Corrected Receiver Room Level (dB)',
                     'Average Receiver Background Level (dB)',
                     'Average RT60 (seconds)',
                     'Noise Reduction, NR (dB)', 
                     'Apparent Transmission Loss, ATL (dB)',
                     'Backgrnd check\nExceptions'
                    ]
                ]

                # Continue with existing table creation code...
                for i in range(len(frequencies)):
                    try:
                        # Access numpy array values directly
                        srs_val = float(onethird_srs[i,1])
                        corr_rec_val = float(self.ASTC_recieve_corr[i])
                        bkg_val = float(onethird_bkgrd[i,1])
                        rt_val = float(rt_thirty[i])
                        NR_table_val = float(self.NR_val[i])
                        atl_val = float(self.ATL_val[i])
                        backgrnd_check_val = self.ASTC_exceptions_backcheck[i]
                        row = [
                            str(frequencies[i]),
                            f"{srs_val:.1f}",
                            f"{corr_rec_val:.1f}",
                            f"{bkg_val:.1f}",
                            f"{rt_val:.3f}",
                            f"{NR_table_val:.1f}",
                            f"{atl_val:.1f}",
                            str(backgrnd_check_val)
                        ]
                        print(f"Created row {i}: {row}")  # Debug output
                        table_data.append(row)
                        
                    except Exception as e:
                        print(f"Error creating row {i}: {str(e)}")
                        print(f"Values at index {i}:")
                        print(f"  NR_val: {self.NR_val[i] if i < len(self.NR_val) else 'index error'}")
                        print(f"  Background: {onethird_bkgrd.iloc[i] if hasattr(onethird_bkgrd, 'iloc') else onethird_bkgrd[i] if i < len(onethird_bkgrd) else 'index error'}")
                        print(f"  RT30: {rt_thirty.iloc[i] if hasattr(rt_thirty, 'iloc') else rt_thirty[i] if i < len(rt_thirty) else 'index error'}")
                        continue

                # Create and style the table
                if len(table_data) > 1:
                         # First, create shorter header texts that will be more readable when rotated
                    header_row = [
                    'Frequency\n(Hz)',
                    'L1,\nAverage\nSource\nRoom\nLevel\n(dB)',
                    'L2,\nAverage\nCorrected\nReceiver\nRoom\nLevel\n(dB)',
                    'Average\nReceiver\nBackground\nLevel\n(dB)',
                    'Average\nRT60\n(seconds)',
                    'Noise\nReduction,\nNR\n(dB)',
                    'Apparent\nTransmission\nLoss,\nATL\n(dB)',
                    'Exceptions\nnoted\nin\nASTM\nE336-16'
                    ]
                
                # Create paragraph style for rotated headers
                    rotated_style = ParagraphStyle(
                    'RotatedHeader',
                    fontName='Helvetica-Bold',
                    fontSize=8,  # Slightly smaller font for headers
                    alignment=TA_CENTER,
                    textColor=colors.black,
                    leading=10  # Controls line spacing for multi-line text
                    )
                
                # Create rotated header paragraphs
                    rotated_headers = [
                    Paragraph(f'<rotate>{text}</rotate>', rotated_style)
                    for text in header_row
                    ]
                    table_data[0] = rotated_headers

                # Create and style the table with adjusted dimensions
                    Test_result_table = Table(
                    table_data, 
                    colWidths=[65, 85, 65, 45, 65, 65, 65, 65],  # Narrower columns
                    rowHeights=[60] + [10]*(len(table_data)-1),  # Shorter rows overall
                    hAlign='LEFT'
                    )
                
                Test_result_table.setStyle(TableStyle([
                    # Header row styling
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, 0), 8),  # Smaller font for header
                    ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
                    
                    # Data rows styling
                    ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                    ('FONTSIZE', (0, 1), (-1, -1), 8),  # Smaller font for data
                    
                    # General table styling
                    ('INNERGRID', (0, 0), (-1, -1), 0.25, colors.black),
                    ('BOX', (0, 0), (-1, -1), 0.25, colors.black),
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                    ('TOPPADDING', (0, 1), (-1, -1), 2),  # Reduced padding for data rows
                    ('BOTTOMPADDING', (0, 1), (-1, -1), 2),
                    ('LEFTPADDING', (0, 0), (-1, -1), 4),
                    ('RIGHTPADDING', (0, 0), (-1, -1), 4),
                ]))
                main_elements.append(Test_result_table)
                main_elements.append(Spacer(1, 20))
                print(f"Table created with {len(table_data)} rows")
            else:
                print("Warning: No data rows were created for the table")

            return main_elements

        except Exception as e:
            print(f"Error in ASTC calculation: {str(e)}")
            print(f"ATL_val shape: {getattr(self.ATL_val, 'shape', 'no shape')}")
            # print(f"STCCurve length: {len(STCCurve)}")
            raise

        except Exception as e:
            print(f"Error in NR calculation: {str(e)}")
            raise

    def get_test_results_paragraph(self):
        return (
            f"The Apparent Sound Transmission Class (ASTC) of {self.ASTC_final_val} was calculated. The ASTC rating is based on Apparent Transmission Loss (ATL), and includes the effects of noise flanking. The ASTC reference contour is shown on the next page, and has been fit to the Apparent Transmission Loss values, in accordance with the procedure of "+standards_text[0][0]
        )
    def get_test_results_table_notes(self):
        return "The results stated in this report represent only the specific construction and acoustical conditions present at the time of the test. Measurements performed in accordance with this test method on nominally identical constructions and acoustical conditions may produce different results."
    
    def get_results_plot(self):
        
        main_elements = []

            # Define frequencies
        freq_series = pd.Series([100, 125, 160, 200, 250, 315, 400, 500, 630, 800, 
                                1000, 1250, 1600, 2000, 2500, 3150, 4000])

        # table_freqs = freq_series[mask].tolist()
        print(f'table_freqs: {freq_series.tolist()}')

        ASTCRef_label = f'ASTC {self.ASTC_final_val} Contour'
        Field_ASTC_label = 'Apparent Transmission Loss, ATL (dB)'
        ASTC_plot_img = plot_curves(
            frequencies=freq_series.tolist(),
            y_label=Field_ASTC_label,
            ref_curve=self.ASTC_contour_val,
            field_curve=np.array(self.ATL_val),  # Use ATL_val instead of NR_val
            ref_label=ASTCRef_label,
            field_label=Field_ASTC_label
        )
        # Create a flowable image that ReportLab can handle
        img = Image(ASTC_plot_img)
        img.drawHeight = 300
        img.drawWidth = 500
        main_elements.append(img)
        return main_elements

class NICTestReport(BaseTestReport):
    def __init__(self, test_data, reportOutputfolder, test_type):
        # Explicitly call parent's __init__
        super().__init__(test_data, reportOutputfolder, test_type)
        # Add any NIC-specific initialization here

    def get_doc_name(self):
        props = vars(self.test_data.room_properties)
        return f"{props['project_name']} NIC Test Report_{props['test_label']}.pdf"

    def get_statement_of_conformance(self):
        return "Testing was conducted in accordance with ASTM E336-20, ASTM E413-16, and ASTM E2235-04(2012), with exceptions noted below. All requrements for measuring abd reporting Airborne Sound Attenuation between Rooms in Buildings (ATL) and Noise Isolation Class (NIC) were met."

    # Implement other methods as needed
    def get_test_procedure(self):
        return super().get_test_procedure()
 
    def get_test_instrumentation(self):
        return super().get_test_instrumentation()

    def get_test_results(self):        
        print('-=-=-=-=-=-=-= Getting NIC test results-=-=-=-=-=-=-=-=-')
        props = vars(self.test_data.room_properties)
        main_elements = []
        
        try:
            # Get stored calculated values
            calculated_values = getattr(self.test_data, 'calculated_values', None)
            if not calculated_values:
                raise ValueError("No calculated values found - please run calculation and storage first")
            
            # Use stored values instead of recalculating
            self.NR_val = calculated_values['NR_val']
            self.sabines = calculated_values['sabines']
            self.NIC_recieve_corr = calculated_values['NIC_recieve_corr']
            self.NIC_final_val = calculated_values['NIC_final_val']
            self.NIC_contour_val = calculated_values['NIC_contour_val']
            
            # Still need to load raw data for the report
            freq_indices = slice(12, 29)  # Include 100, 4kHz
            onethird_srs = self.test_data.srs_data[freq_indices]
            onethird_srs = np.array(onethird_srs, dtype=np.float64).round(1)
            print(f"NIC onethird_srs: {onethird_srs}")
            
            onethird_rec = self.test_data.recive_data[freq_indices]
            # onethird_rec = onethird_rec[:,2] # 2nd column is overall level, first is frequencies. 
            onethird_rec = np.array(onethird_rec, dtype=np.float64).round(1)
            
            onethird_bkgrd = self.test_data.bkgrnd_data[freq_indices]
            onethird_bkgrd = np.array(onethird_bkgrd, dtype=np.float64).round(1)
            
            rt_thirty = self.test_data.rt # Use all values for NIC
            rt_thirty = np.array(rt_thirty, dtype=np.float64).round(3)

            # Set single number result
            self.test_data.single_number_result = self.NIC_final_val
                        # Calculate background check exceptions
            self.NIC_exceptions_backcheck = []
            for i in range(len(onethird_rec)):
                background_diff = onethird_rec[i,1] - onethird_bkgrd[i,1]
                self.NIC_exceptions_backcheck.append('0' if background_diff > 5 else '1')
            # Define frequencies including 4kHz
            frequencies = [100, 125, 160, 200, 250, 315, 400, 500, 630, 800, 1000, 1250, 1600, 2000, 2500, 3150, 4000]

            # Verify data for table creation
            if self.NR_val is not None and self.NIC_recieve_corr is not None:
                print(f"\nArray lengths before table creation:")
                print(f"frequencies: {len(frequencies)}")
                print(f"Noise Reduction: {len(self.NR_val)}")
                print(f"Background: {len(onethird_bkgrd)}")
                print(f"Source room level: {len(onethird_srs)}")
                print(f"RT30: {len(rt_thirty)}")
                print(f"Average corrected receiver room level: {self.NIC_recieve_corr}")
                print(f"NIC exceptions backcheck: {len(self.NIC_exceptions_backcheck)}")
                # Create table data
                table_data = [
                    ['Frequency (Hz)',
                        'L1, Average Source Room Level (dB)',
                        'L2, Average Corrected Receiver Room Level (dB)',
                        'Average Receiver Background Level (dB)',
                        'Average RT60 (seconds)',
                        'Noise Reduction, NR (dB)',
                        'Backgrnd check\nExceptions'
                    ]
                ]

                # Continue with existing table creation code...
                for i in range(len(frequencies)):
                    try:
                        # Access numpy array values directly
                        srs_val = float(onethird_srs[:,2][i])
                        corr_rec_val = float(self.NIC_recieve_corr[i])
                        bkg_val = float(onethird_bkgrd[:,2][i])
                        rt_val = float(rt_thirty[i])
                        NR_table_val = float(self.NR_val[i])
                        row = [
                            str(frequencies[i]),
                            f"{srs_val:.1f}",
                            f"{corr_rec_val:.1f}",
                            f"{bkg_val:.1f}",
                            f"{rt_val:.3f}",
                            f"{NR_table_val:.1f}",
                            f"{self.NIC_exceptions_backcheck[i]}"
                        ]
                        print(f"Created row {i}: {row}")  # Debug output
                        table_data.append(row)
                        
                    except Exception as e:
                        print(f"Error creating row {i}: {str(e)}")
                        print(f"Values at index {i}:")
                        print(f"  NR_val: {self.NR_val[i] if i < len(self.NR_val) else 'index error'}")
                        print(f"  Background: {onethird_bkgrd.iloc[i] if hasattr(onethird_bkgrd, 'iloc') else onethird_bkgrd[i] if i < len(onethird_bkgrd) else 'index error'}")
                        print(f"  RT30: {rt_thirty.iloc[i] if hasattr(rt_thirty, 'iloc') else rt_thirty[i] if i < len(rt_thirty) else 'index error'}")
                        continue

                # Create and style the table
                if len(table_data) > 1:
                    # First, create shorter header texts that will be more readable when rotated
                    header_row = [
                    'Frequency\n(Hz)',
                    'L1,\nAverage\nSource\nRoom\nLevel\n(dB)',
                    'L2,\nAverage\nCorrected\nReceiver\nRoom\nLevel\n(dB)',
                    'Average\nReceiver\nBackground\nLevel\n(dB)',
                    'Average\nRT60\n(seconds)',
                    'Noise\nReduction,\nNR\n(dB)',
                    'Backgrnd check\nExceptions'
                    ]
                
                # Create paragraph style for rotated headers
                    rotated_style = ParagraphStyle(
                    'RotatedHeader',
                    fontName='Helvetica-Bold',
                    fontSize=8,  # Slightly smaller font for headers
                    alignment=TA_CENTER,
                    textColor=colors.black,
                    leading=10  # Controls line spacing for multi-line text
                    )
                
                # Create rotated header paragraphs
                    rotated_headers = [
                    Paragraph(f'<rotate>{text}</rotate>', rotated_style)
                    for text in header_row
                    ]
                    table_data[0] = rotated_headers

                # Create and style the table with adjusted dimensions
                    Test_result_table = Table(
                    table_data, 
                    colWidths=[65, 85, 65, 45, 65, 65, 65],  # Narrower columns
                    rowHeights=[60] + [10]*(len(table_data)-1),  # Shorter rows overall
                    hAlign='LEFT'
                    )
                
                Test_result_table.setStyle(TableStyle([
                    # Header row styling
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, 0), 8),  # Smaller font for header
                    ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
                    
                    # Data rows styling
                    ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                    ('FONTSIZE', (0, 1), (-1, -1), 8),  # Smaller font for data
                    
                    # General table styling
                    ('INNERGRID', (0, 0), (-1, -1), 0.25, colors.black),
                    ('BOX', (0, 0), (-1, -1), 0.25, colors.black),
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                    ('TOPPADDING', (0, 1), (-1, -1), 2),  # Reduced padding for data rows
                    ('BOTTOMPADDING', (0, 1), (-1, -1), 2),
                    ('LEFTPADDING', (0, 0), (-1, -1), 4),
                    ('RIGHTPADDING', (0, 0), (-1, -1), 4),
                ]))
                
                main_elements.append(Test_result_table)
                main_elements.append(Spacer(1, 20))
                print(f"Table created with {len(table_data)} rows")
            else:
                print("Warning: No data rows were created for the table")

            return main_elements

        except Exception as e:
            print(f"Error in NIC calculation: {str(e)}")
            print(f"NR_val shape: {getattr(self.NR_val, 'shape', 'no shape')}")
            # print(f"STCCurve length: {len(STCCurve)}")
            raise

        except Exception as e:
            print(f"Error in NR calculation: {str(e)}")
            raise
        except Exception as e:
            print(f"Error in NIC test results: {str(e)}")
            print(f"Error type: {type(e)}")
            print(f"Array lengths at error:")

            print(f"NR_val: {len(self.NR_val) if hasattr(self, 'NR_val') else 'not created'}")

            raise     

    def get_test_results_table_notes(self):
        return "The results stated in this report represent only the specific construction and acoustical conditions present at the time of the test. Measurements performed in accordance with this test method on nominally identical constructions and acoustical conditions may produce different results."
    def get_test_results_paragraph(self):
        return (
            f"The Noise Isolation Class (NIC) of {self.NIC_final_val} was calculated. The NIC rating is based on Noise Reduction (NR), and includes the effects of noise flanking. The NIC reference contour is shown on the next page, and has been fit to the Noise Reduction values, in accordance with the procedure of "+standards_text[0][0]
        )

    def get_results_plot(self):
        main_elements = []
        NIC_yAxis = 'Noise Reduction, NR (dB)'
        main_elements = []

            # Define frequencies
        freq_series = pd.Series([100, 125, 160, 200, 250, 315, 400, 500, 630, 800, 
                                1000, 1250, 1600, 2000, 2500, 3150, 4000])
        ## need to add 4k data here 
        # freq_series = freq_series.append(pd.Series([4000]))
        STCCurve = [-16, -13, -10, -7, -4, -1, 0, 1, 2, 3, 4, 4, 4, 4, 4, 4, 4]  # Length 17
        NIC_ref_curve = list()
        for i in range(len(STCCurve)):
            NIC_ref_curve.append(STCCurve[i] + self.NIC_final_val)
        # table_freqs = freq_series[mask].tolist()
        print(f'table_freqs: {freq_series.tolist()}')
        print(f"NIC contour value: {self.NIC_contour_val}")
        print(f"NIC value: {self.NIC_final_val}")
        print(f"NR value: {self.NR_val}")
        NICRef_label = f'NIC {self.NIC_final_val} Contour'
        Field_NIC_label = 'Absorption Normalized Impact Sound Pressure Level, ANISPL (dB)'
        NIC_plot_img = plot_curves(
            frequencies=freq_series.tolist(),
            y_label=NIC_yAxis,
            ref_curve=NIC_ref_curve,
            field_curve=np.array(self.NR_val),
            ref_label=NICRef_label,
            field_label=Field_NIC_label
        )
        # Create a flowable image that ReportLab can handle
        img = Image(NIC_plot_img)
        img.drawHeight = 300
        img.drawWidth = 500
        main_elements.append(img)
        return main_elements

### DTC Report still needs development - needs a calculation method too 
### currently a placeholder class
class DTCTestReport(BaseTestReport):
    pass
    # Implement DTC-specific methods

# Custom exception class for report generation errors
class ReportGenerationError(Exception):
    """Custom exception for errors during report generation"""
    pass