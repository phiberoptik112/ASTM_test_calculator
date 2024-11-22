from reportlab.lib.units import inch
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.platypus import Paragraph, Spacer, Table, TableStyle, KeepInFrame, PageBreak
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.platypus import BaseDocTemplate, Frame, PageTemplate


from config import *
from data_processor import *

class BaseTestReport:
    def __init__(self, test_data, reportOutputfolder, test_type):
        print(f"Debug: BaseTestReport initialization:")
        print(f"- Test data type: {type(test_data)}")
        print(f"- Test data attributes: {dir(test_data)}")
        print(f"- Has room_properties: {hasattr(test_data, 'room_properties')}")

        self.test_data = test_data
        self.reportOutputfolder = reportOutputfolder
        self.test_type = test_type
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
        """Setup the document template with proper frames and margins"""
        try:
            print('-=-=-=-=- INSIDE DOCUMENT SETUP =-=-=-=-=-=-')
            print(f"Test Type: {self.test_type}")
            print(f"Output Folder: {self.reportOutputfolder}")
            print(f"Room Properties: {vars(self.test_data.room_properties)}")  # Print all properties

            # Create output path
            output_path = Path(self.reportOutputfolder) / self.get_doc_name()
            print(f'Output path: {output_path}')
            self.doc = BaseDocTemplate(
                str(output_path),  # Convert Path to string
                pagesize=letter,
                leftMargin=self.left_margin, 
                rightMargin=self.right_margin,
                topMargin=self.top_margin, 
                bottomMargin=self.bottom_margin
            )
            print('Document created')
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
        props = self.test_data.room_properties
        elements.append(Paragraph(self.get_report_title(), self.custom_title_style))
        elements.append(Spacer(1, 10))
        print('Building left side data')
        leftside_data = [
            ["Report Date:", props.report_date],
            ['Test Date:', props.test_date],
            ['DLAA Test No', props.test_number]
        ]
        print('Building right side data')
        rightside_data = [
            ["Source Room:", props.source_room_name],
            ["Receiver Room:", props.recieve_room_name],
            ["Test Assembly:", props.tested_assembly]
        ]

        table_left = Table(leftside_data)
        table_right = Table(rightside_data)
        table_left.setStyle(TableStyle([('GRID', (0, 0), (-1, -1), 1, colors.white)]))
        table_right.setStyle(TableStyle([('GRID', (0, 0), (-1, -1), 1, colors.white)]))

        table_combined_lr = Table([[table_left, table_right]], colWidths=[self.doc.width / 2.0] * 2)
        elements.append(KeepInFrame(maxWidth=self.doc.width, maxHeight=self.header_height, content=[table_combined_lr], hAlign='LEFT'))
        elements.append(Spacer(1, 10))
        elements.append(Paragraph('Test site: ' + props.site_name, self.styles['Normal']))
        elements.append(Spacer(1, 5))
        elements.append(Paragraph('Client: ' + props.client_name, self.styles['Normal']))
        return elements

    def header_footer(self, canvas, doc):
        canvas.saveState()
        print('Building header and footer')
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


    def get_standards_data(self,styles):
        """Get standards data based on test type"""
        common_standards = [
            ['ASTM E336-20', 'Standard Test Method for Measurement of Airborne Sound Attenuation between Rooms in Buildings'],
            ['ASTM E413-22', 'Classification for Rating Sound Insulation']
        ]
        # Test-specific standards
        standards_by_type = {
        TestType.AIIC: [
            ['ASTM E1007-14', Paragraph('Standard Test Method for Field Measurement of Tapping Machine Impact Sound Transmission Through Floor-Ceiling Assemblies and Associated Support Structure',styles['Normal'])],
            ['ASTM E413-16', Paragraph('Standard Classification for Rating Sound Insulation',styles['Normal'])],
            ['ASTM E1007-14', Paragraph('Standard Test Method for Field Measurement of Tapping Machine Impact Sound Transmission Through Floor-Ceiling Assemblies and Associated Support Structure',styles['Normal'])],
            ['ASTM E989-06(2012)', Paragraph('Standard Classification for Determination of Impact Insulation Class (IIC)',styles['Normal'])],
            ['ASTM E2235-04(2012)', Paragraph('Standard Test Method for Determination of Decay Rates for Use in Sound Insulation Test Methods',styles['Normal'])]
        ],
        TestType.ASTC: [
            ['ASTM E336-16', Paragraph('Standard Test Method for Measurement of Airborne Sound Attenuation between Rooms in Buildings',styles['Normal'])],
            ['ASTM E413-16', Paragraph('Classification for Rating Sound Insulation',styles['Normal'])],
            ['ASTM E2235-04(2012)', Paragraph('Standard Test Method for Determination of Decay Rates for Use in Sound Insulation Test Methods',styles['Normal'])]
            *common_standards
        ],
        TestType.NIC: [
            ['ASTM E336-16', Paragraph('Standard Test Method for Measurement of Airborne Sound Attenuation between Rooms in Buildings',styles['Normal'])],
            ['ASTM E413-16', Paragraph('Classification for Rating Sound Insulation',styles['Normal'])],
            ['ASTM E2235-04(2012)', Paragraph('Standard Test Method for Determination of Decay Rates for Use in Sound Insulation Test Methods',styles['Normal'])]
            *common_standards
        ],
        TestType.DTC: [
            ['ASTM E336-20', 'Standard Test Method for Measurement of Airborne Sound Attenuation between Rooms in Buildings'],
            *common_standards
            ]
        }
    
        try:
            return standards_by_type[self.test_type]
        except KeyError:
            raise ValueError(f"Unsupported test type: {self.test_type}")

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
    
    def nic_reporting_note(self):
        if int(self.test_data.room_properties['source room vol']) >= 5300 or int(self.test_data.room_properties['receive room vol']) >= 5300:
            NICreporting_Note = 'The receiver and/or source room had a volume exceeding 150 m3 (5,300 cu. ft.), and the absorption of the receiver and/or source room was greater than the maximum allowed per E336-16, Paragraph 9.4.1.2.'
        elif int(self.test_data.room_properties['source room vol']) <= 833 or int(self.test_data.room_properties['receive room vol']) <= 833:
            NICreporting_Note = 'The receiver and/or source room has a volume less than the minimum volume requirement of 25 m3 (883 cu. ft.).'
        else:
            NICreporting_Note = '---'
        return NICreporting_Note
    
    def get_doc_name(self):
        print('Getting doc name')
        props = vars(self.test_data.room_properties)  # Convert to dictionary
        return f'{props["project_name"]} {self.test_type.value} Test Report_{props["test_label"]}.pdf'

    def save_report(self):
        self.doc.save()
    def create_test_environment_section(self):
        # Heading 'TEST ENVIRONMENT'
        props = self.test_data.room_properties
        main_elements = []
        main_elements.append(Paragraph("<u>TEST ENVIRONMENT:</u>", styleHeading))
        main_elements.append(Paragraph('The source room was '+props.source_room_name+'. The space was'+props.source_room_finish+'. The floor was '+single_test_dataframe['room_properties']['srs_floor'][0]+'. The ceiling was '+single_test_dataframe['room_properties']['srs_ceiling'][0]+". The walls were"+single_test_dataframe['room_properties']['srs_walls'][0]+". All doors and windows were closed during the testing period. The source room had a volume of approximately "+single_test_dataframe['room_properties']['source_room_vol'][0]+"cu. ft."))
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
        return main_elements

    @classmethod
    def create_report(cls, test_data, output_folder: Path, test_type):
        """Factory method to create and build the appropriate test report
        
        Args:
            test_data: TestData instance containing all test measurements and properties
            output_folder: Path object for the output directory
            test_type: Type of test (AIIC, ASTC, NIC, or DTC)
            
        Returns:
            BaseTestReport: The generated report instance
            
        Raises:
            ReportGenerationError: If report creation or saving fails
            ValueError: If test type is unsupported or path is invalid
        """
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
            # Create report instance using test_data which contains room_properties
            print('--=-=-=-=-= Creating Report Class and setting up document -=-=-=-=-=-=-')
            report = report_class(test_data = test_data, 
                                  reportOutputfolder = output_dir, 
                                  test_type = test_type)
            
            # Setup document
            doc = report.setup_document()
            print('Document setup complete')
            # Generate content with error handling
            main_elements = []
            try:
                print('Creating first page')
                main_elements.extend(report.create_first_page())
                main_elements.append(PageBreak())   
                print('Creating second page')
                main_elements.extend(report.create_second_page())
                main_elements.append(PageBreak())
                print('Creating third page')
                main_elements.extend(report.create_third_page())
                main_elements.append(PageBreak())
                print('Creating fourth page')
                main_elements.extend(report.create_fourth_page())
            except ReportGenerationError as e:
                print(f"Error generating report pages: {str(e)}")
                raise
            
            # Build and save document
            output_path = output_dir / report.get_doc_name()
            doc.build(main_elements, filename=str(output_path))
            
            if not output_path.exists():
                raise ReportGenerationError(f"Failed to save report to {output_path}")
                
            print(f"Report saved successfully to: {output_path}")
            return True
            
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
            print('Creating standards section')
            # Standards section
            styleHeading = ParagraphStyle('heading', parent=self.styles['Normal'], spaceAfter=10)
            main_elements.append(Paragraph('<u>STANDARDS:</u>', styleHeading))
            
            # Get and validate standards data
            standards_data = self.get_standards_data(styles)
            if not standards_data:
                raise ValueError("Standards data is missing")
                
            standards_table = Table(standards_data, hAlign='LEFT')
            standards_table.setStyle(TableStyle([
                ('INNERGRID', (0, 0), (-1, -1), 0.25, colors.white),
                ('BOX', (0, 0), (-1, -1), 0.25, colors.white),
                ('ALIGN', (0, 0), (-1, -1), 'LEFT')
            ]))
            main_elements.append(standards_table)

            # Statement of conformance
            main_elements.append(Paragraph("<u>STATEMENT OF CONFORMANCE:</u>", styleHeading))
            conformance_statement = self.get_statement_of_conformance()
            if not conformance_statement:
                raise ValueError("Conformance statement is missing")
            main_elements.append(Paragraph(conformance_statement))

            # Test environment
            test_env = self.create_test_environment_section()
            if not test_env:
                raise ValueError("Test environment information is missing")
            main_elements.append(Paragraph(test_env))

            # # Test assembly
            # test_assembly = self.get_test_assembly()
            # if not test_assembly:
            #     raise ValueError("Test assembly information is missing")
            # main_elements.append(test_assembly)

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
                
            main_elements.append(Paragraph("<u>TEST PROCEDURE:</u>", self.custom_title_style))
            main_elements.append(Paragraph(procedure))
            main_elements.append(Spacer(1, 10))
            
            # Test Instrumentation
            main_elements.append(Paragraph("<u>TEST INSTRUMENTATION:</u>", self.custom_title_style))
            
            instrumentation = self.get_test_instrumentation()
            if not instrumentation:
                raise ValueError("Test instrumentation data is missing")
                
            test_instrumentation_table = Table(instrumentation, hAlign='LEFT')
            test_instrumentation_table.setStyle(TableStyle([
                ('INNERGRID', (0, 0), (-1, -1), 0.25, colors.white),
                ('BOX', (0, 0), (-1, -1), 0.25, colors.white),
                ('LEFTPADDING', (0, 0), (-1, -1), 6),
                ('RIGHTPADDING', (0, 0), (-1, -1), 6),
                ('TOPPADDING', (0, 0), (-1, -1), 6),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
                ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                ('ALIGN', (0, 0), (-1, -1), 'LEFT')
            ]))
            
            main_elements.append(test_instrumentation_table)
            main_elements.append(PageBreak())
            
            return main_elements
            
        except Exception as e:
            raise ReportGenerationError(f"Error generating second page: {str(e)}")

    def create_third_page(self):
        """Creates the third page with test results table and summary.
        
        Returns:
            list: List of report elements for the third page
            
        Raises:
            ReportGenerationError: If there is an error generating the page content
        """
        try:
            main_elements = []
            
            # Add test results heading
            main_elements.append(Paragraph("<u>STATEMENT OF TEST RESULTS:</u>", self.custom_title_style))
            
            # Get and format test results table
            try:
                test_results = self.get_test_results()
                results_table = Table(test_results, hAlign='LEFT')
                results_table.setStyle(TableStyle([
                    ('INNERGRID', (0, 0), (-1, -1), 0.25, colors.white),
                    ('BOX', (0, 0), (-1, -1), 0.25, colors.white),
                    ('LEFTPADDING', (0, 0), (-1, -1), 6),
                    ('RIGHTPADDING', (0, 0), (-1, -1), 6),
                    ('TOPPADDING', (0, 0), (-1, -1), 6),
                    ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
                    ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                    ('ALIGN',(0,0), (-1,-1),'LEFT')
                ]))
                main_elements.append(results_table)
            except Exception as e:
                raise ReportGenerationError(f"Error creating test results table: {str(e)}")

            # Add table notes
            try:
                main_elements.append(self.get_testres_table_notes())
            except Exception as e:
                raise ReportGenerationError(f"Error adding table notes: {str(e)}")

            # Add test type label
            test_type_labels = {
                TestType.AIIC: "AIIC",
                TestType.ASTC: "ASTC", 
                TestType.NIC: "NIC",
                TestType.DTC: "DTC"
            }
            test_label = test_type_labels.get(self.curr_test.test_type, "Unknown")
            main_elements.append(Paragraph(f"<u>{test_label}:</u>", self.custom_title_style))
            
            # Add spacing and test results paragraph
            main_elements.append(Spacer(1,10))
            main_elements.append(self.get_test_results_paragraph())
            main_elements.append(PageBreak())
            
            return main_elements
            
        except Exception as e:
            raise ReportGenerationError(f"Error generating third page: {str(e)}")
        
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


        onethird_bkgrd = format_SLMdata(self.test_data.bkgrnd_data)
        rt_thirty = self.test_data.rt['Unnamed: 10'][25:41]/1000

        calc_NR, sabines, corrected_recieve,Nrec_ANISPL = calc_nr_new(onethird_srs, onethird_rec_Total, onethird_bkgrd, rt_thirty,self.test_data.room_properties['receive_room_vol'][0],NIC_vollimit=883,testtype='AIIC')
        
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
        # frequencies =[125,160,200,250,315,400,500,630,800,1000,1250,1600,2000,2500,3150,4000]
        Test_result_table = pd.DataFrame(
            {
                "Frequency": FREQUENCIES,
                "Absorption Normalized Impact Sound Pressure Level, ANISPL (dB)	": Nrec_ANISPL,
                "Average Receiver Background Level": onethird_bkgrd,
                "Average RT60 (Seconds)": rt_thirty,
                "Exceptions noted to ASTM E1007-14": AIIC_Exceptions
            }
        )
        main_elements.append(Paragraph("The Apparent Impact Insulation Class (AIIC) was calculated. The AIIC rating is based on Apparent Transmission Loss (ATL), and includes the effects of noise flanking. The AIIC reference contour is shown on the next page, and has been “fit” to the Apparent Transmission Loss values, in accordance with the procedure of "+self.get_standards_data()[0][0]))

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
        main_elements = []
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
            "Frequency": FREQUENCIES,
            "L1, Average Source Room Level (dB)": onethird_srs,
            "L2, Average Corrected Receiver Room Level (dB)": corrected_STC_recieve,
            "Average Receiver Background Level (dB)": onethird_bkgrd,
            "Average RT60 (Seconds)": rt_thirty,
            "Noise Reduction, NR (dB)": calc_NR,
            "Apparent Transmission Loss, ATL (dB)": ATL_val,
            "Exceptions": ASTC_Exceptions
        })
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
        main_elements.append(Paragraph("The Apparent Sound Transmission Class (ASTC) was calculated. The ASTC rating is based on Apparent Transmission Loss (ATL), and includes the effects of noise flanking. The ASTC reference contour is shown on the next page, and has been “fit” to the Apparent Transmission Loss values, in accordance with the procedure of "+self.get_standards_data()[0][0]))
        return main_elements
    
    def get_results_plot(self, ATL_curve):
        plot_title = 'ASTC Reference Contour'
        plt.plot(ATL_curve, FREQUENCIES)
        plt.xlabel('Apparent Transmission Loss (dB)')
        plt.ylabel('Frequency (Hz)')
        plt.title('ASTC Reference Contour')
        plt.grid()

class NICTestReport(BaseTestReport):
    def __init__(self, test_data, reportOutputfolder, test_type):
        # Explicitly call parent's __init__
        super().__init__(test_data, reportOutputfolder, test_type)
        # Add any NIC-specific initialization here

    def get_doc_name(self):
        self.props = self.test_data.room_properties
        return f"{self.props['Project_Name'][0]} NIC Test Report_{self.props['Test_Label'][0]}.pdf"

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

    def get_test_results(self):
        props = self.test_data.room_properties
        ## obtain SLM data from overall dataframe
        ## need to convert all of this to use the dataclasses and the data_processor.py functions 
        onethird_rec = format_SLMdata(self.test_data.recive_data)
        onethird_srs = format_SLMdata(self.test_data.srs_data)
        onethird_bkgrd = format_SLMdata(self.test_data.bkgrnd_data)
        rt_thirty = self.test_data.rt['Unnamed: 10'][25:41]/1000
        # Calculation of ATL
        ATL_val,corrected_STC_recieve = calc_atl_val(onethird_srs, onethird_rec, onethird_bkgrd,props.partition_area,props.receive_room_vol,sabines)
        # Calculation of NR
        calc_NR, sabines, corrected_recieve,Nrec_ANISPL = calc_nr_new(onethird_srs, onethird_rec, onethird_bkgrd, rt_thirty,props.receive_room_vol,NIC_vollimit=883,testtype='NIC')
        # creating reference curve for ASTC graph
        STCCurve = [-16, -13, -10, -7, -4, -1, 0, 1, 2, 3, 4, 4, 4, 4, 4, 4]
        ASTC_contour_final = list()
        for vals in STCCurve:
            ASTC_contour_final.append(vals+(ATL_val))
        
        Test_result_table = pd.DataFrame(
            {
                "Frequency": FREQUENCIES,
                "Source OBA": onethird_srs,
                "Reciever OBA": onethird_rec,
                "Background OBA": onethird_bkgrd,
                "NR": calc_NR,
                "Exceptions": NIC_exceptions
            }
        )
        return Test_result_table

    def get_results_plot(self, ATL_curve):
        plot_title = 'NIC Reference Contour'
        # plt.plot(ATL_curve, FREQUENCIES)
        resultsplotfig = plot_curves(FREQUENCIES, 'Apparent Transmission Loss (dB)', ATL_curve, ATL_curve, 'NIC Reference Contour', 'NIC Reference Contour')
        plt.xlabel('Apparent Transmission Loss (dB)')
        plt.ylabel('Frequency (Hz)')
        plt.title('NIC Reference Contour')
        plt.grid()

    def get_report_title(self):
        return "Noise Isolation Class (NIC) Test Report"

class DTCTestReport(BaseTestReport):
    pass
    # Implement DTC-specific methods

# Custom exception class for report generation errors
class ReportGenerationError(Exception):
    """Custom exception for errors during report generation"""
    pass