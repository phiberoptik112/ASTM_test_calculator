from pathlib import Path
from reportlab.lib.units import inch
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.platypus import Paragraph, Spacer, Table, TableStyle, KeepInFrame, PageBreak
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.platypus import BaseDocTemplate, Frame, PageTemplate, Image
from reportlab.lib.enums import TA_CENTER

from typing import List, Dict, Optional
from dataclasses import dataclass
from enum import Enum

from src.core.test_data_manager import TestDataManager
from src.core.data_processor import (
    TestType, 
    RoomProperties,
    TestData,
    plot_curves
)

class BaseReport:
    """Base class for all test reports"""
    
    def __init__(self, test_data: TestData, output_folder: Path, test_type: TestType):
        self.test_data = test_data
        self.output_folder = Path(output_folder)
        self.test_type = test_type
        
        # Document setup
        self.doc = None
        self.styles = getSampleStyleSheet()
        self.setup_styles()
        
        # Define margins and heights
        self.margins = {
            'left': 0.75 * 72,    # 0.75 inch
            'right': 0.75 * 72,   # 0.75 inch
            'top': 0.25 * 72,     # 0.25 inch
            'bottom': 1 * 72      # 1 inch
        }
        
        self.heights = {
            'header': 2 * inch,
            'footer': 0.5 * inch,
            'main_content': None   # Will be calculated in setup_document
        }

    def setup_styles(self):
        """Setup custom styles for the report"""
        self.styles['Title'] = ParagraphStyle(
            'CustomTitle',
            parent=self.styles['Heading1'],
            fontSize=12,
            spaceAfter=30,
            alignment=TA_CENTER
        )
        
        self.styles['SectionHeader'] = ParagraphStyle(
            'SectionHeader',
            parent=self.styles['Heading2'],
            fontSize=10,
            spaceAfter=12
        )
        
        self.styles['TableHeader'] = ParagraphStyle(
            'TableHeader',
            parent=self.styles['Normal'],
            fontSize=8,
            fontName='Helvetica-Bold'
        )

    def setup_document(self) -> BaseDocTemplate:
        """Setup the document template with frames and margins"""
        try:
            # Calculate main content height
            self.heights['main_content'] = (
                letter[1] - 
                self.margins['top'] - 
                self.margins['bottom'] - 
                self.heights['header'] - 
                self.heights['footer']
            )
            
            # Create output path
            output_path = self.output_folder / self.get_doc_name()
            
            # Initialize document
            self.doc = BaseDocTemplate(
                str(output_path),
                pagesize=letter,
                leftMargin=self.margins['left'],
                rightMargin=self.margins['right'],
                topMargin=self.margins['top'],
                bottomMargin=self.margins['bottom']
            )
            
            # Create frames
            frames = self._create_frames()
            
            # Create page template
            template = PageTemplate(
                id='Standard',
                frames=frames,
                onPage=self.header_footer
            )
            
            self.doc.addPageTemplates([template])
            return self.doc
            
        except Exception as e:
            raise ReportGenerationError(f"Failed to setup document: {str(e)}")

    def _create_frames(self) -> List[Frame]:
        """Create frames for the document"""
        header_frame = Frame(
            self.margins['left'],
            letter[1] - self.margins['top'] - self.heights['header'],
            letter[0] - 2 * self.margins['left'],
            self.heights['header'],
            id='header'
        )
        
        main_frame = Frame(
            self.margins['left'],
            self.margins['bottom'] + self.heights['footer'],
            letter[0] - 2 * self.margins['left'],
            self.heights['main_content'],
            id='main'
        )
        
        footer_frame = Frame(
            self.margins['left'],
            self.margins['bottom'],
            letter[0] - 2 * self.margins['left'],
            self.heights['footer'],
            id='footer'
        )
        
        return [main_frame, header_frame, footer_frame]

    def header_footer(self, canvas, doc):
        """Draw header and footer on each page"""
        canvas.saveState()
        
        # Draw header
        header_story = self.header_elements()
        self.doc.handle_header(header_story, canvas)
        
        # Draw footer
        self.draw_footer(canvas, doc)
        
        canvas.restoreState()

    def header_elements(self) -> List:
        """Generate header elements"""
        elements = []
        props = self.test_data.room_properties
        
        # Title
        elements.extend(self.get_report_title())
        elements.append(Spacer(1, 5))
        
        # Create tables for header info
        left_table = self.create_info_table([
            ["Report Date:", props.report_date],
            ['Test Date:', props.test_date],
            ['Test No:', props.test_label],
            ['Test Site:', props.site_name],
            ['Client:', props.client_name]
        ])
        
        right_table = self.create_info_table([
            ["Source Room:", props.source_room],
            ["Receiver Room:", props.receive_room],
            ["Test Assembly:", props.tested_assembly]
        ])
        
        # Combine tables
        table_combined = Table(
            [[left_table, right_table]], 
            colWidths=[self.doc.width/2.0]*2,
            hAlign='LEFT'
        )
        
        elements.append(KeepInFrame(
            maxWidth=self.doc.width,
            maxHeight=self.heights['header'],
            content=[table_combined],
            hAlign='LEFT'
        ))
        
        return elements

    def create_info_table(self, data: List[List[str]]) -> Table:
        """Create formatted information table"""
        table = Table(data, colWidths=[100, 200])
        table.setStyle(TableStyle([
            ('GRID', (0, 0), (-1, -1), 1, colors.white),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),
            ('FONTNAME', (1, 0), (1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 0), (-1, -1), 8)
        ]))
        return table

    def draw_footer(self, canvas, doc):
        """Draw footer on the page"""
        canvas.setFont('Helvetica', 10)
        footer_text = f"This page alone is not a complete report. Page {doc.page} of {self.get_total_pages()}"
        canvas.drawCentredString(
            letter[0] / 2,
            self.margins['bottom'] + self.heights['footer'] / 2,
            footer_text
        )
        
        # Draw company info
        self.draw_company_info(canvas)

    def draw_company_info(self, canvas):
        """Draw company information in footer"""
        canvas.setFont('Helvetica', 8)
        address_lines = [
            "The Company",
            "PO BOX 1234",
            "Anywhere, USA, XXXXX",
            "www.thecompany.com",
            "XXX-XXX-XXXX"
        ]
        
        line_height = 12
        total_height = line_height * len(address_lines)
        start_y = self.margins['bottom'] + self.heights['footer'] - total_height
        
        for i, line in enumerate(address_lines):
            text_width = canvas.stringWidth(line)
            x = letter[0] - self.margins['right'] - text_width - 10
            y = start_y + (i * line_height)
            canvas.drawString(x, y, line)

    def get_doc_name(self) -> str:
        """Generate document name"""
        props = self.test_data.room_properties
        return f'{props.project_name} {self.test_type.value} Test Report_{props.test_label}.pdf'

    def get_total_pages(self) -> int:
        """Get total number of pages in report"""
        return 4  # Default value, override in specific report classes

    @abstractmethod
    def get_report_title(self) -> List[Paragraph]:
        """Get report title elements - must be implemented by subclasses"""
        pass

    def build(self):
        """Build the report"""
        self.setup_document()
        story = self.build_story()
        self.doc.build(story)

    @abstractmethod
    def build_story(self) -> List:
        """Build the main story - must be implemented by subclasses"""
        pass

class ReportGenerationError(Exception):
    """Custom exception for report generation errors"""
    pass 