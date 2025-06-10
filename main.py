from datetime import datetime
import os
import platform
import subprocess
from enum import Enum
import json
import ctypes
import re
import tkinter as tk
from tkinter import filedialog
import io

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.pagesizes import letter

from pypdf import PdfReader, PdfWriter
from pypdf.generic import (DictionaryObject, NumberObject, NameObject, 
                              TextStringObject, ArrayObject, FloatObject)

# Font registration
pdfmetrics.registerFont(TTFont('Arial', 'C:\\Windows\\Fonts\\arial.ttf'))
pdfmetrics.registerFont(TTFont('ArialBold', 'C:\\Windows\\Fonts\\arialbd.ttf'))

#####################################################################################################################
def get_user_choice(prompt, default=False):
    """
    Get user's yes/no response.

    Args:
        prompt (str): Question text for user
        default (bool): Default response if user hits enter

    Returns:
        bool: True for yes, False for no
    """
    options = 'Y/n' if default else 'y/N'
    
    while True:
        answer = input(f"{prompt} [{options}]: ").lower()
        
        if answer == '':
            return default
        elif answer in ['y', 'yes']:
            return True
        elif answer in ['n', 'no']:
            return False
        else:
            print("Please enter 'y' or 'n'")

def get_display_name():
    """
    Gets the full display name of current Windows user using Windows API.

    Args:
        None

    Returns:
        str: Full display name (e.g. "John Smith")
    """
    GetUserNameEx = ctypes.windll.secur32.GetUserNameExW
    NameDisplay = 3

    size = ctypes.pointer(ctypes.c_ulong(0))
    GetUserNameEx(NameDisplay, None, size)

    nameBuffer = ctypes.create_unicode_buffer(size.contents.value) 
    GetUserNameEx(NameDisplay, nameBuffer, size)
    return nameBuffer.value

def open_file(filename):
        """Opens PDF file in default viewer."""
        try:
            if platform.system() == 'Darwin':  # macOS
                subprocess.run(['open', filename])
            elif platform.system() == 'Windows':  # Windows
                os.startfile(filename)
            else:  # Linux or others
                subprocess.run(['xdg-open', filename])
        except Exception as e:
            print(f"Nepodarilo sa otviriť PDF súbor: {e}")

#####################################################################################################################
def add_text_label(pdf_file, text, position, page_number=0):
    """
    Adds text label to PDF file.

    Args:
        pdf_file (str): Path to PDF file
        text (str): Text to add
        page_number (int): Page number where to add text (0-based)
        position (tuple): (x, y) coordinates for text position

    Returns:
        None
    """
    def create_text_page(text, position):
        packet = io.BytesIO()
        can = canvas.Canvas(packet, pagesize=letter)
        
        # Set font for diacritic support
        can.setFont('Arial', 10)
        
        can.drawString(position[0], position[1], text)
        can.save()
        packet.seek(0)
        return PdfReader(packet)

    reader = PdfReader(pdf_file)
    writer = PdfWriter()
    
    # Check if requested page exists
    if page_number >= len(reader.pages):
        raise ValueError(f"PDF has only {len(reader.pages)} pages, cannot add text to page {page_number + 1}")
    
    # Copy all pages
    writer.append_pages_from_reader(reader)
    
    # Add text to specified page
    page = writer.pages[page_number]
    text_pdf = create_text_page(text, position)
    page.merge_page(text_pdf.pages[0])

    # Save changes
    with open(pdf_file, "wb") as fp:
        writer.write(fp)

def add_footer(pdf_file, protocol_number):
    """
    Adds page numbers and protocol number to PDF file footer.

    Args:
        pdf_file (str): Path to PDF file
        protocol_number (str): Protocol number to be added

    Returns:
        None
    """
    
    reader = PdfReader(pdf_file)
    total_pages = len(reader.pages)
    
    for page_num in range(total_pages):
        # Add page numbers
        page_num_text = f"{page_num + 1}/{total_pages}"
        add_text_label(pdf_file, page_num_text, position=(535, 20), page_number=page_num)
        
        # Add protocol number
        page_num_protocol = f"Číslo protokolu: {protocol_number}"  # Protocol number: could be translated if needed
        add_text_label(pdf_file, page_num_protocol, position=(42, 20), page_number=page_num)

def add_comment_to_pdf(pdf_file, title, text_list, position):
    """
    Adds a comment (annotation) to PDF file.

    Args:
        pdf_file (str): Path to PDF file
        title (str): Comment title
        text_list (list): List of text lines to be added in comment
        position (tuple): (x, y) coordinates for comment position

    Returns:
        None
    """
    def create_text_annotation(x, y, title, text):
        text_annotation = DictionaryObject()
        text_annotation.update({
            NameObject("/F"): NumberObject(4),
            NameObject("/Type"): NameObject("/Annot"),
            NameObject("/Subtype"): NameObject("/Text"),
            NameObject("/T"): TextStringObject(title),
            NameObject("/Contents"): TextStringObject(text),
            NameObject("/Rect"): ArrayObject([
                FloatObject(x),
                FloatObject(y),
                FloatObject(x + 20),
                FloatObject(y + 20)
            ]),
            NameObject("/C"): ArrayObject([FloatObject(1), FloatObject(1), FloatObject(0.8)]),
            NameObject("/Open"): NameObject("/true")
        })
        return text_annotation

    reader = PdfReader(pdf_file)
    writer = PdfWriter()
    
    if len(reader.pages) > 0:
        # Kopírovanie všetkých strán
        writer.append_pages_from_reader(reader)
        page = writer.pages[0]
        
        if text_list:
            annotation = create_text_annotation(
                x=position[0],
                y=position[1],
                title=title,
                text=f"{title}:\n" + "\n".join(text_list)
            )
            if "/Annots" in page:
                page["/Annots"].append(annotation)
            else:
                page[NameObject("/Annots")] = ArrayObject([annotation])

    # Uloženie zmien
    with open(pdf_file, "wb") as fp:
        writer.write(fp)

def add_error_comments_to_pdf(pdf_file, fixable_errors, unfixable_errors):
    """
    Adds error comments to PDF file.

    Args:
        pdf_file (str): Path to PDF file
        fixable_errors (list): List of fixable errors
        unfixable_errors (list): List of unfixable errors

    Returns:
        None
    """
    if unfixable_errors:
        add_comment_to_pdf(
            pdf_file,
            "Neopraviteľné zmätky",
            unfixable_errors,
            (525, 607)
        )
    
    if fixable_errors:
        add_comment_to_pdf(
            pdf_file,
            "Opraviteľné zmätky",
            fixable_errors,
            (525, 584)
        )
    
def add_attachments_to_pdf(pdf_file, attachment_list):
    """
    Adds files as attachments to PDF document.

    Args:
        pdf_file (str): Path to PDF file
        attachment_list (list): List of file paths to attach

    Returns:
        None
    """

    reader = PdfReader(pdf_file)
    writer = PdfWriter()
    writer.append_pages_from_reader(reader)

    for attachment in attachment_list:
        if os.path.exists(attachment):
            with open(attachment, "rb") as file:
                writer.add_attachment(attachment, file.read())

    with open(pdf_file, "wb") as output_file:
        writer.write(output_file)

#####################################################################################################################
#####################################################################################################################
class Colors(Enum):
    DEFAULT = None
    WHITE = colors.white
    GREY = colors.Color(0.9, 0.9, 0.9)      # light grey
    RED = colors.red
    LIGHT_RED = colors.Color(1, 0.8, 0.8)    # light red
    GREEN = colors.green
    LIGHT_GREEN = colors.Color(0.8, 1, 0.8)  # light green
    BLUE = colors.blue
    LIGHT_BLUE = colors.Color(0.9, 0.9, 1)   # light blue

#####################################################################################################################
#####################################################################################################################
class ProductionProtocol:
    def __init__(self, 
                 protocol_number=0,
                 product_code="XXXX.Y.Z",
                 min_pn=0,
                 max_pn=0,
                 unrepairable_count=0,
                 repairable_count=0,
                 production_doc="XXXXYYYY_YYMMDD",
                 worker_name="XXXX",
                 check_date="",
                 note="",
                 tests=None):
        
        # Basic data
        self.protocol_number = protocol_number
        self.product_code = product_code
        self.min_pn = min_pn
        self.max_pn = max_pn
        self.total_count = max_pn - min_pn + 1
        self.unrepairable_count = unrepairable_count
        self.repairable_count = repairable_count
        self.ok_count = self.total_count - unrepairable_count - repairable_count
        self.production_doc = production_doc
        self.worker_name = worker_name
        self.check_date = check_date
        self.note = note

        # Process steps
        self.input_check = False
        self.additional_assembly = False
        self.cable_production = False
        self.cable_check = False
        self.isolation_measurement = False
        self.programming = False
        self.electrical_test = False
        self.coating = False
        self.component_fixing = False
        self.structural_assembly = False
        self.calibration = False
        self.product_marking = False
        self.finishing_work = False

        # Test processing
        if tests is not None:
            self.reports = tests
            self.test_names = list(next(iter(self.reports.values())).keys())
        else:
            self.reports = {}
            self.test_names = []
        
        # Layout settings
        self.left_margin = 15
        self.left_margin_inside = 17
        self.row_index_max = 280
        self.row_index = self.row_index_max
        self.background_color = Colors.WHITE

    #################################################################################################################
    def _create_frame(self, c, x, y, width, height, line_width=0.8, background_color=Colors.DEFAULT):
        """
        Create a frame with specified parameters.

        Args:
            c: Canvas object
            x (float): Left top corner X coordinate (mm)
            y (float): Left top corner Y coordinate (mm)
            width (float): Frame width (mm)
            height (float): Frame height (mm)
            line_width (float): Frame line width (mm)
            background_color (Colors): Background color from Colors enum
        """
        # Draw background if color specified
        if background_color and background_color.value:
            c.setFillColor(background_color.value)
            c.rect(x*mm, (y-height)*mm, width*mm, height*mm, fill=1, stroke=0)
    
        # Draw frame
        c.setStrokeColor(colors.black)
        c.setFillColor(colors.black)
        c.setLineWidth(line_width)
        c.rect(x*mm, (y-height)*mm, width*mm, height*mm, fill=0)

    def _write_text(self, c, text, x, y, bold=False, size=9, max_width=None, line_spacing=4):
        """
        Write text at specified position with formatting.

        Args:
            c: Canvas object
            text (str): Text to write
            x (float): X coordinate (mm)
            y (float): Y coordinate (mm)
            bold (bool): Whether text should be bold
            size (int): Font size
            max_width (float): Maximum text width in mm (None for no wrap)
            line_spacing (float): Space between lines in mm

        Returns:
            float: Y position of last line
        """
        font = "ArialBold" if bold else "Arial"
        c.setFont(font, size)
        
        if max_width is None:
            # Single line
            c.drawString(x*mm, y*mm, text)
            return y
        else:
            # Text wrapping
            max_width_pt = max_width * mm * 72 / 25.4
            words = text.split()
            lines = []
            current_line = []
            line_width = 0
            
            for word in words:
                word_width = c.stringWidth(word, font, size)
                if line_width + word_width <= max_width_pt:
                    current_line.append(word)
                    line_width += word_width + c.stringWidth(" ", font, size)
                else:
                    if current_line:
                        lines.append(" ".join(current_line))
                    current_line = [word]
                    line_width = word_width + c.stringWidth(" ", font, size)
            
            if current_line:
                lines.append(" ".join(current_line))
            
            # Write lines
            current_y = y
            for line in lines:
                c.drawString(x*mm, current_y*mm, line)
                current_y -= line_spacing
            
            return current_y

    def _create_checkbox(self, c, x, y, size=4, checked=False):
        """
        Create a checkbox.

        Args:
            c: Canvas object
            x (float): X coordinate (mm)
            y (float): Y coordinate (mm)
            size (float): Checkbox size (mm)
            checked (bool): Whether checkbox is checked
        """
        # Thicker frame
        c.setStrokeColor(colors.black)
        c.setLineWidth(0.8)
        c.rect(x*mm, y*mm, size*mm, size*mm)
        
        if checked:
            # Thicker checkmark
            c.setLineWidth(1.2)
            
            # Adjusted coordinates for more visible checkmark
            c.line((x+0.8)*mm, (y+size/2)*mm,
                (x+size/2)*mm, (y+0.8)*mm)
            c.line((x+size/2)*mm, (y+0.8)*mm,
                (x+size-0.8)*mm, (y+size-0.8)*mm)

    #################################################################################################################
    def _create_header(self, c):
        """
        Create header section of the first page.
        
        Args:
            c: Canvas object
        """
        # SIEMENS logo
        self._write_text(c, "SIEMENS", self.left_margin, self.row_index, bold=True, size=20)
        self.row_index -= 10

        # Document title above frame
        self._write_text(c, "Generované pomocou TMU_ProtocolGenerator.py", self.left_margin, self.row_index, bold=True, size=12)
        self.row_index -= 5

        # Main frame
        self._create_frame(c, self.left_margin, self.row_index, 180, 113, 1.5, background_color=self.background_color)
        self.row_index -= 7
        
        # Section title in frame
        self._write_text(c, "A1: Hlavička", self.left_margin_inside, self.row_index, bold=True, size=12)
        self.row_index -= 10

        # Header data
        data = [
            ["Číslo protokolu:",                      self.protocol_number],
            ["Kód produktu:",                         self.product_code],
            ["Rozsah výrobných čísel:",               f"V{self.min_pn:>06} - V{self.max_pn:>06}"],
            ["Počet - zadané do výroby:",             str(self.total_count)],
            ["Počet - zmätky neopraviteľné / PNR:",   str(self.unrepairable_count)],
            ["Počet - zmätky opraviteľné / PNR:",     str(self.repairable_count)],
            ["Počet - vyrobené OK:",                  str(self.ok_count)],
            ["Výrobná dokumentácia:",                 self.production_doc],
            ["Dátum kontroly:",                       self.check_date],
            ["Meno pracovníka:",                      self.worker_name],
            ["Poznámka:",                             self.note]
        ]

        # Draw data
        for i, (text, value) in enumerate(data):
            # Item text
            self._write_text(c, text, self.left_margin_inside, self.row_index, size=10)
            
            # Value frame
            if i in [4, 5]:  # Two frames in row for defect counts
                self._create_frame(c, 90, self.row_index+5, 60, 7, 0.5, background_color=Colors.LIGHT_BLUE)
                self._create_frame(c, 90+65, self.row_index+5, 30, 7, 0.5, background_color=Colors.LIGHT_BLUE)
                if value:
                    self._write_text(c, value, 93, self.row_index, size=10)
            elif i == 9:  # Worker name - longer frame
                self._create_frame(c, 90, self.row_index+5, 95, 7, 0.5, background_color=Colors.LIGHT_BLUE)
                if value:
                    self._write_text(c, value, 93, self.row_index, size=10)
            elif i == 10:  # Note - longer frame
                self._create_frame(c, 90, self.row_index+5, 95, 14, 0.5, background_color=Colors.LIGHT_BLUE)
                if value:
                    self._write_text(c, value, 93, self.row_index, size=10, max_width=32)
            else:  # Standard items
                self._create_frame(c, 90, self.row_index+5, 60, 7, 0.5, background_color=Colors.LIGHT_BLUE)
                if value:
                    self._write_text(c, value, 93, self.row_index, size=10)
            
            self.row_index -= 8

    def _create_processing(self, c):
        """
        Create processing section of the first page.
        
        Args:
            c: Canvas object
        """
        self.row_index -= 12
        # Main frame
        self._create_frame(c, self.left_margin, self.row_index, 180, 37, 1.5, background_color=self.background_color)
        self.row_index -= 7

        # Section title
        self._write_text(c, "A2: Spracovanie", self.left_margin_inside, self.row_index, bold=True, size=12)
        self.row_index -= 13
        
        # Processing data
        self._write_text(c, "Zaevidovanie do 006HMH HOW-zoznam:", self.left_margin_inside, self.row_index, size=10)
        self._create_frame(c, 90, self.row_index+5, 60, 7, 0.5, background_color=Colors.LIGHT_BLUE)
        self.row_index -= 8

        self._write_text(c, "Platnosť výstupnej kontroly:", self.left_margin_inside, self.row_index, size=10)
        self._create_frame(c, 90, self.row_index+5, 60, 7, 0.5, background_color=Colors.LIGHT_BLUE)
        self.row_index -= 8

    def _create_operations(self, c):
        """
        Create operations section of the first page.
        
        Args:
            c: Canvas object
        """
        self.row_index -= 5
        # Main frame
        self._create_frame(c, self.left_margin, self.row_index, 180, 95, 1.5, background_color=self.background_color)
        self.row_index -= 7

        # Section title
        self._write_text(c, "B1: Evidencia blokov operácií", self.left_margin_inside, self.row_index, bold=True, size=12)
        self.row_index -= 13
        
        # Operations data
        operations = [
            ["Vstupná kontrola:", self.input_check],
            ["Doosadenie, úprava DPS:", self.additional_assembly],
            ["Elektrické prepojenia - výroba:", self.cable_production],
            ["Elektrické prepojenia - kontrola:", self.cable_check],
            ["Meranie izolačných pevností:", self.isolation_measurement],
            ["Programovanie a konfigurácia:", self.programming],
            ["Elektrický test DPS:", self.electrical_test],
            ["Lakovanie a UV kontrola DPS:", self.coating],
            ["Fixácia komponentov na DPS:", self.component_fixing],
            ["Montáž konštrukčných prvkov:", self.structural_assembly],
            ["Kalibrácia:", self.calibration],
            ["Označenie polotovaru:", self.product_marking],
            ["Ukončovacie práce:", self.finishing_work]
        ]
        
        # Draw operations
        for text, value in operations:
            self._write_text(c, text, self.left_margin_inside, self.row_index, size=10)
            self._create_checkbox(c, 90, self.row_index, size=4, checked=value)
            self.row_index -= 6

    def _create_first_page(self, c):
        """
        Creates the first page of the protocol by combining header, processing and operations sections.
        
        Args:
            c: Canvas object
        """
        self._create_header(c)
        self._create_processing(c)
        self._create_operations(c)

    #################################################################################################################
    def _add_page(self, c):
        """Adds new page and resets row position."""
        c.showPage()
        self.row_index = self.row_index_max
    
    def _create_test_pages(self, c, start_pn):
        """
        Creates pages with test tables for 10 modules.
        """
        # First count tests with Report=true
        num_tests = sum(1 for test in self.reports[start_pn]["Tests"].values() if test["Report"])
        
        # Maximum tests per page
        max_tests_per_page = 44
        
        # Calculate number of needed pages
        num_pages = (num_tests + max_tests_per_page - 1) // max_tests_per_page
        
        # Create list of tests with Report=true
        tests_to_display = [
            name for name, test in self.reports[start_pn]["Tests"].items() 
            if test["Report"]
        ]
        
        # For each page
        for page_num in range(num_pages):
            if page_num > 0:
                c.showPage()  # New page
                
            # Reset position for new page
            self.row_index = 280
            
            # Page header
            self._write_text(c, "SIEMENS", self.left_margin, self.row_index, bold=True, size=20)
            self.row_index -= 15

            # Section title
            self._create_frame(c, self.left_margin, self.row_index, 180, 253, 1.5, background_color=self.background_color)
            self.row_index -= 7
            
            self._write_text(c, f"B2: Výsledky testov pre moduly V{start_pn:06d} - V{min(start_pn + 9, self.max_pn):06d}", 
                    self.left_margin + 2, self.row_index, bold=True, size=12)
            self.row_index -= 5

            # Table dimensions setup
            test_column = 65
            pn_column = 9
            spacing = 1
            row_height = 5
            header_height = 20

            # Initial x-positions for tests results columns
            x_test = self.left_margin + 2
            x_unit = x_test + test_column - 10
            x_results_start = x_test + test_column
            x_results = [x_results_start + (i * (pn_column + spacing)) for i in range(10)]

            # Vertical text for PN
            for i in range(10):
                pn = start_pn + i
                if pn > self.max_pn:
                    break
                
                c.saveState()
                c.translate((x_results[i] + 4)*mm, (self.row_index-header_height+5)*mm)
                c.rotate(90)
                c.setFont("ArialBold", 8)
                c.drawString(0, 0, f"V{pn:06d}")
                c.restoreState()
            
            self.row_index -= header_height
            
            # Determine test range for current page
            start_index = page_num * max_tests_per_page
            end_index = min((page_num + 1) * max_tests_per_page, len(tests_to_display))
            
            # Draw test rows for current page
            for test_name in tests_to_display[start_index:end_index]:
                # Test name
                self._write_text(c, test_name, x_test, self.row_index, size=7)
                
                if "Unit" in self.reports[start_pn]["Tests"][test_name]:
                    unit = self.reports[start_pn]["Tests"][test_name]["Unit"]
                    if unit:
                        self._write_text(c, f"[{unit}]", x_unit, self.row_index, size=7)
                        
                # Results for each module
                for i in range(10):
                    pn = start_pn + i
                    if pn > self.max_pn:
                        break
                        
                    test_data = self.reports[pn]["Tests"][test_name]
                    result = test_data["Passed"]
                    resultdesc = test_data["ResultDesc"]
                    
                    if result:
                        color = Colors.LIGHT_GREEN
                        if isinstance(resultdesc, (int, float)):
                            display_text = f"{resultdesc}"[:5]
                        else:
                            display_text = "PASS"
                    else:
                        color = Colors.LIGHT_RED
                        if isinstance(resultdesc, (int, float)):
                            display_text = f"{resultdesc}"[:5]
                        else:
                            display_text = "FAIL"

                    # Colored background for result
                    self._create_frame(c, x_results[i]-1, self.row_index+3, 
                                    pn_column, row_height-1, 0.3, background_color=color)
                    
                    # Result text
                    self._write_text(c, display_text, x_results[i], self.row_index, size=6)
                
                self.row_index -= row_height

    #################################################################################################################
    def create_pdf(self, filename):
        """Creates complete PDF protocol."""
        c = canvas.Canvas(filename, pagesize=A4)
        
        # Create first page
        self._create_first_page(c)
        
        # Check if there are tests to report
        has_reportable_tests = False
        for test in self.reports[self.min_pn]["Tests"].values():
            if test["Report"]:
                has_reportable_tests = True
                break
        
        # Create test pages if needed
        if has_reportable_tests:
            num_modules = self.max_pn - self.min_pn + 1
            num_pages = (num_modules + 9) // 10  # round up
            
            for page in range(num_pages):
                self._add_page(c)  # Always add new page for test pages
                start_pn = self.min_pn + (page * 10)
                self._create_test_pages(c, start_pn)
        
        c.save()

#####################################################################################################################
#####################################################################################################################
class JsonProcessor:
    def __init__(self, min_pn, max_pn, path="C:\\"):
        """
        Initialize JsonProcessor.

        Args:
            min_pn (int): Minimum production number
            max_pn (int): Maximum production number
            path (str): Base path for JSON files
        """
        self.min_pn = min_pn
        self.max_pn = max_pn
        self.reports = {}
        self.first_card_type = None
        self.first_tests_names = None
        self.path = path
        self.repairable_count = 0
        self.repairable_list = []
        self.unrepairable_count = 0
        self.unrepairable_list = []
        self.all_relevant_json_files = []
    
    def _get_list_of_all_json_files(self):
        """
        Create a list of JSON files, keeping only the latest version for each SN number.
        Returns a list of file paths containing only the most recent reports.
        """
        # Dictionary to store files grouped by SN number
        latest_records = {}
        
        # Regular expression pattern to extract SN number and timestamp
        pattern = r'(V\d{6})_(\d{8}_\d{6})'
        
        # Walk through directories and process JSON files
        for root, dirs, files in os.walk(self.path):
            for filename in files:
                if filename.endswith('.json'):
                    full_path = os.path.join(root, filename)
                    match = re.search(pattern, filename)
                    
                    if match:
                        sn_number = match.group(1)
                        timestamp_str = match.group(2)
                        
                        # Convert timestamp string to datetime object
                        timestamp = datetime.strptime(timestamp_str, '%Y%m%d_%H%M%S')
                        
                        # Update dictionary if this is a new SN or if this record is newer
                        if sn_number not in latest_records or timestamp > latest_records[sn_number][0]:
                            latest_records[sn_number] = (timestamp, full_path)
        
        # Extract only the file paths from the latest records, sorted by SN number
        result_files = [record[1] for record in sorted(latest_records.values(), 
                                                    key=lambda x: x[1])]
        
        return result_files

    def _check_all_tests(self, data, filename, pn):
        """
        Check if all tests are done and passed.

        Args:
            data (dict): Test data from JSON
            filename (str): JSON filename
            pn (str): Production number

        Returns:
            bool: True if checks pass, False otherwise
        """
        if not data["AllTestsDone"]:
            print(f"Nevykonané niektoré testy v súbore {filename}")
            return False
        
        if not data["Passed"]:
            print(f"Neúspešné niektoré testy v súbore {filename}")

            if get_user_choice("Želáte si pokračovať?", default=False):
                if get_user_choice(f"Opraviteľná závada? - {pn}", default=True):
                    self.repairable_count += 1
                    self.repairable_list.append(pn)
                else:
                    self.unrepairable_count += 1
                    self.unrepairable_list.append(pn)
            else:
                return False
            print("")
        return True

    def _check_card_type(self, card_type, filename):
        """
        Verify card type consistency across files.

        Args:
            card_type (str): Card type from current file
            filename (str): Current filename

        Returns:
            bool: True if card type matches, False otherwise
        """
        if self.first_card_type is None:
            self.first_card_type = card_type
            return True
        
        if card_type != self.first_card_type:
            print(f"Rozdielny typ karty v súbore {filename}")
            print(f"Očakávaný typ: {self.first_card_type}")
            print(f"Nájdený typ: {card_type}")
            return False
        
        return True

    def _check_test_names(self, tests, filename):
        """
        Verify test names consistency across files.

        Args:
            tests (dict): Tests from current file
            filename (str): Current filename

        Returns:
            bool: True if test names match, False otherwise
        """
        current_test_names = set(tests.keys())
        
        if self.first_tests_names is None:
            self.first_tests_names = current_test_names
            return True
        
        if current_test_names != self.first_tests_names:
            print(f"Rozdielne názvy testov v súbore {filename}")
            extra_tests = current_test_names - self.first_tests_names
            missing_tests = self.first_tests_names - current_test_names
            
            if extra_tests:
                print(f"Testy navyše: {sorted(extra_tests)}")
            if missing_tests:
                print(f"Chýbajúce testy: {sorted(missing_tests)}")
            return False
        
        return True
    
    def process_files(self):
        """
        Spracovanie JSON súborov z určeného adresára a podadresárov.

        Returns:
            bool: True ak je spracovanie úspešné, False inak
        """
        if not os.path.exists(self.path):
            print(f"Adresár {self.path} neexistuje!")
            return False

        all_json_files = self._get_list_of_all_json_files()
            
        for pn in range(self.min_pn, self.max_pn + 1):
            found = False
            for full_path in all_json_files:
                filename = os.path.basename(full_path)
                if f"V{pn:>06}" in filename:
                    self.all_relevant_json_files.append(full_path)
                    try:
                        with open(full_path, 'r', encoding='utf-8') as f:
                            data = json.load(f)
                            
                            if data["SafeBytes"]["SN"] != pn:
                                print(f"Nesedí SN v súbore {filename}")
                                return False

                            if not self._check_all_tests(data, full_path, f"V{pn:>06}"):
                                return False

                            if not self._check_card_type(data["CardTypeName"], full_path):
                                return False
                            
                            tests = {}
                            for test in data["Tests"]:
                                test_data = {
                                    "Passed": test["Passed"],
                                    "Report": test["Report"],
                                    "ResultDesc": test["ResultDesc"],
                                    "Unit": test.get("Unit", "")
                                }
                                
                                if "Min" in test:
                                    test_data["Min"] = test["Min"]
                                if "Max" in test:
                                    test_data["Max"] = test["Max"]
                                
                                tests[test["Name"]] = test_data

                            if not self._check_test_names(tests, full_path):
                                return False
                            
                            self.reports[pn] = {
                                "Tests": tests,
                                "Passed": data["Passed"],
                                "AllTestsDone": data["AllTestsDone"],
                                "UserName": data["UserName"],
                                "CardTypeName": data["CardTypeName"]
                            }
                            
                            found = True
                            break

                    except Exception as e:
                        print(f"Chyba pri čítaní súboru {filename}: {str(e)}")
                        return False
            
            if not found:
                print(f"Nenašiel sa súbor pre V{pn:>06}")
                return False

        return True

    def get_reports(self):
        """Return dictionary with all processed reports."""
        return self.reports
    
    def get_num_of_repairable_pieces(self):
        """Return number of repairable pieces."""
        return self.repairable_count
    
    def get_num_of_unrepairable_pieces(self):
        """Return number of unrepairable pieces."""
        return self.unrepairable_count
    
    def get_list_of_repairable_pieces(self):
        """Return list of repairable piece numbers."""
        return self.repairable_list
    
    def get_list_of_unrepairable_pieces(self):
        """Return list of unrepairable piece numbers."""
        return self.unrepairable_list
    
    def get_card_type(self):
        """Return card type."""
        return self.first_card_type
    
    def get_list_of_relevant_json_files(self):
        """Return list of all relevant json files used for protocol."""
        return self.all_relevant_json_files
    
#####################################################################################################################
#####################################################################################################################    
def main():
    print("Spustené generovanie výrobného protokolu.\n")
    
    # Default path to reports
    default_path = "C:\\Reports_TUS"

    # Get path to reports
    if not get_user_choice(f"Použiť defaultnú cestu ku reportom? {default_path}", default=True):
        # Setup file dialog
        root = tk.Tk()
        root.withdraw()
        root.lift()
        root.focus_force()

        # Open dialog for selecting path to reports
        default_path = filedialog.askdirectory(
            parent=root,
            initialdir=default_path,
            title="Vyber priečinok s umiestnením reportov"
        )
        print(f"Zvolená cesta ku reportom: {default_path}")

    # Get protocol number
    protocol_number = str(input("\nZadaj číslo protokolu: "))

    # Get Windows user
    read_username = get_display_name()

    # Get responsible person
    if get_user_choice(f"Zodpovedná osoba? - {read_username}", default=True):
        worker_name = read_username
    else:
        worker_name = str(input("Zadaj meno zodpovednej osoby: "))
    
    # Get documentation number
    production_doc = input("Zadaj číslo výrobnej dokumentácie (XXXXYYYY_YYMMDD): ")

    # Get production number range
    print("Zadaj rozsah výrobných čísel:")
    min_pn = int(input("Min: V"))
    max_pn = int(input("Max: V"))

    # Get note
    note = input("Zadaj poznámku (nepovinné): ")

    print("")

    # Create JsonProcessor instance
    json_processor = JsonProcessor(min_pn, max_pn, path=default_path)

    # Process JSON files
    if not json_processor.process_files():
        print('Neúspešné spracovanie json súborov.')
        input('Stlač ENTER pre ukončenie!')
        exit()

    # Get product code from json
    product_code = json_processor.get_card_type()

    # Get repairable pieces info
    if json_processor.get_num_of_repairable_pieces():
        repairable_pcs_count = json_processor.get_num_of_repairable_pieces()
        repairable_pcs_list = json_processor.get_list_of_repairable_pieces()
    else:
        repairable_pcs_count = 0
        repairable_pcs_list = ""

    # Get unrepairable pieces info
    if json_processor.get_num_of_unrepairable_pieces():
        unrepairable_pcs_count = json_processor.get_num_of_unrepairable_pieces()
        unrepairable_pcs_list = json_processor.get_list_of_unrepairable_pieces()
    else:
        unrepairable_pcs_count = 0
        unrepairable_pcs_list = ""
           
    try:
        # Create protocol instance
        protocol = ProductionProtocol(
            protocol_number=protocol_number,
            product_code=product_code,
            min_pn=min_pn,
            max_pn=max_pn,
            unrepairable_count=unrepairable_pcs_count,
            repairable_count=repairable_pcs_count,
            production_doc=production_doc,
            worker_name=worker_name,
            check_date=datetime.now().strftime("%d.%m.%Y"),
            note=note,
            tests=json_processor.get_reports()
        )

        # Set protocol properties
        protocol.input_check =              get_user_choice("Vstupná kontrola?",                    default=True)
        protocol.additional_assembly =      get_user_choice("Doosadenie, úprava DPS?",              default=True)
        protocol.cable_production =         get_user_choice("Elektrické prepojenia - výroba?",      default=False)
        protocol.cable_check =              get_user_choice("Elektrické prepojenia - kontrola?",    default=False)
        protocol.isolation_measurement =    get_user_choice("Meranie izolačných pevností?",         default=True)
        protocol.programming =              get_user_choice("Programovanie a konfigurácia?",        default=True)
        protocol.electrical_test =          get_user_choice("Elektrický test DPS?",                 default=True)
        protocol.coating =                  get_user_choice("Lakovanie a UV kontrola DPS?",         default=True)
        protocol.component_fixing =         get_user_choice("Fixácia komponentov na DPS?",          default=True)
        protocol.structural_assembly =      get_user_choice("Montáž konštrukčných prvkov?",         default=False)
        protocol.calibration =              get_user_choice("Kalibrácia?",                          default=False)
        protocol.product_marking =          get_user_choice("Označenie polotovaru?",                default=True)
        protocol.finishing_work =           get_user_choice("Ukončovacie práce?",                   default=True)

        # Setup file dialog
        root = tk.Tk()
        root.withdraw()
        root.lift()
        root.focus_force()

        # Open dialog for selecting save directory
        output_dir = filedialog.askdirectory(
            parent=root,
            initialdir=default_path,
            title="Vyber priečinok pre uloženie protokolu"
        )
        
        # Use default path if no directory selected
        if not output_dir:
            output_dir = default_path
        
        # Create directory if it doesn't exist
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
        
        # Create file path
        output_file = os.path.join(output_dir, f"Protocol_{protocol_number}_{product_code}.pdf")
        
        # Create PDF
        protocol.create_pdf(output_file)
        
        print(f"\nProtokol {protocol_number} úspešne vytvorený.")

        # Add footer
        add_footer(output_file, protocol_number)
        print("Úspešné pridané päty.")

        # Add comments if needed
        if unrepairable_pcs_list or repairable_pcs_list:
            add_error_comments_to_pdf(
                output_file, 
                repairable_pcs_list,
                unrepairable_pcs_list
            )
            print("Úspešné pridané komentáre.")

        # Add attachments
        json_files = json_processor.get_list_of_relevant_json_files()
        add_attachments_to_pdf(output_file, json_files)
        print("Úspešne pridané prílohy")

        if get_user_choice("\nŽeláte si otvoriť protokol?", default=True):
            open_file(output_file)

    except Exception as e:
        print(f"Chyba pri vytváraní protokolu: {e}")
        input('Stlač ENTER pre ukončenie!')
        exit()

if __name__ == '__main__':
    main()