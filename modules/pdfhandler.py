# extract_pdf_to_excel/modules/pdfhandler.py
from pathlib import Path
from spire.ocr import *
from spire.pdf import *

from . import core

class SpirePdf:
    '''This class holds a number of methods, using Python's SpirePDF
    and SpireOCR modules, which enable users to:
        - Load a PDF document
        - Convert the PDF to an image
        - Scan the image using Optical Character Recognition (OCR)
        - Extract relevant text from the scanned image
        - Run other, smaller helper methods
    '''
    def __init__(self, filepath, columns):
        self.filepath = filepath
        self.columns = columns
        self.pdf = self.check_and_load_filepath(filepath)
        if self.pdf is None:
            raise ValueError(
                "Class cannot be instantiated: problem with PDF path."
            )
    
    def check_and_load_filepath(self, filepath):
        # Check if PDF file exists and is the correct filetype.
        f = Path(filepath)
        if f.is_file() and f.suffix.lower().endswith(".pdf"):
            # Load the PDF document.
            pdf = PdfDocument()
            pdf.LoadFromFile(filepath)
        else:
            pdf = None
            # raise ValueError(
                # "Class cannot be instantiated: problem with PDF path.")

        return pdf

    def cols(self):
        # Get the number of columns.
        return len(self.columns)
    
    def name(self):
        # Get the name of the PDF file (not the full path).
        return Path(self.filepath).name

    def pdf_page_count(self):
        # Get the number of pages in the PDF.
        return self.pdf.Pages.Count

    def scanner_init(self, lang, model):
        # Initialize an OCR scanner object and configure options.
        s = OcrScanner()
        conf_options = ConfigureOptions()
        
        # Identify the language of the text to be scanned.
        conf_options.Language = lang
        # Set the path of the OCR model to be used.
        conf_options.ModelPath = model

        # Return configured OCR scanner object.
        s.ConfigureDependencies(conf_options)
        return s

    def save_as_img(self, page_index, img_path):
        # Save the provided page number from the PDF as an image.
        image = self.pdf.SaveAsImage(page_index)
        image.Save(img_path)

    
    def scanner_to_text(self, scan, img):
        # Scan an image and return the text contained within it
        # as a string.
        scan.Scan(img)
        return scan.Text.ToString()

    def split_scanned_text(self, scanned_text, delim):
        # Split a string of text into list on a given delimiter.
        return scanned_text.split(delim)

    def max_column_header(self, text_list):
        # Find the column header with the highest index from the
        # scanned PDF text using the column headers defined when the 
        # object was instantiated.
        headers = []
        for i, item in enumerate(text_list):
            # Ignore case sensitivity/special characters and find the
            # header (sometimes image text isn't translated 100% 
            # accurately).
            if core.list_contains_text(self.columns, item):
                headers.append(text_list.index(item))
        
        # Return the last/highest index found.
        return max(headers)
    
    def close(self):
        # Close the PdfDocument object.
        self.pdf.Close()

# Functions outside the SpirePdf class.
def load_pdf_file(path):
    # Load a PDF file using spire.pdf and return the pdf object.
    doc = PdfDocument()
    doc.LoadFromFile(path)

    return doc

def split_pdf_on_range(doc, path, start, end):
    # Given a loaded PdfDocument object, extract a range of pages from
    # the PDF document and save them as a separate PDF document.
    new_doc = PdfDocument()
    new_doc.InsertPageRange(doc, start, end)
    new_doc.SaveToFile(path)
    new_doc.Close()
