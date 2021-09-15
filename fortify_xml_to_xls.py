import pandas as pd
from sys import argv
from os.path import splitext
from openpyxl import workbook
import xml.etree.ElementTree as ET
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Alignment, PatternFill

class XmlToXls:

    def __init__(self, in_file, out_file):
        self.category = []
        self.priority = []
        self.kingdom = []
        self.abstract = []
        self.source_file_path = []
        self.source_file_num = []
        self.sink_file_path = []
        self.sink_file_no = []
        self.status = []
        self.auditor_comment = []
        self.in_file = in_file
        self.outfile = out_file
        

