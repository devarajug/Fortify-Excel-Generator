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
    
    def zipXmlData(self):
        tree = ET.parse(self.file)
        root = tree.getroot()
        for issue in root.findall("ReportSection/SubSection/IssueListing/Chart/GroupingSection/Issue"):
            auditor_comment_value = ""
            for i in range(len(issue)):
                if issue[i].tag == "Category":
                    category_value = issue[i].text
                elif issue[i].tag == "Kingdom":
                    kingdom_value = issue[i].text
                elif issue[i].tag == "Friority":
                    priority_value = issue[i].text
                elif issue[i].tag == "Abstract":
                    abstract_value = issue[i].text
                elif issue[i].tag == "Tag":
                    status_value = issue[i][1].text
                elif issue[i].tag == "Comment":
                    auditor_comment_value+= str(issue[i][0].text) +"\n"+ str(issue[i][1].text)
                elif issue[i].tag == "Source":
                    source_file_path_value = issue[i][1].text
                    source_line_no_value = issue[i][2].text
                elif issue[i].tag == "Primary":
                    sink_file_path_value = issue[i][1].text
                    sink_line_no_value = issue[i][2].text

            self.category.append(category_value)
            self.kingdom.append(kingdom_value)
            self.priority.append(priority_value)
            self.abstract.append(abstract_value)
            self.source_file_path.append(source_file_path_value)
            self.source_file_num.append(source_line_no_value)
            self.sink_file_path.append(sink_file_path_value)
            self.sink_file_no.append(sink_line_no_value)
            self.status.append(status_value)
            self.auditor_comment.append(auditor_comment_value)
        
        return zip(
            self.category,
            self.priority,
            self.kingdom,
            self.abstract,
            self.source_file_path,
            self.source_file_num,
            self.sink_file_path,
            self.sink_file_no,
            self.status,
            self.auditor_comment
        )
    
    


