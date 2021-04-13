#!/usr/bin/env python3

from openpyxl.chart.series import Series
from openpyxl.utils import cell

from config import *
import pickle
from openpyxl import Workbook
from openpyxl.chart import (
    LineChart,
    Reference, 
)

class College:
    
    def __init__(self, col):
        self.total_wams = {}
        self.std_count = {}
        self.wam_trend = {}
        self.add_college(col)
        for term in self.total_wams.keys():
            # minimum number of residents in term before adding
            if self.std_count[term] > 10:
                self.wam_trend[term] = self.total_wams[term]/self.std_count[term]
            else:
                self.wam_trend[term] = 0
        
        
    def add_college(self, college):
        for student in college.values():
            for (term, wam) in student.wams.items():
                if term not in self.total_wams:
                    self.total_wams[term] = 0
                    self.std_count[term] = 0
                if wam:
                    self.total_wams[term] += float(wam)
                    self.std_count[term] += 1
                    
        

last_mod_load = []
college_data = {}
all_terms = {}


# Get pickled college data
try:
    with open(f'cache/data.pkl', 'rb') as f:
        last_mod_load = pickle.load(f)
        college_data = pickle.load(f)
        all_terms = pickle.load(f)
except Exception:
    print("Failed to unpickle cached data. Please run main college academics script again.\n")
    exit()
    
    
# Set up workbook to contain data and charts
wb = Workbook()
ws = wb.active 
ws.append(["College", "Term", "WAM"])

row_start = 2
chart_row = 2
chart_col = 6

# load college data into spreadsheet and charts
for c in college_data:
    college = College(college_data[c])
    
    
    
    for (term, wam) in sorted(college.wam_trend.items()):
        if wam == 0:
            del college.wam_trend[term]
        else:
           ws.append([c, convert_term_name(term), wam])

    
    
    ch = LineChart()
    ch.title = f'{c} WAM'
    ch.style = 4
    ch.y_axis.title = "WAM"          

    
    
    # add data to chart
    ref = Reference(ws, min_col=3, min_row=row_start, max_row=row_start + len(college.wam_trend)-1)
    ch.add_data(ref, titles_from_data=False)
    ch.legend = None
    
    # x axis labels
    terms = Reference(ws, min_col= 2, min_row=row_start, max_row=row_start + len(college.wam_trend)-1)
    ch.set_categories(terms)
    
    ser = ch.series[0]
    ser.graphicalProperties.line.solidFill = college_colours[c]
    ser.marker.symbol = "circle"
    ser.marker.graphicalProperties.line.solidFill = "555555"
    ser.marker.graphicalProperties.noFill = True
    
    # 3 Charts per column
    if chart_row // 15 > 3:
        chart_col += 9
        chart_row = 2
    
    ws.add_chart(ch, f'{cell.get_column_letter(chart_col)}{chart_row}')
    
    row_start += len(college.wam_trend)
    chart_row += 15

    wb.save("College_Stats.xlsx")