# -*- coding: utf-8 -*-


import sys
from datetime import datetime

from openpyxl import Workbook

def read(fsave):
    skiplines = 16
    lines = [l.strip().split(',') for l in fsave.readlines()]
    channel_atts = []
    data = []
    i = skiplines
    
    while lines[i][0] != "Calc settings":
        channel_atts.append(lines[i])
        i += 1
    i += 5 # skip Data and the two headers
    while i < len(lines):
        data.append(lines[i])
        i += 1
        
    return channel_atts, data

def parse_datetime(time_str):
    return datetime.strptime(time_str, "%Y/%m/%d %H:%M:%S")

        
def parse(channel_atts, data):
    lines = []
    # Write the header
    colnames = ["date", "time", "test-time"]
    for atts in channel_atts:
        name = atts[1][1:-1]
        unit = atts[9][1:-1].replace("deg", "º")
        colnames.append("{} ({})".format(name, unit))
#    header = delimiter.join(colnames)
    
    lines.append(colnames)
        
    dt0 = parse_datetime(data[0][1])
    for d in data:       
        dt = parse_datetime(d[1])
        
        date = dt.date()
        clock = dt.time()
        test_time = dt - dt0
        
        fields = [date, clock, test_time]
        for c, atts in enumerate(channel_atts):
            value = d[3+c]
            value = value.replace("+", "")
            value = value.replace(" ", "")
#            value = value.replace(".", ",")
            fields.append(float(value))
#        l = delimiter.join(fields)
        lines.append(fields)
        
    return lines
        

        
def write_csv(fname, lines, delimiter = ", "):
    f = open(fname, 'w')
    for l in lines:
        f.write(delimiter.join([str(field) for field in l]) + "\n")
    f.close()
        
def write_excel(fname, lines):
    wb = Workbook()
    ws = wb.active
    for l in lines:
        ws.append(l)
    wb.save(fname)
    
# Get filename
if len(sys.argv) == 1:
    fname = input("File > ")
elif len(sys.argv) > 1:
    fname = sys.argv[1]
    
print(sys.argv)

# Load and readfile
fname = 'temps.csv' #overwritten for testing
fload = open(fname, 'r')
channel_atts, data = read(fload)

# Parse data
lines = parse(channel_atts, data)
# Save as CSV
f_csv = 'clean_' + fname
write_csv(f_csv, lines, delimiter = "\t ")

# Save as .xlsx

f_excel = "excel_" + fname.replace(".csv", ".xlsx")
write_excel(f_excel, lines)
