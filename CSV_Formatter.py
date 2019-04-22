# This script is designed to take raw CAN data generated as a CSV file and format it to be read by another Python script

import csv
import pandas as pd
import string
import xlsxwriter


#This file path should be where the raw CAN data output file is stored. Replace accordingly
csv_path = r"C:\Users\cnduaguibe\Downloads\ACLTOP550CTS PT Trial 1.csv"

# Create an Excel Workbook and add a worksheet
workbook = xlsxwriter.Workbook("Data Reformatted from CSV File")
worksheet = workbook.add_worksheet("Formatted for Script")
col_headers = ("Index", "Time(s)", "ID", "RTR", "DLC", "Payload", "Pump Events", "P0", "P1", "P2", "P3", "IPCID",
               "IPC Command Translation", "Phase", "Milestone")
row = 6
col = 0


# Writes out the headers of each column in the newly-created Excel worksheet
for header in col_headers:
    worksheet.write(row, col, header)
    col += 1


# Opens the CSV file, reads line by line, filtering out empty lines and the first three lines of the document which do
# not contain data values. Saves each column of data values as a variable for later porting into the new Excel sheet
analyzer_info = list()
raw_time = list()
ID = list()
RTR = list()
DLC = list()
Payload = list()
with open(csv_path) as opened_file:
    file_reader = csv.reader(opened_file)
    for line in file_reader:
        if len(line) > 2:
            if line[0] != "0:00.000.000" and line[1] != "ID":
                list.append(raw_time, line[0])
                list.append(ID, line[1])
                list.append(RTR, line[2])
                list.append(DLC, line[3])
                list.append(Payload, line[4])
            elif line[0] == "0:00.000.000":
                list.append(analyzer_info, line[4])
            else:
                continue


# Converts RTR and DLC values from strings to integers, then reassigns the RTR and DLC variables to the integers
RTR_int = list()
for value in RTR:
    if value != '':
        list.append(RTR_int, int(value))
RTR = RTR_int

DLC_int = list()
for value in DLC:
    if value != '':
        list.append(DLC_int, int(value))
DLC = DLC_int


# Fills the Index column in the worksheet with a number starting at 1 and increasing by 1 each row down
row = 7
col = 0
counter = 1
while row < (len(ID) + 7):
    worksheet.write(row, col, counter)
    row += 1
    counter += 1


# Converts timestamp strings from CSV file into integers, rounds to 6 decimal places, sets each as time from start(0)
adj_time = list()
for timestamp in raw_time:
    min_1 = int(timestamp[0]) * 60
    sec_10 = int(timestamp[2]) * 10
    sec_1 = int(timestamp[3]) * 1
    msec_100 = int(timestamp[5]) * 0.1
    msec_10 = int(timestamp[6]) * 0.01
    msec_1 = int(timestamp[7]) * 0.001
    usec_100 = int(timestamp[9]) * 0.0001
    usec_10 = int(timestamp[10]) * 0.00001
    usec_1 = int(timestamp[11]) * 0.000001
    int_time = (min_1 + sec_10 + sec_1 + msec_1 + msec_10 + msec_1 + usec_100 + usec_10 + usec_1)
    rounded_time = round(int_time, 6)
    list.append(adj_time, rounded_time)
final_time = list()
for time in adj_time:
    zero_time = (time - adj_time[0])
    rounded_time2 = round(zero_time, 6)
    list.append(final_time, rounded_time2)


# Defines a general function for taking a list of CSV data values and porting them into a column in an Excel sheet
def porting_function(row_start, col_start, list_name):
    for info in list_name:
        worksheet.write(row_start, col_start, info)
        row_start += 1
    return


# Ports the three lines of analyzer info (read from the CSV earlier in the script) and writes to the Excel sheet
porting_function(0, 5, analyzer_info)


# Takes the reformatted timestamps from the CSV file and fills in the Time(s) column in the Excel sheet
porting_function(7, 1, final_time)


# Takes the ID, RTR, DLC, and Payload lists from the CSV file and fills in the respective columns in the Excel sheet
porting_function(7, 2, ID)
porting_function(7, 3, RTR)
porting_function(7, 4, DLC)
porting_function(7, 5, Payload)

# Takes ID string from CSV file, slices it, converts hexadecimal to decimal, then a bitshift right returns the IPCID#
def id_to_ipcid_converter(id):
    global ipcid
    ipcid = list()
    for string in id:
        if string != '':
            id_slice = string[2:10]
            id_slice_int = int(id_slice, 16)
            bitshift_id = id_slice_int >> 4
            list.append(ipcid, bitshift_id)
    return ipcid


# Runs the id to ipcid converter function on the ID list from CSV and fills the ID column in the Excel sheet
id_to_ipcid_converter(ID)
porting_function(7, 11, ipcid)

# File path where the IPCID and IPCCOMMAND reference sheet is stored, replace accordingly
ref_path = r"C:\Users\cnduaguibe\Desktop\CAN Command IPCID Reference Sheet.xlsx"

# Creates a dictionary with values read from two columns in an Excel sheet
def excel_dict_function(file_path, sheet_name, column_header1, column_header2):
    data_frame = pd.read_excel(file_path, sheet_name=sheet_name)
    ipcid_list = list()
    ipccommand_list = list()
    for value in data_frame[column_header1]:
        list.append(ipcid_list, value)
    for command in data_frame[column_header2]:
        list.append(ipccommand_list, command)
    dictionary = dict(zip(ipcid_list, ipccommand_list))
    return dictionary


# Matches a given list of keys to a dictionary's keys and creates a list of values found for the given list of keys
def ipcid_to_ipccommand_function(a_list, dictionary):
    global ipccommand
    ipccommand = list()
    for number in a_list:
        if number in dictionary:
            list.append(ipccommand, dictionary[number])
    return ipccommand


# Creates the IPCID and IPCCOMMAND dictionary, uses it to list commands, and fills in the IPCCOMMAND column in the sheet
ref_dict = excel_dict_function(ref_path, 'Ref', 'IPCID', 'IPCCOMMAND')
ipcid_to_ipccommand_function(ipcid, ref_dict)
porting_function(7, 12, ipccommand)


# Formatting the Excel sheet: Making column headers bold
bold_font = workbook.add_format({'bold': True})
count = 0
while count <= 14:
    cell = string.ascii_uppercase[count] + '7'
    worksheet.write(cell, col_headers[count], bold_font)
    count += 1

# Formatting the Excel sheet: Adjusts width of columns for easier viewing of all values
worksheet.set_column(2, 2, 10)
worksheet.set_column(5, 5, 20)
worksheet.set_column(6, 6, 11.5)
worksheet.set_column(7, 10, 2)
worksheet.set_column(12, 12, 42)
worksheet.set_column(13, 13, 9)

workbook.close()
