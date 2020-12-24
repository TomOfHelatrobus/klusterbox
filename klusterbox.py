"""
 _   _ _                             _
| |/ /| |              _            | |
| | / | | _   _  ___ _| |_ ___  _ _ | |_   __  _  __
|  (  | || | | |/ __/_   _| __|| /_/|   \ /  \\ \/ /
| | \ | |\ \_| |\__ \ | | | _| | |  | () | () |)  (
|_|\_\|_| \____|/___/ |_| |___||_|  |___/ \__//_/\_\

Klusterbox
Copyright 2019 Thomas Weeks

Non-standard libraries: (located in requirements.txt file)
chardet==3.0.4
et-xmlfile==1.0.1
jdcal==1.4.1
openpyxl==3.0.3
pdfminer.six==20181108
pycryptodome==3.9.7
PyPDF2==1.26.0
six==1.14.0
sortedcontainers==2.1.0
pillow==6.0.0

to package with pyinstaller:
make sure that the kb_sub folder (with images) is in the same directory as the source file then input -
pyinstaller -w -F --icon kb_sub/kb_images/kb_icon2.ico klusterbox_v3-002.py

Caution: To ensure proper operation of Klusterbox, make sure to keep the Klusterbox application and the kb_sub folder
in the same folder.

For the newest version of Klusterbox, visit www.klusterbox.com/download. The source code is also available there.

This version of Klusterbox is being released under the GNU General Public License version 3.
"""
# version variables
version = "3.006"
release_date = "December 21, 2020"

# Standard Libraries
from tkinter import *
from tkinter import messagebox
from tkinter import filedialog
from tkinter import simpledialog
from tkinter import ttk
from datetime import datetime
from datetime import timedelta
import sqlite3
from operator import itemgetter
import os
import shutil
import csv
import sys
import subprocess
import io
from io import StringIO  # change from cStringIO to io for py 3x
import time
# Pillow Library
from PIL import ImageTk, Image  # Pillow Library
# Spreadsheet Libraries
from openpyxl import Workbook
from openpyxl.styles import NamedStyle, Font, Border, Side, Alignment, PatternFill
from openpyxl.worksheet.pagebreak import Break
# PDF Converter Libraries
import chardet
from pdfminer.pdfparser import PDFParser
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter, resolve1
from pdfminer.converter import TextConverter, PDFPageAggregator
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
# PDF Splitter Libraries
from PyPDF2 import PdfFileReader, PdfFileWriter

def inquire(sql):
    db = sqlite3.connect("kb_sub/mandates.sqlite")
    cursor = db.cursor()

    try:
        cursor.execute(sql)
        results = cursor.fetchall()
        return results
    except:
        messagebox.showerror("Database Error",
                             "Unable to access database.\n"
                             "\n Attempted Query: {}".format(sql))
    db.close()


def commit(sql):
    db = sqlite3.connect("kb_sub/mandates.sqlite")
    cursor = db.cursor()
    try:
        cursor.execute(sql)
        db.commit()
        db.close()
    except:
        messagebox.showerror("Database Error",
                             "Unable to access database.\n"
                             "\n Attempted Query: {}".format(sql))


def dt_converter(string):  # converts a string of a datetime to an actual datetime
    dt = datetime.strptime(string, '%Y-%m-%d %H:%M:%S')
    return dt

def front_window(self):  # Sets up a tkinter page with buttons on the bottom
    if self != "none": self.destroy()  # close out the previous frame
    F = Frame(root)  # create new frame
    F.pack(fill=BOTH, side=LEFT)
    buttons = Canvas(F)  # button bar
    buttons.pack(fill=BOTH, side=BOTTOM)
    # link up the canvas and scrollbar
    S = Scrollbar(F)
    C = Canvas(F, width=1600)
    S.pack(side=RIGHT, fill=BOTH)
    C.pack(side=LEFT, fill=BOTH)
    S.configure(command=C.yview, orient="vertical")
    C.configure(yscrollcommand=S.set)
    # link the mousewheel - implementation varies by platform
    if sys.platform == "win32":
        C.bind_all('<MouseWheel>', lambda event: C.yview_scroll(int(-1 * (event.delta / 120)), "units"))
    elif sys.platform == "linux":
        C.bind_all('<Button-4>', lambda event: C.yview('scroll', -1, 'units'))
        C.bind_all('<Button-5>', lambda event: C.yview('scroll', 1, 'units'))
    # create the frame inside the canvas
    FF = Frame(C)
    C.create_window((0, 0), window=FF, anchor=NW)
    return F, S, C, FF, buttons
    # page contents - then call rear_window(wd)


def rear_window(wd):  # This closes the window created by front_window()
    root.update()
    wd[2].config(scrollregion=wd[2].bbox("all"))
    mainloop()

def get_custom_nsday(): # get ns day color configurations from dbase and make dictionary
    sql = "SELECT * FROM ns_configuration"
    ns_results = inquire(sql)
    ns_dict = {}  # build dictionary for ns days
    days = ("sat", "mon", "tue", "wed", "thu", "fri")
    for r in ns_results:  # build dictionary for rotating ns days
        ns_dict[r[0]] = r[2]# build dictionary for ns fill colors
    for d in days:  # expand dictionary for fixed days
        ns_dict[d] = "fixed: " + d
    ns_dict["none"] = "none"  # add "none" to dictionary
    return ns_dict

def rpt_impman(list_carrier):
    date = g_date[0]
    dates = []  # array containing days.
    if g_range == "week":
        for i in range(7):
            dates.append(date)
            date += timedelta(days=1)
    if g_range == "day": dates.append(d_date)
    if g_range == "week":
        sql = "SELECT * FROM rings3 WHERE rings_date BETWEEN '%s' AND '%s' ORDER BY rings_date, carrier_name" \
              % (g_date[0], g_date[6])
    else:
        sql = "SELECT * FROM rings3 WHERE rings_date = '%s' ORDER BY rings_date, " \
              "carrier_name" \
              % (d_date)
    rings = inquire(sql)
    sql = "SELECT * FROM tolerances"  # get tolerances
    tolerances = inquire(sql)
    ot_own_rt = tolerances[0][2]
    ot_tol = tolerances[1][2]
    av_tol = tolerances[2][2]  # get tolerances
    daily_list = []  # array
    candidates = []
    dl_nl = []
    dl_wal = []
    dl_otdl = []
    dl_aux = []
    rec = ""
    weekly_summary = []
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")  # create a file name
    filename = "report_improper_mandates" + "_" + stamp + ".txt"
    if os.path.isdir('kb_sub/report') == False:  # create a directory if it does not exist
        os.makedirs('kb_sub/report')
    report = open('kb_sub/report/' + filename, "w")  # create text document
    report.write("Improper Mandates Report\n")
    for day in dates:
        report.write('\n\n   Showing results for:\n')
        report.write('      Station: {}\n'.format(g_station))
        f_date = day.strftime("%A  %b %d, %Y")
        report.write('      Date: {}\n'.format(f_date))
        report.write('      Pay Period: {}\n\n'.format(pay_period))
        del daily_list[:]
        del dl_nl[:]
        del dl_wal[:]
        del dl_otdl[:]
        del dl_aux[:]
        # create a list of carriers for each day.
        for ii in range(len(list_carrier)):
            if list_carrier[ii][0] <= str(day):
                candidates.append(list_carrier[ii])  # put name into candidates array
            jump = "no"  # triggers an analysis of the candidates array
            if ii != len(list_carrier) - 1:  # if the loop has not reached the end of the list
                if list_carrier[ii][1] == list_carrier[ii + 1][1]:  # if the name current and next name are the same
                    jump = "yes"  # bypasses an analysis of the candidates array
            if jump == "no":  # review the list of candidates
                winner = max(candidates, key=itemgetter(0))  # select the most recent
                if winner[5] == g_station: daily_list.append(winner)  # add the record if it matches the station
                del candidates[:]  # empty out the candidates array.
        for item in daily_list:  # sort carriers in daily list by the list they are in
            if item[2] == "nl":
                dl_nl.append(item)
            if item[2] == "wal":
                dl_wal.append(item)
            if item[2] == "otdl":
                dl_otdl.append(item)
            if item[2] == "aux":
                dl_aux.append(item)
        daily_summary = [] # initialize array for the daily summary
        daily_summary.append(day)
        print("DAY: ", day, "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx")

        print("No List -------------------------------------------------------------------")
        daily_ot = 0.0
        daily_ot_off_route = 0.0
        for name in dl_nl:
            ot = 0.0
            ot_off_route = 0.0
            for r in rings:
                if r[0] == str(day) and r[1] == name[1]:
                    rec = r
            moves_array = []
            if rec != "":
                if rec[2] != "":
                    if rec[4] == "ns day":
                        ot = float(rec[2])
                    else:
                        ot = max(float(rec[2]) - float(8), 0)  # calculate overtime
                if ot <= float(ot_own_rt):
                    ot = 0  # adjust sum for tolerance
                if rec[5] != "":  # if there is a moves in the record
                    move_list = rec[5].split(",")  # convert moves from string to an array
                    sub_array_counter = 0  # sort the moves into multidimentional array
                    i = 1
                    for item in move_list:
                        if (i + 2) % 3 == 0:  # add an array to the array every third item
                            moves_array.append([])
                        moves_array[sub_array_counter].append(item)
                        i += 1
                        if (i - 1) % 3 == 0:
                            sub_array_counter += 1
                            i = 1
                    move_segment_total = 0
                    for move_segment in moves_array:
                        move_segment_total += (
                                float(move_segment[1]) - float(move_segment[0]))  # calc off time off route
                    if rec[4] == "ns day":
                        ot_off_route = float(rec[2])
                    else:
                        ot_off_route = min(move_segment_total, ot)
                    print(name[1], "  ", rec[2], " ", rec[4], " ", moves_array[0][0]," ", moves_array[0][1]," ",
                          moves_array[0][2], " ", ot, " ", move_segment_total," ", ot_off_route)
                    if len(moves_array)>1:
                        for i in range (len(moves_array)-1):
                            print (moves_array[i+1][0]," ", moves_array[i+1][1]," ",moves_array[i+1][2])
                else:
                    print(name[1])
            daily_ot += ot
            daily_ot_off_route += ot_off_route
            rec = ""
        daily_summary.append(daily_ot)
        daily_summary.append(daily_ot_off_route)

        print("Work Assignment -------------------------------------------------------------------")
        daily_ot = 0.0
        daily_ot_off_route = 0.0
        for name in dl_wal:
            ot = 0.0
            ot_off_route = 0.0
            for r in rings:
                if r[0] == str(day) and r[1] == name[1]:
                    rec = r
            moves_array = []
            if rec != "":
                if rec[2] != "":
                    if rec[4] == "ns day":  # calculate overtime
                        ot = float(rec[2])
                    else:
                        ot = max(float(rec[2]) - float(8), 0)  # calculate overtime
                if rec[5] != "":  # if there is a moves in the record
                    move_list = rec[5].split(",")  # convert moves from string to an array
                    sub_array_counter = 0  # sort the moves into multidimentional array
                    i = 1
                    for item in move_list:
                        if (i + 2) % 3 == 0:  # add an array to the array every third item
                            moves_array.append([])
                        moves_array[sub_array_counter].append(item)
                        i += 1
                        if (i - 1) % 3 == 0:
                            sub_array_counter += 1
                            i = 1
                    move_segment_total = 0
                    for move_segment in moves_array:
                        move_segment_total += (
                                float(move_segment[1]) - float(move_segment[0]))  # calc off time off route
                    if rec[4] == "ns day":
                        ot_off_route = float(rec[2])
                    else:
                        ot_off_route = min(move_segment_total, ot)  # calc off time off route
                    if ot_off_route <= float(ot_tol):
                        ot_off_route = 0  # adjust sum for tolerance
                    print(name[1], "  ", rec[2], " ", rec[4], " ", moves_array[0][0], " ", moves_array[0][1], " ",
                          moves_array[0][2], " ", ot, " ", move_segment_total, " ", ot_off_route)
                    if len(moves_array) > 1:
                        for i in range(len(moves_array) - 1):
                            print(moves_array[i + 1][0], " ", moves_array[i + 1][1], " ", moves_array[i + 1][2])
                else:
                    print(name[1])
            daily_ot += ot
            daily_ot_off_route += ot_off_route
            rec = ""
        daily_summary.append(daily_ot)
        daily_summary.append(daily_ot_off_route)
        print("Overtime Desired -------------------------------------------------------------------")
        report.write('Overtime Desired List\n\n')
        report.write ('{:>31}{:<22}{:<14}{:<20}\n'.format("","Moves off Route","Overtime","Availability"))
        report.write('{:<15}{:>8}{:>6}{:<7}{:<7}{:<7}{:>7}{:>7}{:>7}{:>7}\n'
                     .format("name","code","5200","  off","  on","   route","total","off rt","to 10","to 12"))
        report.write("------------------------------------------------------------------------------\n")
        daily_to_10 = 0.0
        daily_to_12 = 0.0
        for name in dl_otdl:
            availability_to_10 = 0.0
            availability_to_12 = 0.0
            ot = 0.0
            ot_off_route = 0.0
            for r in rings: # cycle though clock rings and search for match
                if r[0] == str(day) and r[1] == name[1]: # if there is a match
                    rec = r # capture the record
            moves_array = []
            carrier = name[1][:15]
            if rec != "": # if there is a result for the name
                if rec[4] == "none": # if the code is "none", create empty string
                    code = ""
                else: code = rec[4]
                if code == "no call": # if there is a no call, max out availability
                    availability_to_12 = 12
                    availability_to_10 = 10
                if rec[2] != "": # calculate daily overtime if there is a 5200 time
                    if code == "ns day":
                        ot = float(rec[2])
                    else:
                        ot = max(float(rec[2]) - float(8), 0)  # calculate overtime
                    availability_to_10 = max(10 - float(rec[2]), 0) # calculate availability to 10 hours
                    if availability_to_10 <= float(av_tol): availability_to_10 = 0  # adjust sum for tolerance
                    availability_to_12 = max(12 - float(rec[2]), 0) # calculate availability to 12 hours
                    if availability_to_12 <= float(av_tol): availability_to_12 = 0  # adjust sum for tolerance
                if rec[5] != "":  # if there is a moves in the record
                    move_list = rec[5].split(",")  # convert moves from string to an array
                    sub_array_counter = 0  # sort the moves into multidimentional array
                    i = 1
                    for item in move_list:
                        if (i + 2) % 3 == 0:  # add an array to the array every third item
                            moves_array.append([])
                        moves_array[sub_array_counter].append(item)
                        i += 1
                        if (i - 1) % 3 == 0:
                            sub_array_counter += 1
                            i = 1
                    move_segment_total = 0  # calc off time off route
                    for move_segment in moves_array:
                        move_segment_total += (
                                float(move_segment[1]) - float(move_segment[0]))
                    if code == "ns day":
                        ot_off_route = float(rec[2])
                    else:
                        ot_off_route = min(move_segment_total, ot) # calc off time off route
                    # if there are moves
                    print(name[1], "  ", rec[2], " ", code, " ", moves_array[0][0], " ", moves_array[0][1], " ",
                          moves_array[0][2], " ", ot, " ", move_segment_total, " ", ot_off_route, " ",
                          availability_to_10, " ", availability_to_12)
                    report.write('{:<15}{:>8}{:>6}{:>7}{:>7}{:>7}{:>7}{:>7}{:>7}{:>7}\n'.format
                            (carrier,
                            code,
                            "{0:.2f}".format(float(rec[2])),
                            "{0:.2f}".format(float(moves_array[0][0])),
                            "{0:.2f}".format(float(moves_array[0][1])),
                            moves_array[0][2],
                            "{0:.2f}".format(float(ot)),
                            "{0:.2f}".format(float(ot_off_route)),
                            "{0:.2f}".format(float(availability_to_10)),
                            "{0:.2f}".format(float(availability_to_12))
                            ))
                    if len(moves_array) > 1:
                        for i in range(len(moves_array) - 1):
                            print(moves_array[i + 1][0], " ", moves_array[i + 1][1], " ", moves_array[i + 1][2])
                            report.write('{:>29}{:>7}{:>7}{:>7}\n'.format
                                    ("",
                                    "{0:.2f}".format(float(moves_array[i + 1][0])),
                                    "{0:.2f}".format(float(moves_array[i + 1][1])),
                                    moves_array[i + 1][2],
                                    ))
                else: # if there are no moves
                    print(name[1], "  ", rec[2], " ", code, " ", "", " ", "", " ",
                          "", " ", ot, " ", "", " ", ot_off_route, " ",
                          availability_to_10, " ", availability_to_12)
                    report.write('{:<15}{:>8}{:>6}{:>7}{:>7}{:>7}{:>7}{:>7}{:>7}{:>7}\n'.format
                                 (carrier,
                                  code,
                                  rec[2],
                                  "",
                                  "",
                                  "",
                                  "{0:.2f}".format(float(ot)),
                                  "{0:.2f}".format(float(ot_off_route)),
                                  "{0:.2f}".format(float(availability_to_10)),
                                  "{0:.2f}".format(float(availability_to_12))
                                  ))
            daily_to_10 += availability_to_10
            daily_to_12 += availability_to_12
            rec = ""
        daily_summary.append(daily_to_10)
        daily_summary.append(daily_to_12)
        print("Auxiliary -------------------------------------------------------------------")
        daily_to_10 = 0.0
        daily_to_12 = 0.0
        for name in dl_aux:
            availability_to_10 = 0.0
            availability_to_12 = 0.0
            for r in rings:
                if r[0] == str(day) and r[1] == name[1]:
                    rec = r
            moves_array = []
            if rec != "":
                if rec[5] != "":
                    move_list = rec[5].split(",")  # convert moves from string to an array
                    sub_array_counter = 0  # sort the moves into multidimentional array
                    i = 1
                    for item in move_list:
                        if (i + 2) % 3 == 0:  # add an array to the array every third item
                            moves_array.append([])
                        moves_array[sub_array_counter].append(item)
                        i += 1
                        if (i - 1) % 3 == 0:
                            sub_array_counter += 1
                            i = 1
                if (rec[2])== "": # if the 5200 hours/ rec[2] is an empty string, make it a zero.
                    dailyhours = float(0.0)
                else:
                    dailyhours = float(rec[2])# if the 5200 hours/ rec[2] is an empty string, make it a zero.
                availability_to_10 = max(10 - dailyhours, 0)  # calculate availability to 10 hours
                if availability_to_10 <= float(av_tol): availability_to_10 = 0  # adjust sum for tolerance
                availability_to_12 = max(12 - dailyhours, 0)  # calculate availability to 12 hours
                if availability_to_12 <= float(av_tol): availability_to_12 = 0  # adjust sum for tolerance
                print(name[1], "  ", dailyhours, " ", rec[4], " ", availability_to_10, " ", availability_to_12)
            else:
                print(name[1])
            daily_to_10 += availability_to_10
            daily_to_12 += availability_to_12
            rec = ""
        report.write("------------------------------------------------------------------------------\n")
        daily_summary.append(daily_to_10)
        daily_summary.append(daily_to_12)
        weekly_summary.append(daily_summary)
    report.close() # finish up text document
    if sys.platform == "win32": # open the text document
        os.startfile('kb_sub\\report\\' + filename)
    if sys.platform == "linux":
        subprocess.call(["xdg-open", 'kb_sub/report/' + filename])
    if sys.platform == "darwin":
        subprocess.call(["open", 'kb_sub/report/' + filename])
    print("weekly summary: ")
    for line in weekly_summary:
        print(line)


def rpt_carrier(carrier_list): # Generate and display a report of carrier routes and nsday
    ns_dict = get_custom_nsday() # get the ns day names from the dbase
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S") # create a file name
    filename = "report_carrier_route" + "_" + stamp + ".txt"
    if os.path.isdir('kb_sub/report') == False: # create a directory if it does not exist
        os.makedirs('kb_sub/report')
    try:
        report = open('kb_sub/report/' + filename, "w")
        report.write("Carrier Route and NS Day Report\n\n\n")
        report.write('   Showing results for:\n')
        report.write('      Station: {}\n'.format(g_station))
        if g_range == "day":
            f_date = d_date.strftime("%b %d, %Y")
            report.write('      Date: {}\n'.format(f_date))
        else:
            f_date = g_date[0].strftime("%b %d, %Y")
            end_f_date = g_date[6].strftime("%b %d, %Y")
            report.write('      Dates: {} through {}\n'.format(f_date,end_f_date))
        report.write('      Pay Period: {}\n\n'.format(pay_period))
        report.write('{:>4}  {:<22} {:<17}{:<24}\n'.format("", "Carrier Name", "N/S Day", "Route/s"))
        report.write('      ----------------------------------------------------------------\n')
        aforementioned = []
        i = 1
        for line in carrier_list:
            if line[1] not in aforementioned:
                report.write('{:>4}  {:<22} {:<5}{:<12}{:<24}\n'
                             .format(i, line[1], ns_code[line[3]], ns_dict[line[3]], line[4]))
                if i % 3 == 0:
                    report.write('      ----------------------------------------------------------------\n')
                aforementioned.append(line[1])
                i += 1
        report.close()
        if sys.platform == "win32":
            os.startfile('kb_sub\\report\\' + filename)
        if sys.platform == "linux":
            subprocess.call(["xdg-open", 'kb_sub/report/' + filename])
        if sys.platform == "darwin":
            subprocess.call(["open", 'kb_sub/report/' + filename])
    except:
        messagebox.showerror("Report Generator", "The report was not generated.")

def rpt_carrier_route(carrier_list): # Generate and display a report of carrier routes
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = "report_carrier_route" + "_" + stamp + ".txt"
    if os.path.isdir('kb_sub/report') == False:
        os.makedirs('kb_sub/report')
    try:
        report = open('kb_sub/report/' + filename, "w")
        report.write("Carrier Route Report\n\n\n")
        report.write('   Showing results for:\n')
        report.write('      Station: {}\n'.format(g_station))
        if g_range == "day":
            f_date = d_date.strftime("%b %d, %Y")
            report.write('      Date: {}\n'.format(f_date))
        else:
            f_date = g_date[0].strftime("%b %d, %Y")
            end_f_date = g_date[6].strftime("%b %d, %Y")
            report.write('      Dates: {} through {}\n'.format(f_date,end_f_date))
        report.write('      Pay Period: {}\n\n'.format(pay_period))
        report.write('{:>4}  {:<22} {:<24}\n'.format("", "Carrier Name", "Route/s"))
        report.write('      -----------------------------------------------\n')
        aforementioned = []
        i = 1
        for line in carrier_list:
            if line[1] not in aforementioned:
                report.write('{:>4}  {:<22} {:<24}\n'.format(i, line[1], line[4]))
                if i % 3 == 0:
                    report.write('      -----------------------------------------------\n')
                aforementioned.append(line[1])
                i += 1
        report.close()
        if sys.platform == "win32":
            os.startfile('kb_sub\\report\\' + filename)
        if sys.platform == "linux":
            subprocess.call(["xdg-open", 'kb_sub/report/' + filename])
        if sys.platform == "darwin":
            subprocess.call(["open", 'kb_sub/report/' + filename])
    except:
        messagebox.showerror("Report Generator", "The report was not generated.")

def rpt_carrier_nsday(carrier_list): # Generate and display a report of carrier ns day
    ns_dict = get_custom_nsday() # get the ns day names from the dbase
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = "report_carrier_route" + "_" + stamp + ".txt"
    if os.path.isdir('kb_sub/report') == False:
        os.makedirs('kb_sub/report')
    try:
        report = open('kb_sub/report/' + filename, "w")
        report.write("Carrier Routes NS Day\n\n\n")
        report.write('   Showing results for:\n')
        report.write('      Station: {}\n'.format(g_station))
        if g_range == "day":
            f_date = d_date.strftime("%b %d, %Y")
            report.write('      Date: {}\n'.format(f_date))
        else:
            f_date = g_date[0].strftime("%b %d, %Y")
            end_f_date = g_date[6].strftime("%b %d, %Y")
            report.write('      Dates: {} through {}\n'.format(f_date,end_f_date))
        report.write('      Pay Period: {}\n\n'.format(pay_period))
        report.write('{:>4}  {:<22} {:<17}\n'.format("", "Carrier Name", "N/S Day"))
        report.write('      ----------------------------------------\n')
        aforementioned = []
        i = 1
        for line in carrier_list:
            if line[1] not in aforementioned:
                report.write('{:>4}  {:<22} {:<5}{:<12}\n'
                             .format(i, line[1], ns_code[line[3]], ns_dict[line[3]]))
                if i % 3 == 0:
                    report.write('      ----------------------------------------\n')
                aforementioned.append(line[1])
                i += 1
        report.close()
        if sys.platform == "win32":
            os.startfile('kb_sub\\report\\' + filename)
        if sys.platform == "linux":
            subprocess.call(["xdg-open", 'kb_sub/report/' + filename])
        if sys.platform == "darwin":
            subprocess.call(["open", 'kb_sub/report/' + filename])
    except:
        messagebox.showerror("Report Generator", "The report was not generated.")

def clean_rings3_table():
    sql = "SELECT * FROM rings3 WHERE leave_type IS NULL"
    result = inquire(sql)
    type = ""
    time = float(0.0)
    if result:
        sql = "UPDATE rings3 SET leave_type='%s',leave_time='%s'" \
        "WHERE leave_type IS NULL" \
        % ( type, time)
        commit(sql)
        messagebox.showinfo("Clean Rings",
                            "Rings table has been cleared of NULL values in leave type and leave time columns.")
    else:
        messagebox.showinfo("Clean Rings",
                            "No NULL values in leave type and leave time columns were found in the Rings3 "
                            "table of the database. No action taken.")
    return

def overmax_spreadsheet(carrier_list):
    date = g_date[0]
    dates = []  # array containing days.
    if g_range == "week":
        for i in range(7):
            dates.append(date)
            date += timedelta(days=1)
    sql = "SELECT * FROM rings3 WHERE rings_date BETWEEN '%s' AND '%s' ORDER BY rings_date, carrier_name" \
          % (g_date[0], g_date[6])
    r_rings = inquire(sql)
    # Named styles for workbook
    bd = Side(style='thin', color="80808080")  # defines borders
    ws_header = NamedStyle(name="ws_header", font=Font(bold=True, name='Arial', size=12))
    list_header = NamedStyle(name="list_header", font=Font(bold=True, name='Arial', size=10))
    date_dov = NamedStyle(name="date_dov", font=Font(name='Arial', size=8))
    date_dov_title = NamedStyle(name="date_dov_title", font=Font(bold=True, name='Arial', size=8),
                                alignment=Alignment(horizontal='right'))
    col_header = NamedStyle(name="col_header", font=Font(bold=True, name='Arial', size=8),
                            border=Border(left=bd, right=bd, top=bd, bottom=bd),
                            alignment=Alignment(horizontal='left'))
    col_center_header = NamedStyle(name="col_center_header", font=Font(bold=True, name='Arial', size=8),
                            alignment=Alignment(horizontal='center'),
                           border=Border(left=bd, right=bd, top=bd, bottom=bd))
    vert_header = NamedStyle(name="vert_header", font=Font(bold=True, name='Arial', size=8),
                             border=Border(left=bd, right=bd, top=bd, bottom=bd),
                            alignment=Alignment(horizontal='right',text_rotation=90))
    input_name = NamedStyle(name="input_name", font=Font(name='Arial', size=8),
                            border=Border(left=bd, right=bd, top=bd, bottom=bd))
    input_s = NamedStyle(name="input_s", font=Font(name='Arial', size=8),
                         border=Border(left=bd, right=bd, top=bd, bottom=bd),
                         alignment=Alignment(horizontal='right'))
    calcs = NamedStyle(name="calcs", font=Font(name='Arial', size=8),
                       border=Border(left=bd, right=bd, top=bd, bottom=bd),
                       fill=PatternFill(fgColor='e5e4e2', fill_type='solid'),
                       alignment=Alignment(horizontal='right'))
    vert_calcs = NamedStyle(name="vert_calcs", font=Font(name='Arial', size=8),
                       border=Border(left=bd, right=bd, top=bd, bottom=bd),
                       fill=PatternFill(fgColor='e5e4e2', fill_type='solid'),
                       alignment=Alignment(horizontal='right',text_rotation=90))
    instruct_text = NamedStyle(name="instruct_text", font=Font(name='Arial', size=8),
                    alignment=Alignment(horizontal='left',vertical ='top'))
    wb = Workbook()  # define the workbook
    overmax = wb.active  # create first worksheet
    summary = wb.create_sheet("summary")
    instructions = wb.create_sheet("instructions")
    for x in range(2,8): overmax.row_dimensions[x].height = 10# adjust all row height
    overmax.title = "12_60_violations"  # title first worksheet
    overmax.oddFooter.center.text = "&A"
    sheets = (overmax, instructions)
    for sheet in sheets:
        sheet.column_dimensions["A"].width = 13
        sheet.column_dimensions["B"].width = 3
        sheet.column_dimensions["C"].width = 5
        sheet.column_dimensions["D"].width = 4
        sheet.column_dimensions["E"].width = 2
        sheet.column_dimensions["F"].width = 4
        sheet.column_dimensions["G"].width = 2
        sheet.column_dimensions["H"].width = 4
        sheet.column_dimensions["I"].width = 2
        sheet.column_dimensions["J"].width = 4
        sheet.column_dimensions["K"].width = 2
        sheet.column_dimensions["L"].width = 4
        sheet.column_dimensions["M"].width = 2
        sheet.column_dimensions["N"].width = 4
        sheet.column_dimensions["O"].width = 2
        sheet.column_dimensions["P"].width = 4
        sheet.column_dimensions["Q"].width = 2
        sheet.column_dimensions["R"].width = 4
        sheet.column_dimensions['R'].hidden = True
        sheet.column_dimensions["S"].width = 5
        sheet.column_dimensions["T"].width = 5
        sheet.column_dimensions["U"].width = 2
        sheet.column_dimensions["V"].width = 2
        sheet.column_dimensions["W"].width = 2
        sheet.column_dimensions["X"].width = 5
    # summary worksheet - format cells
    summary.oddFooter.center.text = "&A"
    summary.merge_cells('A1:R1')
    summary['A1'] = "12 and 60 Hour Violations Summary"
    summary['A1'].style = ws_header
    summary.column_dimensions["A"].width = 15
    summary.column_dimensions["B"].width = 8
    summary['A3'] = "Date:"
    summary['A3'].style = date_dov_title
    summary.merge_cells('B3:D3')  # blank field for date
    summary['B3'] = dates[0].strftime("%x") + " - " + dates[6].strftime("%x")
    summary['B3'].style = date_dov
    summary.merge_cells('K3:N3')
    summary['F3'] = "Pay Period:" # Pay Period Header
    summary['F3'].style = date_dov_title
    summary.merge_cells('G3:I3')  # blank field for pay period
    summary['G3'] = pay_period
    summary['G3'].style = date_dov
    summary['A4'] = "Station:" # Station Header
    summary['A4'].style = date_dov_title
    summary.merge_cells('B4:D4')  # blank field for station
    summary['B4'] = g_station
    summary['B4'].style = date_dov
    summary['A6'] = "name"
    summary['A6'].style = col_center_header
    summary['B6'] = "violation"
    summary['B6'].style = col_center_header
    # overmax worksheet - format cells
    overmax.merge_cells('A1:R1')
    overmax['A1'] = "12 and 60 Hour Violations Worksheet"
    overmax['A1'].style = ws_header
    overmax['A3'] = "Date:"
    overmax['A3'].style = date_dov_title
    overmax.merge_cells('B3:J3')# blank field for date
    overmax['B3'] = dates[0].strftime("%x") + " - " + dates[6].strftime("%x")
    overmax['B3'].style = date_dov
    overmax.merge_cells('K3:N3')
    overmax['K3'] = "Pay Period:"
    overmax['k3'].style = date_dov_title
    overmax.merge_cells('O3:S3') # blank field for pay period
    overmax['O3'] =  pay_period
    overmax['O3'].style = date_dov
    overmax['A4'] = "Station:"
    overmax['A4'].style = date_dov_title
    overmax.merge_cells('B4:J4')# blank field for station
    overmax['B4'] = g_station
    overmax['B4'].style = date_dov

    overmax.merge_cells('D6:Q6')
    overmax['D6'] = "Daily Paid Leave times with type"
    overmax['D6'].style = col_center_header
    overmax.merge_cells('D7:Q7')
    overmax['D7'] = "Daily 5200 times"
    overmax['D7'].style = col_center_header
    overmax['A8'] = "name"
    overmax['A8'].style = col_header
    overmax['B8'] = "list"
    overmax['B8'].style = col_header
    overmax.merge_cells('C5:C8')
    overmax['C5'] = "Weekly\n5200"
    overmax['C5'].style = vert_header
    overmax.merge_cells('D8:E8')
    overmax['D8'] = "sat"
    overmax['D8'].style = col_center_header
    overmax.merge_cells('F8:G8')
    overmax['F8'] = "sun"
    overmax['F8'].style = col_center_header
    overmax.merge_cells('H8:I8')
    overmax['H8'] = "mon"
    overmax['H8'].style = col_center_header
    overmax.merge_cells('J8:K8')
    overmax['J8'] = "tue"
    overmax['J8'].style = col_center_header
    overmax.merge_cells('L8:M8')
    overmax['L8'] = "wed"
    overmax['L8'].style = col_center_header
    overmax.merge_cells('N8:O8')
    overmax['N8'] = "thr"
    overmax['N8'].style = col_center_header
    overmax.merge_cells('P8:Q8')
    overmax['P8'] = "fri"
    overmax['P8'].style = col_center_header
    overmax.merge_cells('S4:S8')
    overmax['S4'] = " Weekly\nViolation"
    overmax['S4'].style = vert_header
    overmax.merge_cells('T4:T8')
    overmax['T4'] = "Daily\nViolation"
    overmax['T4'].style = vert_header
    overmax.merge_cells('U4:U8')
    overmax['U4'] = "Wed Adj"
    overmax['U4'].style = vert_header
    overmax.merge_cells('V4:V8')
    overmax['V4'] = "Thr Adj"
    overmax['V4'].style = vert_header
    overmax.merge_cells('W4:W8')
    overmax['W4'] = "Fri Adj"
    overmax['W4'].style = vert_header
    overmax.merge_cells('X4:X8')
    overmax['X4'] = "Total\nViolation"
    overmax['X4'].style = vert_header

    # format the instructions cells
    instructions.merge_cells('A1:R1')
    instructions['A1'] = "12 and 60 Hour Violations Instructions"
    instructions['A1'].style = ws_header
    instructions.row_dimensions[3].height = 165
    instructions['A3'].style = instruct_text
    instructions.merge_cells('A3:X3')
    instructions['A3']="Instructions: \n1. Fill in the name \n" \
    "2. Fill in the list. Enter either “otdl”,”wal”,”nl” or “aux” in list columns. Use only lowercase. \n" \
    "   If you do not enter anything, the default is “otdl\n" \
    "\totdl = overtime desired list\n" \
    "\twal = work assignment list\n"  \
    "\tnl = no list \n"  \
    "\taux = auxiliary (this would be a cca or city carrier assistant).\n" \
    "3. Fill in the weekly 5200 time in field C if it exceeds 60 hours or if the sum of all daily non 5200 times (all fields D) plus \n" \
    "   the weekly 5200 time (field C) will  exceed 60 hours.\n" \
    "4. Fill in any daily non 5200 times and types in fields D and E. Enter only paid leave types such as sick leave, annual\n" \
    "   leave and holiday leave. Do not enter unpaid leave types such as LWOP (leave without pay) or AWOL (absent \n" \
    "   without leave).\n" \
    "5. Fill in any daily 5200 times which exceed 12 hours for otdl carriers or 11.50 hours for any other carrier in fields F.\n" \
    "   Failing to fill out the daily values for Wednesday, Thursday and Friday could cause errors in calculating the adjustments,\n" \
    "   so fill those in.\n" \
    "6. The gray fields will fill automatically. Do not enter an information in these fields as it will delete the formulas.\n" \
    "7. Field O will show the violation in hours which you should seek a remedy for. \n"
    for x in range(4,20): instructions.row_dimensions[x].height = 10  # adjust all row height
    instructions.merge_cells('D6:Q6')
    instructions['D6'] = "Daily Paid Leave times with type"
    instructions['D6'].style = col_center_header
    instructions.merge_cells('D7:Q7')
    instructions['D7'] = "Daily 5200 times"
    instructions['D7'].style = col_center_header
    instructions['A8'] = "name"
    instructions['A8'].style = col_header
    instructions['B8'] = "list"
    instructions['B8'].style = col_header
    instructions.merge_cells('C5:C8')
    instructions['C5'] = "Weekly\n5200"
    instructions['C5'].style = vert_header
    instructions.merge_cells('D8:E8')
    instructions['D8'] = "sat"
    instructions['D8'].style = col_center_header
    instructions.merge_cells('F8:G8')
    instructions['F8'] = "sun"
    instructions['F8'].style = col_center_header
    instructions.merge_cells('H8:I8')
    instructions['H8'] = "mon"
    instructions['H8'].style = col_center_header
    instructions.merge_cells('J8:K8')
    instructions['J8'] = "tue"
    instructions['J8'].style = col_center_header
    instructions.merge_cells('L8:M8')
    instructions['L8'] = "wed"
    instructions['L8'].style = col_center_header
    instructions.merge_cells('N8:O8')
    instructions['N8'] = "thr"
    instructions['N8'].style = col_center_header
    instructions.merge_cells('P8:Q8')
    instructions['P8'] = "fri"
    instructions['P8'].style = col_center_header
    instructions.merge_cells('S4:S8')
    instructions['S4'] = " Weekly\nViolation"
    instructions['S4'].style = vert_header
    instructions.merge_cells('T4:T8')
    instructions['T4'] = "Daily\nViolation"
    instructions['T4'].style = vert_header
    instructions.merge_cells('U4:U8')
    instructions['U4'] = "Wed Adj"
    instructions['U4'].style = vert_header
    instructions.merge_cells('V4:V8')
    instructions['V4'] = "Thr Adj"
    instructions['V4'].style = vert_header
    instructions.merge_cells('W4:W8')
    instructions['W4'] = "Fri Adj"
    instructions['W4'].style = vert_header
    instructions.merge_cells('X4:X8')
    instructions['X4'] = "Total\nViolation"
    instructions['X4'].style = vert_header
    instructions['A9'] = "A"
    instructions['A9'].style = col_center_header
    instructions['B9'] = "B"
    instructions['B9'].style = col_center_header
    instructions['C9'] = "C"
    instructions['C9'].style = col_center_header
    instructions['D9'] = "D"
    instructions['D9'].style = col_center_header
    instructions['E9'] = "E"
    instructions['E9'].style = col_center_header
    instructions['F9'] = "G"
    instructions['F9'].style = col_center_header
    instructions.merge_cells('F9:G9')
    instructions['H9'] = "D"
    instructions['H9'].style = col_center_header
    instructions['I9'] = "E"
    instructions['I9'].style = col_center_header
    instructions['J9'] = "D"
    instructions['J9'].style = col_center_header
    instructions['K9'] = "E"
    instructions['K9'].style = col_center_header
    instructions['L9'] = "D"
    instructions['L9'].style = col_center_header
    instructions['M9'] = "E"
    instructions['M9'].style = col_center_header
    instructions['N9'] = "D"
    instructions['N9'].style = col_center_header
    instructions['O9'] = "E"
    instructions['O9'].style = col_center_header
    instructions['P9'] = "D"
    instructions['P9'].style = col_center_header
    instructions['Q9'] = "E"
    instructions['Q9'].style = col_center_header
    instructions['S9'] = "J"
    instructions['S9'].style = col_center_header
    instructions['T9'] = "K"
    instructions['T9'].style = col_center_header
    instructions['U9'] = "L"
    instructions['U9'].style = col_center_header
    instructions['V9'] = "M"
    instructions['V9'].style = col_center_header
    instructions['W9'] = "N"
    instructions['W9'].style = col_center_header
    instructions['X9'] = "O"
    instructions['X9'].style = col_center_header
    i = 10
    # instructions name
    instructions.merge_cells('A' + str(i) + ':A' + str(i + 1))  # merge box for name
    instructions['A10'] = "kubrick, s"
    instructions['A10'].style = input_name
    # instructions list
    instructions.merge_cells('B' + str(i) + ':B' + str(i + 1))  # merge box for list type input
    instructions['B10'] = "wal"
    instructions['B10'].style = input_s
    # instructions weekly
    instructions.merge_cells('C' + str(i) + ':C' + str(i + 1))  # merge box for weekly input
    instructions['C10'] = 75.00
    instructions['C10'].style = input_s
    instructions['C10'].number_format = "#,###.00;[RED]-#,###.00"
    # instructions saturday
    instructions.merge_cells('D' + str(i + 1) + ':E' + str(i + 1))  # merge box for sat 5200
    instructions['D' + str(i)] = ""  # leave time
    instructions['D' + str(i)].style = input_s
    instructions['D' + str(i)].number_format = "#,###.00;[RED]-#,###.00"
    instructions['E' + str(i)] = ""  # leave type
    instructions['E' + str(i)].style = input_s
    instructions['D' + str(i + 1)] = 13.00  # 5200 time
    instructions['D' + str(i + 1)].style = input_s
    instructions['D' + str(i + 1)].number_format = "#,###.00;[RED]-#,###.00"
    # instructions sunday
    instructions.merge_cells('F' + str(i + 1) + ':G' + str(i + 1))  # merge box for sun 5200
    instructions['F' + str(i)] = ""  # leave time
    instructions['F' + str(i)].style = input_s
    instructions['F' + str(i)].number_format = "#,###.00;[RED]-#,###.00"
    instructions['G' + str(i)] = ""  # leave type
    instructions['G' + str(i)].style = input_s
    instructions['F' + str(i + 1)] = ""  # 5200 time
    instructions['F' + str(i + 1)].style = input_s
    instructions['F' + str(i + 1)].number_format = "#,###.00;[RED]-#,###.00"
    # instructions monday
    instructions.merge_cells('H' + str(i + 1) + ':I' + str(i + 1))  # merge box for mon 5200
    instructions['H' + str(i)] = 8  # leave time
    instructions['H' + str(i)].style = input_s
    instructions['H' + str(i)].number_format = "#,###.00;[RED]-#,###.00"
    instructions['I' + str(i)] = "h"  # leave type
    instructions['I' + str(i)].style = input_s
    instructions['H' + str(i + 1)] = ""  # 5200 time
    instructions['H' + str(i + 1)].style = input_s
    instructions['H' + str(i + 1)].number_format = "#,###.00;[RED]-#,###.00"
    # instructions tuesday
    instructions.merge_cells('J' + str(i + 1) + ':K' + str(i + 1))  # merge box for tue 5200
    instructions['J' + str(i)] = ""  # leave time
    instructions['J' + str(i)].style = input_s
    instructions['J' + str(i)].number_format = "#,###.00;[RED]-#,###.00"
    instructions['K' + str(i)] = ""  # leave type
    instructions['K' + str(i)].style = input_s
    instructions['J' + str(i + 1)] = 14  # 5200 time
    instructions['J' + str(i + 1)].style = input_s
    instructions['J' + str(i + 1)].number_format = "#,###.00;[RED]-#,###.00"
    # instructions wednesday
    instructions.merge_cells('L' + str(i + 1) + ':M' + str(i + 1))  # merge box for wed 5200
    instructions['L' + str(i)] = ""  # leave time
    instructions['L' + str(i)].style = input_s
    instructions['L' + str(i)].number_format = "#,###.00;[RED]-#,###.00"
    instructions['M' + str(i)] = ""  # leave type
    instructions['M' + str(i)].style = input_s
    instructions['L' + str(i + 1)] = 14  # 5200 time
    instructions['L' + str(i + 1)].style = input_s
    instructions['M' + str(i + 1)].number_format = "#,###.00;[RED]-#,###.00"
    # instructions thursday
    instructions.merge_cells('N' + str(i + 1) + ':O' + str(i + 1))  # merge box for thr 5200
    instructions['N' + str(i)] = ""  # leave time
    instructions['N' + str(i)].style = input_s
    instructions['N' + str(i)].number_format = "#,###.00;[RED]-#,###.00"
    instructions['O' + str(i)] = ""  # leave type
    instructions['O' + str(i)].style = input_s
    instructions['N' + str(i + 1)] = 13  # 5200 time
    instructions['N' + str(i + 1)].style = input_s
    instructions['N' + str(i + 1)].number_format = "#,###.00;[RED]-#,###.00"
    # instructions friday
    instructions.merge_cells('P' + str(i + 1) + ':Q' + str(i + 1))  # merge box for fri 5200
    instructions['P' + str(i)] = ""  # leave time
    instructions['P' + str(i)].style = input_s
    instructions['P' + str(i)].number_format = "#,###.00;[RED]-#,###.00"
    instructions['Q' + str(i)] = ""  # leave type
    instructions['Q' + str(i)].style = input_s
    instructions['P' + str(i + 1)] = 13  # 5200 time
    instructions['P' + str(i + 1)].style = input_s
    instructions['P' + str(i + 1)].number_format = "#,###.00;[RED]-#,###.00"
    # instructions hidden columns
    page = "instructions"

    formula = "=SUM(%s!D%s:%s!P%s)+%s!D%s + %s!H%s + %s!J%s + %s!L%s + " \
                "%s!N%s + %s!P%s" % (page, str(i + 1), page, str(i + 1),
                                     page, str(i), page, str(i), page, str(i),
                                     page, str(i), page, str(i), page, str(i))
    instructions['R' + str(i)] = formula
    instructions['R' + str(i)].style = calcs
    instructions['R' + str(i)].number_format = "#,###.00;[RED]-#,###.00"
    formula = "=SUM(%s!C%s+%s!D%s+%s!H%s+%s!J%s+%s!L%s+%s!N%s+%s!P%s)" % \
                (page, str(i), page, str(i), page, str(i),
                 page, str(i), page, str(i), page, str(i),
                 page, str(i))
    instructions['R' + str(i + 1)] = formula
    instructions['R' + str(i + 1)].style = calcs
    instructions['R' + str(i + 1)].number_format = "#,###.00;[RED]-#,###.00"
    # instructions weekly violations
    instructions.merge_cells('S' + str(i) + ':S' + str(i + 1))  # merge box for weekly violation
    formula = "=MAX(IF(%s!R%s>%s!R%s,MAX(%s!R%s-60,0),MAX(%s!R%s-60)),0)" % (page, str(i),
                                                                               page, str(i + 1),
                                                                               page, str(i),
                                                                               page, str(i + 1),)
    instructions['S10'] = formula
    instructions['S10'].style = calcs
    instructions['S10'].number_format = "#,###.00;[RED]-#,###.00"
    # instructions daily violations
    formula_d = "=IF(OR(%s!B%s=\"wal\",%s!B%s=\"nl\",%s!B%s=\"aux\")," \
                "(SUM(IF(%s!D%s>11.5,%s!D%s-11.5,0)+IF(%s!H%s>11.5,%s!H%s-11.5,0)+IF(%s!J%s>11.5,%s!J%s-11.5,0)" \
                "+IF(%s!L%s>11.5,%s!L%s-11.5,0)+IF(%s!N%s>11.5,%s!N%s-11.5,0)+IF(%s!P%s>11.5,%s!P%s-11.5,0)))," \
                "(SUM(IF(%s!D%s>12,%s!D%s-12,0)+IF(%s!H%s>12,%s!H%s-12,0)+IF(%s!J%s>12,%s!J%s-12,0)" \
                "+IF(%s!L%s>12,%s!L%s-12,0)+IF(%s!N%s>12,%s!N%s-12,0)+IF(%s!P%s>12,%s!P%s-12,0))))" \
                % (page, str(i), page, str(i), page, str(i),
                   page, str(i + 1), page, str(i + 1), page, str(i + 1),
                   page, str(i + 1), page, str(i + 1), page, str(i + 1),
                   page, str(i + 1), page, str(i + 1), page, str(i + 1),
                   page, str(i + 1), page, str(i + 1), page, str(i + 1),
                   page, str(i + 1), page, str(i + 1), page, str(i + 1),
                   page, str(i + 1), page, str(i + 1), page, str(i + 1),
                   page, str(i + 1), page, str(i + 1), page, str(i + 1),
                   page, str(i + 1), page, str(i + 1), page, str(i + 1))
    instructions['T' + str(i)] = formula_d
    instructions.merge_cells('T' + str(i) + ':T' + str(i + 1))  # merge box for daily violation
    instructions['T' + str(i)].style = calcs
    instructions['T' + str(i)].number_format = "#,###.00"
    # instructions wed adjustment
    instructions.merge_cells('U' + str(i) + ':U' + str(i + 1))  # merge box for wed adj
    formula_e = "=IF(OR(%s!B%s=\"wal\",%s!B%s=\"nl\",%s!B%s=\"aux\")," \
                "IF(AND(%s!S%s-(%s!N%s+%s!N%s+%s!P%s+%s!P%s)>0,%s!L%s>11.5)," \
                "IF(%s!S%s-(%s!N%s+%s!N%s+%s!P%s+%s!P%s)>%s!L%s-11.5,%s!L%s-11.5,%s!S%s-(%s!N%s+%s!N%s+%s!P%s+%s!P%s)),0)," \
                "IF(AND(%s!S%s-(%s!N%s+%s!N%s+%s!P%s+%s!P%s)>0,%s!L%s>12)," \
                "IF(%s!S%s-(%s!N%s+%s!N%s+%s!P%s+%s!P%s)>%s!L%s-12,%s!L%s-12,%s!S%s-(%s!N%s+%s!N%s+%s!P%s+%s!P%s)),0))" \
                % (page, str(i), page, str(i), page, str(i),
                   page, str(i), page, str(i + 1), page, str(i),
                   page, str(i + 1), page, str(i), page, str(i + 1),
                   page, str(i), page, str(i + 1), page, str(i),
                   page, str(i + 1), page, str(i), page, str(i + 1),
                   page, str(i + 1), page, str(i), page, str(i + 1),
                   page, str(i), page, str(i + 1), page, str(i),
                   page, str(i), page, str(i + 1), page, str(i),
                   page, str(i + 1), page, str(i), page, str(i + 1),
                   page, str(i), page, str(i + 1), page, str(i),
                   page, str(i + 1), page, str(i), page, str(i + 1),
                   page, str(i + 1), page, str(i), page, str(i + 1),
                   page, str(i), page, str(i + 1), page, str(i))
    instructions['U' + str(i)] = formula_e
    instructions['U' + str(i)].style = vert_calcs
    instructions['U' + str(i)].number_format = "#,###.00"
    # instructions thr adjustment
    formula_f = "=IF(OR(%s!B%s=\"wal\",%s!B%s=\"nl\",%s!B%s=\"aux\")," \
                "IF(AND(%s!S%s-(%s!P%s+%s!P%s)>0,%s!N%s>11.5)," \
                "IF(%s!S%s-(%s!P%s+%s!P%s)>%s!N%s-11.5,%s!N%s-11.5,%s!S%s-(%s!P%s+%s!P%s)),0)," \
                "IF(AND(%s!S%s-(%s!P%s+%s!P%s)>0,%s!N%s>12)," \
                "IF(%s!S%s-(%s!P%s+%s!P%s)>%s!N%s-12,%s!N%s-12,%s!S%s-(%s!P%s+%s!P%s)),0))" \
                % (page, str(i), page, str(i), page, str(i),
                   page, str(i), page, str(i + 1), page, str(i),
                   page, str(i + 1),
                   page, str(i), page, str(i + 1), page, str(i),
                   page, str(i + 1), page, str(i + 1), page, str(i),
                   page, str(i + 1), page, str(i),
                   page, str(i), page, str(i + 1), page, str(i),
                   page, str(i + 1),
                   page, str(i), page, str(i + 1), page, str(i),
                   page, str(i + 1), page, str(i + 1), page, str(i),
                   page, str(i + 1), page, str(i)
                   )
    instructions.merge_cells('V' + str(i) + ':V' + str(i + 1))  # merge box for thr adj
    instructions['V' + str(i)] = formula_f
    instructions['V' + str(i)].style = vert_calcs
    instructions['V' + str(i)].number_format = "#,###.00"
    # instructions fri adjustment
    instructions.merge_cells('W' + str(i) + ':W' + str(i + 1))  # merge box for fri adj
    formula_g = "=IF(OR(%s!B%s=\"wal\",%s!B%s=\"nl\",%s!B%s=\"aux\")," \
                "IF(AND(%s!S%s>0,%s!P%s>11.5)," \
                "IF(%s!S%s>%s!P%s-11.5,%s!P%s-11.5,%s!S%s),0)," \
                "IF(AND(%s!S%s>0,%s!P%s>12)," \
                "IF(%s!S%s>%s!P%s-12,%s!P%s-12,%s!S%s),0))" \
                % (page, str(i), page, str(i), page, str(i),
                   page, str(i), page, str(i + 1), page, str(i),
                   page, str(i + 1), page, str(i + 1), page, str(i),
                   page, str(i), page, str(i + 1), page, str(i),
                   page, str(i + 1), page, str(i + 1), page, str(i))
    instructions['W' + str(i)] = formula_g
    instructions['W' + str(i)].style = vert_calcs
    instructions['W' + str(i)].number_format = "#,###.00"
    # instructions total violation
    instructions.merge_cells('X' + str(i) + ':X' + str(i + 1))  # merge box for total violation
    formula_h = "=SUM(%s!S%s:%s!T%s)-(%s!U%s+%s!V%s+%s!W%s)" \
                % (page, str(i), page, str(i), page, str(i),
                   page, str(i), page, str(i))
    instructions['X' + str(i)] = formula_h
    instructions['X' + str(i)].style = calcs
    instructions['X' + str(i)].number_format = "#,###.00"
    instructions['D12'] = "F"
    instructions['D12'].style = col_center_header
    instructions.merge_cells('D12:E12')
    instructions['F12'] = "F"
    instructions['F12'].style = col_center_header
    instructions.merge_cells('F12:G12')
    instructions['H12'] = "F"
    instructions['H12'].style = col_center_header
    instructions.merge_cells('H12:I12')
    instructions['J12'] = "F"
    instructions['J12'].style = col_center_header
    instructions.merge_cells('J12:K12')
    instructions['L12'] = "F"
    instructions['L12'].style = col_center_header
    instructions.merge_cells('L12:M12')
    instructions['N12'] = "F"
    instructions['N12'].style = col_center_header
    instructions.merge_cells('N12:O12')
    instructions['P12'] = "F"
    instructions['P12'].style = col_center_header
    instructions.merge_cells('P12:Q12')
    # legend section
    instructions.row_dimensions[14].height = 180
    instructions['A14'].style = instruct_text
    instructions.merge_cells('A14:X14')
    instructions['A14'] = "Legend: \n" \
        "A.  Name \n" \
        "B.  List: Either otdl, wal, nl or aux (always use lowercase to preserve operation of the formulas).\n" \
        "C.  Weekly 5200 Time: Enter the 5200 time for the week. \n" \
        "D.  Daily Non 5200 Time: Enter daily hours for either holiday, annual sick leave or other type of paid leave.\n" \
        "E.  Daily Non 5200 Type: Enter “a” for annual, “s” for sick, “h” for holiday, etc. \n" \
        "F.  Daily 5200 Hours: Enter 5200 hours or hours worked for the day. \n" \
        "G.  No value allowed: No non 5200 times allowed for Sundays.\n" \
        "J.   Weekly Violations: This is the total of violations over 60 hours in a week.\n" \
        "K.  Daily Violations: This is the total of daily violations which have exceeded 11.50 (for wal, nl or aux)\n" \
        "     or 12 hours in a day (for otdl).\n" \
        "L.  Wednesday Adjustment: In cases were the 60 hour limit is reached and a daily violation happens (on Wednesday),\n" \
        "     this column deducts one of the violations so to provide a correct remedy.\n" \
        "M.  Thursday Adjustment: In cases were the 60 hour limit is reached and a daily violation happens (on Thursday), \n" \
        "     this column deducts one of the violations so to provide a correct remedy.\n" \
        "N.  Friday Adjustment: In cases were the 60 hour limit is reached and a daily violation happens (on Friday),\n" \
        "     this column deducts one of the violations so to provide a correct remedy.\n" \
        "O.  Total Violation: This field is the end result of the calculation. This is the addition of the total daily  " \
                          "violations and the\n" \
        "     weekly violation, it shows the sum of the two. This is the value which the steward should seek a remedy for."
    daily_list = []  # array
    candidates = []
    for day in dates:
        del daily_list[:]
        # create a list of carriers for each day.
        for ii in range(len(carrier_list)):
            if carrier_list[ii][0] <= str(day):
                candidates.append(carrier_list[ii])  # put name into candidates array
            jump = "no"  # triggers an analysis of the candidates array
            if ii != len(carrier_list) - 1:  # if the loop has not reached the end of the list
                if carrier_list[ii][1] == carrier_list[ii + 1][1]:  # if the name current and next name are the same
                    jump = "yes"  # bypasses an analysis of the candidates array
            if jump == "no":  # review the list of candidates
                winner = max(candidates, key=itemgetter(0))  # select the most recent
                if winner[5] == g_station: daily_list.append(winner)  # add the record if it matches the station
                del candidates[:]  # empty out the candidates array.
    summary_i = 7
    i = 9
    for line in carrier_list:
        # if there is a ring to match the carrier/ date then printe
        carrier_rings = []
        total = 0.0
        grandtotal = 0.0
        totals_array = ["", "", "", "", "", "", ""]
        leavetype_array = ["", "", "", "", "", "", ""]
        leavetime_array = ["", "", "", "", "", "", ""]
        c = 0
        daily_violation = False
        for day in dates:
            for each in r_rings:
                if each[0] == str(day) and each[1] == line[1]:  # find if there are rings for the carrier
                    carrier_rings.append(each) # add any rings to an array
                    if isfloat(each[2]):
                        totals_array[c] = float(each[2])
                        if float(each[2])> 12 and line[2] == "otdl":
                            daily_violation = True
                        if float(each[2]) > 11.5 and line[2] != "otdl":
                            daily_violation = True
                    else:
                        totals_array[c] = each[2]
                    if each[6]=="annual":
                        leavetype_array[c] = "A"
                    if each[6]=="sick":
                        leavetype_array[c] = "S"
                    if each[6] == "holiday":
                        leavetype_array[c] = "H"
                    if each[6] == "other":
                        leavetype_array[c] = "O"
                    if each[7] == "0.0" or each[7]=="0":
                        leavetime_array[c] = ""
                    elif isfloat(each[7]):
                            leavetime_array[c] = float(each[7])
                    else:
                        leavetime_array[c] = each[7]
            c += 1
        for item in carrier_rings:
            if item[2] == "": # convert empty 5200 strings to zero
                t = 0.0
            else: t = float(item[2])
            if item[7] == "": # convert leave time strings to zero
                l = 0.0
            else: l = float(item[7])
            total = total + t
            grandtotal = grandtotal + t + l
        if grandtotal > 60 or daily_violation == True:
            # output to the gui
            overmax.row_dimensions[i].height = 10# adjust all row height
            overmax.row_dimensions[i+1].height = 10
            overmax.merge_cells('A'+ str(i)+':A' + str(i+1))
            overmax['A' + str(i)] = line[1]  # name
            overmax['A' + str(i)].style = input_name
            overmax.merge_cells('B' + str(i) + ':B' + str(i+1)) # merge box for list
            overmax['B' + str(i)] = line[2]  # list
            overmax['B' + str(i)].style = input_s
            overmax.merge_cells('C' + str(i) + ':C' + str(i+1)) # merge box for weekly 5200
            overmax['C' + str(i)] = float(total)  # total
            overmax['C' + str(i)].style = input_s
            overmax['C' + str(i)].number_format = "#,###.00;[RED]-#,###.00"
            #saturday
            overmax.merge_cells('D' + str(i +1 ) + ':E' + str(i + 1))  # merge box for sat 5200
            overmax['D' + str(i)] = leavetime_array[0] # leave time
            overmax['D' + str(i)].style = input_s
            overmax['D' + str(i)].number_format = "#,###.00;[RED]-#,###.00"
            overmax['E' + str(i)] = leavetype_array[0] # leave type
            overmax['E' + str(i)].style = input_s
            overmax['D' + str(i + 1)] = totals_array[0] # 5200 time
            overmax['D' + str(i + 1)].style = input_s
            overmax['D' + str(i + 1)].number_format = "#,###.00;[RED]-#,###.00"
            # sunday
            overmax.merge_cells('F' + str(i + 1) + ':G' + str(i + 1))  # merge box for sun 5200
            overmax['F' + str(i)] = leavetime_array[1]  # leave time
            overmax['F' + str(i)].style = input_s
            overmax['F' + str(i)].number_format = "#,###.00;[RED]-#,###.00"
            overmax['G' + str(i)] = leavetype_array[1]  # leave type
            overmax['G' + str(i)].style = input_s
            overmax['F' + str(i + 1)] = totals_array[1]  # 5200 time
            overmax['F' + str(i + 1)].style = input_s
            overmax['F' + str(i + 1)].number_format = "#,###.00;[RED]-#,###.00"
            # monday
            overmax.merge_cells('H' + str(i + 1) + ':I' + str(i + 1))  # merge box for mon 5200
            overmax['H' + str(i)] = leavetime_array[2]  # leave time
            overmax['H' + str(i)].style = input_s
            overmax['H' + str(i)].number_format = "#,###.00;[RED]-#,###.00"
            overmax['I' + str(i)] = leavetype_array[2]  # leave type
            overmax['I' + str(i)].style = input_s
            overmax['H' + str(i + 1)] = totals_array[2]  # 5200 time
            overmax['H' + str(i + 1)].style = input_s
            overmax['H' + str(i + 1)].number_format = "#,###.00;[RED]-#,###.00"
            # tuesday
            overmax.merge_cells('J' + str(i + 1) + ':K' + str(i + 1))  # merge box for tue 5200
            overmax['J' + str(i)] = leavetime_array[3]  # leave time
            overmax['J' + str(i)].style = input_s
            overmax['J' + str(i)].number_format = "#,###.00;[RED]-#,###.00"
            overmax['K' + str(i)] = leavetype_array[3]  # leave type
            overmax['K' + str(i)].style = input_s
            overmax['J' + str(i + 1)] = totals_array[3]  # 5200 time
            overmax['J' + str(i + 1)].style = input_s
            overmax['J' + str(i + 1)].number_format = "#,###.00;[RED]-#,###.00"
            # wednesday
            overmax.merge_cells('L' + str(i + 1) + ':M' + str(i + 1))  # merge box for wed 5200
            overmax['L' + str(i)] = leavetime_array[4]  # leave time
            overmax['L' + str(i)].style = input_s
            overmax['L' + str(i)].number_format = "#,###.00;[RED]-#,###.00"
            overmax['M' + str(i)] = leavetype_array[4]  # leave type
            overmax['M' + str(i)].style = input_s
            overmax['L' + str(i + 1)] = totals_array[4]  # 5200 time
            overmax['L' + str(i + 1)].style = input_s
            overmax['M' + str(i + 1)].number_format = "#,###.00;[RED]-#,###.00"
            # thursday
            overmax.merge_cells('N' + str(i + 1) + ':O' + str(i + 1))  # merge box for thr 5200
            overmax['N' + str(i)] = leavetime_array[5]  # leave time
            overmax['N' + str(i)].style = input_s
            overmax['N' + str(i)].number_format = "#,###.00;[RED]-#,###.00"
            overmax['O' + str(i)] = leavetype_array[5]  # leave type
            overmax['O' + str(i)].style = input_s
            overmax['N' + str(i + 1)] = totals_array[5]  # 5200 time
            overmax['N' + str(i + 1)].style = input_s
            overmax['N' + str(i + 1)].number_format = "#,###.00;[RED]-#,###.00"
            # friday
            overmax.merge_cells('P' + str(i + 1) + ':Q' + str(i + 1))  # merge box for fri 5200
            overmax['P' + str(i)] = leavetime_array[6]  # leave time
            overmax['P' + str(i)].style = input_s
            overmax['P' + str(i)].number_format = "#,###.00;[RED]-#,###.00"
            overmax['Q' + str(i)] = leavetype_array[6]  # leave type
            overmax['Q' + str(i)].style = input_s
            overmax['P' + str(i + 1)] = totals_array[6]  # 5200 time
            overmax['P' + str(i + 1)].style = input_s
            overmax['P' + str(i + 1)].number_format = "#,###.00;[RED]-#,###.00"
            # calculated fields
            # hidden columns
            formula_a = "=SUM(%s!D%s:%s!P%s)+%s!D%s + %s!H%s + %s!J%s + %s!L%s + " \
                      "%s!N%s + %s!P%s" % ("12_60_violations",str(i + 1),"12_60_violations",str(i + 1),
                       "12_60_violations",str(i),"12_60_violations",str(i),"12_60_violations",str(i),
                       "12_60_violations",str(i),"12_60_violations",str(i),"12_60_violations",str(i))
            overmax['R' + str(i)]= formula_a
            overmax['R' + str(i)].style = calcs
            overmax['R' + str(i)].number_format = "#,###.00;[RED]-#,###.00"
            formula_b = "=SUM(%s!C%s+%s!D%s+%s!H%s+%s!J%s+%s!L%s+%s!N%s+%s!P%s)" % \
                      ("12_60_violations",str(i),"12_60_violations",str(i),"12_60_violations",str(i),
                       "12_60_violations",str(i),"12_60_violations",str(i),"12_60_violations",str(i),
                       "12_60_violations",str(i))
            overmax['R' + str(i + 1)] = formula_b
            overmax['R' + str(i + 1)].style = calcs
            overmax['R' + str(i + 1)].number_format = "#,###.00;[RED]-#,###.00"
            # weekly violation
            overmax.merge_cells('S' + str(i) + ':S' + str(i + 1))  # merge box for weekly violation
            formula_c = "=MAX(IF(%s!R%s>%s!R%s,MAX(%s!R%s-60,0),MAX(%s!R%s-60)),0)" % ("12_60_violations",str(i),
                     "12_60_violations",str(i + 1),"12_60_violations",str(i),"12_60_violations",str(i + 1),)
            overmax['S' + str(i)] = formula_c
            overmax['S' + str(i)].style = calcs
            overmax['S' + str(i)].number_format = "#,###.00;[RED]-#,###.00"
            # daily violation
            formula_d = "=IF(OR(%s!B%s=\"wal\",%s!B%s=\"nl\",%s!B%s=\"aux\")," \
                      "(SUM(IF(%s!D%s>11.5,%s!D%s-11.5,0)+IF(%s!H%s>11.5,%s!H%s-11.5,0)+IF(%s!J%s>11.5,%s!J%s-11.5,0)" \
                      "+IF(%s!L%s>11.5,%s!L%s-11.5,0)+IF(%s!N%s>11.5,%s!N%s-11.5,0)+IF(%s!P%s>11.5,%s!P%s-11.5,0)))," \
                      "(SUM(IF(%s!D%s>12,%s!D%s-12,0)+IF(%s!H%s>12,%s!H%s-12,0)+IF(%s!J%s>12,%s!J%s-12,0)" \
                      "+IF(%s!L%s>12,%s!L%s-12,0)+IF(%s!N%s>12,%s!N%s-12,0)+IF(%s!P%s>12,%s!P%s-12,0))))" \
                       % ("12_60_violations",str(i),"12_60_violations",str(i),"12_60_violations",str(i),
                          "12_60_violations",str(i+1),"12_60_violations",str(i+1),"12_60_violations",str(i+1),
                          "12_60_violations",str(i+1),"12_60_violations",str(i+1),"12_60_violations",str(i+1),
                          "12_60_violations",str(i+1),"12_60_violations",str(i+1),"12_60_violations",str(i+1),
                          "12_60_violations",str(i+1),"12_60_violations",str(i+1),"12_60_violations",str(i+1),
                          "12_60_violations",str(i+1),"12_60_violations",str(i+1),"12_60_violations",str(i+1),
                          "12_60_violations",str(i+1),"12_60_violations",str(i+1),"12_60_violations",str(i+1),
                          "12_60_violations",str(i+1),"12_60_violations",str(i+1),"12_60_violations",str(i+1),
                          "12_60_violations",str(i+1),"12_60_violations",str(i+1),"12_60_violations",str(i+1))
            overmax['T' + str(i)] = formula_d
            overmax.merge_cells('T' + str(i) + ':T' + str(i + 1))  # merge box for daily violation
            overmax['T' + str(i)].style = calcs
            overmax['T' + str(i)].number_format = "#,###.00"
            # wed adjustment
            overmax.merge_cells('U' + str(i) + ':U' + str(i + 1))  # merge box for wed adj
            formula_e = "=IF(OR(%s!B%s=\"wal\",%s!B%s=\"nl\",%s!B%s=\"aux\")," \
                "IF(AND(%s!S%s-(%s!N%s+%s!N%s+%s!P%s+%s!P%s)>0,%s!L%s>11.5)," \
                "IF(%s!S%s-(%s!N%s+%s!N%s+%s!P%s+%s!P%s)>%s!L%s-11.5,%s!L%s-11.5,%s!S%s-(%s!N%s+%s!N%s+%s!P%s+%s!P%s)),0)," \
                "IF(AND(%s!S%s-(%s!N%s+%s!N%s+%s!P%s+%s!P%s)>0,%s!L%s>12)," \
                "IF(%s!S%s-(%s!N%s+%s!N%s+%s!P%s+%s!P%s)>%s!L%s-12,%s!L%s-12,%s!S%s-(%s!N%s+%s!N%s+%s!P%s+%s!P%s)),0))"\
                % ("12_60_violations",str(i),"12_60_violations",str(i),"12_60_violations",str(i),
                   "12_60_violations", str(i),"12_60_violations",str(i + 1),"12_60_violations",str(i),
                   "12_60_violations",str(i + 1),"12_60_violations",str(i),"12_60_violations",str(i + 1),
                   "12_60_violations", str(i), "12_60_violations", str(i + 1), "12_60_violations",str(i),
                   "12_60_violations", str(i + 1), "12_60_violations", str(i), "12_60_violations",str(i + 1),
                   "12_60_violations", str(i + 1), "12_60_violations", str(i), "12_60_violations",str(i + 1),
                   "12_60_violations", str(i), "12_60_violations", str(i + 1), "12_60_violations",str(i),
                   "12_60_violations", str(i), "12_60_violations", str(i+1), "12_60_violations", str(i),
                   "12_60_violations", str(i+1), "12_60_violations", str(i), "12_60_violations", str(i+1),
                   "12_60_violations", str(i), "12_60_violations", str(i + 1), "12_60_violations", str(i),
                   "12_60_violations", str(i + 1), "12_60_violations", str(i), "12_60_violations",str(i + 1),
                   "12_60_violations", str(i+1), "12_60_violations", str(i), "12_60_violations",str(i + 1),
                   "12_60_violations", str(i), "12_60_violations", str(i + 1), "12_60_violations", str(i))
            overmax['U' + str(i)] = formula_e
            overmax['U' + str(i)].style = vert_calcs
            overmax['U' + str(i)].number_format = "#,###.00"
            # thr adjustment
            formula_f = "=IF(OR(%s!B%s=\"wal\",%s!B%s=\"nl\",%s!B%s=\"aux\")," \
                      "IF(AND(%s!S%s-(%s!P%s+%s!P%s)>0,%s!N%s>11.5)," \
                      "IF(%s!S%s-(%s!P%s+%s!P%s)>%s!N%s-11.5,%s!N%s-11.5,%s!S%s-(%s!P%s+%s!P%s)),0)," \
                      "IF(AND(%s!S%s-(%s!P%s+%s!P%s)>0,%s!N%s>12)," \
                      "IF(%s!S%s-(%s!P%s+%s!P%s)>%s!N%s-12,%s!N%s-12,%s!S%s-(%s!P%s+%s!P%s)),0))" \
                      % ("12_60_violations",str(i),"12_60_violations",str(i),"12_60_violations",str(i),
                         "12_60_violations",str(i),"12_60_violations",str(i+1),"12_60_violations",str(i),
                         "12_60_violations",str(i+1),
                         "12_60_violations", str(i), "12_60_violations", str(i + 1), "12_60_violations", str(i),
                         "12_60_violations", str(i + 1), "12_60_violations", str(i + 1), "12_60_violations", str(i),
                         "12_60_violations", str(i + 1), "12_60_violations", str(i),
                         "12_60_violations", str(i), "12_60_violations", str(i + 1), "12_60_violations", str(i),
                         "12_60_violations", str(i + 1),
                         "12_60_violations", str(i), "12_60_violations", str(i + 1), "12_60_violations", str(i),
                         "12_60_violations", str(i + 1), "12_60_violations", str(i + 1), "12_60_violations", str(i),
                         "12_60_violations", str(i + 1), "12_60_violations", str(i)
                         )
            overmax.merge_cells('V' + str(i) + ':V' + str(i + 1))  # merge box for thr adj
            overmax['V' + str(i)] = formula_f
            overmax['V' + str(i)].style = vert_calcs
            overmax['V' + str(i)].number_format = "#,###.00"
            # fri adjustment
            overmax.merge_cells('W' + str(i) + ':W' + str(i + 1))  # merge box for fri adj
            formula_g = "=IF(OR(%s!B%s=\"wal\",%s!B%s=\"nl\",%s!B%s=\"aux\")," \
                    "IF(AND(%s!S%s>0,%s!P%s>11.5)," \
                    "IF(%s!S%s>%s!P%s-11.5,%s!P%s-11.5,%s!S%s),0)," \
                    "IF(AND(%s!S%s>0,%s!P%s>12)," \
                    "IF(%s!S%s>%s!P%s-12,%s!P%s-12,%s!S%s),0))" \
                    % ("12_60_violations", str(i),"12_60_violations", str(i),"12_60_violations", str(i),
                    "12_60_violations", str(i),"12_60_violations", str(i+1),"12_60_violations", str(i),
                    "12_60_violations", str(i+1),"12_60_violations", str(i+1),"12_60_violations", str(i),
                    "12_60_violations", str(i),"12_60_violations", str(i+1),"12_60_violations", str(i),
                    "12_60_violations", str(i+1),"12_60_violations", str(i+1),"12_60_violations", str(i))
            overmax['W' + str(i)] = formula_g
            overmax['W' + str(i)].style = vert_calcs
            overmax['W' + str(i)].number_format = "#,###.00"
            # total violation
            overmax.merge_cells('X' + str(i) + ':X' + str(i + 1))  # merge box for total violation
            formula_h = "=SUM(%s!S%s:%s!T%s)-(%s!U%s+%s!V%s+%s!W%s)" \
                      % ("12_60_violations", str(i),"12_60_violations", str(i),"12_60_violations", str(i),
                         "12_60_violations", str(i),"12_60_violations", str(i))
            overmax['X' + str(i)] = formula_h
            overmax['X' + str(i)].style = calcs
            overmax['X' + str(i)].number_format = "#,###.00"

            formula_i = "=%s!A%s" % ("12_60_violations", str(i))
            summary['A' + str(summary_i)] = formula_i
            summary['A' + str(summary_i)].style = input_name
            formula_j = "=%s!X%s" % ("12_60_violations", str(i))
            summary['B' + str(summary_i)] = formula_j
            summary['B' + str(summary_i)].style = input_s
            summary['B' + str(summary_i)].number_format = "#,###.00"
            summary.row_dimensions[summary_i].height = 10  # adjust all row height
            i += 2
            summary_i += 1
    # display totals for all violations
    overmax.merge_cells('P' + str(i + 1) + ':T' + str(i + 1))
    overmax['P' + str(i + 1)] = "Total Violations"
    overmax['P' + str(i + 1)].style = col_header
    overmax.merge_cells('V' + str(i+1) + ':X' + str(i+1))
    formula_k = "=SUM(%s!X%s:%s!X%s)" % ("12_60_violations", "9", "12_60_violations", str(i))
    overmax['V' + str(i+1)] = formula_k
    overmax['V' + str(i+1)].style = calcs
    overmax['V' + str(i+1)].number_format = "#,###.00"
    overmax.row_dimensions[i].height = 10  # adjust all row height
    overmax.row_dimensions[i+1].height = 10  # adjust all row height

    # name the excel file
    xl_filename = "kb_om" + str(format(g_date[0], "_%y_%m_%d")) + ".xlsx"
    ok = messagebox.askokcancel("Spreadsheet generator", "Do you want to generate a spreadsheet?")
    if ok == True:
        if os.path.isdir('kb_sub/over_max_spreadsheet') == False:
            os.makedirs('kb_sub/over_max_spreadsheet')
        try:
            wb.save('kb_sub/over_max_spreadsheet/' + xl_filename)
            messagebox.showinfo("Spreadsheet generator", "Your spreadsheet was successfully generated. \n"
                                                         "File is named: {}".format(xl_filename))
            if sys.platform == "win32":
                os.startfile('kb_sub\\over_max_spreadsheet\\' + xl_filename)
            if sys.platform == "linux":
                subprocess.call(["xdg-open", 'kb_sub/over_max_spreadsheet/' + xl_filename])
            if sys.platform == "darwin":
                subprocess.call(["open", 'kb_sub/over_max_spreadsheet/' + xl_filename])
        except:
            messagebox.showerror("Spreadsheet generator", "The spreadsheet was not generated. \n"
                                                          "Suggestion: "
                                                          "Make sure that identically named spreadsheets are closed "
                                                          "(the file can't be overwritten while open).")


def ns_config_apply(frame, text_array, color_array):
    for t in text_array:
        if len(t.get()) > 6:
            messagebox.showerror("Non_Scheduled Day Configuration",
                                 "Names must not be longer than 6 characters.")
            return
        if len(t.get()) < 1:
            messagebox.showerror("Non_Scheduled Day Configuration",
                                 "Names must not be shorter than 1 character.")
            return
    colors = ("yellow", "blue", "green", "brown", "red", "black")
    for i in range(6):
        sql = "UPDATE ns_configuration SET custom_name ='%s' WHERE ns_name = '%s'" % (text_array[i].get(), colors[i])
        commit(sql)
        sql = "UPDATE ns_configuration SET fill_color ='%s' WHERE ns_name = '%s'" % (color_array[i].get(), colors[i])
        commit(sql)
    ns_config(frame)


def ns_config_reset(frame):
    fill = ("gold", "navy", "forest green", "saddle brown", "red3", "gray10")
    colors = ("yellow", "blue", "green", "brown", "red", "black")
    for i in range(6):
        sql = "UPDATE ns_configuration SET custom_name ='%s' WHERE ns_name = '%s'" % (colors[i], colors[i])
        commit(sql)
        sql = "UPDATE ns_configuration SET fill_color ='%s' WHERE ns_name = '%s'" % (fill[i], colors[i])
        commit(sql)
    ns_config(frame)


def ns_config(frame):
    if gs_day == "x":
        messagebox.showerror("Non-Scheduled Day Configurations",
                             "You must set the Investigation Range before changing the NS Day Configurations.")
        return
    sql = "SELECT * FROM ns_configuration"
    result = inquire(sql)
    wd = front_window(frame)
    Label(wd[3], text="Non-Scheduled Day Configurations", font="bold", anchor="w").grid(row=0, sticky="w", columnspan=4)
    Label(wd[3], text=" ").grid(row=1, column=0)
    Label(wd[3], text="Change Configuration").grid(row=2, sticky="w", columnspan=4)
    f_date = g_date[0].strftime("%a - %b %d, %Y")
    end_f_date = g_date[6].strftime("%a - %b %d, %Y")
    Label(wd[3], text="Investigation Range: {0} through {1}".format(f_date, end_f_date),
          foreground="red").grid(row=3, column=0, sticky="w", columnspan=4)
    Label(wd[3], text="Pay Period: {0}".format(pay_period),
          foreground="red").grid(row=4, column=0, sticky="w", columnspan=4)
    Label(wd[3], text=" ").grid(row=5, column=0, sticky="w", columnspan=4)
    Label(wd[3], text="Day", foreground="grey").grid(row=6, column=0, sticky="w")  # column headers
    Label(wd[3], text="Name", foreground="grey").grid(row=6, column=1, sticky="w")
    Label(wd[3], text="Color", foreground="grey").grid(row=6, column=2, sticky="w")
    Label(wd[3], text="Default", foreground="grey").grid(row=6, column=3, sticky="w")
    yellow_text = StringVar(wd[3])  # declare variables
    blue_text = StringVar(wd[3])
    green_text = StringVar(wd[3])
    brown_text = StringVar(wd[3])
    red_text = StringVar(wd[3])
    black_text = StringVar(wd[3])
    text_array = [yellow_text, blue_text, green_text, brown_text, red_text, black_text]
    color_array = (
    "black", "blue", "brown", "brown4", "dark green", "deep pink", "forest green", "gold", "gray10", "green",
    "navy", "orange", "purple", "red", "red3", "saddle brown", "yellow", "yellow2")
    yellow_color = StringVar(wd[3])
    blue_color = StringVar(wd[3])
    green_color = StringVar(wd[3])
    brown_color = StringVar(wd[3])
    red_color = StringVar(wd[3])
    black_color = StringVar(wd[3])
    fill_array = [yellow_color, blue_color, green_color, brown_color, red_color, black_color]
    Label(wd[3], text="{}".format(ns_code['yellow'])).grid(row=7, column=0, sticky="w")  # yellow row
    Entry(wd[3], textvariable=yellow_text, width=10).grid(row=7, column=1, sticky="w")
    yellow_text.set(result[0][2])
    om_yellow = OptionMenu(wd[3], yellow_color, *color_array)
    yellow_color.set(result[0][1])
    om_yellow.config(width=13, anchor="w")
    om_yellow.grid(row=7, column=2, sticky="w")
    Label(wd[3], text="yellow").grid(row=7, column=3, sticky="w")
    Label(wd[3], text="{}".format(ns_code['blue'])).grid(row=8, column=0, sticky="w")  # blue row
    Entry(wd[3], textvariable=blue_text, width=10).grid(row=8, column=1, sticky="w")
    blue_text.set(result[1][2])
    om_blue = OptionMenu(wd[3], blue_color, *color_array)
    blue_color.set(result[1][1])
    om_blue.config(width=13, anchor="w")
    om_blue.grid(row=8, column=2, sticky="w")
    Label(wd[3], text="blue").grid(row=8, column=3, sticky="w")
    Label(wd[3], text="{}".format(ns_code['green'])).grid(row=9, column=0, sticky="w")  # green row
    Entry(wd[3], textvariable=green_text, width=10).grid(row=9, column=1, sticky="w")
    green_text.set(result[2][2])
    om_green = OptionMenu(wd[3], green_color, *color_array)
    green_color.set(result[2][1])
    om_green.config(width=13, anchor="w")
    om_green.grid(row=9, column=2, sticky="w")
    Label(wd[3], text="green").grid(row=9, column=3, sticky="w")
    Label(wd[3], text="{}".format(ns_code['brown'])).grid(row=10, column=0, sticky="w")  # brown row
    Entry(wd[3], textvariable=brown_text, width=10).grid(row=10, column=1, sticky="w")
    brown_text.set(result[3][2])
    om_brown = OptionMenu(wd[3], brown_color, *color_array)
    brown_color.set(result[3][1])
    om_brown.config(width=13, anchor="w")
    om_brown.grid(row=10, column=2, sticky="w")
    Label(wd[3], text="brown").grid(row=10, column=3, sticky="w")
    Label(wd[3], text="{}".format(ns_code['red'])).grid(row=11, column=0, sticky="w")  # red row
    Entry(wd[3], textvariable=red_text, width=10).grid(row=11, column=1, sticky="w")
    red_text.set(result[4][2])
    om_red = OptionMenu(wd[3], red_color, *color_array)
    red_color.set(result[4][1])
    om_red.config(width=13, anchor="w")
    om_red.grid(row=11, column=2, sticky="w")
    Label(wd[3], text="red").grid(row=11, column=3, sticky="w")
    Label(wd[3], text="{}".format(ns_code['black'])).grid(row=12, column=0, sticky="w")  # black row
    Entry(wd[3], textvariable=black_text, width=10).grid(row=12, column=1, sticky="w")
    black_text.set(result[5][2])
    om_black = OptionMenu(wd[3], black_color, *color_array)
    black_color.set(result[5][1])
    om_black.config(width=13, anchor="w")
    om_black.grid(row=12, column=2, sticky="w")
    Label(wd[3], text="black").grid(row=12, column=3, sticky="w")
    Label(wd[3], text=" ").grid(row=13)
    Button(wd[3], text="set", width=10, command=lambda: ns_config_apply(wd[0], text_array, fill_array)).grid(row=14,
                                                                                                             column=3)
    Label(wd[3], text=" ").grid(row=15)
    Label(wd[3], text="Restore Defaults").grid(row=16)
    Button(wd[3], text="reset", width=10, command=lambda: ns_config_reset(wd[0])).grid(row=17, column=3)

    Button(wd[4], text="Go Back", width=20, anchor="w",
           command=lambda: (wd[0].destroy(), main_frame())).pack(side=LEFT)
    rear_window(wd)


def get_file_path(subject_path):  # Created for pdf splitter - gets a pdf file
    file_path = filedialog.askopenfilename(initialdir=os.getcwd(),
                                           filetypes=[("PDF files", "*.pdf")], title="Select PDF")  # get the pdf file
    subject_path.set(file_path)


def get_new_path(new_path):  # Created for pdf splitter - creates/overwrites a pdf file
    save_filename = filedialog.asksaveasfilename(initialdir=os.getcwd(),
                                                 filetypes=[("PDF files", "*.pdf")], title="Overwrite/Create PDF")
    new_path.set(save_filename)


def pdf_splitter_apply(subject_path, firstpage, lastpage, new_path):
    # check for empty fields / return if there are any errors
    if subject_path == "":
        messagebox.showerror("Klusterbox PDF Splitter", "You must select a pdf file to split.")
        return
    if new_path == "":
        messagebox.showerror("Klusterbox PDF Splitter", "You must designate a destination"
                                                        " and a name for the df file you are creating.")
        return
    # if the last characters are not .pdf then add the extension
    if new_path[-4:] != ".pdf":
        new_path = new_path + ".pdf"
    if firstpage > lastpage:
        messagebox.showerror("Klusterbox PDF Splitter", "The First Page of the document can not be "
                                                        "higher than the Last Page.")
        return
    try:
        pdf = PdfFileReader(subject_path, "rb")
        pdf_writer = PdfFileWriter()
        for page in range(firstpage - 1, lastpage):
            pdf_writer.addPage(pdf.getPage(page))
        with open(new_path, 'wb') as out:
            pdf_writer.write(out)
        ok = messagebox.askokcancel("Klusterbox PDF Splitter", "PDF file has been split sucessfully."
                                                               "Do you want to open the pdf file?")
        if ok == True:
            if sys.platform == "win32":
                os.startfile(new_path)
            if sys.platform == "linux":
                subprocess.call(["xdg-open", new_path])
            if sys.platform == "darwin":
                subprocess.call(["open", new_path])
    except:
        messagebox.showerror("Klusterbox PDF Splitter", "The PDF splitting has failed. \n"
                                                        "It could be that that the pages set to be split don't exist \n"
                                                        "or \n"
                                                        "the pdf can't be split by this program due to formatting issues. \n"
                                                        "For better results try www.sodapdf.com, google chrome or Adobe Acrobat "
                                                        "Pro DC")


def pdf_splitter(frame):  # PDF Splitter
    wd = front_window(frame)
    Label(wd[3], text="PDF Splitter", font="bold", anchor="w") \
        .grid(row=1, column=1, columnspan=4, sticky="w")
    Label(wd[3], text="").grid(row=2)
    Label(wd[3], text="Select pdf file you want to split:") \
        .grid(row=3, column=1, columnspan=4, sticky="w")
    subject_path = StringVar(wd[3])
    Entry(wd[3], textvariable=subject_path, width=95).grid(row=4, column=1, columnspan=4)
    Button(wd[3], text="Select", width="10", command=lambda: get_file_path(subject_path)) \
        .grid(row=5, column=1, sticky="w")
    Label(wd[3], text="").grid(row=6)
    Label(wd[3], text="Select range of pages you want to use to create the new file:") \
        .grid(row=7, column=1, columnspan=4, sticky="w")
    Label(wd[3], text="First Page:  ").grid(row=8, column=1, sticky="e")
    firstpage = IntVar(wd[3])
    Entry(wd[3], textvariable=firstpage, width=8).grid(row=8, column=2, sticky="w")
    firstpage.set(1)
    Label(wd[3], text="Last Page:  ").grid(row=9, column=1, sticky="e")
    lastpage = IntVar(wd[3])
    Entry(wd[3], textvariable=lastpage, width=8).grid(row=9, column=2, sticky="w")
    lastpage.set(1)
    Label(wd[3], text="").grid(row=10)
    Label(wd[3], text="Select pdf file you want to over write or a create a new file:") \
        .grid(row=11, column=1, columnspan=4, sticky="w")
    new_path = StringVar(wd[3])
    Entry(wd[3], textvariable=new_path, width=95) \
        .grid(row=12, column=1, columnspan=4, sticky="w")
    Button(wd[3], text="Select", width="10", command=lambda: get_new_path(new_path)) \
        .grid(row=13, column=1, sticky="w")
    Label(wd[3], text="").grid(row=14)
    Label(wd[3], text="If all fields are filled out, split the file.") \
        .grid(row=15, column=1, columnspan=3, sticky="w")
    Button(wd[3], text="Split PDF", width="10", command=lambda: pdf_splitter_apply(
        subject_path.get().strip(),
        firstpage.get(),
        lastpage.get(),
        new_path.get().strip())).grid(row=15, column=4, sticky="e")
    Button(wd[4], text="Go Back", width=20, anchor="w",
           command=lambda: (wd[0].destroy(), main_frame())).pack(side=LEFT)
    rear_window(wd)


def pdf_converter_settings_apply(frame, error, raw, txt):
    sql = "UPDATE tolerances SET tolerance='%s'WHERE category='%s'" % (error.get(), "pdf_error_rpt")
    commit(sql)
    sql = "UPDATE tolerances SET tolerance='%s'WHERE category='%s'" % (raw.get(), "pdf_raw_rpt")
    commit(sql)
    sql = "UPDATE tolerances SET tolerance='%s'WHERE category='%s'" % (txt.get(), "pdf_text_reader")
    commit(sql)
    pdf_converter_settings(frame)


def pdf_converter_settings(frame):
    sql = "SELECT tolerance FROM tolerances WHERE category ='%s'" % ("pdf_error_rpt")
    result = inquire(sql)
    wd = front_window(frame)
    Label(wd[3], text="PDF Converter Settings", font="bold", anchor="w").grid(row=0, sticky="w", columnspan=4)
    Label(wd[3], text=" ").grid(row=1, column=0)
    Label(wd[3], text="Generate Reports for PDF Converter").grid(row=2, sticky="w", columnspan=4)
    Label(wd[3], text=" ").grid(row=3, column=0)
    Label(wd[3], text="Error Report", width=15, anchor="w").grid(row=4, column=0, sticky="w")
    error_selection = StringVar(wd[3])
    om_error = OptionMenu(wd[3], error_selection, "on", "off")
    om_error.config(width=5, anchor="w")
    om_error.grid(row=4, column=1)
    error_selection.set(result[0][0])
    sql = "SELECT tolerance FROM tolerances WHERE category ='%s'" % ("pdf_raw_rpt")
    result = inquire(sql)
    Label(wd[3], text="Raw Output Report", width=15, anchor="w").grid(row=5, column=0, sticky="w")
    raw_selection = StringVar(wd[3])
    om_raw = OptionMenu(wd[3], raw_selection, "on", "off")
    om_raw.config(width=5, anchor="w")
    om_raw.grid(row=5, column=1)
    raw_selection.set(result[0][0])
    Label(wd[3], text=" ").grid(row=6, column=0)
    # allow user to read from a text file to bypass the pdfminer
    Label(wd[3], text="Generate Reports from Text file").grid(row=7, sticky="w", columnspan=4)
    Label(wd[3], text="     (where a text file of pdfminer output has been generated)").grid(row=8, sticky="w",
                                                                                             columnspan=4)
    Label(wd[3], text=" ").grid(row=9, column=0)
    sql = "SELECT tolerance FROM tolerances WHERE category ='%s'" % ("pdf_text_reader")
    result = inquire(sql)
    Label(wd[3], text="Read from txt file", width=15, anchor="w").grid(row=10, column=0, sticky="w")
    txt_selection = StringVar(wd[3])
    om_txt = OptionMenu(wd[3], txt_selection, "on", "off")
    om_txt.config(width=5, anchor="w")
    om_txt.grid(row=10, column=1)
    txt_selection.set(result[0][0])
    Label(wd[3], text=" ").grid(row=11, column=0)

    Button(wd[3], text="set", width=10, command=lambda:
    pdf_converter_settings_apply(wd[0], error_selection, raw_selection, txt_selection)) \
        .grid(row=12, column=2)
    Button(wd[4], text="Go Back", width=20, anchor="w",
           command=lambda: (wd[0].destroy(), main_frame())).pack(side=LEFT)
    rear_window(wd)


def pdf_converter_pagecount(filepath):  # gives a page count for pdf_to_text
    file = open(filepath, 'rb')
    parser = PDFParser(file)
    document = PDFDocument(parser)
    page_count = resolve1(document.catalog['Pages'])['Count']  # This will give you the count of pages
    return (page_count)


def pdf_to_text(filepath):  # Called by pdf_converter() to read pdfs with pdfminer
    codec = 'utf-8'
    password = ""
    maxpages = 0
    caching = (True,True)
    pagenos = set()
    laparams = (
        LAParams(
            line_overlap=.1, #best results
            char_margin=2,
            line_margin=.5,
            word_margin=.5,
            boxes_flow=0,
            detect_vertical=True,
            all_texts=True),
        LAParams(
            line_overlap=.5, # default settings
            char_margin=2,
            line_margin=.5,
            word_margin=.5,
            boxes_flow=.5,
            detect_vertical=False,
            all_texts=False)
        )
    for i in range(2):
        retstr = StringIO()
        rsrcmgr = PDFResourceManager()
        device = TextConverter(rsrcmgr, retstr, codec=codec, laparams=laparams[i])
        interpreter = PDFPageInterpreter(rsrcmgr, device)
        page_count = pdf_converter_pagecount(filepath)  # get page count
        with open(filepath, 'rb') as filein:
            # create progressbar
            pb_root = Tk()  # create a window for the progress bar
            pb_root.geometry("%dx%d+%d+%d" % (450, 75, 200, 300))
            pb_root.title("Klusterbox PDF Converter - reading pdf")
            Label(pb_root, text="This process takes several minutes. Please wait for results.").pack(anchor="w", padx=20)
            pb_label = Label(pb_root, text="Reading PDF: ")  # make label for progress bar
            pb_label.pack(anchor="w", padx=20)
            pb = ttk.Progressbar(pb_root, length=400, mode="determinate")  # create progress bar
            pb.pack(anchor="w", padx=20)
            pb["maximum"] = page_count  # set length of progress bar
            pb.start()
            count = 0
            for page in PDFPage.get_pages(filein, pagenos, maxpages=maxpages, password=password, caching=caching[i],
                                          check_extractable=True):
                interpreter.process_page(page)
                pb["value"] = count  # increment progress bar
                pb_root.update()
                count += 1
            text = retstr.getvalue()
            device.close()
            retstr.close()
        pb.stop()  # stop and destroy the progress bar
        pb_label.destroy()  # destroy the label for the progress bar
        pb.destroy()
        pb_root.destroy()
        # test the results
        text = text.replace("","")
        page = text.split("")  # split the document into page
        result = re.search("Restricted USPS T&A Information(.*)Employee Everything Report", page[0], re.DOTALL)
        try:
            station = result.group(1).strip()
            break
        except:
            if i<1:
                result = messagebox.askokcancel("Klusterbox PDF Converter",
                                     "PDF Conversion has failed and will not generate a file.  \n\n"
                                     "We will try again.")
                if result == False:
                    return text
            else:
                messagebox.showerror("Klusterbox PDF Converter",
                                     "PDF Conversion has failed and will not generate a file.  \n\n"
                                     "You will either have to obtain the Employee Everything Report "
                                     "in the csv format from management or manually enter in the "
                                     "information")
    return text


def pdf_converter_reorder_founddays(found_days):
    new_order = []
    correct_series = ("Saturday", "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday")
    for cs in correct_series:
        if cs in found_days:
            new_order.append(cs)
    return new_order


def pdf_converter_path_generator(file_path, add_on, extension):  # generate csv file name and path
    file_parts = file_path.split("/")  # split path into folders and file
    file_name_xten = file_parts[len(file_parts) - 1]  # get the file name from the end of the path
    file_name = file_name_xten[:-4]  # remove the file extension from the file name
    file_name = file_name.replace("_raw_kbpc", "")
    path = file_path[:-len(file_name_xten)]  # get the path back to the source folder
    new_fname = file_name + add_on  # add suffix to to show converted pdf to csv
    new_file_path = path + new_fname + extension  # new path with modified file name
    return new_file_path


def pdf_converter_short_name(file_path):
    file_parts = file_path.split("/")  # split path into folders and file
    file_name_xten = file_parts[len(file_parts) - 1]  # get the file name from the end of the path
    return file_name_xten


def pdf_converter():
    # inquire as to if the pdf converter reports have been opted for by the user
    sql = "SELECT tolerance FROM tolerances WHERE category ='%s'" % ("pdf_error_rpt")
    result = inquire(sql)
    gen_error_report = result[0][0]
    sql = "SELECT tolerance FROM tolerances WHERE category ='%s'" % ("pdf_raw_rpt")
    result = inquire(sql)
    gen_raw_report = result[0][0]
    starttime = time.time()  # start the timer
    # make it possible for user to select text file
    sql = "SELECT tolerance FROM tolerances WHERE category ='%s'" % ("pdf_text_reader")
    result = inquire(sql)
    allow_txt_reader = result[0][0]
    if allow_txt_reader == "on":
        preference = messagebox.askyesno("PDF Converter",
                                         "Did you want to read from a text file of data output by pdfminer?")
    else:
        preference = False
    if preference == False:  # user opts to read from pdf file
        file_path = filedialog.askopenfilename(initialdir=os.getcwd(),
                                               filetypes=[("PDF files", "*.pdf")])  # get the pdf file
        new_file_path = pdf_converter_path_generator(file_path, "_kbpc", ".csv")  # generate csv file name and path
        short_file_name = pdf_converter_short_name(new_file_path)
        # if the file path already exist - ask for confirmation
        if os.path.exists(new_file_path):
            proceed = messagebox.askokcancel("Possible File Name Discrepancy", "There is already a file named {}. "
               "If you proceed, the file will be overwritten. Did you want to proceed?".format(
                short_file_name))
            if proceed == False:
                return
        # warn user that the process can take several minutes
        proceed = messagebox.askokcancel("PDF Converter", "This process will take several minutes. "
                                                          "Did you want to proceed?")
        if proceed == False:
            return
        text = pdf_to_text(file_path)  # read the pdf with pdfminer
    else:  # user opts to read from text file
        file_path = filedialog.askopenfilename(initialdir=os.getcwd(),
                                               filetypes=[("text files", "*.txt")])  # get the pdf file
        new_file_path = pdf_converter_path_generator(file_path, "_kbpc", ".csv")  # generate csv file name and path
        short_file_name = pdf_converter_short_name(new_file_path)
        # if the file path already exist - ask for confirmation
        if os.path.exists(new_file_path):
            proceed = messagebox.askokcancel(
                "Possible File Name Discrepancy",
                "There is already a file named {}. If you proceed, the file will be overwritten. "
                "Did you want to proceed?".format(short_file_name))
            if proceed == False:
                return
        gen_raw_report = "off"  # since you are reading a raw report, turn off the generator
        with open(file_path, 'r') as file:  # read the txt file and put it in the text variable
            text = file.read()
    # put the raw output from the pdf conversion into a text file
    if gen_raw_report == "on":
        kbpc_raw_rpt_file_path = pdf_converter_path_generator \
            (file_path, "_raw_kbpc", ".txt")  # generate csv file name and path
        kbpc_raw_rpt = open(kbpc_raw_rpt_file_path, "w")
        kbpc_raw_rpt.write("KLUSTERBOX PDF CONVERSION REPORT \n\n")
        kbpc_raw_rpt.write("Raw output from pdf miner\n\n")
        input = "subject file: {}\n\n".format(file_path)
        kbpc_raw_rpt.write(input)
        kbpc_raw_rpt.write(text)
        kbpc_raw_rpt.close()
    # create text document for data extracted from the raw pdfminer output
    if gen_error_report == "on":
        kbpc_rpt_file_path = pdf_converter_path_generator(file_path, "_kbpc", ".txt")  # generate csv file name and path
        kbpc_rpt = open(kbpc_rpt_file_path, "w")
        kbpc_rpt.write("KLUSTERBOX PDF CONVERSION REPORT \n\n")
        kbpc_rpt.write("Data extracted from pdfminer output and error reports\n\n")
        input = "subject file: {}\n\n".format(file_path)
        kbpc_rpt.write(input)
    # define csv writer parameters
    csv.register_dialect('myDialect',
                         delimiter=',',
                         quoting=csv.QUOTE_NONE,
                         skipinitialspace=True,
                         lineterminator="\r"
                         )
    # create the csv file and write the first line
    line = ["TAC500R3 - Employee Everything Report"]
    with open(new_file_path, 'w') as writeFile:
        writer = csv.writer(writeFile, dialect='myDialect')
        writer.writerow(line)
    # define csv writer parameters
    csv.register_dialect('myDialect',
                         delimiter=',',
                         quoting=csv.QUOTE_ALL,
                         skipinitialspace=True,
                         lineterminator=",\r"
                         )
    line = ["YrPPWk", "Finance No", "Organization Name", "Sub-Unit", "Employee Id", "Last Name", "FI", "MI",
            "Pay Loc/Fin Unit", "Var. EAS", "Borrowed", "Auto H/L", "Annual Lv Bal", "Sick Lv Bal", "LWOP Lv Bal",
            "FMLA Hrs", "FMLA Used", "SLDC Used", "Job", "D/A", "LDC", "Oper/Lu", "RSC", "Lvl", "FLSA", "Route #",
            "Loaned Fin #", "Effective Start", "Effective End", "Begin Tour", "End Tour", "Lunch Amt", "1261 Ind",
            "Lunch Ind", "Daily Sched Ind", "Time Zone", "FTF", "OOS", "Day", ]
    with open(new_file_path, 'a') as writeFile:
        writer = csv.writer(writeFile, dialect='myDialect')
        writer.writerow(line)
    text = text.replace("","")
    page = text.split("")  # split the document into pages
    whole_line = []
    page_num = 1  # initialize var to count pages
    eid_count = 0  # initialize var to count underscore dash items
    underscore_slash = []  # arrays for building daily array
    daily_underscoreslash = []
    mv_holder = []
    time_holder = []
    timezone_holder = []
    finance_holder = []
    foundday_holder = []
    daily_array = []
    franklin_array = []
    mv_desigs = ("BT", "MV", "ET", "OT", "OL", "IL", "DG")
    days = ("Saturday", "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday")
    saved_pp = ""  # hold the pp to identify if it changes
    pp_days = []  # array of date/time objs for each day in the week
    found_days = []  # array for holding days worked
    base_time = []  # array for holding hours worked during the day
    eid = ""  # hold the employee id
    lastname = ""  # holds the last name of the employee
    fi = ""
    jobs = []  # holds the d/a code
    routes = []  # holds the route
    level = []  # hold the level (one or two normally)
    base_temp = ("Base", "Temp")
    eid_label = False
    lookforname = False
    lookforfi = False
    lookforroute = False
    lookfor2route = False
    lookforlevel = False
    lookfor2level = False
    base_counter = 0
    base_chg = 0
    lookfortimes = False
    unprocessedrings = ""
    new_page = False
    unprocessed_counter = 0
    mcgrath_indicator = False
    mcgrath_carryover = ""
    rod_rpt = []  # error reports
    frank_rpt = []
    rose_rpt = []
    robert_rpt = []
    stevens_rpt = []
    carroll_rpt = []
    nguyen_rpt = []
    salih_rpt = []
    unruh_rpt = []
    mcgrath_rpt = []
    unresolved = []
    basecounter_error = []
    failed = []
    result = re.search('Restricted USPS T&A Information(.*?)Employee Everything Report', page[0], re.DOTALL)
    try:
        station = result.group(1).strip()
    except:
        messagebox.showerror("Klusterbox PDF Converter",
                             "This file does not appear to be an Employee Everything Report. \n\n"
                             "The PDF Converter will not generate a file")
        os.remove(new_file_path)
        if gen_error_report == "on":
            kbpc_rpt.close()
            os.remove(kbpc_rpt_file_path)
        if gen_raw_report == "on":
            os.remove(kbpc_raw_rpt_file_path)
        return
    # start the progress bar
    pb_root = Tk()  # create a window for the progress bar
    pb_root.geometry("%dx%d+%d+%d" % (450, 75, 200, 300))
    pb_root.title("Klusterbox PDF Converter - translating pdf")
    Label(pb_root, text="This process takes several minutes. Please wait for results.").pack(anchor="w", padx=20)
    pb_label = Label(pb_root, text="Translating PDF: ")  # make label for progress bar
    pb_label.pack(anchor="w", padx=20)
    pb = ttk.Progressbar(pb_root, length=400, mode="determinate")  # create progress bar
    pb.pack(anchor="w", padx=20)
    pb["maximum"] = len(page) - 1  # set length of progress bar
    pb.start()
    pb_count = 0
    for a in page:
        if gen_error_report == "on": kbpc_rpt.write(
            "\nxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
            "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx\n")
        if a[0:6] == "Report" or a[0:6] == "":
            restrictedisfirst = False
        else:
            if gen_error_report == "on": kbpc_rpt.write("Out of Sequence Problem!\n")
            restrictedisfirst = True
            eid_count = 0
        if gen_error_report == "on":
            input = "Page: {}\n".format(page_num)
            kbpc_rpt.write(input)
        try:  # if the page has no station information, then break the loop.
            result = re.search("Restricted USPS T&A Information(.*)Employee Everything Report", a, re.DOTALL)
            station = result.group(1).strip()
            station = station.split('\n')[0]
            if len(station) == 0:
                result = re.search("Employee Everything Report(.*)Weekly", a, re.DOTALL)
                station = result.group(1).strip()
                station = station.split('\n')[0]
        except:
            break
        try:
            result = re.search("YrPPWk:\nSub-Unit:\n\n(.*)\n", a)
            yyppwk = result.group(1)
        except:
            result = re.search("YrPPWk:\n\n(.*)\n\nFin. #:", a)
            yyppwk = result.group(1)
        if saved_pp != yyppwk:
            exploded = yyppwk.split("-")  # break up the year/pp string from the ee rpt pdf
            year = exploded[0]  # get the year
            if gen_error_report == "on":
                input = "Year: {}\n".format(year)
                kbpc_rpt.write(input)
            pp = exploded[1]  # get the pay period
            if gen_error_report == "on":
                input = "Pay Period: {}\n".format(pp)
                kbpc_rpt.write(input)
            pp_wk = exploded[2]  # get the week of the pay period
            if gen_error_report == "on":
                input = "Pay Period Week: {}\n".format(pp_wk)
                kbpc_rpt.write(input)
            pp = pp + pp_wk  # join the pay period and the week
            first_date = find_pp(int(year), pp)  # get the first day of the pay period
            if gen_error_report == "on":
                input = "{}\n".format(str(first_date))
                kbpc_rpt.write(input)
            pp_days = []  # build an array of date/time objects for each day in the pay period
            daily_array_days = []  # build an array of formatted days with just month/ day
            for _ in range(7):
                pp_days.append(first_date)
                daily_array_days.append(first_date.strftime("%m/%d"))
                first_date += timedelta(days=1)
            if gen_error_report == "on":
                input = "Days in Pay Period: {}\n".format(pp_days)
                kbpc_rpt.write(input)
            saved_pp = yyppwk  # hold the year/pp to check if it changes
        page_num += 1
        b = a.split("\n\n")
        for c in b:
            # find, categorize and record daily times
            if lookfortimes == True:
                if re.match(r"0[0-9]{4}\:\s0[0-9]{2}\.[0-9]{2}$", c):
                    to_add = [base_counter, c]
                    base_time.append(to_add)
                    base_chg = base_counter  # value to check for errors+
                # solve for robertson basetime problem / Base followed by H/L
                elif re.match(r"0[0-9]{4}\:\s0[0-9]{2}\.[0-9]{2}\n0[0-9]{4}\:\s0[0-9]{2}\.[0-9]{2}", c):
                    if "\n" not in c:  # check that there are no multiple times in the line
                        to_add = [base_counter, c]
                        base_time.append(to_add)
                        base_chg = base_counter  # value to check for errors
                        robert_rpt.append(lastname)  # data for robertson baseline problem
                    elif "\n" in c:  # if there are multiple times in the line
                        split_base = c.split("\n")  # split the times by the line break
                        for sb in split_base:  # add each time individually
                            to_add = [base_counter, sb]  # combine the base counter with the time
                            base_time.append(to_add)  # add that time to the array of base times
                            base_chg = base_counter  # value to check for errors
                else:
                    base_counter += 1
                    lookfortimes = False
            if re.match(r"Base", c):
                lookfortimes = True
            # solve for stevens problem / H/L base times not being read
            if len(finance_holder) == 0 and re.match(r"H/L\s", c):  # set trap to catch daily times
                lookfortimes = True
                stevens_rpt.append(lastname)
            checker = False
            one_mistake = False
            underscore_slash = c.split("\n")
            for us in underscore_slash:  # loop through items to detect matches
                if re.match(r"[0-1][0-9]\/[0-9][0-9]", us) or us == "__/__":
                    checker = True
                else:
                    one_mistake = True
            if len(underscore_slash) > 1 and checker == True and one_mistake == False:
                daily_underscoreslash.append(underscore_slash)
            underscore_slash = []
            d = c.split("\n")
            for e in d:
                try:
                    # build the daily array
                    if re.match(r"[0-9]{6}$", e) and len(movecode_holder) != 0:  # get the route following the chain
                        movecode_holder.append(e)
                        route_holder = movecode_holder
                        if unprocessedrings == "":
                            daily_array.append(route_holder)
                        else:
                            unprocessed_counter += 1  # handle carroll problem
                            carroll_rpt.append(lastname)  # append carroll report
                    movecode_holder = []
                    if len(finance_holder) != 0:  # get the move code following the chain
                        if re.match(r"[0-9]{4}\-[0-9]{2}$", e):
                            finance_holder.append(e)
                            movecode_holder = finance_holder
                        # solve for robertson problem / "H/L" is in move code
                        if re.match(r"H/L", e):  # if the move code is a higher level assignment
                            finance_holder.append(e)
                            finance_holder.append("000000")  # insert zeros for route number
                            if unprocessedrings == "":
                                daily_array.append(finance_holder)  # skip getting the route and create append daily array
                            else:
                                unprocessed_counter += 1  # handle carroll problem
                                carroll_rpt.append(lastname)  # append carroll report
                    finance_holder = []
                    if len(timezone_holder) != 0:  # get the finance number following the chain
                        timezone_holder.append(e)
                        finance_holder = timezone_holder
                    timezone_holder = []
                    if re.match(r"[A-Z]{2}T", e) and len(time_holder) != 0:  # look for the time zone following chain
                        time_holder.append(e)
                        timezone_holder = time_holder
                    elif len(
                            time_holder) != 0 and unprocessedrings != "":  # solve for salih problem / missing time zone in ...
                        unprocessed_counter += 1  # unprocessed rings
                        salih_rpt.append(lastname)
                    time_holder = []
                    if re.match(r" [0-2][0-9]\.[0-9][0-9]$", e) and len(
                            date_holder) != 0:  # look for time following date/mv desig
                        date_holder.append(e)
                        time_holder = date_holder
                    # look for items in franklin array to solve for franklin problem
                    if len(franklin_array) > 0 and re.match(r"[0-1][0-9]\/[0-3][0-9]$", e):  # if franklin array and date
                        frank = franklin_array.pop(0)  # pop out the earliest mv desig
                        mv_holder = []
                        mv_holder.append(eid)  # rebuild the mv holder array
                        mv_holder.append(frank)  # place in a holder and check the next line for a date
                    # solve for rodriguez problem / multiple consecutive mv desigs
                    if len(franklin_array) > 0:
                        if re.match(r"0[0-9]{4}$", e) or re.match(r"0[0-9]{2}$",
                                                                  e) or e in mv_desigs:  # look for move desig
                            franklin_array.append(e)
                            rod_rpt.append(lastname)
                    date_holder = []
                    if re.match(r"[0-1][0-9]\/[0-3][0-9]$", e) and len(
                            mv_holder) != 0:  # look for date following move desig
                        mv_holder.append(e)
                        date_holder = mv_holder
                    # solve for franklin problem: two mv desigs appear consecutively
                    if len(mv_holder) > 0:
                        if re.match(r"0[0-9]{4}$", e) or re.match(r"0[0-9]{2}$",
                                                                  e) or e in mv_desigs:  # look for move desig
                            franklin_array.append(mv_holder[1])
                            franklin_array.append(e)
                            frank_rpt.append(lastname)
                    mv_holder = []
                    if len(franklin_array) == 0:
                        if re.match(r"0[0-9]{4}$", e) or re.match(r"0[0-9]{2}$",
                                                                  e) or e in mv_desigs:  # look for move desig
                            mv_holder.append(eid)
                            mv_holder.append(e)  # place in a holder and check the next line for a date
                    # solve for rose problem: mv desig and date appearing on same line
                    if re.match(r"0[0-9]{4}\s[0-2][0-9]\/[0-9][0-9]$", e):
                        rose = e.split(" ")
                        mv_holder.append(eid)  # add the emp id to the daily array
                        mv_holder.append(rose[0])  # add the mv desig to the daily array
                        mv_holder.append(rose[1])  # add the date to the mv desig array
                        date_holder = mv_holder  # transfer array items to date holder
                        rose_rpt.append(lastname)
                    if e in days:  # find and record all days on the report
                        if eid_label == True:
                            found_days.append(e)
                        if eid_label == False:
                            foundday_holder.append(e)
                    if e == "Processed Clock Rings":
                        eid_count = 0
                    # if e =="Employee ID" and restrictedisfirst==False: # find the employee id label
                    if e == "Employee ID":
                        eid_label = True
                        if gen_error_report == "on":
                            if len(jobs) > 0:
                                input = "Jobs: {}\n".format(jobs)
                                kbpc_rpt.write(input)
                            if len(routes) > 0:
                                input = "Routes: {}\n".format(routes)
                                kbpc_rpt.write(input)
                            if len(level) > 0:
                                input = "Levels: {}\n".format(level)
                                kbpc_rpt.write(input)
                            if len(base_time) > 0:
                                kbpc_rpt.write("Base / Times:")
                                for bt in base_time:
                                    input = "{}\n".format(bt)
                                    kbpc_rpt.write(input)
                        if len(daily_underscoreslash) > 0:  # bind all underscore slash items in one array
                            underscore_slash_result = sum(daily_underscoreslash, [])
                        # write to csv file
                        prime_info = [yyppwk.replace("-", ""), '"{}"'.format("000000"), '"{}"'.format(station),
                                      '"{}"'.format("0000"), '"{}"'.format(eid), '"{}"'.format(lastname),
                                      '"{}"'.format(fi[:1]),
                                      '"_"', '"010/0000"', '"N"', '"N"', '"N"', '"0"', '"0"', '"0"', '"0"', '"0"',
                                      '"0"']
                        count = 0
                        for array in daily_array:
                            array.append(underscore_slash_result[count])
                            array.append(underscore_slash_result[count + 1])
                            count += 2
                        if base_chg + 1 != len(found_days):  # add to basecounter error array
                            to_add = (lastname, base_chg, len(found_days))
                            if len(found_days) > 0:
                                basecounter_error.append(to_add)
                        # set up array for each day in the week
                        csv_sat = []
                        csv_sun = []
                        csv_mon = []
                        csv_tue = []
                        csv_wed = []
                        csv_thr = []
                        csv_fri = []
                        csv_output = [csv_sat, csv_sun, csv_mon, csv_tue, csv_wed, csv_thr, csv_fri]
                        # reorder the found days to ensure the correct order
                        found_days = pdf_converter_reorder_founddays(found_days)
                        # fix problem with miscounted base times
                        high_array = []
                        for bt in base_time:
                            high_array.append(bt[0])
                        if len(high_array) > 0:
                            high_num = max(high_array)
                            comp_array = []
                            for i in range(high_num + 1):
                                comp_array.append(i)
                            del_array = []
                            for num in comp_array:
                                if num in high_array:
                                    del_array.append(num)
                            error_array = comp_array
                            error_array = [x for x in error_array if x not in del_array]
                            error_array.reverse()
                            if len(error_array) > 0:
                                for error_num in error_array:
                                    for bt in base_time:
                                        if bt[0] > error_num:
                                            bt[0] = bt[0] - 1
                        # load the multi array with array for each day
                        if len(foundday_holder) > 0:
                            # solve for nguyen problem / day of week occurs prior to "employee id" label
                            found_days = found_days + foundday_holder
                            ordered_days = ["Saturday", "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
                            for day in days:  # re order days into correct order
                                if day not in found_days:
                                    ordered_days.remove(day)
                            found_days = ordered_days
                            # foundday_holder = []
                            nguyen_rpt.append(lastname)
                        if len(found_days) > 0:  # printe out found days
                            # reorder the found days to ensure the correct order
                            found_days = pdf_converter_reorder_founddays(found_days)
                            if gen_error_report == "on":
                                input = "Found days: {}\n".format(found_days)
                                kbpc_rpt.write(input)
                        if gen_error_report == "on":
                            input = "proto emp id counter: {}\n".format(eid_count)
                            kbpc_rpt.write(input)
                        for i in range(7):
                            for bt in base_time:
                                if found_days[bt[0]] == days[i]:
                                    csv_output[i].append(bt)
                            for da in daily_array:
                                if da[2] == pp_days[i].strftime("%m/%d"):
                                    csv_output[i].append(da)
                        for co in csv_output:  # for each time in the array, printe a line
                            for array in co:
                                if gen_error_report == "on":
                                    input = "{}\n".format(array)
                                    kbpc_rpt.write(input)
                                # put the data into the csv file
                                if len(array) == 2:  # if the line comes from base/time data
                                    add_this = [found_days[int(array[0])], '"_0-00"', '"{}"'.format(array[1])]
                                    whole_line = prime_info + add_this
                                    with open(new_file_path, 'a') as writeFile:  # add the line to the csv file
                                        writer = csv.writer(writeFile, dialect='myDialect')
                                        writer.writerow(whole_line)
                                if len(array) == 10:  # if the line comes from daily array
                                    if array[9] != "__/__":
                                        end_notes = "(W)Ring Deleted From PC"
                                    else:
                                        end_notes = ""
                                    add_this = ["000-00", '"{}"'.format(array[1]),
                                                '"{}"'.format(
                                                    pp_days[daily_array_days.index(array[2])].strftime("%d-%b-%y").upper()),
                                                '"{}"'.format(array[3].strip()), '"{}"'.format(array[5]),
                                                '"{}"'.format(array[6]),
                                                '"{}"'.format(array[7]), '""', '""', '""', '"0"', '""', '""', '"0"',
                                                '"{}"'.format(end_notes)]
                                    whole_line = prime_info + add_this
                                    with open(new_file_path, 'a') as writeFile:  # add the line to the csv file
                                        writer = csv.writer(writeFile, dialect='myDialect')
                                        writer.writerow(whole_line)
                        # define csv writer parameters
                        csv.register_dialect('myDialect',
                                             delimiter=',',
                                             quotechar="'",
                                             skipinitialspace=True,
                                             lineterminator=",\r"
                                             )
                        if len(jobs) > 0:
                            for i in range(len(jobs)):
                                base_line = [base_temp[i], '"{}"'.format(jobs[i].replace("-", "").strip()),
                                             '"0000"', '"7220-10"',
                                             '"Q0"', '"{}"'.format(level[i]), '"N"', '"{}"'.format(routes[i]), '""',
                                             '"0000000"',
                                             '"0000000"', '"0"', '"0"', '"0"', '"N"', '"N"', '"N"', '"MDT"', '"N"']
                                whole_line = prime_info + base_line
                                with open(new_file_path, 'a') as writeFile:
                                    writer = csv.writer(writeFile, dialect='myDialect')
                                    writer.writerow(whole_line)
                        found_days = []  # initialized arrays
                        lookfortimes = False
                        base_time = []
                        eid = ""
                        base_chg = 0
                        base_counter = 0
                        daily_array = []
                        daily_underscoreslash = []
                        unprocessed_counter = 0
                        jobs = []
                        level = []
                        if gen_error_report == "on":
                            input = "{}\n".format(e)
                            kbpc_rpt.write(input)
                        eid_count = 0
                    if lookforfi == True:  # look for first initial
                        if re.fullmatch("[A-Z]\s[A-Z]", e) or re.fullmatch("([A-Z])", e):
                            if gen_error_report == "on":
                                input = "FI: {}\n".format(e)
                                kbpc_rpt.write(input)
                            fi = e
                            lookforfi = False
                    if lookforname == True:  # look for the name
                        if re.fullmatch(r"([A-Z]+)", e) \
                                or re.fullmatch(r"([A-Z]+.[A-Z]+)", e) \
                                or re.fullmatch(r"([A-Z]+.[A-Z]+.[A-Z]+)", e) \
                                or re.fullmatch(r"([A-Z]+.[A-Z]+.[A-Z]+.[A-Z]+)", e) \
                                or re.fullmatch(r"([A-Z]+.[A-Z]+.[A-Z]+.[A-Z]+.[A-Z]+)", e):
                            lastname = e.replace("'"," ")
                            print(lastname)
                            if gen_error_report == "on":
                                input = "Name: {}\n".format(e)
                                kbpc_rpt.write(input)
                            lookforname = False
                            lookforfi = True
                    if re.match(r"\s[0-9]{2}\-[0-9]$", e):  # find the job or d/a code - there might be two
                        jobs.append(e)
                    if lookfor2route == True:  # look for temp route
                        if re.match(r"[0-9]{6}$", e):
                            routes.append(e)  # add route to routes array
                        lookfor2route = False
                    if lookforroute == True:  # look for main route
                        if re.match(r"[0-9]{6}$", e):  #
                            routes.append(e)  # add route to routes array
                            lookfor2route = True
                        lookforroute = False
                    if e == "Route #":  # set trap to catch route # on the next line
                        lookforroute = True
                    if lookfor2level == True:  # intercept the second level
                        if re.match(r"[0-9]{2}$", e):
                            level.append(e)
                        lookfor2level = False
                    if lookforlevel == True:  # intercept the level
                        if re.match(r"[0-9]{2}$", e):
                            level.append(e)
                            lookfor2level = True  # set trap to catch the second level next line
                        lookforlevel = False
                    if e == "Lvl":  # set trap to catch Lvl on the next line
                        lookforlevel = True
                    if eid != "" and new_page == False:
                        if re.match(r"[0-9]{8}", e):  # find the underscore dash string
                            eid_count += 1
                        if re.match(r"xxx\-xx\-[0-9]{4}", e):
                            eid_count += 1
                        if re.match(r"XXX\-XX\-[0-9]{4}", e):
                            eid_count += 1
                        if e == "___-___-____":
                            eid_count += 1
                        # solve for rose problem: time object is fused to emp id object - just increment the eid counter
                        if re.match(r"\s[0-9]{2}\.[0-9]{10}", e) \
                                or re.match(r"__.__[0-9]{8}", e) \
                                or re.match(r"__._____-___-____", e):
                            eid_count += 1
                            rose_rpt.append(lastname)
                    # solve for carroll problem/ unprocessed rings do not have underscore slash counterparts
                    if e == "Un-Processed Rings":  # after unprocessed rings label, add no new rings to daily array
                        unprocessedrings = eid
                    if re.match(r"[0-9]{8}", e):  # find the emp id / it is the first 8 digit number on the page
                        if eid_count == 0:
                            eid = e
                            if gen_error_report == "on":
                                input = "Employee ID: {}\n".format(e)
                                kbpc_rpt.write(input)
                            lookforname = True
                            if eid != unprocessedrings:  # set unprocessedrings and new_page variables
                                unprocessedrings = ""
                                new_page = False
                            else:
                                new_page = True
                                eid_count += 1  # increment the eid counter to stop new eid from being set
                                if gen_error_report == "on": kbpc_rpt.write("NEW PAGE!!!\n")
                except:
                    failed.append(lastname)
                    input = "READING FAILURE: {}\n".format(e)
                    kbpc_rpt.write(input)
        if gen_error_report == "on":  # write to error report
            input = "Station: {}\n".format(station)
            kbpc_rpt.write(input)
            input = "Pay Period: {}\n".format(yyppwk)
            kbpc_rpt.write(input)  # show the pay period
            if len(jobs) > 0:
                input = "Jobs: {}\n".format(jobs)
                kbpc_rpt.write(input)
            if len(routes) > 0:
                input = "Routes: {}\n".format(routes)
                kbpc_rpt.write(input)
            if len(level) > 0:
                input = "Levels: {}\n".format(level)
                kbpc_rpt.write(input)
        # define csv writer parameters
        csv.register_dialect('myDialect',
                             delimiter=',',
                             quotechar="'",
                             skipinitialspace=True,
                             lineterminator=",\r"
                             )
        # write to csv file
        prime_info = [yyppwk.replace("-", ""), '"{}"'.format("000000"), '"{}"'.format(station),
                      '"{}"'.format("0000"), '"{}"'.format(eid), '"{}"'.format(lastname), '"{}"'.format(fi[:1]),
                      '"_"', '"010/0000"', '"N"', '"N"', '"N"', '"0"', '"0"', '"0"', '"0"', '"0"', '"0"']
        if len(jobs) > 0:
            for i in range(len(jobs)):
                base_line = [base_temp[i], '"{}"'.format(jobs[i].replace("-", "").strip()), '"0000"', '"7220-10"',
                             '"Q0"', '"{}"'.format(level[i]), '"N"', '"{}"'.format(routes[i]), '""', '"0000000"',
                             '"0000000"', '"0"', '"0"', '"0"', '"N"', '"N"', '"N"', '"MDT"', '"N"']
                whole_line = prime_info + base_line
                with open(new_file_path, 'a') as writeFile:
                    writer = csv.writer(writeFile, dialect='myDialect')
                    writer.writerow(whole_line)
        if len(foundday_holder) > 0:
            # solve for nguyen problem / day of week occurs prior to "employee id" label
            found_days = found_days + foundday_holder
            ordered_days = ["Saturday", "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
            for day in days:  # re order days into correct order
                if day not in found_days:
                    ordered_days.remove(day)
            found_days = ordered_days
            foundday_holder = []
            nguyen_rpt.append(lastname)
        if len(found_days) > 0:  # printe out found days
            # reorder the found days to ensure the correct order
            found_days = pdf_converter_reorder_founddays(found_days)
            if gen_error_report == "on":
                input = "Found days: {}\n".format(found_days)
                kbpc_rpt.write(input)
        if gen_error_report == "on":
            input = "proto emp id counter: {}\n".format(eid_count)
            kbpc_rpt.write(input)
        if len(daily_underscoreslash) > 0:  # bind all underscore slash items in one array
            underscore_slash_result = sum(daily_underscoreslash, [])
        if mcgrath_indicator == True and len(underscore_slash_result) > 0:  # solve for mcgrath indicator
            mcgrath_carryover.append(underscore_slash_result[0])  # add underscore slash to carryover
            mcgrath_indicator = False  # reset the indicator
            if gen_error_report == "on":
                input = "MCGRATH CARRYOVER: {}\n".format(mcgrath_carryover)
                kbpc_rpt.write(input)  # printe out a notice.
            del underscore_slash_result[0]  # delete the ophan underscore slash
        count = 0
        for array in daily_array:
            array.append(underscore_slash_result[count])
            try:
                array.append(underscore_slash_result[count + 1])
            except:  # solve for the mcgrath problem
                mcgrath_carryover = array
                mcgrath_indicator = True
                mcgrath_rpt.append(lastname)
                if gen_error_report == "on": kbpc_rpt.write("MCGRATH ERROR DETECTED!!!\n")
            # if mcgrath_indicator == False:
            count += 2
        if mcgrath_carryover in daily_array:  # if there is a carryover, remove the daily array item from the list
            daily_array.remove(mcgrath_carryover)
        if mcgrath_indicator == False and mcgrath_carryover != "":  # if there is a carryover to be added
            daily_array.insert(0, mcgrath_carryover)  # put the carryover at the front of the daily array
            mcgrath_carryover = ""  # reset the carryover
            eid_count += 1  # increment the emp id counter
        # set up array for each day in the week
        csv_sat = []
        csv_sun = []
        csv_mon = []
        csv_tue = []
        csv_wed = []
        csv_thr = []
        csv_fri = []
        csv_output = [csv_sat, csv_sun, csv_mon, csv_tue, csv_wed, csv_thr, csv_fri]
        # reorder the found days to ensure the correct order
        found_days = pdf_converter_reorder_founddays(found_days)
        # fix problem with miscounted base times
        high_array = []
        for bt in base_time:
            high_array.append(bt[0])
        if len(high_array) > 0:
            high_num = max(high_array)
            comp_array = []
            for i in range(high_num + 1):
                comp_array.append(i)
            del_array = []
            for num in comp_array:
                if num in high_array:
                    del_array.append(num)
            error_array = comp_array
            error_array = [x for x in error_array if x not in del_array]
            error_array.reverse()
            if len(error_array) > 0:
                for error_num in error_array:
                    for bt in base_time:
                        if bt[0] > error_num:
                            bt[0] = bt[0] - 1
        # load the multi array with array for each day
        for i in range(7):
            for bt in base_time:
                if found_days[bt[0]] == days[i]:
                    csv_output[i].append(bt)
            for da in daily_array:
                if da[2] == pp_days[i].strftime("%m/%d"):
                    csv_output[i].append(da)
        for co in csv_output:  # for each time in the array, printe a line
            for array in co:
                if gen_error_report == "on":
                    input = "{}\n".format(str(array))
                    kbpc_rpt.write(input)
                # put the data into the csv file
                if len(array) == 2:  # if the line comes from base/time data
                    add_this = [found_days[int(array[0])], '"_0-00"', '"{}"'.format(array[1])]
                    whole_line = prime_info + add_this
                    with open(new_file_path, 'a') as writeFile:  # add the line to the csv file
                        writer = csv.writer(writeFile, dialect='myDialect')
                        writer.writerow(whole_line)
                if len(array) == 10:  # if the line comes from daily array
                    if array[9] != "__/__":
                        end_notes = "(W)Ring Deleted From PC"
                    else:
                        end_notes = ""
                    add_this = ["000-00", '"{}"'.format(array[1]),
                                '"{}"'.format(pp_days[daily_array_days.index(array[2])].strftime("%d-%b-%y").upper()),
                                '"{}"'.format(array[3].strip()), '"{}"'.format(array[5]), '"{}"'.format(array[6]),
                                '"{}"'.format(array[7]), '""', '""', '""', '"0"', '""', '""', '"0"',
                                '"{}"'.format(end_notes)]
                    whole_line = prime_info + add_this
                    with open(new_file_path, 'a') as writeFile:  # add the line to the csv file
                        writer = csv.writer(writeFile, dialect='myDialect')
                        writer.writerow(whole_line)
        # Handle Carroll problems
        if mcgrath_indicator == False:
            if eid_count == 1:  # handle widows
                eid_count = 0
                if gen_error_report == "on":
                    input = "WIDOW HANDLING: Carroll Mod emp id counter: {}\n".format(eid_count)
                    kbpc_rpt.write(input)
            elif eid_count % 2 != 0:  # handle eid counts where there has been a cut off
                eid_count += 1
                if gen_error_report == "on":
                    input = "CUT OFF CONTROL: Carroll Mod emp id counter: {}\n".format(eid_count)
                    kbpc_rpt.write(input)
        else:
            eid_count -= 1
        eid_count = eid_count - (unprocessed_counter * 2)

        if unprocessed_counter > 0:
            if gen_error_report == "on":
                input = "Unprocessed Rings: {}\n".format(unprocessed_counter)
                kbpc_rpt.write(input)
            if len(daily_array) == (eid_count) / 2:
                pass
            # Solve for Unruh error / when a underscore dash is missing after unprocessed rings
            elif len(daily_array) == max((eid_count + 2) / 2, 0):
                if gen_error_report == "on":
                    input = "Unruh Mod emp id counter: {}\n".format(eid_count + 2)
                    kbpc_rpt.write(input)
                    kbpc_rpt.write("UNRUH PROBLEM DETECTED!!!")
                unruh_rpt.append(lastname)
            else:
                if gen_error_report == "on": kbpc_rpt.write(
                    "FRANKLIN ERROR DETECTED!!! ALERT! (Unprocessed counter)!\n")
                unresolved.append(lastname)
        else:
            if len(daily_array) != max((eid_count) / 2, 0):
                if gen_error_report == "on": kbpc_rpt.write("FRANKLIN ERROR DETECTED!!! ALERT! ALERT!\n")
                unresolved.append(lastname)
        if base_chg + 1 != len(found_days):  # add to basecounter error array
            to_add = (lastname, base_chg, len(found_days))
            if len(found_days) > 0:
                basecounter_error.append(to_add)
        if gen_error_report == "on":
            input = "daily array lenght: {}\n".format(len(daily_array))
            kbpc_rpt.write(input)
        found_days = []  # initialize arrays
        foundday_holder = []
        base_time = []
        eid = ""
        eid_label = False
        # perez_switch = False
        base_counter = 0
        base_chg = 0
        daily_array = []
        daily_underscoreslash = []
        unprocessed_counter = 0
        jobs = []
        routes = []
        level = []
        franklin_array = []
        if gen_error_report == "on":
            input = "emp id counter: {}\n".format(max(eid_count, 0))
            kbpc_rpt.write(input)
        pb["value"] = pb_count  # increment progress bar
        pb_root.update()
        pb_count += 1
    # end loop
    endtime = time.time()
    pb.stop()  # stop and destroy the progress bar
    pb_label.destroy()  # destroy the label for the progress bar
    pb.destroy()
    pb_root.destroy()
    if gen_error_report == "on":
        kbpc_rpt.write("Potential Problem Reports _________________________________________________\n")
        input = "runtime: {} seconds\n".format(round(endtime - starttime, 4))
        kbpc_rpt.write(input)
        kbpc_rpt.write("Franklin Problems: Consecutive MV Desigs \n")
        input = "\t>>> {}\n".format(frank_rpt)
        kbpc_rpt.write(input)
        kbpc_rpt.write("Rodriguez Problem: This is the Franklin Problem X 4. \n")
        input = "\t>>> {}\n".format(rod_rpt)
        kbpc_rpt.write(input)
        kbpc_rpt.write("Rose Problem: The MV Desig and date are on the same line.\n")
        input = "\t>>> {}\n".format(rose_rpt)
        kbpc_rpt.write(input)
        kbpc_rpt.write("Robertson Baseline Problem: The base count is jumping when H/L basetimes "
                       "are put into the basetime array.\n")
        input = "\t>>> {}\n".format(robert_rpt)
        kbpc_rpt.write(input)
        kbpc_rpt.write("Stevens Problem: Basetimes begining with H/L do not show up and are "
                       "not entered into the basetime array.\n")
        input = "\t>>> {}\n".format(stevens_rpt)
        kbpc_rpt.write(input)
        kbpc_rpt.write("Carroll Problem: Unprocessed rings at the end of the page do not contain __/__ or times.'n")
        input = ">>> {}\n".format(carroll_rpt)
        kbpc_rpt.write(input)
        kbpc_rpt.write("Nguyen Problem: Found day appears above the Emp ID.\n")
        input = "\t>>> {}\n".format(nguyen_rpt)
        kbpc_rpt.write(input)
        kbpc_rpt.write("Unruh Problem: Underscore dash cut off in unprecessed rings.\n")
        input = "\t>>> {}\n".format(unruh_rpt)
        kbpc_rpt.write(input)
        kbpc_rpt.write(
            "Salih Problem: Unprocessed rings are missing a timezone, so that unprocessed rings counter is not" \
            " incremented.\n")
        input = "\t>>> {}\n".format(salih_rpt)
        kbpc_rpt.write(input)
        kbpc_rpt.write("McGrath Problem: \n")
        input = " \t>>> {}\n".format(mcgrath_rpt)
        kbpc_rpt.write(input)
        input = "Unresolved: {}\n".format(unresolved)
        kbpc_rpt.write(input)
        input = "Base Counter Error: {}\n".format(basecounter_error)
        kbpc_rpt.write(input)
    if len(failed)>0: # create messagebox to show any errors
        failed_daily = ""
        for f in failed:
            failed_daily = failed_daily + " \n " + f
        messagebox.showerror("Klusterbox PDF Converter", "Errors have occured for the following carriers {}."
                             .format(failed_daily))
    # create messagebox for completion
    messagebox.showinfo("Klusterbox PDF Converter", "The PDF Convertion is complete. "
                                                    "The file name is {}. ".format(short_file_name))


def informalc_grvchange(frame, passed_result, old_num, new_num):
    l_passed_result = [list(x) for x in passed_result]  # chg tuple of tuples to list of lists
    ok = messagebox.askokcancel("Grievance Number Change", "This will change the grievance number from {} to {} in all "
                                                           "records. Are you sure you want to proceed?".format(old_num,
                                                                                                               new_num.get()))
    if ok == True:
        if new_num.get().strip() == "":
            messagebox.showerror("Invalid Data Entry", "You must enter a grievance number")
            return "fail"
        if new_num.get().isalnum() == False:
            messagebox.showerror("Invalid Data Entry",
                                 "The grievance number can only contain numbers and letters. No other "
                                 "characters are allowed")
            return "fail"
        if len(new_num.get()) < 8:
            messagebox.showerror("Invalid Data Entry", "The grievance number must be at least eight characters long")
            return "fail"
        if len(new_num.get()) > 16:
            messagebox.showerror("Invalid Data Entry", "The grievance number must not exceed 16 characters in lenght.")
            return "fail"

        sql = "SELECT grv_no FROM informalc_grv WHERE grv_no = '%s'" % new_num.get().lower()
        result = inquire(sql)
        if result:
            messagebox.showerror("Grievance Number Error", "This number is already being used for another grievance.")
            return "fail"

        sql = "UPDATE informalc_grv SET grv_no = '%s' WHERE grv_no = '%s'" % (new_num.get().lower(), old_num)
        commit(sql)
        sql = "UPDATE informalc_awards SET grv_no = '%s' WHERE grv_no = '%s'" % (new_num.get().lower(), old_num)
        commit(sql)
        for record in l_passed_result:
            if record[0] == old_num:
                record[0] = new_num.get().lower()
        msg = "The grievance number has been changed."
        informalc_edit(frame, l_passed_result, new_num.get().lower(), msg)


def informalc_edit_apply(frame, grv_no, incident_start, incident_end,date_signed, station, gats_number, docs,
                         description, lvl):
    check = informalc_check_grv_2(incident_start, incident_end, date_signed, gats_number, description)
    if check == "fail":
        return
    dates = [incident_start, incident_end, date_signed]
    in_start = datetime(1, 1, 1)
    in_end = datetime(1, 1, 1)
    d_sign = datetime(1, 1, 1)
    dt_dates = [in_start, in_end, d_sign]
    i = 0
    for date in dates:
        d = date.get().split("/")
        new_date = datetime(int(d[2].lstrip("0")), int(d[0].lstrip("0")), int(d[1].lstrip("0")))
        dt_dates[i] = new_date
        i += 1
    if dt_dates[0] > dt_dates[1]:
        messagebox.showerror("Data Entry Error", "The Incident Start Date can not be later that the Incident End "
                                                 "Date.")
        return
    if dt_dates[0] > dt_dates[2]:
        messagebox.showerror("Data Entry Error", "The Incident Start Date can not be later that the Date Signed.")
        return
    sql = "UPDATE informalc_grv SET indate_start='%s',indate_end='%s',date_signed='%s',station='%s',gats_number='%s'," \
          "docs='%s',description='%s', level='%s' WHERE grv_no='%s'" \
          % (dt_dates[0], dt_dates[1], dt_dates[2], station.get(),gats_number.get().strip(), docs.get(),
             description.get(),lvl.get(),grv_no.get())
    commit(sql)
    messagebox.showerror("Sucessful Update", "Grievance number: {} succesfully updated.".format(grv_no.get()))
    informalc_grvlist(frame)


def informalc_delete(frame, grv_no):
    check = messagebox.askokcancel("Delete Grievance", "Are you sure you want to delete his grievance and all the "
                                                       "data associated with it?")
    if check == False:
        return
    else:
        sql = "DELETE FROM informalc_grv WHERE grv_no='%s'" % grv_no.get()
        commit(sql)
        informalc_grvlist(frame)


def informalc_edit(frame, result, grv_num, msg):
    wd = front_window(frame)
    Label(wd[3], text="Informal C: Edit Grievance", font="bold").grid(row=0, columnspan=2, sticky="w")
    Label(wd[3], text="").grid(row=1)
    Label(wd[3], text="Grievance Number: ").grid(row=2, column=0, sticky="w")
    grv_no = StringVar(wd[0])
    Entry(wd[3], textvariable=grv_no, justify='right').grid(row=2, column=1, sticky="w")
    Button(wd[3], width=9, text="update", command=lambda:
    informalc_grvchange(wd[0], result, grv_num, grv_no)).grid(row=3, column=1, sticky="e")
    grv_no.set(grv_num)
    Label(wd[3], text="Incident Date").grid(row=4, column=0, sticky="w")
    Label(wd[3], text="  Start (mm/dd/yyyy): ").grid(row=5, column=0, sticky="w")
    incident_start = StringVar(wd[0])
    Entry(wd[3], textvariable=incident_start, justify='right').grid(row=5, column=1, sticky="w")
    Label(wd[3], text="  End (mm/dd/yyyy): ").grid(row=6, column=0, sticky="w")
    incident_end = StringVar(wd[0])
    Entry(wd[3], textvariable=incident_end, justify='right').grid(row=6, column=1, sticky="w")
    Label(wd[3], text="Date Signed (mm/dd/yyyy): ").grid(row=7, column=0, sticky="w")
    date_signed = StringVar(wd[0])
    Entry(wd[3], textvariable=date_signed, justify='right').grid(row=7, column=1, sticky="w")

    Label(wd[3], text="Settlement Level: ").grid(row=8, column=0, sticky="w")  # select settlement level
    lvl = StringVar(wd[0])
    lvl_options = ("informal a", "formal a", "step b", "pre arb", "arbitration")
    lvl_om = OptionMenu(wd[3], lvl, *lvl_options)
    lvl_om.config(width=13)
    lvl_om.grid(row=8, column=1)

    Label(wd[3], text="Station: ").grid(row=9, column=0, sticky="w")  # select a station
    station = StringVar(wd[0])
    station_options = list_of_stations
    if "out of station" in station_options:
        station_options.remove("out of station")
    station_om = OptionMenu(wd[3], station, *station_options)
    station_om.config(width=40)
    station_om.grid(row=10, column=0, columnspan=2, sticky="e")
    Label(wd[3], text="GATS Number: ").grid(row=11, column=0, sticky="w")
    gats_number = StringVar(wd[0])
    Entry(wd[3], textvariable=gats_number, justify='right').grid(row=11, column=1, sticky="w")
    Label(wd[3], text="Documentation: ").grid(row=12, column=0, sticky="w")
    docs = StringVar(wd[0])
    doc_options = ("moot","no","partial","yes","incomplete","verified")
    docs_om = OptionMenu(wd[3], docs, *doc_options)
    docs_om.config(width=13)
    docs_om.grid(row=12, column=1)
    Label(wd[3], text="Description: ").grid(row=16, column=0, sticky="w")
    description = StringVar(wd[0])
    Entry(wd[3], textvariable=description, width=47, justify='right')\
        .grid(row=17, column=0, sticky="e", columnspan=2)
    Label(wd[3], text="").grid(row=18, column=0)
    sql = "SELECT * FROM informalc_grv WHERE grv_no='%s'" % grv_num
    search = inquire(sql)
    if search:
        in_start = datetime.strptime(search[0][1], '%Y-%m-%d %H:%M:%S')
        in_end = datetime.strptime(search[0][2], '%Y-%m-%d %H:%M:%S')
        sign_date = datetime.strptime(search[0][3], '%Y-%m-%d %H:%M:%S')
        incident_start.set(in_start.strftime("%m/%d/%Y"))
        incident_end.set(in_end.strftime("%m/%d/%Y"))
        date_signed.set(sign_date.strftime("%m/%d/%Y"))
        station.set(search[0][4])
        gats_number.set(search[0][5])
        docs.set(search[0][6])
        description.set(search[0][7])
        if search[0][8] == None:
            lvl.set("unknown")
        else:
            lvl.set(search[0][8])
    Label(wd[3], text=" ").grid(row=20)
    Label(wd[3], text="Delete Grievance").grid(row=21, column=0, sticky="w")
    Button(wd[3], text="Delete", width=9, command=lambda: informalc_delete(wd[0], grv_no)).grid(row=21, column=1,
                                                                                                sticky="e")
    Label(wd[3], text=" ").grid(row=22)
    Label(wd[3], text=msg, fg="red", anchor="w").grid(row=23, column=0, columnspan=5, sticky="w")
    Button(wd[4], text="Go Back", width=20, command=lambda: informalc_grvlist_result(wd[0], result)).grid(row=0,
                                                                                                          column=0)
    Button(wd[4], text="Enter", width=20, command=lambda: informalc_edit_apply(wd[0], grv_no, incident_start,
        incident_end, date_signed, station, gats_number, docs, description, lvl)).grid(row=0, column=1)
    rear_window(wd)


def informalc_check_grv(grv_no, incident_start, incident_end, date_signed, station, gats_number, description):
    if station.get() == "Select a Station":
        messagebox.showerror("Invalid Data Entry", "You must select a station.")
        return "fail"
    if grv_no.get().strip() == "":
        messagebox.showerror("Invalid Data Entry", "You must enter a grievance number")
        return "fail"
    if re.search('[^1234567890abcdefghijklmnopqrstuvwxyz:ABCDEFGHIJKLMNOPQRSTUVWXYZ,]', grv_no.get()):
        messagebox.showerror("Invalid Data Entry",
                                 "The grievance number can only contain numbers and letters. No other "
                                 "characters are allowed")
        return "fail"
    if len(grv_no.get()) < 8:
        messagebox.showerror("Invalid Data Entry", "The grievance number must be at least eight characters long")
        return "fail"
    if len(grv_no.get()) > 20:
        messagebox.showerror("Invalid Data Entry", "The grievance number must not exceed 20 characters in length.")
        return "fail"
    check = informalc_check_grv_2(incident_start, incident_end, date_signed, gats_number, description)
    return check


def informalc_check_grv_2(incident_start, incident_end,
                          date_signed, gats_number, description):
    dates = [incident_start, incident_end, date_signed]
    date_ids = ("starting incident date", "ending incident date", "date signed")
    i = 0
    for date in dates:
        d = date.get().split("/")
        if len(d) != 3:
            messagebox.showerror("Invalid Data Entry",
                                 "The date for the {} is not properly formatted.".format(date_ids[i]))
            return "fail"
        for num in d:
            if num.isnumeric() == False:
                messagebox.showerror("Invalid Data Entry", "The month, day and year for the {} "
                                                           "must be numeric.".format(date_ids[i]))
                return "fail"
        if len(d[0]) > 2:
            messagebox.showerror("Invalid Data Entry", "The month for the {} must be no more than two digits"
                                                       " long.".format(date_ids[i]))
            return "fail"
        if len(d[1]) > 2:
            messagebox.showerror("Invalid Data Entry", "The day for the {} must be no more than two digits"
                                                       " long.".format(date_ids[i]))
            return "fail"
        if len(d[2]) != 4:
            messagebox.showerror("Invalid Data Entry", "The year for the {} must be four digits long."
                                 .format(date_ids[i]))
            return "fail"
        try:
            date = datetime(int(d[2]), int(d[0]), int(d[1]))
            valid_date = True
        except ValueError:
            valid_date = False
        if valid_date == False:
            messagebox.showerror("Invalid Data Entry", "The date entered for {} is not a valid date."
                                 .format(date_ids[i]))
            return "fail"
        i += 1
    if len(gats_number.get()) > 50:
        messagebox.showerror("Invalid Data Entry", "The GATS number is limited to no more than 20 characters. ")
        return "fail"
    if gats_number.get().strip() != "":
        if all(x.isalnum() or x.isspace() for x in gats_number.get()) == False:
            messagebox.showerror("Invalid Data Entry", "The GATS number can only contain letters and numbers. No "
                                                       "special characters are allowed.")
            return "fail"
    if description.get().strip() != "":
        if all(x.isalnum() or x.isspace() for x in description.get()) == False:
            messagebox.showerror("Invalid Data Entry", "The Description can only contain letters and numbers. No "
                                                       "special characters are allowed.")
            return "fail"
        if len(description.get()) > 40:
            messagebox.showerror("Invalid Data Entry", "The Description is limited to no more than 40 characters. ")
            return "fail"
    return "pass"


def informalc_new_apply(frame, grv_no, incident_start, incident_end, date_signed, station, gats_number, docs,
                        description, lvl):
    check = informalc_check_grv(grv_no, incident_start, incident_end, date_signed, station, gats_number, description)
    if check == "pass":
        dates = [incident_start, incident_end, date_signed]
        in_start = datetime(1, 1, 1)
        in_end = datetime(1, 1, 1)
        d_sign = datetime(1, 1, 1)
        dt_dates = [in_start, in_end, d_sign]
        i = 0
        for date in dates:
            d = date.get().split("/")
            new_date = datetime(int(d[2].lstrip("0")), int(d[0].lstrip("0")), int(d[1].lstrip("0")))
            dt_dates[i] = new_date
            i += 1
        if dt_dates[0] > dt_dates[1]:
            messagebox.showerror("Data Entry Error", "The Incident Start Date can not be later that the Incident End "
                                                     "Date.")
            return
        if dt_dates[0] > dt_dates[2]:
            messagebox.showerror("Data Entry Error", "The Incident Start Date can not be later that the Date Signed.")
            return
        sql = "SELECT grv_no FROM informalc_grv"
        results = inquire(sql)
        existing_grv = []
        for result in results:
            for grv in result:
                existing_grv.append(grv)
        if grv_no.get() in existing_grv:
            messagebox.showerror("Data Entry Error",
                                 "The Grievance Number {} is already present in the database. You can not "
                                 "create a duplicate.".format(grv_no.get()))
            return
        sql = "INSERT INTO informalc_grv (grv_no, indate_start, indate_end, date_signed, station, gats_number, docs," \
              "description, level) VALUES('%s','%s','%s','%s','%s','%s','%s','%s','%s')" \
              % (grv_no.get().lower(), dt_dates[0],dt_dates[1], dt_dates[2], station.get(),gats_number.get().strip(),
                 docs.get(),description.get(), lvl.get())
        commit(sql)
        msg = "Grievance Settlement Added: #{}.".format(grv_no.get().lower())
        informalc_new(frame, msg)


def informalc_gen_clist(start, end, station):
    end += timedelta(weeks=52)
    sql = "SELECT * FROM carriers WHERE effective_date<='%s'and station='%s' " \
          "ORDER BY carrier_name, effective_date DESC" % (end, station)
    result = inquire(sql)
    unique_carriers = []  # create non repeating list of otdl carriers
    for name in result:
        if name[1] not in unique_carriers:
            unique_carriers.append(name[1])
    carrier_list = []
    for name in unique_carriers:
        sql = "SELECT effective_date,carrier_name,station FROM carriers WHERE carrier_name='%s' " \
              "ORDER BY effective_date DESC" % name
        after_start = []  # array for records after start date
        before_start = []  # array for records before start date
        added = False
        result = inquire(sql)
        for rec in result:
            if rec[0] >= str(start):
                after_start.append(rec)
            if rec[0] < str(start):
                before_start.append(rec)
        for rec in after_start:
            if added == False and rec[2] == station:
                carrier_list.append(rec[1])
                added = True
        if added == False and len(before_start) > 0:
            if before_start[0][2] == station:
                carrier_list.append(rec[1])
    return carrier_list


def informalc_addnames(grv_no, c_list, listbox):
    for index in listbox:
        sql = "INSERT INTO informalc_awards (grv_no,carrier_name,hours,rate,amount) VALUES('%s','%s','%s','%s','%s')" \
              % (grv_no, c_list[int(index)], '', '', '')
        commit(sql)


def informalc_root(passed_result, grv_no):
    global informalc_newroot  # initialize the global
    new_root = Tk()
    informalc_newroot = new_root  # set the global
    if sys.platform == "win32":
        try:
            new_root.iconbitmap(r'kb_sub/kb_images/kb_icon2.ico')
        except:
            pass
    if sys.platform == "linux":
        try:
            img = PhotoImage(file='kb_sub/kb_images/kb_icon2.gif')
            new_root.tk.call('wm', 'iconphoto', new_root._w, img)
        except:
            pass
    new_root.title("KLUSTERBOX")
    x_position = root.winfo_x() + 450
    y_position = root.winfo_y() - 25
    new_root.geometry("%dx%d+%d+%d" % (240, 600, x_position, y_position))
    n_F = Frame(new_root)
    n_F.pack()
    n_buttons = Canvas(n_F)  # button bar
    n_buttons.pack(fill=BOTH, side=BOTTOM)
    Label(n_F, text="Add Carriers", font="bold").pack(anchor="w")
    Label(n_F, text="").pack()
    scrollbar = Scrollbar(n_F, orient=VERTICAL)
    listbox = Listbox(n_F, selectmode="multiple", yscrollcommand=scrollbar.set)
    listbox.config(height=100, width=50)
    sql = "SELECT indate_start,indate_end,station FROM informalc_grv WHERE grv_no='%s'" % grv_no
    results = inquire(sql)
    if results:
        start = results[0][0]
        end = results[0][1]
        station = results[0][2]
    start = dt_converter(start)
    end = dt_converter(end)
    c_list = informalc_gen_clist(start, end, station)
    for name in c_list:
        listbox.insert(END, name)
    scrollbar.config(command=listbox.yview)
    scrollbar.pack(side=RIGHT, fill=Y)
    listbox.pack(side=LEFT, expand=1)
    Button(n_buttons, text="Add Carrier", width=10,
           command=lambda: (informalc_addnames(grv_no, c_list, listbox.curselection()),
                            informalc_addaward2(informalc_addframe, passed_result, grv_no))) \
        .pack(side=LEFT, anchor="w")
    Button(n_buttons, text="Clear", width=10,
           command=lambda: (informalc_newroot.destroy(), informalc_root(passed_result, grv_no))) \
        .pack(side=LEFT, anchor="w")
    Button(n_buttons, text="Close", width=10,
           command=lambda: (new_root.destroy())).pack(side=LEFT, anchor="w")


def informalc_deletename(frame, passed_result, grv_no, id):
    sql = "DELETE FROM informalc_awards WHERE rowid='%s'" % (id)
    commit(sql)
    informalc_addaward2(frame, passed_result, grv_no)


def informalc_apply_addaward(frame, buttons, passed_result, grv_no, var_id, var_name, var_hours, var_rate, var_amount):
    pb_label = Label(buttons, text="Updating Changes: ")  # make label for progress bar
    pb_label.grid(row=0, column=2)
    pb = ttk.Progressbar(buttons, length=200, mode="determinate")  # create progress bar
    pb.grid(row=0, column=3)
    pb["maximum"] = len(var_id)  # set length of progress bar
    pb.start()
    ii = 0
    for i in range(len(var_id)):
        pb["value"] = ii  # increment progress bar
        id = var_id[i].get()  # simplify variable names
        name = var_name[i].get()
        hours = var_hours[i].get().strip()
        rate = var_rate[i].get().strip()
        amount = var_amount[i].get().strip()
        if hours and amount:
            messagebox.showerror("Data Input Error", "Input error for {} in row {}. You can not enter both hours and "
                                                     "amount. You can only enter one or another, but not both. Awards can be in the form of "
                                                     "hours at a given rate OR an amount.".format(name, str(i + 1)))
            pb_label.destroy()  # destroy the label for the progress bar
            pb.destroy()  # destroy the progress bar
            return
        if rate and amount:
            messagebox.showerror("Data Input Error", "Input error for {} in row {}. You can not enter both rate and "
                                                     "amount. You can only enter one or another, but not both. Awards can be in the form of "
                                                     "hours at a given rate OR an amount.".format(name, str(i + 1)))
            pb_label.destroy()  # destroy the label for the progress bar
            pb.destroy()  # destroy the progress bar
            return
        if hours and not rate:
            messagebox.showerror("Data Input Error", "Input error for {} in row {}. Hours must be a accompanied by a "
                                                     "rate.".format(name, str(i + 1)))
            pb_label.destroy()  # destroy the label for the progress bar
            pb.destroy()  # destroy the progress bar
            return
        if rate and not hours:
            messagebox.showerror("Data Input Error", "Input error for {} in row {}. Rate must be a accompanied by a "
                                                     "hours.".format(name, str(i + 1)))
            pb_label.destroy()  # destroy the label for the progress bar
            pb.destroy()  # destroy the progress bar
            return
        if hours and isfloat(hours) == False:
            messagebox.showerror("Data Input Error", "Input error for {} in row {}. Hours must be a number."
                                 .format(name, str(i + 1)))
            pb_label.destroy()  # destroy the label for the progress bar
            pb.destroy()  # destroy the progress bar
            return
        if hours and '.' in hours:
            s_hrs = hours.split(".")
            if len(s_hrs[1]) > 2:
                messagebox.showerror("Data Input Error", "Input error for {} in row {}. Hours must have no "
                                                         "more than 2 decimal places.".format(name, str(i + 1)))
                pb_label.destroy()  # destroy the label for the progress bar
                pb.destroy()  # destroy the progress bar
                return
        if rate and amount:
            messagebox.showerror("Data Input Error", "Input error for {} in row {}. You can not enter both rate and "
                                                     "amount. You can only enter one or another, but not both. Awards can be in the form of "
                                                     "hours at a given rate OR an amount.".format(name, str(i + 1)))
            pb_label.destroy()  # destroy the label for the progress bar
            pb.destroy()  # destroy the progress bar
            return
        if rate and amount:
            messagebox.showerror("Data Input Error", "Input error for {} in row {}. You can not enter both rate and "
                                                     "amount. You can only enter one or another, but not both. Awards can be in the form of "
                                                     "hours at a given rate OR an amount.".format(name, str(i + 1)))
            pb_label.destroy()  # destroy the label for the progress bar
            pb.destroy()  # destroy the progress bar
            return
        if rate and isfloat(rate) == False:
            messagebox.showerror("Data Input Error", "Input error for {} in row {}. Rates must be a number."
                                 .format(name, str(i + 1)))
            pb_label.destroy()  # destroy the label for the progress bar
            pb.destroy()  # destroy the progress bar
            return
        if rate and '.' in rate:
            s_rate = rate.split(".")
            if len(s_rate[1]) > 2:
                messagebox.showerror("Data Input Error", "Input error for {} in row {}. Rates must have no "
                                                         "more than 2 decimal places.".format(name, str(i + 1)))
                pb_label.destroy()  # destroy the label for the progress bar
                pb.destroy()  # destroy the progress bar
                return
        if rate and float(rate) > 10:
            messagebox.showerror("Data Input Error", "Input error for {} in row {}. Values greater than 10 are not "
                                                     "accepted. \n"
                                                     "Note the following rates would be expressed as: \n "
                                                     "additional %50         .50 or just .5 \n"
                                                     "straight time rate     1.00 or just 1 \n"
                                                     "overtime rate          1.50 or 1.5 \n"
                                                     "penalty rate           2.00 or just 2".format(name, str(i + 1)))
            pb_label.destroy()  # destroy the label for the progress bar
            pb.destroy()  # destroy the progress bar
            return
        if amount and isfloat(amount) == False:
            messagebox.showerror("Data Input Error", "Input error for {} in row {}. Amounts can only be expressed as "
                                                     "numbers. No special characters, such as $ are allowed.".format(
                name, str(i + 1)))
            pb_label.destroy()  # destroy the label for the progress bar
            pb.destroy()  # destroy the progress bar
            return
        if amount and '.' in amount:
            s_amt = amount.split(".")
            if len(s_amt[1]) > 2:
                messagebox.showerror("Data Input Error", "Input error for {} in row {}. Amounts must have no "
                                                         "more than 2 decimal places.".format(name, str(i + 1)))
                pb_label.destroy()  # destroy the label for the progress bar
                pb.destroy()  # destroy the progress bar
                return
        sql = "UPDATE informalc_awards SET hours='%s',rate='%s',amount='%s' WHERE rowid='%s'" % (
        hours, rate, amount, id)
        commit(sql)
        buttons.update()  # update the progress bar
        ii += 1
    pb.stop()  # stop and destroy the progress bar
    pb_label.destroy()  # destroy the label for the progress bar
    pb.destroy()
    informalc_addaward2(frame, passed_result, grv_no)


def informalc_addaward2(frame, passed_result, grv_no):
    global informalc_addframe
    wd = front_window(frame)
    informalc_addframe = wd[0]
    Label(wd[3], text="Add/Update Settlement Awards", font="bold").grid(row=0, column=0, sticky="w", columnspan=4)
    Label(wd[3], text=" ".format(informalc_addframe)).grid(row=1, column=0)
    Label(wd[3], text="   Grievance Number: {}".format(grv_no), fg="blue").grid(row=2, column=0, sticky="w",
                                                                                columnspan=4)
    sql = "SELECT grv_no,rowid,carrier_name,hours,rate,amount FROM informalc_awards WHERE grv_no ='%s' " \
          "ORDER BY carrier_name" % grv_no
    result = inquire(sql)
    # initialize arrays for names
    var_id = []
    var_name = []
    var_hours = []
    var_rate = []
    var_amount = []
    if len(result) == 0:
        Label(wd[3], text="No records in database").grid(row=3)
    else:
        Label(wd[3], text="Carrier", fg="grey", padx=10).grid(row=3, column=0, sticky="w")
        Label(wd[3], text="Hours", fg="grey", padx=10).grid(row=3, column=1, sticky="w")
        Label(wd[3], text="Rate", fg="grey", padx=10).grid(row=3, column=2, sticky="w")
        Label(wd[3], text="Amount", fg="grey", padx=10).grid(row=3, column=3, sticky="w")
        i = 0
        r = 4
        for re in result:
            var_id.append(StringVar(wd[0]))  # add to arrays
            var_name.append(StringVar(wd[0]))
            var_hours.append(StringVar(wd[0]))
            var_rate.append(StringVar(wd[0]))
            var_amount.append(StringVar(wd[0]))
            Label(wd[3], text=re[2], anchor="w", width=16).grid(row=r, column=0, sticky="w",
                                                                padx=10)  # display name widget
            Entry(wd[3], textvariable=var_hours[i], width=8).grid(row=r, column=1, padx=10)  # display hours widget
            Entry(wd[3], textvariable=var_rate[i], width=8).grid(row=r, column=2, padx=10)  # display rate widget
            Entry(wd[3], textvariable=var_amount[i], width=8).grid(row=r, column=3, padx=10)  # display amount widget
            Button(wd[3], text="delete",
                   command=lambda id=re[1]: informalc_deletename(wd[0], passed_result, grv_no, id)) \
                .grid(row=r, column=4, padx=10)  # display the delete button
            var_id[i].set(re[1])  # set the textvariables
            var_name[i].set(re[2])
            var_hours[i].set(re[3])
            var_rate[i].set(re[4])
            var_amount[i].set(re[5])
            r += 1
            i += 1
    Button(wd[4], text="Go Back", width=15, command=lambda: informalc_call_grvlist_result(wd[0], passed_result)) \
        .grid(row=0, column=0)
    Button(wd[4], text="Apply", width=15,
           command=lambda: informalc_apply_addaward(wd[0], wd[4], passed_result, grv_no, var_id, var_name, var_hours,
                                                    var_rate, var_amount)) \
        .grid(row=0, column=1)
    rear_window(wd)


def informalc_call_grvlist_result(frame, passed_result):
    try:
        informalc_newroot.destroy()
    except:
        pass
    informalc_grvlist_result(frame, passed_result)


def informalc_addaward(frame, passed_result, grv_no):
    informalc_root(passed_result, grv_no)
    informalc_addaward2(frame, passed_result, grv_no)


def informalc_rptgrvsum(result):
    if len(result) > 0:
        result = list(result)
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = "infc_grv_list" + "_" + stamp + ".txt"
        if os.path.isdir('kb_sub/infc_grv') == False:
            os.makedirs('kb_sub/infc_grv')
        try:
            report = open('kb_sub/infc_grv/' + filename, "w")
            report.write("Settlement List\n\n")
            i = 1
            for sett in result:
                sett = list(sett)# correct for legacy problem of NULL Settlement Levels
                if sett[8] == None:
                    sett[8] = "unknown"
                sql = "SELECT * FROM informalc_awards WHERE grv_no='%s'" % sett[0]
                query = inquire(sql)
                num_space = 3 - (len(str(i)))  # number of spaces for number
                awardxhour = 0
                awardxamt = 0
                for rec in query:
                    hour = 0.0
                    rate = 0.0
                    amt = 0
                    if rec[2]: hour = float(rec[2])
                    if rec[3]: rate = float(rec[3])
                    if rec[4]: amt = float(rec[4])
                    if hour and rate:
                        awardxhour = awardxhour + (hour * rate)
                    if amt:
                        awardxamt = awardxamt + amt
                space = " "
                space = space + (num_space * " ")
                if i > 99:
                    report.write(str(i) + "\n" + "    Grievance Number:   " + sett[0] + "\n")
                else:
                    report.write(str(i) + space + "Grievance Number:   " + sett[0] + "\n")
                start = dt_converter(sett[1]).strftime("%m/%d/%Y")
                end = dt_converter(sett[2]).strftime("%m/%d/%Y")
                sign = dt_converter(sett[3]).strftime("%m/%d/%Y")
                report.write("    Dates of Violation: " + start + " - " + end + "\n")
                report.write("    Signing Date:       " + sign + "\n")
                report.write("    Settlement Level    " + sett[8] + "\n")
                report.write("    Station:            " + sett[4] + "\n")
                report.write("    GATS Number:        " + sett[5] + "\n")
                report.write("    Documentation:      " + sett[6] + "\n")
                report.write("    Description:        " + sett[7] + "\n\n")
                report.write("    Carrier Name                Hours      Rate   Adjusted     Amount\n")
                report.write("    -----------------------------------------------------------------\n")
                if len(query) == 0:
                    report.write("         No awards recorded for this settlement.\n")
                c = 1
                for rec in query:
                    if rec[2]:
                        hours = "{0:.2f}".format(float(rec[2]))
                    else:
                        hours = "---"
                    if rec[3]:
                        rate = "{0:.2f}".format(float(rec[3]))
                    else:
                        rate = "---"
                    if rec[2] and rec[3]:
                        adj = "{0:.2f}".format(float(rec[2]) * float(rec[3]))
                    else:
                        adj = "---"
                    if rec[4]:
                        amt = "{0:.2f}".format(float(rec[4]))
                    else:
                        amt = "---"
                    report.write(
                        '    {:<5}{:<22}{:>6}{:>10}{:>10}{:>12}\n'.format(str(c), rec[1], hours, rate, adj, amt))
                    c += 1
                report.write("    -----------------------------------------------------------------\n")
                report.write("         {:<38}{:>10}\n".format("Awards adjusted to straight time", "{0:.2f}"
                                                              .format(float(awardxhour))))
                report.write("         {:<38}{:>22}\n".format("Awards as flat dollar amount", "{0:.2f}"
                                                              .format(float(awardxamt))))
                report.write("\n\n\n")
                i += 1
            report.close()
            if sys.platform == "win32":
                os.startfile('kb_sub\\infc_grv\\' + filename)
            if sys.platform == "linux":
                subprocess.call(["xdg-open", 'kb_sub/infc_grv/' + filename])
            if sys.platform == "darwin":
                subprocess.call(["open", 'kb_sub/infc_grv/' + filename])
        except:
            messagebox.showerror("Report Generator", "The report was not generated.")


def informalc_bycarriers(result):
    unique_carrier = informalc_uniquecarrier(result)
    unique_grv = []  # get a list of all grv numbers in search range
    for grv in result:
        if grv[0] not in unique_grv:
            unique_grv.append(grv[0])  # put these in "unique_grv"
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = "infc_grv_list" + "_" + stamp + ".txt"
    if os.path.isdir('kb_sub/infc_grv') == False:
        os.makedirs('kb_sub/infc_grv')
    try:
        report = open('kb_sub/infc_grv/' + filename, "w")
        report.write("Settlement Report By Carriers\n\n")
        for name in unique_carrier:
            report.write("{:<30}\n\n".format(name))
            report.write("        Grievance Number    Hours    Rate    Adjusted      Amount       docs       level\n")
            report.write("    ------------------------------------------------------------------------------------\n")
            results = []
            for ug in unique_grv:  # do search for each grievance in list of unique grievances
                sql = "SELECT informalc_awards.grv_no, informalc_awards.hours, informalc_awards.rate, " \
                      "informalc_awards.amount, informalc_grv.docs, informalc_grv.level FROM informalc_awards, informalc_grv " \
                      "WHERE informalc_awards.grv_no = informalc_grv.grv_no and informalc_awards.carrier_name='%s'" \
                      "and informalc_awards.grv_no = '%s' " \
                      "ORDER BY informalc_grv.date_signed" % (name, ug)
                query = inquire(sql)
                if query:
                    for q in query:
                        q = list(q)
                        results.append(q)
            if len(results) == 0:
                report.write("    There are no awards on record for this carrier.\n")
            total_adj = 0
            total_amt = 0
            i = 1
            for r in results:
                if r[1]:
                    hours = "{0:.2f}".format(float(r[1]))
                else:
                    hours = "---"
                if r[2]:
                    rate = "{0:.2f}".format(float(r[2]))
                else:
                    rate = "---"
                if r[1] and r[2]:
                    adj = "{0:.2f}".format(float(r[1]) * float(r[2]))
                    total_adj = total_adj + (float(r[1]) * float(r[2]))
                else:
                    adj = "---"
                if r[3]:
                    amt = "{0:.2f}".format(float(r[3]))
                    total_amt = total_amt + float(r[3])
                else:
                    amt = "---"
                if r[5] == None or r[5] == "unknown":
                    r[5] = "---"
                report.write("    {:<4}{:<17}{:>8}{:>8}{:>12}{:>12}{:>11}{:>12}\n"
                             .format(str(i), r[0], hours, rate, adj, amt, r[4],r[5]))
                i += 1
            report.write("    ------------------------------------------------------------------------------------\n")
            t_adj = "{0:.2f}".format(float(total_adj))
            t_amt = "{0:.2f}".format(float(total_amt))
            report.write("        {:<34}{:>11}\n".format("Total hours as straight time", t_adj))
            report.write("        {:<34}{:>23}\n".format("Total as flat dollar amount", t_amt))
            report.write("\n\n\n")
        report.close()
        if sys.platform == "win32":
            os.startfile('kb_sub\\infc_grv\\' + filename)
        if sys.platform == "linux":
            subprocess.call(["xdg-open", 'kb_sub/infc_grv/' + filename])
        if sys.platform == "darwin":
            subprocess.call(["open", 'kb_sub/infc_grv/' + filename])
    except:
        messagebox.showerror("Report Generator", "The report was not generated.")


def informalc_apply_bycarrier(result, names, cursor):
    if len(cursor) == 0:
        return
    unique_grv = []  # get a list of all grv numbers in search range
    for grv in result:
        if grv[0] not in unique_grv:
            unique_grv.append(grv[0])  # put these in "unique_grv"
    name = names[cursor[0]]
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = "infc_grv_list" + "_" + stamp + ".txt"
    if os.path.isdir('kb_sub/infc_grv') == False:
        os.makedirs('kb_sub/infc_grv')
    try:
        report = open('kb_sub/infc_grv/' + filename, "w")
        report.write("Settlement Report By Carrier\n\n")
        report.write("{:<30}\n\n".format(name))
        report.write("        Grievance Number    hours    rate    adjusted      amount       docs       level\n")
        report.write("    ------------------------------------------------------------------------------------\n")
        results = []
        for ug in unique_grv:  # do search for each grievance in list of unique grievances
            sql = "SELECT informalc_awards.grv_no, informalc_awards.hours, informalc_awards.rate, " \
                  "informalc_awards.amount, informalc_grv.docs, informalc_grv.level " \
                  "FROM informalc_awards, informalc_grv " \
                  "WHERE informalc_awards.grv_no = informalc_grv.grv_no and informalc_awards.carrier_name='%s' " \
                  "and informalc_awards.grv_no = '%s'" \
                  "ORDER BY informalc_grv.date_signed" % (name, ug)
            query = inquire(sql)
            if query:
                for q in query:
                    q = list(q)
                    results.append(q)

        if len(results) == 0:
            report.write("    There are no awards on record for this carrier.\n")
        total_adj = 0
        total_amt = 0
        i = 1
        for r in results:
            if r[1]:
                hours = "{0:.2f}".format(float(r[1]))
            else:
                hours = "---"
            if r[2]:
                rate = "{0:.2f}".format(float(r[2]))
            else:
                rate = "---"
            if r[1] and r[2]:
                adj = "{0:.2f}".format(float(r[1]) * float(r[2]))
                total_adj = total_adj + (float(r[1]) * float(r[2]))
            else:
                adj = "---"
            if r[3]:
                amt = "{0:.2f}".format(float(r[3]))
                total_amt = total_amt + float(r[3])
            else:
                amt = "---"
            if r[5] == None or r[5] == "unknown":
                r[5] = "---"
            report.write("    {:<4}{:<18}{:>7}{:>8}{:>12}{:>12}{:>11}{:>12}\n"
                         .format(str(i), r[0], hours, rate, adj, amt, r[4],r[5]))
            i += 1
        report.write("    ------------------------------------------------------------------------------------\n")
        t_adj = "{0:.2f}".format(float(total_adj))
        t_amt = "{0:.2f}".format(float(total_amt))
        report.write("        {:<34}{:>11}\n".format("Total hours as straight time", t_adj))
        report.write("        {:<34}{:>23}\n".format("Total as flat dollar amount", t_amt))
        report.close()
        if sys.platform == "win32":
            os.startfile('kb_sub\\infc_grv\\' + filename)
        if sys.platform == "linux":
            subprocess.call(["xdg-open", 'kb_sub/infc_grv/' + filename])
        if sys.platform == "darwin":
            subprocess.call(["open", 'kb_sub/infc_grv/' + filename])
    except:
        messagebox.showerror("Report Generator", "The report was not generated.")


def informalc_bycarrier(frame, result):
    unique_carrier = informalc_uniquecarrier(result)
    wd = front_window(frame)
    Label(wd[3], text="Informal C: Select Carrier", font="bold").pack(anchor="w")
    Label(wd[3], text="").pack()
    scrollbar = Scrollbar(wd[3], orient=VERTICAL)
    listbox = Listbox(wd[3], selectmode="single", yscrollcommand=scrollbar.set)
    listbox.config(height=30, width=50)
    for name in unique_carrier:
        listbox.insert(END, name)
    scrollbar.config(command=listbox.yview)
    scrollbar.pack(side=RIGHT, fill=Y)
    listbox.pack(side=LEFT, expand=1)
    Button(wd[4], text="Go Back", width=20, command=lambda: informalc_grvlist_result(wd[0], result)).pack(side=LEFT)
    Button(wd[4], text="Report", width=20,
           command=lambda: informalc_apply_bycarrier(result, unique_carrier, listbox.curselection())).pack(side=LEFT)
    rear_window(wd)


def informalc_uniquecarrier(result):
    unique_grv = []
    for grv in result:
        if grv[0] not in unique_grv:
            unique_grv.append(grv[0])
    unique_carrier = []
    for each in unique_grv:
        sql = "SELECT * FROM informalc_awards WHERE grv_no='%s'" % each
        results = inquire(sql)
        for r in results:
            if r[1] not in unique_carrier:
                unique_carrier.append(r[1])
    unique_carrier.sort()
    return unique_carrier


def informalc_rptbygrv(grv_info):
    grv_info = list(grv_info) # correct for legacy problem of NULL Settlement Levels
    if grv_info[8]==None:
        grv_info[8] = "unknown"
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = "infc_grv_list" + "_" + stamp + ".txt"
    if os.path.isdir('kb_sub/infc_grv') == False:
        os.makedirs('kb_sub/infc_grv')
    try:
        report = open('kb_sub/infc_grv/' + filename, "w")
        report.write("Settlement Summary\n\n")
        sql = "SELECT * FROM informalc_awards WHERE grv_no='%s' ORDER BY carrier_name" % grv_info[0]
        query = inquire(sql)
        awardxhour = 0
        awardxamt = 0
        report.write("    Grievance Number:   " + grv_info[0] + "\n")
        start = dt_converter(grv_info[1]).strftime("%m/%d/%Y")
        end = dt_converter(grv_info[2]).strftime("%m/%d/%Y")
        sign = dt_converter(grv_info[3]).strftime("%m/%d/%Y")
        report.write("    Dates of Violation: " + start + " - " + end + "\n")
        report.write("    Signing Date:       " + sign + "\n")
        report.write("    Settlement Level    " + grv_info[8] + "\n")
        report.write("    Station:            " + grv_info[4] + "\n")
        report.write("    GATS Number:        " + grv_info[5] + "\n")
        report.write("    Documentation:      " + grv_info[6] + "\n")
        report.write("    Description:        " + grv_info[7] + "\n\n")
        report.write("    Carrier Name                Hours      Rate   Adjusted     Amount\n")
        report.write("    -----------------------------------------------------------------\n")
        if len(query) == 0:
            report.write("         No awards recorded for this settlement.\n")
        c = 1
        for rec in query:
            hour = 0.0
            rate = 0.0
            amt = 0
            if rec[2]: hour = float(rec[2])
            if rec[3]: rate = float(rec[3])
            if rec[4]: amt = float(rec[4])
            if hour and rate:
                awardxhour = awardxhour + (hour * rate)
            if amt:
                awardxamt = awardxamt + amt
            if rec[2]:
                hours = "{0:.2f}".format(float(rec[2]))
            else:
                hours = "---"
            if rec[3]:
                rate = "{0:.2f}".format(float(rec[3]))
            else:
                rate = "---"
            if rec[2] and rec[3]:
                adj = "{0:.2f}".format(float(rec[2]) * float(rec[3]))
            else:
                adj = "---"
            if rec[4]:
                amt = "{0:.2f}".format(float(rec[4]))
            else:
                amt = "---"
            report.write('    {:<5}{:<22}{:>6}{:>10}{:>10}{:>12}\n'.format(str(c), rec[1], hours, rate, adj, amt))
            c += 1
        report.write("    -----------------------------------------------------------------\n")
        report.write("         {:<38}{:>10}\n".format("Awards adjusted to straight time", "{0:.2f}"
                                                      .format(float(awardxhour))))
        report.write("         {:<38}{:>22}\n".format("Awards as flat dollar amount", "{0:.2f}"
                                                      .format(float(awardxamt))))
        report.write("\n\n\n")
        report.close()
        if sys.platform == "win32":
            os.startfile('kb_sub\\infc_grv\\' + filename)
        if sys.platform == "linux":
            subprocess.call(["xdg-open", 'kb_sub/infc_grv/' + filename])
        if sys.platform == "darwin":
            subprocess.call(["open", 'kb_sub/infc_grv/' + filename])
    except:
        messagebox.showerror("Report Generator", "The report was not generated.")


def informalc_grvlist_setsum(result):
    if len(result) > 0:
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = "infc_grv_list" + "_" + stamp + ".txt"
        if os.path.isdir('kb_sub/infc_grv') == False:
            os.makedirs('kb_sub/infc_grv')
        # try:
        report = open('kb_sub/infc_grv/' + filename, "w")
        report.write("   Settlement List Summary\n")
        report.write("   (ordered by date signed)\n\n")
        report.write('  {:<18}{:<12}{:>9}{:>11}{:>12}{:>12}{:>12}\n'
                     .format("    Grievance #", "Date Signed", "GATS #", "Docs?", "Level", "Hours", "Dollars"))
        report.write("      ----------------------------------------------------------------------------------\n")
        total_hour = 0
        total_amt = 0
        i = 1
        for sett in result:
            sql = "SELECT * FROM informalc_awards WHERE grv_no='%s'" % sett[0]
            query = inquire(sql)
            awardxhour = 0
            awardxamt = 0
            for rec in query:  # calculate total award amounts
                hour = 0.0
                rate = 0.0
                amt = 0
                if rec[2]: hour = float(rec[2])
                if rec[3]: rate = float(rec[3])
                if rec[4]: amt = float(rec[4])
                if hour and rate:
                    awardxhour = awardxhour + (hour * rate)
                if amt:
                    awardxamt = awardxamt + amt
            sign = dt_converter(sett[3]).strftime("%m/%d/%Y")
            s_gats = sett[5].split(" ")
            if sett[8]==None or sett[8]=="unknown":
                lvl = "---"
            else: lvl = sett[8]
            # for gats_no in s_gats:
            for gi in range(len(s_gats)):
                if gi == 0:
                    total_hour += awardxhour
                    total_amt += awardxamt
                    report.write('{:>4}  {:<14}{:<12}{:<9}{:>11}{:>12}{:>12}{:>12}\n'
                                 .format(str(i), sett[0], sign, s_gats[gi], sett[6], lvl
                                         , "{0:.2f}".format(float(awardxhour)), "{0:.2f}".format(float(awardxamt))))
                if gi != 0:
                    report.write('{:<34}{:<12}\n'.format("", s_gats[gi]))
            if i % 3 == 0:
                report.write("      ----------------------------------------------------------------------------------\n")
            i += 1
        report.write("      ----------------------------------------------------------------------------------\n")
        report.write("{:<20}{:>58}\n".format("      Total Hours","{0:.2f}".format(total_hour)))
        report.write("{:<20}{:>70}\n".format("      Total Dollars", "{0:.2f}".format(total_amt)))
        report.close()
        if sys.platform == "win32":
            os.startfile('kb_sub\\infc_grv\\' + filename)
        if sys.platform == "linux":
            subprocess.call(["xdg-open", 'kb_sub/infc_grv/' + filename])
        if sys.platform == "darwin":
            subprocess.call(["open", 'kb_sub/infc_grv/' + filename])


def informalc_grvlist_result(frame, result):
    wd = front_window(frame)
    Label(wd[3], text="Informal C: Search Results", font="bold").grid(row=0, column=0, columnspan=4, sticky="w")
    Label(wd[3], text="").grid(row=1)
    if len(result) == 0:
        Label(wd[3], text="The search has no results.").grid(row=2, column=0, columnspan=4)
    else:
        Label(wd[3], text="Grievance Number", fg="grey", anchor="w").grid(row=2, column=1, sticky="w")
        Label(wd[3], text="Incident Start", fg="grey", anchor="w").grid(row=2, column=2, sticky="w")
        Label(wd[3], text="Incident End", fg="grey", anchor="w").grid(row=2, column=3, sticky="w")
        Label(wd[3], text="Date Signed", fg="grey", anchor="w").grid(row=2, column=4, sticky="w")
    row = 3
    ii = 1
    for r in result:
        Label(wd[3], text=str(ii), anchor="w").grid(row=row, column=0)
        Button(wd[3], text=" " + r[0], anchor="w", width=14, relief=RIDGE).grid(row=row, column=1)
        in_start = datetime.strptime(r[1], '%Y-%m-%d %H:%M:%S')
        in_end = datetime.strptime(r[2], '%Y-%m-%d %H:%M:%S')
        sign_date = datetime.strptime(r[3], '%Y-%m-%d %H:%M:%S')
        Button(wd[3], text=in_start.strftime("%b %d, %Y"), width=11, anchor="w", relief=RIDGE).grid(row=row, column=2)
        Button(wd[3], text=in_end.strftime("%b %d, %Y"), width=11, anchor="w", relief=RIDGE).grid(row=row, column=3)
        Button(wd[3], text=sign_date.strftime("%b %d, %Y"), width=11, anchor="w", relief=RIDGE).grid(row=row, column=4)
        Button(wd[3], text="Edit", width=6, relief=RIDGE, command=lambda x=r[0]: informalc_edit(wd[0], result, x, '')) \
            .grid(row=row, column=5)
        Button(wd[3], text="Report", width=6, relief=RIDGE, command=lambda x=r: informalc_rptbygrv(x)).grid(row=row,
                                                                                                            column=6)
        Button(wd[3], text="Enter Awards", width=10, relief=RIDGE,
               command=lambda x=r[0]: informalc_addaward(wd[0], result, x)).grid(row=row, column=7)
        row += 1
        Label(wd[3], text="         {}".format(r[7]), anchor="w", fg="grey").grid(row=row, column=1, columnspan=5,
                                                                                  sticky="w")
        row += 1
        ii += 1
    Button(wd[4], text="Go Back", width=16, command=lambda: informalc_grvlist(wd[0])).grid(row=0, column=0)
    Label(wd[4], text="Report: ", width=16).grid(row=0, column=1)
    Button(wd[4], text="By Settlements", width=16, command=lambda: informalc_rptgrvsum(result)).grid(row=0, column=2)
    Button(wd[4], text="By Carriers", width=16, command=lambda: informalc_bycarriers(result)).grid(row=0, column=3)
    Button(wd[4], text="By Carrier", width=16, command=lambda: informalc_bycarrier(wd[0], result)).grid(row=0, column=4)
    Label(wd[4], text="Summary: ", width=16).grid(row=1, column=1)
    Button(wd[4], text="By Settlements", width=16, command=lambda: informalc_grvlist_setsum(result)).grid(row=1,
                                                                                                          column=2)
    rear_window(wd)


def informalc_date_checker(date, type):
    d = date.get().split("/")
    if len(d) != 3:
        messagebox.showerror("Invalid Data Entry", "The date for the {} is not properly formatted.".format(type))
        return "fail"
    for num in d:
        if num.isnumeric() == False:
            messagebox.showerror("Invalid Data Entry", "The month, day and year for the {} "
                                                       "must be numeric.".format(type))
            return "fail"
    if len(d[0]) > 2:
        messagebox.showerror("Invalid Data Entry", "The month for the {} must be no more than two digits"
                                                   " long.".format(type))
        return "fail"
    if len(d[1]) > 2:
        messagebox.showerror("Invalid Data Entry", "The day for the {} must be no more than two digits"
                                                   " long.".format(type))
        return "fail"
    if len(d[2]) > 4:
        messagebox.showerror("Invalid Data Entry", "The year for the {} must be no more than four digits long."
                             .format(type))
        return "fail"
    try:
        date = datetime(int(d[2]), int(d[0]), int(d[1]))
        valid_date = True
    except ValueError:
        valid_date = False
    if valid_date == False:
        messagebox.showerror("Invalid Data Entry", "The date entered for {} is not a valid date."
                             .format(type))
        return "fail"


def informalc_grvlist_apply(frame,
                            incident_date, incident_start, incident_end,
                            signing_date, signing_start, signing_end,
                            station, set_lvl, level,
                            gats, have_gats,
                            docs, have_docs):
    conditions = []
    if incident_date.get() == "yes":
        check = informalc_date_checker(incident_start, "starting incident date")
        if check == "fail":
            return
        check = informalc_date_checker(incident_end, "ending incident date")
        if check == "fail":
            return
        d = incident_start.get().split("/")
        start = datetime(int(d[2]), int(d[0]), int(d[1]))
        d = incident_end.get().split("/")
        end = datetime(int(d[2]), int(d[0]), int(d[1]))
        if start > end:
            messagebox.showerror("Invalid Data Entry", "Your starting incident date must be earlier than your "
                                                       "ending incident date.")
            return
        to_add = "indate_start > '{}' and indate_end < '{}'".format(start, end)
        conditions.append(to_add)
    if signing_date.get() == "yes":
        check = informalc_date_checker(signing_start, "starting signing date")
        if check == "fail":
            return
        check = informalc_date_checker(signing_end, "ending signing date")
        if check == "fail":
            return
        d = signing_start.get().split("/")
        start = datetime(int(d[2]), int(d[0]), int(d[1]))
        d = signing_end.get().split("/")
        end = datetime(int(d[2]), int(d[0]), int(d[1]))
        if start > end:
            messagebox.showerror("Invalid Data Entry", "Your starting signing date must be earlier than your "
                                                       "ending signing date.")
            return
        to_add = "date_signed BETWEEN '{}' AND '{}'".format(start, end)
        conditions.append(to_add)
    if station.get() == "Select a Station":
        messagebox.showerror("Invalid Station", "You must select a station.")
        return
    to_add = "station = '{}'".format(station.get())
    conditions.append(to_add)

    if set_lvl.get() == "yes":
        to_add = "level = '{}'".format(level.get())
        conditions.append(to_add)

    if gats.get() == "yes":
        if have_gats.get() == "yes":
            to_add = "gats_number IS NOT ''"
            conditions.append(to_add)
        if have_gats.get() == "no":
            to_add = "gats_number IS ''"
            conditions.append(to_add)
    if docs.get() == "yes":
        to_add = "docs = '{}'".format(have_docs.get())
        conditions.append(to_add)
    where_str = ""
    for i in range(len(conditions)):
        where_str += "{}".format(conditions[i])
        if i + 1 < len(conditions):
            where_str += " and "
    sql = "SELECT * FROM informalc_grv WHERE {} ORDER BY date_signed DESC".format(where_str)
    result = inquire(sql)
    informalc_grvlist_result(frame, result)


def informalc_grvlist(frame):
    wd = front_window(frame)
    Label(wd[3], text="Informal C: Settlement Search Criteria", font="bold").grid(row=0, columnspan=6, sticky="w")
    Label(wd[3], text=" ").grid(row=1, columnspan=6)
    # initialize varibles
    station = StringVar(wd[0])
    incident_date = StringVar(wd[0])
    incident_start = StringVar(wd[0])
    incident_end = StringVar(wd[0])
    signing_date = StringVar(wd[0])
    signing_start = StringVar(wd[0])
    signing_end = StringVar(wd[0])
    set_lvl = StringVar(wd[0])
    level = StringVar(wd[0])
    gats = StringVar(wd[0])
    have_gats = StringVar(wd[0])
    docs = StringVar(wd[0])
    have_docs = StringVar(wd[0])
    # select station
    Label(wd[3], text="Station ").grid(row=2, column=0, columnspan=3, sticky="w")
    station_options = list_of_stations
    if "out of station" in station_options:
        station_options.remove("out of station")
    station_om = OptionMenu(wd[3], station, *station_options)
    station_om.config(width=35)
    station_om.grid(row=2, column=3, columnspan=3, sticky="e")
    station.set("Select a Station")

    Label(wd[3], text="Search For", fg="grey").grid(row=3, column=0, columnspan=2, sticky="w")
    Label(wd[3], text="Category", fg="grey").grid(row=3, column=3)
    Label(wd[3], text="Start", fg="grey").grid(row=3, column=4)
    Label(wd[3], text="End", fg="grey").grid(row=3, column=5)
    # select for starting date
    Radiobutton(wd[3], text="yes", variable=incident_date, value='yes').grid(row=4, column=0, sticky="w")
    Radiobutton(wd[3], text="no", variable=incident_date, value='no').grid(row=4, column=1, sticky="w")
    Label(wd[3], text="", width=2).grid(row=4, column=2)
    Label(wd[3], text="Incident Dates").grid(row=4, column=3, sticky="w")
    Entry(wd[3], textvariable=incident_start, width=12, justify='right').grid(row=4, column=4)
    Entry(wd[3], textvariable=incident_end, width=12, justify='right').grid(row=4, column=5)
    incident_date.set('no')
    # select for signing date
    Radiobutton(wd[3], text="yes", variable=signing_date, value='yes').grid(row=5, column=0, sticky="w")
    Radiobutton(wd[3], text="no", variable=signing_date, value='no').grid(row=5, column=1, sticky="w")
    Label(wd[3], text="Signing Dates").grid(row=5, column=3, sticky="w")
    Entry(wd[3], textvariable=signing_start, width=12, justify='right').grid(row=5, column=4)
    Entry(wd[3], textvariable=signing_end, width=12, justify='right').grid(row=5, column=5)
    signing_date.set('no')

    # select for settlement level
    Radiobutton(wd[3], text="yes", variable=set_lvl, value='yes').grid(row=6, column=0, sticky="w")
    Radiobutton(wd[3], text="no", variable=set_lvl, value='no').grid(row=6, column=1, sticky="w")
    set_lvl.set("no")
    Label(wd[3], text="Settlement Level ").grid(row=6, column=3, sticky="w")
    lvl_options = ("informal a","formal a", "step b", "pre-arb", "arbitration")
    lvl_om = OptionMenu(wd[3], level, *lvl_options)
    lvl_om.config(width=10)
    lvl_om.grid(row=6, column=4, columnspan=3, sticky="e")
    level.set("informal a")

    #select for gats number
    Radiobutton(wd[3], text="yes", variable=gats, value='yes').grid(row=7, column=0, sticky="w")
    Radiobutton(wd[3], text="no", variable=gats, value='no').grid(row=7, column=1, sticky="w")
    Label(wd[3], text="GATS Number").grid(row=7, column=3, sticky="w")
    gats_options =("no","yes")
    gats_om = OptionMenu(wd[3],have_gats, *gats_options)
    gats_om.config(width=10)
    gats_om.grid(row=7, column=4, columnspan=3, sticky="e")
    have_gats.set('no')
    gats.set('no')
    # select for documentation
    Radiobutton(wd[3], text="yes", variable=docs, value='yes').grid(row=9, column=0, sticky="w")
    Radiobutton(wd[3], text="no", variable=docs, value='no').grid(row=9, column=1, sticky="w")
    Label(wd[3], text="Documentation").grid(row=9, column=3, sticky="w")
    doc_options = ("moot", "no", "partial", "yes", "incomplete", "verified")
    docs_om = OptionMenu(wd[3], have_docs, *doc_options)
    docs_om.config(width=10)
    docs_om.grid(row=9, column=4, columnspan=3, sticky="e")
    have_docs.set('no')
    docs.set("no")
    Label(wd[3], text="").grid(row=13)
    # buttons
    Button(wd[4], text="Search", width=20,
           command=lambda: informalc_grvlist_apply(wd[0],incident_date, incident_start, incident_end,
           signing_date, signing_start, signing_end, station, set_lvl, level, gats, have_gats,docs, have_docs))\
            .grid(row=0, column=1)
    Button(wd[4], text="Go Back", width=20, anchor="w", command=lambda: informalc(wd[0])).grid(row=0, column=0)
    rear_window(wd)


def informalc_new(frame, msg):
    wd = front_window(frame)  # F,S,C,FF,buttons
    Label(wd[3], text="New Settlement", font="bold").grid(row=0, column=0, sticky="w")
    Label(wd[3], text="").grid(row=1, column=0, sticky="w")
    Label(wd[3], text="Grievance Number: ").grid(row=2, column=0, sticky="w")
    grv_no = StringVar(wd[0])
    Entry(wd[3], textvariable=grv_no, justify='right').grid(row=2, column=1, sticky="w")
    Label(wd[3], text="Incident Date").grid(row=3, column=0, sticky="w")
    Label(wd[3], text="  Start (mm/dd/yyyy): ").grid(row=4, column=0, sticky="w")
    incident_start = StringVar(wd[0])
    Entry(wd[3], textvariable=incident_start, justify='right').grid(row=4, column=1, sticky="w")
    Label(wd[3], text="  End (mm/dd/yyyy): ").grid(row=5, column=0, sticky="w")
    incident_end = StringVar(wd[0])
    Entry(wd[3], textvariable=incident_end, justify='right').grid(row=5, column=1, sticky="w")
    Label(wd[3], text="Date Signed (mm/dd/yyyy): ").grid(row=6, column=0, sticky="w")
    date_signed = StringVar(wd[0])
    Entry(wd[3], textvariable=date_signed, justify='right').grid(row=6, column=1, sticky="w")
    #select level
    Label(wd[3], text="Settlement Level: ").grid(row=7, column=0, sticky="w")  # select settlement level
    lvl = StringVar(wd[0])
    lvl_options = ("informal a", "formal a", "step b", "pre arb", "arbitration")
    lvl_om = OptionMenu(wd[3], lvl, *lvl_options)
    lvl_om.config(width=13)
    lvl_om.grid(row=7, column=1)
    lvl.set("informal a")
    Label(wd[3], text="Station: ").grid(row=8, column=0, sticky="w")  # select a station
    station = StringVar(wd[0])
    station.set("Select a Station")
    station_options = list_of_stations
    if "out of station" in station_options:
        station_options.remove("out of station")
    station_om = OptionMenu(wd[3], station, *station_options)
    station_om.config(width=40)
    station_om.grid(row=9, column=0, columnspan=2, sticky="e")
    Label(wd[3], text="GATS Number: ").grid(row=10, column=0, sticky="w") # enter gats number
    gats_number = StringVar(wd[0])
    Entry(wd[3], textvariable=gats_number, justify='right').grid(row=10, column=1, sticky="w")
    Label(wd[3], text="Documentation?: ").grid(row=11, column=0, sticky="w") # select documentation
    docs = StringVar(wd[0])
    doc_options = ("moot", "no", "partial", "yes", "incomplete", "verified")
    docs_om = OptionMenu(wd[3],docs, *doc_options)
    docs_om.config(width=13)
    docs_om.grid(row=11, column=1)
    docs.set("no")
    Label(wd[3], text="Description: ").grid(row=15, column=0, sticky="w")
    description = StringVar(wd[0])
    Entry(wd[3], textvariable=description, width=48, justify='right').grid(row=16, column=0, sticky="w", columnspan=2)
    Label(wd[3], text="").grid(row=17, column=0)
    Label(wd[3], text=msg, fg="red").grid(row=18, column=0, columnspan=2, sticky="w")
    Button(wd[4], text="Go Back", width=20, anchor="w", command=lambda: informalc(wd[0])).grid(row=0, column=0)
    Button(wd[4], text="Enter", width=18, command=lambda: informalc_new_apply(wd[0], grv_no, incident_start,
          incident_end, date_signed, station, gats_number, docs, description, lvl)).grid(row=0, column=1)
    rear_window(wd)


def informalc_poe_apply_search(frame, year, station, backdate):
    if year.get().strip() == "":
        messagebox.showerror("Data Entry Error", "You must enter a year.")
        return
    if "." in year.get():
        messagebox.showerror("Data Entry Error", "The year can not contain decimal points.")
        return
    if year.get().isnumeric() == False:
        messagebox.showerror("Data Entry Error", "The year must numeric without any letters or special characters.")
        return
    if float(year.get()) > 9999 or float(year.get()) < 2:
        messagebox.showerror("Data Entry Error", "The year must be between the year 2 and 9999.\nI think I'm being "
                                                 "reasonable.")
        return

        return
    if station.get() == "undefined":
        messagebox.showerror("Data Entry Error", "You must select a station.")
        return
    weeks = int(backdate.get()) * 52
    dt_year = datetime(int(year.get()), int(1), int(1))
    dt_start = dt_year - timedelta(weeks=weeks)
    year = year.get()
    array = []
    selection = "none"
    msg = ""
    informalc_poe_listbox(dt_year, station, dt_start, year)
    informalc_poe_add(frame, array, selection, year, msg)


def informalc_poe_apply_add(frame, name, year, buttons):
    if name == "none":
        messagebox.showerror("Data Entry Error", "You must select a name.")
        return
    for i in range(len(poe_add_pay_periods)):
        pp = poe_add_pay_periods[i].get().strip()
        hr = poe_add_hours[i].get().strip()
        rt = poe_add_rate[i].get().strip()
        amt = poe_add_amount[i].get().strip()
        if pp and not isint(pp):
            messagebox.showerror("Data Input Error", "Input error for {} in row {}. The pay period must be a number"
                                 .format(name, str(i + 1)))
            return
        if pp and int(pp) > 27:
            messagebox.showerror("Data Input Error", "Input error for {} in row {}. The pay period can not be greater "
                                                     "than 27".format(name, str(i + 1)))
            return
        if hr and amt:
            messagebox.showerror("Data Input Error", "Input error for {} in row {}. You can not enter both hours and "
                                                     "amount. You can only enter one or another, but not both. Awards can be in the form of "
                                                     "hours at a given rate OR an amount.".format(name, str(i + 1)))
            return
        if rt and amt:
            messagebox.showerror("Data Input Error", "Input error for {} in row {}. You can not enter both rate and "
                                                     "amount. You can only enter one or another, but not both. Awards can be in the form of "
                                                     "hours at a given rate OR an amount.".format(name, str(i + 1)))
            return
        if hr and not rt:
            messagebox.showerror("Data Input Error", "Input error for {} in row {}. Hours must be a accompanied by a "
                                                     "rate.".format(name, str(i + 1)))
            return
        if rt and not hr:
            messagebox.showerror("Data Input Error", "Input error for {} in row {}. Rate must be a accompanied by a "
                                                     "hours.".format(name, str(i + 1)))
            return
        if hr and isfloat(hr) == False:
            messagebox.showerror("Data Input Error", "Input error for {} in row {}. Hours must be a number."
                                 .format(name, str(i + 1)))
            return
        if hr and '.' in hr:
            s_hrs = hr.split(".")
            if len(s_hrs[1]) > 2:
                messagebox.showerror("Data Input Error", "Input error for {} in row {}. Hours must have no "
                                                         "more than 2 decimal places.".format(name, str(i + 1)))
                return
        if rt and amt:
            messagebox.showerror("Data Input Error", "Input error for {} in row {}. You can not enter both rate and "
                                                     "amount. You can only enter one or the other, but not both. Awards can be in the form of "
                                                     "hours at a given rate OR an amount.".format(name, str(i + 1)))
            return
        if rt and isfloat(rt) == False:
            messagebox.showerror("Data Input Error", "Input error for {} in row {}. Rate must be a number."
                                 .format(name, str(i + 1)))
            return
        if rt and '.' in rt:
            s_rate = rt.split(".")
            if len(s_rate[1]) > 2:
                messagebox.showerror("Data Input Error", "Input error for {} in row {}. Rates must have no "
                                                         "more than 2 decimal places.".format(name, str(i + 1)))
                return
        if rt and float(rt) > 10:
            messagebox.showerror("Data Input Error", "Input error for {} in row {}. Values greater than 10 are not "
                                                     "accepted. \n"
                                                     "Note the following rates would be expressed as: \n "
                                                     "additional %50         .50 or just .5 \n"
                                                     "straight time rate     1.00 or just 1 \n"
                                                     "overtime rate          1.50 or 1.5 \n"
                                                     "penalty rate           2.00 or just 2".format(name, str(i + 1)))
            return
        if amt and isfloat(amt) == False:
            messagebox.showerror("Data Input Error", "Input error for {} in row {}. Amounts can only be expressed as "
                                                     "numbers. No special characters, such as $ are allowed."
                                 .format(name, str(i + 1)))
            return
        if amt and '.' in amt:
            s_amt = amt.split(".")
            if len(s_amt[1]) > 2:
                messagebox.showerror("Data Input Error", "Input error for {} in row {}. Amounts must have no "
                                                         "more than 2 decimal places.".format(name, str(i + 1)))
                return

    pb_label = Label(buttons, text="Updating Changes: ")  # make label for progress bar
    pb_label.grid(row=1, column=2)
    pb = ttk.Progressbar(buttons, length=200, mode="determinate")  # create progress bar
    pb.grid(row=1, column=3)
    pb["maximum"] = len(poe_add_pay_periods) * 2  # set length of progress bar
    pb.start()
    sql = "DELETE FROM informalc_payouts WHERE year='%s' and carrier_name='%s'" % (year, name)
    pb["value"] = len(poe_add_pay_periods)  # increment progress bar
    buttons.update()
    commit(sql)
    ii = len(poe_add_pay_periods)
    count = 0
    paydays = []
    for i in range(len(poe_add_pay_periods)):
        if poe_add_pay_periods[i].get().strip() != "":
            if poe_add_hours[i].get().strip() != "" and poe_add_rate[i].get().strip() != "" \
                    or poe_add_amount[i].get().strip() != "":
                pp = poe_add_pay_periods[i].get().zfill(2)
                one = "1"
                pp = pp + one  # format pp so it can fit in find_pp()
                dt = find_pp(int(year), pp)
                dt += timedelta(days=20)
                paydays.append(dt)
                sql = "INSERT INTO informalc_payouts (year,pp,payday,carrier_name,hours,rate,amount) " \
                      "VALUES('%s','%s','%s','%s','%s','%s','%s')" \
                      % (year, poe_add_pay_periods[i].get().strip(), paydays[i], name, poe_add_hours[i].get().strip()
                         , poe_add_rate[i].get().strip(), poe_add_amount[i].get().strip())
                commit(sql)
                count += 1
                ii += 1
                pb["value"] = ii  # increment progress bar
                buttons.update()
    pb.stop()  # stop and destroy the progress bar
    pb_label.destroy()  # destroy the label for the progress bar
    pb.destroy()
    array = []
    selection = "none"
    msg = "Update: {} records for {} have been recorded in the database.".format(count, name)
    informalc_poe_add(frame, array, selection, year, msg)


def informalc_poe_add_plus(frame, payouts):
    if len(payouts) == 0:
        poe_add_pay_periods.append(StringVar(frame))  # set up array of stringvars for hours,rate,amount
        poe_add_hours.append(StringVar(frame))
        poe_add_rate.append(StringVar(frame))
        poe_add_amount.append(StringVar(frame))
        Entry(frame, textvariable=poe_add_pay_periods[len(poe_add_pay_periods) - 1], width=10) \
            .grid(row=len(poe_add_pay_periods) + 6, column=0, pady=5, padx=5, sticky="w")
        Entry(frame, textvariable=poe_add_hours[len(poe_add_hours) - 1], width=10) \
            .grid(row=len(poe_add_hours) + 6, column=1, pady=5, padx=5)
        Entry(frame, textvariable=poe_add_rate[len(poe_add_rate) - 1], width=10) \
            .grid(row=len(poe_add_rate) + 6, column=2, pady=5, padx=5)
        Entry(frame, textvariable=poe_add_amount[len(poe_add_amount) - 1], width=10) \
            .grid(row=len(poe_add_amount) + 6, column=3, pady=5, padx=5)
    else:
        for i in range(len(payouts)):
            poe_add_pay_periods.append(StringVar(frame))  # set up array of stringvars for hours,rate,amount
            poe_add_hours.append(StringVar(frame))
            poe_add_rate.append(StringVar(frame))
            poe_add_amount.append(StringVar(frame))
            poe_add_pay_periods[i].set(payouts[i][1])
            poe_add_hours[i].set(payouts[i][4])
            poe_add_rate[i].set(payouts[i][5])
            poe_add_amount[i].set(payouts[i][6])
            Entry(frame, textvariable=poe_add_pay_periods[i], width=10) \
                .grid(row=len(poe_add_pay_periods) + 6, column=0, sticky="w")
            Entry(frame, textvariable=poe_add_hours[i], width=10) \
                .grid(row=len(poe_add_hours) + 6, column=1, pady=5, padx=5)
            Entry(frame, textvariable=poe_add_rate[i], width=10) \
                .grid(row=len(poe_add_rate) + 6, column=2, pady=5, padx=5)
            Entry(frame, textvariable=poe_add_amount[i], width=10) \
                .grid(row=len(poe_add_amount) + 6, column=3, pady=5, padx=5)


def informalc_poe_add(frame, array, selection, year, msg):
    empty_array = []
    global poe_add_pay_periods
    global poe_add_hours
    global poe_add_rate
    global poe_add_amount
    poe_add_pay_periods = []
    poe_add_hours = []
    poe_add_rate = []
    poe_add_amount = []
    global informalc_poe_gadd
    wd = front_window(frame)
    informalc_poe_gadd = wd[0]
    Label(wd[3], text="Informal C: Payout Entry", font="bold").grid(row=0, column=0, sticky="w", columnspan=5)
    Label(wd[3], text="").grid(row=1)
    if selection != "none":
        Label(wd[3], text=array[int(selection[0])], font="bold").grid(row=2, column=0, sticky="w", columnspan=5)
        name = array[int(selection[0])]
        Label(wd[3], text="Year: {}".format(year)).grid(row=3, column=0, sticky="w")
        Label(wd[3], text="").grid(row=4)
        Label(wd[3], text="PP", width=10, fg="grey").grid(row=5, column=0, sticky="w")
        Label(wd[3], text="Hours", width=10, fg="grey").grid(row=5, column=1, sticky="w")
        Label(wd[3], text="Rate", width=10, fg="grey").grid(row=5, column=2, sticky="w")
        Label(wd[3], text="Amount", width=10, fg="grey").grid(row=5, column=3, sticky="w")
        Button(wd[3], text="Add Payouts", width=10,
               command=lambda: informalc_poe_add_plus(wd[3], empty_array)).grid(row=5, column=4, sticky="w")
        sql = "SELECT * FROM informalc_payouts WHERE year ='%s' and carrier_name='%s'ORDER BY pp" \
              % (year, name)
        payouts = inquire(sql)
        informalc_poe_add_plus(wd[3], payouts)
    else:
        Label(wd[3], text="Select a carrier from the carrier list.").grid(row=2, column=0, sticky="w", columnspan=5)
        name = "none"
    if msg != "":  # display a message when there is a message
        Label(wd[4], text=msg, fg="red", width=60, anchor="w").grid(row=0, column=0, columnspan=4, sticky="w")
    Button(wd[4], text="Go Back", width=20, command=lambda: informalc_poe_goback(wd[0])) \
        .grid(row=1, column=0, sticky="w")
    Button(wd[4], text="Apply", width=20,
           command=lambda: informalc_poe_apply_add(wd[0], name, year, wd[4])) \
        .grid(row=1, column=1, sticky="w")
    Label(wd[4], text="", width=10).grid(row=1, column=2)
    Label(wd[4], text="", width=10).grid(row=1, column=3)
    rear_window(wd)


def informalc_poe_goback(frame):
    try:
        informalc_poe_lbox.destroy()
    except:
        pass
    informalc_poe_search(frame)


def informalc_poe_listbox(dt_year, station, dt_start, year):
    global informalc_poe_lbox  # initialize the global
    poe_root = Tk()
    informalc_poe_lbox = poe_root  # set the global
    if sys.platform == "win32":
        try:
            poe_root.iconbitmap(r'kb_sub/kb_images/kb_icon2.ico')
        except:
            pass
    if sys.platform == "linux":
        try:
            img = PhotoImage(file='kb_sub/kb_images/kb_icon2.gif')
            poe_root.tk.call('wm', 'iconphoto', poe_root._w, img)
        except:
            pass
    poe_root.title("KLUSTERBOX")
    x_position = root.winfo_x() + 450
    y_position = root.winfo_y() - 25
    poe_root.geometry("%dx%d+%d+%d" % (240, 600, x_position, y_position))
    n_F = Frame(poe_root)
    n_F.pack()
    n_buttons = Canvas(n_F)  # button bar
    n_buttons.pack(fill=BOTH, side=BOTTOM)
    Label(n_F, text="Carrier List", font="bold").pack(anchor="w")
    Label(n_F, text="{} Station:".format(station.get())).pack(anchor="w")
    Label(n_F, text="{} though {}".format(dt_year.strftime("%Y"), dt_start.strftime("%Y"))).pack(anchor="w")
    Label(n_F, text="").pack()
    scrollbar = Scrollbar(n_F, orient=VERTICAL)
    listbox = Listbox(n_F, selectmode="single", yscrollcommand=scrollbar.set)
    listbox.config(height=100, width=50)
    c_list = informalc_gen_clist(dt_start, dt_year, station.get())
    for name in c_list:
        listbox.insert(END, name)
    scrollbar.config(command=listbox.yview)
    scrollbar.pack(side=RIGHT, fill=Y)
    listbox.pack(side=LEFT, expand=1)
    msg = ""
    Button(n_buttons, text="Add Carrier", width=10,
           command=lambda: informalc_poe_add(informalc_poe_gadd, c_list, listbox.curselection(), year, msg)) \
        .pack(side=LEFT, anchor="w")
    Button(n_buttons, text="Close", width=10,
           command=lambda: (poe_root.destroy())).pack(side=LEFT, anchor="w")


def informalc_poe_search(frame):
    wd = front_window(frame)
    the_year = StringVar(wd[0])
    start_year = StringVar(wd[0])
    the_station = StringVar(wd[0])
    station_options = list_of_stations
    if "out of station" in station_options:
        station_options.remove("out of station")
    the_station.set("undefined")
    backdate = StringVar(wd[0])
    backdate.set("1")
    Label(wd[3], text="Informal C: Payout Entry Criteria", font="bold").grid(row=0, column=0, sticky="w", columnspan=4)
    Label(wd[3], text="").grid(row=1)
    Label(wd[3], text="Enter the year and the station to be updated.").grid(row=2, column=0, columnspan=4, sticky="w")
    Label(wd[3], text="\t\t\tYear: ").grid(row=3, column=1, sticky="e")
    Entry(wd[3], textvariable=the_year, width=12).grid(row=3, column=2, sticky="w")
    Label(wd[3], text="Station").grid(row=4, column=1, sticky="e")
    om_station = OptionMenu(wd[3], the_station, *station_options)
    om_station.config(width=28)
    om_station.grid(row=4, column=2, columnspan=2)

    Label(wd[3], text="Build the carrier list by going back how many years?") \
        .grid(row=5, column=0, columnspan=4, sticky="w")
    Label(wd[3], text="Back Date: ").grid(row=6, column=1, sticky="w")
    om_backdate = OptionMenu(wd[3], backdate, "1", "2", "3", "4", "5", "6", "7", "8", "9", "10")
    om_backdate.config(width=5)
    om_backdate.grid(row=6, column=2, sticky="w")
    Button(wd[4], text="Go Back", width=20, command=lambda: informalc(wd[0])).grid(row=0, column=1, sticky="w")
    Button(wd[4], text="Apply", width=20,
           command=lambda: informalc_poe_apply_search(wd[0], the_year, the_station, backdate)) \
        .grid(row=0, column=2, sticky="w")
    rear_window(wd)


def informalc_date_converter(date):  # be sure to run informalc date checker before using this
    sd = date.get().split("/")
    dt = datetime(int(sd[2]), int(sd[0]), int(sd[1]))
    return dt


def informalc_por_all(afterdate, beforedate, station, backdate):
    check = informalc_date_checker(afterdate, "After Date")
    if check == "fail":
        return
    check = informalc_date_checker(beforedate, "Before Date")
    if check == "fail":
        return
    start = informalc_date_converter(afterdate)
    end = informalc_date_converter(beforedate)
    if start > end:
        messagebox.showerror("Data Entry Error", "The After Date can not be earlier than the Before Date")
        return
    if station.get() == "undefined":
        messagebox.showerror("Data Entry Error", "You must select a station. ")
        return
    weeks = int(backdate.get()) * 52
    clist_start = start - timedelta(weeks=weeks)
    carrier_list = informalc_gen_clist(clist_start, end, station.get())

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = "infc_grv_list" + "_" + stamp + ".txt"
    if os.path.isdir('kb_sub/infc_grv') == False:
        os.makedirs('kb_sub/infc_grv')

    report = open('kb_sub/infc_grv/' + filename, "w")
    report.write("  Payouts Report\n\n")
    report.write("  Range of Dates: " + start.strftime("%b %d, %Y") + " - " + end.strftime("%b %d, %Y") + "\n\n")

    for name in carrier_list:
        sql = "SELECT * FROM informalc_payouts WHERE carrier_name = '%s' AND payday BETWEEN '%s' AND '%s' " \
              "ORDER BY payday DESC" % (name, start, end)
        results = inquire(sql)
        if results:
            payxamt = 0
            payxadj = 0
            report.write("  " + name + "\n\n")
            report.write("    PP          Payday          Hours   Rate  Adjusted      Amount\n")
            report.write("    --------------------------------------------------------------\n")
            for result in results:
                hour = 0.0
                rate = 0.0
                amt = 0.0
                if result[4]: hour = float(result[4])
                if result[5]: rate = float(result[5])
                if result[6]: amt = float(result[6])
                if hour and rate:
                    payxadj = payxadj + (hour * rate)
                if amt:
                    payxamt = payxamt + amt
                pp = result[0] + "-" + result[1].zfill(2)
                payday = dt_converter(result[2]).strftime("%b %d, %Y")
                if result[4]:
                    hours = "{0:.2f}".format(float(result[4]))
                else:
                    hours = "---"
                if result[5]:
                    rate = "{0:.2f}".format(float(result[5]))
                else:
                    rate = "---"
                if result[4] and result[5]:
                    adj = "{0:.2f}".format(float(result[4]) * float(result[5]))
                else:
                    adj = "---"
                if result[6]:
                    amt = "{0:.2f}".format(float(result[6]))
                else:
                    amt = "---"
                # report.write(result[0]+" - "+result[1]+result[2]+result[3]+result[4]+result[5]+result[6]+"\n")
                report.write('    {:<5}{:>17}{:>9}{:>7}{:>10}{:>12}\n'.format(pp, payday, hours, rate, adj, amt))
            report.write("    --------------------------------------------------------------\n")
            report.write("    {:<40}{:>10}\n".format("Payouts adjusted to straight time", "{0:.2f}"
                                                     .format(float(payxadj))))
            report.write("    {:<38}{:>24}\n".format("Payouts as flat dollar amount", "{0:.2f}"
                                                     .format(float(payxamt))))
            report.write("\n\n\n")

    report.close()
    if sys.platform == "win32":
        os.startfile('kb_sub\\infc_grv\\' + filename)
    if sys.platform == "linux":
        subprocess.call(["xdg-open", 'kb_sub/infc_grv/' + filename])
    if sys.platform == "darwin":
        subprocess.call(["open", 'kb_sub/infc_grv/' + filename])


def informalc_por(frame):
    wd = front_window(frame)
    afterdate = StringVar(wd[0])
    beforedate = StringVar(wd[0])
    station = StringVar(wd[0])
    station_options = list_of_stations
    if "out of station" in station_options:
        station_options.remove("out of station")
    station.set("undefined")
    backdate = StringVar(wd[0])
    backdate.set("1")
    Label(wd[3], text="Informal C: Payout Report Search Criteria", font="bold") \
        .grid(row=0, column=0, columnspan=4, sticky="w")
    Label(wd[3], text="").grid(row=1)
    Label(wd[3], text="Enter range of dates and select station").grid(row=2, column=0, columnspan=4, sticky="w")
    Label(wd[3], text="\tProvide dates in mm/dd/yyyy format.", fg="grey").grid(row=3, column=0, columnspan=4,
                                                                               sticky="w")
    Label(wd[3], text="", width=20).grid(row=4, column=0)
    Label(wd[3], text="After Date: ").grid(row=4, column=1, sticky="w")
    Entry(wd[3], textvariable=afterdate, width=16).grid(row=4, column=2, sticky="w")
    Label(wd[3], text="Before Date: ").grid(row=5, column=1, sticky="w")
    Entry(wd[3], textvariable=beforedate, width=16).grid(row=5, column=2, sticky="w")
    Label(wd[3], text="Station: ").grid(row=6, column=1, sticky="w")
    om_station = OptionMenu(wd[3], station, *station_options)
    om_station.config(width=28)
    om_station.grid(row=6, column=2, columnspan=2)
    Label(wd[3], text="Build the carrier list by going back how many years?") \
        .grid(row=7, column=0, columnspan=4, sticky="w")
    Label(wd[3], text="Back Date: ").grid(row=8, column=1, sticky="w")
    om_backdate = OptionMenu(wd[3], backdate, "1", "2", "3", "4", "5", "6", "7", "8", "9", "10")
    om_backdate.config(width=5)
    om_backdate.grid(row=8, column=2, sticky="w")
    Button(wd[4], text="Go Back", width=16, command=lambda: informalc(wd[0])).grid(row=0, column=0)
    Label(wd[4], text="Report: ", width=16).grid(row=0, column=1)
    Button(wd[4], text="All Carriers", width=16,
           command=lambda: informalc_por_all(afterdate, beforedate, station, backdate)).grid(row=0, column=2)
    Button(wd[4], text="By Carrier", width=16).grid(row=0, column=3)
    rear_window(wd)

def informalc(frame):
    if os.path.isdir('kb_sub/infc_grv') == True:  # clear contents of temp folder
        shutil.rmtree('kb_sub/infc_grv')

    sql = 'CREATE table IF NOT EXISTS informalc_grv (grv_no varchar, indate_start varchar, indate_end varchar,' \
          'date_signed varchar, station varchar, gats_number varchar, docs varchar, description varchar, level varchar)'
    commit(sql)
    # modify table for legacy version which did not have level column of informalc_grv table.
    sql = 'PRAGMA table_info(informalc_grv)'  # get table info. returns an array of columns.
    result = inquire(sql)
    if len(result) <= 8:  # if there are not enough columns add the leave type and leave time columns
        sql = 'ALTER table informalc_grv ADD COLUMN level varchar'
        commit(sql)
    sql = 'CREATE table IF NOT EXISTS informalc_awards (grv_no varchar,carrier_name varchar, hours varchar, ' \
          'rate varchar, amount varchar)'
    commit(sql)
    sql = 'CREATE table IF NOT EXISTS informalc_payouts(year varchar,pp varchar,payday varchar,carrier_name varchar,' \
          'hours varchar,rate varchar,amount varchar)'
    commit(sql)
    # put out of station back into the list of stations in case it has been removed.
    global list_of_stations
    if "out of station" not in list_of_stations:
        list_of_stations.append("out of station")
    wd = front_window(frame)  # F,S,C,FF,buttons
    Label(wd[3], text="Informal C", font="bold").grid(row=0, sticky="w")
    Label(wd[3], text="The C is for Compliance").grid(row=1, sticky="w")
    Label(wd[3], text="").grid(row=2)
    Button(wd[3], text="New Settlement", width=30, command=lambda: informalc_new(wd[0], " ")).grid(row=3, pady=5)
    Button(wd[3], text="Settlement List", width=30, command=lambda: informalc_grvlist(wd[0])).grid(row=4, pady=5)
    Button(wd[3], text="Payout Entry", width=30, command=lambda: informalc_poe_search(wd[0])).grid(row=5, pady=5)
    Button(wd[3], text="Payout Report", width=30, command=lambda: informalc_por(wd[0])).grid(row=6, pady=5)
    Label(wd[3], text="", width=70).grid(row=7)
    Button(wd[4], text="Go Back", width=20, anchor="w", command=lambda: (wd[0].destroy(), main_frame())).grid(row=0,
                                                                                                              column=0)
    rear_window(wd)

def wkly_avail(frame):  # creates a spreadsheet which shows weekly otdl availability
    file_path = filedialog.askopenfilename(initialdir=os.getcwd(), filetypes=[("Excel files", "*.csv *.xls")])
    if file_path[-4:].lower() == ".csv" or file_path[-4:].lower() == ".xls":
        pass
    else:
        messagebox.showerror("Report Generator", "The file you have selected is not a .csv or .xls file.\n"
                                                 "You must select a file with a .csv or .xls extension.")
        return
    with open(file_path, newline="") as file:
        a_file = csv.reader(file)
        c = 0
        for line in a_file:
            if c == 0 and line[0][:8] != "TAC500R3":
                messagebox.showwarning("File Selection Error", "The selected file does not appear to be an "
                                                               "Employee Everything report.")
                return
            if c == 3:
                tacs_pp = line[0]  # find the pay period
                tacs_station = line[2]  # find the station
                break
            c += 1
        c = 0
        range_days = ["Saturday", "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
        for line in a_file:  # find the range
            if line[18] in range_days:
                range_days.remove(line[18])
            if c == 150: break  # survey 150 lines before breaking to anaylize results.
            c += 1
        if len(range_days) > 5:
            messagebox.showwarning("File Selection Error", "Employee Everything Reports that cover only one day /n"
                                                           "are not supported in version {} of Klusterbox.".format(version))
            return
        else:
            t_range = "week"
    year = int(tacs_pp[:-3])  # set the globals
    pp = tacs_pp[-3:]
    t_date = find_pp(year, pp)
    s_year = t_date.strftime("%Y")
    s_mo = t_date.strftime("%m")
    s_day = t_date.strftime("%d")
    sql = "SELECT kb_station FROM station_index WHERE tacs_station = '%s'" % tacs_station
    station = inquire(sql)  # check to see if station has match in station index
    if not station:
        messagebox.showwarning("Error", "This station has not been matched with Auto Data Entry.")
        return
    set_globals(s_year, s_mo, s_day, t_range, station[0][0], "None")  # set the investigation range
    # get the otdl list from the carriers table
    sql = "SELECT carrier_name FROM carriers WHERE effective_date <= '%s' and station = '%s' and list_status = '%s'" \
          "ORDER BY carrier_name, effective_date desc" % (g_date[6], g_station, 'otdl')
    results = inquire(sql)  # call function to access database
    unique_carriers = []  # create non repeating list of otdl carriers
    for name in results:
        if name[0] not in unique_carriers:
            unique_carriers.append(name[0])
    wkly_list = []  # initialize arrays for data sorting
    otdl_list = []  # pull info from ee for these carriers
    on_list = "no"
    station_anchor = "no"
    for name in unique_carriers:
        ot_wkly = []
        sql = "SELECT emp_id FROM name_index WHERE kb_name='%s'" % (name)
        results = inquire(sql)
        if results:  # record emp id to otdl carrier info
            ot_wkly.append(results[0][0])
        else:  # mark otdl carriers who don't have emp id available
            ot_wkly.append("no index")
        sql = "SELECT effective_date,list_status,station FROM carriers " \
              "WHERE carrier_name='%s' and effective_date<='%s'" \
              "ORDER BY effective_date desc" % (name, g_date[6])
        results = inquire(sql)
        ot_wkly.append(name)
        for date in g_date:  # loop for each day of the week
            for rec in results:  # loop for each record starting from the latest
                if rec[2] == g_station:  # if there is a station match
                    station_anchor = "yes"  # mark the carrier as attached to station
                if datetime.strptime(rec[0],
                                     '%Y-%m-%d %H:%M:%S') <= date:  # if the rec is at or earlier than investigation.
                    if rec[1] == "otdl":  # note whether otdl or not.
                        ot_wkly.append("otdl")
                        on_list = "yes"
                    else:
                        ot_wkly.append("")
                    break  # stop. we only want the first
        if on_list == "yes" and station_anchor == "yes":
            wkly_list.append(ot_wkly)  # fill in array with carrier and otdl data
            otdl_list.append(ot_wkly[0])  # add to list of carriers who will be researched
        on_list = "no"  # reset
        station_anchor == "no"  # reset
    not_indexed = []
    for name in wkly_list:  # check to see if there are any otdl carriers who do not have a rec in name index
        if name[0] == "no index":
            not_indexed.append(name[1])  # add any names who do not into an array
    if len(not_indexed) != 0:  # message box info that some otdl do not have a record in the name index
        messagebox.showwarning("Missing Data", "There are {} name/s which have not been matched with their employee id."
                                               " Please exit and run the Auto Data Entry Feature to ensure that all carriers have "
                                               " employee ids entered into Klusterbox.".format(len(not_indexed)))
    if len(otdl_list) == 0:
        messagebox.showwarning("Empty OTDL", "Klusterbox has no records of any otdl carriers for {} station "
                                             "for the week of {}. This could mean that: \n1. The carrier list is empty. Run the "
                                             "Automatic Data Entry Feature, selecting the Employee Everything Report you used here "
                                             " to remedy this. You do not have to enter the rings data at the final step "
                                             " \n2. The Name Index which matches the carrier name to the employee id "
                                             "empty. As in #1, run the Automatic Data Entry Feature to fix this.\n3. The carrier list "
                                             "has no otdl carriers "
                                             "designated. Use the Multi Input Feature to designate otdl carriers. \n"
                                             "This Weekly Availability Report can not be generated without a list of otdl carriers. "
                                             "Build the carrier list/otdl before re-running Weekly Availability."
                               .format(g_station, g_date[0].strftime("%b %d, %Y")))
        frame.destroy()
        main_frame()
    else:  # if there is an otdl then build array holding hours for each day
        days = ("Saturday", "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday")
        extra_hour_codes = ("49", "52", "55", "56", "57", "58", "59", "60")
        running_total = 0
        with open(file_path, newline="") as file:
            a_file = csv.reader(file)
            c = 0
            all_otdl = []
            good_id = "no"
            day_over = "empty"
            long_day = ["Saturday", "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
            sat = 0
            sun = 0
            mon = 0
            tue = 0
            wed = 0
            thr = 0
            fri = 0
            day_run = [sat, sun, mon, tue, wed, thr, fri]
            for line in a_file:
                if c != 0 and line[4].zfill(8) in otdl_list:  # if the emp_id matches ones we are looking for
                    if line[18] == "Base" and good_id != "no":
                        sql = "SELECT kb_name FROM name_index WHERE emp_id='%s'" % good_id
                        result = inquire(sql)  # get the kb name with the emp id
                        all_day_run = []
                        for i in range(7):
                            all_day_run.append(day_run[i])
                        to_add = ([result[0][0]] + all_day_run + [day_over])
                        all_otdl.append(to_add)
                        for i in range(len(long_day)):
                            day_run[i] = 0  # empty each day in day run
                        day_over = "empty"  # reset
                        running_total = 0  # reset
                    if line[18] == "Base" and line[19] == "844" or line[19] == "134":  # find first line of specific carrier
                        good_id = line[4].zfill(8)  # remember id of carriers who are FT or aux carriers
                    if good_id == line[4].zfill(8) and line[18] != "Base":
                        if line[18] in days:  # get the hours for each day
                            spt_20 = line[20].split(':')  # split to get code and hours
                            hr_type = spt_20[0][1] + spt_20[0][2]  # parse hour code to 2 digits
                            if hr_type in extra_hour_codes:  # if hr_type in hr_codes:
                                running_total += float(spt_20[1])
                                i = 0
                                for ld in long_day:
                                    if ld == line[18]:
                                        day_run[i] += float(spt_20[1])
                                    i += 1
                            if day_over == "empty" and running_total > 60:
                                day_over = line[18]
                c += 1
        # add to the all_otdl for the final carrier after the last line of the file is read
        if good_id != "no":
            sql = "SELECT kb_name FROM name_index WHERE emp_id='%s'" % good_id
            result = inquire(sql)  # get the kb name with the emp id
            all_day_run = []  # gets the total hours for each day
            for i in range(7):
                all_day_run.append(day_run[i])
            to_add = ([result[0][0]] + all_day_run + [day_over])  # add name, daily totals, day over
            all_otdl.append(to_add)
        all_otdl.sort(key=itemgetter(0))  # sort the all otdl array by carrier name
        # define spreadsheet cell formats
        bd = Side(style='thin', color="80808080")  # defines borders
        ws_header = NamedStyle(name="ws_header", font=Font(bold=True, name='Arial', size=14))
        date_dov = NamedStyle(name="date_dov", font=Font(name='Arial', size=10))
        date_dov_title = NamedStyle(name="date_dov_title", font=Font(bold=True, name='Arial', size=10),
                                    alignment=Alignment(horizontal='right'))
        col_header = NamedStyle(name="col_header", font=Font(bold=True, name='Arial', size=10),
                                alignment=Alignment(horizontal='center'))
        col_name = NamedStyle(name="col_name", font=Font(bold=True, name='Arial', size=10),
                              alignment=Alignment(horizontal='left'))
        col_mod = NamedStyle(name="col_mod", font=Font(bold=True, name='Arial', size=10),
                             alignment=Alignment(horizontal='center'),
                             fill=PatternFill(fgColor='FFFFE0', fill_type='solid'),
                             border=Border(left=bd, top=bd, right=bd, bottom=bd))
        input_name = NamedStyle(name="input_name", font=Font(name='Arial', size=10),
                                border=Border(left=bd, top=bd, right=bd, bottom=bd))
        input_s = NamedStyle(name="input_s", font=Font(name='Arial', size=10),
                             border=Border(left=bd, top=bd, right=bd, bottom=bd),
                             alignment=Alignment(horizontal='right'))
        calcs = NamedStyle(name="calcs", font=Font(name='Arial', size=10),
                           border=Border(left=bd, top=bd, right=bd, bottom=bd),
                           fill=PatternFill(fgColor='FFFFE0', fill_type='solid'),
                           alignment=Alignment(horizontal='right'))
        wb = Workbook()  # define the workbook
        wkly_total = wb.active  # create first worksheet
        wkly_total.title = "over_60"  # title first worksheet

        wkly_total["A1"] = "Weekly Availability Summary"
        wkly_total["A1"].style = ws_header
        wkly_total.merge_cells('A1:E1')
        wkly_total['A3'] = "Date:  "  # create date/ pay period/ station header
        wkly_total['A3'].style = date_dov_title
        range_of_dates = format(g_date[0], "%A  %m/%d/%y") + " - " + format(g_date[6], "%A  %m/%d/%y")
        wkly_total['B3'] = range_of_dates
        wkly_total['B3'].style = date_dov
        wkly_total.merge_cells('B3:H3')
        date = datetime(int(gs_year), int(gs_mo), int(gs_day))
        pay_period = pp_by_date(date)
        wkly_total['E4'] = "Pay Period:  "
        wkly_total['E4'].style = date_dov_title
        wkly_total.merge_cells('E4:F4')
        wkly_total['G4'] = pay_period
        wkly_total['G4'].style = date_dov
        wkly_total.merge_cells('G4:H4')
        wkly_total['A4'] = "Station:  "
        wkly_total['A4'].style = date_dov_title
        wkly_total['B4'] = g_station
        wkly_total['B4'].style = date_dov
        wkly_total.merge_cells('B4:D4')
        oi = 6
        # column headers - first row
        wkly_total["A" + str(oi)] = "carrier name"  # carrier name
        wkly_total["B" + str(oi)] = "sat"
        wkly_total["C" + str(oi)] = "sun"
        wkly_total["D" + str(oi)] = "mon"
        wkly_total["E" + str(oi)] = "tue"
        wkly_total["F" + str(oi)] = "wed"
        wkly_total["G" + str(oi)] = "thr"
        wkly_total["H" + str(oi)] = "fri"
        wkly_total["I" + str(oi)] = "day over"  # the day of the violation
        # column headers - second row
        wkly_total["B" + str(oi + 1)] = "cumulative totals"
        wkly_total.merge_cells('B7:H7')
        wkly_total["I" + str(oi + 1)] = "to 60"  # the day of the violation
        # format headers
        wkly_total["A" + str(oi)].style = col_name
        wkly_total["B" + str(oi)].style = col_header
        wkly_total["C" + str(oi)].style = col_header
        wkly_total["D" + str(oi)].style = col_header
        wkly_total["E" + str(oi)].style = col_header
        wkly_total["F" + str(oi)].style = col_header
        wkly_total["G" + str(oi)].style = col_header
        wkly_total["H" + str(oi)].style = col_header
        wkly_total["I" + str(oi)].style = col_header
        wkly_total["B" + str(oi + 1)].style = col_mod
        wkly_total["I" + str(oi + 1)].style = col_mod
        # column widths
        wkly_total.column_dimensions["A"].width = 18
        wkly_total.column_dimensions["B"].width = 7
        wkly_total.column_dimensions["C"].width = 7
        wkly_total.column_dimensions["D"].width = 7
        wkly_total.column_dimensions["E"].width = 7
        wkly_total.column_dimensions["F"].width = 7
        wkly_total.column_dimensions["G"].width = 7
        wkly_total.column_dimensions["H"].width = 7
        wkly_total.column_dimensions["I"].width = 10
        oi += 2
        for otdl in all_otdl:
            # first of two rows
            wkly_total["A" + str(oi)] = otdl[0]  # carrier name
            wkly_total["B" + str(oi)] = otdl[1]
            wkly_total["C" + str(oi)] = otdl[2]
            wkly_total["D" + str(oi)] = otdl[3]
            wkly_total["E" + str(oi)] = otdl[4]
            wkly_total["F" + str(oi)] = otdl[5]
            wkly_total["G" + str(oi)] = otdl[6]
            wkly_total["H" + str(oi)] = otdl[7]
            if otdl[8] == "empty":  # handle "empty" violation days
                violation_day = ""
            else:
                violation_day = otdl[8]
            wkly_total["I" + str(oi)] = violation_day  # the day of the violation
            # format each cell with style
            wkly_total["A" + str(oi)].style = input_name
            wkly_total["B" + str(oi)].style = input_s
            wkly_total["C" + str(oi)].style = input_s
            wkly_total["D" + str(oi)].style = input_s
            wkly_total["E" + str(oi)].style = input_s
            wkly_total["F" + str(oi)].style = input_s
            wkly_total["G" + str(oi)].style = input_s
            wkly_total["H" + str(oi)].style = input_s
            wkly_total["I" + str(oi)].style = input_s
            # set number format for each cell
            wkly_total["B" + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
            wkly_total["C" + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
            wkly_total["D" + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
            wkly_total["E" + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
            wkly_total["F" + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
            wkly_total["G" + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
            wkly_total["H" + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
            # second of two rows - incluces running totals
            formula = "=%s!B%s" % ('over_60', str(oi))
            wkly_total["B" + str(oi + 1)] = formula
            formula = "=SUM(%s!C%s+%s!B%s)" % ('over_60', str(oi), 'over_60', str(oi + 1))
            wkly_total["C" + str(oi + 1)] = formula
            formula = "=SUM(%s!D%s+%s!C%s)" % ('over_60', str(oi), 'over_60', str(oi + 1))
            wkly_total["D" + str(oi + 1)] = formula
            formula = "=SUM(%s!E%s+%s!D%s)" % ('over_60', str(oi), 'over_60', str(oi + 1))
            wkly_total["E" + str(oi + 1)] = formula
            formula = "=SUM(%s!F%s+%s!E%s)" % ('over_60', str(oi), 'over_60', str(oi + 1))
            wkly_total["F" + str(oi + 1)] = formula
            formula = "=SUM(%s!G%s+%s!F%s)" % ('over_60', str(oi), 'over_60', str(oi + 1))
            wkly_total["G" + str(oi + 1)] = formula
            formula = "=SUM(%s!H%s+%s!G%s)" % ('over_60', str(oi), 'over_60', str(oi + 1))
            wkly_total["H" + str(oi + 1)] = formula
            formula = "=MAX(60-%s!H%s,0)" % ('over_60', str(oi + 1))
            wkly_total["I" + str(oi + 1)] = formula
            # format each cell of the second row
            wkly_total["B" + str(oi + 1)].style = calcs
            wkly_total["C" + str(oi + 1)].style = calcs
            wkly_total["D" + str(oi + 1)].style = calcs
            wkly_total["E" + str(oi + 1)].style = calcs
            wkly_total["F" + str(oi + 1)].style = calcs
            wkly_total["G" + str(oi + 1)].style = calcs
            wkly_total["H" + str(oi + 1)].style = calcs
            wkly_total["I" + str(oi + 1)].style = calcs
            # set number format for each cell
            wkly_total["B" + str(oi + 1)].number_format = "#,###.00;[RED]-#,###.00"
            wkly_total["C" + str(oi + 1)].number_format = "#,###.00;[RED]-#,###.00"
            wkly_total["D" + str(oi + 1)].number_format = "#,###.00;[RED]-#,###.00"
            wkly_total["E" + str(oi + 1)].number_format = "#,###.00;[RED]-#,###.00"
            wkly_total["F" + str(oi + 1)].number_format = "#,###.00;[RED]-#,###.00"
            wkly_total["G" + str(oi + 1)].number_format = "#,###.00;[RED]-#,###.00"
            wkly_total["H" + str(oi + 1)].number_format = "#,###.00;[RED]-#,###.00"
            wkly_total["I" + str(oi + 1)].number_format = "#,###.00;[RED]-#,###.00"
            oi += 2
        if len(not_indexed) > 0:
            wkly_total["A" + str(oi)] = "Carriers not included (not in name index):"
            wkly_total.merge_cells('A' + str(oi) + ':D' + str(oi))
            oi += 1
            for name in not_indexed:
                wkly_total['A' + str(oi)] = name
                wkly_total.merge_cells('A' + str(oi) + ':D' + str(oi))
                oi += 1
        # name the excel file
        xl_filename = "kb_wa" + str(format(g_date[0], "_%y_%m_%d")) + ".xlsx"
        ok = messagebox.askokcancel("Spreadsheet generator", "Do you want to generate a spreadsheet?")
        if ok == True:
            if os.path.isdir('kb_sub/weekly_availability') == False:
                os.makedirs('kb_sub/weekly_availability')
            try:
                wb.save('kb_sub/weekly_availability/' + xl_filename)
                messagebox.showinfo("Spreadsheet generator", "Your spreadsheet was successfully generated. \n"
                                                             "File is named: {}".format(xl_filename))
                if sys.platform == "win32":
                    os.startfile('kb_sub\\weekly_availability\\' + xl_filename)
                if sys.platform == "linux":
                    subprocess.call(["xdg-open", 'kb_sub/weekly_availability/' + xl_filename])
                if sys.platform == "darwin":
                    subprocess.call(["open", 'kb_sub/weekly_availability/' + xl_filename])
            except:
                messagebox.showerror("Spreadsheet generator", "The spreadsheet was not generated. \n"
                                                              "Suggestion: "
                                                              "Make sure that identically named spreadsheets are closed "
                                                              "(the file can't be overwritten while open).")
        frame.destroy()
        main_frame()


def station_rec_del(self, tacs, kb):
    sql = "DELETE FROM station_index WHERE tacs_station = '%s' and kb_station='%s'" % (tacs, kb)
    commit(sql)
    self.destroy()
    station_index_mgmt("none")


def station_index_rename_apply(self, tacs, newname):
    sql = "UPDATE station_index SET kb_station='%s' WHERE tacs_station='%s'" % (newname.get(), tacs)
    commit(sql)
    station_index_mgmt(self)


def station_index_rename(self, frame, tacs, kb, newname, button, all_stations):
    button.destroy()
    Button(frame, text=" ", width=6).grid(row=0, column=2)
    if len(all_stations) > 0:
        Label(frame, text="update station name:  ", anchor="e").grid(row=1, column=0, sticky="e")
        # set up station option menu and variable
        om_station = OptionMenu(frame, newname, *all_stations)
        om_station.config(width=28, anchor="w")
        om_station.grid(row=1, column=1)
        newname.set(kb)
        Button(frame, text="rename", command=lambda: station_index_rename_apply(self, tacs, newname)).grid(row=1,
                                                                                                           column=2)
    else:
        Label(frame, text="No Unassigned Stations Available").grid(row=1, column=0, columnspan=2, sticky="e")


def stationindexer_del_all(self):
    sql = "DELETE FROM station_index"
    commit(sql)
    station_index_mgmt("none")


def station_index_mgmt(self):
    wd = front_window(self)  # get window objects 0=F,1=S,2=C,3=FF,4=buttons
    g = 0
    Label(wd[3], text="Station Index Management", font="bold").grid(row=g, column=0, sticky="w")
    Label(wd[3], text="").grid(row=g + 1, column=0)
    g += 2
    all_stations = []
    sql = "SELECT * FROM stations"
    results = inquire(sql)
    for rec in results:
        all_stations.append(rec[0])
    sql = "SELECT * FROM station_index"
    results = inquire(sql)
    for rec in results:
        if rec[1] in all_stations:
            all_stations.remove(rec[1])
    all_stations.remove("out of station")
    if len(results) == 0:
        Label(wd[3], text="There are no stations in the station index").grid(row=g, column=0, sticky="w")
        g += 1
    else:
        header_frame = Frame(wd[3], width=500)
        header_frame.grid(row=g, column=0, sticky="w")
        Label(header_frame, text="TACS Station Name", width="30", anchor="w").grid(row=0, column=0, sticky="w")
        Label(header_frame, text="Klusterbox Station Name", width="30", anchor="w").grid(row=0, column=1, sticky="w")
        g += 1
        f = 0  # initialize number for frame
        frame = []  # initialize array for frame
        si_newname = []
        rename_button = []
        for record in results:
            to_add = "station_frame" + str(f)  # give the new frame a name
            frame.append(to_add)  # add the frame to the array
            frame[f] = Frame(wd[3], width=500)  # create the frame widget
            frame[f].grid(row=g, padx=5, sticky="w")  # grid the widget
            si_newname.append(StringVar(wd[0]))
            Button(frame[f], text=record[0], width=30, anchor="w").grid(row=0, column=0)
            Button(frame[f], text=record[1], width=30, anchor="w").grid(row=0, column=1)
            to_add = Button(frame[f], text="rename", width=6)
            rename_button.append(to_add)
            rename_button[f]['command'] = \
                lambda frame=frame[f], tacs=record[0], kb=record[1], newname=si_newname[f],button=rename_button[f]: \
                station_index_rename(wd[0], frame, tacs, kb, newname, button, all_stations)
            rename_button[f].grid(row=0, column=2)
            delete_button = Button(frame[f], text="delete", width=6,
                                   command=lambda tacs=record[0], kb=record[1]: station_rec_del(wd[0], tacs, kb))
            delete_button.grid(row=0, column=3)
            f += 1
            g += 1
        Button(wd[3], text="Delete All", width="15", command=lambda: (wd[0].destroy(), stationindexer_del_all(wd[0]))) \
            .grid(row=g, column=0, columnspan=3, sticky="e")
    Button(wd[4], text="Go Back", width=20, anchor="w",
           command=lambda: (wd[0].destroy(), main_frame())).pack(side=LEFT)
    rear_window(wd)


def apply_nameindexer_list(self, x):
    sql = "DELETE FROM name_index WHERE emp_id = '%s'" % x
    commit(sql)
    self.destroy()
    name_index_screen()


def del_all_nameindexer(self):
    sql = "DELETE FROM name_index"
    commit(sql)
    self.destroy()
    name_index_screen()


def name_index_screen():
    sql = "SELECT * FROM name_index ORDER BY tacs_name"
    results = inquire(sql)
    wd = front_window("none")  # get window objects
    x = 0
    if len(results) == 0:
        Label(wd[3], text="The Name Index is empty").grid(row=0, column=x)
    else:
        Label(wd[3], text="Name Index Management", font="bold").grid(row=x, column=0, sticky="w",
                                                                     columnspan=2)  # page header
        x += 1
        Label(wd[3], text="").grid(row=x, column=0, sticky="w")
        x += 1
        Label(wd[3], text="TACS Name").grid(row=x, column=1, sticky="w")  # column headers
        Label(wd[3], text="Klusterbox Name").grid(row=x, column=2, sticky="w")
        Label(wd[3], text="Emp ID").grid(row=x, column=3, sticky="w")
        x += 1
        for item in results:  # loop for names in the index
            Label(wd[3], text=str(x - 2), anchor="w").grid(row=x, column=0)
            Button(wd[3], text=" " + item[0], anchor="w", width=20, relief=RIDGE).grid(row=x, column=1)
            Button(wd[3], text=" " + item[1], anchor="w", width=20, relief=RIDGE).grid(row=x, column=2)
            Button(wd[3], text=" " + item[2], anchor="w", width=8, relief=RIDGE).grid(row=x, column=3)
            Button(wd[3], text="delete", anchor="w", width=5, relief=RIDGE, command=lambda x=item[2]:
            apply_nameindexer_list(wd[0], x)).grid(row=x, column=4)
            x += 1
        Button(wd[3], text="Delete All", width="15", command=lambda: del_all_nameindexer(wd[0])) \
            .grid(row=x, column=0, columnspan=5, sticky="e")
    Button(wd[4], text="Go Back", width=20, command=lambda: (wd[0].destroy(), main_frame())).pack(side=LEFT)
    wd[0].update()
    wd[2].config(scrollregion=wd[2].bbox("all"))
    mainloop()


def gen_ns_dict(file_path, to_addname):  # creates a dictionary of ns days
    days = ("Saturday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday")
    mv_codes = ("BT", "MV", "ET")
    good_jobs = ("134", "844")
    results = []
    carrier = []
    id_bank = []
    aux_list = []
    for id in to_addname:
        id_bank.append(id[0].zfill(8))
        if id[3] == "auxiliary": aux_list.append(id[0].zfill(8))  # make an array of auxiliary carrier emp ids
    with open(file_path, newline="") as file:
        a_file = csv.reader(file)
        good_id = "no"
        for line in a_file:
            if len(line) > 4:
                if good_id != line[4].zfill(8) and good_id != "no":  # if new carrier or employee
                    if good_id in aux_list:
                        day = "None"  # ignore auxiliary carriers
                    else:
                        day = ee_ns_detect(carrier)  # process regular carriers
                    to_add = (good_id, day)
                    results.append(to_add)
                    del carrier[:]  # empty array
                    good_id = "no"  # reset trigger
                if line[18] == "Base" and line[19] in good_jobs and line[4].zfill(
                        8) in id_bank:  # find first line of specific carrier
                    good_id = line[4].zfill(8)  # set trigger to id of carriers who are FT or aux carriers
                    carrier.append(line)  # gather times and moves for anaylsis
                if good_id == line[4].zfill(8) and line[18] != "Base":
                    if line[18] in days:  # get the hours for each day
                        carrier.append(line)  # gather times and moves for anaylsis
                    if line[19] in mv_codes and line[32] != "(W)Ring Deleted From PC":
                        carrier.append(line)  # gather times and moves for anaylsis
        if good_id != "no":
            if good_id in aux_list:
                day = "None"  # ignore auxiliary carriers
            else:
                day = ee_ns_detect(carrier)  # process regular carriers
            to_add = (good_id, day)
            results.append(to_add)
        del carrier[:]  # empty array
        return (results)


def auto_precheck():
    # delete any records from name index which don't have corresponding records in carriers table
    sql = "SELECT kb_name FROM name_index"
    kb_name = inquire(sql)
    sql = "SELECT carrier_name FROM carriers"
    results = inquire(sql)
    carriers = []
    for item in results:
        if item not in carriers: carriers.append(item)
    count = 0
    # create progressbar
    pb_root = Tk()  # create a window for the progress bar
    pb_root.geometry("%dx%d+%d+%d" % (500, 50, 200, 300))
    pb_root.title("Database Maintenance")
    pb_label = Label(pb_root, text="Updating Changes: ")  # make label for progress bar
    pb_label.pack(side=LEFT)
    pb = ttk.Progressbar(pb_root, length=400, mode="determinate")  # create progress bar
    pb.pack(side=LEFT)
    pb["maximum"] = len(kb_name)  # set length of progress bar
    pb.start()
    i = 0
    for name in kb_name:
        pb["value"] = i  # increment progress bar
        if name not in carriers:
            sql = "DELETE FROM name_index WHERE kb_name = '%s'" % name
            commit(sql)
            count += 1
        pb_root.update()
        i += 1
    pb.stop()  # stop and destroy the progress bar
    pb_label.destroy()  # destroy the label for the progress bar
    pb.destroy()
    pb_root.destroy()


def gen_carrier_list():
    # generate in range carrier list
    if g_range == "week":  # select sql dependant on range
        sql = "SELECT effective_date, carrier_name,list_status, ns_day,route_s, station, rowid" \
              " FROM carriers WHERE effective_date <= '%s'" \
              "ORDER BY carrier_name, effective_date desc" % (g_date[6])
    if g_range == "day":
        sql = "SELECT effective_date, carrier_name,list_status, ns_day,route_s, station, rowid" \
              " FROM carriers WHERE effective_date <= '%s'" \
              "ORDER BY carrier_name, effective_date desc" % (d_date)
    results = inquire(sql)  # call function to access database
    carrier_list = []  # initialize arrays for data sorting
    candidates = []
    more_rows = []
    pre_invest = []
    for i in range(len(results)):  # take raw data and sort into appropriate arrays
        candidates.append(results[i])  # put name into candidates array
        jump = "no"  # triggers an analysis of the candidates array
        if i != len(results) - 1:  # if the loop has not reached the end of the list
            if results[i][1] == results[i + 1][1]:  # if the name current and next name are the same
                jump = "yes"  # bypasses an analysis of the candidates array
        if jump == "no":
            # sort into records in investigation range and those prior
            for record in candidates:
                if g_range == "week":  # if record falls in investigation range - add it to more rows array
                    if record[0] >= str(g_date[1]) and record[0] <= str(g_date[6]): more_rows.append(record)
                    if record[0] <= str(g_date[0]) and len(pre_invest) == 0: pre_invest.append(record)
                if g_range == "day":
                    if record[0] <= str(d_date) and len(pre_invest) == 0: pre_invest.append(record)
            # find carriers who start in the middle of the investigation range CATEGORY ONE
            if len(more_rows) > 0 and len(pre_invest) == 0:
                station_anchor = "no"
                for each in more_rows:  # check if any records place the carrier in the selected station
                    if each[5] == g_station: station_anchor = "yes"  # if so, set the station anchor
                if station_anchor == "yes":
                    list(more_rows)
                    for each in more_rows:
                        x = list(each)  # convert the tuple to a list
                        carrier_list.append(x)  # add it to the list
            # find carriers with records before and during the investigation range CATEGORY TWO
            if len(more_rows) > 0 and len(pre_invest) > 0:
                station_anchor = "no"
                for each in more_rows + pre_invest:
                    if each[5] == g_station: station_anchor = "yes"
                if station_anchor == "yes":
                    xx = list(pre_invest[0])
                    carrier_list.append(xx)
            # find carrier with records from only before investigation range.CATEGORY THREE
            if len(more_rows) == 0 and len(pre_invest) == 1:
                for each in pre_invest:
                    if each[5] == g_station:
                        x = list(pre_invest[0])
                        carrier_list.append(x)
            del more_rows[:]
            del pre_invest[:]
            del candidates[:]
    return carrier_list


def gen_nameindex_dict():
    sql = "SELECT tacs_name, kb_name, emp_id FROM name_index ORDER BY kb_name"
    results = inquire(sql)
    n_dict = {}
    for line in results:  # loop to fill arrays
        n_dict[line[2]] = line[1]
    return n_dict


def auto_indexer_1(self, file_path):  # pair station from tacs to correct station in klusterbox/ part 1
    auto_precheck()
    with open(file_path, newline="") as file:
        a_file = csv.reader(file)
        c = 0
        for line in a_file:
            if c == 0 and line[0][:8] != "TAC500R3":
                messagebox.showwarning("File Selection Error", "The selected file does not appear to be an "
                                                               "Employee Everything report.")
                return
            if c == 3:
                tacs_pp = line[0]  # find the pay period
                tacs_station = line[2]  # find the station
                break
            c += 1
        c = 0
        range_days = ["Saturday", "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
        for line in a_file:  # find the range
            if line[18] in range_days:
                range_days.remove(line[18])
            if c == 150: break  # survey 150 lines before breaking to anaylize results.
            c += 1
        if len(range_days) > 5:
            t_range = "day"  # set the range
            messagebox.showwarning("File Selection Error", "Employee Everything Reports that cover only one day /n"
                                                           "are not supported in version 3.002 of Klusterbox.")
            return
        else:
            t_range = "week"
    year = int(tacs_pp[:-3])
    pp = tacs_pp[-3:]
    t_date = find_pp(year, pp)
    sql = "SELECT tacs_station, kb_station, finance_num FROM station_index"
    results = inquire(sql)
    station_index = []  # create a list of klusterbox names
    tacs_index = []
    for line in results:
        station_index.append(line[1])
        tacs_index.append(line[0])
    sql = "SELECT station FROM stations"
    results = inquire(sql)
    kb_stations = []
    for record in results:
        kb_stations.append(record[0])
    self.destroy()
    F = Frame(root)
    F.pack(fill=BOTH, side=LEFT)
    S = Scrollbar(F)  # link up the canvas and scrollbar
    C = Canvas(F, width=1600)
    S.pack(side=RIGHT, fill=BOTH)
    C.pack(side=LEFT, fill=BOTH, pady=10, padx=10)
    S.configure(command=C.yview, orient="vertical")
    C.configure(yscrollcommand=S.set)
    if sys.platform == "win32":
        C.bind_all('<MouseWheel>', lambda event: C.yview_scroll(int(-1 * (event.delta / 120)), "units"))
    elif sys.platform == "linux":
        C.bind_all('<Button-4>', lambda event: C.yview('scroll', -1, 'units'))
        C.bind_all('<Button-5>', lambda event: C.yview('scroll', 1, 'units'))
    FF = Frame(C)  # create the frame inside the canvas
    C.create_window((0, 0), window=FF, anchor=NW)
    possible_stations = []
    for item in kb_stations:
        possible_stations.append(item)
    if len(tacs_station) == 0:
        messagebox.showwarning("Auto Data Entry Error",
                               "The Employee Everything Report is corrupt. Data Entry will stop.  \n"
                               "The Employee Everything Report does not include "
                               "information about the station. This could be caused by an error of the pdf "
                               "converter. If you can obtain an Employee Everything Report from management in "
                               "csv format, you should have better results.")
        F.destroy()
        main_frame()
        return

    station_index.append("out of station")
    possible_stations = [x for x in possible_stations if x not in station_index]
    Label(FF, text="Station Pairing", font="bold", pady=10).grid(row=0, column=0, columnspan=4,
                                                                 sticky=W)  # page contents
    Label(FF, text="Match the station detected from TACS with a pre-existing station\n "
                   "or use ADD STATION to add the station if there isn't a match.", justify=LEFT) \
        .grid(row=1, column=0, columnspan=4, sticky=W)
    Label(FF, text="Detected Station: ", anchor="w").grid(row=2, column=0, sticky="w")
    Label(FF, text=tacs_station, fg="blue").grid(row=3, column=0, columnspan=4)
    Label(FF, text="Select Station: ", anchor="w").grid(row=4, column=0, sticky=W)
    station_sorter = StringVar(FF)
    station_options = ["select matching station"] + possible_stations +["ADD STATION"]
    station_sorter.set(station_options[0])
    option_menu = OptionMenu(FF, station_sorter, *station_options)
    option_menu.config(width=30)
    option_menu.grid(row=5, column=0, columnspan=2, sticky=W)
    Label(FF, text=" ", justify=LEFT).grid(row=6, column=0, sticky=W)
    Label(FF, text="If the station is not present in the drop down menu, select  \n "
                   "ADD STATION from the menu and enter the new station name \n"
                   "below to pair it with the station originating the report", justify=LEFT) \
                    .grid(row=7, column=0, columnspan=4, sticky=W)
    Label(FF, text=" ", justify=LEFT).grid(row=8, column=0, sticky=W)
    Label(FF, text="Enter New Station Name: ", anchor="w").grid(row=9, column=0, columnspan=4, sticky=W)
    # insert entry for station name
    station_new = StringVar(FF)
    Entry(FF, width=35, textvariable=station_new).grid(row=10, column=0, columnspan=4, sticky=W)
    Label(FF, text=" ", justify=LEFT).grid(row=11, column=0, sticky=W)
    Button(FF, text="OK", width=8, command=lambda: apply_auto_indexer_1
    (F, file_path, tacs_station, station_sorter.get(), station_new.get(), t_date, t_range)).grid(row=12, column=2, sticky=W)
    Button(FF, text="Cancel", width=8, command=lambda: (F.destroy(), main_frame())).grid(row=12, column=3, sticky=W)
    if tacs_station in tacs_index:
        auto_indexer_2(F, file_path, t_date, tacs_station, t_range)
    else:
        C.config(scrollregion=C.bbox("all"))
        root.update()
        mainloop()


def apply_auto_indexer_1(self, file_path, tacs_station, station_sorter, station_new,  t_date, t_range):
    sql = "SELECT kb_station FROM station_index"
    result = inquire(sql)
    station_index = []
    for s in result:
        station_index.append(s[0])
    global list_of_stations
    station_new = station_new.strip()
    if station_sorter == "select matching station":
        messagebox.showerror("Data Entry Error", "You must select a station or ADD STATION", parent=self)
        return
    elif station_sorter == "ADD STATION" and station_new == "":
        messagebox.showerror("Data Entry Error", "You must provide a name for the new station.", parent=self)
        return
    elif station_sorter == "ADD STATION" and station_new != "":
        if station_new not in list_of_stations:
            sql = "INSERT INTO stations (station) VALUES('%s')" % (station_new)
        commit(sql)
        if station_new not in list_of_stations:
            list_of_stations.append(station_new)
        if len(tacs_station) != 0: # add to the station index to the dbase unless tacs_station is empty.
            sql = "INSERT INTO station_index (tacs_station, kb_station, finance_num) VALUES('%s','%s','%s')" \
                  % (tacs_station, station_new, "")
            commit(sql)
        messagebox.showinfo("Database Updated", "The {} station has been added to the list of stations automatically "
                                                "recognized.".format(station_new))
    elif station_sorter != "ADD STATION" and station_new != "":
        messagebox.showerror("Data Entry Error", "You can not select a station from the drop down menu AND enter "
                                                 "a station in the text field.")
        return
    else:
        sql = "INSERT INTO station_index (tacs_station, kb_station, finance_num) VALUES('%s','%s','%s')" \
              % (tacs_station, station_sorter, "")
        commit(sql)
        messagebox.showinfo("Database Updated",
                            "The {} station has been paired to the {} station. In the future, this association "
                            "will be automatically recognized.".format(tacs_station, station_sorter))
    auto_indexer_2(self, file_path, t_date, tacs_station, t_range)


def auto_indexer_2(self, file_path, t_date, tacs_station, t_range):  # Pairing screen #1
    s_year = t_date.strftime("%Y")
    s_mo = t_date.strftime("%m")
    s_day = t_date.strftime("%d")
    sql = "SELECT kb_station FROM station_index WHERE tacs_station = '%s'" % tacs_station
    station = inquire(sql)
    set_globals(s_year, s_mo, s_day, t_range, station[0][0], "None")
    sql = "SELECT tacs_name, kb_name, emp_id FROM name_index ORDER BY kb_name"
    results = inquire(sql)
    name_index = []  # create a list of klusterbox names
    id_index = []  # create a list of emp ids
    for line in results:
        name_index.append(line[1])
        id_index.append(line[2].zfill(8))
    carrier_list = gen_carrier_list()  # generate an in range carrier list
    c_list = []  # create a list of unique names from carrier list (a set)
    for each in carrier_list:
        if each[1] not in c_list: c_list.append(each[1])
    # Get the names from tacs report
    tacs_list = []
    good_jobs = ("134", "844")
    with open(file_path, newline="") as file:
        a_file = csv.reader(file)
        c = 0
        for line in a_file:
            if c > 1 and line[19] in good_jobs:
                # create a note for carrier's assignment - reg w/route, reg floater or aux
                route = line[25].zfill(6)
                lvl = line[23].zfill(2)
                if line[19] == "134" and lvl == "01":
                    assignment = "reg " + route[1] + route[2] + route[4] + route[5]
                elif line[19] == "134" and lvl == "02":
                    assignment = "reg " + "floater"
                elif line[19] == "844":
                    assignment = "auxiliary"
                else:
                    assignment = "undetected"
                lastname = line[5].lower().replace("\'"," ")
                add_to_list = [line[4].zfill(8), lastname, line[6].lower(),
                               assignment]  # create list to insert in list
                tacs_list.append(add_to_list)
            c += 1
    holder = ["", "", "", ""]  # find the duplicates and remove them where there is both BASE and TEMP
    to_remove = []
    put_back = []
    for item in tacs_list:  # crawler goes down the list to identify Temp entries
        if item[0] == holder[0]:
            if item == holder:
                to_remove.append(holder)  # remove both records
            if item != holder:
                to_remove.append(holder)
                to_remove.append(item)
            put_back.append(item)  # put the later record back in the list
        holder = item  # hold the record to compare in the next loop
    tacs_list = [x for x in tacs_list if x not in to_remove]  # remove the duplicates
    for record in put_back:  # put the Temp record back into the tacs_list
        tacs_list.append(record)
    tacs_list.sort(key=itemgetter(1))  # re-alphabetize the list of carriers
    add = 0  # create tallies for reports
    rec = 0
    out = 0
    to_remove = []  # carriers who are already or newly placed in name index - remove them from further processing
    new_carrier = []  # new carriers who have duplicate names send these to auto indexer 6
    dup_array = []
    check_these = []
    # create progressbar
    pb_root = Tk()  # create a window for the progress bar
    pb_root.geometry("%dx%d+%d+%d" % (500, 50, 200, 300))
    pb_root.title("Database Maintenance")
    pb_label = Label(pb_root, text="Updating Changes: ")  # make label for progress bar
    pb_label.pack(side=LEFT)
    pb = ttk.Progressbar(pb_root, length=400, mode="determinate")  # create progress bar
    pb.pack(side=LEFT)
    pb["maximum"] = len(tacs_list)  # set length of progress bar
    pb.start()
    i = 0
    for each in tacs_list:
        pb["value"] = i  # increment progress bar
        tac_str = "{}, {}".format(each[1], each[2])  # tac str is last name and first initial from tacs report
        if tac_str in c_list and each[0] not in id_index:  # if there is an identical match between kb and tacs names:
            if tac_str in name_index:  # if there is a dup name / need a complete list of carrier names from index
                new_carrier.append(each)  # maybe just pass information via new_carrier and add later
            else:  # go ahead and pair the emp id with the name in carriers
                sql = "INSERT INTO name_index (tacs_name, kb_name, emp_id ) VALUES('%s','%s','%s')" \
                      % (tac_str, tac_str, each[0])
                name_index.append(tac_str)
                id_index.append(each[0])
            add += 1
            commit(sql)
            to_remove.append(each[0])
            name_index.append(tac_str)
        elif each[0] in id_index:  # RECOGNIZED -  the emp id is already in the name index
            to_remove.append(each[0])
            check_these.append(each)
            rec += 1
        else:
            out += 1
        pb_root.update()  # update the progress bar
        i += 1
    pb.stop()  # stop and destroy the progress bar
    pb_label.destroy()  # destroy the label for the progress bar
    pb.destroy()
    pb_root.destroy()  # destroy the progress bar
    # find the carriers in name_index who have records w/ eff dates in the future
    dont_check = []  # remove items from check these if future carriers are found
    for name in check_these:
        sql = "SELECT kb_name FROM name_index WHERE emp_id = '%s'" % name[0]
        result = inquire(sql)
        kb_name = result[0][0]
        sql = "SELECT effective_date,carrier_name FROM carriers WHERE carrier_name = '%s' AND effective_date <= '%s' " \
              "ORDER BY effective_date DESC" % (kb_name, g_date[0])
        result = inquire(sql)
        if not result:
            new_carrier.append(name)  # will add as new carrier in AI 3
            dont_check.append(name[0])  # removes from check these array
            to_remove.append(name[0])  # removes from tacs list
    check_these = [x for x in check_these if x[0] not in dont_check]  # removes don't check from check these
    """
    messagebox.showinfo("Processing Carriers", "{} Carrier names were added to the database\n"
                                               "{} Carrier names were recognized as pre-existing in the database.\n"
                                               "{} Carrier names have not been handled."
                                                .format(add, rec, out))
    """
    tacs_list = [x for x in tacs_list if x[0] not in new_carrier]
    tacs_list = [x for x in tacs_list if x[0] not in to_remove]
    sql = "SELECT tacs_name, kb_name, emp_id FROM name_index ORDER BY kb_name"
    results = inquire(sql)
    name_sorter = []
    tried_names = []
    for item in name_index: tried_names.append(item)
    name_index = []  # create a list of klusterbox names
    for line in results:
        name_index.append(line[1])
    # route to appropriate function based on array contents
    if len(tacs_list) < 1 and len(new_carrier) < 1 and len(check_these) < 1:  # all tacs list resolved/ nothing to check
        auto_indexer_6(self, file_path)  # to straight to entering rings
    elif len(tacs_list) < 1 and len(new_carrier) > 0:  # all tacs list resolved/ new names unresolved
        auto_indexer_4(self, file_path, new_carrier, check_these)  # add new carriers in AI6
    elif len(tacs_list) < 1 and len(new_carrier) < 1 and len(
            check_these) > 0:  # tacs and new carriers resolved/ carriers to check
        auto_indexer_5(self, file_path, check_these)  # step to AI  to check discrepancies
    else:  # If there are candidates sort, generate PAIRING SCREEN 1
        self.destroy()
        F = Frame(root)
        F.pack(fill=BOTH, side=LEFT)
        C1 = Canvas(F)
        C1.pack(fill=BOTH, side=BOTTOM),
        Button(C1, text="Continue", width=8, command=lambda: auto_indexer_3
        (F, file_path, tacs_list, name_sorter, tried_names, new_carrier, check_these)).grid(row=0, column=0)
        Button(C1, text="Cancel", width=8, command=lambda: (F.destroy(), main_frame())).grid(row=0, column=1)
        S = Scrollbar(F)  # link up the canvas and scrollbar
        C = Canvas(F, width=1600)
        S.pack(side=RIGHT, fill=BOTH)
        C.pack(side=LEFT, fill=BOTH, pady=10, padx=10)
        S.configure(command=C.yview, orient="vertical")
        C.configure(yscrollcommand=S.set)
        if sys.platform == "win32":
            C.bind_all('<MouseWheel>', lambda event: C.yview_scroll(int(-1 * (event.delta / 120)), "units"))
        elif sys.platform == "linux":
            C.bind_all('<Button-4>', lambda event: C.yview('scroll', -1, 'units'))
            C.bind_all('<Button-5>', lambda event: C.yview('scroll', 1, 'units'))
        FF = Frame(C)  # create the frame inside the canvas
        C.create_window((0, 0), window=FF, anchor=NW)
        c_list = [x for x in c_list if x not in name_index]
        Label(FF, text="Search for Name Matches #1", font="bold", pady=10).grid(row=0, column=0, sticky="w",
                                                                                columnspan=10)  # page contents
        Label(FF, text=
        "Look for possible matches for each unrecognized name. If the name has already been entered manually, you \n"
        "should be able to find it on this screen or the next. It is possible that the name has no match, if that is \n"
        "the case then select \"ADD NAME\" in the next screen. You can change the default between \"NOT FOUND\" and \n"
        "\"DISCARD\" using the buttons below. Information from TACS is shown in blue\n\n"
        "Investigation Range: {0} through {1}\n\n".format(g_date[0].strftime("%a - %b %d, %Y"),
                                                          g_date[6].strftime("%a - %b %d, %Y")), justify=LEFT) \
            .grid(row=1, column=0, columnspan=10, sticky="w")
        Button(FF, text="DISCARD", width=10, command=lambda: indexer_default(name_sorter, i + 1, name_options, 1)) \
            .grid(row=2, column=3, sticky="w", columnspan=2)
        Label(FF, text="switch default to DISCARD").grid(row=2, column=1, sticky="w", columnspan=2)
        Button(FF, text="NOT FOUND", width=10, command=lambda: indexer_default(name_sorter, i + 1, name_options, 0)) \
            .grid(row=3, column=3, sticky="w", columnspan=2)
        Label(FF, text="switch default to NOT FOUND").grid(row=3, column=1, sticky="w", columnspan=2)
        Label(FF, text="").grid(row=4, column=0)

        Label(FF, text="Name", fg="grey").grid(row="5", column="1", sticky="w")
        Label(FF, text="Assignment", fg="grey").grid(row="5", column="2", sticky="w")
        Label(FF, text="Candidates", fg="grey").grid(row="5", column="3", sticky="w")
        c = 6
        i = 0
        color = "blue"
        for t_name in tacs_list:
            possible_names = []
            Label(FF, text=str(i + 1), anchor="w").grid(row=c, column=0, sticky="w")
            Label(FF, text=t_name[1] + ", " + t_name[2], anchor="w", width=15, fg=color).grid(row=c, column=1,
                                                                                              sticky="w")  # name
            Label(FF, text=t_name[3], anchor="w", width=10, fg=color).grid(row=c, column=2, sticky="w")  # assignment
            # build option menu for unmatched tacs names
            for c_name in c_list:
                if c_name[0] == t_name[1][0]:
                    possible_names.append(c_name)
                    tried_names.append(c_name)
            name_options = ["NOT FOUND", "DISCARD"] + possible_names
            name_sorter.append(StringVar(FF))
            option_menu = OptionMenu(FF, name_sorter[i], *name_options)
            name_sorter[i].set(name_options[0])
            option_menu.config(width=15)
            option_menu.grid(row=c, column=3, sticky="w")  # possible matches
            if len(possible_names) == 1:  # display indicator for possible matches
                Label(FF, text=str(len(possible_names)) + " name").grid(row=c, column=4, sticky="w")
            if len(possible_names) > 1:
                Label(FF, text=str(len(possible_names)) + " names").grid(row=c, column=4, sticky="w")
            c += 1
            i += 1
        root.update()
        C.config(scrollregion=C.bbox("all"))
        mainloop()


def auto_indexer_3(self, file_path, tacs_list, name_sorter, tried_names, new_carrier, check_these):
    # apply pairing screen #1 and create pairing screen #2
    i = 0  # count iterations of loops
    dis = 0  # count of discarded items
    out = 0  # count of unresolved items
    pair = 0  # count of added items
    to_remove = []  # intialized array of names to be removed from tacs names
    not_found = []  # initialize array of names to be futher analyzed.
    to_nameindex = []  # initialize array of names to be be paired in name index
    for item in name_sorter:
        if item.get() == "DISCARD":
            to_remove.append(tacs_list[i][0])
            dis += 1
        elif item.get() == "NOT FOUND":
            not_found.append(tacs_list[i])
            out += 1
        else:
            to_add = [tacs_list[i], item.get()]
            to_nameindex.append(to_add)
            to_remove.append(tacs_list[i][0])
            check_these.append(tacs_list[i])
            pair += 1
        i += 1
    tacs_list = [x for x in tacs_list if x[0] not in to_remove]
    for item in to_nameindex:
        tac_str = "{}, {}".format(item[0][1], item[0][2])  # tac str is last name and first initial from tacs report
        sql = "INSERT INTO name_index (tacs_name, kb_name, emp_id) VALUES('%s','%s','%s')" \
              % (tac_str, item[1], item[0][0])
        commit(sql)
    """
# message screens to summerize output
    messagebox.showinfo("Processing Carriers", "{} Carrier names were paired to names in klusterbox\n"
                                               "{} Carrier names were discarded.\n"
                                               "{} Carrier names have not been handled."
                                                .format(pair, dis, out))
    """
    # build possible names for option menus
    sql = "SELECT kb_name FROM name_index"
    results = inquire(sql)
    name_index = []  # create a list of klusterbox names
    for line in results:
        name_index.append(line[0])
    sql = "SELECT carrier_name FROM carriers ORDER BY carrier_name"  # get all names from the carrier list
    results = inquire(sql)  # call function to access database
    c_list = []
    for item in results:
        if item[0] not in c_list and item[0] not in tried_names and item[0] not in name_index:
            c_list.append(item[0])
    name_sorter = []  # page contents
    self.destroy()  # destroy old frame and build new frame
    F = Frame(root)
    F.pack(fill=BOTH, side=LEFT)
    C1 = Canvas(F)
    C1.pack(fill=BOTH, side=BOTTOM)
    Button(C1, text="Continue", width=8, command=lambda: apply_auto_indexer_3
    (F, C1, file_path, tacs_list, name_sorter, new_carrier, c_list, check_these)).grid(row=0, column=0)
    Button(C1, text="Cancel", width=8, command=lambda: (F.destroy(), main_frame())).grid(row=0, column=1)
    S = Scrollbar(F)  # link up the canvas and scrollbar
    C = Canvas(F, width=1600)
    S.pack(side=RIGHT, fill=BOTH)
    C.pack(side=LEFT, fill=BOTH, pady=10, padx=20)
    S.configure(command=C.yview, orient="vertical")
    C.configure(yscrollcommand=S.set)
    if sys.platform == "win32":
        C.bind_all('<MouseWheel>', lambda event: C.yview_scroll(int(-1 * (event.delta / 120)), "units"))
    elif sys.platform == "linux":
        C.bind_all('<Button-4>', lambda event: C.yview('scroll', -1, 'units'))
        C.bind_all('<Button-5>', lambda event: C.yview('scroll', 1, 'units'))
    FF = Frame(C)  # create the frame inside the canvas
    C.create_window((0, 0), window=FF, anchor=NW)

    # route to functions conditional on arrays
    if len(tacs_list) < 1 and len(check_these) > 0:  # if empty tacs list and something in check these
        auto_indexer_5(F, file_path, check_these)
    elif len(tacs_list) < 1 and len(check_these) < 1:
        auto_indexer_6(F, file_path)
    else:
        Label(FF, text="Search for Name Matches #2", font="bold", pady=10).grid(row=0, column=0, sticky="w",
                                                                                columnspan=10)  # page contents
        Label(FF, text=
        "Look for possible matches for each unrecognized name. If the name has already been entered manually, \n"
        " you should be able to find it on this screen. It is possible that the name has no match, if that is \n"
        "the case then select \"ADD NAME\" in this screen. You can change the default between \"ADD NAME\" and \n"
        "\"DISCARD\" using the buttons below. Information from TACS is shown in blue\n\n"
        "Investigation Range: {0} through {1}\n\n".format(g_date[0].strftime("%a - %b %d, %Y"),
                                                          g_date[6].strftime("%a - %b %d, %Y"))
              , justify=LEFT) \
            .grid(row=1, column=0, columnspan=10, sticky="w")
        Button(FF, text="DISCARD", width=10, command=lambda: indexer_default(name_sorter, i + 1, name_options, 1)) \
            .grid(row=2, column=3, sticky="w", columnspan=2)
        Label(FF, text="switch default to DISCARD").grid(row=2, column=1, sticky="w", columnspan=2)
        Button(FF, text="ADD NAME", width=10, command=lambda: indexer_default(name_sorter, i + 1, name_options, 0)) \
            .grid(row=3, column=3, sticky="w", columnspan=2)
        Label(FF, text="switch default to ADD NAME").grid(row=3, column=1, sticky="w", columnspan=2)
        Label(FF, text="").grid(row=4, column=0)
        Label(FF, text="Name", fg="grey").grid(row="5", column="1", sticky="w")
        Label(FF, text="Assignment", fg="grey").grid(row="5", column="2", sticky="w")
        Label(FF, text="Candidates", fg="grey").grid(row="5", column="3", sticky="w")
        c = 6  # item and grid row counter
        i = 0  # count iterations of the loop
        color = "blue"
        for t_name in tacs_list:
            possible_names = []
            Label(FF, text=str(i + 1), anchor="w").grid(row=c, column=0)
            Label(FF, text=t_name[1] + ", " + t_name[2], anchor="w", width=15, fg=color).grid(row=c, column=1)  # name
            Label(FF, text=t_name[3], anchor="w", width=10, fg=color).grid(row=c, column=2)  # assignment
            # build option menu for unmatched tacs names
            for c_name in c_list:
                if c_name[0] == t_name[1][0]:
                    possible_names.append(c_name)
            name_options = ["ADD NAME", "DISCARD"] + possible_names
            name_sorter.append(StringVar(FF))
            option_menu = OptionMenu(FF, name_sorter[i], *name_options)
            name_sorter[i].set(name_options[0])
            option_menu.config(width=15)
            option_menu.grid(row=c, column=3)  # possible matches
            if len(possible_names) == 1:  # display indicator for possible matches
                Label(FF, text=str(len(possible_names)) + " name").grid(row=c, column=4)
            if len(possible_names) > 1:
                Label(FF, text=str(len(possible_names)) + " names").grid(row=c, column=4)
            c += 1
            i += 1
        root.update()
        C.config(scrollregion=C.bbox("all"))
        mainloop()


def indexer_default(widget, count, options, choice):  # changes the default for the optionmenu widget
    for i in range(count - 1):
        widget[i].set(options[choice])


def apply_auto_indexer_3(self, buttons, file_path, tacs_list, name_sorter, new_carrier, c_list,
                         check_these):  # apply pairing screen 2
    # process incoming data
    i = 0  # count iterations of the loops.
    dis = 0  # count of discarded items
    add = 0  # count of added items
    pair = 0  # count of names paired to klusterbox names
    to_remove = []  # intialized array of names to be removed from tacs names
    to_addname = []  # initialize array of names to be added.
    to_nameindex = []  # initialize array of names to be be paired in name index
    sql = "SELECT tacs_name, kb_name, emp_id FROM name_index ORDER BY kb_name"
    results = inquire(sql)
    n_index = []  # create a list of klusterbox names
    id_index = []  # create a list of emp ids
    for line in results:  # loop to fill arrays
        n_index.append(line[1])
        id_index.append(line[2])
    for item in name_sorter:  # sort passed data from auto index 4
        if item.get() == "DISCARD":
            to_remove.append(tacs_list[i][0])
            dis += 1
        elif item.get() == "ADD NAME":
            to_addname.append(tacs_list[i])
            add += 1
        else:
            to_add = [tacs_list[i], item.get()]
            to_nameindex.append(to_add)
            to_remove.append(tacs_list[i][0])
            check_these.append(tacs_list[i])
            pair += 1
        i += 1

    pb_label = Label(buttons, text="Updating Changes: ")  # make label for progress bar
    pb_label.grid(row=0, column=2)
    pb = ttk.Progressbar(buttons, length=400, mode="determinate")  # create progress bar
    pb.grid(row=0, column=3)
    pb["maximum"] = len(to_nameindex)  # set length of progress bar
    pb.start()
    i = 0
    for item in to_nameindex:  # when a name from the optionmenu was selected
        pb["value"] = i  # increment progress bar
        tac_str = "{}, {}".format(item[0][1], item[0][2])  # tac str is last name and first initial from tacs report
        sql = "INSERT INTO name_index (tacs_name, kb_name, emp_id) VALUES('%s','%s','%s')" \
              % (tac_str, item[1], item[0][0])
        commit(sql)
        buttons.update()  # update the progress bar
        i += 1
    pb.stop()  # stop and destroy the progress bar
    pb_label.destroy()  # destroy the label for the progress bar
    pb.destroy()

    to_chg = []  # array of items from to_addname where the name needs to be modified with emp id
    new_name = []  # array of new names which have been modified with emp id
    for name in new_carrier: to_addname.append(name)  # add new carriers in list to be added to carrier table

    pb_label = Label(buttons, text="Updating Changes: ")  # make label for progress bar
    pb_label.grid(row=0, column=2)
    pb = ttk.Progressbar(buttons, length=400, mode="determinate")  # create progress bar
    pb.grid(row=0, column=3)
    pb["maximum"] = len(to_addname)  # set length of progress bar
    pb.start()
    i = 0
    for item in to_addname:  # when add name was selected from option menu
        pb["value"] = i  # increment progress bar
        tacs_str = "{}, {}".format(item[1], item[2])  # tacs str is last name and first initial from tacs report
        kb_str = "{}, {}".format(item[1], item[2])  # kb str is last name and first initial from tacs report
        if kb_str in n_index or kb_str in c_list:  # detect matches with name index
            sql = "SELECT emp_id, kb_name FROM name_index WHERE emp_id = '%s'" % item[0]
            result = inquire(sql)
            if not result:
                kb_str = "{} {}".format(kb_str, item[0])
                to_chg.append(item)
                mod_name = "{} {}".format(item[2], item[0])
                new_name.append(mod_name)
            if result:  # if the carrier is in the name index
                if result[0][1] != kb_str:  # if the kb name is not the same in the name index record - change name
                    to_chg.append(item)
                    mod_name = result[0][1].split(",")
                    mod_name = mod_name[1].strip()
                    new_name.append(mod_name)
        n_index.append(kb_str)  # add to n_index array so dups can be detected
        sql = "SELECT emp_id FROM name_index WHERE emp_id = '%s'" % item[0]
        result = inquire(sql)
        if not result:
            sql = "INSERT INTO name_index (tacs_name, kb_name, emp_id) VALUES('%s','%s','%s')" \
                  % (tacs_str, str(kb_str), item[0])
            commit(sql)
        buttons.update()  # update the progress bar
        i += 1
    pb.stop()  # stop and destroy the progress bar
    pb_label.destroy()  # destroy the label for the progress bar
    pb.destroy()
    """
    # message screens to summerize output
    messagebox.showinfo("Processing Carriers", "{} Carrier names were added to the database\n"
                                               "{} Carrier names were paired to names in klusterbox\n"
                                               "{} Carrier names were discarded.\n"
                                               .format(add, pair, dis))
    """
    count = 0  # swap out the names which have been modified in to_addname
    for item in to_chg:  # for each item to be swapped
        to_addname.remove(item)  # clear out the old one
        mod_str = [item[0], item[1], new_name[count], item[3]]  # create a modified array with modified name
        to_addname.append(mod_str)  # put in the new one
        count += 1

    if len(to_addname) > 0:
        auto_indexer_4(self, file_path, to_addname, check_these)
    elif len(check_these) > 0:
        auto_indexer_5(self, file_path, check_these)
    else:
        auto_indexer_6(self, file_path)


def auto_indexer_4(self, file_path, to_addname, check_these):  # add new carriers to carrier table / pairing screen #3
    days = ("Saturday", "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday")
    self.destroy()
    opt_nsday = []  # make an array of "day / color" options for option menu
    full_ns_dict = {}
    # get ns structure preference from database
    sql = "SELECT tolerance FROM tolerances WHERE category='%s'" % "ns_auto_pref"
    result = inquire(sql)
    ns_toggle = result[0][0] # modify available ns days per ns_toggle
    if ns_toggle == "rotation":
        remove_array = ("sat", "mon", "tue", "wed", "thu", "fri")
    else:
        remove_array = ("green", "brown", "red", "black", "yellow", "blue")
    ns_code_mod = dict() # copy the ns_code dict to ns_code_mod using dict()
    for key in ns_code:
        ns_code_mod[key]=ns_code[key]
    for key in remove_array:
        if key in ns_code_mod:
            del ns_code_mod[key]  # modify available ns days per ns_toggle

    for each in ns_code_mod:  #
        ns_option = ns_code_mod[each] + " - " + each  # make a string for each day/color
        if each == "none": ns_option = "       " + " - " + each  # if the ns day is "none" - make a special string
        opt_nsday.append(ns_option)
    for each in opt_nsday:  # Make a dictionary to match full days and option menu options
        for day in days:
            if day[:3] == each[:3]:
                full_ns_dict[day] = each  # creates full_ns_dict
        if each[-4:] == "none":
            ns_option = "       " + " - " + "none"  # if the ns day is "none" - make a special string
            full_ns_dict["None"] = ns_option  # creates full_ns_dict None option
    results = gen_ns_dict(file_path, to_addname)  # returns id and name
    ns_dict = {}  # create dictionary for ns day data
    for id in results:  # loop to fill dictionary with ns day info
        ns_dict[id[0]] = id[1]
    F = Frame(root)
    F.pack(fill=BOTH, side=LEFT)
    C1 = Canvas(F)
    C1.pack(fill=BOTH, side=BOTTOM)
    Button(C1, text="Continue", width=8, command=lambda: apply_auto_indexer_4
    (F, C1, file_path, carrier_name, l_s, l_ns, route, check_these)).pack(side=LEFT)
    Button(C1, text="Cancel", width=8, command=lambda: (F.destroy(), main_frame())).pack(side=LEFT)
    S = Scrollbar(F)  # link up the canvas and scrollbar
    C = Canvas(F, width=1600)
    S.pack(side=RIGHT, fill=BOTH)
    C.pack(side=LEFT, fill=BOTH, pady=10, padx=20)
    S.configure(command=C.yview, orient="vertical")
    C.configure(yscrollcommand=S.set)
    if sys.platform == "win32":
        C.bind_all('<MouseWheel>', lambda event: C.yview_scroll(int(-1 * (event.delta / 120)), "units"))
    elif sys.platform == "linux":
        C.bind_all('<Button-4>', lambda event: C.yview('scroll', -1, 'units'))
        C.bind_all('<Button-5>', lambda event: C.yview('scroll', 1, 'units'))
    FF = Frame(C)  # create the frame inside the canvas
    C.create_window((0, 0), window=FF, anchor=NW)
    Label(FF, text="Input New Carriers", font="bold", pady=10).grid(row=0, column=0, sticky="w",
                                                                    columnspan=6)  # Pairing Screen #3
    Label(FF, text=
    "Enter in information for carriers not already recorded in the Klusterbox database. You can use the TACS \n"
    "information (shown in blue),as a guide if it is accurate. As OTDL/WAL information is not in TACS, it is \n"
    "not shown and this information will have to requested from management. Routes must be only 4 digits \n"
    "long. In cases were there are multiple routes, the routes must be separated by a \"/\" backslash.\n\n"
    "Investigation Range: {0} through {1}\n\n".format(g_date[0].strftime("%a - %b %d, %Y"),
                                                      g_date[6].strftime("%a - %b %d, %Y"))
          , justify=LEFT).grid(row=1, column=0, sticky="w", columnspan=6)
    y = 2  # count for the row
    Label(FF, text="Name", fg="Grey").grid(row=y, column=0, sticky="w")
    Label(FF, text="List Status", fg="Grey").grid(row=y, column=1, sticky="w")
    Label(FF, text="NS Day", fg="Grey").grid(row=y, column=2, sticky="w")
    Label(FF, text="Route_s", fg="Grey").grid(row=y, column=3, sticky="w")
    Label(FF, text="Station", fg="Grey").grid(row=y, column=4, sticky="w")
    Label(FF, text="              ", fg="Grey").grid(row=y, column=5, sticky="w")
    y += 1
    i = 0  # count the instances of the array
    carrier_name = []  # create array for carrier names
    l_s = []  # create array for list status
    l_ns = []  # create array for ns days
    route = []  # create array for routes
    color = "blue"
    for name in to_addname:
        Label(FF, text=name[1] + ", " + name[2], fg=color).grid(row=y, column=0, sticky="w")
        carrier_name.append(str(name[1] + ", " + name[2]))
        Label(FF, text="not in record", fg=color).grid(row=y, column=1, sticky="w")
        Label(FF, text=str(ns_dict[name[0]]), fg=color).grid(row=y, column=2, sticky="w")
        Label(FF, text=name[3], fg=color).grid(row=y, column=3, sticky="w")
        Label(FF, text=g_station, fg=color).grid(row=y, column=4, sticky="w")
        y += 1
        # Label(FF, text="               ").grid(row=y, column=0) # 15 spaces / force 15 space width for column 0
        list_options = ("otdl", "wal", "nl", "aux")  # create optionmenu for list status
        if name[3] == "auxiliary":
            lx = 3  # configure defaults for list status
        else:
            lx = 2  # set as 'nl' if not 'aux'
        l_s.append(StringVar(FF))
        l_s[i].set(list_options[lx])  # set the list status
        list_status = OptionMenu(FF, l_s[i], *list_options)
        list_status.config(width=5)
        list_status.grid(row=y, column=1, sticky="w")
        l_ns.append(StringVar(FF))  # create optionmenu for ns days
        l_ns[i].set(full_ns_dict[str(ns_dict[name[0]])])  # set ns day default
        ns_day = OptionMenu(FF, l_ns[i], *opt_nsday)
        ns_day.config(width=12)
        ns_day.grid(row=y, column=2, sticky="w")
        route.append(StringVar(FF))  # create entry field for route
        Entry(FF, width=24, textvariable=route[i]).grid(row=y, column=3, sticky="w")  # create entry for routes
        if name[3][-4:].isnumeric():
            rte = name[3][-4:]
        else:
            rte = ""
        route[i].set(rte)
        y += 1
        i += 1
        Label(FF, text="").grid(row=y, column=0, sticky="w")
        y += 1
    root.update()
    C.config(scrollregion=C.bbox("all"))
    mainloop()


def apply_auto_indexer_4(self, buttons, file_path, carrier_name, l_s, l_ns, route,
                         check_these):  # adds new carriers to the carriers table
    if g_range == "week": eff_date = g_date[0]
    if g_range == "day": eff_date = d_date
    station = StringVar(self)  # put station var in a StringVar object
    station.set(g_station)
    pb_label = Label(buttons, text="Updating Changes: ")  # make label for progress bar
    pb_label.pack(side=LEFT)
    pb = ttk.Progressbar(buttons, length=400, mode="determinate")  # create progress bar
    pb.pack(side=LEFT)
    pb["maximum"] = len(carrier_name)  # set length of progress bar
    pb.start()
    for i in range(len(carrier_name)):
        pb["value"] = i  # increment progress bar
        passed_ns = l_ns[i].get().split(" - ")  # clean the passed ns day data
        clean_ns = StringVar(self)  # put ns day var in StringVar object
        clean_ns.set(passed_ns[1])
        apply_2(eff_date, carrier_name[i], l_s[i], clean_ns, route[i], station, self)
        buttons.update()
    pb.stop()  # stop and destroy the progress bar
    pb_label.destroy()  # destroy the label for the progress bar
    pb.destroy()
    results = []
    for name in carrier_name:
        sql = "SELECT * FROM carriers WHERE carrier_name == '%s' and effective_date == '%s'" % (name, eff_date)
        result = inquire(sql)
        if result: results.append(result)
    if len(results) >= len(carrier_name) and len(check_these) > 0:
        auto_indexer_5(self, file_path, check_these)
    elif len(results) >= len(carrier_name):
        auto_indexer_6(self, file_path)
    else:
        return


def gen_rev_ns_dict():  # creates full day/color ns day dictionary
    days = ("Saturday", "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday")
    colors = ("blue", "green", "brown", "red", "black", "yellow")
    code_ns = {}
    for d in days:
        for c in colors:
            if d[:3] == ns_code[c]:
                code_ns[d] = c
    code_ns["None"] = "none"
    return code_ns


def auto_indexer_5(self, file_path, check_these):  # correct discrepancies
    if len(check_these) == 0: auto_indexer_6(self, file_path)
    check_these.sort(key=itemgetter(1))  # sort the incoming tacs information
    self.destroy()
    opt_nsday = []  # make an array of "day / color" options for option menu
    ns_opt_dict = {}  # creates a dictionary of ns colors/ options for menu
    for each in ns_code:  # creates the option menu options for ns day menu
        ns_option = ns_code[each] + " - " + each  # make a string for each day/color
        ns_opt_dict[each] = ns_option
        if each == "none":
            ns_option = "       " + " - " + each  # if the ns day is "none" - make a special string
            ns_opt_dict[each] = ns_option
        opt_nsday.append(ns_option)
    results = gen_ns_dict(file_path, check_these)  # returns id and name
    ns_dict = {}  # create dictionary for ns day data
    for id in results:  # loop to fill dictionary with ns day info
        ns_dict[id[0]] = id[1]
    carrier_list = gen_carrier_list()  # generate an in range carrier list
    carriers_names_list = []  # generate list of only names from 'in range carrier list'
    for name in carrier_list:
        carriers_names_list.append(name[1])
    name_dict = gen_nameindex_dict()  # generate dictionary for emp id to kb_name
    remainders = []  # find carriers in 'check these' but not in 'in range carrier list' aka 'remainders'
    for name in check_these:
        if name_dict[name[0]] not in carriers_names_list:
            remainders.append(name)
    for name in remainders:  # get carriers data from carriers for remainders
        sql = "SELECT * FROM carriers WHERE carrier_name = '%s' and effective_date <= '%s'" \
              "ORDER BY effective_date desc" % (name_dict[name[0]], g_date[0])
        result = inquire(sql)
        carrier_list.append(list(result[0]))
    carrier_list.sort(key=itemgetter(1))  # resort carrier list after additions
    code_ns = gen_rev_ns_dict()  # generate reverse ns code dictionary
    wd = front_window("none")  # get window objects 0=F,1=S,2=C,3=FF,4=buttons
    header = Frame(wd[3])
    header.grid(row=0, columnspan=6, sticky="w")
    Label(header, text="Discrepancy Resolution Screen", font="bold,", pady=10) \
        .grid(row=0, sticky="w")
    Label(header, text=
    "Correct any discrepancies and inconsistancies that exist between the incoming TACS data (in blue) and the \n"
    "information currently recorded in the Klusterbox database (below in the entry fields and option menus)to reflect \n"
    "the carrier's status acurately. This will update the Klusterbox database. Routes must 4 digits long. In cases \n"
    "were there multiple routes, the routes must be separated by a \"/\" backslash.\n\n"
    "Investigation Range: {0} through {1}\n\n".format(g_date[0].strftime("%a - %b %d, %Y"),
                                                      g_date[6].strftime("%a - %b %d, %Y")),
          justify=LEFT).grid(row=1, sticky="w")
    y = 1  # count for the row
    Label(wd[3], text="    ", fg="Grey").grid(row=y, column=0, sticky="w")
    Label(wd[3], text="List Status", fg="Grey").grid(row=y, column=1, sticky="w")
    Label(wd[3], text="NS Day", fg="Grey").grid(row=y, column=2, sticky="w")
    Label(wd[3], text="Route_s", fg="Grey").grid(row=y, column=3, sticky="w")
    Label(wd[3], text="Station", fg="Grey").grid(row=y, column=4, sticky="w")
    Label(wd[3], text="             ", fg="Grey").grid(row=y, column=5, sticky="w")
    y += 1
    i = 0  # count the instances of the array
    carrier_name = []  # create array for carrier names
    l_s = []  # create array for list status
    l_ns = []  # create array for ns days
    e_route = []  # create array for routes
    l_station = []
    aux_list_tuple = ("aux")
    reg_list_tuple = ("nl", "wal", "otdl")
    skip_this_screen = "yes"
    for name in check_these:
        for k_name in carrier_list:
            if name_dict[name[0]] == k_name[1]:
                if name[3] == "auxiliary":  # parse assignments from tacs list
                    tlist = aux_list_tuple
                    tnsday = "none"
                    troute = ""
                if name[3][-4:].isnumeric() == True:
                    tlist = reg_list_tuple
                    tnsday = code_ns[str(ns_dict[name[0]])]
                    troute = name[3][-4:]
                if name[3][-7:] == "floater":
                    tlist = reg_list_tuple
                    tnsday = code_ns[str(ns_dict[name[0]])]
                    troute = "floater"
                if name[3] == "undetected":
                    tlist = "undetected"
                    tnsday = code_ns[str(ns_dict[name[0]])]
                    troute = "undetected"
                tstation = g_station
                trip_wire = "set"
                # check tacs data against data in carriers table/ klusterbox
                if k_name[2] not in tlist:
                    trip_wire = "sprung"  # check list status
                if k_name[3] != tnsday:
                    trip_wire = "sprung"  # check nsday
                k_rte_len = len(k_name[4].split('/'))  # check route
                if k_rte_len == 0:  # check if route is aux
                    if troute != "":
                        trip_wire = "sprung"
                if k_rte_len == 1:  # check if route is regular
                    if troute != k_name[4]:
                        trip_wire = "sprung"
                if k_rte_len == 5:  # check if route is floater
                    if troute != "floater":
                        trip_wire = "sprung"
                if tstation != k_name[5]:  # check if station is correct
                    trip_wire = "sprung"
                if trip_wire == "sprung":
                    skip_this_screen = "no"  # if there are no discrepancies, then skip the screen
                    # create the page content
                    color = "blue"
                    name_F = Frame(wd[3])  # create separate frame for names
                    name_F.grid(row=y, columnspan=6, sticky="w")
                    Label(name_F, text="Name: ", fg="Grey").grid(row=0, column=0, sticky="w")
                    Label(name_F, text=name[1] + ", " + name[2], fg=color).grid(row=0, column=1, sticky="w")
                    Label(name_F, text=" / " + k_name[1]).grid(row=0, column=2, sticky="w")
                    y += 1
                    Label(wd[3], text="    ", fg=color).grid(row=y, column=0, sticky="w")
                    Label(wd[3], text="not in record", fg=color).grid(row=y, column=1, sticky="w")
                    Label(wd[3], text=str(ns_dict[name[0]]), fg=color).grid(row=y, column=2, sticky="w")
                    Label(wd[3], text=name[3], fg=color).grid(row=y, column=3, sticky="w")
                    Label(wd[3], text=g_station, fg=color).grid(row=y, column=4, sticky="w")
                    y += 1
                    carrier_name.append(k_name[1])  # add kb name to the array
                    list_options = ("otdl", "wal", "nl", "aux")  # create optionmenu for list status
                    l_s.append(StringVar(wd[3]))
                    l_s[i].set(k_name[2])  # set the list status
                    list_status = OptionMenu(wd[3], l_s[i], *list_options)
                    list_status.config(width=6)
                    list_status.grid(row=y, column=1, sticky="w")
                    l_ns.append(StringVar(wd[3]))  # create optionmenu for ns days
                    l_ns[i].set(ns_opt_dict[k_name[3]])  # set ns day default
                    ns_day = OptionMenu(wd[3], l_ns[i], *opt_nsday)
                    ns_day.config(width=12)
                    ns_day.grid(row=y, column=2, sticky="w")
                    e_route.append(StringVar(wd[3]))  # create entry field for route
                    Entry(wd[3], width=25, textvariable=e_route[i]).grid(row=y, column=3,
                                                                         sticky="w")  # create entry for routes
                    e_route[i].set(k_name[4])
                    l_station.append(StringVar(wd[3]))
                    l_station[i].set(k_name[5])
                    list_station = OptionMenu(wd[3], l_station[i], *list_of_stations)
                    list_station.config(width=25)
                    list_station.grid(row=y, column=4, sticky="w")
                    y += 1
                    Label(wd[3], text="").grid(row=y, column=0)
                    y += 1
                    i += 1
    Button(wd[4], text="Continue", width=8,
           command=lambda: apply_auto_indexer_5(wd[0], wd[4], file_path, carrier_name, l_s, l_ns, e_route,
                                                l_station, check_these)).pack(side=LEFT)
    Button(wd[4], text="Cancel", width=8, command=lambda: (wd[0].destroy(), main_frame())).pack(side=LEFT)
    if skip_this_screen == "yes":
        auto_indexer_6(wd[0], file_path)
    else:
        rear_window(wd)  # get rear window objects


def apply_auto_indexer_5(self, buttons, file_path, carrier_name, l_s, l_ns, e_route, l_station, check_these):
    # adds new carriers to the carriers table
    if g_range == "week": eff_date = g_date[0]
    if g_range == "day": eff_date = d_date
    pb_label = Label(buttons, text="Updating Changes: ")  # make label for progress bar
    pb_label.pack(side=LEFT)
    pb = ttk.Progressbar(buttons, length=400, mode="determinate")  # create progress bar
    pb.pack(side=LEFT)
    pb["maximum"] = len(carrier_name)  # set length of progress bar
    pb.start()
    for i in range(len(carrier_name)):
        pb["value"] = i  # increment progress bar
        passed_ns = l_ns[i].get().split(" - ")  # clean the passed ns day data
        clean_ns = StringVar(self)  # put ns day var in StringVar object
        clean_ns.set(passed_ns[1])
        check = apply_2_auto_indexer_5(eff_date, carrier_name[i], l_s[i], clean_ns, e_route[i], l_station[i], self)
        if check == "error":
            self.destroy()
            auto_indexer_5(self, file_path, check_these)
            break
        buttons.update()
    pb.stop()  # stop and destroy the progress bar
    pb_label.destroy()  # destroy the label for the progress bar
    pb.destroy()
    auto_indexer_6(self, file_path)


def apply_2_auto_indexer_5(date, carrier, ls, ns, route, station, self):
    route_list = route.get().split("/")
    if len(route.get()) > 24:
        messagebox.showerror("Route number input error", "There can be no more than five routes per carrier "
                                                         "(for T6 carriers).\n Routes numbers can be no more than four digits long.\n"
                                                         "If there are multiple routes, route numbers must be separated by "
                                                         "the \'/\' character. For example: 1001/1015/1024/1036/1072. Do not use "
                                                         "commas or empty spaces", parent=self)
        return "error"
    for item in route_list:
        item = item.strip()
        if item != "":
            if len(item) != 4:
                messagebox.showerror("Route number input error", 'Routes numbers must be four digits long.\n'
                                                                 'If there are multiple routes, route numbers must be separated by '
                                                                 'the \'/\' character. For example: 1001/1015/1024/1036/1072. Do not use '
                                                                 'commas or empty spaces', parent=self)
                return "error"
        if item.isdigit() == FALSE and item != "":
            messagebox.showerror("Route number input error", "Route numbers must be numbers and can not contain "
                                                             "letters", parent=self)
            return "error"
    route_input = route.get()
    if route_input == "0000":
        route_input = ""
    sql = "SELECT effective_date, carrier_name,list_status, ns_day,route_s, station, rowid FROM carriers " \
          "WHERE carrier_name = '%s' and effective_date = '%s' ORDER BY effective_date" % (carrier, date)
    results = inquire(sql)
    if len(results) == 0:
        sql = "INSERT INTO carriers (effective_date,carrier_name,list_status,ns_day,route_s,station)" \
              " VALUES('%s','%s','%s','%s','%s','%s')" \
              % (date, carrier, ls.get(), ns.get(), route_input, station.get())
        commit(sql)
    elif len(results) == 1:
        sql = "UPDATE carriers SET list_status='%s',ns_day='%s',route_s='%s',station='%s' " \
              "WHERE effective_date = '%s' and carrier_name = '%s'" % \
              (ls.get(), ns.get(), route_input, station.get(), date, carrier)
        commit(sql)
    elif len(results) > 1:
        sql = "DELETE FROM carriers WHERE effective_date ='%s' and carrier_name = '%s'" % (date, carrier)
        commit(sql)
        sql = "INSERT INTO carriers (effective_date,carrier_name,list_status,ns_day,route_s,station)" \
              " VALUES('%s','%s','%s','%s','%s','%s')" \
              % (date, carrier, ls.get(), ns.get(), route_input, station.get())
        commit(sql)


def auto_indexer_6(self,file_path):  # identify and remove any carriers in the carrier
                                    # list who are not in the TACS list
    carrier_list = gen_carrier_list()  # create names_list array
    names_list = []
    for name in carrier_list:
        if name[1] not in names_list:
            names_list.append(name[1])
    tacs_ids = []  # generate tacs list
    good_jobs = ("134", "844")
    with open(file_path, newline="") as file:
        a_file = csv.reader(file)
        to_add = ("x", "x")  # create placeholder for
        for line in a_file:
            if len(line) > 19:  # if there are enough items in the line
                if line[18] == "Temp":
                    to_add = (line[4].zfill(8), line[19])
                elif line[19] != "Temp" or line[19] != "Base":
                    if to_add != ("x", "x"):  # if not placeholder
                        tacs_ids.append(to_add)  # add tacs data to the array
                        to_add = ("x", "x")  # reset placeholder
                if line[18] == "Base":
                    to_add = (line[4].zfill(8), line[19])
    filtered_ids = []  # filter the tacs ids to only good jobs
    for item in tacs_ids:
        if item[1] in good_jobs:
            filtered_ids.append(item)
    del tacs_ids
    t_names = []  # matches emp id to the kb name
    for name in filtered_ids:  #
        sql = "SELECT kb_name FROM name_index WHERE emp_id = '%s'" % (name[0])
        result = inquire(sql)  # check dbase for a match
        if result:  # if there is a match in the dbase, then add data to array
            t_names.append(result[0][0])
    ex_carrier = []  # carriers in carrier list but not tacs data
    for name in names_list:  # for each name in carrier list
        if name not in t_names:  # if they are not also in the tacs data
            ex_carrier.append(name)  # then add them to the array
    wd = front_window(self)  # get window objects 0=F,1=S,2=C,3=FF,4=buttons
    header = Frame(wd[3])
    header.grid(row=0, columnspan=5, sticky="w")
    Label(header, text="Carriers No Longer At Station", font="bold,", pady=10) \
        .grid(row=0, sticky="w")
    Label(header, text=
    "Klusterbox has detected that the following carriers may no longer be at the station. If they are no longer at the\n"
    "station, then please use the option menu below to move them to the correct station (if listed). If the correct \n"
    "is not listed or the carrier is no longer working for the post office, then select \"out of station\".\n\n"
    "Investigation Range: {0} through {1}\n\n".format(g_date[0].strftime("%a - %b %d, %Y"),
                                                      g_date[6].strftime("%a - %b %d, %Y")),
          justify=LEFT).grid(row=1, sticky="w")
    y = 1  # count for the row
    Label(wd[3], text="Name", fg="Grey").grid(row=y, column=0, sticky="w")
    Label(wd[3], text="List Status", fg="Grey").grid(row=y, column=1, sticky="w")
    Label(wd[3], text="Route_s", fg="Grey").grid(row=y, column=2, sticky="w")
    Label(wd[3], text="Station", fg="Grey").grid(row=y, column=3, sticky="w")
    Label(wd[3], text="             ", fg="Grey").grid(row=y, column=4, sticky="w")
    y += 1
    carrier_name = []
    list_status = []
    ns_day = []
    route = []
    station = []
    new_station = []
    c = 0
    for name in ex_carrier:
        sql = "SELECT * FROM carriers WHERE carrier_name = '%s' and effective_date <= '%s' ORDER BY effective_date DESC" \
              % (name, g_date[0])
        result = inquire(sql)
        carrier_name.append(StringVar(wd[3]))  # store name
        carrier_name[c].set(result[0][1])
        Button(wd[3], text=result[0][1], relief=RIDGE, width=25, anchor="w").grid(row=y, column=0, sticky="w")  # name
        list_status.append(StringVar(wd[3]))  # store list status
        list_status[c].set(result[0][2])
        Button(wd[3], text=result[0][2], relief=RIDGE, width=7, anchor="w").grid(row=y, column=1, sticky="w")  # list
        ns_day.append(StringVar(wd[3]))  # store ns day
        ns_day[c].set(result[0][3])
        route.append(StringVar(wd[3]))  # store route
        route[c].set(result[0][4])
        Button(wd[3], text=result[0][4], relief=RIDGE, width=20, anchor="w").grid(row=y, column=2, sticky="w")  # route
        station.append(StringVar(wd[3]))  # store station
        station[c].set(result[0][5])
        new_station.append(StringVar(wd[3]))
        new_station[c].set(result[0][5])
        stat_om = OptionMenu(wd[3], new_station[c], *list_of_stations)  # station
        stat_om.config(width=25, anchor="w")
        stat_om.grid(row=y, column=3, sticky="w")
        Label(wd[3], text="                     ").grid(row=y, column=4)
        c += 1
        y += 1
    if len(carrier_name) == 0:
        auto_skimmer(wd[0], file_path)
    else:
        Button(wd[4], text="Continue", width=8,
               command=lambda: apply_auto_indexer_6(wd[0], wd[4], file_path, carrier_name,
                                                    list_status, ns_day, route, station, new_station)).pack(side=LEFT)
        Button(wd[4], text="Cancel", width=8, command=lambda: (wd[0].destroy(), main_frame())).pack(side=LEFT)
        rear_window(wd)


def apply_auto_indexer_6(self, buttons, file_path, carrier_name, list_status, ns_day, route, station, new_station):
    date = g_date[0]
    pb_label = Label(buttons, text="Updating Changes: ")  # make label for progress bar
    pb_label.pack(side=LEFT)
    pb = ttk.Progressbar(buttons, length=400, mode="determinate")  # create progress bar
    pb.pack(side=LEFT)
    pb["maximum"] = len(carrier_name)  # set length of progress bar
    pb.start()
    for i in range(len(carrier_name)):
        pb["value"] = i  # increment progress bar
        if station[i].get() != new_station[i].get():
            apply_2_auto_indexer_5(date, carrier_name[i].get(), list_status[i], ns_day[i], route[i], new_station[i],
                                   self)
        buttons.update()
    pb.stop()  # stop and destroy the progress bar
    pb_label.destroy()  # destroy the label for the progress bar
    pb.destroy()
    auto_skimmer(self, file_path)


def auto_skimmer(self, file_path):
    self.destroy()
    global allow_zero_top
    global allow_zero_bottom
    global skippers
    sql = "SELECT tolerance FROM tolerances WHERE category='%s'" % "allow_zero_top"
    result = inquire(sql)
    allow_zero_top = result[0][0]
    sql = "SELECT tolerance FROM tolerances WHERE category='%s'" % "allow_zero_bottom"
    result = inquire(sql)
    allow_zero_bottom = result[0][0]
    sql = "SELECT code FROM skippers"  # get skippers data from dbase
    results = inquire(sql)
    skippers = []  # fill the array for skippers
    for item in results:
        skippers.append(item[0])
    carrier_list_cleaning_for_auto_skimmer()
    ok = messagebox.askokcancel("Auto Rings", "Do you want to automatically enter the rings?")
    if ok == False:
        main_frame()
    else:
        days = ("Saturday", "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday")
        mv_codes = ("BT", "MV", "ET")
        carrier = []
        proto_array = []
        pb_root = Tk()  # create a window for the progress bar
        pb_root.geometry("%dx%d+%d+%d" % (500, 50, 200, 300))
        pb_root.title("Entering Carrier Rings")
        pb_label = Label(pb_root, text="Updating Rings: ")  # make label for progress bar
        pb_label.pack(side=LEFT)
        pb = ttk.Progressbar(pb_root, length=400, mode="determinate")  # create progress bar
        pb.pack(side=LEFT)
        with open(file_path, newline="") as file:
            a_file = csv.reader(file)
            row_count = sum(1 for row in a_file)  # get number of rows in csv file
        with open(file_path, newline="") as file:
            a_file = csv.reader(file)
            pb["maximum"] = int(row_count)  # set length of progress bar
            pb.start()
            i = 0
            c = 0
            good_id = "no"
            for line in a_file:
                pb["value"] = i  # increment progress bar
                if c == 0:
                    if line[0][:8] != "TAC500R3":
                        messagebox.showwarning("File Selection Error", "The selected file does not appear to be an "
                                                                       "Employee Everything report.")
                        return
                if c != 0:
                    if good_id != line[4] and good_id != "no":  # if new carrier or employee
                        proto_rings = auto_weekly_analysis(carrier)  # trigger analysis
                        proto_array.append(proto_rings)
                        del carrier[:]  # empty array
                        good_id = "no"  # reset trigger
                    if line[18] == "Base" and line[19] == "844" or line[
                        19] == "134":  # find first line of specific carrier
                        good_id = line[4]  # set trigger to id of carriers who are FT or aux carriers
                        carrier.append(line)  # gather times and moves for anaylsis
                    if good_id == line[4] and line[18] != "Base":
                        if line[18] in days:  # get the hours for each day
                            carrier.append(line)  # gather times and moves for anaylsis
                        if line[19] in mv_codes and line[32] != "(W)Ring Deleted From PC":
                            carrier.append(line)  # gather times and moves for anaylsis
                c += 1
                pb_root.update()
                i += 1
            auto_weekly_analysis(carrier)  # when loop ends, run final analysis
            del carrier[:]  # empty array
            pb.stop()  # stop and destroy the progress bar
            pb_label.destroy()  # destroy the label for the progress bar
            pb.destroy()
            pb_root.destroy()
        messagebox.showinfo("Auto Rings",
                            "The Employee Everything Report has been sucessfully inputed into the database")
        main_frame()


def auto_weekly_analysis(array):
    days = ("Saturday", "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday")
    day_dict = {}
    x = 0
    for item in days:  # make a dictionary for each day in the week
        day_dict[item] = g_date[x]
        x += 1
    rings = []
    input_rings = []
    good_day = "no"
    for line in array:
        if line[18] in days and line[18] != good_day and good_day != "no":
            to_input = auto_daily_analysis(rings)
            input_rings.append(to_input)
            del rings[:]
            good_day = line[18]
        if line[18] == "Base" and line[19] == "844" or line[19] == "134":  # find first line of specific carrier
            continue  # gather base line data
        elif line[18] == "Temp" and line[19] == "844" or line[19] == "134":  # find first line of specific carrier
            continue  # gather base line data
        else:
            if line[18] in days and line[18] == good_day:
                rings.append(line)
            if line[18] in days and good_day == "no":  # day change triggers
                good_day = line[18]
                rings.append(line)
            if line[18] not in days:
                rings.append(line)
    to_input = auto_daily_analysis(rings)  # call function for last line
    input_rings.append(to_input)  # add the proto array for an array
    # return input_rings # send it back to auto skimmer()
    if input_rings[0] != None:
        sql = "SELECT kb_name FROM name_index WHERE emp_id = '%s'" % input_rings[0][1]
        result = inquire(sql)  # check to verify that they are in the name index
        if result:  # if there is a match in the name index, then continue
            kb_name = result[0][0]  # get the kb name which correlates to the emp id
            for line in input_rings:
                sql = "SELECT effective_date, carrier_name, list_status, ns_day, route_s FROM" \
                      " carriers WHERE carrier_name = '%s' and effective_date <= '%s' " \
                      "ORDER BY effective_date DESC" % (kb_name, day_dict[line[0]])
                result = inquire(sql)
                for array in result:  # find the most recent carrier record
                    eff_date = datetime.strptime(array[0], '%Y-%m-%d %H:%M:%S')
                    if eff_date <= day_dict[line[0]]:
                        newest_carrier = array
                        break  # stop. we only need the most recent record
                if not result:
                    return
                # find the code, if any
                if newest_carrier[2] == "nl" or newest_carrier[2] == "wal":
                    if day_dict[line[0]].strftime("%a") == ns_code[newest_carrier[3]] and float(line[2]) > 0:
                        c_code = "ns day"
                    else:
                        c_code = "none"
                elif newest_carrier[2] == "otdl" or newest_carrier[2] == "aux":
                    if line[4] == "":
                        c_code = "none"  # line[4] is the code from proto-array
                    else:
                        c_code = line[4]  # can be sick or annual
                else:
                    c_code = "none"
                routes = []  # create an array for routes
                if newest_carrier[4] != "": routes = newest_carrier[4].split("/")
                # find the moves if any
                mv_triad = []  # triad is route#, start time off route, end time off route
                mv_str = ""
                route_holder = ""
                if len(routes) > 0:  # if the route is in kb
                    pair = "closed"  # trigger opens when a move set needs to be closed
                    for m in line[5]:  # loop through all the rings
                        if m[3] not in routes and pair == "closed":
                            if m[3] == "0000" and m[2] in skippers:  # sometimes off route is not off route
                                continue
                            else:
                                route_holder = m[3]  # hold route to put at end of triad
                                mv_triad.append(m[1])  # add start time to second place of triad
                                pair = "open"
                        if m[3] in routes and pair == "open":
                            mv_triad.append(m[1])  # add end time to third place of triad
                            mv_triad.append(route_holder)
                            pair = "closed"
                    if pair == "open":  # if open at end, then close it with the last ring
                        mv_triad.append(line[5][len(line[5]) - 1][1])
                        mv_triad.append(route_holder)
                if allow_zero_bottom == False:
                    if len(mv_triad) > 0:  # find and remove duplicate ET rings at end
                        if mv_triad[int(len(mv_triad) - 3)] == mv_triad[
                            int(len(mv_triad) - 2)]:  # if the last 2 are the same
                            mv_triad.pop()  # pop out the last triad
                            mv_triad.pop()
                            mv_triad.pop()
                if allow_zero_top == False:
                    if len(mv_triad) > 0:  # find and remove rings in the front
                        if mv_triad[0] == mv_triad[1]:
                            mv_triad.pop(0)  # pop out the triad
                            mv_triad.pop(0)
                            mv_triad.pop(0)
                mv_str = ','.join(mv_triad)  # format array as string to fit in dbase
                # if hours worked > 0 or there is a code or a leave type
                if float(line[2]) > 0 or c_code != "none" or line[6]!="":
                    if float(line[2]) == 0:
                        hr_52 = ""  # don't put zeros in 5200 for rings record
                    else:
                        hr_52 = float(line[2])  # if it is greater than zero, put it in as a float
                    lv_time = float(line[7]) # convert the leave time to a float var
                    current_array = [str(day_dict[line[0]]), kb_name, hr_52, line[3], c_code, mv_str, line[6], lv_time]
                    # check rings table to see if record already exist.
                    sql = "SELECT * FROM rings3 WHERE carrier_name = '%s' and rings_date = '%s'" % (
                    kb_name, day_dict[line[0]])
                    result = inquire(sql)
                    if len(result) == 0:
                        sql = "INSERT INTO rings3 (rings_date, carrier_name, total, rs, code, moves,leave_type,leave_time) " \
                              "VALUES('%s','%s','%s','%s','%s','%s','%s','%s')" % \
                              (current_array[0], current_array[1], current_array[2], current_array[3], current_array[4],
                               current_array[5],current_array[6],current_array[7])
                        commit(sql)
                    else:
                        sql = "UPDATE rings3 SET total='%s',rs='%s' ,code='%s',moves='%s'," \
                              "leave_type ='%s',leave_time = '%s'" \
                              "WHERE rings_date = '%s' and carrier_name = '%s'" \
                              % (
                              current_array[2], current_array[3], current_array[4], current_array[5],
                              current_array[6],current_array[7],
                              current_array[0],current_array[1])
                        commit(sql)


def auto_daily_analysis(rings):
    days = ("Saturday", "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday")
    mv_codes = ("BT", "MV", "ET")
    hr_52 = 0.0 # work hours
    hr_55 = 0.0 # annual leave
    hr_56 = 0.0 # sick leave
    hr_58 = 0.0 # holiday leave
    hr_62 = 0.0 # guaranteed time
    hr_86 = 0.0 # other paid leave
    rs = 0
    code = ""
    moves = []
    leave_type = []
    leave_time = []
    final_leave_type = ""
    final_leave_time = 0.0
    if len(rings) > 0:
        name = rings[0][4].zfill(8)  # Get NAME
        for line in rings:
            if line[18] in days:  # get 5200 or non 5200 times for TOTAL, code, leave_type and leave_time
                dayofweek = line[18]
                spt_20 = line[20].split(':')  # split to get code and hours
                # get second and third digits of the of the split line 20 or spt_20
                spt_20_mod = "".join([spt_20[0][1], spt_20[0][2]])
                if spt_20_mod == "52":
                    hr_52 = spt_20[1]  # get the total hours worked
                if spt_20_mod == "55":
                    hr_55 = spt_20[1]  # get the annual leave hours
                if spt_20_mod == "56":
                    hr_56 = spt_20[1]  # get the sick leave hours
                if spt_20_mod == "58":
                    hr_58 = spt_20[1]  # get the holiday leave hours
                if spt_20_mod == "62":
                    hr_62 = spt_20[1]  # get the guaranteed time hours
                if spt_20_mod == "86":
                    hr_86 = spt_20[1]  # get other leave hours

                # calculate the leave type and time:
                if float(hr_55) > 0 or float(hr_56) > 0 or float(hr_58) > 0 or float(hr_62) > 0 or float(hr_86) > 0:
                    if float(hr_55) > 0:
                        leave_type.append("annual")
                        leave_time.append(hr_55)
                    if float(hr_56) > 0:
                        leave_type.append("sick")
                        leave_time.append(hr_56)
                    if float(hr_58) > 0:
                        leave_type.append("holiday")
                        leave_time.append(hr_58)
                    if float(hr_62) > 0:
                        leave_type.append("guaranteed")
                        leave_time.append(hr_62)
                    if float(hr_86) > 0:
                        leave_type.append("other")
                        leave_time.append(hr_86)
                    if len(leave_type) > 1:
                        final_leave_type = "combo"
                        final_leave_time = float(hr_55) + float(hr_56) + float(hr_58) + float(hr_62) + float(hr_86)
                    elif len(leave_type) == 1:
                        final_leave_type = leave_type[0]
                        final_leave_time = leave_time[0]
                    else:
                        final_leave_type = ""
                        final_leave_time = 0.0
                if float(hr_55) > 1: code = "annual"  # alter CODE if annual leave was used
                if float(hr_56) > 1: code = "sick"  # alter code if sick leave was used
                # clear out non-5200 times
                hr_55 = 0.0  # annual leave
                hr_56 = 0.0  # sick leave
                hr_58 = 0.0  # holiday leave
                hr_62 = 0.0  # guaranteed time
                hr_86 = 0.0  # other paid leave
            if line[19] == "MV" and line[23][:3] == "722":  # get the RETURN TO OFFICE time
                rs = line[21]  # save the last occurrence.

            if line[19] in mv_codes:  # get the MOVES
                route_z = line[24].zfill(6)  # because some reports omit leading zeros
                route = route_z[1] + route_z[2] + route_z[4] + route_z[5]  # reformat route to 4 digit format
                mv_data = [line[19], line[21], line[23][:3], route]
                moves.append(mv_data)

        proto_array = [dayofweek, name, hr_52, rs, code, moves, final_leave_type, final_leave_time]  # form the proto array
        return (proto_array)  # send it back to auto weekly analysis()


def call_indexers(self):
    file_path = filedialog.askopenfilename(initialdir=os.getcwd(), filetypes=[("Excel files", "*.csv *.xls")])
    if file_path[-4:].lower() == ".csv" or file_path[-4:].lower() == ".xls":
        auto_indexer_1(self, file_path)
    else:
        messagebox.showerror("Report Generator", "The file you have selected is not a .csv or .xls file. "
                                                 "You must select a file with a .csv or .xls extension.")


def save_all(self):
    messagebox.showinfo("For Your Information ",
                        "All data has already been saved. Data is saved to the\n"
                        "database whenever an apply or submit button is pressed.\n"
                        "This button does nothing. :)",
                        parent=self)


def find_move_sets(moves):
    mv_sets = []
    pair = "closed"
    for line in moves:
        if line[3] == "off" and pair == "closed":
            mv_sets.append(line[1])
            pair = "open"
        if pair == "open":
            if line[3] == "":
                mv_sets.append(line[1])
                pair = "closed"


def ee_ns_detect(array):  # finds the ns day from ee reports
    days = ("Saturday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday")
    ns_candidates = ["Saturday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
    for d in days:
        hr_52 = 0  # straight hours
        hr_53 = 0  # overtime hours
        hr_43 = 0  # penalty hours
        for line in array:
            if line[18] in ns_candidates:
                ns_candidates.remove(line[18])
            if line[18] == d:
                spt_20 = line[20].split(':')  # split to get code and hours
                if spt_20[0] == "05200": hr_52 = spt_20[1]
                if spt_20[0] == "05300": hr_53 = spt_20[1]
                if spt_20[0] == "04300": hr_43 = spt_20[1]
        if float(hr_52) != 0:
            sum = float(hr_53) + float(hr_43)
            if float(hr_52) == round(sum, 2):
                return d
    if len(ns_candidates) == 1:
        return ns_candidates[0]


def ee_analysis(array, report):
    days = ("Saturday", "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday")
    hr_codes = ("52", "55", "56", "59", "60")
    code_dict = {"52": "total ", "55": "annual", "56": "sick  ", "59": "lwop  ", "60": "lwop  "}
    mv_codes = ("BT", "MV", "ET")
    moves_array = []
    for line in array:
        if line[19] and line[19] not in mv_codes and len(moves_array) > 0:
            find_move_sets(moves_array)  # call function to analyse moves
            del moves_array[:]
        if line[18] == "Base" and line[19] == "844" or line[18] == "Base" and line[
            19] == "134":  # find first line of specific carrier
            if line[19] == "844":
                list = "aux"
                route = ""
                ns_day = ""
            else:
                list = "FT"
                ns_day = ee_ns_detect(array)  # call function to find the ns day
                if line[23].zfill(2) == "01":
                    route = line[25].zfill(6)
                    route = route[1] + route[2] + route[4] + route[5]
                if line[23].zfill(2) == "02":
                    route = "floater"
            report.write("================================================\n")
            report.write(line[5].lower() + ", " + line[6].lower() + "\n")  # write name
            report.write(list + "\n")
            if list == "FT":
                report.write("route:" + route + "\n")
                if ns_day == None:
                    report.write("Klusterbox failed to detect ns day!")
                else:
                    report.write("ns day:" + ns_day + "\n")
            # report.write("================================================\n")
        if line[18] in days:
            spt_20 = line[20].split(':')  # split to get code and hours
            hr_type = spt_20[0][1] + spt_20[0][2]  # parse hour code to 2 digits
            if hr_type in hr_codes:  # compare to array of hour codes
                report.write("------------------------------------------------\n")
                if line[18] == ns_day:  # if the day is the ns day...
                    report.write("{}{}{}{}\n".format(line[18].ljust(12, " "), code_dict[hr_type].ljust(10, " "),
                                                     "{0:.2f}".format(float(spt_20[1])).ljust(6, " "),
                                                     "ns day".rjust(17, " ")))
                else:  # if the day is NOT the ns day...
                    report.write("{}{}{}\n".format(line[18].ljust(12, " "), code_dict[hr_type].ljust(10, " "),
                                                   "{0:.2f}".format(float(spt_20[1])).ljust(6, " ")))
                # report.write("------------------------------------------------\n")
        if line[19] in mv_codes and line[32] != "(W)Ring Deleted From PC":  # printe rings
            r_route = line[24].zfill(6)
            r_route = r_route[1] + r_route[2] + r_route[4] + r_route[5]  # reformat route to 4 digit format
            if route != r_route and list == "FT" and route != "floater" and r_route != "0000":
                off_route = "off"  # marker for off route work
            else:
                off_route = ""  # no marker for off route work
            # make array and call function to makes moves sets
            mv_data = (line[19], float(line[21]), move_translator(line[23][:-4]), off_route)
            moves_array.append(mv_data)
            report.write(
                "\t{}{}{}{}{}\n".format(line[19].ljust(2, " "), "{00:.2f}".format(float(line[21])).rjust(8, " "),
                                        move_translator(line[23][:-4]).rjust(12, " "), r_route.rjust(6, " "),
                                        off_route.rjust(6, " ")))
    if len(moves_array) > 0:
        # call function to analyse moves
        find_move_sets(moves_array)
        del moves_array[:]


def ee_skimmer():
    days = ("Saturday", "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday")
    mv_codes = ("BT", "MV", "ET")
    carrier = []
    file_path = filedialog.askopenfilename(initialdir=os.getcwd())
    if file_path[-4:].lower() == ".csv" or file_path[-4:].lower() == ".xls":
        with open(file_path, newline="") as file:
            a_file = csv.reader(file)
            c = 0
            good_id = "no"
            for line in a_file:
                if c == 0:
                    if line[0][:8] != "TAC500R3":
                        messagebox.showwarning("File Selection Error", "The selected file does not appear to be an "
                                                                       "Employee Everything report.")
                        return
                if c == 2:
                    pp = line[0]  # find the pay period
                    filename = "ee_reader" + "_" + pp + ".txt"
                    if os.path.isdir('kb_sub/ee_reader') == False:
                        os.makedirs('kb_sub/ee_reader')
                    try:
                        report = open('kb_sub/ee_reader/' + filename, "w")
                    except:
                        messagebox.showwarning("Report Generator", "The Employee Everything Report Reader "
                                                                   "was not generated.")
                        return
                    report.write("\nEmployee Everything Report Reader\n")
                    report.write(
                        "pay period: " + pp[:-3] + " " + pp[4] + pp[5] + "-" + pp[6] + "\n\n")  # printe pay period
                if c != 0:
                    if good_id != line[4] and good_id != "no":  # if new carrier or employee
                        ee_analysis(carrier, report)  # trigger analysis
                        del carrier[:]  # empty array
                        good_id = "no"  # reset trigger
                    if line[18] == "Base" and line[19] == "844" or line[
                        19] == "134":  # find first line of specific carrier
                        good_id = line[4]  # set trigger to id of carriers who are FT or aux carriers
                        carrier.append(line)  # gather times and moves for anaylsis
                    if good_id == line[4] and line[18] != "Base":
                        if line[18] in days:  # get the hours for each day
                            carrier.append(line)  # gather times and moves for anaylsis
                        if line[19] in mv_codes and line[32] != "(W)Ring Deleted From PC":
                            carrier.append(line)  # gather times and moves for anaylsis
                c += 1
            ee_analysis(carrier, report)  # when loop ends, run final analysis
            del carrier[:]  # empty array
            report.close()
            if sys.platform == "win32":
                os.startfile('kb_sub\\ee_reader\\' + filename)
            if sys.platform == "linux":
                subprocess.call(["xdg-open", 'kb_sub/ee_reader/' + filename])
            if sys.platform == "darwin":
                subprocess.call(["open", 'kb_sub/ee_reader/' + filename])
    else:
        messagebox.showerror("Report Generator", "The file you have selected is not a .csv or .xls file.\n"
                                                 "You must select a file with a .csv or .xls extension.")
        return


def pay_period_guide(self):
    year = simpledialog.askinteger("Pay Period Guide", "Enter the year you want generated.", parent=self,
                                   minvalue=2, maxvalue=9999)
    if year != None:
        firstday = datetime(1, 12, 22, 0, 0, 0)
        while int(firstday.strftime("%Y")) != year - 1:
            firstday += timedelta(weeks=52)
            if int(firstday.strftime("%m")) <= 12 and int(firstday.strftime("%d")) <= 12:
                firstday += timedelta(weeks=2)
        filename = "pp_guide" + "_" + str(year) + ".txt"  # create the filename for the text doc
        if os.path.isdir('kb_sub/pp_guide') == False:  # check to see if the folder exist
            os.makedirs('kb_sub/pp_guide')  # if not, then create the folder
        try:
            report = open('kb_sub/pp_guide/' + filename, "w")  # create the document
            report.write("\nPay Period Guide\n")
            report.write("Year: " + str(year) + "\n")
            report.write("---------------------------------------------\n\n")
            report.write("                 START (Sat):   END (Fri):         \n")
            for i in range(1, 27):
                # calculate dates
                wk1_start = firstday
                wk1_end = firstday + timedelta(days=6)
                wk2_start = firstday + timedelta(days=7)
                wk2_end = firstday + timedelta(days=13)
                report.write("PP: " + str(i).zfill(2) + "\n")
                report.write(
                    "\t week 1: " + wk1_start.strftime("%b %d, %Y") + " - " + wk1_end.strftime("%b %d, %Y") + "\n")
                report.write(
                    "\t week 2: " + wk2_start.strftime("%b %d, %Y") + " - " + wk2_end.strftime("%b %d, %Y") + "\n")
                # increment the first day by two weeks
                firstday += timedelta(days=14)
            # handle cases where there are 27 pay periods
            if int(firstday.strftime("%m")) <= 12 and int(firstday.strftime("%d")) <= 12:
                i += 1
                wk1_start = firstday
                wk1_end = firstday + timedelta(days=6)
                wk2_start = firstday + timedelta(days=7)
                wk2_end = firstday + timedelta(days=13)
                report.write("PP: " + str(i).zfill(2) + "\n")
                report.write(
                    "\t week 1: " + wk1_start.strftime("%b %d, %Y") + " - " + wk1_end.strftime("%b %d, %Y") + "\n")
                report.write(
                    "\t week 2: " + wk2_start.strftime("%b %d, %Y") + " - " + wk2_end.strftime("%b %d, %Y") + "\n")
            report.close()
            if sys.platform == "win32":
                os.startfile('kb_sub\\pp_guide\\' + filename)
            if sys.platform == "linux":
                subprocess.call(["xdg-open", 'kb_sub/pp_guide/' + filename])
            if sys.platform == "darwin":
                subprocess.call(["open", 'kb_sub/pp_guide/' + filename])
        except:
            messagebox.showerror("Report Generator", "The report was not generated.")


def pp_by_date(sat_range):  # returns a formatted pay period when given the starting date
    year = sat_range.strftime("%Y")
    pp_end = find_pp(int(year) + 1, "011")
    if sat_range >= pp_end:
        year = int(year) + 1
        year = str(year)
    firstday = find_pp(int(year), "011")
    pp_finder = {}
    for i in range(1, 27):
        # update the dictionary
        pp_finder[firstday] = str(i).zfill(2) + "1"
        pp_finder[firstday + timedelta(days=7)] = str(i).zfill(2) + "2"
        # increment the first day by two weeks
        firstday += timedelta(days=14)
    # in cases where there are 27 pay periods
    if int(firstday.strftime("%m")) <= 12 and int(firstday.strftime("%d")) <= 12:
        pp_finder[firstday] = "27" + "1"
        pp_finder[firstday + timedelta(days=7)] = "27" + "2"
    raw_pp = year.zfill(4) + pp_finder[sat_range]  # get the year/pp in a rough format
    return raw_pp[:-3] + "-" + raw_pp[4] + raw_pp[5] + "-" + raw_pp[6]  # return formatted year/pp


def find_pp(year, pp):  # returns the starting date of the pp when given year and pay period
    firstday = datetime(1, 12, 22, 0, 0, 0)
    while int(firstday.strftime("%Y")) != year - 1:
        firstday += timedelta(weeks=52)
        if int(firstday.strftime("%m")) <= 12 and int(firstday.strftime("%d")) <= 12:
            firstday += timedelta(weeks=2)
    pp_finder = {}
    for i in range(1, 27):
        # update the dictionary
        pp_finder[str(i).zfill(2) + "1"] = firstday
        pp_finder[str(i).zfill(2) + "2"] = firstday + timedelta(days=7)
        # increment the first day by two weeks
        firstday += timedelta(days=14)
    # handle cases where there are 27 pay periods
    if int(firstday.strftime("%m")) <= 12 and int(firstday.strftime("%d")) <= 12:
        pp_finder["27" + "1"] = firstday
        pp_finder["27" + "2"] = firstday + timedelta(days=7)
    return pp_finder[pp]


def move_translator(num):  # makes 721, 722 codes readable.
    move_xlr = {"721": "to office", "722": "to street", "354": "standby", "622": "to travel", "613": "steward"}
    if num in move_xlr:  # if the code is in the dictionary...
        return move_xlr[num]  # translate it
    else:  # if the code is not in the dictionary...
        return num  # just return the code


def max_hr():  # generates a report for 12/60 hour violations
    file_path = filedialog.askopenfilename(initialdir=os.getcwd(), filetypes=[("Excel files", "*.csv *.xls")])
    day_xlr = {"Saturday": "sat", "Sunday": "sun", "Monday": "mon", "Tuesday": "tue", "Wednesday": "wed",
               "Thursday": "thr", "Friday": "fri"}
    leave_xlr = {"49": "owcp   ", "55": "annual ", "56": "sick   ", "58": "holiday", "59": "lwop   ", "60": "lwop   "}
    max_hr = []
    max_aux_day = []
    max_ft_day = []
    extra_hours = []
    all_extra = []
    adjustment = []
    days = ("Saturday", "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday")
    day_hours = []
    if file_path[-4:].lower() == ".csv" or file_path[-4:].lower() == ".xls":
        with open(file_path, newline="") as file:
            a_file = csv.reader(file)
            c = 0
            good_id = "no"
            for line in a_file:
                if c == 0:
                    if line[0][:8] != "TAC500R3":
                        messagebox.showwarning("File Selection Error", "The selected file does not appear to be an "
                                                                       "Employee Everything report.")
                        return
                if c == 2: # on the second line
                    pp = line[0]  # find the pay period
                    pp = pp.strip()  # strip whitespace out of pay period information
                if c != 0: # on all but the first line
                    if line[18] == "Base" and good_id and len(day_hours) > 0:
                        # find fri hours for friday adjustment
                        fri_hrs = 0
                        for t in day_hours:  # get the friday hours
                            if t[3] == "Friday":
                                fri_hrs += float(t[2])
                        # find thu hours for thursday adjustment
                        thu_hrs = 0
                        for t in day_hours:  # find the thursday hours
                            if t[3] == "Thursday":
                                thu_hrs += float(t[2])
                        # find wed hours for wednesday adjustment
                        wed_hrs = 0
                        for t in day_hours:  # find the wednesday hours
                            if t[3] == "Wednesday":
                                wed_hrs += float(t[2])
                        # find the weekly total by adding daily totals
                        wkly_total = 0
                        for t in day_hours:
                            wkly_total += float(t[2])
                        if wkly_total > 60:
                            add_maxhr = (day_hours[0][0].lower(), day_hours[0][1].lower(), wkly_total)
                            max_hr.append(add_maxhr)
                            for item in extra_hours:  # get any extra hours codes for non-5200 hours list
                                all_extra.append(item)
                            # find the all adjustments
                            if FT == True:
                                # find friday adjustment
                                fri_post_60 = float(wkly_total - 60)
                                if fri_hrs > 12:
                                    fri_over = fri_hrs - 12
                                    if fri_over < fri_post_60:
                                        fri_adj = fri_over
                                    else:
                                        fri_adj = fri_post_60
                                    add_adjustment = ("fri", day_hours[0][0].lower(), day_hours[0][1].lower(), fri_adj)
                                    adjustment.append(add_adjustment)
                                # find the thursday adjustment
                                thu_post_60 = float(wkly_total - 60) - fri_hrs
                                if thu_hrs > 12 and thu_post_60 > 0:
                                    thu_over = thu_hrs - 12
                                    if thu_over < thu_post_60:
                                        thu_adj = thu_over
                                    else:
                                        thu_adj = thu_post_60
                                    add_adjustment = ("thu", day_hours[0][0].lower(), day_hours[0][1].lower(), thu_adj)
                                    adjustment.append(add_adjustment)
                                # find the wednesday adjustment
                                wed_post_60 = float(wkly_total - 60) - fri_hrs - thu_hrs
                                if wed_hrs > 12 and wed_post_60 > 0:
                                    wed_over = wed_hrs - 12
                                    if wed_over < wed_post_60:
                                        wed_adj = wed_over
                                    else:
                                        wed_adj = wed_post_60
                                    add_adjustment = (
                                        "wed", day_hours[0][0].lower(), day_hours[0][1].lower(), wed_adj)
                                    adjustment.append(add_adjustment)
                        del day_hours[:]
                        del extra_hours[:]
                    if line[18] == "Base" and line[19] == "844" or line[
                        19] == "134":  # find first line of specific carrier
                        good_id = line[4]  # remember id of carriers who are FT or aux carriers
                        if line[19] == "844":
                            FT = False
                        else:
                            FT = True
                    if good_id == line[4] and line[18] != "Base":
                        if line[18] in days:  # get the hours for each day
                            spt_20 = line[20].split(':')  # split to get code and hours
                            hr_type = spt_20[0][1] + spt_20[0][2]  # parse hour code to 2 digits
                            # if hr_type in hr_codes:  # compare to array of hour codes
                            if hr_type == "52":  # compare to array of hour codes
                                if float(spt_20[1]) > 11.5 and FT == False:
                                    add_max_aux = (line[5].lower(), line[6].lower(), line[18], spt_20[1])
                                    max_aux_day.append(add_max_aux)
                                if float(spt_20[1]) > 12 and FT == True:
                                    add_max_ft = (line[5].lower(), line[6].lower(), line[18], spt_20[1])
                                    max_ft_day.append(add_max_ft)
                                if FT == True:  # increment daily totals to find weekly total
                                    add_day_hours = (line[5].lower(), line[6].lower(), spt_20[1], line[18])
                                    day_hours.append(add_day_hours)
                            extra_hour_codes = ("49", "55", "56", "58")  # paid leave types only , (lwop "59", "60")
                            if hr_type in extra_hour_codes and FT == True:  # if there is holiday pay
                                add_day_hours = (line[5].lower(), line[6].lower(), spt_20[1], line[18])
                                day_hours.append(add_day_hours)
                                add_extra_hours = (line[5].lower(), line[6].lower(), line[18], hr_type, spt_20[1])
                                extra_hours.append(add_extra_hours)  # track non 5200 hours
                c += 1
    elif file_path == "":
        return
    else:
        messagebox.showerror("Report Generator", "The file you have selected is not a .csv or .xls file.\n"
                                                 "You must select a file with a .csv or .xls extension.")
        return
    # find the weekly total by adding daily totals for last carrier
    if len(day_hours) > 0:
        wkly_total = 0
        for t in day_hours:
            wkly_total += float(t[2])
        if wkly_total > 60:
            add_maxhr = (day_hours[0][0].lower(), day_hours[0][1].lower(), wkly_total)
            max_hr.append(add_maxhr)
            for item in extra_hours:  # get any extra hours codes for non-5200 hours list
                all_extra.append(item)
        del day_hours[:]
        del extra_hours[:]

    if len(max_hr) == 0 and len(max_ft_day) == 0 and len(max_aux_day) == 0:
        messagebox.showwarning("Report Generator", "No violations were found. "
                                                   "The report was not generated.")
        return
    weekly_max = []  # array hold each carrier's hours for the week
    daily_max = []  # array hold each carrier's sum of maximum daily hours for the week
    if len(max_hr) > 0 or len(max_ft_day) > 0 or len(max_aux_day) > 0:
        pp_str = pp[:-3] + "_" + pp[4] + pp[5] + "_" + pp[6]
        filename = "max" + "_" + pp_str + ".txt"
        if os.path.isdir('kb_sub/over_max') == False:
            os.makedirs('kb_sub/over_max')
        try:
            report = open('kb_sub/over_max/' + filename, "w")
            report.write("12 and 60 Hour Violations Report\n\n")
            report.write("pay period: " + pp[:-3] + " " + pp[4] + pp[5] + "-" + pp[6] + "\n")  # printe pay period
            pp_date = find_pp(int(pp[:-3]), pp[-3:])  # send year and pp to get the date
            pp_date_end = pp_date + timedelta(days=6)  # add six days to get the last part of the range
            report.write(
                "week of: " + pp_date.strftime("%x") + " - " + pp_date_end.strftime("%x") + "\n")  # printe date
            report.write("\n60 hour violations \n\n")
            report.write("name                              total   over\n")
            report.write("-----------------------------------------------\n")
            if len(max_hr) == 0:
                report.write("no violations" + "\n")
            else:
                diff_total = 0
                max_hr.sort(key=itemgetter(0))
                for item in max_hr:
                    tabs = 30 - (len(item[0]))
                    period = "."
                    period = period + (tabs * ".")
                    diff = float(item[2]) - 60
                    diff_total = diff_total + diff
                    report.write(item[0] + ", " + item[1] + period + "{0:.2f}".format(float(item[2]))
                                 + "   " + "{0:.2f}".format(float(diff)).rjust(5, " ") + "\n")
                    wmax_add = (item[0], item[1], diff)
                    weekly_max.append(wmax_add)  # catch totals of violations for the week
                report.write("\n" + "                                   total:  " + "{0:.2f}".format(float(diff_total))
                             + "\n")
            all_extra.sort(key=itemgetter(0))
            report.write("\nNon 5200 codes contributing to 60 hour violations  \n\n")
            report.write("day   name                            hr type   hours\n")
            report.write("-----------------------------------------------------\n")
            if len(all_extra) == 0: report.write("no contributions" + "\n")
            for i in range(len(all_extra)):
                tabs = 28 - (len(all_extra[i][0]))
                period = "."
                period = period + (tabs * ".")
                report.write(day_xlr[all_extra[i][2]] + "   " + all_extra[i][0] + ", " + all_extra[i][1] + period +
                             leave_xlr[all_extra[i][3]] + "  " + "{0:.2f}".format(float(all_extra[i][4])).rjust(5, " ")
                             + "\n")
            report.write("\n\n12 hour full time carrier violations \n\n")
            report.write("day   name                        total   over   sum\n")
            report.write("-----------------------------------------------------\n")
            if len(max_ft_day) == 0: report.write("no violations" + "\n")
            diff_sum = 0
            sum_total = 0
            max_ft_day.sort(key=itemgetter(0))
            for i in range(len(max_ft_day)):
                jump = "no"  # triggers an analysis of the candidates array
                diff = float(max_ft_day[i][3]) - 12
                diff_sum = diff_sum + diff
                if i != len(max_ft_day) - 1:  # if the loop has not reached the end of the list
                    # if the name current and next name are the same
                    if max_ft_day[i][0] == max_ft_day[i + 1][0] and max_ft_day[i][1] == max_ft_day[i + 1][1]:
                        jump = "yes"  # bypasses an analysis of the candidates array
                        tabs = 24 - (len(max_ft_day[i][0]))
                        period = "."
                        period = period + (tabs * ".")
                        report.write(day_xlr[max_ft_day[i][2]] + "   " + max_ft_day[i][0] + ", " + max_ft_day[i][1] +
                                     period + "{0:.2f}".format(
                            float(max_ft_day[i][3])) + "   " + "{0:.2f}".format(float(diff)) + "\n")
                if jump == "no":
                    tabs = 24 - (len(max_ft_day[i][0]))
                    period = "."
                    period = period + (tabs * ".")
                    report.write(day_xlr[max_ft_day[i][2]] + "   " + max_ft_day[i][0] + ", " + max_ft_day[i][1] + period
                                 + "{0:.2f}".format(float(max_ft_day[i][3])) + "   " + "{0:.2f}".format(float(diff)) +
                                 "   " + "{0:.2f}".format(float(diff_sum)) + "\n")
                    dmax_add = (max_ft_day[i][0], max_ft_day[i][1], diff_sum)
                    daily_max.append(dmax_add)  # catch sum of daily violations for the week
                    sum_total = sum_total + diff_sum
                    diff_sum = 0
            report.write("\n" + "                                         total:  " + "{0:.2f}".format(float(sum_total))
                         + "\n")
            report.write("\n11.50 hour auxiliary carrier violations \n\n")
            report.write("day   name                        total   over   sum\n")
            report.write("-----------------------------------------------------\n")
            if len(max_aux_day) == 0: report.write("no violations" + "\n")
            diff_sum = 0
            sum_total = 0
            max_aux_day.sort(key=itemgetter(0))
            for i in range(len(max_aux_day)):
                jump = "no"  # triggers an analysis of the candidates array
                diff = float(max_aux_day[i][3]) - 11.5
                diff_sum = diff_sum + diff
                if i != len(max_aux_day) - 1:  # if the loop has not reached the end of the list
                    # if the current and next name are the same
                    if max_aux_day[i][0] == max_aux_day[i + 1][0] and max_aux_day[i][1] == max_aux_day[i + 1][1]:
                        jump = "yes"  # bypasses an analysis of the candidates array
                        tabs = 24 - (len(max_aux_day[i][0]))
                        period = "."
                        period = period + (tabs * ".")
                        report.write(day_xlr[max_aux_day[i][2]] + "   " + max_aux_day[i][0] + ", "
                                     + max_aux_day[i][1] + period + "{0:.2f}".format(float(max_aux_day[i][3]))
                                     + "   " + "{0:.2f}".format(float(diff)) + "\n")
                if jump == "no":
                    tabs = 24 - (len(max_aux_day[i][0]))
                    period = "."
                    period = period + (tabs * ".")
                    report.write(day_xlr[max_aux_day[i][2]] + "   " + max_aux_day[i][0] + ", "
                                 + max_aux_day[i][1] + period + "{0:.2f}".format(float(max_aux_day[i][3]))
                                 + "   " + "{0:.2f}".format(float(diff)) + "   " + "{0:.2f}".format(float(diff_sum))
                                 + "\n")
                    dmax_add = (max_aux_day[i][0], max_aux_day[i][1], diff_sum)
                    daily_max.append(dmax_add)  # catch sum of daily violations for the week
                    sum_total = sum_total + diff_sum
                    diff_sum = 0
            report.write(
                "\n" + "                                         total:  " + "{0:.2f}".format(float(sum_total)) + "\n")
            weekly_and_daily = []
            d_max_remove = []
            w_max_remove = []
            # find the write the adjustments
            # get the adjustment
            adjustment.sort(key=itemgetter(1))
            adj_sum = 0
            adj_total = []
            report.write("\nPost 60 Hour Adjustments \n\n")
            report.write("day   name                   daily adj    total\n")
            report.write("-----------------------------------------------\n")
            if len(adjustment) == 0: report.write("no adjustments" + "\n")
            for i in range(len(adjustment)):
                jump = "no"  # triggers an analysis of the adjustment array
                adj_sum = adj_sum + adjustment[i][3]
                if i != len(adjustment) - 1:  # if the loop has not reached the end of the list
                    # if the current and next name are the same
                    if adjustment[i][1] == adjustment[i + 1][1] and adjustment[i][2] == adjustment[i + 1][2]:
                        jump = "yes"  # bypasses an analysis of the candidates array
                        tabs = 24 - (len(adjustment[i][1]))
                        period = "."
                        period = period + (tabs * ".")
                        report.write(adjustment[i][0] + "   " + adjustment[i][1] + ", "
                                     + adjustment[i][2] + period + "{0:.2f}".format(float(adjustment[i][3])) + "\n")
                if jump == "no":
                    tabs = 24 - (len(adjustment[i][1]))
                    period = "."
                    period = period + (tabs * ".")
                    report.write(adjustment[i][0] + "   " + adjustment[i][1] + ", "
                                 + adjustment[i][2] + period + "{0:.2f}".format(float(adjustment[i][3]))
                                 + "     " + "{0:.2f}".format(float(adj_sum))
                                 + "\n")
                    adj_add = [adjustment[i][1], adjustment[i][2], adj_sum]
                    adj_sum = 0
                    adj_total.append(adj_add)  # catch sum of adjustments for the week
            for w_max in weekly_max:  # find the total violation
                for d_max in daily_max:
                    if w_max[0] + w_max[1] == d_max[0] + d_max[
                        1]:  # look for names with both weekly and daily violations
                        wk_dy_sum = w_max[2] + d_max[2]  # add the weekly and daily
                        to_add = [w_max[0], w_max[1], wk_dy_sum]
                        weekly_and_daily.append(to_add)
                        d_max_remove.append(d_max)
                        w_max_remove.append(w_max)
            weekly_max = [x for x in weekly_max if x not in w_max_remove]
            daily_max = [x for x in daily_max if x not in d_max_remove]
            d_max_remove = []
            w_max_remove = []
            for d_max in daily_max:
                for w_max in weekly_max:
                    if w_max[0] + w_max[1] == d_max[0] + d_max[1]:  # if the names match
                        wk_dy_sum = w_max[2] + d_max[2]  # add the weekly and daily
                        to_add = [w_max[0], w_max[1], wk_dy_sum]
                        weekly_and_daily.append(to_add)
                        d_max_remove.append(d_max)
                        w_max_remove.append(w_max)
            weekly_max = [x for x in weekly_max if x not in w_max_remove]  # remove
            daily_max = [x for x in daily_max if x not in d_max_remove]
            joint_max = (weekly_max + daily_max + weekly_and_daily)  # add all arrays to get the final array
            joint_max.sort(key=itemgetter(0, 1))
            for j in joint_max:  # cycle through the totals and adjustments
                for a in adj_total:
                    if j[0] + j[1] == a[0] + a[1]:  # if the names match
                        j[2] = j[2] - a[2]  # subtract the adjustment from the total
            report.write("\n\nTotal of the two violations (with adjustments)\n\n")
            report.write("name                              total\n")
            report.write("---------------------------------------\n")
            if len(joint_max) == 0: report.write("no violations" + "\n")
            great_total = 0
            for item in joint_max:
                tabs = 30 - (len(item[0]))
                period = "."
                period = period + (tabs * ".")
                great_total = great_total + item[2]
                report.write(item[0] + ", " + item[1] + period + "{0:.2f}".format(float(item[2])).rjust(5, ".") + "\n")
            report.write(
                "\n" + "                           total:  " + "{0:.2f}".format(float(great_total)) + "\n")
            report.close()
            if sys.platform == "win32":
                os.startfile('kb_sub\\over_max\\' + filename)
            if sys.platform == "linux":
                subprocess.call(["xdg-open", 'kb_sub/over_max/' + filename])
            if sys.platform == "darwin":
                subprocess.call(["open", 'kb_sub/over_max/' + filename])
        except:
            messagebox.showerror("Report Generator", "The report was not generated.")


def file_dialogue(folder):  # opens file folders to access generated reports
    if os.path.isdir(folder) == False:
        os.makedirs(folder)
    file_path = filedialog.askopenfilename(initialdir=os.getcwd() + "/" + folder)
    if file_path:
        if sys.platform == "win32":
            os.startfile(file_path)
        if sys.platform == "linux":
            subprocess.call(["xdg-open", file_path])
        if sys.platform == "darwin":
            subprocess.call([opener, file_path])

def remove_file(folder): # removes a file and all contents
    if os.path.isdir(folder) == True:
        shutil.rmtree(folder)

def remove_file_var(folder): # removes a file and all contents
    folder_name = folder.split("/")
    folder_name = folder_name[1]
    if os.path.isdir(folder) == True:
        if messagebox.askokcancel("Delete Folder Contents",
            "This will delete all the files in the {} archive. ".format(folder_name)):
            shutil.rmtree(folder)
        if os.path.isdir(folder) == False:
            messagebox.showinfo("Delete Folder Contents",
                "Success! All the files in the {} archive have been deleted.".format(folder_name))

    else:
        messagebox.showwarning("Delete Folder Contents", "The {} folder is already empty".format(folder_name))


def location_klusterbox(self):  # provides the location of the program
    messagebox.showinfo("KLUSTERBOX ",
                        "On this computer Klusterbox is located at:\n"
                        "{}".format(os.getcwd()), parent=self)


def about_klusterbox(self):  # gives information about the program
    self.destroy()
    F = Frame(root)
    F.pack(fill=BOTH, side=LEFT)
    C1 = Canvas(F)
    C1.pack(fill=BOTH, side=BOTTOM)
    # apply and close buttons
    Button(C1, text="Go Back", width=20, anchor="w",
           command=lambda: [F.destroy(), main_frame()]).pack(side=LEFT)
    # link up the canvas and scrollbar
    S = Scrollbar(F)
    C = Canvas(F, width=1600)
    S.pack(side=RIGHT, fill=BOTH)
    C.pack(side=LEFT, fill=BOTH, pady=10, padx=20)
    S.configure(command=C.yview, orient="vertical")
    C.configure(yscrollcommand=S.set)
    if sys.platform == "win32":
        C.bind_all('<MouseWheel>', lambda event: C.yview_scroll(int(-1 * (event.delta / 120)), "units"))
    elif sys.platform == "linux":
        C.bind_all('<Button-4>', lambda event: C.yview('scroll', -1, 'units'))
        C.bind_all('<Button-5>', lambda event: C.yview('scroll', 1, 'units'))
    # create the frame inside the canvas
    FF = Frame(C)
    C.create_window((0, 0), window=FF, anchor=NW)
    # page contents
    try:
        photo = ImageTk.PhotoImage(Image.open("kb_sub/kb_images/kb_about.jpg"))
        Label(FF, image=photo).pack(fill=X)
    except:
        pass
    Label(FF, text="Klusterbox", font="bold", fg="red", anchor=W).pack(fill=X)
    Label(FF, text="version: {}".format(version), anchor=W).pack(fill=X)
    Label(FF, text="release date: {}".format(release_date), anchor=W).pack(fill=X)
    Label(FF, text="created by Thomas Weeks", anchor=W).pack(fill=X)
    Label(FF, text="Original release: October 2018", anchor=W).pack(fill=X)
    Label(FF, text=" ", anchor=W).pack(fill=X)
    Label(FF, text="comments and criticisms are welcome", anchor=W, fg="blue").pack(fill=X)
    Label(FF, text=" ", anchor=W).pack(fill=X)
    Label(FF, text="contact information: ", anchor=W).pack(fill=X)
    Label(FF, text="", anchor=W).pack(fill=X)
    Label(FF, text="Thomas Weeks", anchor=W).pack(fill=X)
    Label(FF, text="tomandsusan4ever@msn.com", anchor=W).pack(fill=X)
    Label(FF, text="(please put \"klusterbox\" in the subject line", anchor=W).pack(fill=X)
    Label(FF, text="720.280.0415", anchor=W).pack(fill=X)
    root.update()
    C.config(scrollregion=C.bbox("all"))
    FF.mainloop()


def apply_startup(switch, station, self):
    global list_of_stations
    if switch == "enter":
        if station.get().strip() == "":
            messagebox.showerror("Prohibited Action",
                                 "You can not enter a blank entry for a station.", parent=self)
            return
        sql = "INSERT INTO stations (station) VALUES('%s')" % (station.get().strip())
        commit(sql)
        list_of_stations.append(station.get())
    # access list of stations from database
    sql = "SELECT * FROM stations ORDER BY station"
    results = inquire(sql)
    # define and populate list of stations variable
    del list_of_stations[:]
    for stat in results:
        list_of_stations.append(stat[0])
    self.destroy()  # destroy old frame
    main_frame()  # load new frame


def start_up():  # the start up screen when no information has been entered
    # put records in the skippers table
    skip_these = [["354", "stand by"], ["613", "stewards time"], ["743", "route maintenance"]]
    for rec in skip_these:
        sql = "INSERT OR IGNORE INTO skippers(code, description) VALUES ('%s','%s')" % (rec[0], rec[1])
        commit(sql)
    F = Frame(root)
    F.pack(fill=BOTH, side=LEFT, pady=10, padx=20)
    C = Canvas(F, width=1600)
    C.pack(side=LEFT, fill=BOTH)
    FF = Frame(C)
    C.create_window((0, 0), window=FF, anchor=NW)
    Label(FF, text="Welcome to Klusterbox", font="bold").grid(row=0, columnspan=2, sticky="w")
    Label(FF, text="version: {}".format(version)).grid(row=1, columnspan=2, sticky="w")
    Label(FF, text="", pady=20).grid(row=2, column=0)
    # enter new stations
    new_station = StringVar(FF)
    Label(FF, text="To get started, please enter your station name:", pady=5).grid(row=3, columnspan=2, sticky="w")
    e = Entry(FF, width=35, textvariable=new_station)
    e.grid(row=4, column=0, sticky="w")
    new_station.set("")
    Button(FF, width=5, anchor="w", text="ENTER", command=lambda: apply_startup("enter", new_station, F)). \
        grid(row=4, column=1, sticky="w")
    Label(FF, text="", pady=20).grid(row=5, columnspan=2, sticky="w")
    Label(FF, text="Or you can exit to the main screen and enter your\n"
                   "station by going to Configuration > list of stations.").grid(row=6, columnspan=2, sticky="w")
    Button(FF, width=5, text="EXIT", command=lambda: [F.destroy(), main_frame()]). \
        grid(row=7, columnspan=2, sticky="e")
    root.update()
    mainloop()


def carrier_list_cleaning_for_auto_skimmer():  # cleans the database of duplicate records
    sql = "SELECT * FROM carriers ORDER BY carrier_name, effective_date"
    results = inquire(sql)
    duplicates = []
    for i in range(len(results)):
        if i != len(results) - 1:  # if the loop has not reached the end of the list
            if results[i][1] == results[i + 1][1] and \
                    results[i][2] == results[i + 1][2] and \
                    results[i][3] == results[i + 1][3] and \
                    results[i][4] == results[i + 1][4] and \
                    results[i][5] == results[i + 1][5]:  # if the name current and next name are the same
                duplicates.append(i + 1)
    if len(duplicates) > 0:
        pb_root = Tk()  # create a window for the progress bar
        pb_root.title("Database Maintenance")
        pb_label = Label(pb_root, text="Updating Changes: ")  # make label for progress bar
        pb_label.pack(side=LEFT)
        pb = ttk.Progressbar(pb_root, length=400, mode="determinate")  # create progress bar
        pb.pack(side=LEFT)
        pb["maximum"] = len(duplicates)  # set length of progress bar
        pb.start()
        i = 0
        for d in duplicates:
            pb["value"] = i  # increment progress bar
            sql = "DELETE FROM carriers WHERE effective_date='%s' and carrier_name='%s'" % (
            results[d][0], results[d][1])
            commit(sql)
            pb_root.update()
            i += 1
        pb.stop()  # stop and destroy the progress bar
        pb_label.destroy()  # destroy the label for the progress bar
        pb.destroy()
        pb_root.destroy()
        messagebox.showinfo("Database Maintenance", "All redundancies have been eliminated from the carrier list.")
    del duplicates[:]


def carrier_list_cleaning(self):  # cleans the database of duplicate records
    sql = "SELECT * FROM carriers ORDER BY carrier_name, effective_date"
    results = inquire(sql)
    duplicates = []
    for i in range(len(results)):
        if i != len(results) - 1:  # if the loop has not reached the end of the list
            if results[i][1] == results[i + 1][1] and \
                    results[i][2] == results[i + 1][2] and \
                    results[i][3] == results[i + 1][3] and \
                    results[i][4] == results[i + 1][4] and \
                    results[i][5] == results[i + 1][5]:  # if the name current and next name are the same
                duplicates.append(i + 1)
    ok = False
    if len(duplicates) > 0:
        ok = messagebox.askokcancel("Database Maintenance", "Did you want to eliminate database redundancies? \n"
                                                            "{} redundancies have been found in the database \n"
                                                            "This is recommended maintenance.".format(len(duplicates)))
    if ok == True:
        pb_root = Tk()  # create a window for the progress bar
        pb_root.title("Database Maintenance")
        pb_label = Label(pb_root, text="Updating Changes: ")  # make label for progress bar
        pb_label.pack(side=LEFT)
        pb = ttk.Progressbar(pb_root, length=400, mode="determinate")  # create progress bar
        pb.pack(side=LEFT)
        pb["maximum"] = len(duplicates)  # set length of progress bar
        pb.start()
        i = 0
        for d in duplicates:
            pb["value"] = i  # increment progress bar
            sql = "DELETE FROM carriers WHERE effective_date='%s' and carrier_name='%s'" % (
            results[d][0], results[d][1])
            commit(sql)
            pb_root.update()
            i += 1
        pb.stop()  # stop and destroy the progress bar
        pb_label.destroy()  # destroy the label for the progress bar
        pb.destroy()
        pb_root.destroy()
        messagebox.showinfo("Database Maintenance", "All redundancies have been eliminated from the carrier list.")
        self.destroy()
        main_frame()
    if ok == False: messagebox.showinfo("Database Maintenance", "No redundancies have been found in the carrier list.")
    del duplicates[:]


def isfloat(value):
    try:
        float(value)
        return True
    except ValueError:
        return False


def isint(value):
    try:
        int(value)
        return True
    except ValueError:
        return False


def data_mods_codes_delete(frame, to_delete):
    sql = "DELETE FROM skippers WHERE code='%s'" % to_delete[0]
    commit(sql)
    auto_data_entry_settings(frame)


def data_mods_codes_add(frame, code, description):
    sql = "SELECT code FROM skippers"
    results = inquire(sql)
    existing_codes = []
    for item in results:
        existing_codes.append(item[0])
    prohibited_codes = ('721', '722')
    if code.get() in prohibited_codes:
        messagebox.showerror("Data Entry Error", "It is prohibited to exclude code {}".format(code.get()))
        return
    if code.get() in existing_codes:
        messagebox.showerror("Data Entry Error", "This code had already been entered.")
        return
    if code.get().isdigit() == FALSE:
        messagebox.showerror("Data Entry Error", "TACS code must contain only numbers.")
        return
    if len(code.get()) > 3 or len(code.get()) < 3:
        messagebox.showerror("Data Entry Error", "TACS code must be 3 digits long.")
        return
    if len(description.get()) > 39:
        messagebox.showerror("Data Enty Error", "Please limit description to less than 40 characters.")
        return
    sql = "INSERT INTO skippers(code,description) VALUES('%s','%s')" % (code.get(), description.get())
    commit(sql)
    auto_data_entry_settings(frame)


def data_mods_codes_default(frame):
    sql = "DELETE FROM skippers"
    commit(sql)
    # put records in the skippers table
    skip_these = [["354", "stand by"], ["613", "stewards time"], ["743", "route maintenance"]]
    for rec in skip_these:
        sql = "INSERT OR IGNORE INTO skippers(code, description) VALUES ('%s','%s')" % (rec[0], rec[1])
        commit(sql)
    auto_data_entry_settings(frame)

def apply_auto_ns_structure(frame, ns_structure):
    sql = "UPDATE tolerances SET tolerance='%s'WHERE category='%s'" % (ns_structure.get(), "ns_auto_pref")
    commit(sql)
    messagebox.showinfo("Settings Updated", "Auto Data Entry settings have been updated.")

def data_entry_permit_zero(frame, top, bottom):
    sql = "UPDATE tolerances SET tolerance='%s'WHERE category='%s'" % (top.get(), "allow_zero_top")
    commit(sql)
    sql = "UPDATE tolerances SET tolerance='%s'WHERE category='%s'" % (bottom.get(), "allow_zero_bottom")
    commit(sql)
    messagebox.showinfo("Settings Updated", "Auto Data Entry settings have been updated.")


def auto_data_entry_settings(frame):
    wd = front_window(frame)  # F,S,C,FF,buttons
    r = 0
    Label(wd[3], text="Auto Data Entry Settings", font="bold").grid(row=r, column=0, sticky="w", columnspan=4)
    r += 1
    Label(wd[3], text="").grid(row=r, column=1)
    r += 1
    Label(wd[3], text="NS Day Structure", font="bold").grid(row=r, column=0, columnspan=4, sticky="w")
    r += 1
    ns_structure = StringVar(wd[3])
    sql = "SELECT tolerance FROM tolerances WHERE category='%s'" % "ns_auto_pref"
    result = inquire(sql)
    Radiobutton(wd[3], text="rotation", variable=ns_structure, value="rotation").grid(row=r, column=1, sticky="e")
    Radiobutton(wd[3], text="fixed", variable=ns_structure, value="fixed").grid(row=r, column=2, sticky="w")
    ns_structure.set(result[0][0])
    r += 1
    Button(wd[3], text="Set", width=5, command=lambda:apply_auto_ns_structure(wd[0],ns_structure)).grid(row=r, column=3)
    r += 1
    Label(wd[3], text="List of TACS MODS Codes", font="bold").grid(row=r, column=0, columnspan=4, sticky="w")
    r += 1
    Label(wd[3], text="(to exclude from Auto Data Entry moves).") \
        .grid(row=r, column=0, columnspan=4, sticky="w")
    r += 1
    Label(wd[3], text="code", fg="grey", anchor="w").grid(row=r, column=0)
    Label(wd[3], text="description", fg="grey", anchor="w").grid(row=r, column=1, columnspan=2)
    sql = "SELECT * FROM skippers"
    results = inquire(sql)
    r += 1
    if len(results) > 0:
        for i in range(len(results)):
            Button(wd[3], text=results[i][0], anchor="w", width=5).grid(row=i + r, column=0)  # display code
            Button(wd[3], text=results[i][1], anchor="w", width=30).grid(row=i + r, column=1,
                                                                         columnspan=2)  # display description
            Button(wd[3], text="delete", command=lambda x=i: data_mods_codes_delete(wd[0], results[x])).grid(row=i + r,
                                                                                                             column=3)
    else:
        Label(wd[3], text="No Exceptions Listed.", anchor="w").grid(row=r, column=0, sticky="w", columnspan=3)
        i = 1
    r = r + i
    r += 1
    Label(wd[3], text="").grid(row=r, column=2)
    r += 1
    Label(wd[3], text="Add New Code", font="bold").grid(row=r, column=0, columnspan=3,
                                                        sticky="w")  # add new code labels
    r += 1
    new_code = StringVar(wd[3])
    new_descp = StringVar(wd[3])
    Label(wd[3], text="code", fg="grey", anchor="w").grid(row=r, column=0)
    Label(wd[3], text="description", fg="grey", anchor="w").grid(row=r, column=1, columnspan=2)
    r += 1
    Entry(wd[3], textvariable=new_code, width=6).grid(row=r, column=0)  # add new code
    Entry(wd[3], textvariable=new_descp, width=35).grid(row=r, column=1, columnspan=2)
    Button(wd[3], text="Add", width=5, command=lambda: data_mods_codes_add(wd[0], new_code, new_descp)) \
        .grid(row=r, column=3)
    r += 1
    Label(wd[3], text="").grid(row=r, column=0)
    r += 1
    Label(wd[3], text="Restore Defaults").grid(row=r, column=1, columnspan=2, sticky="e")
    Button(wd[3], text="Set", width=5, command=lambda: data_mods_codes_default(wd[0])).grid(row=r, column=3)
    r += 1
    Label(wd[3], text="").grid(row=r, column=0)
    r += 1
    sql = "SELECT tolerance FROM tolerances WHERE category='%s'" % "allow_zero_top"
    result_top = inquire(sql)
    sql = "SELECT tolerance FROM tolerances WHERE category='%s'" % "allow_zero_bottom"
    result_bottom = inquire(sql)
    Label(wd[3], text="Permit Zero Sums", font="bold").grid(row=r, column=0, columnspan=2, sticky="w")
    text = "Selecting 'allow' will permit entries into moves where the MOVE OFF and MOVE ON times are the same. " \
           "While these entries do not add to the total for Overtime Worked Off route, they might indicate something " \
           "that would merit further investigation. You can always delete them manually. Selecting 'don't allow' will " \
           "hide these entries.\n'Top' refers to the start of the workday and 'Bottom' refers to the end of the workday."
    Button(wd[3], text="info", width=5, command=lambda: messagebox.showinfo("For Your Information", text)) \
        .grid(row=r, column=3)
    zero_top = BooleanVar(wd[3])
    zero_bottom = BooleanVar(wd[3])
    r += 1
    Label(wd[3], text="Allow Zero Sums on the Top").grid(row=r, column=0, sticky="w", columnspan=3)
    r += 1
    Radiobutton(wd[3], text="allow", variable=zero_top, value=True).grid(row=r, column=1, sticky="e")
    Radiobutton(wd[3], text="don't allow", variable=zero_top, value=False).grid(row=r, column=2, sticky="w")
    zero_top.set(result_top[0][0])
    r += 1
    Label(wd[3], text="Allow Zero Sum On Bottom").grid(row=r, column=0, sticky="w", columnspan=3)
    r += 1
    Radiobutton(wd[3], text="allow", variable=zero_bottom, value=True).grid(row=r, column=1, sticky="e")
    Radiobutton(wd[3], text="don't allow", variable=zero_bottom, value=False).grid(row=r, column=2, sticky="w")
    zero_bottom.set(result_bottom[0][0])
    r += 1
    Button(wd[3], text="Set", width=5, command=lambda: data_entry_permit_zero(wd[0], zero_top, zero_bottom)) \
        .grid(row=r, column=0, columnspan=4, sticky="e")

    Button(wd[4], text="Go Back", width=20, command=lambda: (wd[0].destroy(), main_frame())).grid(row=0, column=0,
                                                                                                  sticky="w")
    rear_window(wd)


def min_ss_presets(frame, order):
    if order == "default": num = "25"
    if order == "zero": num = "0"
    types = ("min_ss_nl", "min_ss_wal", "min_ss_otdl", "min_ss_aux")
    for t in types:
        sql = "UPDATE tolerances SET tolerance ='%s' WHERE category = '%s'" % (num, t)
        commit(sql)
    spreadsheet_settings(frame)


def apply_ss_min(frame, tolerance, type):
    if isint(tolerance) == False:
        text = "You must enter a number with no decimals. "
        messagebox.showerror("Tolerance value entry error", text, parent=frame)
        return
    if tolerance.strip() == "":
        text = "You must enter a numeric value for tolerances"
        messagebox.showerror("Tolerance value entry error", text, parent=frame)
        return
    if float(tolerance) < 0:
        text = "Values must be equal to or greater than zero."
        messagebox.showerror("Tolerance value entry error", text, parent=frame)
        return
    if float(tolerance) > 100:
        text = "You must enter a value less than one-hundred."
        messagebox.showerror("Tolerance value entry error", text, parent=frame)
        return
    sql = "UPDATE tolerances SET tolerance ='%s' WHERE category = '%s'" % (tolerance, type)
    commit(sql)
    spreadsheet_settings(frame)


def spreadsheet_settings(frame):
    wd = front_window(frame)  # F,S,C,FF,buttons
    Label(wd[3], text="Spreadsheet Settings", font="bold").grid(row=0, column=0, sticky="w")
    Label(wd[3], text="").grid(row=1, column=0)
    sql = "SELECT tolerance FROM tolerances"
    results = inquire(sql)  # get spreadsheet settings from database
    min_nl = StringVar(wd[3])  # create stringvars
    min_wal = StringVar(wd[3])
    min_otdl = StringVar(wd[3])
    min_aux = StringVar(wd[3])
    min_nl.set(results[3][0])  # set the values for the stringvars from dbase
    min_wal.set(results[4][0])
    min_otdl.set(results[5][0])
    min_aux.set(results[6][0])
    # Lay out widgets for displaying/changing minimum spreadsheet rows
    Label(wd[3], text="Minimum rows for No List", width=30, anchor="w").grid(row=2, column=0, ipady=5, sticky="w")
    Entry(wd[3], width=5, textvariable=min_nl).grid(row=2, column=1, padx=4)
    Button(wd[3], width=5, text="change", command=lambda: apply_ss_min(wd[0], min_nl.get(), "min_ss_nl")) \
        .grid(row=2, column=2, padx=4)
    Button(wd[3], width=5, text="info", command=lambda: tolerance_info(wd[0], "min_nl")).grid(row=2, column=3, padx=4)
    Label(wd[3], text="Minimum rows for Work Assignment", width=30, anchor="w").grid(row=3, column=0, ipady=5,
                                                                                     sticky="w")
    Entry(wd[3], width=5, textvariable=min_wal).grid(row=3, column=1, padx=4)
    Button(wd[3], width=5, text="change", command=lambda: apply_ss_min(wd[0], min_wal.get(), "min_ss_wal")) \
        .grid(row=3, column=2, padx=4)
    Button(wd[3], width=5, text="info", command=lambda: tolerance_info(wd[0], "min_wal")).grid(row=3, column=3, padx=4)
    Label(wd[3], text="Minimum rows for OT Desired", width=30, anchor="w").grid(row=4, column=0, ipady=5, sticky="w")
    Entry(wd[3], width=5, textvariable=min_otdl).grid(row=4, column=1, padx=4)
    Button(wd[3], width=5, text="change", command=lambda: apply_ss_min(wd[0], min_otdl.get(), "min_ss_otdl")) \
        .grid(row=4, column=2, padx=4)
    Button(wd[3], width=5, text="info", command=lambda: tolerance_info(wd[0], "min_otdl")).grid(row=4, column=3, padx=4)
    Label(wd[3], text="Minimum rows for Auxiliary", width=30, anchor="w").grid(row=5, column=0, ipady=5, sticky="w")
    Entry(wd[3], width=5, textvariable=min_aux).grid(row=5, column=1, padx=4)
    Button(wd[3], width=5, text="change", command=lambda: apply_ss_min(wd[0], min_aux.get(), "min_ss_aux")) \
        .grid(row=5, column=2, padx=4)
    Button(wd[3], width=5, text="info", command=lambda: tolerance_info(wd[0], "min_wal")).grid(row=5, column=3, padx=4)
    Label(wd[3], text="_______________________________________________________________________", pady=5) \
        .grid(row=6, columnspan=4, sticky="w")
    Label(wd[3], text="Restore Defaults").grid(row=7, column=0, ipady=5, sticky="w")
    Button(wd[3], width=5, text="set", command=lambda: min_ss_presets(wd[0], "default")).grid(row=7, column=3)
    Label(wd[3], text="Set rows to zero").grid(row=8, column=0, ipady=5, sticky="w")
    Button(wd[3], width=5, text="set", command=lambda: min_ss_presets(wd[0], "zero")).grid(row=8, column=3)
    Button(wd[4], text="Go Back", width=20, anchor="w",
           command=lambda: (wd[0].destroy(), main_frame())).pack(side=LEFT)
    rear_window(wd)


def tolerance_info(self, switch):
    if switch == "OT_own_route":
        text = "Sets the tolerance for no list carrier overtime\n" \
               "\n" \
               "Enter a value in clicks between 0 and .99"
    if switch == "OT_off_route":
        text = "Sets the tolerance for no list and work assignment \n" \
               "list carriers for overtime off their own routes.\n\n" \
               "Enter a value in clicks between 0 and .99"
    if switch == "availability":
        text = "Sets the tolerance for availability of otdl and " \
               "aux carriers. Applies to availability to 10, 11.5 \n" \
               "and 12 hour columns.\n\n" \
               "Enter a value in clicks between 0 and .99"
    if switch == "min_nl":
        text = "Sets the minimum number of rows for the No List " \
               "section of the spreadsheet. \n\n" \
               "Enter a value between 0 and 100"
    if switch == "min_wal":
        text = "Sets the minimum number of rows for the Work Assignment " \
               "section of the spreadsheet. \n\n" \
               "Enter a value between 0 and 100"
    if switch == "min_otdl":
        text = "Sets the minimum number of rows for the OT Desired " \
               "section of the spreadsheet. \n\n" \
               "Enter a value between 0 and 100"
    if switch == "min_aux":
        text = "Sets the minimum number of rows for the Auxiliary " \
               "section of the spreadsheet. \n\n" \
               "Enter a value between 0 and 100"
    messagebox.showinfo("About Tolerances", text, parent=self)


def apply_tolerance(self, tolerance, type):
    if isfloat(tolerance) == False:
        text = "You must enter a number."
        messagebox.showerror("Tolerance value entry error", text, parent=self)
        return
    if tolerance.strip() == "":
        text = "You must enter a numeric value for tolerances"
        messagebox.showerror("Tolerance value entry error", text, parent=self)
        return
    if float(tolerance) < 0:
        text = "Values must be equal to or greater than zero."
        messagebox.showerror("Tolerance value entry error", text, parent=self)
        return
    if float(tolerance) > 1:
        text = "You must enter a value less than one."
        messagebox.showerror("Tolerance value entry error", text, parent=self)
        return
    if float(tolerance) < 1:
        number = tolerance.split('.')
        if len(number) == 2:
            if len(number[1]) > 2:
                text = "Value cannot exceed two decimal places."
                messagebox.showerror("Tolerance value entry error", text, parent=self)
        else:
            if len(number[0]) > 2:
                text = "Value cannot exceed two decimal places."
                messagebox.showerror("Tolerance value entry error", text, parent=self)
    sql = "UPDATE tolerances SET tolerance ='%s' WHERE category = '%s'" % (tolerance, type)
    commit(sql)
    tolerances(self)


def tolerance_presets(self, order):
    if order == "default": num = ".25"
    if order == "zero": num = "0"
    types = ("ot_own_rt", "ot_tol", "av_tol")
    for t in types:
        sql = "UPDATE tolerances SET tolerance ='%s' WHERE category = '%s'" % (num, t)
        commit(sql)
    tolerances(self)


def tolerances(self):
    self.destroy()
    F = Frame(root)
    F.pack(fill=BOTH, side=LEFT)
    C1 = Canvas(F)
    C1.pack(fill=BOTH, side=BOTTOM)
    # apply and close buttons
    Button(C1, text="Go Back", width=20, anchor="w",
           command=lambda: [F.destroy(), main_frame()]).pack(side=LEFT)
    # link up the canvas and scrollbar
    S = Scrollbar(F)
    C = Canvas(F, width=1600)
    S.pack(side=RIGHT, fill=BOTH)
    C.pack(side=LEFT, fill=BOTH, pady=10, padx=20)
    S.configure(command=C.yview, orient="vertical")
    C.configure(yscrollcommand=S.set)
    if sys.platform == "win32":
        C.bind_all('<MouseWheel>', lambda event: C.yview_scroll(int(-1 * (event.delta / 120)), "units"))
    elif sys.platform == "linux":
        C.bind_all('<Button-4>', lambda event: C.yview('scroll', -1, 'units'))
        C.bind_all('<Button-5>', lambda event: C.yview('scroll', 1, 'units'))
    # create the frame inside the canvas
    FF = Frame(C)
    C.create_window((0, 0), window=FF, anchor=NW)
    # page contents
    sql = "SELECT * FROM tolerances"
    results = inquire(sql)
    ot_own_rt = StringVar(FF)
    ot_tol = StringVar(FF)
    av_tol = StringVar(FF)
    Label(FF, text="Tolerances", font="bold", anchor="w").grid(row=0, column=0, columnspan=4, sticky="w")
    Label(FF, text=" ").grid(row=1, column=0, columnspan=4, sticky="w")
    Label(FF, text="Overtime on own route", width=20, anchor="w").grid(row=2, column=0, ipady=5, sticky="w")
    Entry(FF, width=5, textvariable=ot_own_rt).grid(row=2, column=1, padx=4)
    Button(FF, width=5, text="change", command=lambda: apply_tolerance(F, ot_own_rt.get(), "ot_own_rt")).grid(row=2,
                                                                                                              column=2,
                                                                                                              padx=4)
    Button(FF, width=5, text="info", command=lambda: tolerance_info(F, "OT_own_route")).grid(row=2, column=3, padx=4)
    Label(FF, text="Overtime off own route").grid(row=3, column=0, ipady=5, sticky="w")
    Entry(FF, width=5, textvariable=ot_tol).grid(row=3, column=1)
    Button(FF, width=5, text="change", command=lambda: apply_tolerance(F, ot_tol.get(), "ot_tol")).grid(row=3, column=2)
    Button(FF, width=5, text="info", command=lambda: tolerance_info(F, "OT_off_route")).grid(row=3, column=3)
    Label(FF, text="Availability tolerance").grid(row=4, column=0, ipady=5, sticky="w")
    Entry(FF, width=5, textvariable=av_tol).grid(row=4, column=1)
    Button(FF, width=5, text="change", command=lambda: apply_tolerance(F, av_tol.get(), "av_tol")).grid(row=4, column=2)
    Button(FF, width=5, text="info", command=lambda: tolerance_info(F, "availability")).grid(row=4, column=3)
    Label(FF, text="____________________________________________________________", pady=5).grid(row=5, columnspan=4,
                                                                                                sticky="w")
    Label(FF, text="Restore Defaults").grid(row=6, column=0, ipady=5, sticky="w")
    Button(FF, width=5, text="set", command=lambda: tolerance_presets(F, "default")).grid(row=6, column=2)
    Label(FF, text="Set tolerances to zero").grid(row=7, column=0, ipady=5, sticky="w")
    Button(FF, width=5, text="set", command=lambda: tolerance_presets(F, "zero")).grid(row=7, column=2)
    ot_own_rt.set(results[0][2])
    ot_tol.set(results[1][2])
    av_tol.set(results[2][2])
    root.update()
    C.config(scrollregion=C.bbox("all"))


def apply_station(switch, station, self):
    global list_of_stations
    if switch == "enter":
        if station.get().strip() == "":
            messagebox.showerror("Prohibited Action",
                                 "You can not enter a blank entry for a station.", parent=self)
            return
        if station.get() in list_of_stations:
            messagebox.showerror("Prohibited Action",
                                 "That station is already in the list of stations.", parent=self)
            return
    if switch == "enter":
        sql = "INSERT INTO stations (station) VALUES('%s')" % (station.get().strip())
        commit(sql)
        list_of_stations.append(station.get())
    if switch == "delete":
        if station == "out of station":
            text = "You can not delete the \"out of station\" listing."
            messagebox.showerror("Action not allowed", text, parent=self)
            return
        sql = "DELETE FROM stations WHERE station='%s'" % (station)
        commit(sql)
        if g_station == station:
            reset("none")
    # access list of stations from database
    sql = "SELECT * FROM stations ORDER BY station"
    results = inquire(sql)
    # define and populate list of stations variable
    del list_of_stations[:]
    for stat in results:
        list_of_stations.append(stat[0])
    station_list(self)


def station_update_apply(self, old_station, new_station):
    global list_of_stations
    if old_station.get() == "select a station":
        messagebox.showerror("Prohibited Action",
                             "Please select a station.", parent=self)
        return
    if new_station.get().strip() == "" or new_station.get() == "enter a new station name":
        messagebox.showerror("Prohibited Action",
                             "You can not enter a blank entry for a station.", parent=self)
        return
    if g_station == old_station.get():
        reset("none")
    go_ahead = True
    duplicate = False
    if new_station.get() in list_of_stations:
        go_ahead = messagebox.askokcancel("Duplicate Detected", "This station already exist in the list of stations. "
                                                                "If you proceed, all records for {} will be merged with "
                                                                "records from {}. Do you want to proceed?".format(
            old_station.get(), new_station.get()))
        duplicate = True
    if duplicate == True and go_ahead == True:
        sql = "DELETE FROM stations WHERE station='%s'" % old_station.get()
        commit(sql)
        list_of_stations.remove(new_station.get())
    if go_ahead == True:
        sql = "UPDATE stations SET station='%s' WHERE station='%s'" % (new_station.get(), old_station.get())
        commit(sql)
        sql = "UPDATE carriers SET station='%s' WHERE station='%s'" % (new_station.get(), old_station.get())
        commit(sql)
        sql = "UPDATE station_index SET kb_station='%s' WHERE kb_station='%s'" % (new_station.get(), old_station.get())
        commit(sql)
        list_of_stations.append(new_station.get())
        list_of_stations.remove(old_station.get())
        station_list(self)
    if go_ahead == False:
        return


def station_list(self):
    self.destroy()
    F = Frame(root)
    F.pack(fill=BOTH, side=LEFT)
    C1 = Canvas(F)
    C1.pack(fill=BOTH, side=BOTTOM)
    Button(C1, text="Go Back", width=20, anchor="w",
           command=lambda: [F.destroy(), main_frame()]).pack(side=LEFT)
    # link up the canvas and scrollbar
    S = Scrollbar(F)
    C = Canvas(F, width=1600)
    S.pack(side=RIGHT, fill=BOTH)
    C.pack(side=LEFT, fill=BOTH, pady=10, padx=20)
    S.configure(command=C.yview, orient="vertical")
    C.configure(yscrollcommand=S.set)
    if sys.platform == "win32":
        C.bind_all('<MouseWheel>', lambda event: C.yview_scroll(int(-1 * (event.delta / 120)), "units"))
    elif sys.platform == "linux":
        C.bind_all('<Button-4>', lambda event: C.yview('scroll', -1, 'units'))
        C.bind_all('<Button-5>', lambda event: C.yview('scroll', 1, 'units'))
    # create the frame inside the canvas
    FF = Frame(C)
    C.create_window((0, 0), window=FF, anchor=NW)
    # page title
    row = 0
    Label(FF, text="Manage Station List", font="bold").grid(row=row, columnspan=2, sticky="w")
    row += 1
    Label(FF, text="____________________________________________________", pady=5). \
        grid(row=row, columnspan=2, sticky="w")
    row += 1
    # enter new stations
    new_name = StringVar(FF)
    Label(FF, text="Enter New Station", pady=5, font="bold").grid(row=row, columnspan=2, sticky="w")
    row += 1
    e = Entry(FF, width=35, textvariable=new_name)
    e.grid(row=row, column=0, sticky="w")
    new_name.set("")
    Button(FF, width=5, anchor="w", text="ENTER", command=lambda: apply_station("enter", new_name, F)). \
        grid(row=row, column=1, sticky="w")
    row += 1
    Label(FF, text="____________________________________________________", pady=5). \
        grid(row=row, columnspan=2, sticky="w")
    row += 1
    # list current list of stations and delete buttons.
    sql = "SELECT * FROM stations ORDER BY station"
    results = inquire(sql)
    Label(FF, text="List Of Stations", font="bold", pady=5).grid(row=row, columnspan=2, sticky="w")
    row += 1
    for record in results:
        Button(FF, text=record[0], width=30, anchor="w").grid(row=row, column=0, sticky="w")
        Button(FF, text="delete", command=lambda x=record[0]: apply_station("delete", x, F)).grid(row=row, column=1,
                                                                                                  sticky="w")
        row += 1
    Label(FF, text="____________________________________________________", pady=5). \
        grid(row=row, columnspan=2, sticky="w")
    row += 1
    # change names of stations
    Label(FF, text="Change Station Name", font="bold").grid(row=row, column=0, sticky="w")
    row += 1
    all_stations = []
    for rec in results:
        all_stations.append(rec[0])
    if "out of station" in all_stations:
        all_stations.remove("out of station")
    old_station = StringVar(FF)
    om = OptionMenu(FF, old_station, *all_stations)
    om.config(width="35")
    om.grid(row=row, column=0, sticky="w", columnspan=2)
    row += 1
    old_station.set("select a station")
    Label(FF, text="enter a new name:").grid(row=row, column=0, sticky="w")
    row += 1
    new_station = StringVar(FF)
    Entry(FF, textvariable=new_station, width="30").grid(row=row, column=0, sticky="w")
    new_station.set("enter a new station name")
    Button(FF, text="update", command=lambda: station_update_apply(F, old_station, new_station)) \
        .grid(row=row, column=1, sticky="w")
    row += 1
    Label(FF, text="____________________________________________________", pady=5). \
        grid(row=row, columnspan=2, sticky="w")
    row += 1
    # find and display list of unique stations
    Label(FF, text="List Of Stations", pady=5, font="bold") \
        .grid(row=row, columnspan=3, sticky="w")
    row += 1
    Label(FF, text="(referenced in carrier database)", pady=5) \
        .grid(row=row, columnspan=3, sticky="w")
    row += 1
    unique_station = []
    sql = "SELECT * FROM carriers"
    results = inquire(sql)
    for name in results:
        if name[5] not in unique_station:
            unique_station.append(name[5])
    unique_station = sorted(unique_station, key=str.lower)
    count = 1
    for s in unique_station:
        Label(FF, text="{}.  {}".format(count, s)).grid(row=row, columnspan=2, sticky="w")
        count += 1
        row += 1
    root.update()
    C.config(scrollregion=C.bbox("all"))


def apply_mi(self, array_var, ls, ns, station, route, date):  # enter changes from multiple input into database
    x = date.get()
    year = IntVar()
    month = IntVar()
    day = IntVar()
    y = g_date[x].strftime("%Y").lstrip("0")
    m = g_date[x].strftime("%m").lstrip("0")
    d = g_date[x].strftime("%d").lstrip("0")
    year.set(y)
    month.set(m)
    day.set(d)
    for i in range(len(array_var)):
        passed_ns = ns[i].get().split("  ")
        ns[i].set(passed_ns[1])
        if array_var[i][2] != ls[i].get() or array_var[i][3] != ns[i].get() or array_var[i][5] != station[i].get():
            apply(year, month, day, array_var[i][1], ls[i], ns[i], route[i], station[i], self)


def mass_input(self, day, sort):
    self.destroy()
    switchF7 = Frame(root)
    switchF7.pack()
    C1 = Canvas(switchF7)
    C1.pack(fill=BOTH, side=BOTTOM)
    # apply and close buttons
    Button(C1, text="Submit", width=10, anchor="w",
           command=lambda: [switchF7.destroy(), apply_mi(switchF7, array_var, mi_list, mi_nsday, mi_station, mi_route,
                                                         pass_date), main_frame()]).pack(side=LEFT)
    Button(C1, text="Apply", width=10, anchor="w",
           command=lambda: [apply_mi(switchF7, array_var, mi_list, mi_nsday, mi_station, mi_route, pass_date),
                            mass_input(switchF7, day, sort)]).pack(side=LEFT)
    Button(C1, text="Go Back", width=10, anchor="w",
           command=lambda: [switchF7.destroy(), main_frame()]).pack(side=LEFT)
    # link up the canvas and scrollbar
    S = Scrollbar(switchF7)
    C = Canvas(switchF7, height=800, width=1600)
    S.pack(side=RIGHT, fill=BOTH)
    C.pack(side=LEFT, fill=BOTH, pady=10, padx=20)
    S.configure(command=C.yview, orient="vertical")
    C.configure(yscrollcommand=S.set)
    if sys.platform == "win32":
        C.bind_all('<MouseWheel>', lambda event: C.yview_scroll(int(-1 * (event.delta / 120)), "units"))
    elif sys.platform == "linux":
        C.bind_all('<Button-4>', lambda event: C.yview('scroll', -1, 'units'))
        C.bind_all('<Button-5>', lambda event: C.yview('scroll', 1, 'units'))
    # create the frame inside the canvas
    head_F = Frame(C)
    C.create_window((0, 0), window=head_F, anchor=NW)
    F = Frame(C)
    C.create_window((0, 50), window=F, anchor=NW)
    # set up the option menus to order results by day and sort criteria.
    mi_date = StringVar()
    mi_sort = StringVar()
    opt_day = ["Sat", "Sun", "Mon", "Tue", "Wed", "Thu", "Fri"]
    opt_sort = ["name", "list", "ns day"]
    mi_date.set(day)
    if g_range == "week":
        mi_date.set(day)
        om1 = OptionMenu(head_F, mi_date, *opt_day)
        om1.config(width="5")
        om1.grid(row=0, column=0)
    mi_sort.set(sort)
    om2 = OptionMenu(head_F, mi_sort, *opt_sort)
    om2.grid(row=0, column=1)
    om2.config(width="8")
    Button(head_F, text="set", width=6, command=lambda: mass_input(switchF7, mi_date.get(), mi_sort.get())).grid(row=0,
                                                                                                                 column=2)
    # figure out the day and display
    pass_date = IntVar()
    if g_range == "week":
        for i in range(len(g_date)):
            if opt_day[i] == day:
                f_date = g_date[i].strftime("%a - %b %d, %Y")
                pass_date.set(i)
                Label(F, text="Showing results for {}".format(f_date), font="bold", justify=LEFT) \
                    .grid(row=0, column=0, columnspan=4, sticky=W)
    if g_range == "day":
        for i in range(len(opt_day)):
            if d_date.strftime("%a") == opt_day[i]:
                f_date = d_date.strftime("%a - %b %d, %Y")
                pass_date.set(i)
                Label(F, text="Showing results for {}".format(f_date), font="bold", justify=LEFT) \
                    .grid(row=0, column=0, columnspan=4, sticky=W)
    # access database
    for i in range(len(g_date)):
        if opt_day[i] == day:
            if g_range == "week":
                sql = "SELECT effective_date, carrier_name,list_status, ns_day,route_s, station, rowid" \
                      " FROM carriers WHERE effective_date <= '%s'" \
                      "ORDER BY carrier_name, effective_date" % (g_date[i])
            else:
                sql = "SELECT effective_date, carrier_name,list_status, ns_day,route_s, station, rowid" \
                      " FROM carriers WHERE effective_date <= '%s'" \
                      "ORDER BY carrier_name, effective_date" % (d_date)
    results = inquire(sql)
    # initialize arrays for data sorting
    carrier_list = []
    candidates = []
    if sort == "list":
        otdl_array = []
        wal_array = []
        nl_array = []
        aux_array = []
    if sort == "ns day":
        yellow_array = []
        blue_array = []
        green_array = []
        brown_array = []
        red_array = []
        black_array = []
        none_array = []
    # take raw data and sort into appropiate arrays
    for i in range(len(results)):
        candidates.append(results[i])  # put name into candidates array
        jump = "no"  # triggers an analysis of the candidates array
        if i != len(results) - 1:  # if the loop has not reached the end of the list
            if results[i][1] == results[i + 1][1]:  # if the name current and next name are the same
                jump = "yes"  # bypasses an analysis of the candidates array
        if jump == "no":
            winner = max(candidates, key=itemgetter(0))  # select the most recent record
            if winner[5] == g_station:  # if that record matches the current station...
                carrier_list.append(winner)  # then insert that record in the carrier list
                if sort == "list":  # sort carrier list by ot list if selected
                    if winner[2] == "otdl": otdl_array.append(winner)
                    if winner[2] == "wal": wal_array.append(winner)
                    if winner[2] == "nl": nl_array.append(winner)
                    if winner[2] == "aux": aux_array.append(winner)
                if sort == "ns day":  # sort carrier list by ns day if selected
                    if winner[3] == "yellow": yellow_array.append(winner)
                    if winner[3] == "blue": blue_array.append(winner)
                    if winner[3] == "green": green_array.append(winner)
                    if winner[3] == "brown": brown_array.append(winner)
                    if winner[3] == "red": red_array.append(winner)
                    if winner[3] == "black": black_array.append(winner)
                    if winner[3] == "none": none_array.append(winner)
        del candidates[:]
    # Display results XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
    i = 1
    array_var = []
    # set up first header
    if sort == "name":
        for c in carrier_list: array_var.append(c)
        list_header = "carrier list"
    if sort == "list":
        array_var = nl_array + wal_array + otdl_array + aux_array
        if len(nl_array) > 0:
            list_header = "nl"
        else:
            list_header = " "
    if sort == "ns day":
        array_var = yellow_array + blue_array + green_array + brown_array + red_array + black_array + none_array
        if len(yellow_array) > 0:
            list_header = "yellow"
        else:
            list_header = " "
    Label(F, text=list_header).grid(row=i, column=0)
    i += 1
    # intialize arrays for option menus
    mi_list = []
    opt_list = "nl", "wal", "otdl", "aux"
    mi_nsday = []
    nsk = []
    days = ("sat", "mon", "tue", "wed", "thu", "fri")
    for each in ns_code.keys(): nsk.append(each)  # make an array of ns_code keys
    opt_nsday = []  # make an array of "day / color" options for option menu
    for each in ns_code:  #
        ns_option = ns_code[each] + "  " + each  # make a string for each day/color
        if each in days:
            ns_option = "fixed:" + "  " + each  # if the ns day is fixed - make a special string
        if each == "none": ns_option = "---" + "  " + each  # if the ns day is "none" - make a special string
        opt_nsday.append(ns_option)
    mi_station = []
    mi_route = []
    c = 0
    for record in array_var:  # loop to put information on to window
        # set up color
        if i & 1:
            color = "light yellow"
        else:
            color = "white"
        if sort == "list":
            if list_header != record[2]:
                list_header = record[2]
                Label(F, text=list_header).grid(row=i, column=0)
                i += 1
        if sort == "ns day":
            if list_header != record[3]:
                list_header = record[3]
                Label(F, text=list_header).grid(row=i, column=0)
                i += 1
        # set up carrier name button and variable
        Button(F, text=record[1], width=24, anchor="w", bg=color, bd=0).grid(row=i, column=0)
        # removed button function: command = lambda x=record[1]: [switchF7.destroy(), edit_carrier(x)])
        # set up list status option menu and variable
        mi_list.append(StringVar(F))
        om_list = OptionMenu(F, mi_list[c], *opt_list)
        om_list.config(width=5, anchor="w", bg=color, relief='ridge', bd=0)
        om_list.grid(row=i, column=1, ipadx=0)
        mi_list[c].set(record[2])
        # set up ns day option menu and variable
        mi_nsday.append(StringVar(F))
        om_nsday = OptionMenu(F, mi_nsday[c], *opt_nsday)
        om_nsday.config(width=10, anchor="w", bg=color, relief='ridge', bd=0)
        om_nsday.grid(row=i, column=2)
        ns_index = nsk.index(record[3])
        mi_nsday[c].set(opt_nsday[ns_index])
        # mi_nsday[c].set(record[3])
        # set up station option menu and variable
        mi_station.append(StringVar(F))
        om_station = OptionMenu(F, mi_station[c], *list_of_stations)
        om_station.config(width=28, anchor="w", bg=color, relief='ridge', bd=0)
        om_station.grid(row=i, column=3)
        mi_station[c].set(record[5])
        # set up route variable - not visible but passed along with other variables
        mi_route.append(StringVar(F))
        mi_route[c].set(record[4])
        c += 1
        i += 1
    del carrier_list[:]
    root.update()
    C.config(scrollregion=C.bbox("all"))


def spreadsheet(list_carrier, r_rings):
    date = g_date[0]
    dates = []  # array containing days.
    if g_range == "week":
        for i in range(7):
            dates.append(date)
            date += timedelta(days=1)
    if g_range == "day": dates.append(d_date)
    if r_rings == "x":
        if g_range == "week":
            sql = "SELECT * FROM rings3 WHERE rings_date BETWEEN '%s' AND '%s' ORDER BY rings_date, carrier_name" \
                  % (g_date[0], g_date[6])
        else:
            sql = "SELECT * FROM rings3 WHERE rings_date = '%s' ORDER BY rings_date, " \
                  "carrier_name" \
                  % (d_date)
        r_rings = inquire(sql)
    # Named styles for workbook
    bd = Side(style='thin', color="80808080")  # defines borders
    ws_header = NamedStyle(name="ws_header", font=Font(bold=True, name='Arial', size=12))
    list_header = NamedStyle(name="list_header", font=Font(bold=True, name='Arial', size=10))
    date_dov = NamedStyle(name="date_dov", font=Font(name='Arial', size=8))
    date_dov_title = NamedStyle(name="date_dov_title", font=Font(bold=True, name='Arial', size=8),
                                alignment=Alignment(horizontal='right'))
    col_header = NamedStyle(name="col_header", font=Font(bold=True, name='Arial', size=8),
                            alignment=Alignment(horizontal='right'))
    input_name = NamedStyle(name="input_name", font=Font(name='Arial', size=8),
                            border=Border(left=bd, top=bd, right=bd, bottom=bd))
    input_s = NamedStyle(name="input_s", font=Font(name='Arial', size=8),
                         border=Border(left=bd, top=bd, right=bd, bottom=bd),
                         alignment=Alignment(horizontal='right'))
    calcs = NamedStyle(name="calcs", font=Font(name='Arial', size=8),
                       border=Border(left=bd, top=bd, right=bd, bottom=bd),
                       fill=PatternFill(fgColor='e5e4e2', fill_type='solid'),
                       alignment=Alignment(horizontal='right'))
    daily_list = []  # array
    candidates = []
    dl_nl = []
    dl_wal = []
    dl_otdl = []
    dl_aux = []
    av_to_10_day = []  # arrays to hold totals for summary sheet.
    av_to_10_row = []
    av_to_12_day = []
    av_to_12_row = []
    man_ot_day = []
    man_ot_row = []
    nl_ot_day = []
    nl_ot_row = []
    day_finder = ["Sat", "Sun", "Mon", "Tue", "Wed", "Thu", "Fri"]
    day_of_week = ["saturday", "sunday", "monday", "tuesday", "wednesday", "thursday", "friday"]
    ws_list = ["saturday", "sunday", "monday", "tuesday", "wednesday", "thursday", "friday"]
    wb = Workbook()  # define the workbook
    if g_range == "day":
        for ii in range(len(day_finder)):
            if d_date.strftime("%a") == day_finder[ii]:  # find the correct day
                i = ii
        ws_list[i] = wb.active  # create first worksheet
        ws_list[i].title = day_of_week[i]  # title first worksheet
        summary = wb.create_sheet("summary")
        reference = wb.create_sheet("reference")
    if g_range == "week":
        ws_list[0] = wb.active  # create first worksheet
        ws_list[0].title = "saturday"  # title first worksheet
        for i in range(1, len(ws_list)):  # create worksheet for remaining six days
            ws_list[i] = wb.create_sheet(ws_list[i])
            i = 0
        ws_list[i].title = day_of_week[i]  # title first worksheet
        summary = wb.create_sheet("summary")
        reference = wb.create_sheet("reference")
    # get spreadsheet row minimums from tolerance table
    sql = "SELECT tolerance FROM tolerances"
    result = inquire(sql)
    min_ss_nl = int(result[3][0])
    min_ss_wal = int(result[4][0])
    min_ss_otdl = int(result[5][0])
    min_ss_aux = int(result[6][0])
    for day in dates:
        del daily_list[:]
        del dl_nl[:]
        del dl_wal[:]
        del dl_otdl[:]
        del dl_aux[:]
        # create a list of carriers for each day.
        for ii in range(len(list_carrier)):
            if list_carrier[ii][0] <= str(day):
                candidates.append(list_carrier[ii])  # put name into candidates array
            jump = "no"  # triggers an analysis of the candidates array
            if ii != len(list_carrier) - 1:  # if the loop has not reached the end of the list
                if list_carrier[ii][1] == list_carrier[ii + 1][1]:  # if the name current and next name are the same
                    jump = "yes"  # bypasses an analysis of the candidates array
            if jump == "no":  # review the list of candidates
                winner = max(candidates, key=itemgetter(0))  # select the most recent
                if winner[5] == g_station: daily_list.append(winner)  # add the record if it matches the station
                del candidates[:]  # empty out the candidates array.
        for item in daily_list:  # sort carriers in daily list by the list they are in
            if item[2] == "nl":
                dl_nl.append(item)
            if item[2] == "wal":
                dl_wal.append(item)
            if item[2] == "otdl":
                dl_otdl.append(item)
            if item[2] == "aux":
                dl_aux.append(item)
        ws_list[i].oddFooter.center.text = "&A"
        ws_list[i].column_dimensions["A"].width = 14
        ws_list[i].column_dimensions["B"].width = 5
        ws_list[i].column_dimensions["C"].width = 6
        ws_list[i].column_dimensions["D"].width = 6
        ws_list[i].column_dimensions["E"].width = 6
        ws_list[i].column_dimensions["F"].width = 6
        ws_list[i].column_dimensions["G"].width = 6
        ws_list[i].column_dimensions["H"].width = 6
        ws_list[i].column_dimensions["I"].width = 6
        ws_list[i].column_dimensions["J"].width = 6
        ws_list[i].column_dimensions["K"].width = 6
        ws_list[i]['A1'] = "Improper Mandate Worksheet"
        ws_list[i]['A1'].style = ws_header
        ws_list[i].merge_cells('A1:E1')
        ws_list[i]['A3'] = "Date:  "  # create date/ pay period/ station header
        ws_list[i]['A3'].style = date_dov_title
        ws_list[i]['B3'] = format(day, "%A  %m/%d/%y")
        ws_list[i]['B3'].style = date_dov
        ws_list[i].merge_cells('B3:D3')
        ws_list[i]['E3'] = "Pay Period:  "
        ws_list[i]['E3'].style = date_dov_title
        ws_list[i].merge_cells('E3:F3')
        ws_list[i]['G3'] = pay_period
        ws_list[i]['G3'].style = date_dov
        ws_list[i].merge_cells('G3:H3')
        ws_list[i]['A4'] = "Station:  "
        ws_list[i]['A4'].style = date_dov_title
        ws_list[i]['B4'] = g_station
        ws_list[i]['B4'].style = date_dov
        ws_list[i].merge_cells('B4:D4')
        # no list carriers *********************************************************************************************
        ws_list[i]['A6'] = "No List Carriers"
        ws_list[i]['A6'].style = list_header
        # column headers
        ws_list[i]['A7'] = "Name"
        ws_list[i]['A7'].style = col_header
        ws_list[i]['B7'] = "note"
        ws_list[i]['B7'].style = col_header
        ws_list[i]['C7'] = "5200"
        ws_list[i]['C7'].style = col_header
        ws_list[i]['D7'] = "RS"
        ws_list[i]['D7'].style = col_header
        ws_list[i]['E7'] = "MV off"
        ws_list[i]['E7'].style = col_header
        ws_list[i]['F7'] = "MV on"
        ws_list[i]['F7'].style = col_header
        ws_list[i]['G7'] = "Route"
        ws_list[i]['G7'].style = col_header
        ws_list[i]['H7'] = "MV total"
        ws_list[i]['H7'].style = col_header
        ws_list[i]['I7'] = "OT"
        ws_list[i]['I7'].style = col_header
        ws_list[i]['J7'] = "off rt"
        ws_list[i]['J7'].style = col_header
        ws_list[i]['K7'] = "OT off rt"
        ws_list[i]['K7'].style = col_header
        oi = 8  # rows: start at 8th row
        move_totals = []  # list of totals of each set of moves
        ot_total = 0  # running total for OT
        ot_off_total = 0  # running total for OT off route
        daily_ot_off_rt = 0
        daily_avail = 0
        nl_oi_start = oi  # start counting the number of rows in nl
        for line in dl_nl:
            match = "miss"
            del move_totals[:]  # empty array of moves totals.
            # if there is a ring to match the carrier/ date then printe
            for each in r_rings:
                if each[0] == str(day) and each[1] == line[1]:  # if the rings record is a match
                    match = "hit"
                    if match == "hit":
                        s_moves = each[5].split(",")  # sort out the moves
                        c = 0
                        for e in range(int(len(s_moves) / 3)):  # tally totals for each set of moves
                            total = float(s_moves[c + 1]) - float(s_moves[c])  # calc off time off route
                            c = c + 3
                            move_totals.append(total)
                        off_route = 0.0
                        if str(each[2]) != "":  # in case the 5200 time is blank
                            time5200 = each[2]
                        else:
                            time5200 = 0
                        if each[4] == "ns day":  # if the carrier worked on their ns day
                            off_route = float(time5200)  # cal >off route
                            ot = float(time5200)  # cal > ot
                        else:  # if carrier did not work ns day
                            ot = max(float(time5200) - float(8), 0)  # calculate overtime
                            for mt in move_totals:  # calc total off route work.
                                off_route += float(mt)
                        ot_total += ot
                        ot_off_route = min(off_route, ot)  # calculate the ot off route
                        ot_off_total += ot_off_route
                        move_count = (int(len(s_moves) / 3))  # find the number of sets of moves
                        # output to the gui
                        ws_list[i]['A' + str(oi)] = each[1]  # name
                        ws_list[i]['A' + str(oi)].style = input_name
                        if each[4] == "none":
                            code = ""  # leave code field blank if 'none'
                        else:
                            code = each[4]
                        ws_list[i]['B' + str(oi)] = code  # code
                        ws_list[i]['B' + str(oi)].style = input_s
                        if time5200 == 0:
                            ws_list[i]['C' + str(oi)] = ""  # 5200
                        else:
                            ws_list[i]['C' + str(oi)] = float(time5200)  # 5200
                        ws_list[i]['C' + str(oi)].style = input_s
                        ws_list[i]['C' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
                        if isfloat(each[3]) == True:
                            ws_list[i]['D' + str(oi)] = float(each[3])
                        else:
                            ws_list[i]['D' + str(oi)] = each[3]
                        ws_list[i]['D' + str(oi)].style = input_s
                        ws_list[i]['D' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
                        count = 0
                        if move_count == 0:  # if there are no moves then format the empty cells
                            ws_list[i]['E' + str(oi)] = ""  # move off
                            ws_list[i]['E' + str(oi)].style = input_s
                            ws_list[i]['E' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
                            count += 1
                            ws_list[i]['F' + str(oi)] = ""  # move on
                            ws_list[i]['F' + str(oi)].style = input_s
                            ws_list[i]['F' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
                            count += 1
                            ws_list[i]['G' + str(oi)] = ""  # route
                            ws_list[i]['G' + str(oi)].style = input_s
                            ws_list[i]['G' + str(oi)].number_format = "####"
                            count += 1
                            formula = "=SUM(%s!F%s - %s!E%s)" % (day_of_week[i], str(oi), day_of_week[i], str(oi))
                            ws_list[i]['H' + str(oi)] = formula  # move total
                            ws_list[i]['H' + str(oi)].style = input_s
                            ws_list[i]['H' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
                        elif move_count == 1:  # if there is only one set of moves
                            ws_list[i]['E' + str(oi)] = float(s_moves[0])  # move off
                            ws_list[i]['E' + str(oi)].style = input_s
                            ws_list[i]['E' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
                            count += 1
                            ws_list[i]['F' + str(oi)] = float(s_moves[1])  # move on
                            ws_list[i]['F' + str(oi)].style = input_s
                            ws_list[i]['F' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
                            count += 1
                            ws_list[i]['G' + str(oi)] = float(s_moves[2])  # route
                            ws_list[i]['G' + str(oi)].style = input_s
                            ws_list[i]['G' + str(oi)].number_format = "####"
                            count += 1
                            formula = "=SUM(%s!F%s - %s!E%s)" % (day_of_week[i], str(oi), day_of_week[i], str(oi))
                            ws_list[i]['H' + str(oi)] = formula  # move total
                            ws_list[i]['H' + str(oi)].style = input_s
                            ws_list[i]['H' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
                        else:  # There are multiple moves
                            ws_list[i]['E' + str(oi)] = "*"  # move off
                            ws_list[i]['E' + str(oi)].style = input_s
                            ws_list[i]['F' + str(oi)] = "*"  # move on
                            ws_list[i]['F' + str(oi)].style = input_s
                            ws_list[i]['G' + str(oi)] = "*"  # route
                            ws_list[i]['G' + str(oi)].style = input_s
                            formula = "=SUM(%s!H%s:%s!H%s)" % (day_of_week[i], str(oi + move_count),
                                                               day_of_week[i], str(oi + 1))
                            ws_list[i]['H' + str(oi)] = formula  # move total
                            ws_list[i]['H' + str(oi)].style = input_s
                            ws_list[i]['H' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"

                            formula = "=IF(%s!B%s =\"ns day\", %s!C%s,IF(%s!C%s <= 8 + reference!C3, 0, MAX(%s!C%s - 8, 0)))" \
                                      % (day_of_week[i], str(oi), day_of_week[i], str(oi), day_of_week[i], str(oi),
                                         day_of_week[i], str(oi))

                            ws_list[i]['I' + str(oi)] = formula  # overtime
                            ws_list[i]['I' + str(oi)].style = calcs
                            ws_list[i]['I' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
                            formula = "=%s!H%s" % (day_of_week[i], str(oi))  # copy data from column H/ MV total
                            ws_list[i]['J' + str(oi)] = formula  # off route
                            ws_list[i]['J' + str(oi)].style = calcs
                            ws_list[i]['J' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
                            # formula for OT off route
                            formula = "=IF(OR(%s!B%s=\"ns day\",%s!J%s >= %s!C%s),%s!C%s, IF(%s!C%s <= 8 + reference!C4, 0, MIN" \
                                      "(MAX(%s!C%s - 8, 0),IF(%s!J%s <= reference!C4,0, %s!J%s))))" \
                                      % (day_of_week[i], str(oi), day_of_week[i], str(oi),day_of_week[i], str(oi),
                                         day_of_week[i], str(oi), day_of_week[i], str(oi),
                                         day_of_week[i], str(oi), day_of_week[i], str(oi), day_of_week[i], str(oi))
                            ws_list[i]['K' + str(oi)] = formula  # OT off route
                            ws_list[i]['K' + str(oi)].style = calcs
                            ws_list[i]['K' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
                            oi += 1
                            for ii in range(move_count):  # if there are multiple moves, create + populate cells
                                ws_list[i]['E' + str(oi)] = float(s_moves[count])  # move off
                                ws_list[i]['E' + str(oi)].style = input_s
                                ws_list[i]['E' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
                                count += 1
                                ws_list[i]['F' + str(oi)] = float(s_moves[count])  # move on
                                ws_list[i]['F' + str(oi)].style = input_s
                                ws_list[i]['F' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
                                count += 1
                                ws_list[i]['G' + str(oi)] = float(s_moves[count])  # route
                                ws_list[i]['G' + str(oi)].style = input_s
                                ws_list[i]['G' + str(oi)].number_format = "####"
                                count += 1
                                formula = "=SUM(%s!F%s - %s!E%s)" % (day_of_week[i], str(oi), day_of_week[i], str(oi))
                                ws_list[i]['H' + str(oi)] = formula  # move total
                                ws_list[i]['H' + str(oi)].style = input_s
                                ws_list[i]['H' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
                                if ii < move_count - 1: oi += 1  # create another row
                            oi += 1
                        if move_count < 2:
                            # input formula for overtime
                            formula = "=IF(%s!B%s =\"ns day\", %s!C%s,IF(%s!C%s <= 8+ reference!C3, 0, MAX(%s!C%s - 8, 0)))" \
                                      % (day_of_week[i], str(oi), day_of_week[i], str(oi), day_of_week[i], str(oi),
                                         day_of_week[i], str(oi))
                            ws_list[i]['I' + str(oi)] = formula  # overtime
                            ws_list[i]['I' + str(oi)].style = calcs
                            ws_list[i]['I' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
                            # formula for off route
                            formula = "=SUM(%s!F%s - %s!E%s)" \
                                      % (day_of_week[i], str(oi), day_of_week[i], str(oi))
                            ws_list[i]['J' + str(oi)] = formula  # off route
                            ws_list[i]['J' + str(oi)].style = calcs
                            ws_list[i]['J' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
                            # formula for OT off route
                            formula = "=IF(OR(%s!B%s=\"ns day\",%s!J%s >= %s!C%s),%s!C%s, IF(%s!C%s <= 8 + reference!C4, 0, MIN" \
                                      "(MAX(%s!C%s - 8, 0),IF(%s!J%s <= reference!C4,0, %s!J%s))))" \
                                      % (day_of_week[i], str(oi), day_of_week[i], str(oi), day_of_week[i], str(oi),
                                         day_of_week[i], str(oi), day_of_week[i], str(oi),
                                         day_of_week[i], str(oi), day_of_week[i], str(oi), day_of_week[i], str(oi))
                            ws_list[i]['K' + str(oi)] = formula  # OT off route
                            ws_list[i]['K' + str(oi)].style = calcs
                            ws_list[i]['K' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
                            oi += 1
            #  if there is no match, then just printe the name.
            if match == "miss":
                ws_list[i]['A' + str(oi)] = line[1]  # name
                ws_list[i]['A' + str(oi)].style = input_name
                ws_list[i]['B' + str(oi)].style = input_s
                ws_list[i]['B' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
                ws_list[i]['C' + str(oi)].style = input_s
                ws_list[i]['C' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
                ws_list[i]['D' + str(oi)].style = input_s
                ws_list[i]['D' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
                ws_list[i]['E' + str(oi)].style = input_s
                ws_list[i]['E' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
                ws_list[i]['F' + str(oi)].style = input_s
                ws_list[i]['F' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
                ws_list[i]['G' + str(oi)].style = input_s
                ws_list[i]['G' + str(oi)].number_format = "####"
                formula = "=SUM(%s!F%s - %s!E%s)" % (day_of_week[i], str(oi), day_of_week[i], str(oi))
                ws_list[i]['H' + str(oi)] = formula  # move total
                ws_list[i]['H' + str(oi)].style = input_s
                ws_list[i]['H' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
                formula = "=IF(%s!B%s =\"ns day\", %s!C%s,IF(%s!C%s <= 8 + reference!C3, 0, MAX(%s!C%s - 8, 0)))" \
                          % (day_of_week[i], str(oi), day_of_week[i], str(oi), day_of_week[i], str(oi),
                             day_of_week[i], str(oi))
                ws_list[i]['I' + str(oi)] = formula  # overtime
                ws_list[i]['I' + str(oi)].style = calcs
                ws_list[i]['I' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
                formula = "=SUM(%s!F%s - %s!E%s)" \
                          % (day_of_week[i], str(oi), day_of_week[i], str(oi))
                ws_list[i]['J' + str(oi)] = formula  # off route
                ws_list[i]['J' + str(oi)].style = calcs
                ws_list[i]['J' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
                formula = "=IF(OR(%s!B%s=\"ns day\",%s!J%s >= %s!C%s),%s!C%s, IF(%s!C%s <= 8 + reference!C4, 0, MIN" \
                          "(MAX(%s!C%s - 8, 0),IF(%s!J%s <= reference!C4,0, %s!J%s))))" \
                          % (day_of_week[i], str(oi), day_of_week[i], str(oi), day_of_week[i], str(oi),
                             day_of_week[i], str(oi), day_of_week[i], str(oi),
                             day_of_week[i], str(oi), day_of_week[i], str(oi), day_of_week[i], str(oi))
                ws_list[i]['K' + str(oi)] = formula  # OT off route
                ws_list[i]['K' + str(oi)].style = calcs
                ws_list[i]['K' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
                oi += 1
        nl_oi_end = oi
        nl_oi_diff = nl_oi_end - nl_oi_start  # find how many lines exist in nl
        # if the minimum number of rows are not reached, insert blank rows
        e_range = min_ss_nl - nl_oi_diff
        if e_range <= 0:
            e_range = 0
        for e in range(e_range):
            ws_list[i]['A' + str(oi)] = ""  # name
            ws_list[i]['A' + str(oi)].style = input_name
            ws_list[i]['B' + str(oi)].style = input_s
            ws_list[i]['B' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
            ws_list[i]['C' + str(oi)].style = input_s
            ws_list[i]['C' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
            ws_list[i]['D' + str(oi)].style = input_s
            ws_list[i]['D' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
            ws_list[i]['E' + str(oi)].style = input_s
            ws_list[i]['E' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
            ws_list[i]['F' + str(oi)].style = input_s
            ws_list[i]['F' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
            ws_list[i]['G' + str(oi)].style = input_s
            ws_list[i]['G' + str(oi)].number_format = "####"
            formula = "=SUM(%s!F%s - %s!E%s)" % (day_of_week[i], str(oi), day_of_week[i], str(oi))
            ws_list[i]['H' + str(oi)] = formula  # move total
            ws_list[i]['H' + str(oi)].style = input_s
            ws_list[i]['H' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
            formula = "=IF(%s!B%s =\"ns day\", %s!C%s,IF(%s!C%s <= 8 + reference!C3, 0, MAX(%s!C%s - 8, 0)))" \
                      % (day_of_week[i], str(oi), day_of_week[i], str(oi), day_of_week[i], str(oi),
                         day_of_week[i], str(oi))
            ws_list[i]['I' + str(oi)] = formula  # overtime
            ws_list[i]['I' + str(oi)].style = calcs
            ws_list[i]['I' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
            formula = "=SUM(%s!F%s - %s!E%s)" \
                      % (day_of_week[i], str(oi), day_of_week[i], str(oi))
            ws_list[i]['J' + str(oi)] = formula  # off route
            ws_list[i]['J' + str(oi)].style = calcs
            ws_list[i]['J' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
            formula = "=IF(OR(%s!B%s=\"ns day\",%s!J%s >= %s!C%s),%s!C%s, IF(%s!C%s <= 8 + reference!C4, 0, MIN" \
                      "(MAX(%s!C%s - 8, 0),IF(%s!J%s <= reference!C4,0, %s!J%s))))" \
                      % (day_of_week[i], str(oi), day_of_week[i], str(oi), day_of_week[i], str(oi),
                         day_of_week[i], str(oi), day_of_week[i], str(oi),
                         day_of_week[i], str(oi), day_of_week[i], str(oi), day_of_week[i], str(oi))
            ws_list[i]['K' + str(oi)] = formula  # OT off route
            ws_list[i]['K' + str(oi)].style = calcs
            ws_list[i]['K' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
            oi += 1
        cell = str(oi - 1)
        oi += 1
        ws_list[i]['H' + str(oi)] = "Total NL Overtime"
        ws_list[i]['H' + str(oi)].style = col_header
        formula = "=SUM(%s!I8:%s!I%s)" % (day_of_week[i], day_of_week[i], cell)
        ws_list[i]['I' + str(oi)] = formula  # OT
        nl_ot_row.append(str(oi))  # get the cell information to reference in summary tab
        nl_ot_day.append(i)
        ws_list[i]['I' + str(oi)].style = calcs
        ws_list[i]['I' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
        oi += 2
        ws_list[i]['J' + str(oi)] = "Total NL Mandates"
        ws_list[i]['J' + str(oi)].style = col_header
        formula = "=SUM(%s!K8:%s!K%s)" % (day_of_week[i], day_of_week[i], cell)
        ws_list[i]['K' + str(oi)] = formula  # OT off route
        ws_list[i]['K' + str(oi)].style = calcs
        ws_list[i]['K' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
        nl_totals = oi
        oi += 1
        try:
            ws_list[i].page_breaks.append(Break(id=oi))
        except:
            ws_list[i].row_breaks.append(Break(id=oi))
        oi += 1
        # # work assignment carriers ************************************************************************************
        ws_list[i]['A' + str(oi)] = "Work Assignment Carriers"
        ws_list[i]['A' + str(oi)].style = list_header
        oi += 1
        # column headers
        ws_list[i]['A' + str(oi)] = "Name"
        ws_list[i]['A' + str(oi)].style = col_header
        ws_list[i]['B' + str(oi)] = "note"
        ws_list[i]['B' + str(oi)].style = col_header
        ws_list[i]['C' + str(oi)] = "5200"
        ws_list[i]['C' + str(oi)].style = col_header
        ws_list[i]['D' + str(oi)] = "RS"
        ws_list[i]['D' + str(oi)].style = col_header
        ws_list[i]['E' + str(oi)] = "MV off"
        ws_list[i]['E' + str(oi)].style = col_header
        ws_list[i]['F' + str(oi)] = "MV on"
        ws_list[i]['F' + str(oi)].style = col_header
        ws_list[i]['G' + str(oi)] = "Route"
        ws_list[i]['G' + str(oi)].style = col_header
        ws_list[i]['H' + str(oi)] = "MV total"
        ws_list[i]['H' + str(oi)].style = col_header
        ws_list[i]['I' + str(oi)] = "OT"
        ws_list[i]['I' + str(oi)].style = col_header
        ws_list[i]['J' + str(oi)] = "off rt"
        ws_list[i]['J' + str(oi)].style = col_header
        ws_list[i]['K' + str(oi)] = "OT off rt"
        ws_list[i]['K' + str(oi)].style = col_header
        oi += 1
        wal_oi_start = oi
        top_cell = str(oi)
        move_totals = []  # list of totals of each set of moves
        ot_total = 0  # running total for OT
        ot_off_total = 0  # running total for OT off route
        for line in dl_wal:
            match = "miss"
            del move_totals[:]  # empty array of moves totals.
            # if there is a ring to match the carrier/ date then printe
            for each in r_rings:
                if each[0] == str(day) and each[1] == line[1]:  # if the rings record is a match
                    match = "hit"
                    if match == "hit":
                        s_moves = each[5].split(",")  # sort out the moves
                        c = 0
                        for e in range(int(len(s_moves) / 3)):  # tally totals for each set of moves
                            total = float(s_moves[c + 1]) - float(s_moves[c])
                            c = c + 3
                            move_totals.append(total)
                        off_route = 0.0
                        if str(each[2]) != "":  # in case the 5200 time is blank
                            time5200 = each[2]
                        else:
                            time5200 = 0
                        if each[4] == "ns day":  # if the carrier worked on their ns day
                            off_route = float(time5200)  # cal >off route
                            ot = float(time5200)  # cal > ot
                        else:  # if carrier did not work ns day
                            ot = max(float(time5200) - float(8), 0)  # calculate overtime
                            for mt in move_totals:  # calc total off route work.
                                off_route += float(mt)
                        ot_total += ot
                        ot_off_route = min(off_route, ot)  # calculate the ot off route
                        ot_off_total += ot_off_route
                        move_count = (int(len(s_moves) / 3))  # find the number of sets of moves
                        # output to the gui
                        ws_list[i]['A' + str(oi)] = each[1]  # name
                        ws_list[i]['A' + str(oi)].style = input_name
                        if each[4] == "none":
                            code = ""  # leave code field blank if 'none'
                        else:
                            code = each[4]
                        ws_list[i]['B' + str(oi)] = code  # code
                        ws_list[i]['B' + str(oi)].style = input_s
                        if time5200 == 0:
                            ws_list[i]['C' + str(oi)] = ""  # 5200
                        else:
                            ws_list[i]['C' + str(oi)] = float(time5200)  # 5200
                        ws_list[i]['C' + str(oi)].style = input_s
                        ws_list[i]['C' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
                        if isfloat(each[3]) == True:
                            ws_list[i]['D' + str(oi)] = float(each[3])
                        else:
                            ws_list[i]['D' + str(oi)] = each[3]
                        ws_list[i]['D' + str(oi)].style = input_s
                        ws_list[i]['D' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
                        count = 0
                        if move_count == 0:  # if there are no moves then format the empty cells
                            ws_list[i]['E' + str(oi)] = ""  # move off
                            ws_list[i]['E' + str(oi)].style = input_s
                            ws_list[i]['E' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
                            count += 1
                            ws_list[i]['F' + str(oi)] = ""  # move on
                            ws_list[i]['F' + str(oi)].style = input_s
                            ws_list[i]['F' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
                            count += 1
                            ws_list[i]['G' + str(oi)] = ""  # route
                            ws_list[i]['G' + str(oi)].style = input_s
                            ws_list[i]['G' + str(oi)].number_format = "####"
                            count += 1
                            formula = "=SUM(%s!F%s - %s!E%s)" % (day_of_week[i], str(oi), day_of_week[i], str(oi))
                            ws_list[i]['H' + str(oi)] = formula  # move total
                            ws_list[i]['H' + str(oi)].style = input_s
                            ws_list[i]['H' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
                        elif move_count == 1:  # if there is only one set of moves
                            ws_list[i]['E' + str(oi)] = float(s_moves[0])  # move off
                            ws_list[i]['E' + str(oi)].style = input_s
                            ws_list[i]['E' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
                            count += 1
                            ws_list[i]['F' + str(oi)] = float(s_moves[1])  # move on
                            ws_list[i]['F' + str(oi)].style = input_s
                            ws_list[i]['F' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
                            count += 1
                            ws_list[i]['G' + str(oi)] = float(s_moves[2])  # route
                            ws_list[i]['G' + str(oi)].style = input_s
                            ws_list[i]['G' + str(oi)].number_format = "####"
                            count += 1
                            formula = "=SUM(%s!F%s - %s!E%s)" % (day_of_week[i], str(oi), day_of_week[i], str(oi))
                            ws_list[i]['H' + str(oi)] = formula  # move total
                            ws_list[i]['H' + str(oi)].style = input_s
                            ws_list[i]['H' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
                        else:  # There are multiple moves
                            ws_list[i]['E' + str(oi)] = "*"  # move off
                            ws_list[i]['E' + str(oi)].style = input_s
                            ws_list[i]['F' + str(oi)] = "*"  # move on
                            ws_list[i]['F' + str(oi)].style = input_s
                            ws_list[i]['G' + str(oi)] = "*"  # route
                            ws_list[i]['G' + str(oi)].style = input_s
                            formula = "=SUM(%s!H%s:%s!H%s)" % (day_of_week[i], str(oi + move_count),
                                                               day_of_week[i], str(oi + 1))
                            ws_list[i]['H' + str(oi)] = formula  # move total
                            ws_list[i]['H' + str(oi)].style = input_s
                            ws_list[i]['H' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
                            # input formula for overtime
                            formula = "=IF(%s!B%s =\"ns day\", %s!C%s, MAX(%s!C%s - 8, 0))" \
                                      % (day_of_week[i], str(oi), day_of_week[i], str(oi), day_of_week[i], str(oi))
                            ws_list[i]['I' + str(oi)] = formula  # overtime
                            ws_list[i]['I' + str(oi)].style = calcs
                            ws_list[i]['I' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
                            formula = "=%s!H%s" % (day_of_week[i], str(oi))  # copy data from column H/ MV total
                            ws_list[i]['J' + str(oi)] = formula  # off route
                            ws_list[i]['J' + str(oi)].style = calcs
                            ws_list[i]['J' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
                            # formula for OT off route
                            formula = "=IF(OR(%s!B%s=\"ns day\",%s!J%s >= %s!C%s),%s!C%s, IF(%s!C%s <= 8 + reference!C4, 0, MIN" \
                                      "(MAX(%s!C%s - 8, 0),IF(%s!J%s <= reference!C4,0, %s!J%s))))" \
                                      % (day_of_week[i], str(oi), day_of_week[i], str(oi), day_of_week[i], str(oi),
                                         day_of_week[i], str(oi), day_of_week[i], str(oi),
                                         day_of_week[i], str(oi), day_of_week[i], str(oi), day_of_week[i], str(oi))
                            ws_list[i]['K' + str(oi)] = formula  # OT off route
                            ws_list[i]['K' + str(oi)].style = calcs
                            ws_list[i]['K' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
                            oi += 1
                            for ii in range(move_count):  # if there are multiple moves, create + populate cells
                                ws_list[i]['E' + str(oi)] = float(s_moves[count])  # move off
                                ws_list[i]['E' + str(oi)].style = input_s
                                ws_list[i]['E' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
                                count += 1
                                ws_list[i]['F' + str(oi)] = float(s_moves[count])  # move on
                                ws_list[i]['F' + str(oi)].style = input_s
                                ws_list[i]['F' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
                                count += 1
                                ws_list[i]['G' + str(oi)] = float(s_moves[count])  # route
                                ws_list[i]['G' + str(oi)].style = input_s
                                ws_list[i]['G' + str(oi)].number_format = "####"
                                count += 1
                                formula = "=SUM(%s!F%s - %s!E%s)" % (day_of_week[i], str(oi), day_of_week[i], str(oi))
                                ws_list[i]['H' + str(oi)] = formula  # move total
                                ws_list[i]['H' + str(oi)].style = input_s
                                ws_list[i]['H' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
                                if ii < move_count - 1: oi += 1
                            oi += 1
                        if move_count < 2:
                            # input formula for overtime
                            formula = "=IF(%s!B%s =\"ns day\", %s!C%s, MAX(%s!C%s - 8, 0))" \
                                      % (day_of_week[i], str(oi), day_of_week[i], str(oi), day_of_week[i], str(oi))
                            ws_list[i]['I' + str(oi)] = formula  # overtime
                            ws_list[i]['I' + str(oi)].style = calcs
                            ws_list[i]['I' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
                            # formula for off route
                            formula = "=SUM(%s!F%s - %s!E%s)" \
                                      % (day_of_week[i], str(oi), day_of_week[i], str(oi))
                            ws_list[i]['J' + str(oi)] = formula  # off route
                            ws_list[i]['J' + str(oi)].style = calcs
                            ws_list[i]['J' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
                            # formula for OT off route
                            formula = "=IF(OR(%s!B%s=\"ns day\",%s!J%s >= %s!C%s),%s!C%s, IF(%s!C%s <= 8 + reference!C4, 0, MIN" \
                                      "(MAX(%s!C%s - 8, 0),IF(%s!J%s <= reference!C4,0, %s!J%s))))" \
                                      % (day_of_week[i], str(oi), day_of_week[i], str(oi), day_of_week[i], str(oi),
                                         day_of_week[i], str(oi), day_of_week[i], str(oi),
                                         day_of_week[i], str(oi), day_of_week[i], str(oi), day_of_week[i], str(oi))
                            ws_list[i]['K' + str(oi)] = formula  # OT off route
                            ws_list[i]['K' + str(oi)].style = calcs
                            ws_list[i]['K' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
                            oi += 1
            #  if there is no match, then just printe the name.
            if match == "miss":
                ws_list[i]['A' + str(oi)] = line[1]  # name
                ws_list[i]['A' + str(oi)].style = input_name
                ws_list[i]['B' + str(oi)].style = input_s
                ws_list[i]['B' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
                ws_list[i]['C' + str(oi)].style = input_s
                ws_list[i]['C' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
                ws_list[i]['D' + str(oi)].style = input_s
                ws_list[i]['D' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
                ws_list[i]['E' + str(oi)].style = input_s
                ws_list[i]['E' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
                ws_list[i]['F' + str(oi)].style = input_s
                ws_list[i]['F' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
                ws_list[i]['G' + str(oi)].style = input_s
                ws_list[i]['G' + str(oi)].number_format = "####"
                formula = "=SUM(%s!F%s - %s!E%s)" % (day_of_week[i], str(oi), day_of_week[i], str(oi))
                ws_list[i]['H' + str(oi)] = formula  # move total
                ws_list[i]['H' + str(oi)].style = input_s
                ws_list[i]['H' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
                formula = "=IF(%s!B%s =\"ns day\", %s!C%s, MAX(%s!C%s - 8, 0))" \
                          % (day_of_week[i], str(oi), day_of_week[i], str(oi), day_of_week[i], str(oi))
                ws_list[i]['I' + str(oi)] = formula  # overtime
                ws_list[i]['I' + str(oi)].style = calcs
                ws_list[i]['I' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
                formula = "=SUM(%s!F%s - %s!E%s)" \
                          % (day_of_week[i], str(oi), day_of_week[i], str(oi))
                ws_list[i]['J' + str(oi)] = formula  # off route
                ws_list[i]['J' + str(oi)].style = calcs
                ws_list[i]['J' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
                formula = "=IF(OR(%s!B%s=\"ns day\",%s!J%s >= %s!C%s),%s!C%s, IF(%s!C%s <= 8 + reference!C4, 0, MIN" \
                          "(MAX(%s!C%s - 8, 0),IF(%s!J%s <= reference!C4,0, %s!J%s))))" \
                          % (day_of_week[i], str(oi), day_of_week[i], str(oi), day_of_week[i], str(oi),
                             day_of_week[i], str(oi), day_of_week[i], str(oi),
                             day_of_week[i], str(oi), day_of_week[i], str(oi), day_of_week[i], str(oi))
                ws_list[i]['K' + str(oi)] = formula  # OT off route
                ws_list[i]['K' + str(oi)].style = calcs
                ws_list[i]['K' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
                oi += 1
        wal_oi_end = oi
        wal_oi_diff = wal_oi_end - wal_oi_start  # find how many lines exist in nl
        # if the minimum number of rows are not reached, insert blank rows
        e_range = min_ss_wal - wal_oi_diff
        if e_range <= 0:
            e_range = 0
        for e in range(e_range):
            # ws_list[i]['A' + str(oi)] = ""  # name
            ws_list[i]['A' + str(oi)].style = input_name
            ws_list[i]['B' + str(oi)].style = input_s
            ws_list[i]['B' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
            ws_list[i]['C' + str(oi)].style = input_s
            ws_list[i]['C' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
            ws_list[i]['D' + str(oi)].style = input_s
            ws_list[i]['D' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
            ws_list[i]['E' + str(oi)].style = input_s
            ws_list[i]['E' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
            ws_list[i]['F' + str(oi)].style = input_s
            ws_list[i]['F' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
            ws_list[i]['G' + str(oi)].style = input_s
            ws_list[i]['G' + str(oi)].number_format = "####"
            formula = "=SUM(%s!F%s - %s!E%s)" % (day_of_week[i], str(oi), day_of_week[i], str(oi))
            ws_list[i]['H' + str(oi)] = formula  # move total
            ws_list[i]['H' + str(oi)].style = input_s
            ws_list[i]['H' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
            formula = "=IF(%s!B%s =\"ns day\", %s!C%s,IF(%s!C%s <= 8 + reference!C3, 0, MAX(%s!C%s - 8, 0)))" \
                      % (day_of_week[i], str(oi), day_of_week[i], str(oi), day_of_week[i], str(oi),
                         day_of_week[i], str(oi))
            ws_list[i]['I' + str(oi)] = formula  # overtime
            ws_list[i]['I' + str(oi)].style = calcs
            ws_list[i]['I' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
            formula = "=SUM(%s!F%s - %s!E%s)" \
                      % (day_of_week[i], str(oi), day_of_week[i], str(oi))
            ws_list[i]['J' + str(oi)] = formula  # off route
            ws_list[i]['J' + str(oi)].style = calcs
            ws_list[i]['J' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
            formula = "=IF(OR(%s!B%s=\"ns day\",%s!J%s >= %s!C%s),%s!C%s, IF(%s!C%s <= 8 + reference!C4, 0, MIN" \
                      "(MAX(%s!C%s - 8, 0),IF(%s!J%s <= reference!C4,0, %s!J%s))))" \
                      % (day_of_week[i], str(oi), day_of_week[i], str(oi), day_of_week[i], str(oi),
                         day_of_week[i], str(oi), day_of_week[i], str(oi),
                         day_of_week[i], str(oi), day_of_week[i], str(oi), day_of_week[i], str(oi))
            ws_list[i]['K' + str(oi)] = formula  # OT off route
            ws_list[i]['K' + str(oi)].style = calcs
            ws_list[i]['K' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
            oi += 1
        cell = str(oi - 1)
        oi += 1
        ws_list[i]['J' + str(oi)] = "Total WAL Mandates"
        ws_list[i]['J' + str(oi)].style = col_header
        formula = "=SUM(%s!K%s:%s!K%s)" % (day_of_week[i], top_cell, day_of_week[i], cell)
        ws_list[i]['K' + str(oi)] = formula  # OT off route
        ws_list[i]['K' + str(oi)].style = calcs
        ws_list[i]['K' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
        # formula = "=SUM(%s!K%s + %s!K%s)"%(day_of_week[i],str(oi),day_of_week[i],str(daily_ot_off_rt))
        formula = "=SUM(%s!K%s + %s!K%s)" % (day_of_week[i], str(oi), day_of_week[i], str(nl_totals))
        oi += 2
        ws_list[i]['J' + str(oi)] = "Total Mandates"
        ws_list[i]['J' + str(oi)].style = col_header
        ws_list[i]['K' + str(oi)] = formula  # total ot off route for nl and wal
        ws_list[i]['K' + str(oi)].style = calcs
        ws_list[i]['K' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
        man_ot_day.append(i)  # get the cell information to reference in the summary tab
        man_ot_row.append(oi)
        oi += 1
        try:
            ws_list[i].page_breaks.append(Break(id=oi))
        except:
            ws_list[i].row_breaks.append(Break(id=oi))
        oi += 1
        #  overtime desired list xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
        ws_list[i]['A' + str(oi)] = "Overtime Desired List Carriers"
        ws_list[i]['A' + str(oi)].style = list_header
        oi += 1
        # column headers
        ws_list[i]['E' + str(oi)] = "Availability to:"
        ws_list[i]['E' + str(oi)].style = col_header
        oi += 1
        ws_list[i]['A' + str(oi)] = "Name"
        ws_list[i]['A' + str(oi)].style = col_header
        ws_list[i]['B' + str(oi)] = "note"
        ws_list[i]['B' + str(oi)].style = col_header
        ws_list[i]['C' + str(oi)] = "5200"
        ws_list[i]['C' + str(oi)].style = col_header
        ws_list[i]['D' + str(oi)] = "RS"
        ws_list[i]['D' + str(oi)].style = col_header
        ws_list[i]['E' + str(oi)] = "to 10"
        ws_list[i]['E' + str(oi)].style = col_header
        ws_list[i]['F' + str(oi)] = "to 12"
        ws_list[i]['F' + str(oi)].style = col_header
        oi += 1
        top_cell = str(oi)
        otdl_oi_start = oi
        aval_10_total = 0
        aval_12_total = 0
        for line in dl_otdl:
            match = "miss"
            for each in r_rings:
                if each[0] == str(day) and each[1] == line[1]:  # if the rings record is a match
                    match = "hit"
                    if match == "hit":
                        # find 10 hour availability pending code status
                        if each[4] == "light" or each[4] == "sch chg":
                            aval_10 = 0.00
                        elif each[4] == "no call":
                            aval_10 = 10.00
                        elif each[2].strip() == "":
                            aval_10 = 0.00
                        else:
                            aval_10 = max(10 - float(each[2]), 0)
                        aval_10_total += aval_10  # add to availability total
                        # find 12 hour availability pending code status
                        if each[4] == "light" or each[4] == "sch chg":
                            aval_12 = 0.00
                        elif each[4] == "no call":
                            aval_12 = 12.00
                        elif each[2].strip() == "":
                            aval_12 = 0.00
                        else:
                            aval_12 = max(12 - float(each[2]), 0)
                        aval_12_total += aval_12  # add to availability total

                        # output to the gui
                        ws_list[i]['A' + str(oi)] = each[1]  # name
                        ws_list[i]['A' + str(oi)].style = input_name
                        if each[4] == "none":
                            code = ""  # leave code field blank if 'none'
                        else:
                            code = each[4]
                        ws_list[i]['B' + str(oi)] = code  # code
                        ws_list[i]['B' + str(oi)].style = input_s
                        if each[2].strip() == "":
                            ws_list[i]['C' + str(oi)] = each[2]  # 5200
                        else:
                            ws_list[i]['C' + str(oi)] = float(each[2])  # 5200
                        ws_list[i]['C' + str(oi)].style = input_s
                        ws_list[i]['C' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
                        if each[3].strip() == "":
                            rs = ""  # handle empty RS strings
                        else:
                            rs = float(each[3])
                        ws_list[i]['D' + str(oi)] = rs  # rs
                        ws_list[i]['D' + str(oi)].style = input_s
                        ws_list[i]['D' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
                        formula = "=IF(OR(%s!B%s = \"light\",%s!B%s = \"excused\", %s!B%s = \"sch chg\", %s!B%s = \"annual\", " \
                                  "%s!B%s = \"sick\", %s!C%s >= 10 - reference!C5), 0, IF(%s!B%s = \"no call\", " \
                                  "10, IF(%s!C%s = 0, 0, MAX(10 - %s!C%s, 0))))" % (
                                      day_of_week[i], str(oi), day_of_week[i], str(oi),
                                      day_of_week[i], str(oi), day_of_week[i], str(oi),
                                      day_of_week[i], str(oi), day_of_week[i], str(oi),
                                      day_of_week[i], str(oi), day_of_week[i], str(oi),
                                      day_of_week[i], str(oi))
                        ws_list[i]['E' + str(oi)] = formula  # availability to 10
                        ws_list[i]['E' + str(oi)].style = calcs
                        ws_list[i]['E' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
                        formula = "=IF(OR(%s!B%s = \"light\",%s!B%s = \"excused\", %s!B%s = \"sch chg\", %s!B%s = \"annual\", " \
                                  "%s!B%s = \"sick\", %s!C%s >= 12 - reference!C5), 0, IF(%s!B%s = \"no call\", " \
                                  "12, IF(%s!C%s = 0, 0, MAX(12 - %s!C%s, 0))))" % (
                                      day_of_week[i], str(oi), day_of_week[i], str(oi),
                                      day_of_week[i], str(oi), day_of_week[i], str(oi),
                                      day_of_week[i], str(oi), day_of_week[i], str(oi),
                                      day_of_week[i], str(oi), day_of_week[i], str(oi),
                                      day_of_week[i], str(oi))
                        ws_list[i]['F' + str(oi)] = formula  # availability to 12
                        ws_list[i]['F' + str(oi)].style = calcs
                        ws_list[i]['F' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
                        oi += 1
            # if there is no match, then just printe the name.
            if match == "miss":
                ws_list[i]['A' + str(oi)] = line[1]  # name
                ws_list[i]['A' + str(oi)].style = input_name
                ws_list[i]['B' + str(oi)].style = input_s
                ws_list[i]['B' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
                ws_list[i]['C' + str(oi)].style = input_s
                ws_list[i]['C' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
                ws_list[i]['D' + str(oi)].style = input_s
                ws_list[i]['D' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
                formula = "=IF(OR(%s!B%s = \"light\",%s!B%s = \"excused\", %s!B%s = \"sch chg\", %s!B%s = \"annual\", " \
                          "%s!B%s = \"sick\", %s!C%s >= 10 - reference!C5), 0, IF(%s!B%s = \"no call\", " \
                          "10, IF(%s!C%s = 0, 0, MAX(10 - %s!C%s, 0))))" % (
                              day_of_week[i], str(oi), day_of_week[i], str(oi),
                              day_of_week[i], str(oi), day_of_week[i], str(oi),
                              day_of_week[i], str(oi), day_of_week[i], str(oi),
                              day_of_week[i], str(oi), day_of_week[i], str(oi),
                              day_of_week[i], str(oi))
                ws_list[i]['E' + str(oi)] = formula  # availability to 10
                ws_list[i]['E' + str(oi)].style = calcs
                ws_list[i]['E' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
                formula = "=IF(OR(%s!B%s = \"light\",%s!B%s = \"excused\", %s!B%s = \"sch chg\", %s!B%s = \"annual\", " \
                          "%s!B%s = \"sick\", %s!C%s >= 12 - reference!C5), 0, IF(%s!B%s = \"no call\", " \
                          "12, IF(%s!C%s = 0, 0, MAX(12 - %s!C%s, 0))))" % (
                              day_of_week[i], str(oi), day_of_week[i], str(oi),
                              day_of_week[i], str(oi), day_of_week[i], str(oi),
                              day_of_week[i], str(oi), day_of_week[i], str(oi),
                              day_of_week[i], str(oi), day_of_week[i], str(oi),
                              day_of_week[i], str(oi))
                ws_list[i]['F' + str(oi)] = formula  # availability to 12
                ws_list[i]['F' + str(oi)].style = calcs
                ws_list[i]['F' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
                oi += 1
        otdl_oi_end = oi
        otdl_oi_diff = otdl_oi_end - otdl_oi_start  # find how many lines exist in otdl
        # if the minimum number of rows are not reached, insert blank rows
        e_range = min_ss_otdl - otdl_oi_diff
        if e_range <= 0:
            e_range = 0
        for e in range(e_range):
            ws_list[i]['A' + str(oi)] = ""  # name
            ws_list[i]['A' + str(oi)].style = input_name
            ws_list[i]['B' + str(oi)].style = input_s
            ws_list[i]['B' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
            ws_list[i]['C' + str(oi)].style = input_s
            ws_list[i]['C' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
            ws_list[i]['D' + str(oi)].style = input_s
            ws_list[i]['D' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
            formula = "=IF(OR(%s!B%s = \"light\",%s!B%s = \"excused\", %s!B%s = \"sch chg\", %s!B%s = \"annual\", " \
                      "%s!B%s = \"sick\", %s!C%s >= 10 - reference!C5), 0, IF(%s!B%s = \"no call\", " \
                      "10, IF(%s!C%s = 0, 0, MAX(10 - %s!C%s, 0))))" % (
                          day_of_week[i], str(oi), day_of_week[i], str(oi),
                          day_of_week[i], str(oi), day_of_week[i], str(oi),
                          day_of_week[i], str(oi), day_of_week[i], str(oi),
                          day_of_week[i], str(oi), day_of_week[i], str(oi),
                          day_of_week[i], str(oi))
            ws_list[i]['E' + str(oi)] = formula  # availability to 10
            ws_list[i]['E' + str(oi)].style = calcs
            ws_list[i]['E' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
            formula = "=IF(OR(%s!B%s = \"light\",%s!B%s = \"excused\", %s!B%s = \"sch chg\", %s!B%s = \"annual\", " \
                      "%s!B%s = \"sick\", %s!C%s >= 12 - reference!C5), 0, IF(%s!B%s = \"no call\", " \
                      "12, IF(%s!C%s = 0, 0, MAX(12 - %s!C%s, 0))))" % (
                          day_of_week[i], str(oi), day_of_week[i], str(oi),
                          day_of_week[i], str(oi), day_of_week[i], str(oi),
                          day_of_week[i], str(oi), day_of_week[i], str(oi),
                          day_of_week[i], str(oi), day_of_week[i], str(oi),
                          day_of_week[i], str(oi))
            ws_list[i]['F' + str(oi)] = formula  # availability to 12
            ws_list[i]['F' + str(oi)].style = calcs
            ws_list[i]['F' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
            oi += 1
        oi += 1
        cell = str(oi - 2)
        ws_list[i]['D' + str(oi)] = "Total OTDL Availability"
        ws_list[i]['D' + str(oi)].style = col_header
        formula = "=SUM(%s!E%s:%s!E%s)" % (day_of_week[i], top_cell, day_of_week[i], cell)
        otdl_total = oi
        ws_list[i]['E' + str(oi)] = formula  # availability to 10
        ws_list[i]['E' + str(oi)].style = calcs
        ws_list[i]['E' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
        formula = "=SUM(%s!F%s:%s!F%s)" % (day_of_week[i], top_cell, day_of_week[i], cell)
        ws_list[i]['F' + str(oi)] = formula  # availability to 12
        ws_list[i]['F' + str(oi)].style = calcs
        ws_list[i]['F' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
        oi += 1
        try:
            ws_list[i].page_breaks.append(Break(id=oi))
        except:
            ws_list[i].row_breaks.append(Break(id=oi))
        oi += 1
        # Auxiliary assistance xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
        ws_list[i]['A' + str(oi)] = "Auxiliary Assistance"
        ws_list[i]['A' + str(oi)].style = list_header
        oi += 1
        # column headers
        ws_list[i]['E' + str(oi)] = "Availability to:"
        ws_list[i]['E' + str(oi)].style = col_header
        oi += 1
        ws_list[i]['A' + str(oi)] = "Name"
        ws_list[i]['A' + str(oi)].style = col_header
        ws_list[i]['B' + str(oi)] = "note"
        ws_list[i]['B' + str(oi)].style = col_header
        ws_list[i]['C' + str(oi)] = "5200"
        ws_list[i]['C' + str(oi)].style = col_header
        ws_list[i]['D' + str(oi)] = "RS"
        ws_list[i]['D' + str(oi)].style = col_header
        ws_list[i]['E' + str(oi)] = "to 10"
        ws_list[i]['E' + str(oi)].style = col_header
        ws_list[i]['F' + str(oi)] = "to 11.5"
        ws_list[i]['F' + str(oi)].style = col_header
        oi += 1
        aux_oi_start = oi
        top_cell = str(oi)
        aval_10_total = 0  # initialize variables for availability totals.
        aval_115_total = 0
        for line in dl_aux:
            match = "miss"
            for each in r_rings:
                if each[0] == str(day) and each[1] == line[1]:  # if the rings record is a match
                    match = "hit"
                    if match == "hit":
                        # find 10 hour availability pending code status
                        if each[4] == "light" or each[4] == "sch chg":
                            aval_10 = 0.00
                        elif each[4] == "no call":
                            aval_10 = 10.00
                        elif each[2].strip() == "":
                            aval_10 = 0.00
                        else:
                            aval_10 = max(10 - float(each[2]), 0)
                        aval_10_total += aval_10  # add to availability total
                        # find 11.5 hour availability pending code status
                        if each[4] == "light" or each[4] == "sch chg" or each[4] == "excused":
                            aval_115 = 0.00
                        elif each[4] == "no call":
                            aval_115 = 12.00
                        elif each[2].strip() == "":
                            aval_115 = 0.00
                        else:
                            aval_115 = max(12 - float(each[2]), 0)
                        aval_115_total += aval_115  # add to availability total
                        # output to the gui
                        ws_list[i]['A' + str(oi)] = each[1]  # name
                        ws_list[i]['A' + str(oi)].style = input_name
                        if each[4] == "none":
                            code = ""  # leave code field blank if 'none'
                        else:
                            code = each[4]
                        ws_list[i]['B' + str(oi)] = code  # code
                        ws_list[i]['B' + str(oi)].style = input_s
                        if each[2].strip() == "":
                            ws_list[i]['C' + str(oi)] = each[2]  # 5200
                        else:
                            ws_list[i]['C' + str(oi)] = float(each[2])  # 5200
                        ws_list[i]['C' + str(oi)].style = input_s
                        ws_list[i]['C' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
                        if each[3].strip() == "":
                            rs = ""  # handle empty RS strings
                        else:
                            rs = float(each[3])
                        ws_list[i]['D' + str(oi)] = rs  # rs
                        ws_list[i]['D' + str(oi)].style = input_s
                        ws_list[i]['D' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
                        formula = "=IF(OR(%s!B%s = \"light\",%s!B%s = \"excused\", %s!B%s = \"sch chg\", %s!B%s = \"annual\", " \
                                  "%s!B%s = \"sick\", %s!C%s >= 10 - reference!C5), 0, IF(%s!B%s = \"no call\", " \
                                  "10, IF(%s!C%s = 0, 0, MAX(10 - %s!C%s, 0))))" % (
                                      day_of_week[i], str(oi), day_of_week[i], str(oi),
                                      day_of_week[i], str(oi), day_of_week[i], str(oi),
                                      day_of_week[i], str(oi), day_of_week[i], str(oi),
                                      day_of_week[i], str(oi), day_of_week[i], str(oi),
                                      day_of_week[i], str(oi))
                        ws_list[i]['E' + str(oi)] = formula  # availability to 10
                        ws_list[i]['E' + str(oi)].style = calcs
                        ws_list[i]['E' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
                        formula = "=IF(OR(%s!B%s = \"light\",%s!B%s = \"excused\", %s!B%s = \"sch chg\", %s!B%s = \"annual\", " \
                                  "%s!B%s = \"sick\", %s!C%s >= 11.5 - reference!C5), 0, IF(%s!B%s = \"no call\", " \
                                  "11.5, IF(%s!C%s = 0, 0, MAX(11.5 - %s!C%s, 0))))" % (
                                      day_of_week[i], str(oi), day_of_week[i], str(oi),
                                      day_of_week[i], str(oi), day_of_week[i], str(oi),
                                      day_of_week[i], str(oi), day_of_week[i], str(oi),
                                      day_of_week[i], str(oi), day_of_week[i], str(oi),
                                      day_of_week[i], str(oi))
                        ws_list[i]['F' + str(oi)] = formula  # availability to 12
                        ws_list[i]['F' + str(oi)].style = calcs
                        ws_list[i]['F' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
                        oi += 1
            # if there is no match, then just printe the name.
            if match == "miss":
                if match == "miss":
                    ws_list[i]['A' + str(oi)] = line[1]  # name
                    ws_list[i]['A' + str(oi)].style = input_name
                    ws_list[i]['B' + str(oi)].style = input_s
                    ws_list[i]['B' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
                    ws_list[i]['C' + str(oi)].style = input_s
                    ws_list[i]['C' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
                    ws_list[i]['D' + str(oi)].style = input_s
                    ws_list[i]['D' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
                    formula = "=IF(OR(%s!B%s = \"light\",%s!B%s = \"excused\", %s!B%s = \"sch chg\", %s!B%s = \"annual\", " \
                              "%s!B%s = \"sick\", %s!C%s >= 10 - reference!C5), 0, IF(%s!B%s = \"no call\", " \
                              "10, IF(%s!C%s = 0, 0, MAX(10 - %s!C%s, 0))))" % (
                                  day_of_week[i], str(oi), day_of_week[i], str(oi),
                                  day_of_week[i], str(oi), day_of_week[i], str(oi),
                                  day_of_week[i], str(oi), day_of_week[i], str(oi),
                                  day_of_week[i], str(oi), day_of_week[i], str(oi),
                                  day_of_week[i], str(oi))
                    ws_list[i]['E' + str(oi)] = formula  # availability to 10
                    ws_list[i]['E' + str(oi)].style = calcs
                    ws_list[i]['E' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
                    formula = "=IF(OR(%s!B%s = \"light\",%s!B%s = \"excused\", %s!B%s = \"sch chg\", %s!B%s = \"annual\", " \
                              "%s!B%s = \"sick\", %s!C%s >= 11.5 - reference!C5), 0, IF(%s!B%s = \"no call\", " \
                              "11.5, IF(%s!C%s = 0, 0, MAX(11.5 - %s!C%s, 0))))" % (
                                  day_of_week[i], str(oi), day_of_week[i], str(oi),
                                  day_of_week[i], str(oi), day_of_week[i], str(oi),
                                  day_of_week[i], str(oi), day_of_week[i], str(oi),
                                  day_of_week[i], str(oi), day_of_week[i], str(oi),
                                  day_of_week[i], str(oi))
                    ws_list[i]['F' + str(oi)] = formula  # availability to 12
                    ws_list[i]['F' + str(oi)].style = calcs
                    ws_list[i]['F' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
                    oi += 1
        aux_oi_end = oi
        aux_oi_diff = aux_oi_end - aux_oi_start  # find how many lines exist in aux
        # if the minimum number of rows are not reached, insert blank rows
        e_range = min_ss_aux - aux_oi_diff
        if e_range <= 0:
            e_range = 0
        for e in range(e_range):
            ws_list[i]['A' + str(oi)] = ""  # name
            ws_list[i]['A' + str(oi)].style = input_name
            ws_list[i]['B' + str(oi)].style = input_s
            ws_list[i]['B' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
            ws_list[i]['C' + str(oi)].style = input_s
            ws_list[i]['C' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
            ws_list[i]['D' + str(oi)].style = input_s
            ws_list[i]['D' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
            formula = "=IF(OR(%s!B%s = \"light\",%s!B%s = \"excused\", %s!B%s = \"sch chg\", %s!B%s = \"annual\", " \
                      "%s!B%s = \"sick\", %s!C%s >= 10 - reference!C5), 0, IF(%s!B%s = \"no call\", " \
                      "10, IF(%s!C%s = 0, 0, MAX(10 - %s!C%s, 0))))" % (
                          day_of_week[i], str(oi), day_of_week[i], str(oi),
                          day_of_week[i], str(oi), day_of_week[i], str(oi),
                          day_of_week[i], str(oi), day_of_week[i], str(oi),
                          day_of_week[i], str(oi), day_of_week[i], str(oi),
                          day_of_week[i], str(oi))
            ws_list[i]['E' + str(oi)] = formula  # availability to 10
            ws_list[i]['E' + str(oi)].style = calcs
            ws_list[i]['E' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
            formula = "=IF(OR(%s!B%s = \"light\",%s!B%s = \"excused\", %s!B%s = \"sch chg\", %s!B%s = \"annual\", " \
                      "%s!B%s = \"sick\", %s!C%s >= 12 - reference!C5), 0, IF(%s!B%s = \"no call\", " \
                      "12, IF(%s!C%s = 0, 0, MAX(12 - %s!C%s, 0))))" % (
                          day_of_week[i], str(oi), day_of_week[i], str(oi),
                          day_of_week[i], str(oi), day_of_week[i], str(oi),
                          day_of_week[i], str(oi), day_of_week[i], str(oi),
                          day_of_week[i], str(oi), day_of_week[i], str(oi),
                          day_of_week[i], str(oi))
            ws_list[i]['F' + str(oi)] = formula  # availability to 12
            ws_list[i]['F' + str(oi)].style = calcs
            ws_list[i]['F' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
            oi += 1
        oi += 1
        cell = str(oi - 2)
        ws_list[i]['D' + str(oi)] = "Total AUX Availability"
        ws_list[i]['D' + str(oi)].style = col_header
        formula = "=SUM(%s!E%s:%s!E%s)" % (day_of_week[i], top_cell, day_of_week[i], cell)
        aux_total = oi
        ws_list[i]['E' + str(oi)] = formula  # availability to 10
        ws_list[i]['E' + str(oi)].style = calcs
        ws_list[i]['E' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
        formula = "=SUM(%s!F%s:%s!F%s)" % (day_of_week[i], top_cell, day_of_week[i], cell)
        ws_list[i]['F' + str(oi)] = formula  # availability to 11.5
        ws_list[i]['F' + str(oi)].style = calcs
        ws_list[i]['F' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
        oi += 2
        ws_list[i]['D' + str(oi)] = "Total Availability"
        ws_list[i]['D' + str(oi)].style = col_header
        formula = "=SUM(%s!E%s + %s!E%s)" % (day_of_week[i], otdl_total, day_of_week[i], aux_total)
        ws_list[i]['E' + str(oi)] = formula  # availability to 10
        ws_list[i]['E' + str(oi)].style = calcs
        ws_list[i]['E' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
        av_to_10_day.append(i)
        av_to_10_row.append(oi)
        formula = "=SUM(%s!F%s + %s!F%s)" % (day_of_week[i], otdl_total, day_of_week[i], aux_total)
        ws_list[i]['F' + str(oi)] = formula  # availability to 11.5
        ws_list[i]['F' + str(oi)].style = calcs
        ws_list[i]['F' + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
        av_to_12_day.append(i)
        av_to_12_row.append(oi)
        oi += 1
        i += 1
    # summary page xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
    summary.column_dimensions["A"].width = 14
    summary.column_dimensions["B"].width = 9
    summary.column_dimensions["C"].width = 9
    summary.column_dimensions["D"].width = 9
    summary.column_dimensions["E"].width = 2
    summary.column_dimensions["F"].width = 9
    summary.column_dimensions["G"].width = 9
    summary.column_dimensions["H"].width = 9
    summary['A1'] = "Improper Mandate Worksheet"
    summary['A1'].style = ws_header
    summary.merge_cells('A1:E1')
    summary['B3'] = "Summary Sheet"
    summary['B3'].style = date_dov_title
    summary['A5'] = "Pay Period:  "
    summary['A5'].style = date_dov_title
    summary['B5'] = pay_period
    summary['B5'].style = date_dov
    summary.merge_cells('B5:D5')

    summary['A6'] = "Station:  "
    summary['A6'].style = date_dov_title
    summary['B6'] = g_station
    summary['B6'].style = date_dov
    summary.merge_cells('B6:D6')
    summary['B8'] = "Availability"
    summary['B8'].style = date_dov_title
    summary['B9'] = "to 10"
    summary['B9'].style = date_dov_title
    summary['C8'] = "No list"
    summary['C8'].style = date_dov_title
    summary['C9'] = "overtime"
    summary['C9'].style = date_dov_title
    summary['D9'] = "violations"
    summary['D9'].style = date_dov_title
    summary['F8'] = "Availability"
    summary['F8'].style = date_dov_title
    summary['F9'] = "to 12"
    summary['F9'].style = date_dov_title
    summary['G8'] = "Off route"
    summary['G8'].style = date_dov_title
    summary['G9'] = "mandates"
    summary['G9'].style = date_dov_title
    summary['H9'] = "violations"
    summary['H9'].style = date_dov_title
    row = 10
    if g_range == "week":
        range_num = 7
    if g_range == "day":
        range_num = 1
    for i in range(range_num):
        summary['A' + str(row)] = format(dates[i], "%m/%d/%y %a")
        summary['A' + str(row)].style = date_dov_title
        summary['A' + str(row)].number_format = "#,###.00;[RED]-#,###.00"
        summary['B' + str(row)] = "=%s!E%s" % (day_of_week[av_to_10_day[i]], av_to_10_row[i])  # availability to 10
        summary['B' + str(row)].style = input_s
        summary['B' + str(row)].number_format = "#,###.00;[RED]-#,###.00"
        summary['C' + str(row)] = "=%s!I%s" % (day_of_week[nl_ot_day[i]], nl_ot_row[i])  # no list OT
        summary['C' + str(row)].style = input_s
        summary['C' + str(row)].number_format = "#,###.00;[RED]-#,###.00"
        summary['D' + str(row)] = "=IF(%s!B%s<%s!C%s,%s!B%s,%s!C%s)" \
                                  % ('summary', str(row), 'summary', str(row), 'summary', str(row), 'summary', str(row))
        summary['D' + str(row)].style = calcs
        summary['D' + str(row)].number_format = "#,###.00;[RED]-#,###.00"
        summary['F' + str(row)] = "=%s!F%s" % (day_of_week[av_to_12_day[i]], av_to_12_row[i])  # availability to 12
        summary['F' + str(row)].style = input_s
        summary['F' + str(row)].number_format = "#,###.00;[RED]-#,###.00"
        summary['G' + str(row)] = "=%s!K%s" % (day_of_week[man_ot_day[i]], man_ot_row[i])  # total mandates
        summary['G' + str(row)].style = input_s
        summary['G' + str(row)].number_format = "#,###.00;[RED]-#,###.00"
        summary['H' + str(row)] = "=IF(%s!F%s<%s!G%s,%s!F%s,%s!G%s)" \
                                  % ('summary', str(row), 'summary', str(row), 'summary', str(row), 'summary', str(row))
        summary['H' + str(row)].style = calcs
        summary['H' + str(row)].number_format = "#,###.00;[RED]-#,###.00"
        row = row + 2
    # reference page xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
    reference.column_dimensions["A"].width = 14
    reference.column_dimensions["B"].width = 8
    reference.column_dimensions["C"].width = 8
    reference.column_dimensions["D"].width = 2
    reference.column_dimensions["E"].width = 6
    sql = "SELECT tolerance FROM tolerances"
    tolerances = inquire(sql)
    reference['B2'].style = list_header
    reference['B2'] = "Tolerances"
    reference['C3'] = float(tolerances[0][0])  # overtime on own route tolerance
    reference['C3'].style = input_s
    reference['C3'].number_format = "#,###.00;[RED]-#,###.00"
    reference['E3'] = "overtime on own route"
    reference['C4'] = float(tolerances[1][0])  # overtime off own route tolerance
    reference['C4'].style = input_s
    reference['C4'].number_format = "#,###.00;[RED]-#,###.00"
    reference['E4'] = "overtime off own route"
    reference['C5'] = float(tolerances[2][0])  # availability tolerance
    reference['C5'].style = input_s
    reference['C5'].number_format = "#,###.00;[RED]-#,###.00"
    reference['E5'] = "availability tolerance"
    reference['B7'].style = list_header
    reference['B7'] = "Code Guide"
    reference['C8'] = "ns day"
    reference['C8'].style = input_s
    reference['E8'] = "Carrier worked on their non scheduled day"
    reference['C10'] = "no call"
    reference['C10'].style = input_s
    reference['E10'] = "Carrier was not scheduled for overtime"
    reference['C11'] = "light"
    reference['C11'].style = input_s
    reference['E11'] = "Carrier on light duty and unavailable for overtime"
    reference['C12'] = "sch chg"
    reference['C12'].style = input_s
    reference['E12'] = "Schedule change: unavailable for overtime"
    reference['C13'] = "annual"
    reference['C13'].style = input_s
    reference['E13'] = "Annual leave"
    reference['C14'] = "sick"
    reference['C14'].style = input_s
    reference['E14'] = "Sick leave"
    reference['C15'] = "excused"
    reference['C15'].style = input_s
    reference['E15'] = "Carrier excused from mandatory overtime"
    # name the excel file
    r = "_w"
    if g_range == "day": r = "_d"
    xl_filename = "kb" + str(format(dates[0], "_%y_%m_%d")) + r + ".xlsx"
    ok = messagebox.askokcancel("Spreadsheet generator", "Do you want to generate a spreadsheet?")
    if ok == True:
        if os.path.isdir('kb_sub/spreadsheets') == False:
            os.makedirs('kb_sub/spreadsheets')
        try:
            wb.save('kb_sub/spreadsheets/' + xl_filename)
            messagebox.showinfo("Spreadsheet generator", "Your spreadsheet was successfully generated. \n"
                                                         "File is named: {}".format(xl_filename))
            if sys.platform == "win32":
                os.startfile('kb_sub\\spreadsheets\\' + xl_filename)
            if sys.platform == "linux":
                subprocess.call(["xdg-open", 'kb_sub/spreadsheets/' + xl_filename])
            if sys.platform == "darwin":
                subprocess.call(["open", 'kb_sub/spreadsheets/' + xl_filename])
        except:
            messagebox.showerror("Spreadsheet generator", "The spreadsheet was not generated. \n"
                                                          "Suggestion: "
                                                          "Make sure that identically named spreadsheets are closed "
                                                          "(the file can't be overwritten while open).")


def tab_selected(t):  # attach notebook tab for
    global current_tab
    current_tab = t


def output_tab(self, list_carrier):
    self.destroy()
    switchF5 = Frame(root, bg="white")
    switchF5.pack(fill=BOTH, side=LEFT)
    C1 = Canvas(switchF5)
    C1.pack(fill=BOTH, side=BOTTOM)
    # Button(C1, text="new carrier", command=lambda: input_carriers(switchF5),
    #        width=15).pack(side=LEFT)
    Button(C1, text="spreadsheet", width=15, anchor="w",
           command=lambda: [spreadsheet(list_carrier, r_rings)]).pack(side=LEFT)
    Button(C1, text="Go Back", width=15, anchor="w",
           command=lambda: [switchF5.destroy(), main_frame()]).pack(side=LEFT)
    dates = []  # array containing days
    if g_range == "week": dates = g_date
    if g_range == "day": dates.append(d_date)

    if g_range == "week":
        sql = "SELECT * FROM rings3 WHERE rings_date BETWEEN '%s' AND '%s' ORDER BY rings_date, carrier_name" \
              % (g_date[0], g_date[6])
    else:
        sql = "SELECT * FROM rings3 WHERE rings_date = '%s' ORDER BY rings_date, " \
              "carrier_name" % (d_date)
    r_rings = inquire(sql)
    sql = "SELECT * FROM tolerances"  # get tolerances
    tolerances = inquire(sql)
    ot_own_rt = tolerances[0][2]
    ot_tol = tolerances[1][2]
    av_tol = tolerances[2][2]
    daily_list = []  # array
    candidates = []
    dl_nl = []
    dl_wal = []
    dl_otdl = []
    dl_aux = []
    # list the names of the tabs
    tab = ["saturday", "sunday", "monday", "tuesday", "wednesday", "thursday", "friday"]
    C = ["C0", "C1", "C2", "C3", "C4", "C5", "C6"]
    global current_tab
    current_tab = 0
    tabControl = ttk.Notebook(switchF5)  # Create Tab Control
    tabControl.pack(expand=1, fill="both")
    t = 0
    # for day in dates:
    for day in dates:
        del daily_list[:]
        del dl_nl[:]
        del dl_wal[:]
        del dl_otdl[:]
        del dl_aux[:]
        # create a list of carriers for each day.
        for i in range(len(list_carrier)):
            if list_carrier[i][0] <= str(day):
                candidates.append(list_carrier[i])  # put name into candidates array
            jump = "no"  # triggers an analysis of the candidates array
            if i != len(list_carrier) - 1:  # if the loop has not reached the end of the list
                if list_carrier[i][1] == list_carrier[i + 1][1]:  # if the name current and next name are the same
                    jump = "yes"  # bypasses an analysis of the candidates array
            if jump == "no":  # review the list of candidates
                winner = max(candidates, key=itemgetter(0))  # select the most recent
                if winner[5] == g_station: daily_list.append(winner)  # add the record if it matches the station
                del candidates[:]  # empty out the candidates array.
        for item in daily_list:  # sort carriers in daily list by the list they are in
            if item[2] == "nl":
                dl_nl.append(item)
            if item[2] == "wal":
                dl_wal.append(item)
            if item[2] == "otdl":
                dl_otdl.append(item)
            if item[2] == "aux":
                dl_aux.append(item)
        tabs = Frame(tabControl)  # put frame in notebook
        tabs.pack(fill=BOTH, side=LEFT)
        if g_range == "week": tabControl.add(tabs, text="{}".format(tab[t]))  # Add the tab
        C[t] = Canvas(tabs, width=1600, bg="white")  # put canvas inside notebook frame
        S = Scrollbar(tabs, command=C[t].yview)  # define and bind the scrollbar with the canvas
        C[t].config(yscrollcommand=S.set, scrollregion=(0, 0, 100, 5000))  # bind the canvas with the scrollbar
        #   Enable mousewheel
        C[t].bind("<Map>", lambda event, t=t: tab_selected(t))
        # C[current_tab].bind_all('<MouseWheel>', lambda event: C[current_tab].yview_scroll(int(-1 * (event.delta / 120)), "units"))
        if sys.platform == "win32":
            C[current_tab].bind_all('<MouseWheel>',
                                    lambda event: C[current_tab].yview_scroll(int(-1 * (event.delta / 120)), "units"))
        elif sys.platform == "linux":
            C[current_tab].bind_all('<Button-4>', lambda event: C[current_tab].yview('scroll', -1, 'units'))
            C[current_tab].bind_all('<Button-5>', lambda event: C[current_tab].yview('scroll', 1, 'units'))

        S.pack(side=RIGHT, fill=BOTH)
        C[t].pack(side=LEFT, fill=BOTH, expand=True)
        F = Frame(C[t], bg="white")  # put a frame in the canvas
        F.pack()
        C[t].create_window((0, 0), window=F, anchor=NW)  # create window with frame
        oi = 0
        Label(F, text=day.strftime("%A  %m/%d/%y"), justify=LEFT, anchor=W, font="bold",
              pady=5, bg="white").grid(row=oi, column=0, columnspan=10, sticky=W)
        in_color = "white"
        out_color = "light goldenrod yellow"
        oi += 1
        #  no list xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
        Label(F, text="no list", justify=LEFT, bg="white", font=('Helvetica 10 bold')) \
            .grid(sticky=W, row=oi, column=0, columnspan=9)
        oi += 1
        Label(F, text="move", bg="white").grid(row=oi, column=7)  # top of move total
        Label(F, text="off", bg="white").grid(row=oi, column=9)  # top of off route
        Label(F, text="ot off", bg="white").grid(row=oi, column=10)  # top of ot off route
        oi += 1
        Label(F, text="Carrier", bg="white").grid(row=oi, column=0, sticky=W)
        Label(F, text="note", bg="white").grid(row=oi, column=1)
        Label(F, text="5200", bg="white").grid(row=oi, column=2)
        Label(F, text="RS", bg="white").grid(row=oi, column=3)
        Label(F, text="MV off", bg="white").grid(row=oi, column=4)
        Label(F, text="MV on", bg="white").grid(row=oi, column=5)
        Label(F, text="Route", bg="white").grid(row=oi, column=6)
        Label(F, text="total", bg="white").grid(row=oi, column=7)
        Label(F, text="OT", bg="white").grid(row=oi, column=8)
        Label(F, text="route", bg="white").grid(row=oi, column=9)
        Label(F, text="route", bg="white").grid(row=oi, column=10)
        oi += 1
        move_totals = []
        ot_total = 0
        ot_off_total = 0
        for line in dl_nl:
            match = "miss"
            del move_totals[:]  # empty array of moves totals.
            # if there is a ring to match the carrier/ date then printe
            for each in r_rings:
                if each[0] == str(day) and each[1] == line[1]:  # if the rings record is a match
                    match = "hit"
                    if match == "hit":
                        s_moves = each[5].split(",")  # converts str to array
                        c = 0
                        for i in range(int(len(s_moves) / 3)):
                            total = float(s_moves[c + 1]) - float(s_moves[c])  # calc off time off route
                            c = c + 3
                            move_totals.append(total)
                        off_route = 0.0
                        if str(each[2]) != "":  # in case the 5200 time is blank
                            time5200 = each[2]
                        else:
                            time5200 = 0
                        if each[4] == "ns day":  # if the carrier worked on their ns day
                            off_route = float(time5200)  # cal >off route
                            ot = float(time5200)  # cal > ot
                        else:  # if carrier did not work ns day
                            ot = max(float(time5200) - float(8), 0)  # calculate overtime
                            if ot <= float(ot_own_rt): ot = 0  # adjust sum for tolerance
                            for mt in move_totals:  # cal off route work.
                                off_route += float(mt)
                        ot_total += ot
                        ot_off_route = min(off_route, ot)  # calculate the ot off route
                        if ot_off_route <= float(ot_tol): ot_off_route = 0  # adjust sum for tolerance
                        ot_off_total += ot_off_route
                        move_count = (int(len(s_moves) / 3))  # find the number of sets of moves
                        # output to the gui
                        Label(F, text=each[1], anchor=W, width=21, relief=RIDGE, bg=in_color) \
                            .grid(row=oi, column=0)  # name
                        if each[4] == "none":
                            code = ""
                        else:
                            code = each[4]
                        Label(F, text=code, justify=LEFT, width=6, relief=RIDGE, bg=in_color) \
                            .grid(row=oi, column=1)  # code
                        if each[2] == "" or each[2] == " ":  # handle empty 5200 strings
                            t_hrs = ""
                        else:
                            t_hrs = format(float(each[2]), '.2f')
                        Label(F, text=t_hrs, justify=LEFT, width=6, relief=RIDGE, bg=in_color) \
                            .grid(row=oi, column=2)  # 5200
                        if each[3] == "" or each[3] == " ":
                            rs = ""  # handle empty RS strings
                        else:
                            rs = format(float(each[3]), '.2f')
                        Label(F, text=rs, justify=LEFT, width=6, relief=RIDGE, bg=in_color) \
                            .grid(row=oi, column=3)  # return to station
                        count = 0
                        if move_count == 0:  # if there are no moves, fill in with empty cells.
                            for i in range(4, 8):
                                if i < 7:
                                    color = in_color
                                else:
                                    color = out_color
                                Label(F, text="", justify=LEFT, width=6,
                                      relief=RIDGE, bg=color).grid(row=oi, column=i)
                        for i in range(move_count):  # if there are moves, create + populate cells
                            Label(F, text=format(float(s_moves[count]), '.2f'), justify=LEFT, width=6,
                                  relief=RIDGE, bg=in_color).grid(row=oi, column=4)  # move off
                            count += 1
                            Label(F, text=format(float(s_moves[count]), '.2f'), justify=LEFT, width=6,
                                  relief=RIDGE, bg=in_color).grid(row=oi, column=5)  # move on
                            count += 1
                            Label(F, text=s_moves[count], justify=LEFT, width=6,
                                  relief=RIDGE, bg=in_color).grid(row=oi, column=6)  # route
                            count += 1
                            Label(F, text=format(move_totals[i], '.2f'), justify=LEFT, width=6,
                                  relief=RIDGE, bg=out_color).grid(row=oi, column=7)  # move total
                            if i < move_count - 1: oi += 1
                        Label(F, text=format(ot, '.2f'), justify=LEFT, width=6, relief=RIDGE, bg=out_color) \
                            .grid(row=oi, column=8)  # overtime
                        Label(F, text=format(off_route, '.2f'), justify=LEFT, width=6, relief=RIDGE, bg=out_color) \
                            .grid(row=oi, column=9)  # off route
                        Label(F, text=format(ot_off_route, '.2f'), justify=LEFT, width=6, relief=RIDGE, bg=out_color) \
                            .grid(row=oi, column=10)  # OT off route
                        oi += 1
            #  if there is no match, then just printe the name.
            if match == "miss":
                Label(F, text=line[1], anchor=W, width=21, relief=RIDGE, bg=in_color) \
                    .grid(row=oi, column=0)  # name
                for i in range(10):
                    if i < 6:
                        color = in_color
                    else:
                        color = out_color
                    Label(F, text="", width=6, relief=RIDGE, bg=color) \
                        .grid(row=oi, column=i + 1)  # generate blank cells
                oi += 1
        oi += 1
        Label(F, text="", height=2, bg="white").grid(row=oi, column=0)
        Label(F, text=format(ot_total, '.2f'), justify=LEFT, width=6, relief=RIDGE, bg=out_color) \
            .grid(row=oi, column=8)  # overtime
        Label(F, text=format(ot_off_total, '.2f'), justify=LEFT, width=6, relief=RIDGE, bg=out_color) \
            .grid(row=oi, column=10)  # OT off route
        oi += 2
        # work assignment list xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
        Label(F, text="work assignment list", justify=LEFT, font=('Helvetica 10 bold'), bg="white") \
            .grid(sticky=W, row=oi, column=0, columnspan=9)
        oi += 1
        Label(F, text="move", bg="white").grid(row=oi, column=7)  # top of move total
        Label(F, text="off", bg="white").grid(row=oi, column=9)  # top of off route
        Label(F, text="ot off", bg="white").grid(row=oi, column=10)  # top of ot off route
        oi += 1
        Label(F, text="Carrier", bg="white").grid(row=oi, column=0, sticky=W)
        Label(F, text="note", bg="white").grid(row=oi, column=1)
        Label(F, text="5200", bg="white").grid(row=oi, column=2)
        Label(F, text="RS", bg="white").grid(row=oi, column=3)
        Label(F, text="MV off", bg="white").grid(row=oi, column=4)
        Label(F, text="MV on", bg="white").grid(row=oi, column=5)
        Label(F, text="Route", bg="white").grid(row=oi, column=6)
        Label(F, text="total", bg="white").grid(row=oi, column=7)
        Label(F, text="OT", bg="white").grid(row=oi, column=8)
        Label(F, text="route", bg="white").grid(row=oi, column=9)
        Label(F, text="route", bg="white").grid(row=oi, column=10)
        oi += 1
        move_totals = []
        ot_total = 0
        ot_off_total = 0
        for line in dl_wal:
            match = "miss"
            del move_totals[:]  # empty array of moves totals.
            # if there is a ring to match the carrier/ date then printe
            for each in r_rings:
                if each[0] == str(day) and each[1] == line[1]:  # if the rings record is a match
                    match = "hit"
                    if match == "hit":
                        s_moves = each[5].split(",")
                        c = 0
                        for i in range(int(len(s_moves) / 3)):
                            total = float(s_moves[c + 1]) - float(s_moves[c])  # calc off time off route
                            c = c + 3
                            move_totals.append(total)
                        off_route = 0.0
                        if str(each[2]) != "":  # in case the 5200 time is blank
                            time5200 = each[2]
                        else:
                            time5200 = 0
                        if each[4] == "ns day":  # if the carrier worked on their ns day
                            off_route = float(time5200)  # cal >off route
                            ot = float(time5200)  # cal > ot
                        else:  # if carrier did not work ns day
                            ot = max(float(time5200) - float(8), 0)  # calculate overtime
                            for mt in move_totals:  # cal off route work.
                                off_route += float(mt)
                        ot_total += ot
                        ot_off_route = min(off_route, ot)  # calculate the ot off route
                        if ot_off_route <= float(ot_tol): ot_off_route = 0  # adjust sum for tolerance
                        ot_off_total += ot_off_route
                        move_count = (int(len(s_moves) / 3))  # find the number of sets of moves
                        # output to the gui
                        Label(F, text=each[1], anchor=W, width=21, relief=RIDGE, bg=in_color) \
                            .grid(row=oi, column=0)  # name
                        if each[4] == "none":
                            code = ""
                        else:
                            code = each[4]
                        Label(F, text=code, justify=LEFT, width=6, relief=RIDGE, bg=in_color) \
                            .grid(row=oi, column=1)  # code
                        if each[2] == "" or each[2] == " ":  # handle empty 5200 strings
                            t_hrs = ""
                        else:
                            t_hrs = format(float(each[2]), '.2f')
                        Label(F, text=t_hrs, justify=LEFT, width=6, relief=RIDGE, bg=in_color) \
                            .grid(row=oi, column=2)  # 5200
                        if each[3] == "" or each[3] == " ":
                            rs = ""  # handle empty RS strings
                        else:
                            rs = format(float(each[3]), '.2f')
                        Label(F, text=rs, justify=LEFT, width=6, relief=RIDGE, bg=in_color) \
                            .grid(row=oi, column=3)  # return to station
                        count = 0
                        if move_count == 0:  # if there are no moves, fill in with empty cells.
                            for i in range(4, 8):
                                if i < 7:
                                    color = in_color
                                else:
                                    color = out_color
                                Label(F, text="", justify=LEFT, width=6,
                                      relief=RIDGE, bg=color).grid(row=oi, column=i)
                        for i in range(move_count):  # if there are moves, create + populate cells
                            Label(F, text=format(float(s_moves[count]), '.2f'), justify=LEFT, width=6,
                                  relief=RIDGE, bg=in_color).grid(row=oi, column=4)  # move off
                            count += 1
                            Label(F, text=format(float(s_moves[count]), '.2f'), justify=LEFT, width=6,
                                  relief=RIDGE, bg=in_color).grid(row=oi, column=5)  # move on
                            count += 1
                            Label(F, text=s_moves[count], justify=LEFT, width=6,
                                  relief=RIDGE, bg=in_color).grid(row=oi, column=6)  # route
                            count += 1
                            Label(F, text=format(move_totals[i], '.2f'), justify=LEFT, width=6,
                                  relief=RIDGE, bg=out_color).grid(row=oi, column=7)  # move total
                            if i < move_count - 1: oi += 1
                        Label(F, text=format(ot, '.2f'), justify=LEFT, width=6, relief=RIDGE, bg=out_color) \
                            .grid(row=oi, column=8)  # overtime
                        Label(F, text=format(off_route, '.2f'), justify=LEFT, width=6, relief=RIDGE, bg=out_color) \
                            .grid(row=oi, column=9)  # off route
                        Label(F, text=format(ot_off_route, '.2f'), justify=LEFT, width=6, relief=RIDGE, bg=out_color) \
                            .grid(row=oi, column=10)  # OT off route
                        oi += 1
            # if there is no match, then just printe the name.
            if match == "miss":
                Label(F, text=line[1], anchor=W, width=21, relief=RIDGE, bg=in_color) \
                    .grid(row=oi, column=0)  # name
                for i in range(10):
                    if i < 6:
                        color = in_color
                    else:
                        color = out_color
                    Label(F, text="", width=6, relief=RIDGE, bg=color) \
                        .grid(row=oi, column=i + 1)  # generate blank cells
                oi += 1
        oi += 1
        Label(F, text="", height=2, bg="white").grid(row=oi, column=0)
        Label(F, text=format(ot_total, '.2f'), justify=LEFT, width=6, relief=RIDGE, bg=out_color) \
            .grid(row=oi, column=8)  # overtime
        Label(F, text=format(ot_off_total, '.2f'), justify=LEFT, width=6, relief=RIDGE, bg=out_color) \
            .grid(row=oi, column=10)  # OT off route
        oi += 2
        #  overtime desired list xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
        Label(F, text="overtime desired list", justify=LEFT, font=('Helvetica 10 bold'), bg="white") \
            .grid(sticky=W, row=oi, column=0, columnspan=9)
        oi += 1
        Label(F, text="Availability to: ", bg="white").grid(row=oi, column=4, columnspan=3, sticky=W)
        oi += 1
        Label(F, text="Carrier", bg="white").grid(row=oi, column=0, sticky=W)
        Label(F, text="note", bg="white").grid(row=oi, column=1)
        Label(F, text="5200", bg="white").grid(row=oi, column=2)
        Label(F, text="RS", bg="white").grid(row=oi, column=3)
        Label(F, text="to 10", bg="white").grid(row=oi, column=4)
        Label(F, text="to 12", bg="white").grid(row=oi, column=5)
        oi += 1
        aval_10_total = 0
        aval_12_total = 0
        for line in dl_otdl:
            match = "miss"
            for each in r_rings:
                if each[0] == str(day) and each[1] == line[1]:  # if the rings record is a match
                    match = "hit"
                    if match == "hit":
                        # find 10 hour availability pending code status
                        if each[4] == "light" or each[4] == "sch chg" or each[4] == "excused":
                            aval_10 = 0.00
                        elif each[4] == "no call":
                            aval_10 = 10.00
                        elif each[4] == "sick" or each[4] == "annual":
                            aval_10 = 0.00
                        elif each[2].strip() == "":
                            aval_10 = 0.00
                        else:
                            aval_10 = max(10 - float(each[2]), 0)
                        if aval_10 <= float(av_tol): aval_10 = 0  # adjust sum for tolerance
                        aval_10_total += aval_10  # add to availability total
                        # find 12 hour availability pending code status
                        if each[4] == "light" or each[4] == "sch chg" or each[4] == "excused":
                            aval_12 = 0.00
                        elif each[4] == "no call":
                            aval_12 = 12.00
                        elif each[4] == "sick" or each[4] == "annual":
                            aval_12 = 0.00
                        elif each[2].strip() == "":
                            aval_12 = 0.00
                        else:
                            aval_12 = max(12 - float(each[2]), 0)
                        if aval_12 <= float(av_tol): aval_12 = 0  # adjust sum for tolerance
                        aval_12_total += aval_12  # add to availability total
                        # output to the gui
                        Label(F, text=each[1], anchor=W, width=21, relief=RIDGE, bg=in_color) \
                            .grid(row=oi, column=0)  # name
                        if each[4] == "none":
                            code = ""
                        else:
                            code = each[4]
                        Label(F, text=code, justify=LEFT, width=6, relief=RIDGE, bg=in_color) \
                            .grid(row=oi, column=1)  # code
                        if each[2] == "" or each[2] == " ":  # handle empty 5200 strings
                            t_hrs = ""
                        else:
                            t_hrs = format(float(each[2]), '.2f')
                        Label(F, text=t_hrs, justify=LEFT, width=6, relief=RIDGE, bg=in_color) \
                            .grid(row=oi, column=2)  # 5200
                        if each[3] == "" or each[3] == " ":  # handle empty RS strings
                            rs = ""
                        else:
                            rs = format(float(each[3]), '.2f')
                        Label(F, text=rs, justify=LEFT, width=6, relief=RIDGE, bg=in_color) \
                            .grid(row=oi, column=3)  # return to station
                        Label(F, text=format(float(aval_10), '.2f'), justify=LEFT, width=6, relief=RIDGE, bg=out_color) \
                            .grid(row=oi, column=4)  # availability to 10
                        Label(F, text=format(float(aval_12), '.2f'), justify=LEFT, width=6, relief=RIDGE, bg=out_color) \
                            .grid(row=oi, column=5)  # availability to 12
                        oi += 1
                    # if there is no match, then just printe the name.
            if match == "miss":
                Label(F, text=line[1], anchor=W, width=21, relief=RIDGE, bg=in_color) \
                    .grid(row=oi, column=0)  # name
                for i in range(5):
                    if i < 3:
                        color = in_color
                    else:
                        color = out_color
                    Label(F, text="", width=6, relief=RIDGE, bg=color) \
                        .grid(row=oi, column=i + 1)  # generate blank cells
                oi += 1
        oi += 1
        Label(F, text="", height=2, bg="white").grid(row=oi, column=0)
        Label(F, text=format(aval_10_total, '.2f'), justify=LEFT, width=6, relief=RIDGE, bg=out_color) \
            .grid(row=oi, column=4)  # availability to 10 total
        Label(F, text=format(aval_12_total, '.2f'), justify=LEFT, width=6, relief=RIDGE, bg=out_color) \
            .grid(row=oi, column=5)  # availability to 12 total
        oi += 2
        # auxiliary assistance xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
        Label(F, text="auxiliary assistance", justify=LEFT, font=('Helvetica 10 bold'), bg="white") \
            .grid(sticky=W, row=oi, column=0, columnspan=9)
        oi += 1
        Label(F, text="Availability to: ", bg="white").grid(row=oi, column=4, columnspan=3, sticky=W)
        oi += 1
        Label(F, text="Carrier", bg="white").grid(row=oi, column=0, sticky=W)
        Label(F, text="note", bg="white").grid(row=oi, column=1)
        Label(F, text="5200", bg="white").grid(row=oi, column=2)
        Label(F, text="RS", bg="white").grid(row=oi, column=3)
        Label(F, text="to 10", bg="white").grid(row=oi, column=4)
        Label(F, text="to 11.5", bg="white").grid(row=oi, column=5)
        oi += 1
        aval_10_total = 0  # initialize variables for availability totals.
        aval_115_total = 0
        for line in dl_aux:
            match = "miss"
            for each in r_rings:
                if each[0] == str(day) and each[1] == line[1]:  # if the rings record is a match
                    match = "hit"
                    if match == "hit":
                        # find 10 hour availability pending code status
                        if each[4] == "light" or each[4] == "sch chg" or each[4] == "excused":
                            aval_10 = 0.00
                        elif each[4] == "no call":
                            aval_10 = 10.00
                        elif each[4] == "sick" or each[4] == "annual":
                            aval_10 = 0.00
                        elif each[2].strip() == "":
                            aval_10 = 0.00
                        else:
                            aval_10 = max(10 - float(each[2]), 0)
                        if aval_10 <= float(av_tol): aval_10 = 0  # adjust sum for tolerance
                        aval_10_total += aval_10  # add to availability total
                        # find 11.5 hour availability pending code status
                        if each[4] == "light" or each[4] == "sch chg" or each[4] == "excused":
                            aval_115 = 0.00
                        elif each[4] == "no call":
                            aval_115 = 12.00
                        elif each[4] == "sick" or each[4] == "annual":
                            aval_115 = 0.00
                        elif each[2].strip() == "":
                            aval_115 = 0.00
                        else:
                            aval_115 = max(12 - float(each[2]), 0)
                        if aval_115 <= float(av_tol): aval_115 = 0  # adjust sum for tolerance
                        aval_115_total += aval_115  # add to availability total
                        # output to the gui
                        Label(F, text=each[1], anchor=W, width=21, relief=RIDGE, bg=in_color) \
                            .grid(row=oi, column=0)  # name
                        if each[4] == "none":
                            code = ""
                        else:
                            code = each[4]
                        Label(F, text=code, justify=LEFT, width=6, relief=RIDGE, bg=in_color) \
                            .grid(row=oi, column=1)  # code
                        if each[2] == "" or each[2] == " ":  # handle empty 5200 strings
                            t_hrs = ""
                        else:
                            t_hrs = format(float(each[2]), '.2f')
                        Label(F, text=t_hrs, justify=LEFT, width=6, relief=RIDGE, bg=in_color) \
                            .grid(row=oi, column=2)  # 5200
                        if each[3] == "" or each[3] == " ":  # handle empty RS strings
                            rs = ""
                        else:
                            rs = format(float(each[3]), '.2f')
                        Label(F, text=rs, justify=LEFT, width=6, relief=RIDGE, bg=in_color) \
                            .grid(row=oi, column=3)  # return to station
                        Label(F, text=format(float(aval_10), '.2f'), justify=LEFT, width=6, relief=RIDGE, bg=out_color) \
                            .grid(row=oi, column=4)  # availability to 10
                        Label(F, text=format(float(aval_115), '.2f'), justify=LEFT, width=6, relief=RIDGE, bg=out_color) \
                            .grid(row=oi, column=5)  # availability to 12
                        oi += 1
            # if there is no match, then just printe the name.
            if match == "miss":
                Label(F, text=line[1], anchor=W, width=21, relief=RIDGE, bg=in_color) \
                    .grid(row=oi, column=0)  # name
                for i in range(5):
                    if i < 3:
                        color = in_color
                    else:
                        color = out_color
                    Label(F, text="", width=6, relief=RIDGE, bg=color) \
                        .grid(row=oi, column=i + 1)  # generate blank cells
            oi += 1
        oi += 1
        Label(F, text="", height=2, bg="white").grid(row=oi, column=0)
        Label(F, text=format(aval_10_total, '.2f'), justify=LEFT, width=6, relief=RIDGE, bg=out_color) \
            .grid(row=oi, column=4)  # availability to 10 total
        Label(F, text=format(aval_115_total, '.2f'), justify=LEFT, width=6, relief=RIDGE, bg=out_color) \
            .grid(row=oi, column=5)  # availability to 11.5 total
        oi += 2
        t += 1  # t increaments tabs
    root.mainloop()


def apply_rings(origin_frame, frame, carrier, total, RS, code,lv_type, lv_time, go_return):
    day = ("Saturday", "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday")
    days = (sat_mm, sun_mm, mon_mm, tue_mm, wed_mm, thr_mm, fri_mm)
    c = 0
    for d in days:  # check for bad inputs in moves
        x = len(d)
        for i in range(x):
            if triad_col_finder(i) == 0:  # find the first of the triad
                if isfloat(d[i].get()) == False or isfloat(d[i + 1].get()) == False:
                    if d[i].get().strip() == "" and d[i + 1].get().strip() == "":
                        continue
                    text = "You must enter a numeric value on moves for {}.".format(day[c])
                    messagebox.showerror("Move entry error", text, parent=frame)
                    return
                if float(d[i].get()) == float(d[i + 1].get()):  # if earlier greater than later
                    text = "The earlier value can not be greater equal to the later value on moves for {}.".format(
                        day[c])
                    messagebox.showerror("Move entry error", text, parent=frame)
                    return
                if float(d[i].get()) > float(d[i + 1].get()):  # if earlier greater than later
                    text = "The earlier value can not be greater than the later value on moves for {}.".format(day[c])
                    messagebox.showerror("Move entry error", text, parent=frame)
                    return
                if float(d[i].get()) > 24 or float(d[i + 1].get()) > 24:
                    text = "Values greater than 24 are not accepted on moves for {}.".format(day[c])
                    messagebox.showerror("Move entry error", text, parent=frame)
                    return
                if float(d[i].get()) <= 0 or float(d[i + 1].get()) < 0:
                    text = "Values less than 0 are not accepted on moves for {}.".format(day[c])
                    messagebox.showerror("Move entry error", text, parent=frame)
                    return
            if triad_col_finder(i) == 2:  # find the third of the triad
                if int(d[i].get().isnumeric()) == False:
                    if d[i].get().strip() == "":
                        continue
                    text = "You must enter a numeric value on route for {}.".format(day[c])
                    messagebox.showerror("Move entry error", text, parent=frame)
                    return
                if d[i].get() != "":  # if the route field is not blank
                    if len(d[i].get()) != 4:  # it must contain four digits.
                        text = "The route number for {} must be four digits long.".format(day[c])
                        messagebox.showerror("Move entry error", text, parent=frame)
                        return
        c += 1
    c = -1
    for t in total:  # check for bad inputs in 5200 fields
        c += 1
        if isfloat(t.get()) == False:
            if t.get().strip() == "":
                continue
            text = "You must enter a numeric value in 5200 for {}.".format(day[c])
            messagebox.showerror("Move entry error", text, parent=frame)
            return
        if float(t.get()) > 24:
            text = "Values greater than 24 are not accepted in 5200 for {}.".format(day[c])
            messagebox.showerror("Move entry error", text, parent=frame)
            return
        if float(t.get()) <= 0:
            text = "Values less than or equal to 0 are not accepted in 5200 for {}.".format(day[c])
            messagebox.showerror("Move entry error", text, parent=frame)
            return
    ttotal = []
    for t in total:
        t = str(t.get()).strip()
        if isfloat(t) == TRUE:
            ttotal.append(format(float(str(t)), '.2f'))
        else:
            ttotal.append(str(t))
    c = -1
    for r in RS:  # check for bad inputs in RS fields
        c += 1
        if isfloat(r.get()) == False:
            if r.get().strip() == "":
                continue
            text = "You must enter a numeric value in RS for {}.".format(day[c])
            messagebox.showerror("Move entry error", text, parent=frame)
            return
        if float(r.get()) > 24:
            text = "Values greater than 24 are not accepted in RS for {}.".format(day[c])
            messagebox.showerror("Move entry error", text, parent=frame)
            return
        if float(r.get()) < 0:
            text = "Values less than 0 are not accepted in RS for {}.".format(day[c])
            messagebox.showerror("Move entry error", text, parent=frame)
            return
    rRS = []
    for r in RS:
        r = str(r.get()).strip()
        if isfloat(r) == TRUE:
            rRS.append(format(float(str(r)), '.2f'))
        else:
            rRS.append(str(r))
    # check for bad inputs in lv_time fields
    c = -1
    for t in lv_time:
        c += 1
        if isfloat(t.get()) == False:
            if t.get().strip() == "":
                continue
            text = "You must enter a numeric value for leave times {}.".format(day[c])
            messagebox.showerror("5200 entry error", text, parent=frame)
            return
        if float(t.get()) > 8:
            text = "Values greater than 8 are not accepted for leave times for {}.".format(day[c])
            messagebox.showerror("5200 entry error", text, parent=frame)
            return
        # if float(t.get()) <= 0:
        #     text = "Values less than or equal to 0 are not accepted for leave time for {}.".format(day[c])
        #     messagebox.showerror("5200 entry error", text, parent=frame)
        #     return
    llv_time = [] # create new array to keep formated leave times
    for t in lv_time:
        t = str(t.get()).strip()
        if isfloat(t) == TRUE: # if the leave time can be a float
            if float(t)<= 0: # if the leave time is less than or equal to zero
                llv_time.append(str("")) # insert a blank in the array
            else: # if the leave time can be a float
                llv_time.append(format(float(str(t)), '.2f')) # format it as a float with 2 decimal places
        else:
            llv_time.append(str(t)) # otherwise input the string as it appears

    dates = []
    if g_range == "week": dates = g_date
    if g_range == "day": dates.append(d_date)
    if g_range == "week":
        sql = "SELECT * FROM rings3 WHERE carrier_name = '%s' and rings_date BETWEEN '%s' AND '%s'" \
              % (carrier[1], dates[0], dates[6])
    if g_range == "day":
        sql = "SELECT * FROM rings3 WHERE carrier_name = '%s' and rings_date = '%s'" \
              % (carrier[1], d_date)
    results = inquire(sql)
    d_sat_mm = []  # format moves for database
    d_sun_mm = []
    d_mon_mm = []
    d_tue_mm = []
    d_wed_mm = []
    d_thr_mm = []
    d_fri_mm = []
    d_mm = [d_sat_mm, d_sun_mm, d_mon_mm, d_tue_mm, d_wed_mm, d_thr_mm, d_fri_mm]
    all_moves = []
    # inserts moves into a daily list/ formats moves to float
    i = 0
    field1 = ""
    field2 = ""
    for d in days:
        ii = 0
        for each in d:
            if triad_col_finder(ii) == 0:  # find the first of the triad
                field1 = each.get().strip()
            if triad_col_finder(ii) == 1:  # find the seoond of the triad
                field2 = each.get().strip()
            if triad_col_finder(ii) == 2:
                # only write where MV fields are filled in
                if field1 != "":
                    d_mm[i].append(format(float(field1), '.2f'))
                    d_mm[i].append(format(float(field2), '.2f'))
                    d_mm[i].append(format(each.get()))

                field1 = ""
                field2 = ""
            ii += 1
        i += 1
    # remove the quotes around the items in the list
    for i in range(len(d_mm)):
        x = ','.join(d_mm[i])
        if x.replace(',', '') == "":
            x = ""
        all_moves.append(x)
    updates = []  # sort rings and moves and execute sql
    for i in range(len(dates)):
        for each in results:
            if str(dates[i]) == each[0]:
                updates.append(i)
                sql = "UPDATE rings3 SET total='%s',rs='%s',code='%s',moves='%s',leave_type = '%s',leave_time = '%s'" \
                      "WHERE rings_date = '%s' and carrier_name = '%s'" \
                      % (ttotal[i], rRS[i], code[i].get(),
                         all_moves[i],lv_type[i].get(), llv_time[i], dates[i], carrier[1])
                commit(sql)
    if g_range == "week":
        inserts = [0, 1, 2, 3, 4, 5, 6, ]  # seven inserts for a week and one for a day
    else:
        inserts = [0]
    for num in updates:
        if num in inserts:
            inserts.remove(num)
    for i in inserts:  # for each day, insert the information
        sql = "INSERT INTO rings3 (rings_date, carrier_name, total, rs, code, moves, leave_type, leave_time )" \
              "VALUES('%s','%s','%s','%s','%s','%s','%s','%s') " \
              % (dates[i], carrier[1], ttotal[i], rRS[i], code[i].get(), all_moves[i], lv_type[i].get(), llv_time[i])
        commit(sql)
    sql = "DELETE FROM rings3 WHERE total='%s' and code='%s' and leave_time ='%s'" % ("", 'none', 'none')

    commit(sql)
    # destroy the old rings entry window
    if go_return == "no_return":
        frame.destroy()
    else:
        frame.destroy()
        rings2(carrier, origin_frame)


def triad_row_finder(index):
    if index % 3 == 0:
        row = index / 3
    elif (index - 1) % 3 == 0:
        row = (index - 1) / 3
    elif (index - 2) % 3 == 0:
        row = (index - 2) / 3
    return int(row)


def triad_col_finder(index):
    if index % 3 == 0:  # first column
        col = 0
    elif (index - 1) % 3 == 0:  # second column
        col = 1
    elif (index - 2) % 3 == 0:  # third column
        col = 2
    return int(col)


def rings_triad_placement(iteration):
    if iteration % 3 == 0:
        place = 2
    elif (iteration - 1) % 3 == 0:
        place = 0
    elif (iteration - 2) % 3 == 0:
        place = 1
    return place


def new_entry(self, day, moves):  # creates new entry fields for moves
    if day == "sat":
        mm = sat_mm  # find the day in question and use the correlating  array
    elif day == "sun":
        mm = sun_mm
    elif day == "mon":
        mm = mon_mm
    elif day == "tue":
        mm = tue_mm
    elif day == "wed":
        mm = wed_mm
    elif day == "thr":
        mm = thr_mm
    elif day == "fri":
        mm = fri_mm
    # what to do depending on the moves
    if moves == 0:  # if there are no moves sent to the function
        mm.append(StringVar(self))  # create first entry field for new entries
        Entry(self, width=8, textvariable=mm[len(mm) - 1]) \
            .grid(row=triad_row_finder(len(mm) - 1) + 2, column=triad_col_finder(len(mm) - 1) + 2)  # route
        mm.append(StringVar(self))  # create second entry field for new entries
        Entry(self, width=8, textvariable=mm[len(mm) - 1]) \
            .grid(row=triad_row_finder(len(mm) - 1) + 2, column=triad_col_finder(len(mm) - 1) + 2)  # move off
        mm.append(StringVar(self))  # create second entry field for new entries
        Entry(self, width=8, textvariable=mm[len(mm) - 1]) \
            .grid(row=triad_row_finder(len(mm) - 1) + 2, column=triad_col_finder(len(mm) - 1) + 2)  # move on
    else:  # if there are moves which need to be set
        moves = moves.split(",")
        iterations = len(moves)
        for i in range(int(iterations)):
            mm.append(StringVar(self))  # create entry field for moves from database
            mm[i].set(moves[i])
            Entry(self, width=8, textvariable=mm[i]) \
                .grid(row=triad_row_finder(i) + 2, column=triad_col_finder(i) + 2)


def rings2(carrier, origin_frame):
    root = Tk()
    root.title("KLUSTERBOX")
    root.geometry("%dx%d+%d+%d" % (origin_frame.winfo_width(), origin_frame.winfo_height(),
                                   origin_frame.winfo_rootx(), origin_frame.winfo_rooty() - 30))
    switchF2 = Frame(root)
    switchF2.pack(fill=BOTH, side=LEFT)
    C1 = Canvas(switchF2)
    C1.pack(fill=BOTH, side=BOTTOM)
    # apply and close buttons
    Button(C1, text="Submit", width=10, bg="light yellow", anchor="w",
           command=lambda: [apply_rings(origin_frame, root, carrier, total, RS, code,lv_type, lv_time, "no_return")])\
        .pack(side=LEFT)
    Button(C1, text="Apply", width=10, bg="light yellow", anchor="w",
           command=lambda: [apply_rings(origin_frame, root, carrier, total, RS, code,lv_type, lv_time, "do_return")])\
        .pack(side=LEFT)
    Button(C1, text="Go Back", width=10, bg="light yellow", anchor="w",
           command=lambda: root.destroy()).pack(side=LEFT)
    # define scrollbar and canvas
    S = Scrollbar(switchF2)
    C = Canvas(switchF2, width=1600)
    # link up the canvas and scrollbar
    S.pack(side=RIGHT, fill=BOTH)
    C.pack(side=LEFT, fill=BOTH, pady=10, padx=20)
    S.configure(command=C.yview, orient="vertical")
    C.configure(yscrollcommand=S.set)
    if sys.platform == "win32":
        C.bind_all('<MouseWheel>', lambda event: C.yview_scroll(int(-1 * (event.delta / 120)), "units"))
    elif sys.platform == "linux":
        C.bind_all('<Button-4>', lambda event: C.yview('scroll', -1, 'units'))
        C.bind_all('<Button-5>', lambda event: C.yview('scroll', 1, 'units'))
    # create the frame inside the canvas
    F = Frame(C, height=1000)
    C.create_window((0, 0), window=F, anchor=NW)
    del sat_mm[:]  # initialize the daily arrays for moves...
    del sun_mm[:]
    del mon_mm[:]
    del tue_mm[:]
    del wed_mm[:]
    del thr_mm[:]
    del fri_mm[:]
    date = datetime(int(gs_year), int(gs_mo), int(gs_day))
    dates = []
    list_carrier = []
    if g_range == "week":
        start_invest = date
        end_invest = date + timedelta(days=6)
        for i in range(7):
            dates.append(date)
            date += timedelta(days=1)
    else:
        end_invest = date
        dates.append(date)
    sql = "SELECT * FROM" \
          " carriers WHERE carrier_name = '%s' and effective_date <= '%s' " \
          "ORDER BY effective_date" % (carrier[1], end_invest)
    results = inquire(sql)
    day = ("sat", "sun", "mon", "tue", "wed", "thr", "fri")
    frame = ["F0", "F1", "F2", "F3", "F4", "F5", "F6"]
    color = ["red", "light blue", "yellow", "green", "brown", "gold", "purple", "grey", "light grey"]
    nolist_codes = ("none", "ns day")
    ot_aux_codes = ("none", "no call", "light", "sch chg", "annual", "sick", "excused")
    lv_options = ("none","annual","sick","holiday","other")
    option_menu = ["om0", "om1", "om2", "om3", "om4", "om5", "om6"]
    lv_option_menu = ["lom0", "lom1", "lom2", "lom3", "lom4", "lom5", "lom6"]
    total_widget = ["tw0", "tw1", "tw2", "tw3", "tw4", "tw5", "tw6"]
    total = []
    RS = []
    lv_type = []
    lv_time = []
    code = []
    if g_range == "week":  # Get carrier list information
        in_range = []
        candidates = []
        station_anchor = "no"
        sat_rec = "false"
        in_station = "false"
        is_winner = "false"
        # create list carrier array: most recent records of carriers currently in the station for any day of the week
        for r in results:
            if str(start_invest) <= r[0] <= str(end_invest) and r[5] == g_station:
                station_anchor = "yes"
                if str(start_invest) == r[0]:
                    sat_rec = "true"  # hit on saturday is true
                if r[5] == g_station:
                    in_station = "true"  # hit if in station at any time
        for r in results:
            if (str(start_invest) <= r[0]):
                in_range.append(r)
            if (r[0] < str(start_invest) and sat_rec == "false"):
                candidates.append(r)
        if candidates and sat_rec == "false":
            winner = max(candidates, key=itemgetter(0))
            if winner[5] == g_station or station_anchor == "yes":
                list_carrier.append(winner)
                if len(in_range) > 0:
                    is_winner = "true"
        if len(in_range) > 0 and in_station == "true" or station_anchor == "yes" or is_winner == "true":
            for each in in_range:
                list_carrier.append(each)
        short_list = []  # create an array of candidates of possible valid records for each day
        daily_record = []
        for d in dates:
            del short_list[:]
            for l in list_carrier:
                if l[0] <= str(d): short_list.append(l)
            try:
                winner = max(short_list, key=itemgetter(0))
                daily_record.append(winner)
            except:
                no_record = (str(d), l[1], '', '', '', 'no record')
                daily_record.append(no_record)
    elif g_range == "day":
        daily_record = []
        candidates = []
        for record in results:
            candidates.append(record)
        if candidates:
            winner = max(candidates, key=itemgetter(0))
            if winner[5] == g_station:
                list_carrier.append(winner)
                daily_record.append(winner)
    if g_range == "week":
        sql = "SELECT * FROM rings3 WHERE carrier_name = '%s' and rings_date BETWEEN '%s' AND '%s'" \
              % (carrier[1], start_invest, end_invest)
    else:
        sql = "SELECT * FROM rings3 WHERE carrier_name = '%s' and rings_date = '%s'" \
              % (carrier[1], end_invest)
    r_rings = inquire(sql)
    frame_i = 0  # counter for the frame
    header_frame = Frame(F, width=500)  # header  frame
    header_frame.grid(row=frame_i, padx=5, sticky="w")
    # Header at top of window: name
    Label(header_frame, text="carrier name: ", fg="Grey", font="bold").grid(row=0, column=0, sticky="w")
    Label(header_frame, text="{}".format(carrier[1]), font="bold").grid(row=0, column=1, sticky="w")
    Label(header_frame, text="list status: {}".format(carrier[2])).grid(row=1, sticky="w", columnspan=2)
    if carrier[4] != "":
        Label(header_frame, text="route/s: {}".format(carrier[4])).grid(row=2, sticky="w", columnspan=2)
    frame_i += 2
    if g_range == "week":
        i_range = 7  # loop 7 times for week or once for day
    else:
        i_range = 1
    for i in range(i_range):
        now_total = ""
        now_rs = ""
        now_code = "none"
        now_moves = ""
        now_lv_type = "none"
        now_lv_time = ""
        for ring in r_rings:
            if ring[0] == str(dates[i]):  # if the dates match set the corresponding rings
                now_total = ring[2]
                now_rs = ring[3]
                now_code = ring[4]
                now_moves = ring[5]
                if ring[6]=='': # format the leave type
                    now_lv_type = "none"
                else:
                    now_lv_type = ring[6]
                if str(ring[7])=='None': # format the leave time to be blank or a float
                    now_lv_time = ""
                elif isfloat(ring[7]) == TRUE and float(ring[7])==0: # if the leave time can be a float
                    now_lv_time = ""
                else:
                    now_lv_time = ring[7]
        grid_i = 0  # counter for the grid within the frame
        frame[i] = Frame(F, width=500)
        frame[i].grid(row=frame_i, padx=5, sticky="w")
        # Display the day and date
        if ns_code[carrier[3]] == dates[i].strftime("%a"):
            Label(frame[i], text="{} NS DAY".format(dates[i].strftime("%a %b %d, %Y")), fg="red") \
                .grid(row=grid_i, column=0, columnspan=5, sticky="w")
        else:
            Label(frame[i], text=dates[i].strftime("%a %b %d, %Y"), fg="blue") \
                .grid(row=grid_i, column=0, columnspan=5, sticky="w")
        grid_i += 1
        if daily_record[i][5] == g_station:
            Label(frame[i], text="5200", fg=color[7]).grid(row=grid_i, column=0)  # Display all labels
            Label(frame[i], text="RS", fg=color[7]).grid(row=grid_i, column=1)
            if daily_record[i][2] == "wal" or daily_record[i][2] == "nl":
                Label(frame[i], text="MV off", fg=color[7]).grid(row=grid_i, column=2)
                Label(frame[i], text="MV on", fg=color[7]).grid(row=grid_i, column=3)
                Label(frame[i], text="Route", fg=color[7]).grid(row=grid_i, column=4)
                Label(frame[i], text="code", fg=color[7]).grid(row=grid_i, column=6)
                Label(frame[i], text="LV type", fg=color[7]).grid(row=grid_i, column=7)
                Label(frame[i], text="LV time", fg=color[7]).grid(row=grid_i, column=8)
            else:
                Label(frame[i], text="code", fg=color[7]).grid(row=grid_i, column=3)
                Label(frame[i], text="LV type", fg=color[7]).grid(row=grid_i, column=4)
                Label(frame[i], text="LV time", fg=color[7]).grid(row=grid_i, column=5)

            grid_i += 1
            # Display the entry widgets
            total.append(StringVar(frame[i]))  # 5200 entry widget
            total_widget[i] = Entry(frame[i], width=8, textvariable=total[i])
            total_widget[i].grid(row=grid_i, column=0)
            total[i].set(now_total)  # set the starting value for total
            RS.append(StringVar(frame[i]))  # RS entry widget
            Entry(frame[i], width=8, textvariable=RS[i]).grid(row=grid_i, column=1)
            RS[i].set(now_rs)  # set the starting value for RS
            if daily_record[i][2] == "wal" or daily_record[i][2] == "nl":
                if now_moves.strip() != "":
                    new_entry(frame[i], day[i], now_moves)  # MOVES on and off entry widgets
                else:
                    new_entry(frame[i], day[i], 0)
                Button(frame[i], text="more moves", command=lambda x=i: new_entry(frame[x], day[x], 0)) \
                    .grid(row=grid_i, column=5)
            code.append(StringVar(frame[i]))  # code entry widget
            if daily_record[i][2] == "wal" or daily_record[i][2] == "nl":
                option_menu[i] = OptionMenu(frame[i], code[i], *nolist_codes)
            else:
                option_menu[i] = OptionMenu(frame[i], code[i], *ot_aux_codes)
            code[i].set(now_code)
            option_menu[i].configure(width=7)
            lv_type.append(StringVar(frame[i]))  # leave type entry widget
            lv_option_menu[i] = OptionMenu(frame[i],lv_type[i], *lv_options)
            lv_option_menu[i].configure(width=7)
            lv_time.append(StringVar(frame[i]))  # leave time entry widget
            lv_type[i].set(now_lv_type)  # set the starting value for leave type
            lv_time[i].set(now_lv_time)  # set the starting value for leave type
            # put code widgets on the grid
            if daily_record[i][2] == "wal" or daily_record[i][2] == "nl":
                option_menu[i].grid(row=grid_i, column=6) # code widget
                lv_option_menu[i].grid(row=grid_i, column=7)  # leave type widget
                Entry(frame[i], width=8, textvariable=lv_time[i]).grid(row=grid_i, column=8) # leave time widget
            else:
                option_menu[i].grid(row=grid_i, column=3) # code widget
                lv_option_menu[i].grid(row=grid_i, column=4)  # leave type widget
                Entry(frame[i], width=8, textvariable=lv_time[i]).grid(row=grid_i, column=5) # leave time widget

        else:
            total.append(StringVar(frame[i]))  # 5200 entry widget
            RS.append(StringVar(frame[i]))  # RS entry

            if daily_record[i][5] != "no record":  # display for records that are out of station
                Label(frame[i], text="out of station: {}".format(daily_record[i][5]), fg="white", bg="grey", width=55,
                      height=2, anchor="w") \
                    .grid(row=grid_i, column=0)
            else:  # display for when there is no record relevant for that day.
                Label(frame[i], text="no record", fg="white", bg="grey", width=55,
                      height=2, anchor="w") \
                    .grid(row=grid_i, column=0)
        frame_i += 1
    # total_widget[0].focus_set() # set the focus for the first 5200 widget.
    F7 = Frame(F)
    F7.grid(row=frame_i)
    Label(F7, height=50).grid(row=1, column=0)  # extra white space on bottom of form to facilitate moves
    root.update()
    C.config(scrollregion=C.bbox("all"))
    mainloop()


def apply_update_carrier(year, month, day, name, ls, ns, route, station, rowid, self):
    if year.get() > 9999:
        messagebox.showerror("Year Input Error", "Year must be between 1 and 9999", parent=self)
        return
    if year.get() < 1:
        messagebox.showerror("Year Input Error", "Year must be between 1 and 9999", parent=self)
        return
    try:
        date = datetime(year.get(), month.get(), day.get())
    except:
        messagebox.showerror("Invalid Date", "Date entered is not valid", parent=self)
        return
    route_list = []
    route_list = route.get().split("/")
    if len(route.get()) > 24:
        messagebox.showerror("Route number input error", "There can be no more than five routes per carrier "
                                                         "(for T6 carriers).\n Routes numbers can be no more than four digits long.\n"
                                                         "If there are multiple routes, route numbers must be separated by "
                                                         "the \'/\' character. For example: 1001/1015/1024/1036/1072. Do not use "
                                                         "commas or empty spaces", parent=self)
        return
    for item in route_list:
        item = item.strip()
        if item != "":
            if len(item) != 4:
                messagebox.showerror("Route number input error",
                                     'Routes numbers must be four digits long.\n'
                                     'If there are multiple routes, route numbers must be separated by '
                                     'the \'/\' character. For example: 1001/1015/1024/1036/0972. Do not use '
                                     'commas or empty spaces', parent=self)
                return
        if item.isdigit() == FALSE and item != "":
            messagebox.showerror("Route number input error", "Route numbers must be numbers and can not contain "
                                                             "letters", parent=self)
            return
    route_input = route.get()
    if route_input == "0000":
        route_input = ""
    sql = "UPDATE carriers SET effective_date='%s',list_status='%s',ns_day='%s',route_s='%s',station='%s' " \
          "WHERE rowid = '%s'" % \
          (date, ls.get(), ns.get(), route_input, station.get(), rowid)
    commit(sql)
    self.destroy()
    edit_carrier(name)


def delete_carrier(name):
    sql = "DELETE FROM carriers WHERE rowid = '%s'" % name[6]
    commit(sql)
    sql = "SELECT carrier_name FROM carriers WHERE carrier_name = '%s'" % name[1]
    results = inquire(sql)
    if len(results) > 0:
        edit_carrier(name[1])
    else:
        main_frame()


def apply(year, month, day, c_name, ls, ns, route, station, self):
    if year.get() > 9999:
        messagebox.showerror("Year Input Error", "Year must be between 1 and 9999", parent=self)
        return
    if year.get() < 1:
        messagebox.showerror("Year Input Error", "Year must be between 1 and 9999", parent=self)
        return

    try:
        date = datetime(year.get(), month.get(), day.get())
    except:
        messagebox.showerror("Invalid Date", "Date entered is not valid", parent=self)
        return
    carrier = c_name.strip().lower()
    if len(carrier) > 30:
        messagebox.showerror("Name input error", "Names must not exceed 30 characters.", parent=self)
        return
    if len(carrier) < 1:
        messagebox.showerror("Name input error", "You must enter a name.", parent=self)
        return
    apply_2(date, carrier, ls, ns, route, station, self)


def apply_2(date, carrier, ls, ns, route, station, self):
    route_list = route.get().split("/")
    if len(route.get()) > 24:
        messagebox.showerror("Route number input error", "There can be no more than five routes per carrier "
                                                         "(for T6 carriers).\n Routes numbers can be no more than four digits long.\n"
                                                         "If there are multiple routes, route numbers must be separated by "
                                                         "the \'/\' character. For example: 1001/1015/1024/1036/0972. Do not use "
                                                         "commas or empty spaces", parent=self)
        return
    for item in route_list:
        item = item.strip()
        if item != "":
            if len(item) != 4:
                messagebox.showerror("Route number input error", 'Routes numbers must be four digits long.\n'
                                                                 'If there are multiple routes, route numbers must be separated by '
                                                                 'the \'/\' character. For example: 1001/1015/1024/1036/1072. Do not use '
                                                                 'commas or empty spaces', parent=self)
                return
        if item.isdigit() == FALSE and item != "":
            messagebox.showerror("Route number input error", "Route numbers must be numbers and can not contain "
                                                             "letters", parent=self)
            return
    # find all matches for date and name
    route_input = route.get()
    if route_input == "0000":
        route_input = ""
    sql = "SELECT effective_date, carrier_name,list_status, ns_day,route_s, station, rowid FROM carriers " \
          "WHERE carrier_name = '%s' and effective_date = '%s' ORDER BY effective_date" % (carrier, date)
    results = inquire(sql)
    if len(results) == 0:
        sql = "INSERT INTO carriers (effective_date,carrier_name,list_status,ns_day,route_s,station)" \
              " VALUES('%s','%s','%s','%s','%s','%s')" \
              % (date, carrier, ls.get(), ns.get(), route_input, station.get())
        commit(sql)
    elif len(results) == 1:
        sql = "UPDATE carriers SET list_status='%s',ns_day='%s',route_s='%s',station='%s' " \
              "WHERE effective_date = '%s' and carrier_name = '%s'" % \
              (ls.get(), ns.get(), route_input, station.get(), date, carrier)
        commit(sql)
    elif len(results) > 1:
        sql = "DELETE FROM carriers WHERE effective_date ='%s' and carrier_name = '%s'" % (date, carrier)
        commit(sql)
        sql = "INSERT INTO carriers (effective_date,carrier_name,list_status,ns_day,route_s,station)" \
              " VALUES('%s','%s','%s','%s','%s','%s')" \
              % (date, carrier, ls.get(), ns.get(), route_input, station.get())
        commit(sql)


def name_change(name, c_name, self):
    c_name = c_name.get().strip().lower()
    if messagebox.askokcancel("Name Change", "This will change the name {} to {} in all records. "
                                             "Are you sure?".format(name, c_name)):
        if len(c_name) > 42:
            messagebox.showerror("Name input error", "Names must not exceed 42 characters.", parent=self)
            return
        if len(c_name) < 1:
            messagebox.showerror("Name input error", "You must enter a name.", parent=self)
            return
        sql = "SELECT kb_name FROM name_index WHERE kb_name = '%s'" % c_name
        result = inquire(sql)
        if result:
            messagebox.showerror("Name input error", "This name is already being used for another carrier.",
                                 parent=self)
            return
        sql = "SELECT carrier_name FROM carriers WHERE carrier_name = '%s'" % c_name
        result = inquire(sql)
        if result:
            messagebox.showerror("Name input error", "This name is already being used for another carrier.",
                                 parent=self)
            return
        sql = "UPDATE carriers SET carrier_name = '%s' WHERE carrier_name = '%s'" % (c_name, name)
        commit(sql)
        sql = "UPDATE rings3 SET carrier_name = '%s' WHERE carrier_name = '%s'" % (c_name, name)
        commit(sql)
        sql = "SELECT kb_name FROM name_index WHERE kb_name = '%s'" % name
        result = inquire(sql)
        if result:
            sql = "UPDATE name_index SET kb_name = '%s' WHERE kb_name = '%s'" % (c_name, name)
            commit(sql)
        self.destroy()
        main_frame()


def update_carrier(a):
    switchF4 = Frame(root)
    switchF4.pack(fill=BOTH, side=LEFT)
    C1 = Canvas(switchF4)
    C1.pack(fill=BOTH, side=BOTTOM)
    # define scrollbar and canvas
    S = Scrollbar(switchF4)
    C = Canvas(switchF4, width=1600)
    # link up the canvas and scrollbar
    S.pack(side=RIGHT, fill=BOTH)
    C.pack(side=LEFT, fill=BOTH, pady=10, padx=20)
    S.configure(command=C.yview, orient="vertical")
    C.configure(yscrollcommand=S.set)
    if sys.platform == "win32":
        C.bind_all('<MouseWheel>', lambda event: C.yview_scroll(int(-1 * (event.delta / 120)), "units"))
    elif sys.platform == "linux":
        C.bind_all('<Button-4>', lambda event: C.yview('scroll', -1, 'units'))
        C.bind_all('<Button-5>', lambda event: C.yview('scroll', 1, 'units'))
    # create the frame inside the canvas
    F = Frame(C)
    C.create_window((0, 0), window=F, anchor=NW)
    # page title
    title_F = Frame(F)
    Label(title_F, text="Update Carrier Information", font="bold").grid(row=0, column=0, columnspan=4)
    title_F.grid(row=0, sticky=W)  # put frame on grid
    # date
    date_frame = Frame(F)  # define frame
    year = IntVar(date_frame)  # define variables for date
    month = IntVar(date_frame)
    day = IntVar(date_frame)
    # pre set values for date
    month.set(int(a[0][5:7]))
    day.set(int(a[0][8:10]))
    year.set(int(a[0][:4]))
    Label(date_frame, text="date (month/day/year):").grid(row=0, column=0, sticky=W, columnspan=3)  # date label
    OptionMenu(date_frame, month, "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12") \
        .grid(row=1, column=0, sticky=W)
    OptionMenu(date_frame, day, "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14",
               "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28", "29",
               "30", "31").grid(row=1, column=1, sticky=W)
    Entry(date_frame, width=6, textvariable=year).grid(row=1, column=2, sticky=W)
    date_frame.grid(row=1, sticky=W)  # put frame on grid
    # carrier name
    name_frame = Frame(F, pady=2)
    name = StringVar(name_frame)
    name = a[1]  # name value if name is not changed
    Label(name_frame, width=15, text="carrier name: ", anchor="w").grid(row=0, column=0, sticky=W)
    Label(name_frame, text="{}".format(a[1].upper()), anchor="w", width=37).grid(row=1, column=0, sticky=W)
    name_frame.grid(row=2, sticky=W)
    # list status
    list_frame = Frame(F, bd=1, relief=RIDGE, pady=2)
    Label(list_frame, width=15, text="list status", anchor="w").grid(row=0, column=0, sticky=W)
    ls = StringVar(list_frame)
    ls.set(value=a[2])
    Radiobutton(list_frame, text="OTDL", variable=ls, value='otdl', justify=LEFT).grid(row=1, column=0, sticky=W)
    Radiobutton(list_frame, text="Work Assignment", variable=ls, value='wal', justify=LEFT).grid(row=1, column=1,
                                                                                                 sticky=W)
    Radiobutton(list_frame, text="No List", variable=ls, value='nl', justify=LEFT).grid(row=2, column=0, sticky=W)
    Radiobutton(list_frame, text="Auxiliary", variable=ls, value='aux', justify=LEFT).grid(row=2, column=1, sticky=W)
    list_frame.grid(row=3, sticky=W)
    # set non scheduled day
    ns_frame = Frame(F, pady=2)
    Label(ns_frame, width=15, text="non scheduled day", anchor="w").grid(row=0, column=0, sticky=W)
    ns = StringVar(ns_frame)
    ns.set(a[3])
    Radiobutton(ns_frame, text="{}:   yellow".format(ns_code['yellow']), variable=ns, value="yellow", indicatoron=0,
                width=15, anchor="w",
                bg="grey", fg="white", selectcolor="yellow").grid(row=1, column=0)
    Radiobutton(ns_frame, text="{}:   blue".format(ns_code['blue']), variable=ns, value="blue", indicatoron=0, width=15,
                anchor="w",
                bg="grey", fg="white", selectcolor="blue").grid(row=2, column=0)
    Radiobutton(ns_frame, text="{}:   green".format(ns_code['green']), variable=ns, value="green", indicatoron=0,
                width=15, anchor="w",
                bg="grey", fg="white", selectcolor="green").grid(row=3, column=0)
    Radiobutton(ns_frame, text="{}:   brown".format(ns_code['brown']), variable=ns, value="brown", indicatoron=0,
                width=15, anchor="w",
                bg="grey", fg="white", selectcolor="brown").grid(row=1, column=1)
    Radiobutton(ns_frame, text="{}:   red".format(ns_code['red']), variable=ns, value="red", indicatoron=0, width=15,
                anchor="w",
                bg="grey", fg="white", selectcolor="red").grid(row=2, column=1)
    Radiobutton(ns_frame, text="{}:   black".format(ns_code['black']), variable=ns, value="black", indicatoron=0,
                width=15, anchor="w",
                bg="grey", fg="white", selectcolor="black").grid(row=3, column=1)
    Radiobutton(ns_frame, text="none", variable=ns, value="none", indicatoron=0, width=15, anchor="w").grid(row=4,
                                                                                                            column=1)
    ns_frame.grid(row=4, sticky=W)
    # set route entry field
    route_frame = Frame(F, bd=1, relief=RIDGE, pady=2)
    Label(route_frame, text="route/s", width=15, anchor="w").grid(row=0, column=0, sticky=W)
    route = StringVar(route_frame)
    route.set(a[4])
    Entry(route_frame, width=37, textvariable=route).grid(row=1, column=0, sticky=W)
    route_frame.grid(row=5, sticky=W)
    # set station option menu
    station_frame = Frame(F, pady=2)
    Label(station_frame, text="station", width=10, anchor="w").grid(row=0, column=0, sticky=W)
    station = StringVar(station_frame)
    station.set(a[5])  # default value
    OptionMenu(station_frame, station, *list_of_stations).grid(row=0, column=1, sticky=W)
    station_frame.grid(row=6, sticky=W)
    # set rowid
    rowid = StringVar(F)
    rowid = a[6]
    root.update()
    C.config(scrollregion=C.bbox("all"))
    # apply and close buttons
    Button(C1, text="Apply", width=15, anchor="w",
           command=lambda: apply_update_carrier(year, month, day, name, ls, ns, route, station, rowid, switchF4)) \
        .pack(side=LEFT)
    Button(C1, text="Go Back", width=15, anchor="w",
           command=lambda: [switchF4.destroy(), main_frame()]).pack(side=LEFT)


def edit_carrier(e_name):
    sql = "SELECT effective_date, carrier_name,list_status, ns_day,route_s, station, rowid" \
          " FROM carriers WHERE carrier_name = '%s' ORDER BY effective_date DESC" % e_name
    results = inquire(sql)
    sql = "SELECT * FROM ns_configuration"
    ns_results = inquire(sql)
    ns_dict = {}  # build dictionary for ns days
    ns_color_dict = {}
    days = ("sat", "mon", "tue", "wed", "thu", "fri")
    for r in ns_results:  # build dictionary for rotating ns days
        ns_dict[r[0]] = r[2]
        ns_color_dict[r[0]] = r[1]  # build dictionary for ns fill colors
    for d in days:  # expand dictionary for fixed days
        ns_dict[d] = "fixed: " + d
        ns_color_dict[d] = "teal"
    ns_dict["none"] = "none"  # add "none" to dictionary
    ns_color_dict["none"] = "teal"
    switchF3 = Frame(root)
    switchF3.pack(fill=BOTH, side=LEFT)
    C1 = Canvas(switchF3)
    C1.pack(fill=BOTH, side=BOTTOM)
    # define scrollbar and canvas
    S = Scrollbar(switchF3)
    C = Canvas(switchF3, width=1600)
    # link up the canvas and scrollbar
    S.pack(side=RIGHT, fill=BOTH)
    C.pack(side=LEFT, fill=BOTH, pady=10, padx=20)
    S.configure(command=C.yview, orient="vertical")
    C.configure(yscrollcommand=S.set)
    if sys.platform == "win32":
        C.bind_all('<MouseWheel>', lambda event: C.yview_scroll(int(-1 * (event.delta / 120)), "units"))
    elif sys.platform == "linux":
        C.bind_all('<Button-4>', lambda event: C.yview('scroll', -1, 'units'))
        C.bind_all('<Button-5>', lambda event: C.yview('scroll', 1, 'units'))
    # create the frame inside the canvas
    F = Frame(C)
    C.create_window((0, 0), window=F, anchor=NW)
    # page title
    title_F = Frame(F)
    Label(title_F, text="Edit Carrier Information", font="bold").grid(row=0, column=0, columnspan=4)
    title_F.grid(row=0, sticky=W)  # put frame on grid
    # current date
    year = IntVar(F)
    month = IntVar(F)
    day = IntVar(F)
    # pre set values for date
    month.set(gs_mo)
    day.set(gs_day)
    year.set(gs_year)
    # define frame
    date_frame = Frame(F)
    Label(date_frame, text="date (month/day/year):") \
        .grid(row=0, column=0, sticky=W, columnspan=3)  # date label
    OptionMenu(date_frame, month, "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12") \
        .grid(row=1, column=0, sticky=W)  # option menu for month
    OptionMenu(date_frame, day, "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14",
               "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28", "29", "30", "31") \
        .grid(row=1, column=1, sticky=W)  # option menu for day
    Entry(date_frame, width=6, textvariable=year).grid(row=1, column=2, sticky=W)  # entry field for year
    date_frame.grid(row=1, sticky=W)  # put frame on grid
    # carrier name
    name_frame = Frame(F, pady=2)
    c_name = StringVar(name_frame)
    name = StringVar(name_frame)
    name = e_name  # name value if name is not changed
    c_name.set(e_name)  # name value for name changes
    Label(name_frame, text="carrier name: {}".format(e_name), anchor="w", width=37).grid(row=0, column=0, sticky=W)
    Entry(name_frame, width=37, textvariable=c_name).grid(row=1, column=0, sticky=W)
    Button(name_frame, width=7, text="update", command=lambda: name_change(name, c_name, switchF3)) \
        .grid(row=2, column=0, sticky=W)
    name_frame.grid(row=2, sticky=W)
    # list status
    list_frame = Frame(F, bd=1, relief=RIDGE, pady=2)
    Label(list_frame, width=15, text="list status", anchor="w").grid(row=0, column=0, sticky=W)
    ls = StringVar(list_frame)
    try:
        ls.set(value=results[0][2])
    except:
        switchF3.destroy(), main_frame()
    Radiobutton(list_frame, text="OTDL", variable=ls, value='otdl', justify=LEFT).grid(row=1, column=0, sticky=W)
    Radiobutton(list_frame, text="Work Assignment", variable=ls, value='wal', justify=LEFT).grid(row=1, column=1,
                                                                                                 sticky=W)
    Radiobutton(list_frame, text="No List", variable=ls, value='nl', justify=LEFT).grid(row=2, column=0, sticky=W)
    Radiobutton(list_frame, text="Auxiliary", variable=ls, value='aux', justify=LEFT).grid(row=2, column=1, sticky=W)
    list_frame.grid(row=3, sticky=W)
    # set non scheduled day
    ns_frame = Frame(F, pady=2)
    Label(ns_frame, width=15, text="non scheduled day", anchor="w").grid(row=0, column=0, sticky=W)
    ns = StringVar(ns_frame)
    ns.set(results[0][3])
    Radiobutton(ns_frame, text="{}:   {}".format(ns_code['yellow'], ns_results[0][2]), variable=ns, value="yellow",
                indicatoron=0, width=15, anchor="w",
                bg="grey", fg="white", selectcolor=ns_color_dict["yellow"]).grid(row=1, column=0)
    Radiobutton(ns_frame, text="{}:   {}".format(ns_code['blue'], ns_results[1][2]), variable=ns, value="blue",
                indicatoron=0, width=15, anchor="w",
                bg="grey", fg="white", selectcolor=ns_color_dict["blue"]).grid(row=1, column=1)
    Radiobutton(ns_frame, text="{}:   {}".format(ns_code['green'], ns_results[2][2]), variable=ns, value="green",
                indicatoron=0, width=15, anchor="w",
                bg="grey", fg="white", selectcolor=ns_color_dict["green"]).grid(row=2, column=0)
    Radiobutton(ns_frame, text="{}:   {}".format(ns_code['brown'], ns_results[3][2]), variable=ns, value="brown",
                indicatoron=0, width=15, anchor="w",
                bg="grey", fg="white", selectcolor=ns_color_dict["brown"]).grid(row=2, column=1)
    Radiobutton(ns_frame, text="{}:   {}".format(ns_code['red'], ns_results[4][2]), variable=ns, value="red",
                indicatoron=0, width=15, anchor="w",
                bg="grey", fg="white", selectcolor=ns_color_dict["red"]).grid(row=3, column=0)
    Radiobutton(ns_frame, text="{}:   {}".format(ns_code['black'], ns_results[5][2]), variable=ns, value="black",
                indicatoron=0, width=15, anchor="w",
                bg="grey", fg="white", selectcolor=ns_color_dict["black"]).grid(row=3, column=1)
    Label(ns_frame, text="Fixed:", anchor="w").grid(row=4, column=0, sticky="w")
    Radiobutton(ns_frame, text="none", variable=ns, value="none", indicatoron=0, width=15,
                bg="grey", fg="white", selectcolor=ns_color_dict["none"], anchor="w") \
        .grid(row=4, column=1)
    Radiobutton(ns_frame, text="Sat:   fixed", variable=ns, value="sat",
                bg="grey", fg="white", selectcolor=ns_color_dict["sat"], indicatoron=0, width=15, anchor="w") \
        .grid(row=5, column=0)
    Radiobutton(ns_frame, text="Mon:   fixed", variable=ns, value="mon",
                bg="grey", fg="white", selectcolor=ns_color_dict["mon"], indicatoron=0, width=15, anchor="w") \
        .grid(row=5, column=1)
    Radiobutton(ns_frame, text="Tue:   fixed", variable=ns, value="tue",
                bg="grey", fg="white", selectcolor=ns_color_dict["tue"], indicatoron=0, width=15, anchor="w") \
        .grid(row=6, column=0)
    Radiobutton(ns_frame, text="Wed:   fixed", variable=ns, value="wed",
                bg="grey", fg="white", selectcolor=ns_color_dict["wed"], indicatoron=0, width=15, anchor="w") \
        .grid(row=6, column=1)
    Radiobutton(ns_frame, text="Thu:   fixed", variable=ns, value="thu",
                bg="grey", fg="white", selectcolor=ns_color_dict["thu"], indicatoron=0, width=15, anchor="w") \
        .grid(row=7, column=0)
    Radiobutton(ns_frame, text="Fri:   fixed", variable=ns, value="fri",
                bg="grey", fg="white", selectcolor=ns_color_dict["fri"], indicatoron=0, width=15, anchor="w") \
        .grid(row=7, column=1)

    ns_frame.grid(row=4, sticky=W)
    # set route entry field
    route_frame = Frame(F, bd=1, relief=RIDGE, pady=2)
    Label(route_frame, text="route/s", width=15, anchor="w").grid(row=0, column=0, sticky=W)
    route = StringVar(route_frame)
    route.set(results[0][4])
    Entry(route_frame, width=37, textvariable=route).grid(row=1, column=0, sticky=W)
    route_frame.grid(row=5, sticky=W)
    # set station option menu
    station_frame = Frame(F, pady=2)
    Label(station_frame, text="station", width=10, anchor="w").grid(row=0, column=0, sticky=W)
    station = StringVar(station_frame)
    station.set(results[0][5])  # default value
    OptionMenu(station_frame, station, *list_of_stations).grid(row=0, column=1, sticky=W)
    # set rowid
    rowid = StringVar(F)
    rowid = results[0][6]
    station_frame.grid(row=6, sticky=W)
    #   History of status changes
    history_frame = Frame(F, pady=2)
    row_line = 0
    Label(history_frame, width=25, text="Status Change History", anchor="w", font="bold") \
        .grid(row=row_line, column=0, sticky=W, columnspan=4)
    row_line += 1
    for line in results:
        con_date = datetime.strptime(line[0], "%Y-%m-%d %H:%M:%S")  # convert str to datetime obj.
        Label(history_frame, width=25, text="date: {}".format(str(con_date.strftime("%b %d, %Y"))), anchor="w") \
            .grid(row=row_line, column=0, sticky=W, columnspan=4)
        row_line += 1
        Label(history_frame, width=25, text="list status: {}".format(line[2]), anchor="w") \
            .grid(row=row_line, column=0, sticky=W, columnspan=4)
        row_line += 1
        Label(history_frame, width=25, text="ns day: {}".format(ns_dict[line[3]]), anchor="w") \
            .grid(row=row_line, column=0, sticky=W, columnspan=4)
        row_line += 1
        Label(history_frame, width=25, text="route: {}".format(line[4]), anchor="w") \
            .grid(row=row_line, column=0, sticky=W, columnspan=4)
        row_line += 1
        Label(history_frame, width=25, text="station: {}".format(line[5]), anchor="w") \
            .grid(row=row_line, column=0, sticky=W, columnspan=4)
        row_line += 1
        Button(history_frame, width=14, text="edit", anchor="w",
               command=lambda x=line: [switchF3.destroy(), update_carrier(x)]) \
            .grid(row=row_line, column=0, sticky=W, )
        Button(history_frame, width=14, text="delete", anchor="w",
               command=lambda x=line: [switchF3.destroy(), delete_carrier(x)]) \
            .grid(row=row_line, column=1, sticky=W)
        row_line += 1
    history_frame.grid(row=7, sticky=W)
    root.update()
    C.config(scrollregion=C.bbox("all"))
    # apply and close buttons
    Button(C1, text="Apply", width=15, anchor="w",
           command=lambda: [apply(year, month, day, name, ls, ns, route, station, switchF3), switchF3.destroy(),
                            main_frame()]).pack(side=LEFT)
    Button(C1, text="Go Back", width=15, anchor="w",
           command=lambda: [switchF3.destroy(), main_frame()]).pack(side=LEFT)


def nc_apply(year, month, day, nc_name, nc_fname, nc_ls, nc_ns, nc_route, nc_station, self):
    if year.get() > 9999:
        messagebox.showerror("Year Input Error", "Year must be between 1 and 9999", parent=self)
        return
    if year.get() < 1:
        messagebox.showerror("Year Input Error", "Year must be between 1 and 9999", parent=self)
        return
    try:
        date = datetime(year.get(), month.get(), day.get())
    except:
        messagebox.showerror("Invalid Date", "Date entered is not valid", parent=self)
        return
    carrier = nc_name.get().strip().lower() + ", " + nc_fname.get().strip().lower()
    if len(nc_name.get()) > 30 or len(nc_fname.get()) > 12:
        messagebox.showerror("Name input error", "Names must not exceed 30 characters."
                                                 "First names must not exceed 12 characters", parent=self)
        return
    if len(nc_name.get()) < 1:
        messagebox.showerror("Name input error", "You must enter a name.", parent=self)
        return
    if len(nc_fname.get()) < 1:
        messagebox.showerror("Name input error", "You must enter a first initial or name.", parent=self)
        return
    if len(nc_fname.get()) > 1:
        answer = messagebox.askyesno("Caution", "It is recommended that you use only the first initial of the first"
                                                "name unless it is necessary to create a unique identifier, such as"
                                                "when you have two identical names that must be distinquished."
                                                "Do you want to proceed?", parent=self)
        if answer == False: return
    nc_route_list = nc_route.get().split("/")
    if len(nc_route.get()) > 24:
        messagebox.showerror("Route number input error", "There can be no more than five routes per carrier "
                                                         "(for T6 carriers).\n Routes numbers can be no more than four digits long.\n"
                                                         "If there are multiple routes, route numbers must be separated by "
                                                         "the \'/\' character. For example: 1001/1015/1024/1036/1072. Do not use "
                                                         "commas or empty spaces"
                             , parent=self)
        return
    for item in nc_route_list:
        item = item.strip()
        if item != "":
            if len(item) != 4:
                messagebox.showerror("Route number input error", 'Routes numbers must be four digits long.\n'
                                                                 'If there are multiple routes, route numbers must be separated by '
                                                                 'the \'/\' character. For example: 1001/1015/1024/1036/1072. Do not use '
                                                                 'commas or empty spaces'
                                     , parent=self)
                return
        if item.isdigit() == FALSE and item != "":
            messagebox.showerror("Route number input error", "Route numbers must be numbers and can not contain "
                                                             "letters", parent=self)
            return
    route_input = nc_route.get()
    if route_input == "0000":
        route_input = ""
    # check to see if new carrier name is already in carrier table
    match = False
    sql = "SELECT carrier_name, effective_date FROM carriers"
    results = inquire(sql)
    name_set = set()
    for x in results:
        name_set.add(x[0])
    sql = "INSERT INTO carriers (effective_date,carrier_name,list_status,ns_day,route_s,station)" \
          " VALUES('%s','%s','%s','%s','%s','%s')" \
          % (date, carrier, nc_ls.get(), nc_ns.get(), route_input, nc_station.get())
    if carrier in name_set:
        ok = messagebox.askokcancel("New Carrier Input Warning", "This carrier name is already in the database.\n"
                                                                 "Did you want to proceed?", parent=self)
        if ok == True:
            for pair in results:
                if pair[0] == carrier and pair[1] == str(datetime(year.get(), month.get(), day.get(), 00, 00, 00)):
                    messagebox.showwarning("New Carrier - Prohibited Action",
                                           "There is a pre existing record for this carrier on this day.\n"
                                           "You can not update that record using this window.\n"
                                           "To edit/ delete this record, return to the main page and press\n"
                                           "\"edit\" to the right of the carrier's name. ",
                                           parent=self)
                    match = True
        if ok == False:
            match = True
    if match == False: commit(sql)

    self.destroy()
    main_frame()


def input_carriers(frame):  # window for inputting new carriers
    # get ns day color configurations
    sql = "SELECT * FROM ns_configuration"
    ns_results = inquire(sql)
    ns_dict = {}  # build dictionary for ns days
    ns_color_dict = {}
    days = ("sat", "mon", "tue", "wed", "thu", "fri")
    for r in ns_results:  # build dictionary for rotating ns days
        ns_dict[r[0]] = r[2]
        ns_color_dict[r[0]] = r[1]  # build dictionary for ns fill colors
    for d in days:  # expand dictionary for fixed days
        ns_dict[d] = "fixed: " + d
        ns_color_dict[d] = "teal"
    ns_dict["none"] = "none"  # add "none" to dictionary
    ns_color_dict["none"] = "teal"
    frame.destroy()
    switchF6 = Frame(root)
    switchF6.pack(fill=BOTH, side=LEFT)
    C1 = Canvas(switchF6)
    C1.pack(fill=BOTH, side=BOTTOM)
    Button(C1, text="Apply", width=15, anchor="w",
           command=lambda: (
               nc_apply(year, month, day, nc_name, nc_fname, nc_ls, nc_ns, nc_route, nc_station, switchF6))) \
        .pack(side=LEFT)
    Button(C1, text="Go Back", width=15, anchor="w",
           command=lambda: [switchF6.destroy(), main_frame()]).pack(side=LEFT)
    # set up variable for scrollbar and canvas
    S = Scrollbar(switchF6)
    C = Canvas(switchF6, width=1600)
    # link up the canvas and scrollbar
    S.pack(side=RIGHT, fill=BOTH)
    C.pack(side=LEFT, fill=BOTH, pady=10, padx=20)
    S.configure(command=C.yview, orient="vertical")
    C.configure(yscrollcommand=S.set)
    if sys.platform == "win32":
        C.bind_all('<MouseWheel>', lambda event: C.yview_scroll(int(-1 * (event.delta / 120)), "units"))
    elif sys.platform == "linux":
        C.bind_all('<Button-4>', lambda event: C.yview('scroll', -1, 'units'))
        C.bind_all('<Button-5>', lambda event: C.yview('scroll', 1, 'units'))
    # create the frame inside the canvas
    nc_F = Frame(C)
    C.create_window((0, 0), window=nc_F, anchor=NW)
    # page title
    title_F = Frame(nc_F)
    Label(title_F, text="Enter New Carrier", font="bold").grid(row=0, column=0, columnspan=4)
    title_F.grid(row=0, sticky=W)  # put frame on grid
    # date
    date_frame = Frame(nc_F)  # define frame
    year = IntVar(date_frame)  # define variables for date
    month = IntVar(date_frame)
    day = IntVar(date_frame)
    month.set(gs_mo)  # set values for variables
    day.set(gs_day)
    year.set(gs_year)
    Label(date_frame, text="date (month/day/year):").grid(row=0, column=0, sticky=W, columnspan=3)  # date label
    OptionMenu(date_frame, month, "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12") \
        .grid(row=1, column=0, sticky=W)
    OptionMenu(date_frame, day, "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14",
               "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28", "29",
               "30", "31").grid(row=1, column=1, sticky=W)
    Entry(date_frame, width=6, textvariable=year).grid(row=1, column=2, sticky=W)
    date_frame.grid(row=1, sticky=W)  # put frame on grid
    # carrier name:
    name_frame = Frame(nc_F, pady=2)
    Label(name_frame, text="last name: ", anchor="w").grid(row=0, column=0, sticky=W)
    Label(name_frame, text="1st initial: ", anchor="w").grid(row=0, column=1, sticky=W)
    nc_name = StringVar(nc_F)
    nc_fname = StringVar(nc_F)
    Entry(name_frame, width=29, textvariable=nc_name).grid(row=1, column=0, sticky=W)
    Entry(name_frame, width=8, textvariable=nc_fname).grid(row=1, column=1, sticky=W)
    name_frame.grid(row=2, sticky=W)
    # list status
    list_frame = Frame(nc_F, bd=1, relief=RIDGE, pady=2)
    Label(list_frame, width=15, text="list status", anchor="w").grid(row=0, column=0, sticky=W)
    nc_ls = StringVar(list_frame)
    nc_ls.set(value="nl")
    Radiobutton(list_frame, text="OTDL", variable=nc_ls, value='otdl', justify=LEFT).grid(row=1, column=0, sticky=W)
    Radiobutton(list_frame, text="Work Assignment", variable=nc_ls, value='wal', justify=LEFT).grid(row=1, column=1,
                                                                                                    sticky=W)
    Radiobutton(list_frame, text="No List", variable=nc_ls, value='nl', justify=LEFT).grid(row=2, column=0, sticky=W)
    Radiobutton(list_frame, text="Auxiliary", variable=nc_ls, value='aux', justify=LEFT).grid(row=2, column=1, sticky=W)
    list_frame.grid(row=3, sticky=W)
    # set non scheduled day
    ns_frame = Frame(nc_F, pady=2)
    Label(ns_frame, width=15, text="non scheduled day", anchor="w").grid(row=0, column=0, sticky=W)
    nc_ns = StringVar(ns_frame)
    nc_ns.set("none")
    Radiobutton(ns_frame, text="{}:   yellow".format(ns_code['yellow']), variable=nc_ns, value="yellow", indicatoron=0,
                width=15, anchor="w", bg="grey", fg="white", selectcolor=ns_color_dict["yellow"]).grid(row=1, column=0)
    Radiobutton(ns_frame, text="{}:   blue".format(ns_code['blue']), variable=nc_ns, value="blue", indicatoron=0,
                width=15, anchor="w", bg="grey", fg="white", selectcolor=ns_color_dict["blue"]).grid(row=2, column=0)
    Radiobutton(ns_frame, text="{}:   green".format(ns_code['green']), variable=nc_ns, value="green", indicatoron=0,
                width=15, anchor="w", bg="grey", fg="white", selectcolor=ns_color_dict["green"]).grid(row=3, column=0)
    Radiobutton(ns_frame, text="{}:   brown".format(ns_code['brown']), variable=nc_ns, value="brown", indicatoron=0,
                width=15, anchor="w", bg="grey", fg="white", selectcolor=ns_color_dict["brown"]).grid(row=1, column=1)
    Radiobutton(ns_frame, text="{}:   red".format(ns_code['red']), variable=nc_ns, value="red", indicatoron=0, width=15,
                anchor="w", bg="grey", fg="white", selectcolor=ns_color_dict["red"]).grid(row=2, column=1)
    Radiobutton(ns_frame, text="{}:   black".format(ns_code['black']), variable=nc_ns, value="black", indicatoron=0,
                width=15, anchor="w", bg="grey", fg="white", selectcolor=ns_color_dict["black"]).grid(row=3, column=1)
    Radiobutton(ns_frame, text="none", variable=nc_ns, value="none", indicatoron=0, width=15, anchor="w")\
                .grid(row=4,column=1)
    Label(ns_frame, text="Fixed:", anchor="w").grid(row=4, column=0, sticky="w")
    Radiobutton(ns_frame, text="none", variable=nc_ns, value="none", indicatoron=0, width=15, bg="grey", fg="white",
                selectcolor=ns_color_dict["none"], anchor="w").grid(row=4, column=1)
    Radiobutton(ns_frame, text="Sat:   fixed", variable=nc_ns, value="sat", bg="grey", fg="white",
                selectcolor=ns_color_dict["sat"], indicatoron=0, width=15, anchor="w").grid(row=5, column=0)
    Radiobutton(ns_frame, text="Mon:   fixed", variable=nc_ns, value="mon", bg="grey", fg="white",
                selectcolor=ns_color_dict["mon"], indicatoron=0, width=15, anchor="w").grid(row=5, column=1)
    Radiobutton(ns_frame, text="Tue:   fixed", variable=nc_ns, value="tue", bg="grey", fg="white",
                selectcolor=ns_color_dict["tue"], indicatoron=0, width=15, anchor="w").grid(row=6, column=0)
    Radiobutton(ns_frame, text="Wed:   fixed", variable=nc_ns, value="wed", bg="grey", fg="white",
                selectcolor=ns_color_dict["wed"], indicatoron=0, width=15, anchor="w").grid(row=6, column=1)
    Radiobutton(ns_frame, text="Thu:   fixed", variable=nc_ns, value="thu", bg="grey", fg="white",
                selectcolor=ns_color_dict["thu"], indicatoron=0, width=15, anchor="w").grid(row=7, column=0)
    Radiobutton(ns_frame, text="Fri:   fixed", variable=nc_ns, value="fri", bg="grey", fg="white",
                selectcolor=ns_color_dict["fri"], indicatoron=0, width=15, anchor="w").grid(row=7, column=1)
    ns_frame.grid(row=4, sticky=W)
    # set route entry field
    route_frame = Frame(nc_F, bd=1, relief=RIDGE, pady=2)
    Label(route_frame, text="route/s", width=15, anchor="w").grid(row=0, column=0, sticky=W)
    nc_route = StringVar(route_frame)
    nc_route.set("")
    Entry(route_frame, width=37, textvariable=nc_route).grid(row=1, column=0, sticky=W)
    route_frame.grid(row=5, sticky=W)
    # set station option menu
    station_frame = Frame(nc_F, pady=2)
    Label(station_frame, text="station", width=10, anchor="w").grid(row=0, column=0, sticky=W)
    nc_station = StringVar(station_frame)
    nc_station.set(g_station)  # default value
    OptionMenu(station_frame, nc_station, *list_of_stations).grid(row=0, column=1, sticky=W)
    station_frame.grid(row=6, sticky=W)
    root.update()
    C.config(scrollregion=C.bbox("all"))


def reset(frame):
    global gs_year
    global gs_mo
    global gs_day
    global g_range
    global g_station
    global ns_code
    global g_date
    global d_date
    # reset initial value of globals
    gs_year = "x"
    gs_mo = "x"
    gs_day = "x"
    g_range = "x"
    g_station = "x"
    g_date = []
    if frame != "none":
        frame.destroy()
        main_frame()


def set_globals(s_year, s_mo, s_day, i_range, station, frame):
    global gs_year
    global gs_mo
    global gs_day
    global g_range
    global g_station
    global ns_code
    global g_date
    global d_date
    global pay_period
    g_range = i_range
    if station == "undefined":
        messagebox.showerror("Investigation station setting", 'Please select a station.                 ', parent=frame)
        return
    # error check for valid date
    try:
        date = datetime(int(s_year), int(s_mo), int(s_day))
        valid_date = True
    except ValueError:
        valid_date = False
    if valid_date == True:
        d_date = date
        wkdy_name = date.strftime("%a")
        while wkdy_name != "Sat":  # while date enter is not a saturday
            date -= timedelta(days=1)  # walk back the date until it is a saturday
            wkdy_name = date.strftime("%a")
        sat_range = date  # sat range = sat or the sat most prior
        pay_period = pp_by_date(sat_range)
        gs_year = int(date.strftime("%Y"))  # format that sat to form the global
        gs_mo = int(date.strftime("%m"))
        gs_day = int(date.strftime("%d"))
        del g_date[:]  # empty out the array for the global date variable
        d = datetime(int(gs_year), int(gs_mo), int(gs_day))
        # set the g_date variable
        g_date.append(d)
        for i in range(6):
            d += timedelta(days=1)
            g_date.append(d)
        # define color sequence tuple
        pat = ("blue", "green", "brown", "red", "black", "yellow")
        # calculate the n/s day of sat/first day of investigation range
        end_date = sat_range + timedelta(days=-1)
        cdate = datetime(2017, 1, 7)
        x = 0
        if sat_range > cdate:
            while cdate < end_date:
                if x > 0:
                    x -= 1
                    cdate += timedelta(days=7)
                else:
                    x = 5
                    cdate += timedelta(days=7)
        else:
            # IN REVERSE
            while cdate > sat_range:
                if x < 5:
                    x += 1
                    cdate -= timedelta(days=7)
                else:
                    x = 0
                    cdate -= timedelta(days=7)
        # find ns day for each day in range
        date = sat_range
        ns_code = {}
        for i in range(7):
            if i == 0:
                ns_code[pat[x]] = date.strftime("%a")
                date += timedelta(days=1)
            elif i == 1:
                date += timedelta(days=1)
                if x > 4:
                    x = 0
                else:
                    x += 1
            else:
                ns_code[pat[x]] = date.strftime("%a")
                date += timedelta(days=1)
                if x > 4:
                    x = 0
                else:
                    x += 1
        ns_code["none"] = "  "
        if i_range == "day":
            date = datetime(int(s_year), int(s_mo), int(s_day))
            # f_date = date.strftime("%A - %B %d,%Y")
            gs_year = int(s_year)
            gs_mo = int(s_mo)
            gs_day = int(s_day)
        ns_code["sat"] = "Sat"
        ns_code["mon"] = "Mon"
        ns_code["tue"] = "Tue"
        ns_code["wed"] = "Wed"
        ns_code["thu"] = "Thu"
        ns_code["fri"] = "Fri"
    else:
        messagebox.showerror("Investigation date/range", 'The date entered is not valid.', parent=frame)
        return
    g_station = station
    if frame != "None":
        frame.destroy()
        main_frame()


def main_frame():
    F = Frame(root)
    F.pack(fill=BOTH, side=LEFT)
    C1 = Canvas(F)
    C1.pack(fill=BOTH, side=BOTTOM)
    if gs_day != "x":
        Button(C1, text="New Carrier", command=lambda: input_carriers(F),
               width=12).pack(side=LEFT)
        Button(C1, text="Multi Input", command=lambda dd="Sat", ss="name": mass_input(F, dd, ss),
               width=12).pack(side=LEFT)
        Button(C1, text="Report", command=lambda: output_tab(F, carrier_list),
               width=12).pack(side=LEFT)
        r_rings = "x"
        Button(C1, text="Spreadsheet", width=12,
               command=lambda: spreadsheet(carrier_list, r_rings)).pack(side=LEFT)
        Button(C1, text="Quit", width=12, command=root.destroy).pack(side=LEFT)
    # link up the canvas and scrollbar
    S = Scrollbar(F)
    C = Canvas(F, width=1600)
    S.pack(side=RIGHT, fill=BOTH)
    C.pack(side=LEFT, fill=BOTH, pady=10, padx=10)
    S.configure(command=C.yview, orient="vertical")
    C.configure(yscrollcommand=S.set)
    if sys.platform == "win32":
        C.bind_all('<MouseWheel>', lambda event: C.yview_scroll(int(-1 * (event.delta / 120)), "units"))
    elif sys.platform == "darwin":
        C.bind_all('<MouseWheel>', lambda event: C.yview_scroll(int(event.delta)))  # Maybe: this is a guess
    elif sys.platform == "linux":
        C.bind_all('<Button-4>', lambda event: C.yview('scroll', -1, 'units'))
        C.bind_all('<Button-5>', lambda event: C.yview('scroll', 1, 'units'))
    # create a pulldown menu, and add it to the menu bar
    menubar = Menu(F)
    # file menu
    basic_menu = Menu(menubar, tearoff=0)
    basic_menu.add_command(label="Save All", command=lambda: save_all(F))
    basic_menu.add_separator()
    basic_menu.add_command(label="New Carrier", command=lambda: input_carriers(F))
    basic_menu.add_command(label="Multiple Input", command=lambda dd="Sat", ss="name": mass_input(F, dd, ss))
    basic_menu.add_command(label="Report Summary", command=lambda: output_tab(F, carrier_list))
    basic_menu.add_command(label="Create Spreadsheet", command=lambda: spreadsheet(carrier_list, r_rings))
    basic_menu.add_command(label="Over Max Spreadsheet", command=lambda r_rings="x": overmax_spreadsheet(carrier_list))
    if gs_day == "x":
        basic_menu.entryconfig(2, state=DISABLED)
        basic_menu.entryconfig(3, state=DISABLED)
        basic_menu.entryconfig(4, state=DISABLED)
        basic_menu.entryconfig(5, state=DISABLED)
        basic_menu.entryconfig(6, state=DISABLED)
    basic_menu.add_separator()
    basic_menu.add_command(label="Informal C", command=lambda: informalc(F))
    basic_menu.add_separator()
    basic_menu.add_command(label="Quit", command=lambda: root.destroy())
    menubar.add_cascade(label="Basic", menu=basic_menu)
    # automated menu
    automated_menu = Menu(menubar, tearoff=0)
    automated_menu.add_command(label="Automatic Data Entry", command=lambda: call_indexers(F))
    automated_menu.add_command(label=" Auto Over Max Finder", command=lambda: max_hr())
    automated_menu.add_separator()
    automated_menu.add_command(label="Everything Report Reader", command=lambda: ee_skimmer())
    automated_menu.add_command(label="Pay Period Guide Generator", command=lambda: pay_period_guide(F))
    automated_menu.add_command(label="Weekly Availability", command=lambda: wkly_avail(F))
    automated_menu.add_separator()
    automated_menu.add_command(label="PDF Converter", command=lambda: pdf_converter())
    automated_menu.add_command(label="PDF Splitter", command=lambda: pdf_splitter(F))
    menubar.add_cascade(label="Automated", menu=automated_menu)

    # reports menu
    reports_menu = Menu(menubar, tearoff=0)
    reports_menu.add_command(label="Carrier Route and NS Day", command=lambda: rpt_carrier(carrier_list))
    reports_menu.add_command(label="Carrier Route", command=lambda: rpt_carrier_route(carrier_list))
    reports_menu.add_command(label="Carrier NS Day", command=lambda: rpt_carrier_nsday(carrier_list))
    # reports_menu.add_command(label="Improper Mandates", command=lambda: rpt_impman(carrier_list))
    if gs_day == "x":
        reports_menu.entryconfig(0, state=DISABLED)
        reports_menu.entryconfig(1, state=DISABLED)
        reports_menu.entryconfig(2, state=DISABLED)
        # reports_menu.entryconfig(3, state=DISABLED)
    menubar.add_cascade(label="Reports", menu=reports_menu)
    # library menu
    reportsarchive_menu = Menu(menubar, tearoff=0)
    reportsarchive_menu.add_command(label="Spreadsheet Archive", command=lambda: file_dialogue('kb_sub/spreadsheets'))
    reportsarchive_menu.add_command(label="Over Max Finder Archive", command=lambda: file_dialogue('kb_sub/over_max'))
    reportsarchive_menu.add_command(label="Over Max Spreadsheet Archive",
                             command=lambda: file_dialogue('kb_sub/over_max_spreadsheet'))
    reportsarchive_menu.add_command(label="Everything Report Archive", command=lambda: file_dialogue('kb_sub/ee_reader'))
    reportsarchive_menu.add_command(label="Pay Period Guide Archive", command=lambda: file_dialogue('kb_sub/pp_guide'))
    reportsarchive_menu.add_command(label="Weekly Availability Archive",
                             command=lambda: file_dialogue('kb_sub/weekly_availability'))
    reportsarchive_menu.add_separator()
    reportsarchive_menu.add_command(label="Empty Spreadsheet Archive",
                                    command=lambda: remove_file_var('kb_sub/spreadsheets'))
    reportsarchive_menu.add_command(label="Empty Over Max Finder Archive",
                                    command=lambda: remove_file_var('kb_sub/over_max'))
    reportsarchive_menu.add_command(label="Empty Over Max Spreadsheet Archive",
                                    command=lambda: remove_file_var('kb_sub/over_max_spreadsheet'))
    reportsarchive_menu.add_command(label="Empty Everything Report Archive",
                                    command=lambda: remove_file_var('kb_sub/ee_reader'))
    reportsarchive_menu.add_command(label="Empty Pay Period Guide Archive",
                                    command=lambda: remove_file_var('kb_sub/pp_guide'))
    reportsarchive_menu.add_command(label="Empty Weekly Availability Archive Archive",
                                    command=lambda: remove_file_var('kb_sub/weekly_availability'))
    menubar.add_cascade(label="Archive", menu=reportsarchive_menu)
    # management menu
    management_menu = Menu(menubar, tearoff=0)
    management_menu.add_command(label="List of Stations", command=lambda: station_list(F))
    management_menu.add_command(label="Tolerances", command=lambda: tolerances(F))
    management_menu.add_command(label="Spreadsheet Settings", command=lambda: spreadsheet_settings(F))
    management_menu.add_command(label="Auto Data Entry Settings", command=lambda: auto_data_entry_settings(F))
    management_menu.add_command(label="Clean Carrier List", command=lambda: carrier_list_cleaning(F))
    management_menu.add_command(label="Clean Rings", command=lambda: clean_rings3_table())
    management_menu.add_command(label="Name Index", command=lambda: (F.destroy(), name_index_screen()))
    management_menu.add_command(label="Station Index", command=lambda: station_index_mgmt(F))
    management_menu.add_command(label="PDF Converter Settings", command=lambda: pdf_converter_settings(F))
    management_menu.add_command(label="NS Day Configurations", command=lambda: ns_config(F))
    management_menu.add_separator()
    management_menu.add_command(label="Location", command=lambda: location_klusterbox(F))
    management_menu.add_command(label="About Klusterbox", command=lambda: about_klusterbox(F))
    menubar.add_cascade(label="Management", menu=management_menu)
    root.config(menu=menubar)
    # create the frame inside the canvas
    preF = Frame(C)
    C.create_window((0, 0), window=preF, anchor=NW)
    FF = Frame(C)
    C.create_window((0, 108), window=FF, anchor=NW)
    # set up tkinter variables for time and place:
    now = datetime.now()
    start_year = IntVar(preF)
    start_month = IntVar(preF)
    start_day = IntVar(preF)
    i_range = StringVar(preF)
    # set up labels for the investigation range and station
    Label(preF, text="INVESTIGATION DATE").grid(row=1, column=1, columnspan=2)
    if gs_mo == "x":
        start_month.set(now.month)
    else:
        start_month.set(gs_mo)
    om_month = OptionMenu(preF, start_month, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12)
    om_month.config(width=2)
    om_month.grid(row=1, column=3)
    if gs_day == "x":
        start_day.set(now.day)
    else:
        start_day.set(gs_day)
    om_day = OptionMenu(preF, start_day, "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14",
                        "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28", "29", "30",
                        "31")
    om_day.config(width=2)
    om_day.grid(row=1, column=4)
    date_year = Entry(preF, width=6, textvariable=start_year)
    if gs_year == "x":
        start_year.set(now.year)
    else:
        start_year.set(gs_year)
    date_year.grid(row=1, column=5)
    Label(preF, text="RANGE").grid(row=1, column=6)
    if g_range == "x":
        i_range.set("week")
    else:
        i_range.set(g_range)
    Radiobutton(preF, text="weekly", variable=i_range, value="week", width=6, anchor="w").grid(row=1, column=7)
    Radiobutton(preF, text="daily", variable=i_range, value="day", width=5, anchor="w").grid(row=1, column=8)
    # set station option menu
    Label(preF, text="STATION", anchor="w").grid(row=2, column=1)
    station = StringVar(F)
    if g_station == "x":
        station.set("undefined")  # default value
    else:
        station.set(g_station)
    om = OptionMenu(preF, station, *list_of_stations)
    om.config(width="35")
    om.grid(row=2, column=2, columnspan=5, sticky=W)
    # set and reset buttons for investigation range
    Button(preF, text="Set", anchor="w", width=8,
           command=lambda: set_globals(start_year.get(), start_month.get(), start_day.get(), i_range.get(),
                                       station.get(), F)) \
        .grid(row=2, column=7)
    Button(preF, text="Reset", anchor="w", width=8, command=lambda: reset(F)).grid(row=2, column=8)
    # Investigation date SET/NOT SET notification
    if g_range == "x":
        Label(preF, text="Investigation date/range not set", foreground="red") \
            .grid(row=3, column=1, columnspan=8, sticky="w")
    elif g_range == "day":
        f_date = d_date.strftime("%a - %b %d, %Y")
        Label(preF, text="Investigation Date Set: {}".format(f_date),
              foreground="red").grid(row=3, column=1, columnspan=8, sticky="w")
        Label(preF, text="Pay Period: {}".format(pay_period),
              foreground="red").grid(row=4, column=1, columnspan=8, sticky="w")
    else:
        f_date = g_date[0].strftime("%a - %b %d, %Y")
        end_f_date = g_date[6].strftime("%a - %b %d, %Y")
        Label(preF, text="Investigation Range: {0} through {1}".format(f_date, end_f_date),
              foreground="red").grid(row=3, column=1, columnspan=8, sticky="w")
        Label(preF, text="Pay Period: {0}".format(pay_period),
              foreground="red").grid(row=4, column=1, columnspan=8, sticky="w")
    if gs_day == "x":
        Label(F, text=" Please input Investigation Range and Station").pack()
    else:
        if g_range == "week":
            sql = "SELECT effective_date, carrier_name,list_status, ns_day,route_s, station, rowid" \
                  " FROM carriers WHERE effective_date <= '%s'" \
                  "ORDER BY carrier_name, effective_date desc" % (g_date[6])
        if g_range == "day":
            sql = "SELECT effective_date, carrier_name,list_status, ns_day,route_s, station, rowid" \
                  " FROM carriers WHERE effective_date <= '%s'" \
                  "ORDER BY carrier_name, effective_date desc" % (d_date)
        results = inquire(sql)
        # initialize arrays for data sorting
        carrier_list = []
        candidates = []
        more_rows = []
        pre_invest = []
        # take raw data and sort into appropriate arrays
        for i in range(len(results)):
            candidates.append(results[i])  # put name into candidates array
            jump = "no"  # triggers an analysis of the candidates array
            if i != len(results) - 1:  # if the loop has not reached the end of the list
                if results[i][1] == results[i + 1][1]:  # if the name current and next name are the same
                    jump = "yes"  # bypasses an analysis of the candidates array
            if jump == "no":
                # sort into records in investigation range and those prior
                for record in candidates:
                    if g_range == "week":  # if record falls in investigation range - add it to more rows array
                        if record[0] >= str(g_date[1]) and record[0] <= str(g_date[6]): more_rows.append(record)
                        if record[0] <= str(g_date[0]) and len(pre_invest) == 0: pre_invest.append(record)
                    if g_range == "day":
                        if record[0] <= str(d_date) and len(pre_invest) == 0: pre_invest.append(record)
                # find carriers who start in the middle of the investigation range CATEGORY ONE
                if len(more_rows) > 0 and len(pre_invest) == 0:
                    station_anchor = "no"
                    for each in more_rows:  # check if any records place the carrier in the selected station
                        if each[5] == g_station: station_anchor = "yes"  # if so, set the station anchor
                    # since the carrier starts in the middle of the week and there is no record prior, create one
                    if station_anchor == "yes":
                        filler = (str(g_date[0]), each[1], " ", "none", " ", "out of station", 0, "A_out")
                        carrier_list.append(list(filler))
                        list(more_rows)
                        more_rows.reverse()  # reverse the tuple
                        for each in more_rows:
                            # carrier_list.append(list(each))
                            x = list(each)  # convert the tuple to a list
                            if x[5] == g_station:
                                x.append("B_in")  # tag if the record is the first in the list
                            else:
                                x.append("B_out")
                            carrier_list.append(x)  # add it to the list
                # find carriers with records before and during the investigation range CATEGORY TWO
                if len(more_rows) > 0 and len(pre_invest) > 0:
                    station_anchor = "no"
                    for each in more_rows + pre_invest:
                        if each[5] == g_station: station_anchor = "yes"
                    if station_anchor == "yes":
                        # handle records prior to or on first day of investigation range.
                        xx = list(pre_invest[0])
                        if xx[5] == g_station:
                            xx.append("A_in")
                        else:
                            xx.append("A_out")
                        carrier_list.append(xx)
                        # handle records inside the investigation range
                        list(more_rows)
                        more_rows.reverse()
                        for each in more_rows:
                            x = list(each)
                            if x[5] == g_station:
                                x.append("B_in")
                            else:
                                x.append("B_out")
                            carrier_list.append(x)
                # find carrier with records from only before investigation range.CATEGORY THREE
                if len(more_rows) == 0 and len(pre_invest) == 1:
                    for each in pre_invest:
                        if each[5] == g_station:
                            x = list(pre_invest[0])
                            x.append("A_in")
                            carrier_list.append(x)
                del more_rows[:]
                del pre_invest[:]
                del candidates[:]
        # This code displays the records that have been selected above and placed in the carrier list array.
        r = 0
        i = 0
        ii = 1
        if len(carrier_list) == 0:
            Label(FF, text="").grid(row=0, column=0)
            Label(FF, text="The carrier list is empty. ", font="bold").grid(row=1, column=0, sticky="w")
            Label(FF, text="").grid(row=2, column=0)
            Label(FF, text="Build the carrier list with the New Carrier feature\nor by running "
                           "the Automatic Data Entry Feature.").grid(row=3, column=0)
        if len(carrier_list) > 0:
            Label(FF, text="Name (click for Rings)", fg="grey").grid(row=r, column=1, sticky="w")
            Label(FF, text="List", fg="grey").grid(row=r, column=2, sticky="w")
            Label(FF, text="N/S", fg="grey").grid(row=r, column=3, sticky="w")
            Label(FF, text="Route", fg="grey").grid(row=r, column=4, sticky="w")
            Label(FF, text="Edit", fg="grey").grid(row=r, column=5, sticky="w")
            r += 1
        for line in carrier_list:
            # if the row is even, then choose a color for it
            if i & 1:
                color = "light yellow"
            else:
                color = "white"
            if carrier_list[i][len(carrier_list[i])-1] == "A_in" or carrier_list[i][len(carrier_list[i])-1] == "A_out":
                Label(FF, text=ii).grid(row=r, column=0)
                ii += 1
                Button(FF, text=line[1], width=24, bg=color, anchor="w",
                       command=lambda x=line: rings2(x, root)).grid(row=r, column=1)
            else:
                dt = datetime.strptime(line[0], "%Y-%m-%d %H:%M:%S")
                Button(FF, text=dt.strftime("%a"), width=24, bg=color, anchor="e",
                       command=lambda x=line: rings2(x, root)).grid(row=r, column=1)
            if line[len(line)-1] == "A_in" or line[len(line)-1] == "B_in":
                Button(FF, text=line[2], width=3, bg=color, anchor="w").grid(row=r, column=2)
                day_off = ns_code[line[3]].lower()
                Button(FF, text=day_off, width=4, bg=color, anchor="w").grid(row=r, column=3)
                Button(FF, text=line[4], width=20, bg=color, anchor="w").grid(row=r, column=4)
            else:
                Button(FF, text="out of station", width=30, bg=color).grid(row=r, column=2, columnspan=3)
            Button(FF, text="edit", width=4, bg=color, anchor="w",
                   command=lambda x=line[1]: [F.destroy(), edit_carrier(x)]).grid(row=r, column=5)
            i += 1
            r += 1
    if g_station == "x":
        Button(FF, text="Automatic Data Entry", width=30, command=lambda: call_indexers(F)).grid(row=0, pady=5)
        Button(FF, text="Auto Over Max Finder", width=30, command=lambda: max_hr()).grid(row=1, pady=5)
        Button(FF, text="Informal C", width=30, command=lambda: informalc(F)).grid(row=2, pady=5)
        Label(FF, text="", width=70).grid(row=3)
    root.update()
    C.config(scrollregion=C.bbox("all"))
    mainloop()


if __name__ == "__main__":
    pb_root = Tk()  # create a window for the progress bar
    pb_root.title("Starting Klusterbox")
    pb_label = Label(pb_root, text="Running Setup: ")  # make label for progress bar
    pb_label.pack(side=LEFT)
    pb = ttk.Progressbar(pb_root, length=400, mode="determinate")  # create progress bar
    pb.pack(side=LEFT)
    steps = 20
    pb["maximum"] = steps  # set length of progress bar
    pb.start()
    pb["value"] = 1  # increment progress bar
    pb_root.update()
    global gs_year
    global gs_mo
    global gs_day
    global g_date
    global d_date
    global g_range
    global g_station
    global list_of_stations
    global pay_period
    # set initial value of globals
    gs_year = "x"
    gs_mo = "x"
    gs_day = "x"
    g_range = "x"
    g_station = "x"
    g_date = []
    # initialize arrays for multiple move functionality
    sat_mm = []
    sun_mm = []
    mon_mm = []
    tue_mm = []
    wed_mm = []
    thr_mm = []
    fri_mm = []
    # initialize position and size for root window
    position_x = 100
    position_y = 50
    size_x = 625
    size_y = 600
    pb["value"] = 1  # increment progress bar
    pb_root.update()
    # create kb_sub folder if it does not exist
    if os.path.isdir('kb_sub') == False:
        os.makedirs('kb_sub')
    # set up database if it does not exist
    sql = 'CREATE table IF NOT EXISTS stations (station varchar primary key)'
    commit(sql)
    sql = 'INSERT OR IGNORE INTO stations (station) VALUES ("out of station")'
    commit(sql)
    sql = 'CREATE table IF NOT EXISTS tolerances (row_id integer primary key, category varchar, tolerance varchar)'
    commit(sql)
    pb["value"] = 5  # increment progress bar
    pb_root.update()
    sql = 'INSERT OR IGNORE INTO tolerances (row_id, category, tolerance) VALUES (0,"ot_own_rt", .25)'
    commit(sql)
    sql = 'INSERT OR IGNORE INTO tolerances (row_id, category, tolerance) VALUES (1,"ot_tol", .25)'
    commit(sql)
    sql = 'INSERT OR IGNORE INTO tolerances (row_id, category, tolerance) VALUES (2,"av_tol", .25)'
    commit(sql)
    sql = 'INSERT OR IGNORE INTO tolerances (row_id, category, tolerance) VALUES (3,"min_ss_nl", 25)'
    commit(sql)
    sql = 'INSERT OR IGNORE INTO tolerances (row_id, category, tolerance) VALUES (4,"min_ss_wal", 25)'
    commit(sql)
    pb["value"] = 10  # increment progress bar
    pb_root.update()
    sql = 'INSERT OR IGNORE INTO tolerances (row_id, category, tolerance) VALUES (5,"min_ss_otdl", 25)'
    commit(sql)
    sql = 'INSERT OR IGNORE INTO tolerances (row_id, category, tolerance) VALUES (6,"min_ss_aux", 25)'
    commit(sql)
    sql = 'INSERT OR IGNORE INTO tolerances(row_id,category,tolerance)VALUES(7,"allow_zero_top","False")'
    commit(sql)
    sql = 'INSERT OR IGNORE INTO tolerances(row_id,category,tolerance)VALUES(8,"allow_zero_bottom","True")'
    commit(sql)
    sql = 'INSERT OR IGNORE INTO tolerances(row_id,category,tolerance)VALUES(9,"pdf_error_rpt","off")'
    commit(sql)
    sql = 'INSERT OR IGNORE INTO tolerances(row_id,category,tolerance)VALUES(10,"pdf_raw_rpt","off")'
    commit(sql)
    sql = 'INSERT OR IGNORE INTO tolerances(row_id,category,tolerance)VALUES(11,"pdf_text_reader","off")'
    commit(sql)
    sql = 'INSERT OR IGNORE INTO tolerances(row_id,category,tolerance)VALUES(12,"ns_auto_pref","rotation")'
    commit(sql)
    sql = 'CREATE table IF NOT EXISTS carriers (effective_date date, carrier_name varchar, list_status varchar, ' \
          ' ns_day varchar, route_s varchar, station varchar)'
    commit(sql)
    sql = 'CREATE table IF NOT EXISTS rings3 ' \
          '(rings_date date, carrier_name varchar, total varchar, rs varchar, code varchar, moves varchar, ' \
          'leave_type varchar, leave_time varchar)'
    commit(sql)
    # modify table for legacy version which did not have leave type and leave time columns of rings3 table.
    sql = 'PRAGMA table_info(rings3)' # get table info. returns an array of columns.
    result = inquire (sql)
    if len(result)<= 6: # if there are not enough columns add the leave type and leave time columns
        sql = 'ALTER table rings3 ADD COLUMN leave_type varchar'
        commit(sql)
        sql = 'ALTER table rings3 ADD COLUMN leave_time varchar'
        commit(sql)
    sql = 'CREATE table IF NOT EXISTS name_index (tacs_name varchar, kb_name varchar, emp_id varchar)'
    commit(sql)
    pb["value"] = 15  # increment progress bar
    pb_root.update()
    sql = 'CREATE table IF NOT EXISTS station_index (tacs_station varchar, kb_station varchar, finance_num varchar)'
    commit(sql)  # access list of stations from database
    sql = 'CREATE table IF NOT EXISTS skippers (code varchar primary key, description varchar)'
    commit(sql)
    sql = 'CREATE table IF NOT EXISTS ns_configuration (ns_name varchar primary key, fill_color varchar, ' \
          'custom_name varchar)'
    commit(sql)
    sql = 'INSERT OR IGNORE INTO ns_configuration(ns_name,fill_color,custom_name)VALUES("yellow","gold","yellow")'
    commit(sql)
    sql = 'INSERT OR IGNORE INTO ns_configuration(ns_name,fill_color,custom_name)VALUES("blue","navy","blue")'
    commit(sql)
    sql = 'INSERT OR IGNORE INTO ns_configuration(ns_name,fill_color,custom_name)VALUES("green","forest green","green")'
    commit(sql)
    sql = 'INSERT OR IGNORE INTO ns_configuration(ns_name,fill_color,custom_name)VALUES("brown","saddle brown","brown")'
    commit(sql)
    sql = 'INSERT OR IGNORE INTO ns_configuration(ns_name,fill_color,custom_name)VALUES("red","red3","red")'
    commit(sql)
    sql = 'INSERT OR IGNORE INTO ns_configuration(ns_name,fill_color,custom_name)VALUES("black","gray10","black")'
    commit(sql)
    sql = "SELECT * FROM stations ORDER BY station"
    results = inquire(sql)
    # define and populate list of stations variable
    list_of_stations = []
    for stat in results:
        list_of_stations.append(stat[0])
    pb["value"] = 20  # increment progress bar
    pb_root.update()
    pb.stop()  # stop and destroy the progress bar
    pb_label.destroy()  # destroy the label for the progress bar
    pb.destroy()
    pb_root.destroy()
    root = Tk()
    if sys.platform == "win32":
        try:
            root.iconbitmap(r'kb_sub/kb_images/kb_icon2.ico')
        except:
            pass
    if sys.platform == "linux":
        try:
            img = PhotoImage(file='kb_sub/kb_images/kb_icon2.gif')
            root.tk.call('wm', 'iconphoto', root._w, img)
        except:
            pass
    root.title("KLUSTERBOX version {}".format(version))
    root.geometry("%dx%d+%d+%d" % (size_x, size_y, position_x, position_y))
    # if there are no stations in the stations list
    if len(list_of_stations) < 2:
        start_up()
    else:
        remove_file('kb_sub/report') # empty out folders
        remove_file('kb_sub/infc_grv')
        main_frame()

