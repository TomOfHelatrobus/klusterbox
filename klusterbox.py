# custom modules
from kbreports import Reports, Messenger
from kbtoolbox import *
from kbspreadsheets import OvermaxSpreadsheet, ImpManSpreadsheet
from kbdatabase import DataBase, setup_plaformvar, setup_dirs_by_platformvar
from kbspeedsheets import SpeedSheetGen, OpenText
from kbequitability import QuarterRecs, OTEquitSpreadsheet, OTDistriSpreadsheet
from kbcsv_repair import CsvRepair
# Standard Libraries
from tkinter import *
from tkinter import messagebox, filedialog, ttk
from datetime import datetime, timedelta
import sqlite3
from operator import itemgetter
import os
import shutil
import csv
import sys
import subprocess
from io import StringIO  # change from cStringIO to io for py 3x
import time
import webbrowser  # for hyper link at about_klusterbox()
from threading import *  # run load workbook while progress bar runs
# Pillow Library
from PIL import ImageTk, Image  # Pillow Library
# Spreadsheet Libraries
from openpyxl import load_workbook
from openpyxl import Workbook
from openpyxl.styles import NamedStyle, Font, Border, Side, Alignment, PatternFill
# PDF Converter Libraries
from pdfminer.pdfparser import PDFParser
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter, resolve1
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
# PDF Splitter Libraries
from PyPDF2 import PdfFileReader, PdfFileWriter
# version variables
version = "4.006"
release_date = "Undetermined"
"""
 _   _ _                             _
| |/ /| |              _            | |
| | / | | _   _  ___ _| |_ ___  _ _ | |_   __  _  __
|  (  | || | | |/ __/_   _| __|| /_/|   \ /  \\ \/ /
| | \ | |\ \_| |\__ \ | | | _| | |  | () | () |)  (
|_|\_\|_| \____|/___/ |_| |___||_|  |___/ \__//_/\_\

Klusterbox
Copyright 2019 Thomas Weeks

Caution: To ensure proper operation of Legacy Klusterbox outside Program Files (Windows) or Applications (mac OS),
make sure to keep the Klusterbox.exe and the kb_sub folder in the same folder.

For the newest version of Klusterbox, visit www.klusterbox.com/download.
Visit https://github.com/TomOfHelatrobus/klusterbox for the most recent source code.

This version of Klusterbox is being released under the GNU General Public License version 3.
"""


class ProgressBarIn:  # Indeterminate Progress Bar
    def __init__(self, title="", label="", text=""):
        self.title = title
        self.label = label
        self.text = text
        self.pb_root = Tk()  # create a window for the progress bar
        self.pb_label = Label(self.pb_root, text=self.label)  # make label for progress bar
        self.pb = ttk.Progressbar(self.pb_root, length=400, mode="indeterminate")  # create progress bar
        self.pb_text = Label(self.pb_root, text=self.text, anchor="w")

    def start_up(self):
        titlebar_icon(self.pb_root)  # place icon in titlebar
        self.pb_root.title(self.title)
        self.pb_label.grid(row=0, column=0, sticky="w")
        self.pb.grid(row=1, column=0, sticky="w")
        self.pb_text.grid(row=2, column=0, sticky="w")
        while pb_flag:  # use global as a flag. stop loop when flag is False
            projvar.root.update()
            self.pb['value'] += 1
            time.sleep(.01)

    def stop(self):
        self.pb.stop()  # stop and destroy the progress bar
        self.pb_text.destroy()
        self.pb_label.destroy()  # destroy the label for the progress bar
        self.pb.destroy()
        self.pb_root.destroy()


class RefusalWin:  # create a window for refusals for otdl equitability
    def __init__(self):
        self.frame = None
        self.win = None
        self.row = 0
        self.carrier_name = ""
        self.startdate = datetime(1, 1, 1)
        self.enddate = datetime(1, 1, 1)
        self.station = ""
        self.time_vars = []  # a list of stringvars of refusal times
        self.type_vars = []  # a list of stringvars of refusal types/indicators.
        self.ref_dates = []  # a list of datetime objects corrosponding to refusal times and types
        self.displaydate = []  # a list of strings providing the date of the refusals
        self.refset = []  # a list of refusals for the quarter
        self.onrec_time = []  # a list of the refusal time in the database
        self.onrec_type = []  # a list of the refusal type/indicator in the database
        self.onrec_displaydate = []
        self.status_update = None

    def create(self, frame, carrier, startdate, enddate, station):
        self.frame = frame
        self.carrier_name = carrier
        self.startdate = startdate
        self.enddate = enddate
        self.station = station
        self.win = MakeWindow()
        self.win.create(self.frame)
        self.get_refset()
        self.setup_vars_and_stringvars()
        self.build_header()
        self.build()
        self.build_bottom()
        self.buttons_frame()
        self.win.finish()

    def get_refset(self):
        sql = "SELECT * FROM refusals WHERE refusal_date between '%s' and '%s' and carrier_name = '%s' " \
              "ORDER BY refusal_date" % (self.startdate, self.enddate, self.carrier_name)
        self.refset = inquire(sql)

    def setup_vars_and_stringvars(self):
        i = 0
        date = self.startdate  # this will be the first date
        while date != self.enddate + timedelta(days=1):  # for each date in the quarter
            self.time_vars.append(StringVar(self.win.body))  # create a stringvar for time
            self.type_vars.append(StringVar(self.win.body))  # create a stringvar for type
            self.ref_dates.append(date)  # create a list of datetime objs corrosponding to the time/type vars
            displaydate = date.strftime("%m") + "/" + date.strftime("%d")  # make a string of date eg 07/29
            self.displaydate.append(displaydate)  # create a list of dates as string corrosponding to time/type vars
            self.onrec_time.append("")  # create the onrec time array
            self.onrec_type.append("")  # create the onrec type array
            for line in self.refset:  # loop through refset for refusals on that date
                if dt_converter(line[0]) == date:  # if there is a match
                    self.type_vars[i].set(line[2])  # set the stringvar for type
                    self.time_vars[i].set(line[3])  # set the stringvar for time
                    self.onrec_type[i] = line[2]  # change the onrec type to type from refset
                    self.onrec_time[i] = line[3]  # change the onrec time to time from refset
                    # create list of dates with records in the database as a string date eg 07/29
                    self.onrec_displaydate.append(dt_converter(line[0]).strftime("%m") + "/" + date.strftime("%d"))
            date += timedelta(days=1)
            i += 1  # increment the counter

    def start_column(self):  # returns the column position of the startdate
        days = ("Saturday", "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday")
        i = 0
        for day in days:  # loop through tuple of days
            if self.startdate.strftime("%A") == day:  # if the startdate matches the day
                return i  # return the index of the tuple
            i += 3  # increment the counter

    def build_header(self):
        Label(self.win.body, text="Refusals: {}".format(self.carrier_name),
              font=macadj("bold", "Helvetica 18"), anchor="w") \
            .grid(row=self.row, column=0, sticky="w", columnspan=27)
        self.row += 1
        Label(self.win.body, text="").grid(row=self.row)
        self.row += 1
        Label(self.win.body, text="Investigation Range: {} though {}"
              .format(self.startdate.strftime("%m/%d/%Y"), self.enddate.strftime("%m/%d/%Y")), fg="red")\
            .grid(row=self.row, columnspan=macadj(20, 27), sticky="w")
        self.row += 1
        Label(self.win.body, text="Station: {}".format(self.station)) \
            .grid(row=self.row, columnspan=macadj(20, 27), sticky="w")
        self.row += 1
        text = "Fill in the Refusal Indicator (optional) in the small field and any Refusal " \
               "Times in the large field. "
        Label(self.win.body, text=text, anchor="e", justify=LEFT).grid(row=self.row, columnspan=macadj(20, 27))
        self.row += 1
        Label(self.win.body, text="").grid(row=self.row)
        self.row += 1
        column = 0
        days = ("Sat", "Sun", "Mon", "Tue", "Wed", "Thu", "Fri")
        for day in days:
            Label(self.win.body, width=macadj(7, 3), text=day, anchor="w", fg="Blue")\
                .grid(row=self.row, column=column+1, columnspan=3, sticky="w")
            column += 3
        self.row += 1

    def build(self):
        column = self.start_column()
        for i in range(len(self.time_vars)):
            Label(self.win.body, width=macadj(2, 0), text="").grid(row=self.row, column=column)  # blank column
            column += 1
            Label(self.win.body, width=macadj(7, 4), text=self.displaydate[i], fg="Gray", anchor="w")\
                .grid(row=self.row, column=column, columnspan=2, sticky="w")  # display date
            Entry(self.win.body, width=macadj(2, 1), textvariable=self.type_vars[i])\
                .grid(row=self.row+1, column=column, sticky="w")  # entry field for type
            column += 1
            Entry(self.win.body, width=macadj(5, 4), textvariable=self.time_vars[i])\
                .grid(row=self.row+1, column=column, sticky="w")  # entry field for time
            column += 1
            if column >= 21:  # if the row is full
                column = 0  # reset column position to begining
                self.row += 2  # and start a new row

    def build_bottom(self):
        for _ in range(3):
            self.row += 1
            Label(self.win.body, text="").grid(row=self.row)

    def status_report(self, updates):
        # msg = "{} Record{} Updated.".format(updates, Handler(updates).plurals())
        msg = "hello there"
        self.status_update.config(text="{}".format(msg))

    def buttons_frame(self):
        button = Button(self.win.buttons)
        button.config(text="Submit", width=macadj(20, 21),
                      command=lambda: self.apply(True))  # apply and do no return to main screen
        if sys.platform == "win32":
            button.config(anchor="w")
        button.pack(side=LEFT)

        button = Button(self.win.buttons)
        button.config(text="Apply", width=macadj(20, 21),
                      command=lambda: self.apply(False))  # apply and do no return to main screen
        if sys.platform == "win32":
            button.config(anchor="w")
        button.pack(side=LEFT)

        button = Button(self.win.buttons)
        button.config(text="Go Back", width=macadj(20, 21),
                      command=lambda: OtEquitability()
                      .create_from_refusals(self.win.topframe, self.enddate, self.station))
        if sys.platform == "win32":
            button.config(anchor="w")
        button.pack(side=LEFT)

        self.status_update = Label(self.win.buttons, text="", fg="red")
        self.status_update.pack(side=LEFT)

    def apply(self, home):
        # loop through all stringvars and check for errors
        for i in range(len(self.type_vars)):
            if not self.checktypes(i):  # check the refusal indicator
                return  # return if there is an error
            if not self.checktimes(i):  # check the refusal time
                return  # return if there is an error
        # if all checks pass - input/update/dalete the refusals database table
        for i in range(len(self.type_vars)):
            time_var = self.time_vars[i].get().strip()
            if not self.match_type(i) or not self.match_time(i):
                if self.displaydate[i] not in self.onrec_displaydate:  # if there is no record with that date
                    self.insert(i)
                if self.displaydate[i] in self.onrec_displaydate:  # if there is a record with that date
                    if not time_var:  # if the time is blank
                        self.delete(i)  # delete the record
                    else:  # if there is a time
                        self.update(i)  # update the record
        if home:  # return to the OT Preference screen
            OtEquitability().create_from_refusals(self.win.topframe, self.enddate, self.station)
        else:  # create a new object and recreate the window
            RefusalWin().create(self.win.topframe, self.carrier_name, self.startdate, self.enddate, self.station)

    def checktypes(self, i):
        type_var = self.type_vars[i].get().strip()
        time_var = self.time_vars[i].get().strip()
        if RefusalTypeChecker(type_var).is_empty():
            return True
        if not RefusalTypeChecker(type_var).is_one():
            messagebox.showerror("Refusal Tracking",
                                 "The Refusal indicator for {} must be only one character".format(self.displaydate[i]),
                                 parent=self.win.body)
            return False
        if not RefusalTypeChecker(type_var).is_letter():
            messagebox.showerror("Refusal Tracking",
                                 "The Refusal indicator for {} must be a letter".format(self.displaydate[i]),
                                 parent=self.win.body)
            return False
        if type_var and not time_var:
            messagebox.showerror("Refusal Tracking",
                                 "The refusal indicator for {} is not accompanied with a refusal time."
                                 .format(self.displaydate[i]),
                                 parent=self.win.body)
        return True

    def checktimes(self, i):
        time_var = self.time_vars[i].get().strip()
        if RingTimeChecker(time_var).check_for_zeros():  # if blank or zero, skip all other checks
            return True
        if not RingTimeChecker(time_var).check_numeric():
            text = "The Refusal time for {} must be a numeric value.".format(self.displaydate[i])
            messagebox.showerror("Refusal Tracking", text, parent=self.win.topframe)
            return False
        if not RingTimeChecker(time_var).over_24():
            text = "The Refusal time for {} must be less than 24.".format(self.displaydate[i])
            messagebox.showerror("Refusal Tracking", text, parent=self.win.topframe)
            return False
        if not RingTimeChecker(time_var).less_than_zero():
            text = "The Refusal time for {} must be greater than or equal to 0.".format(self.displaydate[i])
            messagebox.showerror("Refusal Tracking", text, parent=self.win.topframe)
            return False
        if not RingTimeChecker(time_var).count_decimals_place():
            text = "The Refusal time for {} must not have more than 2 decimal places.".format(self.displaydate[i])
            messagebox.showerror("Refusal Tracking", text, parent=self.win.topframe)
            return False
        return True

    def match_type(self, i):  # check if the newly inputed type matchs the type in the database
        type_var = self.type_vars[i].get().strip()  # the newly inputed type
        onrec = self.onrec_type[i]  # the type on record in the database
        if type_var == onrec:
            return True
        return False

    def match_time(self, i):  # check if the newly inputed time matchs the time in the database
        time_var = self.time_vars[i].get().strip()  # the newly inputed time
        onrec = self.onrec_time[i]  # the time on record in the database
        if time_var == onrec:
            return True
        return False

    def insert(self, i):  # insert a new record into the dbase
        type_var = self.type_vars[i].get().strip()
        time_var = Convert(self.time_vars[i].get().strip()).hundredths()
        sql = "INSERT INTO Refusals (refusal_date, carrier_name, refusal_type, refusal_time) " \
              "VALUES('%s', '%s', '%s', '%s')" % (self.ref_dates[i], self.carrier_name, type_var, time_var)
        commit(sql)

    def update(self, i):  # update an existing record in the dbase
        type_var = self.type_vars[i].get().strip()
        time_var = Convert(self.time_vars[i].get().strip()).hundredths()
        # "UPDATE informalc_grv SET grv_no = '%s' WHERE grv_no = '%s'" % (new_num.get().lower(), old_num)
        sql = "UPDATE Refusals SET refusal_type = '%s', refusal_time = '%s' WHERE refusal_date = '%s' " \
              "and carrier_name = '%s'" % (type_var, time_var, self.ref_dates[i], self.carrier_name)
        commit(sql)

    def delete(self, i):  # delete the record from the dbase
        sql = "DELETE FROM Refusals WHERE refusal_date = '%s' and carrier_name = '%s'" \
              % (self.ref_dates[i], self.carrier_name)
        commit(sql)


class OtDistribution:
    def __init__(self):
        self.frame = None
        self.win = None
        self.row = 0
        self.quartinvran_year = None  # StringVar for investigation range
        self.quartinvran_quarter = None
        self.quartinvran_station = None
        self.new_quartinvran_year = None
        self.new_quartinvran_quarter = None
        self.new_quartinvran_station = None
        self.stations_minus_outofstation = None
        self.carrierlist = []  # distinct list of carriers by station and quarter
        self.recset = []  # recset of otdl carriers
        self.eligible_carriers = []  # all carriers on otdl during quarter
        self.ineligible_carriers = []  # carriers with no otdl rec during quarter, but a rec in otdl prefs
        self.startdate = datetime(1, 1, 1)
        self.enddate = datetime(1, 1, 1)
        self.station = ""
        self.quarter = ""
        self.range = None
        self.list_option_otdl = None
        self.list_option_wal = None
        self.list_option_nl = None
        self.list_option_aux = None
        self.list_option_ptf = None
        self.list_option_array = []
        self.status_update = ""

    def create(self, frame):  # called from the main screen to build ot preferences screen
        self.frame = frame
        self.win = MakeWindow()
        self.win.create(self.frame)
        self.startup_stringvars()
        self.setup_listoption_stringvars()
        self.create_lower()

    def re_create(self, frame):  # called from the ot preferences screen when invran is changed.
        self.row = 0  # re initialize vars
        self.startdate = datetime(1, 1, 1)
        self.enddate = datetime(1, 1, 1)
        self.station = ""
        self.quarter = ""
        self.frame = frame  # define the frame
        self.win = MakeWindow()
        self.re_startup_stringvars()
        self.setup_listoption_stringvars()
        self.create_lower()

    def create_lower(self):
        self.get_quarter()
        self.get_stations_list()
        self.get_dates()  # get startdate, enddate and station
        self.win.create(self.frame)
        self.build_quarterinvran()
        self.investigation_status()
        self.build_range()
        self.build_list_options()
        self.buttons_frame()
        self.win.finish()

    def get_stations_list(self):  # get a list of stations for station optionmenu
        self.stations_minus_outofstation = projvar.list_of_stations[:]
        self.stations_minus_outofstation.remove("out of station")
        if len(self.stations_minus_outofstation) == 0:
            self.stations_minus_outofstation.append("undefined")

    def get_dates(self):  # find startdate, enddate and station
        year = int(self.quartinvran_year.get())
        startdate = (datetime(year, 1, 1), datetime(year, 4, 1), datetime(year, 7, 1), datetime(year, 10, 1))
        enddate = (datetime(year, 3, 31), datetime(year, 6, 30), datetime(year, 9, 30), datetime(year, 12, 31))
        self.startdate = startdate[int(self.quartinvran_quarter.get())-1]
        self.enddate = enddate[int(self.quartinvran_quarter.get())-1]
        if self.quartinvran_station.get() == "undefined":
            self.station = ""
        else:
            self.station = self.quartinvran_station.get()

    def build_quarterinvran(self):
        Label(self.win.body, text="Overtime Distribution", font=macadj("bold", "Helvetica 18"), anchor="w") \
            .grid(row=self.row, column=0, sticky="w", columnspan=20)
        self.row += 1
        Label(self.win.body, text="").grid(row=self.row, column=0)
        self.row += 1
        Label(self.win.body, text="QUARTERLY INVESTIGATION RANGE")\
            .grid(row=self.row, column=0, columnspan=20, sticky="w")
        self.row += 1
        Label(self.win.body, text=macadj("Year: ", "Year:"), fg="Gray", anchor="w")\
            .grid(row=self.row, column=0, sticky="w")
        Entry(self.win.body, width=macadj(5, 4), textvariable=self.quartinvran_year)\
            .grid(row=self.row, column=1, sticky="w")
        Label(self.win.body, text=macadj("Quarter: ", "Quarter:"), fg="Gray")\
            .grid(row=self.row, column=2, sticky="w")
        Entry(self.win.body, width=macadj(2, 1), textvariable=self.quartinvran_quarter)\
            .grid(row=self.row, column=3, sticky="w")
        Label(self.win.body, text=macadj("Station: ", "Station:"), fg="Gray")\
            .grid(row=self.row, column=4, sticky="w")
        om_station = OptionMenu(self.win.body, self.quartinvran_station, *self.stations_minus_outofstation)
        om_station.config(width=macadj(31, 23))
        om_station.grid(row=self.row, column=5, columnspan=4, sticky=W, padx=2)
        # set and reset buttons for investigation range
        Button(self.win.body, text="Set", width=macadj(5, 6), bg=macadj("green", "SystemButtonFace"),
               fg=macadj("white", "green"), command=lambda: self.set_invran()).grid(row=self.row, column=9, padx=2)
        Button(self.win.body, text="Reset", width=macadj(5, 6), bg=macadj("red", "SystemButtonFace"),
               fg=macadj("white", "red")).grid(row=self.row, column=10, padx=2)
        self.row += 1
        self.win.fill(self.row, 30)  # fill the bottom of the window for scrolling

    def investigation_status(self):  # provide message on status of investigation range
        Label(self.win.body, text="").grid(row=self.row, column=0)
        self.row += 1
        Label(self.win.body, text="WEEKLY INVESTIGATION RANGE") \
            .grid(row=self.row, column=0, columnspan=20, sticky="w")
        self.row += 1
        # Investigation date SET/NOT SET notification
        if projvar.invran_weekly_span is None:
            Label(self.win.body, text="Investigation date/range not set", fg="red") \
                .grid(row=self.row, column=0, columnspan=8, sticky="w")
        elif projvar.invran_weekly_span == 0:  # if the investigation range is one day
            f_date = projvar.invran_date.strftime("%a - %b %d, %Y")
            Label(self.win.body, text="Day Set: {} --> Pay Period: {}".format(f_date, projvar.pay_period), fg="red")\
                .grid(row=self.row, column=0, columnspan=8, sticky="w")
        else:
            # if the investigation range is weekly
            f_date = projvar.invran_date_week[0].strftime("%a - %b %d, %Y")
            end_f_date = projvar.invran_date_week[6].strftime("%a - %b %d, %Y")
            Label(self.win.body, text="{0} through {1} --> Pay Period: {2}"
                  .format(f_date, end_f_date, projvar.pay_period), fg="red")\
                .grid(row=self.row, column=0, columnspan=8, sticky="w")

    def build_range(self):
        self.row += 1
        Label(self.win.body, text="").grid(row=self.row, column=0)
        self.row += 1
        Label(self.win.body, text="Spread Sheet Range: ").grid(row=self.row, column=0, columnspan=8, sticky="w")
        self.row += 1
        self.range = StringVar(self.win.body)
        self.range.set('weekly')
        Radiobutton(self.win.body, text="Quarterly", variable=self.range, value='quarterly', justify=LEFT) \
            .grid(row=self.row, column=1, sticky=W, columnspan=3)
        self.row += 1
        Radiobutton(self.win.body, text="Weekly", variable=self.range, value='weekly', justify=LEFT) \
            .grid(row=self.row, column=1, sticky=W, columnspan=3)

    def build_list_options(self):
        self.row += 1
        Label(self.win.body, text="").grid(row=self.row, column=0)
        self.row += 1
        Label(self.win.body, text="List Options: ").grid(row=self.row, column=0, columnspan=8, sticky="w")
        self.row += 1
        Checkbutton(self.win.body, text="OTDL", variable=self.list_option_otdl, justify=LEFT) \
            .grid(row=self.row, column=1, sticky=W, columnspan=3)
        self.row += 1
        Checkbutton(self.win.body, text="Work Assignment", variable=self.list_option_wal, justify=LEFT) \
            .grid(row=self.row, column=1, sticky=W, columnspan=3)
        self.row += 1
        Checkbutton(self.win.body, text="No List", variable=self.list_option_nl, justify=LEFT) \
            .grid(row=self.row, column=1, sticky=W, columnspan=3)
        self.row += 1
        Checkbutton(self.win.body, text="Auxiliary", variable=self.list_option_aux, justify=LEFT) \
            .grid(row=self.row, column=1, sticky=W, columnspan=3)
        self.row += 1
        Checkbutton(self.win.body, text="Part Time Flex", variable=self.list_option_ptf, justify=LEFT) \
            .grid(row=self.row, column=1, sticky=W, columnspan=3)
        self.row += 1

    def set_invran(self):
        if not self.check_quarterinvran():
            return
        self.re_create(self.win.topframe)

    def check_quarterinvran(self):
        if not isint(self.quartinvran_year.get()):
            self.error_msg("The year must be a numeric.")
            return False
        if not len(self.quartinvran_year.get()) == 4:
            self.error_msg("Year must have four digits.")
            return False
        if not isint(self.quartinvran_quarter.get()):
            self.error_msg("The quarter must be an integer.")
            return False
        if int(self.quartinvran_quarter.get()) not in (1, 2, 3, 4):
            self.error_msg("Acceptable values for Quarter are limited to 1, 2, 3 or 4.")
            return False
        if self.quartinvran_station.get() == "undefined":
            self.error_msg("You must select a station to set the investigation range.")
            return False
        self.new_quartinvran_year = self.quartinvran_year.get()
        self.new_quartinvran_quarter = self.quartinvran_quarter.get()
        self.new_quartinvran_station = self.quartinvran_station.get()
        return True

    def startup_stringvars(self):
        if projvar.invran_weekly_span is None:  # if no investigation range is set
            date = datetime.now()
            station = "undefined"
        elif projvar.invran_weekly_span:  # if the investigation range is weekly
            date = projvar.invran_date_week[6]
            station = projvar.invran_station
        else:
            date = projvar.invran_date  # if the investigation range is daily
            station = projvar.invran_station
        year = date.strftime("%Y")
        month = date.strftime("%m")
        quarter = Quarter(month).find()  # get the quarter from the month
        self.quartinvran_year = StringVar(self.win.body)
        self.quartinvran_quarter = StringVar(self.win.body)
        self.quartinvran_station = StringVar(self.win.body)
        self.quartinvran_year.set(year)
        self.quartinvran_quarter.set(quarter)
        self.quartinvran_station.set(station)

    def re_startup_stringvars(self):
        self.quartinvran_year = StringVar(self.win.body)
        self.quartinvran_quarter = StringVar(self.win.body)
        self.quartinvran_station = StringVar(self.win.body)
        self.quartinvran_year.set(self.new_quartinvran_year)
        self.quartinvran_quarter.set(self.new_quartinvran_quarter)
        self.quartinvran_station.set(self.new_quartinvran_station)

    def setup_listoption_stringvars(self):
        self.list_option_otdl = IntVar(self.win.body)
        self.list_option_wal = IntVar(self.win.body)
        self.list_option_nl = IntVar(self.win.body)
        self.list_option_aux = IntVar(self.win.body)
        self.list_option_ptf = IntVar(self.win.body)
        self.list_option_otdl.set(0)
        self.list_option_wal.set(1)
        self.list_option_nl.set(1)
        self.list_option_aux.set(0)
        self.list_option_ptf.set(0)

    def get_quarter(self):  # creates quarter in format "2021-3"
        self.quarter = self.quartinvran_year.get() + "-" + self.quartinvran_quarter.get()

    def buttons_frame(self):
        button = Button(self.win.buttons)
        button.config(text="Go Back", width=macadj(18, 12),
                      command=lambda: MainFrame().start(frame=self.win.topframe))
        if sys.platform == "win32":
            button.config(anchor="w")
        button.pack(side=LEFT)
        # generate spreadsheet
        button = Button(self.win.buttons)
        button.config(text="SpreadSheet", width=macadj(17, 12),
                      command=lambda: (self.set_listoption_array(), OTDistriSpreadsheet().create
                      (self.win.topframe, self.startdate, self.quartinvran_station.get(), self.range.get(),
                       self.list_option_array)))
        if sys.platform == "win32":
            button.config(anchor="w")
        button.pack(side=LEFT)
        self.status_update = Label(self.win.buttons, text="", fg="red")
        self.status_update.pack(side=LEFT)

    def set_listoption_array(self):
        self.list_option_array = []
        options = ("otdl", "wal", "nl", "aux", "ptf")
        strvars = (self.list_option_otdl.get(), self.list_option_wal.get(), self.list_option_nl.get(),
                   self.list_option_aux.get(), self.list_option_ptf.get())
        for i in range(len(strvars)):
            if strvars[i]:
                self.list_option_array.append(options[i])


class OtEquitability:
    def __init__(self):
        self.frame = None
        self.win = None
        self.row = 0
        self.quartinvran_year = None  # StringVar for investigation range
        self.quartinvran_quarter = None
        self.quartinvran_station = None
        self.new_quartinvran_year = None  # place values in these when setting a new investigation range
        self.new_quartinvran_quarter = None
        self.new_quartinvran_station = None
        self.stations_minus_outofstation = None
        self.carrierlist = []
        self.recset = []
        self.startdate = datetime(1, 1, 1)
        self.enddate = datetime(1, 1, 1)
        self.station = ""
        self.quarter = ""
        self.pref_var = []  # build an array of stringvars for ot preference
        self.makeup_var = []  # build an array of stringvars for ot makeups
        self.onrec_prefs_carriers = []
        self.onrec_prefs = []
        self.onrec_makeups = []
        self.status_update = None
        self.delete_report = []  # list of ineligible carriers to be deleted from otdl prefence table
        self.eligible_carriers = []  # carriers on the otdl during the quarter from carriers table
        self.ineligible_carriers = []  # carriers with no otdl rec during quarter, but a rec in otdl prefs

    def create(self, frame):  # called from the main screen to build ot preferences screen
        self.frame = frame
        self.win = MakeWindow()
        self.startup_stringvars()
        self.create_lower()
        self.win.finish()

    def create_from_refusals(self, frame, enddate, station):
        self.frame = frame
        self.station = station
        self.win = MakeWindow()
        self.setup_stringvars_from_refusals(enddate, station)
        self.create_lower()
        self.win.finish()

    def re_create(self, frame):  # called from the ot preferences screen when invran is changed.
        self.row = 0  # re initialize vars
        self.carrierlist = []  # distinct list of carriers by station and quarter
        self.recset = []  # recset of otdl carriers
        self.eligible_carriers = []  # all carriers on otdl during quarter
        self.ineligible_carriers = []  # carriers with no otdl rec during quarter, but a rec in otdl prefs
        self.startdate = datetime(1, 1, 1)
        self.enddate = datetime(1, 1, 1)
        self.station = ""
        self.quarter = ""
        self.pref_var = []
        self.makeup_var = []
        self.onrec_prefs_carriers = []
        self.onrec_prefs = []
        self.onrec_makeups = []
        # self.status_update = None
        self.delete_report = []  # list of ineligible carriers to be deleted from otdl prefence table
        self.frame = frame  # define the frame
        self.win = MakeWindow()
        self.re_startup_stringvars()
        self.create_lower()
        self.win.finish()

    def create_lower(self):
        self.get_quarter()
        self.get_stations_list()
        self.win.create(self.frame)
        self.build_invran()
        self.get_dates()  # get startdate, enddate and station
        self.get_carrierlist()
        self.get_recsets()
        self.get_eligible_carriers()
        self.get_onrecs_set_stringvars()
        self.get_onrec_pref_carriers()
        self.get_ineligible()
        self.delete_ineligible()
        self.build_header()
        self.build_main()
        self.deletion_report()
        self.buttons_frame()

    def startup_stringvars(self):
        if projvar.invran_weekly_span is None:  # if no investigation range is set
            date = datetime.now()
            station = "undefined"
        elif projvar.invran_weekly_span:  # if the investigation range is weekly
            date = projvar.invran_date_week[6]
            station = projvar.invran_station
        else:
            date = projvar.invran_date  # if the investigation range is daily
            station = projvar.invran_station
        year = date.strftime("%Y")
        month = date.strftime("%m")
        quarter = Quarter(month).find()  # get the quarter from the month
        self.quartinvran_year = StringVar(self.win.body)
        self.quartinvran_quarter = StringVar(self.win.body)
        self.quartinvran_station = StringVar(self.win.body)
        self.quartinvran_year.set(year)
        self.quartinvran_quarter.set(quarter)
        self.quartinvran_station.set(station)

    def setup_stringvars_from_refusals(self, enddate, station):
        year = enddate.strftime("%Y")
        month = enddate.strftime("%m")
        quarter = Quarter(month).find()  # get the quarter from the month
        self.quartinvran_year = StringVar(self.win.body)
        self.quartinvran_quarter = StringVar(self.win.body)
        self.quartinvran_station = StringVar(self.win.body)
        self.quartinvran_year.set(year)
        self.quartinvran_quarter.set(quarter)
        self.quartinvran_station.set(station)

    def re_startup_stringvars(self):
        self.quartinvran_year = StringVar(self.win.body)
        self.quartinvran_quarter = StringVar(self.win.body)
        self.quartinvran_station = StringVar(self.win.body)
        self.quartinvran_year.set(self.new_quartinvran_year)
        self.quartinvran_quarter.set(self.new_quartinvran_quarter)
        self.quartinvran_station.set(self.new_quartinvran_station)

    def get_quarter(self):  # creates quarter in format "2021-3"
        self.quarter = self.quartinvran_year.get() + "-" + self.quartinvran_quarter.get()

    def get_stations_list(self):  # get a list of stations for station optionmenu
        self.stations_minus_outofstation = projvar.list_of_stations[:]
        self.stations_minus_outofstation.remove("out of station")
        if len(self.stations_minus_outofstation) == 0:
            self.stations_minus_outofstation.append("undefined")

    def get_dates(self):  # find startdate, enddate and station
        year = int(self.quartinvran_year.get())
        startdate = (datetime(year, 1, 1), datetime(year, 4, 1), datetime(year, 7, 1), datetime(year, 10, 1))
        enddate = (datetime(year, 3, 31), datetime(year, 6, 30), datetime(year, 9, 30), datetime(year, 12, 31))
        self.startdate = startdate[int(self.quartinvran_quarter.get())-1]
        self.enddate = enddate[int(self.quartinvran_quarter.get())-1]
        if self.quartinvran_station.get() == "undefined":
            self.station = ""
        else:
            self.station = self.quartinvran_station.get()

    def get_carrierlist(self):
        self.carrierlist = CarrierList(self.startdate, self.enddate, self.station).get_distinct()

    def get_recsets(self):
        for carrier in self.carrierlist:
            otlist = ("otdl", )
            rec = QuarterRecs(carrier[0], self.startdate, self.enddate, self.station).get_filtered_recs(otlist)
            if rec:
                self.recset.append(rec)

    def build_invran(self):
        Label(self.win.body, text="OTDL Preferences", font=macadj("bold", "Helvetica 18"), anchor="w") \
            .grid(row=self.row, column=0, sticky="w", columnspan=20)
        self.row += 1
        Label(self.win.body, text="").grid(row=self.row, column=0)
        self.row += 1
        Label(self.win.body, text="QUARTERLY INVESTIGATION RANGE").grid(row=self.row, column=0, columnspan=20,
                                                                        sticky="w")
        self.row += 1
        Label(self.win.body, text=macadj("Year: ", "Year:"), fg="Gray", anchor="w")\
            .grid(row=self.row, column=0, sticky="w")
        Entry(self.win.body, width=macadj(5, 4), textvariable=self.quartinvran_year)\
            .grid(row=self.row, column=1, sticky="w")
        Label(self.win.body, text=macadj("Quarter: ", "Quarter:"), fg="Gray")\
            .grid(row=self.row, column=2, sticky="w")
        Entry(self.win.body, width=macadj(2, 1), textvariable=self.quartinvran_quarter)\
            .grid(row=self.row, column=3, sticky="w")
        Label(self.win.body, text=macadj("Station: ", "Station:"), fg="Gray")\
            .grid(row=self.row, column=4, sticky="w")
        om_station = OptionMenu(self.win.body, self.quartinvran_station, *self.stations_minus_outofstation)
        om_station.config(width=macadj(31, 23))
        om_station.grid(row=self.row, column=5, columnspan=4, sticky=W, padx=2)
        # set and reset buttons for investigation range
        Button(self.win.body, text="Set", width=macadj(5, 6), bg=macadj("green", "SystemButtonFace"),
               fg=macadj("white", "green"), command=lambda: self.set_invran()).grid(row=self.row, column=9, padx=2)
        Button(self.win.body, text="Reset", width=macadj(5, 6), bg=macadj("red", "SystemButtonFace"),
               fg=macadj("white", "red")).grid(row=self.row, column=10, padx=2)
        self.row += 1
        self.win.fill(self.row, 30)  # fill the bottom of the window for scrolling

    def set_invran(self):
        if not self.check_quarterinvran():
            return
        self.re_create(self.win.topframe)

    def error_msg(self, text):
        messagebox.showerror("OTDL Preferences", text, parent=self.win.topframe)

    def check_quarterinvran(self):
        if not isint(self.quartinvran_year.get()):
            self.error_msg("The year must be a numeric.")
            return False
        if not len(self.quartinvran_year.get()) == 4:
            self.error_msg("Year must have four digits.")
            return False
        if not isint(self.quartinvran_quarter.get()):
            self.error_msg("The quarter must be an integer.")
            return False
        if int(self.quartinvran_quarter.get()) not in (1, 2, 3, 4):
            self.error_msg("Acceptable values for Quarter are limited to 1, 2, 3 or 4.")
            return False
        if self.quartinvran_station.get() == "undefined":
            self.error_msg("You must select a station to set the investigation range.")
            return False
        self.new_quartinvran_year = self.quartinvran_year.get()
        self.new_quartinvran_quarter = self.quartinvran_quarter.get()
        self.new_quartinvran_station = self.quartinvran_station.get()
        return True

    def get_status(self, recs):  # returns true if the carrier's last record is otdl and the station is correct.
        if recs[0][2] == "otdl" and recs[0][5] == self.station:
            return "on"
        return "off"

    @staticmethod
    def check_consistancy(recs):  # check that carriers on list have not gotten off then on again.
        off_list = False
        on_list = False
        for rec in reversed(recs):
            if off_list:
                if rec[2] == "otdl":
                    on_list = True
            if rec[2] != "otdl":
                off_list = True
        if off_list and on_list:
            return True
        return False

    def get_eligible_carriers(self):  # builds array of carriers on otdl at any point during quarter from carrier table
        for carrier in self.recset:
            self.eligible_carriers.append(carrier[0][1])

    def get_pref(self, carrier):  # pull otdl preferences from dbase - insert if there is no preference.
        sql = "SELECT preference FROM otdl_preference WHERE carrier_name = '%s' and quarter = '%s' and station = '%s'" \
              % (carrier, self.quarter, self.station)
        pref = inquire(sql)
        if not pref:
            sql = "INSERT INTO otdl_preference (quarter, carrier_name, preference, station, makeups) " \
                  "VALUES('%s', '%s', '%s', '%s', '%s')" \
                  % (self.quarter, carrier, "12", self.station, "")
            commit(sql)
            return ('12',)
        else:
            return pref[0]

    def get_makeups(self, carrier):  # pull makeups from the dbase
        sql = "SELECT makeups FROM otdl_preference WHERE carrier_name = '%s' and quarter = '%s' and station = '%s'" \
              % (carrier, self.quarter, self.station)
        makeups = inquire(sql)
        if not makeups:
            return 0
        return makeups[0]

    def get_onrecs_set_stringvars(self):
        i = 0
        for carrier in self.eligible_carriers:
            self.pref_var.append(StringVar(self.win.body))  # build array of string vars for otdl preferences
            self.makeup_var.append(StringVar(self.win.body))  # build array of string vars for make ups
            pref = self.get_pref(carrier)  # call method to inquire otdl preference table
            makeup = self.get_makeups(carrier)[0]  # call method to inquire otdl preference table
            makeup = Convert(makeup).empty_not_zero()  # use empty string instead of zero
            self.pref_var[i].set(pref[0])  # set the preference stringvar
            self.makeup_var[i].set(makeup)
            self.onrec_prefs.append(pref[0])  # build the array of otdl preferences from otdl preferences table.
            self.onrec_makeups.append(makeup)
            i += 1

    def get_onrec_pref_carriers(self):
        sql = "SELECT carrier_name FROM otdl_preference WHERE quarter = '%s'and station = '%s'" \
              % (self.quarter, self.station)
        pref = inquire(sql)
        for carrier in pref:
            self.onrec_prefs_carriers.append(carrier[0])

    def get_ineligible(self):
        for pref_carrier in self.onrec_prefs_carriers:
            if pref_carrier not in self.eligible_carriers:
                self.ineligible_carriers.append(pref_carrier)

    def delete_ineligible(self):
        for carrier in self.ineligible_carriers:
            sql = "DELETE FROM otdl_preference WHERE quarter = '%s' AND carrier_name = '%s' AND station = '%s'" \
                  % (self.quarter, carrier, self.station)
            commit(sql)
            self.delete_report.append(carrier)

    def deletion_report(self):
        if len(self.delete_report) > 0:
            deleted_list = ""
            for name in self.delete_report:
                deleted_list += "      " + name + "\n"
            msg = "The OTDL Preference records has been deleted for quarter {} for the following " \
                  "carriers:\n\n{}\nThis is a routine maintenance action.".format(self.quarter, deleted_list)
            messagebox.showinfo("OTDL Preferences", msg, parent=self.win.body)

    def carrier_report(self, recs, consistant):
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = "report_carrier_history" + "_" + stamp + ".txt"
        report = open(dir_path('report') + filename, "w")
        report.write("\nCarrier List Status History\n\n")
        report.write('   Showing all list status changes for {} during quarter {}\n\n'.format(recs[0][1], self.quarter))
        report.write('{:<16}{:<8}{:<25}\n'.format("Date Effective", "List", "Station"))
        report.write('---------------------------------------------\n')
        i = 1
        for line in recs:
            report.write('{:<16}{:<8}{:<25}\n'
                         .format(dt_converter(line[0]).strftime("%m/%d/%Y"), line[2], line[5]))
            if i % 3 == 0:
                report.write('---------------------------------------------\n')
            i += 1
        if consistant == "error":
            report.write('\n')
            report.write('>>>Consistency Error: \n'
                         'OTDL Carriers can not get back on the Over Time Desired List once they \n'
                         'have gotten off during the quarter. This will raise an  \"error\" \n'
                         'message in the Check column. If this is a mistake, edit the carrier\'s \n'
                         'status history. \n')
        report.close()
        if sys.platform == "win32":  # open the text document
            os.startfile(dir_path('report') + filename)
        if sys.platform == "linux":
            subprocess.call(["xdg-open", 'kb_sub/report/' + filename])
        if sys.platform == "darwin":
            subprocess.call(["open", dir_path('report') + filename])

    def build_header(self):
        Label(self.win.body, text="Name", fg="Gray").grid(row=self.row, column=1, sticky="w")
        Label(self.win.body, text="Preference", fg="Gray").grid(row=self.row, column=5, sticky="w")
        Label(self.win.body, text="Make up", fg="Gray").grid(row=self.row, column=6, sticky="w")
        Label(self.win.body, text="Status", fg="Gray").grid(row=self.row, column=7, sticky="w")
        Label(self.win.body, text="Check", fg="Gray").grid(row=self.row, column=8, sticky="w")
        Label(self.win.body, text="Report", fg="Gray").grid(row=self.row, column=9, sticky="w")
        Label(self.win.body, text="Refusal", fg="Gray").grid(row=self.row, column=10, sticky="w")
        self.row += 1

    def build_main(self):
        i = 0
        for carrier in self.recset:
            Label(self.win.body, text=i+1, anchor="w").grid(row=self.row, column=0, sticky="w")
            Label(self.win.body, text=carrier[0][1], anchor="w").grid(row=self.row, column=1, columnspan=4, sticky="w")
            om_pref = OptionMenu(self.win.body, self.pref_var[i], "12", "10", "track")
            om_pref.config(width=4)
            om_pref.grid(row=self.row, column=5, sticky="w")
            Entry(self.win.body, textvariable=self.makeup_var[i], width=macadj(8, 6), justify='right')\
                .grid(row=self.row, column=6, sticky="w")  # make ups entry field
            status = "on"
            fg = "black"
            if self.get_status(carrier) == "off":  # if there is an error, display in red
                status = "off"
                fg = "red"
            Label(self.win.body, text=status, anchor="w", fg=fg).grid(row=self.row, column=7, sticky="w")
            consistant = "ok"
            fg = "black"
            if self.check_consistancy(carrier):  # if there is an error, display in red
                consistant = "error"
                fg = "red"
            Label(self.win.body, text=consistant, fg=fg, anchor="w").grid(row=self.row, column=8, sticky="w")
            Button(self.win.body, text="report",
                   command=lambda car=carrier, con=consistant: self.carrier_report(car, con))\
                .grid(row=self.row, column=9, sticky="w")
            Button(self.win.body, text="refusals",
                   command=lambda car=carrier[0][1]: RefusalWin().create(self.win.topframe, car,
                                                                         self.startdate, self.enddate, self.station))\
                .grid(row=self.row, column=10, sticky="w")
            self.row += 1
            i += 1

    def check_all(self):
        for i in range(len(self.onrec_makeups)):
            if not self.check_each(i):
                return False
        return True

    def check_each(self, i):
        carrier = self.recset[i][0][1]
        makeup = self.makeup_var[i].get()  # call method to inquire otdl preference table
        if RingTimeChecker(makeup).check_for_zeros():
            return True
        if not RingTimeChecker(makeup).check_numeric():
            text = "The Make up value for {} must be a number.".format(carrier)
            self.error_msg(text)
            return False
        if not RingTimeChecker(makeup).over_5000():
            text = "The Make up value for {} must not exceed 5000.".format(carrier)
            self.error_msg(text)
            return False
        if not RingTimeChecker(makeup).less_than_zero():
            text = "The Make up value for {} must not be less than zero.".format(carrier)
            self.error_msg(text)
            return False
        if not RingTimeChecker(makeup).count_decimals_place():
            text = "The Make up value for {} can not have more than two decimal places.".format(carrier)
            self.error_msg(text)
            return False
        return True

    def apply(self, home):
        if not self.check_all():
            return
        updates = 0
        for i in range(len(self.onrec_prefs)):
            update = False
            if self.onrec_prefs[i] != self.pref_var[i].get():
                carrier = self.recset[i][0][1]
                sql = "UPDATE otdl_preference SET preference = '%s' WHERE carrier_name = '%s' AND quarter = '%s' " \
                      "AND station = '%s'" % (self.pref_var[i].get(), carrier, self.quarter, self.station)
                commit(sql)
                update = True
            if self.onrec_makeups[i] != self.makeup_var[i].get():
                carrier = self.recset[i][0][1]
                makeup = Convert(self.makeup_var[i].get()).empty_not_zero()
                makeup = Convert(makeup).empty_or_hunredths()
                sql = "UPDATE otdl_preference SET makeups = '%s' WHERE carrier_name = '%s' AND quarter = '%s' " \
                      "AND station = '%s'" % (makeup, carrier, self.quarter, self.station)
                commit(sql)
                update = True
            if update:
                updates += 1
        if home:
            MainFrame().start(self.win.topframe)
        else:
            self.status_report(updates)
            self.reset_onrecs_and_vars()

    def reset_onrecs_and_vars(self):
        for i in range(len(self.pref_var)):
            pref = self.pref_var[i].get()
            makeup = Convert(self.makeup_var[i].get()).empty_not_zero()
            makeup = Convert(makeup).empty_or_hunredths()
            self.onrec_prefs[i] = pref
            self.onrec_makeups[i] = makeup
            self.makeup_var[i].set(makeup)

    def status_report(self, updates):
        msg = "{} Record{} Updated.".format(updates, Handler(updates).plurals())
        self.status_update.config(text="{}".format(msg))

    def buttons_frame(self):
        button = Button(self.win.buttons)
        button.config(text="Submit", width=macadj(17, 12),
                      command=lambda: self.apply(True))  # apply and return to main screen
        if sys.platform == "win32":
            button.config(anchor="w")
        button.pack(side=LEFT)
        button = Button(self.win.buttons)
        button.config(text="Apply", width=macadj(18, 12),
                      command=lambda: self.apply(False))  # apply and no not return to main
        if sys.platform == "win32":
            button.config(anchor="w")
        button.pack(side=LEFT)
        button = Button(self.win.buttons)
        button.config(text="Go Back", width=macadj(18, 12),
                      command=lambda: MainFrame().start(frame=self.win.topframe))
        if sys.platform == "win32":
            button.config(anchor="w")
        button.pack(side=LEFT)
        # generate spreadsheet
        button = Button(self.win.buttons)
        button.config(text="SpreadSheet", width=macadj(17, 12),
                      command=lambda: OTEquitSpreadsheet()
                      .create(self.win.topframe, self.startdate, self.quartinvran_station.get()))
        if sys.platform == "win32":
            button.config(anchor="w")
        button.pack(side=LEFT)
        self.status_update = Label(self.win.buttons, text="", fg="red")
        self.status_update.pack(side=LEFT)


class SpeedConfigGui:
    def __init__(self, frame):
        self.frame = frame
        self.win = MakeWindow()
        self.ns_mode = StringVar(self.win.body)
        self.abc_breakdown = StringVar(self.win.body)  # create stringvars
        self.min_empid = StringVar(self.win.body)
        self.min_alpha = StringVar(self.win.body)
        self.min_abc = StringVar(self.win.body)
        self.status_update = Label(self.win.buttons, text="", fg="red")

    def create(self):
        self.win.create(self.frame)
        Label(self.win.body, text="SpeedSheet Configurations", font=macadj("bold", "Helvetica 18"), anchor="w") \
            .grid(row=0, sticky="w", columnspan=4)
        Label(self.win.body, text=" ").grid(row=1, column=0)
        self.set_stringvars()
        Label(self.win.body, text="NS Day Preferred Mode: ", width=macadj(40, 30), anchor="w") \
            .grid(row=3, column=0, ipady=5, sticky="w")
        ns_pref = OptionMenu(self.win.body, self.ns_mode, "rotating", "fixed")
        ns_pref.config(width=macadj(9, 9))
        if sys.platform == "win32":
            ns_pref.config(anchor="w")
        ns_pref.grid(row=3, column=1, columnspan=2, sticky="w", padx=4)
        Button(self.win.body, width=5, text="change",
               command=lambda: self.apply_ns_mode()).grid(row=3, column=3, padx=4)
        Label(self.win.body, text="Minimum rows for SpeedSheets", width=macadj(30, 30), anchor="w") \
            .grid(row=4, column=0, ipady=5, sticky="w")
        Label(self.win.body, text="Alphabetical Breakdown (multiple tabs)", width=macadj(40, 30), anchor="w") \
            .grid(row=5, column=0, ipady=5, sticky="w")
        opt_breakdown = OptionMenu(self.win.body, self.abc_breakdown, "True", "False")
        opt_breakdown.config(width=macadj(9, 9))
        if sys.platform == "win32":
            opt_breakdown.config(anchor="w")
        opt_breakdown.grid(row=5, column=1, columnspan=2, sticky="w", padx=4)
        Button(self.win.body, width=5, text="change",
               command=lambda: self.apply_abc_breakdown()).grid(row=5, column=3, padx=4)
        Label(self.win.body, text="Minimum rows for Employee ID tab", width=macadj(40, 30), anchor="w") \
            .grid(row=6, column=0, ipady=5, sticky="w")
        Entry(self.win.body, width=5, textvariable=self.min_empid).grid(row=6, column=1, padx=4)
        Button(self.win.body, width=5, text="change",
               command=lambda: self.apply_min_empid()).grid(row=6, column=2, padx=4)
        Button(self.win.body, width=5, text="info",
               command=lambda: self.info("min_spd_empid")) \
            .grid(row=6, column=3, padx=4)
        Label(self.win.body, text="Minimum rows for Alphabetically tab", width=macadj(40, 30), anchor="w") \
            .grid(row=7, column=0, ipady=5, sticky="w")
        Entry(self.win.body, width=5, textvariable=self.min_alpha).grid(row=7, column=1, padx=4)
        Button(self.win.body, width=5, text="change",
               command=lambda: self.apply_min_alpha()).grid(row=7, column=2, padx=4)
        Button(self.win.body, width=5, text="info",
               command=lambda: self.info("min_spd_alpha")) \
            .grid(row=7, column=3, padx=4)
        Label(self.win.body, text="Minimum rows for Alphabetical breakdown tabs", width=macadj(40, 35), anchor="w") \
            .grid(row=8, column=0, ipady=5, sticky="w")
        Entry(self.win.body, width=5, textvariable=self.min_abc).grid(row=8, column=1, padx=4)
        Button(self.win.body, width=5, text="change",
               command=lambda: self.apply_min_abc()) \
            .grid(row=8, column=2, padx=4)
        Button(self.win.body, width=5, text="info", command=lambda: self.info("min_spd_abc")) \
            .grid(row=8, column=3, padx=4)
        dash_line = "________________________________________________________________________________________"
        if sys.platform == "darwin":
            dash_line = "__________________________________________________________________"
        Label(self.win.body,
              text=dash_line, pady=5).grid(row=9, columnspan=5, sticky="w")
        Label(self.win.body, text="Restore Defaults").grid(row=10, column=0, ipady=5, sticky="w")
        Button(self.win.body, width=5, text="set",
               command=lambda: self.preset_default()).grid(row=10, column=3)
        Label(self.win.body, text="High Settings").grid(row=11, column=0, ipady=5, sticky="w")
        Button(self.win.body, width=5, text="set",
               command=lambda: self.preset_high()).grid(row=11, column=3)
        Label(self.win.body, text="Low Settings").grid(row=12, column=0, ipady=5, sticky="w")
        Button(self.win.body, width=5, text="set",
               command=lambda: self.preset_low()).grid(row=12, column=3)
        self.win.fill(11, 20)  # fill the bottom of the window for scrolling
        self.buttons_frame()

    def buttons_frame(self):
        button = Button(self.win.buttons)
        button.config(text="Go Back", width=20, command=lambda: MainFrame().start(frame=self.win.topframe))
        if sys.platform == "win32":
            button.config(anchor="w")
        button.pack(side=LEFT)
        self.status_update.pack(side=LEFT)
        self.win.finish()

    def apply_ns_mode(self):
        if self.ns_mode.get() == "rotating":
            value = True
        else:
            value = False
        msg = "NS Day Preferred Mode updated: {}".format(self.ns_mode.get())
        self.commit_to_base(value, "speedcell_ns_rotate_mode", msg)

    def apply_abc_breakdown(self):
        msg = "Alphabetical Breakdown (multiple tabs) updated: {}".format(self.abc_breakdown.get())
        self.commit_to_base(self.abc_breakdown.get(), "abc_breakdown", msg)

    def apply_min_empid(self):
        if self.check(self.min_empid.get()) is None:
            msg = "Minimum rows for Employee ID tab updated: {}".format(self.min_empid.get())
            self.commit_to_base(self.min_empid.get(), "min_spd_empid", msg)

    def apply_min_alpha(self):
        if self.check(self.min_alpha.get()) is None:
            msg = "Minimum rows for Alphabetically tab updated: {}".format(self.min_alpha.get())
            self.commit_to_base(self.min_alpha.get(), "min_spd_alpha", msg)

    def apply_min_abc(self):
        if self.check(self.min_abc.get()) is None:
            if self.check_abc(self.min_abc.get()) is None:
                msg = "Minimum rows for Alphabetical breakdown tabs updated: {}".format(self.min_abc.get())
                self.commit_to_base(self.min_abc.get(), "min_spd_abc", msg)

    def commit_to_base(self, value, setting, msg):
        sql = "UPDATE tolerances SET tolerance ='%s' WHERE category = '%s'" % \
              (value, setting)
        commit(sql)
        self.set_stringvars()
        self.status_update.config(text="{}".format(msg))

    def check(self, value):  # check values for minimum rows
        if not isint(value):
            text = "You must enter a number with no decimals. "
            messagebox.showerror("Tolerance value entry error",
                                 text,
                                 parent=self.win.topframe)
            return False
        if value.strip() == "":
            text = "You must enter a numeric value for tolerances"
            messagebox.showerror("Tolerance value entry error",
                                 text,
                                 parent=self.win.topframe)
            return False
        if float(value) < 5:
            text = "Values must be equal to or greater than five."
            messagebox.showerror("Tolerance value entry error",
                                 text,
                                 parent=self.win.topframe)

            return False
        if float(value) > 500:
            text = "You must enter a value less five hundred."
            messagebox.showerror("Tolerance value entry error",
                                 text,
                                 parent=self.win.topframe)
            return False

    def check_abc(self, value):
        if float(value) > 50:
            text = "You must enter a value less than fifty."
            messagebox.showerror("Tolerance value entry error",
                                 text,
                                 parent=self.win.topframe)
            return False

    def preset_default(self):
        empid = "50"
        alpha = "50"
        abc = "10"
        self.preset_to_base(self, empid, alpha, abc)
        self.status_update.config(text="Default Minimum Row Settings Restored")

    def preset_high(self):
        empid = "150"
        alpha = "150"
        abc = "40"
        self.preset_to_base(self, empid, alpha, abc)
        self.status_update.config(text="High Minimum Row Settings Enabled")

    def preset_low(self):
        empid = "10"
        alpha = "10"
        abc = "5"
        self.preset_to_base(self, empid, alpha, abc)
        self.status_update.config(text="Low Minimum Row Settings Enabled")

    @staticmethod
    def preset_to_base(self, empid, alpha, abc):
        #  abc breakdown is false in all cases
        sql = "UPDATE tolerances SET tolerance ='%s' WHERE category = '%s'" % ("False", "abc_breakdown")
        commit(sql)
        sql = "UPDATE tolerances SET tolerance ='%s' WHERE category = '%s'" % (empid, "min_spd_empid")
        commit(sql)
        sql = "UPDATE tolerances SET tolerance ='%s' WHERE category = '%s'" % (alpha, "min_spd_alpha")
        commit(sql)
        sql = "UPDATE tolerances SET tolerance ='%s' WHERE category = '%s'" % (abc, "min_spd_abc")
        commit(sql)
        self.set_stringvars()

    def set_stringvars(self):
        setting = SpeedSettings()  # retrieve settings from tolerance table in dbase
        if setting.speedcell_ns_rotate_mode:
            self.ns_mode.set("rotating")
        else:
            self.ns_mode.set("fixed")
        self.abc_breakdown.set(str(setting.abc_breakdown))  # convert to str, else you get a 0 or 1
        self.min_empid.set(setting.min_empid)
        self.min_alpha.set(setting.min_alpha)
        self.min_abc.set(setting.min_abc)

    def info(self, switch):
        text = ""
        if switch == "min_spd_empid":
            text = "Sets the minimum number of rows for the " \
                   "Employee Id tab of the All Inclusive Speedsheet. \n\n" \
                   "Enter a value between 5 and 500"
        if switch == "min_spd_alpha":
            text = "Sets the minimum number of rows for the " \
                   "Alphabetical tab of the All Inclusive Speedsheet. \n\n" \
                   "Enter a value between 5 and 500"
        if switch == "min_spd_abc":
            text = "Sets the minimum number of rows for the " \
                   "Alphabetical breakdown tabs of the All Inclusive Speedsheet. \n\n" \
                   "Enter a value between 5 and 50"
        messagebox.showinfo("SpeedSheet Minimum Rows", text, parent=self.win.topframe)


class SpeedLoadThread(Thread):  # use multithreading to load workbook while progress bar runs
    def __init__(self, path):
        Thread.__init__(self)
        self.path = path
        self.workbook = ""

    def run(self):
        global pb_flag  # this will signal when the thread has ended to end the progress bar
        wb = load_workbook(self.path)  # load xlsx doc with openpyxl
        self.workbook = wb
        pb_flag = False


class SpeedWorkBookGet:
    @staticmethod
    def get_filepath():
        if projvar.platform == "macapp" or projvar.platform == "winapp":
            return os.path.join(os.path.sep, os.path.expanduser("~"), 'Documents', 'klusterbox', 'speedsheets')
        else:
            return 'kb_sub/speedsheets'

    def get_file(self):
        path = self.get_filepath()
        file_path = filedialog.askopenfilename(initialdir=path, filetypes=[("Excel files", "*.xlsx")])
        if file_path[-5:].lower() == ".xlsx":
            return file_path
        elif file_path == "":
            return "no selection"
        else:
            return "invalid selection"

    def open_file(self, frame, interject):
        global pb_flag
        pb_flag = True
        file_path = self.get_file()
        if file_path == "no selection":
            return
        elif file_path == "invalid selection":
            messagebox.showerror("Report Generator",
                                 "The file you have selected is not an .xlsx file. "
                                 "You must select a file with a .xlsx extension.",
                                 parent=frame)
            return
        else:
            pb = ProgressBarIn(title="Klusterbox", label="SpeedSheeets Loading",
                               text="Loading and reading workbook. This could take a minute")
            wb = SpeedLoadThread(file_path)  # open workbook in separate thread
            wb.start()  # start loading workbook
            pb.start_up()  # start progress bar
            wb.join()  # wait for loading workbook to finish
            pb.stop()  # stop the progress bar and destroy the object
            SpeedSheetCheck(frame, wb.workbook, file_path, interject).check()  # check the speedsheet


class SpeedSheetCheck:
    def __init__(self, frame, wb, path, interject):
        self.frame = frame
        self.wb = wb
        self.path = path
        self.interject = interject  # True = add to database/ False = pre-check
        self.carrier_count = 0
        self.rings_count = 0
        self.fatal_rpt = 0
        self.fyi_rpt = 0
        self.add_rpt = 0
        self.rings_fatal_rpt = 0
        self.rings_fyi_rpt = 0
        self.rings_add_rpt = 0
        self.ns_xlate = {}
        self.ns_rotate_mode = True
        self.ns_true_rev = {}
        self.ns_false_rev = {}
        self.ns_custom = {}
        self.filename = ReportName("speedsheet_precheck").create()  # generate a name for the report
        self.report = open(dir_path('report') + self.filename, "w")  # open the report
        self.station = ""
        self.i_range = True  # investigation range is one week unless changed
        self.start_date = datetime(1, 1, 1, 0, 0, 0)
        self.end_date = datetime(1, 1, 1, 0, 0, 0)
        self.name = ""
        self.allowaddrecs = True
        self.name_mentioned = False
        self.pb = ProgressBarDe(title="Klusterbox", label="SpeedSheet Checking", text="Stand by...")
        self.sheets = []
        self.sheet_count = 0
        self.sheet_rowcount = []
        self.all_inclusive = True
        self.start_row = 6
        self.modulus = 8
        self.step = 2

    def check(self):
        try:
            date_array = [1, 1, 1]
            self.set_ns_preference()
            if self.ns_rotate_mode is not None and self.set_all_inclusive():
                self.set_sheet_facts()
                self.set_dates()
                self.set_ns_dictionaries()
                self.set_station()
                self.start_reporter()
                self.checking()
                self.reporter()
                date_array = Convert(self.start_date).datetime_separation()  # get the date to reset globals
                set_globals(date_array[0], date_array[1], date_array[2], self.i_range, self.station, self.frame)
            else:
                self.pb.delete()  # stop and destroy progress bar
                self.showerror()
        except KeyError:  # if wrong type of file is selected, there will be an error
            self.pb.delete()  # stop and destroy progress bar
            self.showerror()

    def set_ns_preference(self):  # are ns day preferences rotating or fixed?
        rotation = self.wb["by employee id"].cell(row=3, column=10).value  # get the ns day mode preference.
        if rotation.lower() not in ("r", "f"):
            self.ns_rotate_mode = None
        elif rotation == "r":
            self.ns_rotate_mode = True
        else:
            self.ns_rotate_mode = False

    def set_all_inclusive(self):
        # is the speedsheet all inclusive/ carrier only.
        all_in = self.wb["by employee id"].cell(row=1, column=1).value
        if all_in == "Speedsheet - All Inclusive Weekly":
            return True  # default settings from __init__ do not need changing
        elif all_in == "Speedsheet - All Inclusive Daily":
            self.step = 0
            self.modulus = 2
            return True
        elif all_in == "Speedsheet - Carriers":
            self.all_inclusive = False
            self.start_row = 5
            self.step = 0
            self.modulus = 1
            return True
        else:
            return False

    def set_sheet_facts(self):
        self.sheets = self.wb.sheetnames  # get the names of the worksheets as a list
        self.sheet_count = len(self.sheets)  # get the number of worksheets

    def set_dates(self):  # set the dates and the investigation range based on speedsheet input
        datecell = self.wb[self.sheets[0]].cell(row=2, column=2).value  # get the date or range of dates
        if len(datecell) < 12:  # if the investigation range is daily
            self.start_date = Convert(datecell).backslashdate_to_datetime()  # convert formatted date to datetime
            self.end_date = self.start_date  # since daily, dates are the same
            self.i_range = False  # change the range since it is daily
        else:  # if the investigation range is weekly
            d = datecell.split(" through ")  # split the date into two
            self.start_date = Convert(d[0]).backslashdate_to_datetime()  # convert formatted date to datetime
            self.end_date = Convert(d[1]).backslashdate_to_datetime()

    def set_ns_dictionaries(self):
        ns_obj = NsDayDict(self.start_date)  # get the ns day object
        self.ns_xlate = ns_obj.get()  # get ns day dictionary
        self.ns_true_rev = ns_obj.get_rev(True)  # get ns day dictionary for rotating days
        self.ns_false_rev = ns_obj.get_rev(False)  # get ns day dictionary for fixed days
        self.ns_custom = ns_obj.custom_config()  # shows custom ns day configurations for  printout / reports

    def set_station(self):
        self.station = self.wb[self.sheets[0]].cell(row=2, column=9).value  # get the station.

    def start_reporter(self):
        self.report.write("\nSpeedSheet Pre-Check Report \n")
        self.report.write(">>> {}\n".format(self.path))

    def row_count(self):  # get a count of all rows for all sheets - need for progress bar
        total_rows = 0
        for i in range(self.sheet_count):
            ws = self.wb[self.sheets[i]]  # assign the worksheet object
            row_count = ws.max_row  # get the total amount of rows in the worksheet
            self.sheet_rowcount.append(row_count)
            total_rows += row_count
        return total_rows

    def showerror(self):
        messagebox.showerror("Klusterbox SpeedSheets",
                             "SpeedSheets Precheck or Input has failed. \n"
                             "Either you have selected a spreadsheet that is not \n"
                             "a SpeedSheet or your Speedsheet is corrupted. \n"
                             "Suggestion: Verify that the file you are selecting \n "
                             "is a SpeedSheet. \n"
                             "Suggestion: Try re-generating the SpeedSheet.",
                             parent=self.frame)

    def checking(self):
        is_name = False  # initialize bool for speedcell name
        count_diff = self.sheet_count * self.start_row  # subtract top five/six rows from the row count
        self.pb.max_count(self.row_count() - count_diff)  # get total count of rows for the progress bar
        self.pb.start_up()  # start up the progress bar
        pb_counter = 0  # initialize the progress bar counter
        for i in range(self.sheet_count):
            ws = self.wb[self.sheets[i]]  # assign the worksheet object
            row_count = ws.max_row  # get the total amount of rows in the worksheet
            for ii in range(self.start_row, row_count):  # loop through all rows, start with row 5 or 6 until the end
                self.pb.move_count(pb_counter)
                if (ii + self.step) % self.modulus == 0:  # if the row is a carrier record
                    if ws.cell(row=ii, column=2).value is not None:  # if the carrier record has a carrier name
                        self.name_mentioned = False  # keeps names from being repeated in reports
                        self.carrier_count += 1  # get a count of the carriers for reports
                        is_name = True  # bool: the speedcell has a name
                        day = Handler(ws.cell(row=ii, column=1).value).nonetype()
                        name = Handler(ws.cell(row=ii, column=2).value).nonetype()
                        list_stat = Handler(ws.cell(row=ii, column=5).value).nonetype()
                        nsday = Handler(ws.cell(row=ii, column=6).value).ns_nonetype()
                        route = Handler(ws.cell(row=ii, column=7).value).nonetype()
                        empid = Handler(ws.cell(row=ii, column=10).value).nonetype()
                        self.name = name
                        self.pb.change_text("Reading Speedcell: {}".format(name))  # update text for progress bar
                        SpeedCarrierCheck(self, self.sheets[i], ii, name, day, list_stat, nsday, route,
                                          empid).check_all()
                    else:
                        is_name = False  # the speedcell does not have a name
                        self.pb.change_text("Detected empty Speedcell.")  # update text for progress bar
                else:
                    # if the speedcell has a name and passed carrier test, get the rings
                    if is_name and self.allowaddrecs:
                        self.rings_count += 1
                        # Handler().nonetype will convert any nonetypes to empty stings
                        day = Handler(ws.cell(row=ii, column=1).value).nonetype()
                        hours = Handler(ws.cell(row=ii, column=2).value).nonetype()
                        moves = Handler(ws.cell(row=ii, column=3).value).nonetype()
                        rs = Handler(ws.cell(row=ii, column=7).value).nonetype()
                        codes = Handler(ws.cell(row=ii, column=8).value).nonetype()
                        lv_type = Handler(ws.cell(row=ii, column=9).value).nonetype()
                        lv_time = Handler(ws.cell(row=ii, column=10).value).nonetype()
                        SpeedRingCheck(self, self.sheets[i], ii, day, hours, moves, rs, codes, lv_type, lv_time).check()
                pb_counter += 1
        self.pb.stop()

    def reporter(self):
        self.report.write("\n\n----------------------------------")
        # build report summary for carrier checks
        self.report.write("\n\nSpeedSheet Carrier Check Complete.\n\n")
        msg = "carrier{} checked".format(Handler(self.carrier_count).plurals())
        self.report.write('{:>6}  {:<40}\n'.format(self.carrier_count, msg))
        msg = "fatal error{} found".format(Handler(self.fatal_rpt).plurals())
        self.report.write('{:>6}  {:<40}\n'.format(self.fatal_rpt, msg))
        if self.interject:
            msg = "addition{} made".format(Handler(self.add_rpt).plurals())
            self.report.write('{:>6}  {:<40}\n'.format(self.add_rpt, msg))
        else:
            msg = "fyi notification{}".format(Handler(self.fyi_rpt).plurals())
            self.report.write('{:>6}  {:<40}\n'.format(self.fyi_rpt, msg))
        # build report summary for rings checks
        self.report.write("\n\nSpeedSheet Rings Check Complete.\n\n")
        msg = "ring{} checked".format(Handler(self.rings_count).plurals())
        self.report.write('{:>6}  {:<40}\n'.format(self.rings_count, msg))
        msg = "fatal error{} found".format(Handler(self.rings_fatal_rpt).plurals())
        self.report.write('{:>6}  {:<40}\n'.format(self.rings_fatal_rpt, msg))
        if self.interject:
            msg = "addition{} made".format(Handler(self.rings_add_rpt).plurals())
            self.report.write('{:>6}  {:<40}\n'.format(self.rings_add_rpt, msg))
        else:
            msg = "fyi notification{}".format(Handler(self.rings_fyi_rpt).plurals())
            self.report.write('{:>6}  {:<40}\n'.format(self.rings_fyi_rpt, msg))
        # close out the report and open in notepad
        self.report.close()
        if sys.platform == "win32":  # open the text document
            os.startfile(dir_path('report') + self.filename)
        if sys.platform == "linux":
            subprocess.call(["xdg-open", 'kb_sub/report/' + self.filename])
        if sys.platform == "darwin":
            subprocess.call(["open", dir_path('report') + self.filename])


class SpeedCarrierCheck:  # accepts carrier records from SpeedSheets
    def __init__(self, parent, sheet, row, name, day, list_stat, nsday, route, empid):
        self.parent = parent  # get objects from SpeedSheetCheck
        self.sheet = sheet  # input here is coming directly from the speedcell
        self.row = str(row)
        self.name = name  # get information passed from SpeedCell
        self.day = day
        self.list_stat = list_stat
        self.nsday = nsday.lower()
        self.route = route
        self.empid = empid
        self.tacs_name = ""  # get names and employee id numbers from name index
        self.kb_name = ""
        self.index_id = ""
        sql = "SELECT * FROM name_index WHERE kb_name = '%s'" % self.name  # access dbase to check emp id
        result = inquire(sql)
        if result:
            self.tacs_name = result[0][0]
            self.kb_name = result[0][1]
            self.index_id = result[0][2]
        self.filtered_recset = []
        self.onrec_date = ""  # get carrier information "on record" from the database
        self.onrec_name = ""
        self.onrec_list = ""
        self.onrec_nsday = ""
        self.onrec_route = ""
        self.addday = []  # checked input formatted for entry into database
        self.addlist = ["empty"]
        self.addnsday = "empty"
        self.addroute = "empty"
        self.addempid = ""
        self.parent.allowaddrecs = True  # if False, records will not be added to database
        self.error_array = []  # arrays for error, fyi and add reports
        self.fyi_array = []
        self.attn_array = []
        self.add_array = []
        self.ns_dict = \
            {"s": "sat", "m": "mon", "tu": "tue", "u": "tue", "w": "wed", "th": "thu", "h": "thu", "f": "fri",
             "fs": "sat", "fm": "mon", "ftu": "tue", "fu": "tue", "fw": "wed", "fth": "thu", "fh": "thu", "ff": "fri",
             "rs": "sat", "rm": "mon", "rtu": "tue", "ru": "tue", "rw": "wed", "rth": "thu", "rh": "thu", "rf": "fri",
             "sat": "sat", "mon": "mon", "tue": "tue", "wed": "wed", "thu": "thu", "fri": "fri",
             "rsat": "sat", "rmon": "mon", "rtue": "tue", "rwed": "wed", "rthu": "thu", "rfri": "fri",
             "fsat": "sat", "fmon": "mon", "ftue": "tue", "fwed": "wed", "fthu": "thu", "ffri": "fri"}

    def check_all(self):
        self.get_carrec()  # get carrier records and condense them into one array
        self.check_name()  # check for errors with the carrier name
        self.check_employee_id_format()
        self.check_employee_id_situation()
        self.check_employee_id_use()
        self.check_list_status()
        self.check_ns()
        self.check_route()
        if self.parent.interject:  # True = add to database/ False = pre-check
            self.add_recs()
        self.generate_report()

    def get_carrec(self):  # get carrier records and condense them into one array
        carrec = CarrierRecSet(self.name, self.parent.start_date, self.parent.end_date, self.parent.station).get()
        self.filtered_recset = CarrierRecFilter(carrec, self.parent.start_date).filter_nonlist_recs()
        carrec = CarrierRecFilter(self.filtered_recset, self.parent.start_date).condense_recs_ns()
        self.onrec_date = carrec[0]
        self.onrec_name = carrec[1]
        self.onrec_list = carrec[2]
        self.onrec_nsday = carrec[3]
        self.onrec_route = carrec[4]

    def check_name(self):  # check for errors with the carrier name
        if self.name == self.onrec_name:
            return
        if not NameChecker(self.name).check_characters():
            error = "     ERROR: Carrier name can not contain numbers or most special characters\n"
            self.error_array.append(error)
            self.parent.allowaddrecs = False  # do not allow this speedcell be be input into database
        if not NameChecker(self.name).check_length():
            error = "     ERROR: Carrier name must not exceed 42 characters\n"
            self.error_array.append(error)
            self.parent.allowaddrecs = False  # do not allow this speedcell be be input into database
        if not NameChecker(self.name).check_comma():
            error = "     ERROR: Carrier name must contain one comma to separate last name and first initial\n"
            self.error_array.append(error)
            self.parent.allowaddrecs = False  # do not allow this speedcell be be input into database
        if not NameChecker(self.name).check_initial():
            attn = "     ATTENTION: Carrier name should must contain one initial ideally, \n" \
                   "                unless more are needed to create a distinct carrier name.\n"
            self.attn_array.append(attn)
        # self.name = NameChecker(self.name).add_comma_spacing()  # make sure there is a space after the comma

    def check_employee_id_situation(self):
        if self.index_id == "" and self.empid == "":  # if both emp id and name index are blank
            pass
        elif self.index_id == self.empid:  # if the emp id from the name index and the speedsheet match
            pass
        elif self.index_id != "" and self.empid == "":  # if value in name index but spdcell is blank
            attn = "     ATTENTION: employee id can not be deleted from speedsheet\n"
            self.attn_array.append(attn)  # place this on "addition" report for user's information
            return
        elif self.index_id == "" and self.empid != "":  # if name index blank and spd cell has a value
            self.addempid = self.empid
            attn = "     ATTENTION: Possible new employee id\n"  # report
            self.attn_array.append(attn)
        else:
            error = "     ERROR: Employee id contridiction. \n" \
                    "            You can not change employee id with speedsheet\n"  # report
            self.error_array.append(error)
            self.parent.allowaddrecs = False  # do not allow this speedcell be be input into database

    def check_employee_id_format(self):  # verifies the employee id
        if self.empid == "":  # allow empty strings
            pass
        elif str(self.empid).isnumeric():  # allow integers and numeric strings
            self.empid = str(self.empid).zfill(8)  # change self.empid to string and zero fill to 8 places
            pass
        else:  # don't allow anything else
            error = "     ERROR: employee id is not numeric\n"  # report
            self.error_array.append(error)
            self.parent.allowaddrecs = False
            return

    def check_employee_id_use(self):  # make sure the employee id is not being used by another carrier
        kb_name = ""
        emp_id = ""
        if self.empid != "":
            sql = "SELECT * FROM name_index WHERE emp_id = '%s'" % self.empid
            result = inquire(sql)
            if result:
                kb_name = result[0][1]
                emp_id = result[0][2]
        if emp_id == "":
            return
        elif kb_name == self.name:
            pass
        else:
            error = "     ERROR: employee id is in use by another carrier\n"
            self.error_array.append(error)
            self.parent.allowaddrecs = False

    def add_list_status(self, dlsn_array, dlsn_day_array):
        if not self.filtered_recset:  # if the carrier is new
            self.addlist = dlsn_array
            self.addday = dlsn_day_array
            fyi = "     FYI: New List status will be entered: {}\n".format(dlsn_array)
            self.fyi_array.append(fyi)
        elif self.onrec_list != Convert(dlsn_array).array_to_string():  # if the list has changed
            self.addlist = dlsn_array
            self.addday = dlsn_day_array
            fyi = "     FYI: List status will be updated to: {}\n".format(dlsn_array)
            self.fyi_array.append(fyi)
        elif self.onrec_date != Convert(dlsn_day_array).array_to_string():  # if the days have changed
            self.addlist = dlsn_array
            self.addday = dlsn_day_array
            fyi = "     FYI: List status will be updated to: {}\n".format(dlsn_array)
            self.fyi_array.append(fyi)
        else:  # if there has been no change, do not change add___ vars.
            pass

    def check_list_status(self):
        self.list_stat = str(self.list_stat)
        self.list_stat = self.list_stat.strip()
        if self.list_stat == "":  # if the list_stat is empty
            self.add_list_status(["nl"], [])
            return
        dlsn_array = []  # dynamic list status notation array
        if self.list_stat != "":
            dlsn_array = Convert(self.list_stat).string_to_array()
        if len(dlsn_array) > 6:  # check number of list status changes
            error = "     ERROR: More than six changes in list status are not allowed\n"
            self.error_array.append(error)
            self.parent.allowaddrecs = False
            return
        for ls in dlsn_array:  # check for any input that does not conform with list status notation
            ls = ls.strip()  # strip any whitespace
            ls = ls.lower()  # make lowercase
            if ls in ("n", "w", "o", "a", "p", "c"):  # acceptable values
                pass
            elif ls in ("nl", "wal", "otdl", "odl", "aux", "cca", "ptf"):  # acceptable values
                pass
            else:
                error = "     ERROR: No such list status or list status notation {}\n".format(ls)
                self.error_array.append(error)
                self.parent.allowaddrecs = False
                return
        dlsn_array = self.dlsn_baseready(dlsn_array)  # format the list status/es for database
        # check days
        self.day = str(self.day)
        self.day = self.day.strip()
        dlsn_day_array = []
        if self.day != "":
            dlsn_day_array = Convert(self.day).string_to_array()
        if len(dlsn_day_array) > 7:
            error = "     ERROR: More than seven changes in days are not allowed\n"
            self.error_array.append(error)
            self.parent.allowaddrecs = False
        if len(dlsn_day_array) == 0 and len(dlsn_array) == 0:
            return
        elif len(dlsn_day_array) + 1 > len(dlsn_array):
            error = "     ERROR: Too many days compared to the list status {}\n" \
                    "            (hint: SpeedCell notation does not mention the \n" \
                    "            first day.) \n".format(self.day)
            self.error_array.append(error)
            self.parent.allowaddrecs = False
            return
        elif len(dlsn_day_array) + 1 < len(dlsn_array):
            error = "     ERROR: Too many list statuses compared to days {}\n" \
                    "            (SpeedCell notation requires that list status \n" \
                    "            changes be accompanied by the day of the change.) \n".format(self.day)
            self.error_array.append(error)
            self.parent.allowaddrecs = False
            return
        else:
            pass
        for d in dlsn_day_array:
            d = d.strip()  # strip any whitespace
            d = d.lower()  # make lowercase
            if d in ("s", "m", "tu", "u", "w", "th", "h", "f"):
                pass
            elif d in ("sat", "mon", "tue", "wed", "thu", "fri"):
                pass
            else:
                error = "     ERROR: No such day or day notation {}\n".format(d)
                self.error_array.append(error)
                self.parent.allowaddrecs = False
                return
        dlsn_day_array = self.day_baseready(dlsn_day_array)  # format the day/s for the database
        if self.check_day_sequence(dlsn_day_array) is False:  # check days for correct sequence
            return
        self.add_list_status(dlsn_array, dlsn_day_array)

    @staticmethod
    def dlsn_baseready(array):  # format dynamic list status notation into database ready
        new = []
        for ls in array:  # for each list status
            if ls in ("nl", "n"):
                new.append("nl")
            if ls in ("wal", "w"):
                new.append("wal")
            if ls in ("otdl", "odl", "o"):
                new.append("otdl")
            if ls in ("aux", "a", "cca", "c"):
                new.append("aux")
            if ls in ("ptf", "p"):
                new.append("ptf")
        return new

    def check_day_sequence(self, array):  # check the day/s for correct sequence
        sequence = ("sat", "sun", "mon", "tue", "wed", "thu", "fri")
        past = []
        for a in array:
            if a in past:
                error = "     ERROR: Days are out of sequence {}\n".format(self.day)
                self.error_array.append(error)
                self.parent.allowaddrecs = False
                return False
            for s in sequence:
                if s == a:
                    past.append(s)
                    break
                past.append(s)

    @staticmethod
    def day_baseready(array):  # format dynamic list status notation into database ready
        new = []
        for d in array:
            if d in ("sat", "s"):
                new.append("sat")
            if d in ("mon", "m"):
                new.append("mon")
            if d in ("tue", "tu", "u"):
                new.append("tue")
            if d in ("wed", "w"):
                new.append("wed")
            if d in ("thu", "th", "h"):
                new.append("thu")
            if d in ("fri", "f"):
                new.append("fri")
        return new

    def ns_baseready(self, ns, mode):  # formats provided ns day into a fixed or rotating ns day for database input
        baseready = self.parent.ns_true_rev[ns]  # if True is passed use rotate mode
        if not mode:  # if False is passed use fixed mode
            baseready = self.parent.ns_false_rev[ns]
        return baseready

    def add_ns(self, baseready):
        if self.onrec_nsday == baseready:
            pass  # keep value of addnsday var as "empty"
        else:
            fyi = "     FYI: New or updated nsday: {}.\n".format(self.parent.ns_custom[baseready])  # report
            self.fyi_array.append(fyi)
            self.addnsday = baseready

    def check_ns(self):
        #  self.parent.ns_rotate_mode: True for rotate, False for fixed
        ns = "none"  # initialize ns variable
        if not self.nsday:  # if string is empty
            self.add_ns(ns)  # ns day is "none"
        if self.nsday in ("sat", "mon", "tue", "wed", "thu", "fri"):
            baseready = self.ns_baseready(self.nsday, self.parent.ns_rotate_mode)  # format for dbase input
        elif self.nsday in ("s", "m", "tu", "u", "w", "th", "h", "f"):
            ns = self.ns_dict[self.nsday]  # translate the notation
            baseready = self.ns_baseready(ns, self.parent.ns_rotate_mode)
        elif self.nsday == "  ":  # if the string is almost empty
            baseready = ns  # ns day is "none"
        elif self.nsday in ("rsat", "rmon", "rtue", "rwed", "rthu", "rfri",
                            "rs", "rm", "rtu", "ru", "rw", "rth", "rh", "rf"):
            ns = self.ns_dict[self.nsday]  # use dictionary to get the day
            baseready = self.ns_baseready(ns, True)  # use ns rotate mode to get correct dictionary for day
        elif self.nsday in ("fsat", "fmon", "ftue", "fwed", "fthu", "ffri",
                            "fs", "fm", "ftu", "fu", "fw", "fth", "fh", "ff"):
            ns = self.ns_dict[self.nsday]
            baseready = self.ns_baseready(ns, False)  # use ns rotate mode to get correct dictionary for day
        else:
            error = "     ERROR: No such nsday: \"{}\"\n".format(self.nsday)  # report
            self.error_array.append(error)
            self.parent.allowaddrecs = False  # do not allow speedcell to be input into dbase
            return
        self.add_ns(baseready)

    def add_route(self):
        if self.route == self.onrec_route:
            pass  # retain "empty" value for addroute variable
        else:
            fyi = "     FYI: New or updated route: {}\n".format(self.route)
            self.fyi_array.append(fyi)
            self.addroute = self.route  # save to input to dbase

    def check_route(self):
        self.route = str(self.route)
        self.route = self.route.strip()
        if self.route == "":
            self.add_route()
        elif 4 > len(self.route) > 0:  # zero fill any inputs with between 0 and 4 digits
            self.route = self.route.zfill(4)
        if not RouteChecker(self.route).check_all():
            error = "     ERROR: Improper route formatting\n"  # report
            self.error_array.append(error)
            self.parent.allowaddrecs = False  # do not allow speedcell to be input into dbase
            return
        else:
            self.route = Handler(self.route).routes_adj()
            self.add_route()

    def add_recs(self):
        chg_these = []
        list_place = []
        ns_place = ""
        route_place = ""
        if not self.parent.allowaddrecs:  # if all checks passed
            return
        if self.addlist != ["empty"]:

            add = "     INPUT: List Status added or updated to database >>{}\n" \
                .format(Convert(self.addlist).array_to_string())  # report
            self.add_array.append(add)
            chg_these.append("list")
            list_place = self.addlist
        else:
            list_place = Convert(self.onrec_list).string_to_array()
        if self.addnsday != "empty":
            add = "     INPUT: Nsday added or updated to database >>{}\n".format(self.addnsday)  # report
            self.add_array.append(add)
            chg_these.append("ns")
            ns_place = self.addnsday
        else:
            ns_place = self.onrec_nsday
        if self.addroute != "empty":
            add = "     INPUT: Route added or updated to database >>{}\n".format(self.addroute)  # report
            self.add_array.append(add)
            chg_these.append("route")
            route_place = self.addroute
        else:
            route_place = self.onrec_route
        if self.addempid != "":
            sql = "INSERT INTO name_index (tacs_name, kb_name, emp_id) VALUES('%s', '%s', '%s')" \
                  % ("", self.name, str(self.empid).zfill(8))
            commit(sql)
            add = "     INPUT: Employee id added or updated to database >>{}\n".format(self.addempid)  # report
            self.add_array.append(add)
        # is the earliest car rec a Relevent Preceeding Record or a sat range:
        rpr = True  # Relevent Preceeding Record
        if self.filtered_recset:
            lastrec = self.filtered_recset.pop()  # get the earliest rec from rec set
            if lastrec[0] == str(self.parent.start_date):  # if last rec is the saturday in range
                rpr = False  # then there is no RPR
        if len(chg_these) != 0:  # build the first rec
            if rpr:  # insert the first rec
                sql = "INSERT INTO carriers(effective_date, carrier_name, list_status, ns_day, route_s, " \
                      "station) VALUES('%s','%s','%s','%s','%s','%s')" \
                      % (self.parent.start_date, self.name, list_place[0], ns_place, route_place, self.parent.station)
            else:  # update the first rec to replace pre existing record.
                sql = "UPDATE carriers SET list_status = '%s', ns_day = '%s', route_s = '%s', station = '%s'" \
                      "WHERE carrier_name = '%s' and effective_date = '%s'" \
                      % (list_place[0], ns_place, route_place, self.parent.station, self.name, self.parent.start_date)
            commit(sql)
        if self.addlist != ["empty"] and "list" in chg_these:
            second_date = self.parent.start_date + timedelta(days=1)
            seventh_date = self.parent.end_date  # delete all dates in service week except sat range
            sql = "DELETE FROM carriers WHERE carrier_name = '%s' and effective_date BETWEEN '%s' and '%s'" % \
                  (self.name, second_date, seventh_date)
            commit(sql)  # delete any records in investigation range except saturday
            for i in range(len(self.addlist)):
                if i == 0:
                    pass  # the first rec has already been entered
                else:
                    date = Convert(self.addday[i - 1]).day_to_datetime_str(self.parent.start_date)
                    sql = "INSERT INTO carriers(effective_date, carrier_name, list_status, ns_day, route_s, " \
                          "station) VALUES('%s','%s','%s','%s','%s','%s')" \
                          % (date, self.name, list_place[i], ns_place, route_place, self.parent.station)
                    commit(sql)

    def generate_report(self):  # generate a report
        self.parent.fatal_rpt += len(self.error_array)
        self.parent.add_rpt += len(self.add_array)
        self.parent.fyi_rpt += len(self.fyi_array)
        if not self.parent.interject:
            master_array = self.error_array + self.attn_array + self.fyi_array  # use these reports for precheck
        else:
            master_array = self.error_array + self.attn_array + self.add_array  # use these reports for input
        if len(master_array) > 0:
            if not self.parent.name_mentioned:
                self.parent.report.write("\n{}\n".format(self.name))
                self.parent.name_mentioned = True
            self.parent.report.write("   >>> sheet: \"{}\" --> row: \"{}\"  <<<\n".format(self.sheet, self.row))
            if not self.parent.allowaddrecs:
                self.parent.report.write("     SPEEDCELL ENTRY PROHIBITED: Correct errors!\n")
                # self.parent.fatal_rpt += 1
            for rpt in master_array:  # write all reports that have been keep in arrays.
                self.parent.report.write(rpt)


class SpeedRingCheck:  # accepts carrier rings from SpeedSheets
    def __init__(self, parent, sheet, row, day, hours, moves, rs, codes, lv_type, lv_time):
        self.parent = parent
        self.sheet = sheet
        self.row = row
        self.day = day
        self.hours = hours
        self.moves = moves
        self.rs = rs
        self.codes = codes
        self.lv_type = lv_type
        self.lv_time = lv_time
        self.allowaddrings = True
        self.error_array = []
        self.fyi_array = []
        self.attn_array = []
        self.add_array = []
        self.onrec_list = ""  # get carrier information "on record" from the database
        self.onrec_nsday = ""
        self.onrec_route = ""
        self.onrec_date = ""  # get rings information "on record" from the database
        self.onrec_name = ""
        self.onrec_5200 = ""
        self.onrec_rs = ""
        self.onrec_codes = ""
        self.onrec_moves = ""
        self.onrec_leave_type = ""
        self.onrec_leave_time = ""
        self.adddate = "empty"  # checked input formatted for entry into database
        self.add5200 = "empty"
        self.addrs = "empty"
        self.addcode = "empty"
        self.addmoves = "empty"
        self.addlvtype = "empty"
        self.addlvtime = "empty"

    def check(self):
        if self.check_day():  # if the day is a valid day
            self.get_onrecs()  # get existing "on record" records from the database
            self.check_5200()  # check 5200/ hours
            self.check_leave_time()  # check leave time
            if not self.check_empty():  # checks if the record should be deleted
                self.check_rs()   # check "return to station"
                self.check_codes()  # check the codes/notes
                self.check_leave_type()  # check leave type
                self.check_moves()  # check moves
                if self.parent.interject:  # if user wants to update database
                    self.add_recs()  # format and input rings into database
        self.generate_report()

    def get_day_as_datetime(self):  # get the datetime object for the day in use
        day = Convert(self.day).day_to_datetime_str(self.parent.start_date)
        self.adddate = day
        return day

    def get_onrecs(self):
        carrec = CarrierRecSet(self.parent.name, self.parent.start_date, self.parent.end_date,
                               self.parent.station).get()
        if carrec:
            self.onrec_list = carrec[0][2]  # get carrier information "on record" from the database
            self.onrec_nsday = carrec[0][3]
            self.onrec_route = carrec[0][4]
            ringrec = Rings(self.parent.name, self.get_day_as_datetime()).get_for_day()
            if ringrec[0]:  # if there is a result for clock rings on that day
                self.onrec_date = ringrec[0][0]  # get rings information "on record" from the database
                self.onrec_name = ringrec[0][1]
                self.onrec_5200 = ringrec[0][2]
                self.onrec_rs = ringrec[0][3]
                self.onrec_codes = ringrec[0][4]
                self.onrec_moves = ringrec[0][5]
                self.onrec_leave_type = ringrec[0][6]
                self.onrec_leave_time = ringrec[0][7]

    def check_day(self):
        days = ("sat", "sun", "mon", "tue", "wed", "thu", "fri")
        self.day = self.day.strip()
        self.day = str(self.day)
        self.day = self.day.lower()
        if self.day not in days:
            error = "     ERROR: Rings day is not correctly formatted. Acceptable values: sat, sun \n" \
                    "     mon, tue, wed, thu, or fri. Got instead \"{}\": \n".format(self.day)
            self.error_array.append(error)
            self.allowaddrings = False  # do not allow speedcell to be input into dbase
            return False
        return True

    def check_empty(self):
        # determine conditions where existing record is deleted
        if not self.hours:
            if not self.lv_time:
                if self.codes != "no call":
                    if self.onrec_date:  # if there is an existing record to delete
                        self.delete_recs()  # delete any pre existing record
                    return True
        return False

    def add_5200(self):
        if self.hours == "0.0" and self.onrec_5200 in ("0", "0.00", "0.0", "", 0, 0.0):
            pass
        elif self.hours != self.onrec_5200:  # compare 5200 time against 5200 from database,
            self.add5200 = self.hours  # if different, the add
            fyi = "     FYI: New or updated 5200 time: {}\n".format(self.hours)
            self.fyi_array.append(fyi)

    def check_5200(self):
        if type(self.hours) == str and not self.hours:  # pass if value is an empty string
            self.add_5200()
            return
        ring = RingTimeChecker(self.hours).make_float()  # returns float or False
        if ring is not False:
            self.hours = ring  # convert the item to a float, if not already
        else:  # if fail, create error msg and return
            error = "     ERROR: 5200 time must be a number. Got instead \"{}\": \n".format(self.hours)
            self.error_array.append(error)
            self.allowaddrings = False
            return
        if not RingTimeChecker(self.hours).over_24():
            error = "     ERROR: 5200 time can not exceed 24.00. Got instead \"{}\": \n".format(self.hours)
            self.error_array.append(error)
            self.allowaddrings = False
            return
        if not RingTimeChecker(self.hours).less_than_zero():
            error = "     ERROR: 5200 time can not be negative. Got instead \"{}\": \n".format(self.hours)
            self.error_array.append(error)
            self.allowaddrings = False
            return
        if not RingTimeChecker(self.hours).count_decimals_place():
            error = "     ERROR 5200 time can have no more than two decimal places. Got instead \"{}\": \n"\
                .format(self.hours)
            self.error_array.append(error)
            self.allowaddrings = False
            return
        self.hours = str(self.hours)  # convert float back to string
        self.hours = Convert(self.hours).hundredths()  # make number a string with 2 decimal places
        self.add_5200()

    def add_rs(self):
        if self.rs == "0.0" and self.onrec_rs in ("0", "0.00", "0.0", "", 0, 0.0):
            pass
        elif self.rs != self.onrec_rs:  # compare 5200 time against 5200 from database,
            self.addrs = self.rs  # if different, the add
            fyi = "     FYI: New or updated return to station: {}\n".format(self.rs)
            self.fyi_array.append(fyi)

    def check_rs(self):
        if type(self.rs) == str and not self.rs:  # pass if value is an empty string
            self.add_rs()
            return
        ring = RingTimeChecker(self.rs).make_float()  # returns float or False
        if ring is not False:
            self.rs = ring  # convert the attribute to a float, if not already
        else:  # if fail, create error msg and return
            error = "     ERROR: RS must be a number. Got instead \"{}\": \n".format(self.rs)
            self.error_array.append(error)
            self.allowaddrings = False
            return
        if not RingTimeChecker(self.rs).over_24():
            error = "     ERROR: RS time can not exceed 24.00. Got instead \"{}\": \n".format(self.rs)
            self.error_array.append(error)
            self.allowaddrings = False
            return
        if not RingTimeChecker(self.rs).less_than_zero():
            error = "     ERROR: RS time can not be negative. Got instead \"{}\": \n".format(self.rs)
            self.error_array.append(error)
            self.allowaddrings = False
            return
        if not RingTimeChecker(self.rs).count_decimals_place():
            error = "     ERROR: RS time can have no more than two decimal places. Got instead \"{}\": \n".\
                format(self.rs)
            self.error_array.append(error)
            self.allowaddrings = False
            return
        self.rs = str(self.rs)  # convert float back to string
        self.rs = Convert(self.rs).hundredths()  # make number a string with 2 decimal places
        self.add_rs()

    def add_moves(self, baseready):
        if baseready != self.onrec_moves:  # if the moves are different from on record moves from dbase,
            self.addmoves = baseready  # add the moves
            fyi = "     FYI: New or updated moves: {}\n".format(baseready)
            self.fyi_array.append(fyi)

    def check_moves(self):
        self.moves = str(self.moves)
        self.moves = self.moves.strip()
        if type(self.moves) == str and not self.moves:
            self.add_moves("")
            return
        self.moves = self.moves.replace("+", ",").replace("/", ",").replace("//", ",")\
            .replace("-", ",").replace("*", ",")  # replace all delimiters with commas
        moves_array = Convert(self.moves).string_to_array()  # convert the moves string to an array
        if not MovesChecker(moves_array).length():  # check number of items is multiple of three
            error = "     ERROR: Moves must be given in multiples of three. Got instead \"{}\": \n"\
                .format(len(moves_array))
            self.error_array.append(error)
            self.allowaddrings = False
            return
        for i in range(len(moves_array)):
            if i % 3 == 0 or (i + 2) % 3 == 0:  # check the time components of the moves triad
                move_ring = RingTimeChecker(moves_array[i]).make_float()  # try to convert moves_array[i] to a float.
                if move_ring is not False:  # if fail, create error msg and return
                    moves_array[i] = move_ring  # convert the item to a float, if not already
                else:
                    error = "     ERROR: Move times must be a number. Got instead \"{}\": \n".format(moves_array[i])
                    self.error_array.append(error)
                    self.allowaddrings = False
                    return
                if not RingTimeChecker(moves_array[i]).over_24():
                    error = "     ERROR: Move time can not exceed 24.00. Got instead \"{}\": ".format(moves_array[i])
                    self.error_array.append(error)
                    self.allowaddrings = False
                    return
                if not RingTimeChecker(moves_array[i]).less_than_zero():
                    error = "     ERROR: Move time can not be negative. Got instead \"{}\": \n".format(moves_array[i])
                    self.error_array.append(error)
                    self.allowaddrings = False
                    return
                if not RingTimeChecker(moves_array[i]).count_decimals_place():
                    error = "     ERROR: Move time can have no more than two decimal places. Got instead \"{}\": \n"\
                        .format(moves_array[i])
                    self.error_array.append(error)
                    self.allowaddrings = False
                    return
            if (i + 1) % 3 == 0:  # check the route component of the move triad
                if not RouteChecker(moves_array[i]).check_numeric():
                    error = "     ERROR: Routes in move triads must be numeric. Got instead \"{}\": \n" \
                        .format(moves_array[i])
                    self.error_array.append(error)
                    self.allowaddrings = False
                    return
                if not RouteChecker(moves_array[i]).check_length():
                    error = "     ERROR: Routes in move triads must have 4 or 5 digits. Got instead \"{}\": \n" \
                        .format(moves_array[i])
                    self.error_array.append(error)
                    self.allowaddrings = False
                    return
        for i in range(0, len(moves_array), 3):
            if moves_array[i] > moves_array[i + 1]:
                error = "     ERROR: first value \"{}\" must be lesser than the second \n" \
                        "            value \"{}\" in moves.\n".format(moves_array[i], moves_array[i + 1])
                self.error_array.append(error)
                self.allowaddrings = False
                return
            else:  # convert the items back into strings with 2 decimal places
                moves_array[i] = str(moves_array[i])
                moves_array[i] = Convert(moves_array[i]).hundredths()
                moves_array[i + 1] = str(moves_array[i + 1])
                moves_array[i + 1] = Convert(moves_array[i + 1]).hundredths()

        baseready = Convert(moves_array).array_to_string()  # convert the moves array to a baseready string
        self.add_moves(baseready)

    def add_codes(self):
        if self.codes == self.onrec_codes:  # compare 5200 time against 5200 from database,
            pass
        else:
            self.addcode = self.codes  # if different, the add
            fyi = "     FYI: New or updated code/note: {}\n".format(self.codes)
            self.fyi_array.append(fyi)

    def check_codes(self):
        all_codes = ("none", "ns day", "no call", "light", "sch chg", "annual", "sick", "excused")
        self.codes = self.codes.strip()
        self.codes = str(self.codes)
        self.codes = self.codes.lower()
        if not self.codes:
            self.codes = "none"
            self.add_codes()
            return
        if self.codes not in all_codes:
            error = "     ERROR: There is no such code/note. Got instead: \"{}\" \n" \
                .format(self.codes)
            self.error_array.append(error)
            self.allowaddrings = False
            return
        if self.onrec_list in ("nl", "wal"):
            if self.codes in ("no call", "light", "sch chg", "annual", "sick", "excused"):
                attn = "     ATTENTION: The code/note you entered is not consistant with the list status \n" \
                       "                for the day. Only \"none\" and \"ns day\" are useful for {} carriers. \n" \
                       "                Got instead: {}\n".format(self.onrec_list, self.codes)  # report
                self.attn_array.append(attn)
        # deleted otdl from list below. as of version 4.003 otdl carrier are allowed the ns day code.
        if self.onrec_list in ("aux", "ptf"):
            if self.codes in ("ns day",):
                attn = "     ATTENTION: The code/note you entered is not consistant with the list status \n" \
                       "                for the day. Only \"none\", \"no call\", \"light\", \"sch chg\", \n" \
                       "                \"annual\", \"sick\", \"excused\" are useful for {} carriers. \n" \
                       "                Got instead: {}\n".format(self.onrec_list, self.codes)
                self.attn_array.append(attn)
        self.add_codes()

    def add_lvtype(self):  # store the leave type if it has changed and passes checks
        if self.lv_type == self.onrec_leave_type:  # compare 5200 time against 5200 from database,
            pass  # take no action if they are the same
        else:
            self.addlvtype = self.lv_type  # if different, the add
            fyi = "     FYI: New or updated leave type: {}\n".format(self.lv_type)
            self.fyi_array.append(fyi)

    def check_leave_type(self):  # check the leave type
        all_codes = ("none", "annual", "sick", "holiday", "other", "combo")
        self.lv_type = str(self.lv_type)  # make sure lv type is a string
        self.lv_type = self.lv_type.strip()  # remove whitespace
        self.lv_type = self.lv_type.lower()  # force lv type to be lowercase
        if not self.lv_type:
            self.lv_type = "none"
            self.add_lvtype()  # store the leave type if it has changed and passes checks
            return
        if self.lv_type not in all_codes:
            error = "     ERROR: There is no such leave type. Acceptable types are: \"none\", \n" \
                    "            \"annual\", \"sick\", \"holiday\", \"other\" \n" \
                    "            Got instead: \"{}\"\n".format(self.lv_type)
            self.error_array.append(error)
            self.allowaddrings = False
            return
        self.add_lvtype()  # store the leave type if it has changed and passes checks

    def add_leave_time(self):
        if self.lv_time == "0.0" and self.onrec_leave_time in ("0", "0.00", "0.0", "", 0, 0.0):
            pass  # if new and old lv times are both empty, take no action
        elif self.lv_time != self.onrec_leave_time:  # compare lv type time against lv type from database,
            self.addlvtime = self.lv_time  # if different, the add
            fyi = "     FYI: New or updated leave time: {}\n".format(self.lv_time)
            self.fyi_array.append(fyi)

    def check_leave_time(self):
        if type(self.lv_time) == str and not self.lv_time:  # pass if value is an empty string
            self.add_leave_time()
            return
        ring = RingTimeChecker(self.lv_time).make_float()  # try to convert moves_array[i] to a float.
        if ring is not False:  # if fail, create error msg and return
            self.lv_time = ring  # convert the item to a float, if not already
        else:
            error = "     ERROR: Leave time must be a number. Got instead \"{}\": \n".format(self.lv_time)
            self.error_array.append(error)
            self.allowaddrings = False
            return
        if not RingTimeChecker(self.lv_time).over_8():
            error = "     ERROR: Leave time can not exceed 8.00. Got instead \"{}\": \n".format(self.lv_time)
            self.error_array.append(error)
            self.allowaddrings = False
            return
        if not RingTimeChecker(self.lv_time).less_than_zero():
            error = "     ERROR: Leave time can not be negative. Got instead \"{}\": \n".format(self.lv_time)
            self.error_array.append(error)
            self.allowaddrings = False
            return
        if not RingTimeChecker(self.lv_time).count_decimals_place():
            error = "     ERROR: Leave time can have no more than two decimal places. Got instead \"{}\": \n" \
                .format(self.lv_time)
            self.error_array.append(error)
            self.allowaddrings = False
            return
        self.lv_time = str(self.lv_time)  # make lv time back into a string
        self.lv_time = Convert(self.lv_time).hundredths()  # make lv time into a string number with 2 decimal places
        self.add_leave_time()

    def delete_recs(self):  # delete any pre existing record
        if not self.parent.interject:
            fyi = "     FYI: Clock Rings record will be deleted from database\n"
            self.fyi_array.append(fyi)
            return
        sql = "DELETE FROM rings3 WHERE rings_date = '%s' and carrier_name = '%s'" % (self.adddate, self.parent.name)
        commit(sql)
        add = "     DELETE: Clock Rings record deleted from database\n"  # report
        self.add_array.append(add)

    def add_recs(self):
        chg_these = []
        hours_place = ""
        rs_place = ""
        code_place = ""
        moves_place = ""
        lv_type_place = ""
        lv_time_place = ""
        if not self.allowaddrings:
            return
        # determine conditions where existing record is deleted
        if not self.hours:
            if not self.lv_time:
                if self.codes != "no call":
                    if self.onrec_date:  # if there is an existing record to delete
                        self.delete_recs()  # delete any pre existing record
                        return
        # contruct the sql command to commit to the database.
        if self.add5200 != "empty":  # 5200 place of sql command
            add = "     INPUT: 5200 time added or updated to database >>{}\n".format(self.add5200)  # report
            self.add_array.append(add)
            chg_these.append("hours")
            hours_place = self.add5200
        else:
            hours_place = self.onrec_5200
        if self.addrs != "empty":  # rs place of sql command
            add = "     INPUT: RS time added or updated to database >>{}\n".format(self.addrs)  # report
            self.add_array.append(add)
            chg_these.append("rs")
            rs_place = self.addrs
        else:
            rs_place = self.onrec_rs
        if self.addcode != "empty":  # code place of sql command
            add = "     INPUT: Code/note added or updated to database >>{}\n".format(self.addcode)  # report
            self.add_array.append(add)
            chg_these.append("code")
            code_place = self.addcode
        else:
            code_place = self.onrec_codes
        if self.addmoves != "empty":  # moves place of sql command
            add = "     INPUT: Moves added or updated to database >>{}\n".format(self.addmoves)  # report
            self.add_array.append(add)
            chg_these.append("moves")
            moves_place = self.addmoves
        else:
            moves_place = self.onrec_moves
        if self.addlvtype != "empty":  # lv type place of sql command
            add = "     INPUT: Leave type added or updated to database >>{}\n".format(self.addlvtype)  # report
            self.add_array.append(add)
            chg_these.append("lv type")
            lv_type_place = self.addlvtype
        else:
            lv_type_place = self.onrec_leave_type
        if self.addlvtime != "empty":  # lv time place of sql command
            add = "     INPUT: Leave time added or updated to database >>{}\n".format(self.addlvtime)  # report
            self.add_array.append(add)
            chg_these.append("lv time")
            lv_time_place = self.addlvtime
        else:
            lv_time_place = self.onrec_leave_time
        # if there are items to change, construct the sql command
        if chg_these:
            if not self.onrec_date:  # if there is no rings record for the date
                sql = "INSERT INTO rings3(rings_date, carrier_name, total, rs, code, " \
                      "moves, leave_type, leave_time) VALUES('%s','%s','%s','%s','%s','%s','%s','%s')" \
                      % (self.adddate, self.parent.name, hours_place, rs_place, code_place, moves_place,
                         lv_type_place, lv_time_place)
            else:  # if a record already exist
                sql = "UPDATE rings3 SET total = '%s', rs = '%s', code = '%s', moves = '%s', leave_type = '%s', " \
                      "leave_time = '%s' WHERE rings_date = '%s' and carrier_name = '%s'" % (hours_place,
                      rs_place, code_place, moves_place, lv_type_place, lv_time_place, self.adddate, self.parent.name)
            commit(sql)

    def generate_report(self):  # generate a report
        self.parent.rings_fatal_rpt += len(self.error_array)
        self.parent.rings_add_rpt += len(self.add_array)
        self.parent.rings_fyi_rpt += len(self.fyi_array)
        if not self.parent.interject:
            master_array = self.error_array + self.attn_array + self.fyi_array  # use these reports for precheck
        else:
            master_array = self.error_array + self.attn_array + self.add_array  # use these reports for input
        if len(master_array) > 0:
            if not self.parent.name_mentioned:
                self.parent.report.write("\n{}\n".format(self.parent.name))
                self.parent.name_mentioned = True
            self.parent.report.write("   >>> sheet: \"{}\" --> row: \"{}\" <<<\n".format(self.sheet, self.row))
            if not self.allowaddrings:
                self.parent.report.write("     CLOCK RINGS ENTRY PROHIBITED: Correct errors!\n")
                # self.parent.fatal_rpt += 1
            for rpt in master_array:  # write all reports that have been keep in arrays.
                self.parent.report.write(rpt)


class GuiConfig:
    def __init__(self, frame):
        self.frame = frame
        self.win = MakeWindow()
        self.wheel_selection = StringVar(self.win.body)
        self.invran_mode = StringVar(self.win.body)
        self.ot_rings_limiter = StringVar(self.win.body)
        self.status_update = Label(self.win.buttons, text="", fg="red")
        self.rings_limiter = None
        self.invran_result = None
        self.row = 0

    def create(self):
        self.get_settings()
        self.build()
        self.button_frame()
        self.win.fill(self.row + 1, 25)
        self.win.finish()

    def get_settings(self):
        sql = "SELECT tolerance FROM tolerances WHERE category = '%s'" % "mousewheel"
        results = inquire(sql)
        projvar.mousewheel = int(results[0][0])
        sql = "SELECT tolerance FROM tolerances WHERE category = '%s'" % "invran_mode"
        results = inquire(sql)
        self.invran_result = results[0][0]
        sql = "SELECT tolerance FROM tolerances WHERE category = '%s'" % "ot_rings_limiter"
        results = inquire(sql)
        rings_limiter = results[0][0]
        self.rings_limiter = Convert(rings_limiter).bool_to_onoff()  # convert the bool to on or off

    def build(self):
        self.win.create(self.frame)
        Label(self.win.body, text="GUI Configuration", font=macadj("bold", "Helvetica 18"), anchor="w") \
            .grid(row=self.row, sticky="w", columnspan=4)
        self.row += 1
        Label(self.win.body, text=" ").grid(row=self.row, column=0)
        self.row += 1
        # mousewheel scrolling direction
        Label(self.win.body, text="Mouse Wheel Scrolling:  ", anchor="w").grid(row=self.row, column=0, sticky="w")
        om_wheel = OptionMenu(self.win.body, self.wheel_selection, "natural", "reverse")  # option menu configuration
        om_wheel.config(width=7)
        om_wheel.grid(row=self.row, column=1)
        if projvar.mousewheel == 1:
            self.wheel_selection.set("natural")
        else:
            self.wheel_selection.set("reverse")
        Button(self.win.body, text="set", width=7, command=lambda: self.apply_mousewheel()).grid(row=self.row, column=2)
        self.row += 1
        Label(self.win.body, text=" ").grid(row=self.row, column=0)
        self.row += 1
        # investigation range mode
        Label(self.win.body, text="Investigation Range Mode:  ", anchor="w").grid(row=self.row, column=0, sticky="w")
        om_invran_mode = OptionMenu(self.win.body, self.invran_mode, "original", "simple", "no labels")
        om_invran_mode.config(width=7)
        om_invran_mode.grid(row=self.row, column=1)
        self.invran_mode.set(self.invran_result)
        Button(self.win.body, text="set", width=7,
               command=lambda: self.apply_invran_mode()).grid(row=self.row, column=2)
        self.row += 1
        Label(self.win.body, text=" ").grid(row=self.row, column=0)
        self.row += 1
        # overtime rings limiter
        Label(self.win.body, text="Overtime Rings Limiter:  ", anchor="w").grid(row=self.row, column=0, sticky="w")
        om_rings = OptionMenu(self.win.body, self.ot_rings_limiter, "on", "off")  # option menu configuration below
        om_rings.config(width=7)
        om_rings.grid(row=self.row, column=1)
        self.ot_rings_limiter.set(self.rings_limiter)
        Button(self.win.body, text="set", width=7,
               command=lambda: self.apply_rings_limiter()).grid(row=self.row, column=2)

    def button_frame(self):  # Display buttons and status update message
        button = Button(self.win.buttons)
        button.config(text="Go Back", width=20, command=lambda: MainFrame().start(frame=self.win.topframe))
        if sys.platform == "win32":
            button.config(anchor="w")
        button.pack(side=LEFT)
        self.status_update.pack(side=LEFT)

    def apply_rings_limiter(self):
        if self.ot_rings_limiter.get() == "on":
            rings_limiter = int(1)
        else:
            rings_limiter = int(0)
        sql = "UPDATE tolerances SET tolerance='%s'WHERE category='%s'" % (rings_limiter, "ot_rings_limiter")
        commit(sql)
        msg = "Overtime Rings Limiter updated: {}".format(self.ot_rings_limiter.get())
        self.status_update.config(text="{}".format(msg))

    def apply_invran_mode(self):
        sql = "UPDATE tolerances SET tolerance='%s'WHERE category='%s'" % (self.invran_mode.get(), "invran_mode")
        commit(sql)
        msg = "Investigation Range mode updated: {}".format(self.invran_mode.get())
        self.status_update.config(text="{}".format(msg))

    def apply_mousewheel(self):
        if self.wheel_selection.get() == "natural":
            wheel_multiple = int(1)
            projvar.mousewheel = int(1)  # sets the project variable
        else:  # if the self.wheel_selection.get() == "reverse"
            wheel_multiple = int(-1)
            projvar.mousewheel = int(-1)  # sets the project variable
        sql = "UPDATE tolerances SET tolerance='%s'WHERE category='%s'" % (wheel_multiple, "mousewheel")
        commit(sql)
        msg = "Mousescroll direction updated: {}".format(self.wheel_selection.get())
        self.status_update.config(text="{}".format(msg))


def database_rings_report(frame, station):
    #  generate a report summary of all clock rings for the station
    gross_dates = []  # captures all dates of rings for given station
    master_dates = []  # a distinct collection of dates for given station
    unique_dates = []
    sql = "SELECT DISTINCT carrier_name FROM carriers WHERE station = '%s' ORDER BY carrier_name" \
          % station
    results = inquire(sql)
    for name in results:
        active_station = []
        # get all records for the carrier
        sql = "SELECT * FROM carriers WHERE carrier_name= '%s' ORDER BY effective_date" % name[0]
        result_1 = inquire(sql)
        start_search = True
        start = ''
        end = ''
        # build the active_station array - find dates where carrier entered/left station
        for r in result_1:
            if r[5] == station and start_search:
                start = r
                start_search = False
            if r[5] != station and not start_search:
                end = r
                active_station.append([start, end])
                start = ''
                end = ''
                start_search = True
        if not start_search:
            active_station.append([start, end])
        for active in active_station:
            if active[1] != '':
                sql = "SELECT rings_date FROM rings3 WHERE rings_date " \
                      "BETWEEN '%s' AND '%s' AND carrier_name = '%s' " \
                      % (active[0][0], active[1][0], name[0])
                the_dates = inquire(sql)
                for td in the_dates:
                    gross_dates.append(td[0])
            else:
                sql = "SELECT rings_date FROM rings3 WHERE rings_date >= '%s' AND carrier_name = '%s' " \
                      % (active[0][0], name[0])
                the_dates = inquire(sql)
                for td in the_dates:
                    gross_dates.append(td[0])
    for gd in gross_dates:  # get a list of unique dates
        if gd not in unique_dates:
            unique_dates.append(gd)
    unique_dates.sort(reverse=True)  # sort the unique dates in reverse order
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = "clock_rings_summary" + "_" + stamp + ".txt"
    try:
        report = open(dir_path('report') + filename, "w")
        report.write("\nClock Rings Summary Report\n\n\n")
        report.write('   Showing results for:\n')
        report.write('   Station: {}\n'.format(station))
        report.write('\n')
        report.write('{:>4}  {:<26} {:<24}\n'.format("", "Date", "Records Available"))
        report.write('      --------------------------------------------\n')
        i = 1
        for line in unique_dates:
            report.write('{:>4}  {:<26} {:<24}\n'
                         .format("", dt_converter(line).strftime("%m/%d/%Y - %a"), gross_dates.count(line)))
            if i % 3 == 0:
                report.write('      --------------------------------------------\n')
            i += 1
        report.write('\n')
        report.write('Total distinct dates for which clock ring records are available: {:<9}\n'.format(i - 1))
        report.close()
        if sys.platform == "win32":  # open the text document
            os.startfile(dir_path('report') + filename)
        if sys.platform == "linux":
            subprocess.call(["xdg-open", 'kb_sub/report/' + filename])
        if sys.platform == "darwin":
            subprocess.call(["open", dir_path('report') + filename])
    except PermissionError:
        messagebox.showerror("Report Generator",
                             "The report failed to generate.",
                             parent=frame)


def database_delete_carriers_apply(frame, station, vars):
    if station.get() == "Select a station":
        station_string = "x"
    else:
        station_string = station.get()

    del_holder = []
    for pair in vars:
        if pair[1].get():
            del_holder.append(pair[0])
    if len(del_holder) > 0:
        if messagebox.askokcancel("Delete Carrier Records",
                                  "Are you sure you want to delete {} carriers, \n"
                                  "along with all their clock rings and name indexes? \n\n"
                                  "This action is not reversible.".format(len(del_holder)),
                                  parent=frame):
            pb_root = Tk()  # create a window for the progress bar
            pb_root.title("Deleting Carrier Records")
            titlebar_icon(pb_root)
            pb_label = Label(pb_root, text="Running Process: ", anchor="w")  # make label for progress bar
            pb_label.grid(row=0, column=0, sticky="w")
            pb = ttk.Progressbar(pb_root, length=400, mode="determinate")  # create progress bar
            pb.grid(row=0, column=1, sticky="w")
            pb_text = Label(pb_root, text="", anchor="w")
            pb_text.grid(row=1, column=0, columnspan=2, sticky="w")
            steps = len(del_holder)
            pb_count = 0
            pb["maximum"] = steps  # set length of progress bar
            pb.start()
            for name in del_holder:
                pb_count += 1
                # change text for progress bar
                pb_text.config(text="Deleting records for: {}".format(name))
                pb_root.update()
                pb["value"] = pb_count  # increment progress bar
                sql = "DELETE FROM rings3 WHERE carrier_name = '%s'" % name
                commit(sql)
                sql = "DELETE FROM carriers WHERE carrier_name = '%s'" % name
                commit(sql)
                sql = "DELETE FROM name_index WHERE kb_name = '%s'" % name
                commit(sql)
            pb.stop()  # stop and destroy the progress bar
            pb_label.destroy()  # destroy the label for the progress bar
            pb.destroy()
            pb_root.destroy()
            database_delete_carriers(frame, station_string)
        else:
            return


def database_chg_station(frame, station):
    if station.get() == "Select a station":
        station_string = "x"
    else:
        station_string = station.get()
    database_delete_carriers(frame, station_string)


def database_delete_carriers(frame, station):
    wd = front_window(frame)
    Label(wd[3], text="Delete Carriers", font=macadj("bold", "Helvetica 18")) \
        .grid(row=0, column=0, sticky="w")
    Label(wd[3], text="").grid(row=1, column=0)
    Label(wd[3], text="Select the station to see all carriers who have ever worked "
                      "at the station - past and present. \nDeleting the carrier will"
                      "result in all records for that carrier being deleted. This "
                      "includes clock \nrings and name indexes. ", justify=LEFT) \
        .grid(row=2, column=0, sticky="w", columnspan=6)
    Label(wd[3], text="").grid(row=3, column=0)
    Label(wd[3], text="Select Station: ", anchor="w").grid(row=4, column=0, sticky="w")
    station_selection = StringVar(wd[3])
    om_station = OptionMenu(wd[3], station_selection, *projvar.list_of_stations)
    om_station.config(width=30, anchor="w")
    om_station.grid(row=5, column=0, columnspan=2, sticky="w")
    if station == "x":
        station_selection.set("Select a station")
    else:
        station_selection.set(station)
    Button(wd[3], text="select", width=macadj(14, 12), anchor="w",
           command=lambda: database_chg_station(wd[0], station_selection)) \
        .grid(row=5, column=2, sticky="w")
    Label(wd[3], text="                ",
          anchor="w").grid(row=5, column=3, sticky="w")
    Label(wd[3], text="").grid(row=6, column=0)
    sql = "SELECT DISTINCT carrier_name FROM carriers WHERE station = '%s' " \
          "ORDER BY carrier_name ASC" % station
    results = inquire(sql)
    if station != "x":
        Label(wd[3], text="Carriers of {}".format(station), anchor="w").grid(row=7, column=0, sticky="w")
    results_frame = Frame(wd[3])
    results_frame.grid(row=8, columnspan=4)
    i = 0
    vars = []
    if len(results) == 0 and station != "x":
        Label(results_frame, text="", anchor="w").grid(row=i, column=2, sticky="w")
        i += 1
        Label(results_frame, text="After a search, no carrier records were found in the Klustebox database",
              anchor="w").grid(row=i, column=0, columnspan=3, sticky="w")
        Label(results_frame, text="                                    ",
              anchor="w").grid(row=i, column=3, sticky="w")
    for name in results:
        sql = "SELECT MAX(effective_date), station FROM carriers WHERE carrier_name = '%s'" % name
        top_rec = inquire(sql)
        var = BooleanVar()
        chk = Checkbutton(results_frame, text=name[0], variable=var, anchor="w")
        chk.grid(row=i, column=0, sticky="w")
        vars.append((name[0], var))
        Label(results_frame, text=dt_converter(top_rec[0][0]).strftime("%m/%d/%Y"), anchor="w") \
            .grid(row=i, column=1, sticky="w")
        Label(results_frame, text="     ", anchor="w").grid(row=i, column=2, sticky="w")
        Label(results_frame, text=top_rec[0][1], anchor="w").grid(row=i, column=3, sticky="w")
        Label(results_frame, text="                 ", anchor="w").grid(row=i, column=4, sticky="w")
        i += 1
    # apply and close buttons
    button_apply = Button(wd[4])
    button_back = Button(wd[4])
    button_apply.config(text="Apply", width=15,
                        command=lambda: database_delete_carriers_apply(wd[0], station_selection, vars))
    button_back.config(text="Go Back", width=15, command=lambda: MainFrame().start(frame=wd[0]))
    if sys.platform == "win32":
        button_apply.config(anchor="w")
        button_back.config(anchor="w")
    button_apply.pack(side=LEFT)
    button_back.pack(side=LEFT)
    rear_window(wd)


def database_delete_records(masterframe, frame, time_range, date, end_date, table, stations):
    db_date = datetime(1, 1, 1, 0, 0)
    db_end_date = datetime(1, 1, 1, 0, 0)
    table_array = []
    if time_range.get() != "all":
        if informalc_date_checker(frame, date, "date") == "fail":
            return
    if time_range.get() == "between":
        if informalc_date_checker(frame, end_date, "end date") == "fail":
            return
    if table.get() == "" or stations.get() == "":
        if messagebox.showerror("Database Maintenance",
                                "You must select a table and a station. ",
                                parent=frame):
            return
    if not messagebox.askokcancel("Database Maintenance",
                                  "This action will delete records from the database. \n\n"
                                  "This action is irreversible. \n\n"
                                  "Are you sure you want to proceed?",
                                  parent=frame):
        return
    #  convert date to format usable by sqlite
    if time_range.get() != "all":
        d = date.get().split("/")
        db_date = datetime(int(d[2]), int(d[0]), int(d[1]))
    if time_range.get() == "between":
        d = end_date.get().split("/")
        db_end_date = datetime(int(d[2]), int(d[0]), int(d[1]))
    # define the station array to loop
    if stations.get() == "all stations":
        station_array = projvar.list_of_stations[:]
    else:
        station_array = [stations.get()]
    # define the table array to loop
    if table.get() == "all":
        table_array = ["rings3", "name_index", "carriers", "stations", "station_index"]
    elif table.get() == "carriers + index":
        table_array = ["carriers", "name_index"]
    elif table.get() == "carriers":
        table_array = ["carriers"]
    elif table.get() == "name index":
        table_array = ["name_index"]
    elif table.get() == "clock rings":
        table_array = ["rings3"]
    #  short cuts to delete all records in table
    if time_range.get() == "all" and stations.get() == "all stations":
        for tab in table_array:
            if tab == "stations":
                sql = "DELETE FROM stations"
                commit(sql)
            if tab == "station_index":
                sql = "DELETE FROM station_index"
                commit(sql)
            if tab == "name_index":
                sql = "DELETE FROM name_index"
                commit(sql)
            if tab == "carriers":
                sql = "DELETE FROM carriers"
                commit(sql)
            if tab == "rings3":
                sql = "DELETE FROM rings3"
                commit(sql)
            if tab == "stations":
                sql = "DELETE FROM stations"
                commit(sql)
                sql = "INSERT INTO stations (station) VALUES('%s')" % "out of station"
                commit(sql)
                del projvar.list_of_stations[:]
                projvar.list_of_stations.append("out of station")

                reset("none")  # reset investigation range
        messagebox.showinfo("Database Maintenance",
                            "Success! The database has been cleaned of the specified records.",
                            parent=frame)
        frame.destroy()
        database_maintenance(masterframe)
        return
    # loop for great justice
    operator = ""
    for stat in station_array:
        for tab in table_array:
            # delete all rings associated with station
            if tab == "stations":
                if stat != "out of station":
                    sql = "DELETE FROM stations WHERE station = '%s'" % stat
                    commit(sql)
                if stat != "out of station":
                    projvar.list_of_stations.remove(stat)
                if projvar.invran_station == stat:
                    reset("none")  # reset initial value of globals
            if tab == "station_index":
                if stat != "out of station":
                    sql = "DELETE FROM station_index WHERE kb_station = '%s'" % stat
                    commit(sql)
            if tab == "rings3":
                # determine operator based on time_range
                if time_range.get() == "before":
                    operator = " AND rings_date <= '%s'" % db_date
                elif time_range.get() == "this_date":
                    operator = " AND rings_date = '%s'" % db_date
                elif time_range.get() == "after":
                    operator = " AND rings_date >= '%s'" % db_date
                elif time_range.get() == "all":
                    operator = ""
                elif time_range.get() == "between":
                    operator = " AND rings_date BETWEEN '%s' AND '%s'" % (db_date, db_end_date)
                sql = "SELECT DISTINCT carrier_name FROM carriers WHERE station = '%s' ORDER BY carrier_name" \
                      % stat
                result = inquire(sql)
                pb_root = Tk()  # create a window for the progress bar
                pb_root.title("Deleting Clock Rings from {}".format(stat))
                titlebar_icon(pb_root)
                pb_label = Label(pb_root, text="Running Process: ", anchor="w")  # make label for progress bar
                pb_label.grid(row=0, column=0, sticky="w")
                pb = ttk.Progressbar(pb_root, length=400, mode="determinate")  # create progress bar
                pb.grid(row=0, column=1, sticky="w")
                pb_text = Label(pb_root, text="", anchor="w")
                pb_text.grid(row=1, column=0, columnspan=2, sticky="w")
                steps = len(result)
                pb_count = 0
                pb["maximum"] = steps  # set length of progress bar
                pb.start()
                for name in result:
                    pb_count += 1
                    active_station = []
                    # get all records for the carrier
                    sql = "SELECT * FROM carriers WHERE carrier_name= '%s' ORDER BY effective_date" % name[0]
                    result_1 = inquire(sql)
                    start_search = True
                    start = ''
                    end = ''
                    # build the active_station array - find dates where carrier entered/left station
                    for r in result_1:
                        if r[5] == stat and start_search is True:
                            start = r
                            start_search = False
                        if r[5] != stat and start_search is False:
                            end = r
                            active_station.append([start, end])
                            start = ''
                            end = ''
                            start_search = True
                    if not start_search:
                        active_station.append([start, end])
                    for active in active_station:
                        if active[1] != '':
                            sql = "DELETE FROM rings3 WHERE rings_date " \
                                  "BETWEEN '%s' AND '%s'{} AND carrier_name = '%s' " \
                                      .format(operator) % (active[0][0], active[1][0], name[0])
                            commit(sql)
                            # change text for progress bar
                            pb_text.config(text="Deleting in range rings for: {} - {} through {}"
                                           .format(name[0], active[0][0], active[1][0]))
                            pb["value"] = pb_count  # increment progress bar
                            pb_root.update()
                        else:
                            sql = "DELETE FROM rings3 WHERE rings_date >= '%s'{} AND carrier_name = '%s' " \
                                      .format(operator) % (active[0][0], name[0])
                            commit(sql)
                            # change text for progress bar
                            pb_text.config(text="Deleting in range rings for: {} - {} +".format(name[0], active[0][0]))
                            pb_root.update()
                            pb["value"] = pb_count  # increment progress bar
                pb.stop()  # stop and destroy the progress bar
                pb_label.destroy()  # destroy the label for the progress bar
                pb.destroy()
                pb_root.destroy()
            if tab == "name_index":
                sql = "SELECT DISTINCT carrier_name FROM carriers WHERE station = '%s'" \
                      % stat
                results = inquire(sql)
                pb_root = Tk()  # create a window for the progress bar
                pb_root.title("Deleting Clock Rings from {}".format(stat))
                titlebar_icon(pb_root)
                pb_label = Label(pb_root, text="Running Process: ", anchor="w")  # make label for progress bar
                pb_label.grid(row=0, column=0, sticky="w")
                pb = ttk.Progressbar(pb_root, length=400, mode="determinate")  # create progress bar
                pb.grid(row=0, column=1, sticky="w")
                pb_text = Label(pb_root, text="", anchor="w")
                pb_text.grid(row=1, column=0, columnspan=2, sticky="w")
                steps = len(results)
                pb_count = 0
                pb["maximum"] = steps  # set length of progress bar
                pb.start()
                for car in results:
                    sql = "DELETE FROM name_index WHERE kb_name='%s'" % car[0]
                    commit(sql)
                    pb_count += 1
                    pb_text.config(text="Deleting name index for: {}".format(car[0]))
                    pb["value"] = pb_count  # increment progress bar
                    pb_root.update()
                pb.stop()  # stop and destroy the progress bar
                pb_label.destroy()  # destroy the label for the progress bar
                pb.destroy()
                pb_root.destroy()
            if tab == "carriers":
                # determine operator based on time_range
                if time_range.get() == "before":
                    operator = "AND effective_date <= '%s'" % db_date
                elif time_range.get() == "this_date":
                    operator = "AND effective_date = '%s'" % db_date
                elif time_range.get() == "after":
                    operator = "AND effective_date >= '%s'" % db_date
                elif time_range.get() == "all":
                    operator = ""
                elif time_range.get() == "between":
                    operator = "AND '%s' <= effective_date <= '%s'" % (db_date, db_end_date)
                sql = "SELECT DISTINCT carrier_name FROM carriers WHERE station = '%s' {}" \
                          .format(operator) % stat
                results = inquire(sql)
                pb_root = Tk()  # create a window for the progress bar
                pb_root.title("Deleting Carrier Records from {}".format(stat))
                titlebar_icon(pb_root)
                pb_label = Label(pb_root, text="Running Process: ", anchor="w")  # make label for progress bar
                pb_label.grid(row=0, column=0, sticky="w")
                pb = ttk.Progressbar(pb_root, length=400, mode="determinate")  # create progress bar
                pb.grid(row=0, column=1, sticky="w")
                pb_text = Label(pb_root, text="", anchor="w")
                pb_text.grid(row=1, column=0, columnspan=2, sticky="w")
                steps = len(results)
                pb_count = 0
                pb["maximum"] = steps  # set length of progress bar
                pb.start()
                for car in results:
                    pb_text.config(text="Deleting clock rings for: {}".format(car[0]))  # change text for progress bar
                    pb_count += 1
                    pb["value"] = pb_count  # increment progress bar
                    pb_root.update()
                    sql = "SELECT * FROM carriers WHERE  carrier_name = '%s' {}".format(operator) % car[0]
                    car_ask = inquire(sql)
                    outside_station = False
                    for cc in car_ask:  # look for rings where the station doesn't match or out of station
                        if cc[5] != "out of station" or cc[5] != stat:
                            outside_station = True
                    if not outside_station:
                        for carr in results:
                            # update all records where station/carrier match to 'out of station'
                            sql = "UPDATE carriers SET station='%s' WHERE carrier_name ='%s' AND station='%s' {}" \
                                      .format(operator) % ("out of station", carr[0], stat)
                            commit(sql)
                            # find redundancies where two 'out of station' records are adjacent.
                            sql = "SELECT * FROM carriers WHERE carrier_name ='%s' " \
                                  "ORDER BY carrier_name, effective_date" % carr[0]
                            car_results = inquire(sql)
                            duplicates = []
                            for i in range(len(car_results)):
                                if i != len(car_results) - 1:  # if the loop has not reached the end of the list
                                    # if the name current and next name are the same
                                    if car_results[i][5] == 'out of station' and \
                                            car_results[i + 1][5] == 'out of station':
                                        duplicates.append(i + 1)
                            for d in duplicates:
                                sql = "DELETE FROM carriers WHERE effective_date='%s' and carrier_name='%s'" % (
                                    car_results[d][0], car_results[d][1])
                                commit(sql)
                            # find and delete records where a carrier has only 'one out of station' record
                            sql = "SELECT station FROM carriers WHERE carrier_name = '%s'" \
                                  % carr[0]
                            if len(inquire(sql)) == 1:
                                sql = "DELETE FROM carriers WHERE carrier_name = '%s' AND station = '%s'" \
                                      % (carr[0], "out of station")
                                commit(sql)
                    else:
                        sql = "DELETE FROM carriers WHERE carrier_name = '%s' {}".format(operator) % car[0]
                        commit(sql)
                pb.stop()  # stop and destroy the progress bar
                pb_label.destroy()  # destroy the label for the progress bar
                pb.destroy()
                pb_root.destroy()
    messagebox.showinfo("Database Maintenance",
                        "Success! The database has been cleaned of the specified records.",
                        parent=frame)
    frame.destroy()
    database_maintenance(masterframe)


def database_reset(masterframe, frame):  # deletes the database and rebuilds it.
    if not messagebox.askokcancel("Delete Database",
                                  "This action will delete your database and all information inside it."
                                  "This includes carrier information, rings information, settings as "
                                  "well as any informal c data. The database will be rebuilt and will be "
                                  "like new. "
                                  "\n\n This action can not be reversed."
                                  "\n\n Are you sure you want to proceed?", parent=frame):
        return
    path = "kb_sub/mandates.sqlite"
    if projvar.platform == "macapp":
        path = os.path.expanduser("~") + '/Documents/.klusterbox/mandates.sqlite'
    if projvar.platform == "winapp":
        path = os.path.expanduser("~") + '\\Documents\\.klusterbox\\mandates.sqlite'
    if projvar.platform == "py":
        path = "kb_sub/mandates.sqlite"
    try:
        if os.path.exists(path):
            os.remove(path)
    except sqlite3.OperationalError:
        messagebox.showerror("Access Error",
                             "Klusterbox can not delete the database as it is being used by another "
                             "application. Close the database in the other application and retry.",
                             parent=frame)
    frame.destroy()
    masterframe.destroy()
    reset("none")  # reset initial value of globals
    DataBase().setup()
    StartUp().start()


def database_clean_carriers():  # delete carrier records where station no longer exist
    sql = "SELECT DISTINCT station FROM carriers"
    all_stations = inquire(sql)
    sql = "SELECT station FROM stations"
    good_stations = inquire(sql)
    deceased = [x for x in all_stations if x not in good_stations]
    pb_root = Tk()  # create a window for the progress bar
    pb_root.title("Deleting Orphaned Clock Rings")
    titlebar_icon(pb_root)  # place icon in titlebar
    pb_label = Label(pb_root, text="Running Process: ", anchor="w")  # make label for progress bar
    pb_label.grid(row=0, column=0, sticky="w")
    pb = ttk.Progressbar(pb_root, length=400, mode="determinate")  # create progress bar
    pb.grid(row=0, column=1, sticky="w")
    pb_text = Label(pb_root, text="", anchor="w")
    pb_text.grid(row=1, column=0, columnspan=2, sticky="w")
    steps = len(deceased)
    pb_count = 0
    pb["maximum"] = steps  # set length of progress bar
    pb.start()
    for dead in deceased:
        sql = "DELETE FROM carriers WHERE station ='%s'" % (dead[0])
        commit(sql)
        pb_count += 1
        pb_text.config(text="Deleting carrier records for: {}".format(dead[0]))
        pb["value"] = pb_count  # increment progress bar
        pb_root.update()
    sql = "DELETE FROM rings3 WHERE carrier_name IS Null"
    commit(sql)
    pb.stop()  # stop and destroy the progress bar
    pb_label.destroy()  # destroy the label for the progress bar
    pb.destroy()
    pb_root.destroy()


def database_clean_rings():
    sql = "SELECT DISTINCT carrier_name FROM carriers"
    carriers_results = inquire(sql)
    sql = "SELECT DISTINCT carrier_name FROM rings3"
    rings_results = inquire(sql)
    deceased = [x for x in rings_results if x not in carriers_results]
    pb_root = Tk()  # create a window for the progress bar
    pb_root.title("Deleting Orphaned Clock Rings")
    titlebar_icon(pb_root)  # place icon in titlebar
    pb_label = Label(pb_root, text="Running Process: ", anchor="w")  # make label for progress bar
    pb_label.grid(row=0, column=0, sticky="w")
    pb = ttk.Progressbar(pb_root, length=400, mode="determinate")  # create progress bar
    pb.grid(row=0, column=1, sticky="w")
    pb_text = Label(pb_root, text="", anchor="w")
    pb_text.grid(row=1, column=0, columnspan=2, sticky="w")
    steps = len(deceased)
    pb_count = 0
    pb["maximum"] = steps  # set length of progress bar
    pb.start()
    for dead in deceased:
        pb_text.config(text="Deleting clock rings for: {}".format(dead[0]))
        pb_count += 1
        pb["value"] = pb_count  # increment progress bar
        # change text for progress bar
        pb_root.update()
        sql = "DELETE FROM rings3 WHERE carrier_name='%s'" % dead
        commit(sql)
    pb_text.config(text="Deleting NULL clock rings.")
    pb_root.update()
    sql = "DELETE FROM rings3 WHERE carrier_name IS Null"
    commit(sql)
    sql = "DELETE FROM rings3 WHERE total='%s' and code='%s' and leave_type ='%s'" % ("", 'none', '0.0')
    commit(sql)
    pb.stop()  # stop and destroy the progress bar
    pb_label.destroy()  # destroy the label for the progress bar
    pb.destroy()
    pb_root.destroy()


def database_maintenance(frame):
    wd = front_window(frame)
    r = 0
    Label(wd[3], text="Database Maintenance", font=macadj("bold", "Helvetica 18"), anchor="w") \
        .grid(row=r, sticky="w", columnspan=4)
    r += 1
    Label(wd[3], text="").grid(row=r)
    r += 1
    Label(wd[3], text="Database Records").grid(row=r, sticky="w", columnspan=4)
    r += 1
    Label(wd[3], text="                    ").grid(row=r, column=0, sticky="w")
    r += 1
    # get and display number of records for rings3
    sql = "SELECT COUNT (*) FROM rings3"
    results = inquire(sql)
    Label(wd[3], text=results, anchor="e", fg="red").grid(row=r, column=0, sticky="e")
    Label(wd[3], text=" total records in rings table").grid(row=r, column=1, sticky="w")
    r += 1
    # get and display number of records for unique carriers in rings3
    sql = "SELECT COUNT (DISTINCT carrier_name) FROM rings3"
    results = inquire(sql)
    Label(wd[3], text=results, anchor="e", fg="red").grid(row=r, column=0, sticky="e")
    Label(wd[3], text=" distinct carrier names in rings table").grid(row=r, column=1, sticky="w")
    r += 1
    # get and display number of records for unique days in rings3
    sql = "SELECT COUNT (DISTINCT rings_date) FROM rings3"
    results = inquire(sql)
    Label(wd[3], text=results, anchor="e", fg="red").grid(row=r, column=0, sticky="e")
    Label(wd[3], text=" distinct days in rings table").grid(row=r, column=1, sticky="w")
    r += 1
    # get and display number of records for carriers
    sql = "SELECT COUNT (*) FROM carriers"
    results = inquire(sql)
    Label(wd[3], text=results, anchor="e", fg="red").grid(row=r, column=0, sticky="e")
    Label(wd[3], text=" total records in carriers table").grid(row=r, column=1, sticky=W)
    r += 1
    # get and display number of records for distinct carrier names from carriers
    sql = "SELECT COUNT (DISTINCT carrier_name) FROM carriers"
    results = inquire(sql)
    Label(wd[3], text=results, anchor="e", fg="red").grid(row=r, column=0, sticky="e")
    Label(wd[3], text=" distinct carrier names in carriers table").grid(row=r, column=1, sticky=W)
    r += 1
    # get and display number of records for stations
    sql = "SELECT COUNT (*) FROM stations"
    results = inquire(sql)
    Label(wd[3], text=results, anchor="e", fg="red").grid(row=r, column=0, sticky="e")
    Label(wd[3], text=" total records in station table (this includes \'out of station\')") \
        .grid(row=r, column=1, sticky="w")
    r += 1
    # find orphaned rings from deceased carriers
    sql = "SELECT DISTINCT carrier_name FROM carriers"
    carriers_results = inquire(sql)
    sql = "SELECT DISTINCT carrier_name FROM rings3"
    rings_results = inquire(sql)
    deceased = [x for x in rings_results if x not in carriers_results]
    Label(wd[3], text=len(deceased), anchor="e", fg="red").grid(row=r, column=0, sticky="e")
    Label(wd[3], text=" \'deceased\' carriers in rings table").grid(row=r, column=1, sticky=W)
    r += 1
    if len(deceased) > 0:
        Label(wd[3], text="").grid(row=r, column=0, sticky="w")
        r += 1
        Button(wd[3], text="clean",
               command=lambda: (database_clean_rings(), wd[0].destroy(), database_maintenance(frame))) \
            .grid(row=r, column=0, sticky="w")
        Label(wd[3], text="Delete rings records where carriers no longer exist (recommended)") \
            .grid(row=r, column=1, sticky="w", columnspan=6)
        r += 1
        Label(wd[3], text="").grid(row=r, column=0, sticky="w")
        r += 1
    sql = "SELECT DISTINCT station FROM carriers"
    all_stations = inquire(sql)
    sql = "SELECT station FROM stations"
    good_stations = inquire(sql)
    deceased_cars = [x for x in all_stations if x not in good_stations]
    Label(wd[3], text=len(deceased_cars), anchor="e", fg="red").grid(row=r, column=0, sticky="e")
    Label(wd[3], text=" \'deceased\' stations in carriers table").grid(row=r, column=1, sticky=W)
    r += 1
    if len(deceased_cars) > 0:
        Label(wd[3], text="").grid(row=r, column=0, sticky="w")
        r += 1
        Button(wd[3], text="clean",
               command=lambda: (database_clean_carriers(), wd[0].destroy(), database_maintenance(frame))) \
            .grid(row=r, column=0, sticky="w")
        Label(wd[3], text="Delete carrier records where station no longer exist (recommended)") \
            .grid(row=r, column=1, sticky="w", columnspan=6)
        r += 1
    if projvar.invran_station is None:
        Label(wd[3], text="").grid(row=r, column=0, sticky="w")
        r += 1
        Label(wd[3], text="Database Records, {} Specific".format(projvar.invran_station)) \
            .grid(row=r, sticky="w", columnspan=4)
        r += 1
        Label(wd[3], text="To see results from other stations, change station "
                          "in the investigation range", fg="grey") \
            .grid(row=r, column=0, sticky="w", columnspan=6)
        r += 1
        Label(wd[3], text="                    ").grid(row=r, column=0, sticky="w")
        r += 1
        # get and display number of records for carriers
        sql = "SELECT COUNT (*) FROM carriers WHERE station = '%s'" % projvar.invran_station
        results = inquire(sql)
        Label(wd[3], text=results, anchor="e", fg="red").grid(row=r, column=0, sticky="e")
        Label(wd[3], text=" total records in carriers table").grid(row=r, column=1, sticky=W)
        r += 1
        # get and display number of records for distinct carrier names from carriers
        sql = "SELECT COUNT (DISTINCT carrier_name) FROM carriers WHERE station = '%s'" % projvar.invran_station
        results = inquire(sql)
        Label(wd[3], text=results, anchor="e", fg="red").grid(row=r, column=0, sticky="e")
        Label(wd[3], text=" distinct carrier names in carriers table").grid(row=r, column=1, sticky=W)
        r += 1
    if "out of station" in projvar.list_of_stations:
        Label(wd[3], text="").grid(row=r, column=0, sticky="w")
        r += 1
        Label(wd[3], text="Database Records, for \"{}\"".format("out of station")) \
            .grid(row=r, sticky="w", columnspan=4)
        r += 1
        Label(wd[3], text="                    ").grid(row=r, column=0, sticky="w")
        r += 1
        # get and display number of records for carriers
        sql = "SELECT COUNT (*) FROM carriers WHERE station = '%s'" % "out of station"
        results = inquire(sql)
        Label(wd[3], text=results, anchor="e", fg="red").grid(row=r, column=0, sticky="e")
        Label(wd[3], text=" total records in carriers table").grid(row=r, column=1, sticky=W)
        r += 1
        # get and display number of records for distinct carrier names from carriers
        sql = "SELECT COUNT (DISTINCT carrier_name) FROM carriers WHERE station = '%s'" % "out of station"
        results = inquire(sql)
        Label(wd[3], text=results, anchor="e", fg="red").grid(row=r, column=0, sticky="e")
        Label(wd[3], text=" distinct carrier names in carriers table").grid(row=r, column=1, sticky=W)
        r += 1
        Label(wd[3], text="").grid(row=r)
        r += 1
    #  Clock Rings summary
    rings_frame = Frame(wd[3])
    rings_frame.grid(row=r, column=0, columnspan=6, sticky=W)
    rings_station = StringVar(rings_frame)
    rr = 0
    Label(rings_frame, text="Clock Rings Summary Report by Station:").grid(row=rr, column=0, columnspan=6, sticky=W)
    rr += 1
    Label(rings_frame, text="").grid(row=rr)
    rr += 1
    Label(rings_frame, text="Station: ").grid(row=rr, column=0, sticky=W)
    om_rings = OptionMenu(rings_frame, rings_station, *projvar.list_of_stations)
    om_rings.config(width=20)
    if projvar.invran_station is None:
        present_station = projvar.invran_station
    else:
        present_station = "select a station"
    rings_station.set(present_station)
    om_rings.grid(row=rr, column=1, sticky=W)
    Button(rings_frame, text="Report", width=8, command=lambda: database_rings_report(wd[0], rings_station.get())) \
        .grid(row=rr, column=2, sticky=W, padx=20)
    rr += 1
    Label(rings_frame, text="").grid(row=rr)
    r += 1
    # declare variables for Delete Database Records
    clean1_range = StringVar(wd[3])
    clean1_date = StringVar(wd[3])
    clean1_table = StringVar(wd[3])
    clean1_station = StringVar(wd[3])
    # create frame and widgets for Delete Database Records
    cleaner_frame1 = Frame(wd[3])
    cleaner_frame1.grid(row=r, columnspan=6)
    rr = 0
    Label(cleaner_frame1, text="Delete Database Records (Remove records from database per given specifications)",
          anchor="w").grid(row=rr, sticky="w", columnspan=4, column=0)
    rr += 1
    Label(cleaner_frame1, text="* format all date fields as mm/dd/yyyy, failure to do so will return an error",
          anchor="w", fg="grey").grid(row=rr, sticky="w", columnspan=4, column=0)
    rr += 1
    Label(cleaner_frame1, text="                                               ").grid(row=rr, column=5)
    rr += 1
    Label(cleaner_frame1, text="Delete Records: ", anchor="w").grid(row=rr, sticky="w", column=0)
    Radiobutton(cleaner_frame1, text="before and on date", variable=clean1_range, value="before",
                anchor="w").grid(row=rr, sticky="w", column=1)
    rr += 1
    Radiobutton(cleaner_frame1, text="entered date only", variable=clean1_range, value="this_date",
                anchor="w").grid(row=rr, sticky="w", column=1)
    rr += 1
    Radiobutton(cleaner_frame1, text="after and on date", variable=clean1_range, value="after",
                anchor="w").grid(row=rr, sticky="w", column=1)
    rr += 1
    Radiobutton(cleaner_frame1, text="all dates", variable=clean1_range, value="all", anchor="w") \
        .grid(row=rr, sticky="w", column=1)
    clean1_range.set("after")
    r += 1
    # create frame and widgets for Delete Database Records
    cleaner_frame2 = Frame(wd[3])
    cleaner_frame2.grid(row=r, columnspan=6, sticky="w")
    rrr = 0
    Label(cleaner_frame2, text="date* ", anchor="e").grid(row=rrr, column=0, sticky="e")
    Entry(cleaner_frame2, textvariable=clean1_date, width=macadj(12, 8), justify='right') \
        .grid(row=rrr, column=1, sticky="w")
    Label(cleaner_frame2, text="         table", anchor="e").grid(row=rrr, column=2, sticky="e")
    table_options = ("carriers + index", "carriers", "name index", "clock rings", "all")
    om1_table = OptionMenu(cleaner_frame2, clean1_table, *table_options)
    clean1_table.set(table_options[-1])
    if sys.platform != "darwin":
        om1_table.config(width=20, anchor="w")
    else:
        om1_table.config(width=20)
    om1_table.grid(row=rrr, column=3, sticky="w")
    rrr += 1
    station_options = projvar.list_of_stations[:]  # use splice to make copy of list without creating alias
    station_options.append("all stations")
    Label(cleaner_frame2, text="stations", anchor="e").grid(row=rrr, column=2, sticky="e")
    om1_station = OptionMenu(cleaner_frame2, clean1_station, *station_options)
    clean1_station.set(station_options[-1])
    if sys.platform != "darwin":
        om1_station.config(width=20, anchor="w")
    else:
        om1_station.config(width=20)
    om1_station.grid(row=rrr, column=3, sticky="w")
    Button(cleaner_frame2, text="delete", width=macadj(6, 5),
           command=lambda: database_delete_records
           (frame, wd[0], clean1_range, clean1_date, "x", clean1_table, clean1_station)) \
        .grid(row=rrr, column=4, sticky="w")
    rrr += 1
    Label(cleaner_frame2, text="").grid(row=rrr)
    rrr += 1
    # declare variables for Delete Database Records
    clean2_range = StringVar(wd[3])
    clean2_startdate = StringVar(wd[3])
    clean2_enddate = StringVar(wd[3])
    clean2_table = StringVar(wd[3])
    clean2_station = StringVar(wd[3])
    rr += 1
    Label(cleaner_frame2, text="Delete Records within a specified range: ", anchor="w") \
        .grid(row=rrr, sticky="w", column=0, columnspan=6)
    rrr += 1
    Label(cleaner_frame2, text="* format all date fields as mm/dd/yyyy, failure to do so will return an error",
          anchor="w", fg="grey").grid(row=rrr, sticky="w", columnspan=4)
    rrr += 1
    # declare range as "between by default
    clean2_range.set("between")
    Label(cleaner_frame2, text="     start date* ", anchor="e").grid(row=rrr, column=0, sticky="e")
    Entry(cleaner_frame2, textvariable=clean2_startdate, width=macadj(12, 8), justify='right') \
        .grid(row=rrr, column=1, sticky="w")
    Label(cleaner_frame2, text="         table", anchor="e").grid(row=rrr, column=2, sticky="e")
    om2_table = OptionMenu(cleaner_frame2, clean2_table, *table_options)
    clean2_table.set(table_options[-1])
    if sys.platform != "darwin":
        om2_table.config(width=20, anchor="w")
    else:
        om2_table.config(width=20)
    om2_table.grid(row=rrr, column=3, sticky="w")
    rrr += 1
    Label(cleaner_frame2, text="end date* ", anchor="e").grid(row=rrr, column=0, sticky="e")
    Entry(cleaner_frame2, textvariable=clean2_enddate, width=macadj(12, 8), justify='right') \
        .grid(row=rrr, column=1, sticky="w")
    Label(cleaner_frame2, text="stations", anchor="e").grid(row=rrr, column=2, sticky="e")
    om2_station = OptionMenu(cleaner_frame2, clean2_station, *station_options)
    clean2_station.set(station_options[-1])
    if sys.platform != "darwin":
        om2_station.config(width=20, anchor="w")
    else:
        om2_station.config(width=20)
    om2_station.grid(row=rrr, column=3, sticky="w")
    Button(cleaner_frame2, text="delete", width=macadj(6, 5),
           command=lambda: database_delete_records(frame, wd[0], clean2_range, clean2_startdate, clean2_enddate,
                                                   clean2_table, clean2_station)) \
        .grid(row=rrr, column=4, sticky="w")
    rrr += 1
    Label(cleaner_frame2, text="").grid(row=rrr)
    rrr += 1
    Label(cleaner_frame2, text="Reset Database - Delete and Rebuild the Database (all information will be lost)") \
        .grid(row=rrr, sticky="w", column=0, columnspan=6)
    rrr += 1
    Label(cleaner_frame2, text="").grid(row=rrr)
    rrr += 1
    Button(cleaner_frame2, text="Reset", width=10, padx=5, fg=macadj("white", "red"), bg=macadj("red", "white"),
           command=lambda: database_reset(frame, wd[0])) \
        .grid(row=rrr, column=0, sticky="w")
    rrr += 1
    Label(cleaner_frame2, text="").grid(row=rrr)
    rrr += 1
    Label(cleaner_frame2, text="").grid(row=rrr)
    r += 1
    button = Button(wd[4])
    button.config(text="Go Back", width=20, command=lambda: MainFrame().start(frame=wd[0]))
    if sys.platform == "win32":  # center the widget text for mac
        button.config(anchor="w")
    button.pack(side=LEFT)
    rear_window(wd)


class RptWin:
    def __init__(self, frame):
        self.frame = frame

    def rpt_chg_station(self, frame, station):
        self.frame = frame
        if station.get() == "Select a station":
            station_string = "x"
        else:
            station_string = station.get()
        self.rpt_find_carriers(station_string)

    def rpt_find_carriers(self, station):
        win = MakeWindow()
        win.create(self.frame)
        Label(win.body, text="Carriers Status History", font=macadj("bold", "Helvetica 18")) \
            .grid(row=0, column=0, sticky="w")
        Label(win.body, text="").grid(row=1, column=0)
        Label(win.body, text="Select the station to see all carriers who have ever worked "
                          "at the station - past and present. \n ", justify=LEFT) \
            .grid(row=2, column=0, sticky="w", columnspan=6)
        Label(win.body, text="").grid(row=3, column=0)
        Label(win.body, text="Select Station: ", anchor="w").grid(row=4, column=0, sticky="w")
        station_selection = StringVar(win.body)
        om_station = OptionMenu(win.body, station_selection, *projvar.list_of_stations)
        if sys.platform != "darwin":
            om_station.config(width=30, anchor="w")
        else:
            om_station.config(width=30)
        om_station.grid(row=5, column=0, columnspan=2, sticky="w")
        if station == "x":
            station_selection.set("Select a station")
        else:
            station_selection.set(station)
        Button(win.body, text="select", width=macadj(14, 12), anchor="w",
               command=lambda: self.rpt_chg_station(win.topframe, station_selection)) \
            .grid(row=5, column=2, sticky="w")
        Label(win.body, text="").grid(row=6, column=0)
        sql = "SELECT DISTINCT carrier_name FROM carriers WHERE station = '%s' " \
              "ORDER BY carrier_name ASC" % station
        results = inquire(sql)
        if station != "x":
            Label(win.body, text="Carriers of {}".format(station), anchor="w").grid(row=7, column=0, sticky="w")
        results_frame = Frame(win.body)
        results_frame.grid(row=8, columnspan=4)
        i = 0
        if station != "x":
            if len(results) > 0:
                Label(results_frame, text="Name", anchor="w", fg="grey").grid(row=i, column=0, sticky="w")
                Label(results_frame, text="Last Date", anchor="w", fg="grey") \
                    .grid(row=i, column=1, columnspan=2, sticky="w")
                Label(results_frame, text="Station", anchor="w", fg="grey").grid(row=i, column=3, sticky="w")
            elif len(results) == 0:
                Label(results_frame, text="", anchor="w").grid(row=i, column=0, sticky="w")
                i += 1
                Label(results_frame, text="After a search, no results were found in the klusterbox database.",
                      anchor="w") \
                    .grid(row=i, column=0, sticky="w")
        i += 1
        for name in results:
            sql = "SELECT MAX(effective_date), station FROM carriers WHERE carrier_name = '%s'" % name
            top_rec = inquire(sql)
            Label(results_frame, text=name[0], anchor="w").grid(row=i, column=0, sticky="w")
            Label(results_frame, text=dt_converter(top_rec[0][0]).strftime("%m/%d/%Y"), anchor="w") \
                .grid(row=i, column=1, sticky="w")
            Label(results_frame, text="     ", anchor="w").grid(row=i, column=2, sticky="w")
            Label(results_frame, text=top_rec[0][1], anchor="w").grid(row=i, column=3, sticky="w")
            Label(results_frame, text="     ", anchor="w").grid(row=i, column=4, sticky="w")
            Button(results_frame, text="Report", anchor="w",
                   command=lambda in_line=name: Reports(self.frame).rpt_carrier_history(in_line[0])) \
                .grid(row=i, column=5, sticky="w")
            Label(results_frame, text="         ", anchor="w").grid(row=i, column=6, sticky="w")
            i += 1
        # apply and close buttons
        button = Button(win.buttons)
        button.config(text="Go Back", width=macadj(20, 20),
                      command=lambda: MainFrame().start(frame=win.topframe))
        if sys.platform == "win32":
            button.config(anchor="w")
        button.pack(side=LEFT)
        win.finish()


def clean_rings3_table():  # database maintenance
    sql = "SELECT * FROM rings3 WHERE leave_type IS NULL"
    result = inquire(sql)
    types = ""
    times = float(0.0)
    if result:
        sql = "UPDATE rings3 SET leave_type='%s',leave_time='%s'" \
              "WHERE leave_type IS NULL" \
              % (types, times)
        commit(sql)
        messagebox.showinfo("Clean Rings",
                            "Rings table has been cleared of NULL values in leave type and leave time columns.")
    else:
        messagebox.showinfo("Clean Rings",
                            "No NULL values in leave type and leave time columns were found in the Rings3 "
                            "table of the database. No action taken.")
    return


def ns_config_apply(frame, text_array, color_array):
    # set ns configurations from Non-Scheduled Day Configurations page
    for t in text_array:
        if len(t.get()) > 6:
            messagebox.showerror("Non_Scheduled Day Configuration",
                                 "Names must not be longer than 6 characters.",
                                 parent=frame)
            return
        if len(t.get()) < 1:
            messagebox.showerror("Non_Scheduled Day Configuration",
                                 "Names must not be shorter than 1 character.",
                                 parent=frame)
            return
    color = ("yellow", "blue", "green", "brown", "red", "black")
    for i in range(6):
        sql = "UPDATE ns_configuration SET custom_name ='%s' WHERE ns_name = '%s'" % (text_array[i].get(), color[i])
        commit(sql)
        sql = "UPDATE ns_configuration SET fill_color ='%s' WHERE ns_name = '%s'" % (color_array[i].get(), color[i])
        commit(sql)
    ns_config(frame)


def ns_config_reset(frame):  # reset ns day configurations from Non-Scheduled Day Configurations page
    fill = ("gold", "navy", "forest green", "saddle brown", "red3", "gray10")
    color = ("yellow", "blue", "green", "brown", "red", "black")
    for i in range(6):
        sql = "UPDATE ns_configuration SET custom_name ='%s' WHERE ns_name = '%s'" % (color[i], color[i])
        commit(sql)
        sql = "UPDATE ns_configuration SET fill_color ='%s' WHERE ns_name = '%s'" % (fill[i], color[i])
        commit(sql)
    ns_config(frame)


def ns_config(frame):  # generate Non-Scheduled Day Configurations page to configure ns day settings
    if projvar.invran_day is None:
        messagebox.showerror("Non-Scheduled Day Configurations",
                             "You must set the Investigation Range before changing the NS Day Configurations.",
                             parent=frame)
        return
    sql = "SELECT * FROM ns_configuration"
    result = inquire(sql)
    wd = front_window(frame)
    Label(wd[3], text="Non-Scheduled Day Configurations", font=macadj("bold", "Helvetica 18"), anchor="w") \
        .grid(row=0, sticky="w", columnspan=4)
    Label(wd[3], text=" ").grid(row=1, column=0)
    Label(wd[3], text="Change Configuration").grid(row=2, sticky="w", columnspan=4)
    f_date = projvar.invran_date_week[0].strftime("%a - %b %d, %Y")
    end_f_date = projvar.invran_date_week[6].strftime("%a - %b %d, %Y")
    Label(wd[3], text="Investigation Range: {0} through {1}".format(f_date, end_f_date),
          foreground="red").grid(row=3, column=0, sticky="w", columnspan=4)
    Label(wd[3], text="Pay Period: {0}".format(projvar.pay_period),
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
    Label(wd[3], text="{}".format(projvar.ns_code['yellow'])).grid(row=7, column=0, sticky="w")  # yellow row
    Entry(wd[3], textvariable=yellow_text, width=10).grid(row=7, column=1, sticky="w")
    yellow_text.set(result[0][2])
    om_yellow = OptionMenu(wd[3], yellow_color, *color_array)
    yellow_color.set(result[0][1])
    om_yellow.config(width=13, anchor="w")
    om_yellow.grid(row=7, column=2, sticky="w")
    Label(wd[3], text="yellow").grid(row=7, column=3, sticky="w")
    Label(wd[3], text="{}".format(projvar.ns_code['blue'])).grid(row=8, column=0, sticky="w")  # blue row
    Entry(wd[3], textvariable=blue_text, width=10).grid(row=8, column=1, sticky="w")
    blue_text.set(result[1][2])
    om_blue = OptionMenu(wd[3], blue_color, *color_array)
    blue_color.set(result[1][1])
    om_blue.config(width=13, anchor="w")
    om_blue.grid(row=8, column=2, sticky="w")
    Label(wd[3], text="blue").grid(row=8, column=3, sticky="w")
    Label(wd[3], text="{}".format(projvar.ns_code['green'])).grid(row=9, column=0, sticky="w")  # green row
    Entry(wd[3], textvariable=green_text, width=10).grid(row=9, column=1, sticky="w")
    green_text.set(result[2][2])
    om_green = OptionMenu(wd[3], green_color, *color_array)
    green_color.set(result[2][1])
    om_green.config(width=13, anchor="w")
    om_green.grid(row=9, column=2, sticky="w")
    Label(wd[3], text="green").grid(row=9, column=3, sticky="w")
    Label(wd[3], text="{}".format(projvar.ns_code['brown'])).grid(row=10, column=0, sticky="w")  # brown row
    Entry(wd[3], textvariable=brown_text, width=10).grid(row=10, column=1, sticky="w")
    brown_text.set(result[3][2])
    om_brown = OptionMenu(wd[3], brown_color, *color_array)
    brown_color.set(result[3][1])
    om_brown.config(width=13, anchor="w")
    om_brown.grid(row=10, column=2, sticky="w")
    Label(wd[3], text="brown").grid(row=10, column=3, sticky="w")
    Label(wd[3], text="{}".format(projvar.ns_code['red'])).grid(row=11, column=0, sticky="w")  # red row
    Entry(wd[3], textvariable=red_text, width=10).grid(row=11, column=1, sticky="w")
    red_text.set(result[4][2])
    om_red = OptionMenu(wd[3], red_color, *color_array)
    red_color.set(result[4][1])
    om_red.config(width=13, anchor="w")
    om_red.grid(row=11, column=2, sticky="w")
    Label(wd[3], text="red").grid(row=11, column=3, sticky="w")
    Label(wd[3], text="{}".format(projvar.ns_code['black'])).grid(row=12, column=0, sticky="w")  # black row
    Entry(wd[3], textvariable=black_text, width=10).grid(row=12, column=1, sticky="w")
    black_text.set(result[5][2])
    om_black = OptionMenu(wd[3], black_color, *color_array)
    black_color.set(result[5][1])
    om_black.config(width=13, anchor="w")
    om_black.grid(row=12, column=2, sticky="w")
    Label(wd[3], text="black").grid(row=12, column=3, sticky="w")
    Label(wd[3], text=" ").grid(row=13)
    Button(wd[3], text="set", width=10, command=lambda: ns_config_apply(wd[0], text_array, fill_array)) \
        .grid(row=14, column=3)
    Label(wd[3], text=" ").grid(row=15)
    Label(wd[3], text="Restore Defaults").grid(row=16)
    Button(wd[3], text="reset", width=10, command=lambda: ns_config_reset(wd[0])).grid(row=17, column=3)
    button_back = Button(wd[4])
    button_back.config(text="Go Back", width=20, command=lambda: MainFrame().start(frame=wd[0]))
    if sys.platform == "win32":
        button_back.config(anchor="w")
    button_back.pack(side=LEFT)
    rear_window(wd)


def get_file_path(subject_path):  # Created for pdf splitter - gets a pdf file
    path = dir_filedialog()
    file_path = filedialog.askopenfilename(initialdir=path,
                                           filetypes=[("PDF files", "*.pdf")], title="Select PDF")  # get the pdf file
    subject_path.set(file_path)


def get_new_path(new_path):  # Created for pdf splitter - creates/overwrites a pdf file
    path = dir_filedialog()
    save_filename = filedialog.asksaveasfilename(initialdir=path,
                                                 filetypes=[("PDF files", "*.pdf")], title="Overwrite/Create PDF")
    new_path.set(save_filename)


# check for empty fields / return if there are any errors
def pdf_splitter_apply(frame, subject_path, firstpage, lastpage, new_path):
    if subject_path == "":
        messagebox.showerror("Klusterbox PDF Splitter",
                             "You must select a pdf file to split.",
                             parent=frame)
        return
    if new_path == "":
        messagebox.showerror("Klusterbox PDF Splitter",
                             "You must designate a destination"
                             " and a name for the df file you are creating.",
                             parent=frame)
        return
    # if the last characters are not .pdf then add the extension
    if new_path[-4:] != ".pdf":
        new_path = new_path + ".pdf"
    if firstpage > lastpage:
        messagebox.showerror("Klusterbox PDF Splitter",
                             "The First Page of the document can not be "
                             "higher than the Last Page.",
                             parent=frame)
        return
    try:
        pdf = PdfFileReader(subject_path, "rb")
        pdf_writer = PdfFileWriter()
        for page in range(firstpage - 1, lastpage):
            pdf_writer.addPage(pdf.getPage(page))
        with open(new_path, 'wb') as out:
            pdf_writer.write(out)
        if messagebox.askokcancel("Klusterbox PDF Splitter",
                                  "PDF file has been split sucessfully."
                                  "Do you want to open the pdf file?",
                                  parent=frame):
            if sys.platform == "win32":
                os.startfile(new_path)
            if sys.platform == "linux":
                subprocess.call(["xdg-open", new_path])
            if sys.platform == "darwin":
                subprocess.call(["open", new_path])
    except:
        messagebox.showerror("Klusterbox PDF Splitter",
                             "The PDF splitting has failed. \n"
                             "It could be that that the pages set to be split don't exist \n"
                             "or \n"
                             "the pdf can't be split by this program due to formatting issues. \n"
                             "For better results try www.sodapdf.com, google chrome or Adobe Acrobat "
                             "Pro DC",
                             parent=frame)


def pdf_splitter(frame):  # PDF Splitter
    wd = front_window(frame)
    Label(wd[3], text="PDF Splitter", font=macadj("bold", "Helvetica 18"), anchor="w") \
        .grid(row=1, column=1, columnspan=4, sticky="w")
    Label(wd[3], text="").grid(row=2)
    Label(wd[3], text="Select pdf file you want to split:") \
        .grid(row=3, column=1, columnspan=4, sticky="w")
    subject_path = StringVar(wd[3])
    Entry(wd[3], textvariable=subject_path, width=macadj(95, 50)).grid(row=4, column=1, columnspan=4)
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
    Entry(wd[3], textvariable=new_path, width=macadj(95, 50)) \
        .grid(row=12, column=1, columnspan=4, sticky="w")
    Button(wd[3], text="Select", width="10", command=lambda: get_new_path(new_path)) \
        .grid(row=13, column=1, sticky="w")
    Label(wd[3], text="").grid(row=14)
    Label(wd[3], text="If all fields are filled out, split the file.") \
        .grid(row=15, column=1, columnspan=3, sticky="w")
    Button(wd[3], text="Split PDF", width="10",
           command=lambda: pdf_splitter_apply(
               wd[0],
               subject_path.get().strip(),
               firstpage.get(),
               lastpage.get(),
               new_path.get().strip())) \
        .grid(row=15, column=4, sticky="e")
    button_back = Button(wd[4])
    button_back.config(text="Go Back", width=20, command=lambda: MainFrame().start(frame=wd[0]))
    if sys.platform == "win32":
        button_back.config(anchor="w")
    button_back.pack(side=LEFT)
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
    sql = "SELECT tolerance FROM tolerances WHERE category ='%s'" % "pdf_error_rpt"
    result = inquire(sql)
    wd = front_window(frame)
    Label(wd[3], text="PDF Converter Settings", font=macadj("bold", "Helvetica 18"), anchor="w") \
        .grid(row=0, sticky="w", columnspan=4)
    Label(wd[3], text=" ").grid(row=1, column=0)
    Label(wd[3], text="Generate Reports for PDF Converter").grid(row=2, sticky="w", columnspan=4)
    Label(wd[3], text=" ").grid(row=3, column=0)
    Label(wd[3], text="Error Report", width=15, anchor="w").grid(row=4, column=0, sticky="w")
    error_selection = StringVar(wd[3])
    om_error = OptionMenu(wd[3], error_selection, "on", "off")  # option menu configuration below
    om_error.grid(row=4, column=1)
    error_selection.set(result[0][0])
    sql = "SELECT tolerance FROM tolerances WHERE category ='%s'" % "pdf_raw_rpt"
    result = inquire(sql)
    Label(wd[3], text="Raw Output Report", width=15, anchor="w").grid(row=5, column=0, sticky="w")
    raw_selection = StringVar(wd[3])
    om_raw = OptionMenu(wd[3], raw_selection, "on", "off")  # option menu configuration below
    om_raw.grid(row=5, column=1)
    raw_selection.set(result[0][0])
    Label(wd[3], text=" ").grid(row=6, column=0)
    # allow user to read from a text file to bypass the pdfminer
    Label(wd[3], text="Generate Reports from Text file").grid(row=7, sticky="w", columnspan=4)
    Label(wd[3], text="     (where a text file of pdfminer output has been generated)") \
        .grid(row=8, sticky="w", columnspan=4)
    Label(wd[3], text=" ").grid(row=9, column=0)
    sql = "SELECT tolerance FROM tolerances WHERE category ='%s'" % "pdf_text_reader"
    result = inquire(sql)
    Label(wd[3], text="Read from txt file", width=15, anchor="w").grid(row=10, column=0, sticky="w")
    txt_selection = StringVar(wd[3])
    om_txt = OptionMenu(wd[3], txt_selection, "on", "off")
    om_txt.grid(row=10, column=1)  # option menu configuration below
    txt_selection.set(result[0][0])
    Label(wd[3], text=" ").grid(row=11, column=0)
    if sys.platform == "darwin":  # option menu configuration
        om_error.config(width=5)
        om_raw.config(width=5)
        om_txt.config(width=5)
    else:
        om_error.config(width=5, anchor="w")
        om_raw.config(width=5, anchor="w")
        om_txt.config(width=5, anchor="w")
    Button(wd[3], text="set", width=10, command=lambda:
    pdf_converter_settings_apply(wd[0], error_selection, raw_selection, txt_selection)) \
        .grid(row=12, column=2)
    button = Button(wd[4])
    button.config(text="Go Back", width=20, command=lambda: MainFrame().start(frame=wd[0]))
    if sys.platform == "win32":
        button.config(anchor="w")
    button.pack(side=LEFT)
    rear_window(wd)


def pdf_converter_pagecount(filepath):  # gives a page count for pdf_to_text
    file = open(filepath, 'rb')
    parser = PDFParser(file)
    document = PDFDocument(parser)
    page_count = resolve1(document.catalog['Pages'])['Count']  # This will give you the count of pages
    return page_count


def pdf_to_text(frame, filepath):  # Called by pdf_converter() to read pdfs with pdfminer
    codec = 'utf-8'
    password = ""
    maxpages = 0
    caching = (True, True)
    pagenos = set()
    laparams = (
        LAParams(
            line_overlap=.1,  # best results
            char_margin=2,
            line_margin=.5,
            word_margin=.5,
            boxes_flow=0,
            detect_vertical=True,
            all_texts=True),
        LAParams(
            line_overlap=.5,  # default settings
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
            titlebar_icon(pb_root)  # place icon in titlebar
            Label(pb_root, text="This process takes several minutes. Please wait for results.") \
                .grid(row=0, column=0, columnspan=2, sticky="w")
            pb_label = Label(pb_root, text="Reading PDF: ")  # make label for progress bar
            pb_label.grid(row=1, column=0, sticky="w")
            pb = ttk.Progressbar(pb_root, length=350, mode="determinate")  # create progress bar
            pb.grid(row=1, column=1, sticky="w")
            pb_text = Label(pb_root, text="", anchor="w")
            pb_text.grid(row=2, column=0, columnspan=2, sticky="w")
            pb["maximum"] = page_count  # set length of progress bar
            pb.start()
            count = 0
            for page in PDFPage.get_pages(filein, pagenos, maxpages=maxpages, password=password, caching=caching[i],
                                          check_extractable=True):
                interpreter.process_page(page)
                pb["value"] = count  # increment progress bar
                pb_text.config(text="Reading page: {}/{}".format(count, page_count))
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
        text = text.replace("", "")
        page = text.split("")  # split the document into page
        result = re.search("Restricted USPS T&A Information(.*)Employee Everything Report", page[0], re.DOTALL)
        try:
            station = result.group(1).strip()
            break
        except:
            if i < 1:
                result = messagebox.askokcancel("Klusterbox PDF Converter",
                                                "PDF Conversion has failed and will not generate a file.  \n\n"
                                                "We will try again.",
                                                parent=frame)
                if not result:
                    return text
            else:
                messagebox.showerror("Klusterbox PDF Converter",
                                     "PDF Conversion has failed and will not generate a file.  \n\n"
                                     "You will either have to obtain the Employee Everything Report "
                                     "in the csv format from management or manually enter in the "
                                     "information",
                                     parent=frame)

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


def pdf_converter(frame):
    date_holder = []
    # inquire as to if the pdf converter reports have been opted for by the user
    sql = "SELECT tolerance FROM tolerances WHERE category ='%s'" % "pdf_error_rpt"
    result = inquire(sql)
    gen_error_report = result[0][0]
    sql = "SELECT tolerance FROM tolerances WHERE category ='%s'" % "pdf_raw_rpt"
    result = inquire(sql)
    gen_raw_report = result[0][0]
    starttime = time.time()  # start the timer
    # make it possible for user to select text file
    sql = "SELECT tolerance FROM tolerances WHERE category ='%s'" % "pdf_text_reader"
    result = inquire(sql)
    allow_txt_reader = result[0][0]
    if allow_txt_reader == "on":
        preference = messagebox.askyesno("PDF Converter",
                                         "Did you want to read from a text file of data output by pdfminer?",
                                         parent=frame)
    else:
        preference = False
    if not preference:  # user opts to read from pdf file
        path = dir_filedialog()
        file_path = filedialog.askopenfilename(initialdir=path,
                                               filetypes=[("PDF files", "*.pdf")])  # get the pdf file
        new_file_path = pdf_converter_path_generator(file_path, "_kbpc", ".csv")  # generate csv file name and path
        short_file_name = pdf_converter_short_name(new_file_path)
        # if the file path already exist - ask for confirmation
        if os.path.exists(new_file_path):
            if not messagebox.askokcancel("Possible File Name Discrepancy",
                                          "There is already a file named {}. "
                                          "If you proceed, the file will be overwritten. "
                                          "Did you want to proceed?".format(short_file_name),
                                          parent=frame):
                return
        # warn user that the process can take several minutes
        if not messagebox.askokcancel("PDF Converter", "This process will take several minutes. "
                                                       "Did you want to proceed?",
                                      parent=frame):
            return
        else:
            text = pdf_to_text(frame, file_path)  # read the pdf with pdfminer
    else:  # user opts to read from text file
        path = dir_filedialog()
        file_path = filedialog.askopenfilename(initialdir=path,
                                               filetypes=[("text files", "*.txt")])  # get the pdf file
        new_file_path = pdf_converter_path_generator(file_path, "_kbpc", ".csv")  # generate csv file name and path
        short_file_name = pdf_converter_short_name(new_file_path)
        # if the file path already exist - ask for confirmation
        if os.path.exists(new_file_path):
            if not messagebox.askokcancel(
                    "Possible File Name Discrepancy",
                    "There is already a file named {}. If you proceed, the file will be overwritten. "
                    "Did you want to proceed?".format(short_file_name),
                    parent=frame):
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
    text = text.replace("", "")
    page = text.split("")  # split the document into pages
    # whole_line = []
    page_num = 1  # initialize var to count pages
    eid_count = 0  # initialize var to count underscore dash items
    # underscore_slash = []  # arrays for building daily array
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
    daily_array_days = []  # build an array of formatted days with just month/ day
    result = re.search('Restricted USPS T&A Information(.*?)Employee Everything Report', page[0], re.DOTALL)
    try:
        station = result.group(1).strip()
    except:
        messagebox.showerror("Klusterbox PDF Converter",
                             "This file does not appear to be an Employee Everything Report. \n\n"
                             "The PDF Converter will not generate a file",
                             parent=frame)
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
    titlebar_icon(pb_root)  # place icon in titlebar
    Label(pb_root, text="This process takes several minutes. Please wait for results.").pack(anchor="w", padx=20)
    pb_label = Label(pb_root, text="Translating PDF: ")  # make label for progress bar
    pb_label.pack(anchor="w", padx=20)
    pb = ttk.Progressbar(pb_root, length=400, mode="determinate")  # create progress bar
    pb.pack(anchor="w", padx=20)
    pb["maximum"] = len(page) - 1  # set length of progress bar
    pb.start()
    pb_count = 0
    for a in page:
        if gen_error_report == "on":
            kbpc_rpt.write(
                "\nxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
                "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx\n")
        if a[0:6] == "Report" or a[0:6] == "":
            pass
        else:
            if gen_error_report == "on":
                kbpc_rpt.write("Out of Sequence Problem!\n")
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
        # get the pay period
        try:
            result = re.search("YrPPWk:\nSub-Unit:\n\n(.*)\n", a)
            yyppwk = result.group(1)
        except:
            try:
                result = re.search("YrPPWk:\n\n(.*)\n\nFin. #:", a)
                yyppwk = result.group(1)
            except:
                try:
                    result = re.findall(r'[0-9][0-9][0-9][0-9]-[0-9][0-9]-[0-9]', text)
                    yyppwk = result[-1]
                except:
                    pass
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
            if lookfortimes:
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
            # underscore_slash = []
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
                                daily_array.append(
                                    finance_holder)  # skip getting the route and create append daily array
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
                    # solve for salih problem / missing time zone in ...
                    elif len(time_holder) != 0 and unprocessedrings != "":
                        unprocessed_counter += 1  # unprocessed rings
                        salih_rpt.append(lastname)
                    time_holder = []
                    # look for time following date/mv desig
                    if re.match(r" [0-2][0-9]\.[0-9][0-9]$", e) and len(date_holder) != 0:
                        date_holder.append(e)
                        time_holder = date_holder
                    # look for items in franklin array to solve for franklin problem
                    if len(franklin_array) > 0 and re.match(r"[0-1][0-9]\/[0-3][0-9]$",
                                                            e):  # if franklin array and date
                        frank = franklin_array.pop(0)  # pop out the earliest mv desig
                        mv_holder = [eid, frank]
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
                        if eid_label:
                            found_days.append(e)
                        if not eid_label:
                            foundday_holder.append(e)
                    if e == "Processed Clock Rings":
                        eid_count = 0
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
                            ordered_days = ["Saturday", "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday",
                                            "Friday"]
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
                                                    pp_days[daily_array_days.index(array[2])].strftime(
                                                        "%d-%b-%y").upper()),
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
                    if lookforfi:  # look for first initial
                        if re.fullmatch("[A-Z]\s[A-Z]", e) or re.fullmatch("([A-Z])", e):
                            if gen_error_report == "on":
                                input = "FI: {}\n".format(e)
                                kbpc_rpt.write(input)
                            fi = e
                            lookforfi = False
                    if lookforname:  # look for the name
                        if re.fullmatch(r"([A-Z]+)", e) \
                                or re.fullmatch(r"([A-Z]+.[A-Z]+)", e) \
                                or re.fullmatch(r"([A-Z]+.[A-Z]+.[A-Z]+)", e) \
                                or re.fullmatch(r"([A-Z]+.[A-Z]+.[A-Z]+.[A-Z]+)", e) \
                                or re.fullmatch(r"([A-Z]+.[A-Z]+.[A-Z]+.[A-Z]+.[A-Z]+)", e):
                            lastname = e.replace("'", " ")
                            if gen_error_report == "on":
                                input = "Name: {}\n".format(e)
                                kbpc_rpt.write(input)
                            lookforname = False
                            lookforfi = True
                    if re.match(r"\s[0-9]{2}\-[0-9]$", e):  # find the job or d/a code - there might be two
                        jobs.append(e)
                    if lookfor2route:  # look for temp route
                        if re.match(r"[0-9]{6}$", e):
                            routes.append(e)  # add route to routes array
                        lookfor2route = False
                    if lookforroute:  # look for main route
                        if re.match(r"[0-9]{6}$", e):  #
                            routes.append(e)  # add route to routes array
                            lookfor2route = True
                        lookforroute = False
                    if e == "Route #":  # set trap to catch route # on the next line
                        lookforroute = True
                    if lookfor2level:  # intercept the second level
                        if re.match(r"[0-9]{2}$", e):
                            level.append(e)
                        lookfor2level = False
                    if lookforlevel:  # intercept the level
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
            # if the route count is less than the jobs count, fill the route count
            routes = PdfConverterFix(routes).route_filler(len(jobs))
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
        if len(daily_underscoreslash) > 0:  # bind all underscore slash items in one array
            underscore_slash_result = sum(daily_underscoreslash, [])
        if mcgrath_indicator and len(underscore_slash_result) > 0:  # solve for mcgrath indicator
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
                if gen_error_report == "on":
                    kbpc_rpt.write("MCGRATH ERROR DETECTED!!!\n")
            # if mcgrath_indicator == False:
            count += 2
        if mcgrath_carryover in daily_array:  # if there is a carryover, remove the daily array item from the list
            daily_array.remove(mcgrath_carryover)
        if not mcgrath_indicator and mcgrath_carryover != "":  # if there is a carryover to be added
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
        if not mcgrath_indicator:
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
            if len(daily_array) == eid_count / 2:
                pass
            # Solve for Unruh error / when a underscore dash is missing after unprocessed rings
            elif len(daily_array) == max((eid_count + 2) / 2, 0):
                if gen_error_report == "on":
                    input = "Unruh Mod emp id counter: {}\n".format(eid_count + 2)
                    kbpc_rpt.write(input)
                    kbpc_rpt.write("UNRUH PROBLEM DETECTED!!!")
                unruh_rpt.append(lastname)
            else:
                if gen_error_report == "on":
                    kbpc_rpt.write(
                        "FRANKLIN ERROR DETECTED!!! ALERT! (Unprocessed counter)!\n")
                unresolved.append(lastname)
        else:
            if len(daily_array) != max(eid_count / 2, 0):
                if gen_error_report == "on":
                    kbpc_rpt.write("FRANKLIN ERROR DETECTED!!! ALERT! ALERT!\n")
                unresolved.append(lastname)
        if base_chg + 1 != len(found_days):  # add to basecounter error array
            to_add = (lastname, base_chg, len(found_days))
            if len(found_days) > 0:
                basecounter_error.append(to_add)
        if gen_error_report == "on":
            input = "daily array lenght: {}\n".format(len(daily_array))
            kbpc_rpt.write(input)
        # initialize arrays
        found_days = []
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
            "Salih Problem: Unprocessed rings are missing a timezone, so that unprocessed rings counter is not"
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
    if len(failed) > 0:  # create messagebox to show any errors
        failed_daily = ""
        for f in failed:
            failed_daily = failed_daily + " \n " + f
        messagebox.showerror("Klusterbox PDF Converter",
                             "Errors have occured for the following carriers {}."
                             .format(failed_daily),
                             parent=frame)
    # create messagebox for completion
    messagebox.showinfo("Klusterbox PDF Converter",
                        "The PDF Convertion is complete. "
                        "The file name is {}. ".format(short_file_name),
                        parent=frame)


def informalc_grvchange(frame, passed_result, old_num, new_num):
    l_passed_result = [list(x) for x in passed_result]  # chg tuple of tuples to list of lists
    if messagebox.askokcancel("Grievance Number Change",
                              "This will change the grievance number from {} to {} in all "
                              "records. Are you sure you want to proceed?".format(old_num, new_num.get()),
                              parent=frame):
        if new_num.get().strip() == "":
            messagebox.showerror("Invalid Data Entry",
                                 "You must enter a grievance number",
                                 parent=frame)
            return "fail"
        if not new_num.get().isalnum():
            messagebox.showerror("Invalid Data Entry",
                                 "The grievance number can only contain numbers and letters. No other "
                                 "characters are allowed",
                                 parent=frame)
            return "fail"
        if len(new_num.get()) < 8:
            messagebox.showerror("Invalid Data Entry",
                                 "The grievance number must be at least eight characters long",
                                 parent=frame)
            return "fail"
        if len(new_num.get()) > 16:
            messagebox.showerror("Invalid Data Entry",
                                 "The grievance number must not exceed 16 characters in length.",
                                 parent=frame)
            return "fail"
        sql = "SELECT grv_no FROM informalc_grv WHERE grv_no = '%s'" % new_num.get().lower()
        result = inquire(sql)
        if result:
            messagebox.showerror("Grievance Number Error",
                                 "This number is already being used for another grievance.",
                                 parent=frame)
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


def informalc_edit_apply(frame, grv_no, incident_start, incident_end, date_signed, station, gats_number, docs,
                         description, lvl):
    check = informalc_check_grv_2(frame, incident_start, incident_end, date_signed, gats_number, description)
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
        messagebox.showerror("Data Entry Error",
                             "The Incident Start Date can not be later that the Incident End "
                             "Date.",
                             parent=frame)
        return
    if dt_dates[0] > dt_dates[2]:
        messagebox.showerror("Data Entry Error",
                             "The Incident Start Date can not be later that the Date Signed.",
                             parent=frame)
        return
    sql = "UPDATE informalc_grv SET indate_start='%s',indate_end='%s',date_signed='%s',station='%s',gats_number='%s'," \
          "docs='%s',description='%s', level='%s' WHERE grv_no='%s'" \
          % (dt_dates[0], dt_dates[1], dt_dates[2], station.get(), gats_number.get().strip(), docs.get(),
             description.get(), lvl.get(), grv_no.get())
    commit(sql)
    messagebox.showerror("Sucessful Update",
                         "Grievance number: {} succesfully updated.".format(grv_no.get()),
                         parent=frame)
    informalc_grvlist(frame)


def informalc_delete(frame, grv_no):
    check = messagebox.askokcancel("Delete Grievance",
                                   "Are you sure you want to delete his grievance and all the "
                                   "data associated with it?",
                                   parent=frame)
    if not check:
        return
    else:
        sql = "DELETE FROM informalc_grv WHERE grv_no='%s'" % grv_no.get()
        commit(sql)
        informalc_grvlist(frame)


def informalc_edit(frame, result, grv_num, msg):
    wd = front_window(frame)
    Label(wd[3], text="Informal C: Edit Grievance", font=macadj("bold", "Helvetica 18")).grid(row=0, columnspan=2,
                                                                                              sticky="w")
    Label(wd[3], text="").grid(row=1)
    Label(wd[3], text="Grievance Number: ", background=macadj("gray95", "grey"),
          fg=macadj("black", "white"), width=macadj(22, 20), anchor="w", height=macadj(1, 1)) \
        .grid(row=2, column=0, sticky="w")
    grv_no = StringVar(wd[0])
    Entry(wd[3], textvariable=grv_no, justify='right', width=macadj(20, 15)) \
        .grid(row=2, column=1, sticky="w")
    Button(wd[3], width=9, text="update", command=lambda:
    informalc_grvchange(wd[0], result, grv_num, grv_no)).grid(row=3, column=1, sticky="e")
    grv_no.set(grv_num)
    Label(wd[3], text="Incident Date", background=macadj("gray95", "grey"),
          fg=macadj("black", "white"), width=macadj(22, 20), anchor="w", height=macadj(1, 1)) \
        .grid(row=4, column=0, sticky="w")
    Label(wd[3], text="  Start (mm/dd/yyyy): ", background=macadj("gray95", "grey"),
          fg=macadj("black", "white"), width=macadj(22, 20), anchor="w", height=macadj(1, 1)) \
        .grid(row=5, column=0, sticky="w")
    incident_start = StringVar(wd[0])
    Entry(wd[3], textvariable=incident_start, justify='right', width=macadj(20, 15)) \
        .grid(row=5, column=1, sticky="w")
    Label(wd[3], text="  End (mm/dd/yyyy): ", background=macadj("gray95", "grey"),
          fg=macadj("black", "white"), width=macadj(22, 20), anchor="w", height=macadj(1, 1)) \
        .grid(row=6, column=0, sticky="w")
    incident_end = StringVar(wd[0])
    Entry(wd[3], textvariable=incident_end, justify='right', width=macadj(20, 15)) \
        .grid(row=6, column=1, sticky="w")
    Label(wd[3], text="Date Signed (mm/dd/yyyy): ", background=macadj("gray95", "grey"),
          fg=macadj("black", "white"), width=macadj(22, 20), anchor="w", height=macadj(1, 1)) \
        .grid(row=7, column=0, sticky="w")
    date_signed = StringVar(wd[0])
    Entry(wd[3], textvariable=date_signed, justify='right', width=macadj(20, 15)) \
        .grid(row=7, column=1, sticky="w")
    Label(wd[3], text="Settlement Level: ", background=macadj("gray95", "grey"),
          fg=macadj("black", "white"), width=macadj(22, 20), anchor="w", height=macadj(1, 1)) \
        .grid(row=8, column=0, sticky="w")  # select settlement level
    lvl = StringVar(wd[0])
    lvl_options = ("informal a", "formal a", "step b", "pre arb", "arbitration")
    lvl_om = OptionMenu(wd[3], lvl, *lvl_options)
    lvl_om.config(width=13)
    lvl_om.grid(row=8, column=1)
    Label(wd[3], text="Station: ", background=macadj("gray95", "grey"),
          fg=macadj("black", "white"), width=macadj(22, 20), anchor="w", height=macadj(1, 1)) \
        .grid(row=9, column=0, sticky="w")  # select a station
    station = StringVar(wd[0])
    station_options = projvar.list_of_stations
    if "out of station" in station_options:
        station_options.remove("out of station")
    station_om = OptionMenu(wd[3], station, *station_options)
    station_om.config(width=macadj(40, 34))
    station_om.grid(row=10, column=0, columnspan=2, sticky="e")
    Label(wd[3], text="GATS Number: ", background=macadj("gray95", "grey"),
          fg=macadj("black", "white"), width=macadj(22, 20), anchor="w", height=macadj(1, 1)) \
        .grid(row=11, column=0, sticky="w")
    gats_number = StringVar(wd[0])
    Entry(wd[3], textvariable=gats_number, justify='right', width=macadj(20, 15)) \
        .grid(row=11, column=1, sticky="w")
    Label(wd[3], text="Documentation: ", background=macadj("gray95", "grey"),
          fg=macadj("black", "white"), width=macadj(22, 20), anchor="w", height=macadj(1, 1)) \
        .grid(row=12, column=0, sticky="w")
    docs = StringVar(wd[0])
    doc_options = ("moot", "no", "partial", "yes", "incomplete", "verified")
    docs_om = OptionMenu(wd[3], docs, *doc_options)
    docs_om.config(width=13)
    docs_om.grid(row=12, column=1)
    Label(wd[3], text="Description: ", background=macadj("gray95", "grey"),
          fg=macadj("black", "white"), width=macadj(22, 20), anchor="w", height=macadj(1, 1)) \
        .grid(row=16, column=0, sticky="w")
    description = StringVar(wd[0])
    Entry(wd[3], textvariable=description, width=macadj(47, 36), justify='right') \
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
        if search[0][8] is None:
            lvl.set("unknown")
        else:
            lvl.set(search[0][8])
    Label(wd[3], text=" ").grid(row=20)
    Label(wd[3], text="Delete Grievance", background=macadj("gray95", "grey"),
          fg=macadj("black", "white"), width=macadj(22, 20), anchor="w", height=macadj(1, 1)) \
        .grid(row=21, column=0, sticky="w")
    Button(wd[3], text="Delete", width=9, command=lambda: informalc_delete(wd[0], grv_no)) \
        .grid(row=21, column=1, sticky="e")
    Label(wd[3], text=" ").grid(row=22)
    Label(wd[3], text=msg, fg="red", anchor="w").grid(row=23, column=0, columnspan=5, sticky="w")
    Button(wd[4], text="Go Back", width=macadj(19, 18),
           command=lambda: informalc_grvlist_result(wd[0], result)).grid(row=0, column=0)
    Button(wd[4], text="Enter", width=macadj(19, 18),
           command=lambda: informalc_edit_apply(wd[0], grv_no, incident_start,
                                                incident_end, date_signed, station, gats_number, docs, description,
                                                lvl)).grid(row=0, column=1)
    rear_window(wd)


def informalc_check_grv(frame, grv_no, incident_start, incident_end, date_signed, station, gats_number, description):
    if station.get() == "Select a Station":
        messagebox.showerror("Invalid Data Entry",
                             "You must select a station.",
                             parent=frame)
        return "fail"
    if grv_no.get().strip() == "":
        messagebox.showerror("Invalid Data Entry",
                             "You must enter a grievance number",
                             parent=frame)
        return "fail"
    if re.search('[^1234567890abcdefghijklmnopqrstuvwxyz:ABCDEFGHIJKLMNOPQRSTUVWXYZ,]', grv_no.get()):
        messagebox.showerror("Invalid Data Entry",
                             "The grievance number can only contain numbers and letters. No other "
                             "characters are allowed",
                             parent=frame)
        return "fail"
    if len(grv_no.get()) < 8:
        messagebox.showerror("Invalid Data Entry",
                             "The grievance number must be at least eight characters long",
                             parent=frame)
        return "fail"
    if len(grv_no.get()) > 20:
        messagebox.showerror("Invalid Data Entry",
                             "The grievance number must not exceed 20 characters in length.",
                             parent=frame)
        return "fail"
    check = informalc_check_grv_2(frame, incident_start, incident_end, date_signed, gats_number, description)
    return check


def informalc_check_grv_2(frame, incident_start, incident_end, date_signed, gats_number, description):
    dates = [incident_start, incident_end, date_signed]
    date_ids = ("starting incident date", "ending incident date", "date signed")
    i = 0
    for date in dates:
        d = date.get().split("/")
        if len(d) != 3:
            messagebox.showerror("Invalid Data Entry",
                                 "The date for the {} is not properly formatted.".format(date_ids[i]),
                                 parent=frame)
            return "fail"
        for num in d:
            if not num.isnumeric():
                messagebox.showerror("Invalid Data Entry",
                                     "The month, day and year for the {} "
                                     "must be numeric.".format(date_ids[i]),
                                     parent=frame)
                return "fail"
        if len(d[0]) > 2:
            messagebox.showerror("Invalid Data Entry",
                                 "The month for the {} must be no more than two digits"
                                 " long.".format(date_ids[i]),
                                 parent=frame)
            return "fail"
        if len(d[1]) > 2:
            messagebox.showerror("Invalid Data Entry",
                                 "The day for the {} must be no more than two digits"
                                 " long.".format(date_ids[i]),
                                 parent=frame)
            return "fail"
        if len(d[2]) != 4:
            messagebox.showerror("Invalid Data Entry",
                                 "The year for the {} must be four digits long."
                                 .format(date_ids[i]),
                                 parent=frame)
            return "fail"
        try:
            date = datetime(int(d[2]), int(d[0]), int(d[1]))
            valid_date = True
        except ValueError:
            valid_date = False
        if not valid_date:
            messagebox.showerror("Invalid Data Entry",
                                 "The date entered for {} is not a valid date."
                                 .format(date_ids[i]),
                                 parent=frame)
            return "fail"
        i += 1
    if len(gats_number.get()) > 50:
        messagebox.showerror("Invalid Data Entry",
                             "The GATS number is limited to no more than 20 characters. ",
                             parent=frame)
        return "fail"
    if gats_number.get().strip() != "":
        if not all(x.isalnum() or x.isspace() for x in gats_number.get()):
            messagebox.showerror("Invalid Data Entry",
                                 "The GATS number can only contain letters and numbers. No "
                                 "special characters are allowed.",
                                 parent=frame)
            return "fail"
    if description.get().strip() != "":
        if not all(x.isalnum() or x.isspace() for x in description.get()):
            messagebox.showerror("Invalid Data Entry",
                                 "The Description can only contain letters and numbers. No "
                                 "special characters are allowed.",
                                 parent=frame)
            return "fail"
        if len(description.get()) > 40:
            messagebox.showerror("Invalid Data Entry",
                                 "The Description is limited to no more than 40 characters. ",
                                 parent=frame)
            return "fail"
    return "pass"


def informalc_new_apply(frame, grv_no, incident_start, incident_end, date_signed, station, gats_number, docs,
                        description, lvl):
    check = informalc_check_grv(frame, grv_no, incident_start, incident_end, date_signed, station, gats_number,
                                description)
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
            messagebox.showerror("Data Entry Error",
                                 "The Incident Start Date can not be later that the Incident End "
                                 "Date.",
                                 parent=frame)
            return
        if dt_dates[0] > dt_dates[2]:
            messagebox.showerror("Data Entry Error",
                                 "The Incident Start Date can not be later that the Date Signed.",
                                 parent=frame)
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
                                 "create a duplicate.".format(grv_no.get()),
                                 parent=frame)
            return
        sql = "INSERT INTO informalc_grv (grv_no, indate_start, indate_end, date_signed, station, gats_number, docs," \
              "description, level) VALUES('%s','%s','%s','%s','%s','%s','%s','%s','%s')" \
              % (grv_no.get().lower(), dt_dates[0], dt_dates[1], dt_dates[2], station.get(), gats_number.get().strip(),
                 docs.get(), description.get(), lvl.get())
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
            if not added and rec[2] == station:
                carrier_list.append(rec[1])
                added = True
        if not added and len(before_start) > 0:
            if before_start[0][2] == station:
                carrier_list.append(rec[1])
    return carrier_list


def informalc_addnames(grv_no, c_list, listbox):
    for index in listbox:
        sql = "INSERT INTO informalc_awards (grv_no,carrier_name,hours,rate,amount) VALUES('%s','%s','%s','%s','%s')" \
              % (grv_no, c_list[int(index)], '', '', '')
        commit(sql)


def informalc_root(passed_result, grv_no):
    start = None
    end = None
    global informalc_newroot  # initialize the global
    new_root = Tk()
    informalc_newroot = new_root  # set the global
    new_root.title("KLUSTERBOX")
    titlebar_icon(new_root)  # place icon in titlebar
    x_position = projvar.root.winfo_x() + 450
    y_position = projvar.root.winfo_y() - 25
    new_root.geometry("%dx%d+%d+%d" % (240, 600, x_position, y_position))
    n_f = Frame(new_root)
    n_f.pack()
    n_buttons = Canvas(n_f)  # button bar
    n_buttons.pack(fill=BOTH, side=BOTTOM)
    Label(n_f, text="Add Carriers", font=macadj("bold", "Helvetica 18")).pack(anchor="w")
    Label(n_f, text="").pack()
    scrollbar = Scrollbar(n_f, orient=VERTICAL)
    listbox = Listbox(n_f, selectmode="multiple", yscrollcommand=scrollbar.set)
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


def informalc_deletename(frame, passed_result, grv_no, ids):
    sql = "DELETE FROM informalc_awards WHERE rowid='%s'" % ids
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
        id_no = var_id[i].get()  # simplify variable names
        name = var_name[i].get()
        hours = var_hours[i].get().strip()
        rate = var_rate[i].get().strip()
        amount = var_amount[i].get().strip()
        if hours and amount:
            messagebox.showerror("Data Input Error",
                                 "Input error for {} in row {}. You can not enter both hours and "
                                 "amount. You can only enter one or another, but not both. "
                                 "Awards can be in the form of "
                                 "hours at a given rate OR an amount.".format(name, str(i + 1)),
                                 parent=frame)
            pb_label.destroy()  # destroy the label for the progress bar
            pb.destroy()  # destroy the progress bar
            return
        if rate and amount:
            messagebox.showerror("Data Input Error",
                                 "Input error for {} in row {}. You can not enter both rate and "
                                 "amount. You can only enter one or another, but not both. "
                                 "Awards can be in the form of "
                                 "hours at a given rate OR an amount.".format(name, str(i + 1)),
                                 parent=frame)
            pb_label.destroy()  # destroy the label for the progress bar
            pb.destroy()  # destroy the progress bar
            return
        if hours and not rate:
            messagebox.showerror("Data Input Error",
                                 "Input error for {} in row {}. Hours must be a accompanied by a "
                                 "rate.".format(name, str(i + 1)),
                                 parent=frame)
            pb_label.destroy()  # destroy the label for the progress bar
            pb.destroy()  # destroy the progress bar
            return
        if rate and not hours:
            messagebox.showerror("Data Input Error",
                                 "Input error for {} in row {}. Rate must be a accompanied by a "
                                 "hours.".format(name, str(i + 1)),
                                 parent=frame)
            pb_label.destroy()  # destroy the label for the progress bar
            pb.destroy()  # destroy the progress bar
            return
        if hours and isfloat(hours) == False:
            messagebox.showerror("Data Input Error",
                                 "Input error for {} in row {}. Hours must be a number."
                                 .format(name, str(i + 1)),
                                 parent=frame)
            pb_label.destroy()  # destroy the label for the progress bar
            pb.destroy()  # destroy the progress bar
            return
        if hours and '.' in hours:
            s_hrs = hours.split(".")
            if len(s_hrs[1]) > 2:
                messagebox.showerror("Data Input Error",
                                     "Input error for {} in row {}. Hours must have no "
                                     "more than 2 decimal places.".format(name, str(i + 1)),
                                     parent=frame)
                pb_label.destroy()  # destroy the label for the progress bar
                pb.destroy()  # destroy the progress bar
                return
        if rate and amount:
            messagebox.showerror("Data Input Error",
                                 "Input error for {} in row {}. You can not enter both rate and "
                                 "amount. You can only enter one or another, but not both. "
                                 "Awards can be in the form of "
                                 "hours at a given rate OR an amount.".format(name, str(i + 1)),
                                 parent=frame)
            pb_label.destroy()  # destroy the label for the progress bar
            pb.destroy()  # destroy the progress bar
            return
        if rate and amount:
            messagebox.showerror("Data Input Error",
                                 "Input error for {} in row {}. You can not enter both rate and "
                                 "amount. You can only enter one or another, but not both. "
                                 "Awards can be in the form of "
                                 "hours at a given rate OR an amount.".format(name, str(i + 1)),
                                 parent=frame)
            pb_label.destroy()  # destroy the label for the progress bar
            pb.destroy()  # destroy the progress bar
            return
        if rate and not isfloat(rate):
            messagebox.showerror("Data Input Error", "Input error for {} in row {}. Rates must be a number."
                                 .format(name, str(i + 1)),
                                 parent=frame)
            pb_label.destroy()  # destroy the label for the progress bar
            pb.destroy()  # destroy the progress bar
            return
        if rate and '.' in rate:
            s_rate = rate.split(".")
            if len(s_rate[1]) > 2:
                messagebox.showerror("Data Input Error",
                                     "Input error for {} in row {}. Rates must have no "
                                     "more than 2 decimal places.".format(name, str(i + 1)),
                                     parent=frame)
                pb_label.destroy()  # destroy the label for the progress bar
                pb.destroy()  # destroy the progress bar
                return
        if rate and float(rate) > 10:
            messagebox.showerror("Data Input Error",
                                 "Input error for {} in row {}. Values greater than 10 are not "
                                 "accepted. \n"
                                 "Note the following rates would be expressed as: \n "
                                 "additional %50         .50 or just .5 \n"
                                 "straight time rate     1.00 or just 1 \n"
                                 "overtime rate          1.50 or 1.5 \n"
                                 "penalty rate           2.00 or just 2".format(name, str(i + 1)),
                                 parent=frame)
            pb_label.destroy()  # destroy the label for the progress bar
            pb.destroy()  # destroy the progress bar
            return
        if amount and isfloat(amount) == False:
            messagebox.showerror("Data Input Error",
                                 "Input error for {} in row {}. Amounts can only be expressed as "
                                 "numbers. No special characters, such as $ are allowed.".format(name, str(i + 1)),
                                 parent=frame)
            pb_label.destroy()  # destroy the label for the progress bar
            pb.destroy()  # destroy the progress bar
            return
        if amount and '.' in amount:
            s_amt = amount.split(".")
            if len(s_amt[1]) > 2:
                messagebox.showerror("Data Input Error",
                                     "Input error for {} in row {}. "
                                     "Amounts must have no more than 2 decimal places.".format(name, str(i + 1)),
                                     parent=frame)
                pb_label.destroy()  # destroy the label for the progress bar
                pb.destroy()  # destroy the progress bar
                return
        sql = "UPDATE informalc_awards SET hours='%s',rate='%s',amount='%s' WHERE rowid='%s'" % (
            hours, rate, amount, id_no)
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
    Label(wd[3], text="Add/Update Settlement Awards", font=macadj("bold", "Helvetica 18")) \
        .grid(row=0, column=0, sticky="w", columnspan=4)
    Label(wd[3], text=" ".format(informalc_addframe)).grid(row=1, column=0)
    Label(wd[3], text="   Grievance Number: {}".format(grv_no), fg="blue") \
        .grid(row=2, column=0, sticky="w", columnspan=4)
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
                                                    var_rate, var_amount)).grid(row=0, column=1)
    rear_window(wd)


def informalc_call_grvlist_result(frame, passed_result):
    try:
        informalc_newroot.destroy()
    except TclError:
        pass
    informalc_grvlist_result(frame, passed_result)


def informalc_addaward(frame, passed_result, grv_no):
    informalc_root(passed_result, grv_no)
    informalc_addaward2(frame, passed_result, grv_no)


def informalc_rptgrvsum(frame, result):
    if len(result) > 0:
        result = list(result)
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = "infc_grv_list" + "_" + stamp + ".txt"
        report = open(dir_path('infc_grv') + filename, "w")
        report.write("Settlement List\n\n")
        i = 1
        for sett in result:
            sett = list(sett)  # correct for legacy problem of NULL Settlement Levels
            if sett[8] is None:
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
                if rec[2]:
                    hour = float(rec[2])
                if rec[3]:
                    rate = float(rec[3])
                if rec[4]:
                    amt = float(rec[4])
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
            cc = 1
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
                    '    {:<5}{:<22}{:>6}{:>10}{:>10}{:>12}\n'.format(str(cc), rec[1], hours, rate, adj, amt))
                cc += 1
            report.write("    -----------------------------------------------------------------\n")
            report.write("         {:<38}{:>10}\n".format("Awards adjusted to straight time", "{0:.2f}"
                                                          .format(float(awardxhour))))
            report.write("         {:<38}{:>22}\n".format("Awards as flat dollar amount", "{0:.2f}"
                                                          .format(float(awardxamt))))
            report.write("\n\n\n")
            i += 1
        report.close()
        try:
            if sys.platform == "win32":
                os.startfile(dir_path('infc_grv') + filename)
            if sys.platform == "linux":
                subprocess.call(["xdg-open", 'kb_sub/infc_grv/' + filename])
            if sys.platform == "darwin":
                subprocess.call(["open", dir_path('infc_grv') + filename])
        except PermissionError:
            messagebox.showerror("Report Generator", "The report was not generated.", parent=frame)


def informalc_bycarriers(frame, result):
    unique_carrier = informalc_uniquecarrier(result)
    unique_grv = []  # get a list of all grv numbers in search range
    for grv in result:
        if grv[0] not in unique_grv:
            unique_grv.append(grv[0])  # put these in "unique_grv"
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = "infc_grv_list" + "_" + stamp + ".txt"
    report = open(dir_path('infc_grv') + filename, "w")
    report.write("Settlement Report By Carriers\n\n")
    for name in unique_carrier:
        report.write("{:<30}\n\n".format(name))
        report.write("        Grievance Number    Hours    Rate    Adjusted      Amount       docs       level\n")
        report.write("    ------------------------------------------------------------------------------------\n")
        results = []
        for ug in unique_grv:  # do search for each grievance in list of unique grievances
            sql = "SELECT informalc_awards.grv_no, informalc_awards.hours, informalc_awards.rate, " \
                  "informalc_awards.amount, informalc_grv.docs, informalc_grv.level " \
                  "FROM informalc_awards, informalc_grv " \
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
            if r[5] is None or r[5] == "unknown":
                r[5] = "---"
            report.write("    {:<4}{:<17}{:>8}{:>8}{:>12}{:>12}{:>11}{:>12}\n"
                         .format(str(i), r[0], hours, rate, adj, amt, r[4], r[5]))
            i += 1
        report.write("    ------------------------------------------------------------------------------------\n")
        t_adj = "{0:.2f}".format(float(total_adj))
        t_amt = "{0:.2f}".format(float(total_amt))
        report.write("        {:<34}{:>11}\n".format("Total hours as straight time", t_adj))
        report.write("        {:<34}{:>23}\n".format("Total as flat dollar amount", t_amt))
        report.write("\n\n\n")
    report.close()
    try:
        if sys.platform == "win32":
            os.startfile(dir_path('infc_grv') + filename)
        if sys.platform == "linux":
            subprocess.call(["xdg-open", 'kb_sub/infc_grv/' + filename])
        if sys.platform == "darwin":
            subprocess.call(["open", dir_path('infc_grv') + filename])
    except PermissionError:
        messagebox.showerror("Report Generator", "The report was not generated.", parent=frame)


def informalc_apply_bycarrier(frame, result, names, cursor):
    if len(cursor) == 0:
        return
    unique_grv = []  # get a list of all grv numbers in search range
    for grv in result:
        if grv[0] not in unique_grv:
            unique_grv.append(grv[0])  # put these in "unique_grv"
    name = names[cursor[0]]
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = "infc_grv_list" + "_" + stamp + ".txt"
    report = open(dir_path('infc_grv') + filename, "w")
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
        if r[5] is None or r[5] == "unknown":
            r[5] = "---"
        report.write("    {:<4}{:<18}{:>7}{:>8}{:>12}{:>12}{:>11}{:>12}\n"
                     .format(str(i), r[0], hours, rate, adj, amt, r[4], r[5]))
        i += 1
    report.write("    ------------------------------------------------------------------------------------\n")
    t_adj = "{0:.2f}".format(float(total_adj))
    t_amt = "{0:.2f}".format(float(total_amt))
    report.write("        {:<34}{:>11}\n".format("Total hours as straight time", t_adj))
    report.write("        {:<34}{:>23}\n".format("Total as flat dollar amount", t_amt))
    report.close()
    try:
        if sys.platform == "win32":
            os.startfile(dir_path('infc_grv') + filename)
        if sys.platform == "linux":
            subprocess.call(["xdg-open", 'kb_sub/infc_grv/' + filename])
        if sys.platform == "darwin":
            subprocess.call(["open", dir_path('infc_grv') + filename])
    except PermissionError:
        messagebox.showerror("Report Generator", "The report was not generated.", parent=frame)


def informalc_bycarrier(frame, result):
    unique_carrier = informalc_uniquecarrier(result)
    wd = front_window(frame)
    Label(wd[3], text="Informal C: Select Carrier", font=macadj("bold", "Helvetica 18")).pack(anchor="w")
    Label(wd[3], text="").pack()
    scrollbar = Scrollbar(wd[3], orient=VERTICAL)
    listbox = Listbox(wd[3], selectmode="single", yscrollcommand=scrollbar.set)
    listbox.config(height=30, width=50)
    for name in unique_carrier:
        listbox.insert(END, name)
    scrollbar.config(command=listbox.yview)
    scrollbar.pack(side=RIGHT, fill=Y)
    listbox.pack(side=LEFT, expand=1)
    Button(wd[4], text="Go Back", width=20,
           command=lambda: informalc_grvlist_result(wd[0], result)).pack(side=LEFT)
    Button(wd[4], text="Report", width=20,
           command=lambda: informalc_apply_bycarrier
           (frame, result, unique_carrier, listbox.curselection())).pack(side=LEFT)
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


def informalc_rptbygrv(frame, grv_info):
    grv_info = list(grv_info)  # correct for legacy problem of NULL Settlement Levels
    if grv_info[8] is None:
        grv_info[8] = "unknown"
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = "infc_grv_list" + "_" + stamp + ".txt"
    report = open(dir_path('infc_grv') + filename, "w")
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
    cc = 1
    for rec in query:
        hour = 0.0
        rate = 0.0
        amt = 0
        if rec[2]:
            hour = float(rec[2])
        if rec[3]:
            rate = float(rec[3])
        if rec[4]:
            amt = float(rec[4])
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
        report.write('    {:<5}{:<22}{:>6}{:>10}{:>10}{:>12}\n'.format(str(cc), rec[1], hours, rate, adj, amt))
        cc += 1
    report.write("    -----------------------------------------------------------------\n")
    report.write("         {:<38}{:>10}\n".format("Awards adjusted to straight time", "{0:.2f}"
                                                  .format(float(awardxhour))))
    report.write("         {:<38}{:>22}\n".format("Awards as flat dollar amount", "{0:.2f}"
                                                  .format(float(awardxamt))))
    report.write("\n\n\n")
    report.close()
    try:
        if sys.platform == "win32":
            os.startfile(dir_path('infc_grv') + filename)
        if sys.platform == "linux":
            subprocess.call(["xdg-open", 'kb_sub/infc_grv/' + filename])
        if sys.platform == "darwin":
            subprocess.call(["open", dir_path('infc_grv') + filename])
    except PermissionError:
        messagebox.showerror("Report Generator", "The report was not generated.", parent=frame)


def informalc_grvlist_setsum(result):
    if len(result) > 0:
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = "infc_grv_list" + "_" + stamp + ".txt"
        report = open(dir_path('infc_grv') + filename, "w")
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
                if rec[2]:
                    hour = float(rec[2])
                if rec[3]:
                    rate = float(rec[3])
                if rec[4]:
                    amt = float(rec[4])
                if hour and rate:
                    awardxhour = awardxhour + (hour * rate)
                if amt:
                    awardxamt = awardxamt + amt
            sign = dt_converter(sett[3]).strftime("%m/%d/%Y")
            s_gats = sett[5].split(" ")
            if sett[8] is None or sett[8] == "unknown":
                lvl = "---"
            else:
                lvl = sett[8]
            # for gats_no in s_gats:
            for gi in range(len(s_gats)):
                if gi == 0:
                    total_hour += awardxhour
                    total_amt += awardxamt
                    report.write('{:>4}  {:<14}{:<12}{:<9}{:>11}{:>12}{:>12}{:>12}\n'
                                 .format(str(i), sett[0], sign, s_gats[gi], sett[6], lvl,
                                         "{0:.2f}".format(float(awardxhour)), "{0:.2f}".format(float(awardxamt))))
                if gi != 0:
                    report.write('{:<34}{:<12}\n'.format("", s_gats[gi]))
            if i % 3 == 0:
                report.write(
                    "      ----------------------------------------------------------------------------------\n")
            i += 1
        report.write("      ----------------------------------------------------------------------------------\n")
        report.write("{:<20}{:>58}\n".format("      Total Hours", "{0:.2f}".format(total_hour)))
        report.write("{:<20}{:>70}\n".format("      Total Dollars", "{0:.2f}".format(total_amt)))
        report.close()
        if sys.platform == "win32":
            os.startfile(dir_path('infc_grv') + filename)
        if sys.platform == "linux":
            subprocess.call(["xdg-open", 'kb_sub/infc_grv/' + filename])
        if sys.platform == "darwin":
            subprocess.call(["open", dir_path('infc_grv') + filename])


def informalc_grvlist_result(frame, result):
    wd = front_window(frame)
    Label(wd[3], text="Informal C: Search Results", font=macadj("bold", "Helvetica 18")) \
        .grid(row=0, column=0, columnspan=4, sticky="w")
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
        Label(wd[3], text=str(ii), anchor="w", width=macadj(4, 2)).grid(row=row, column=0)
        Button(wd[3], text=" " + r[0], anchor="w", width=macadj(14, 12), relief=RIDGE).grid(row=row, column=1)
        in_start = datetime.strptime(r[1], '%Y-%m-%d %H:%M:%S')
        in_end = datetime.strptime(r[2], '%Y-%m-%d %H:%M:%S')
        sign_date = datetime.strptime(r[3], '%Y-%m-%d %H:%M:%S')
        Button(wd[3], text=in_start.strftime("%b %d, %Y"), width=macadj(11, 10), anchor="w", relief=RIDGE) \
            .grid(row=row, column=2)
        Button(wd[3], text=in_end.strftime("%b %d, %Y"), width=macadj(11, 10), anchor="w", relief=RIDGE) \
            .grid(row=row, column=3)
        Button(wd[3], text=sign_date.strftime("%b %d, %Y"), width=macadj(11, 10), anchor="w", relief=RIDGE) \
            .grid(row=row, column=4)
        Button(wd[3], text="Edit", width=macadj(6, 5), relief=RIDGE,
               command=lambda x=r[0]: informalc_edit(wd[0], result, x, '')) \
            .grid(row=row, column=5)
        Button(wd[3], text="Report", width=macadj(6, 5), relief=RIDGE,
               command=lambda x=r: informalc_rptbygrv(wd[0], x)).grid(row=row, column=6)
        Button(wd[3], text=macadj("Enter Awards", "Awards"), width=macadj(10, 6), relief=RIDGE,
               command=lambda x=r[0]: informalc_addaward(wd[0], result, x)).grid(row=row, column=7)
        row += 1
        Label(wd[3], text="         {}".format(r[7]), anchor="w", fg="grey") \
            .grid(row=row, column=1, columnspan=5, sticky="w")
        row += 1
        ii += 1
    Button(wd[4], text="Go Back", width=macadj(16, 13), command=lambda: informalc_grvlist(wd[0])) \
        .grid(row=0, column=0)
    Label(wd[4], text="Report: ", width=macadj(16, 11)).grid(row=0, column=1)
    Button(wd[4], text="By Settlements", width=macadj(16, 13), command=lambda: informalc_rptgrvsum(wd[0], result)) \
        .grid(row=0, column=2)
    Button(wd[4], text="By Carriers", width=macadj(16, 13), command=lambda: informalc_bycarriers(wd[0], result)) \
        .grid(row=0, column=3)
    Button(wd[4], text="By Carrier", width=macadj(16, 13), command=lambda: informalc_bycarrier(wd[0], result)) \
        .grid(row=0, column=4)
    Label(wd[4], text="Summary: ", width=macadj(16, 11)).grid(row=1, column=1)
    Button(wd[4], text="By Settlements", width=macadj(16, 13),
           command=lambda: informalc_grvlist_setsum(result)).grid(row=1, column=2)
    rear_window(wd)


def informalc_date_checker(frame, date, type):
    d = date.get().split("/")
    if len(d) != 3:
        messagebox.showerror("Invalid Data Entry",
                             "The date for the {} is not properly formatted.".format(type),
                             parent=frame)
        return "fail"
    for num in d:
        if not num.isnumeric():
            messagebox.showerror("Invalid Data Entry",
                                 "The month, day and year for the {} "
                                 "must be numeric.".format(type),
                                 parent=frame)
            return "fail"
    if len(d[0]) > 2:
        messagebox.showerror("Invalid Data Entry",
                             "The month for the {} must be no more than two digits"
                             " long.".format(type),
                             parent=frame)
        return "fail"
    if len(d[1]) > 2:
        messagebox.showerror("Invalid Data Entry",
                             "The day for the {} must be no more than two digits"
                             " long.".format(type),
                             parent=frame)
        return "fail"
    if len(d[2]) != 4:
        messagebox.showerror("Invalid Data Entry",
                             "The year for the {} must be four digits long."
                             .format(type),
                             parent=frame)
        return "fail"
    try:
        date = datetime(int(d[2]), int(d[0]), int(d[1]))
        valid_date = True
    except ValueError:
        valid_date = False
    if not valid_date:
        messagebox.showerror("Invalid Data Entry",
                             "The date entered for {} is not a valid date."
                             .format(type),
                             parent=frame)
        return "fail"


def informalc_grvlist_apply(frame,
                            incident_date, incident_start, incident_end,
                            signing_date, signing_start, signing_end,
                            station, set_lvl, level,
                            gats, have_gats,
                            docs, have_docs):
    conditions = []
    if incident_date.get() == "yes":
        check = informalc_date_checker(frame, incident_start, "starting incident date")
        if check == "fail":
            return
        check = informalc_date_checker(frame, incident_end, "ending incident date")
        if check == "fail":
            return
        d = incident_start.get().split("/")
        start = datetime(int(d[2]), int(d[0]), int(d[1]))
        d = incident_end.get().split("/")
        end = datetime(int(d[2]), int(d[0]), int(d[1]))
        if start > end:
            messagebox.showerror("Invalid Data Entry",
                                 "Your starting incident date must be earlier than your "
                                 "ending incident date.",
                                 parent=frame)
            return
        to_add = "indate_start > '{}' and indate_end < '{}'".format(start, end)
        conditions.append(to_add)
    if signing_date.get() == "yes":
        check = informalc_date_checker(frame, signing_start, "starting signing date")
        if check == "fail":
            return
        check = informalc_date_checker(frame, signing_end, "ending signing date")
        if check == "fail":
            return
        d = signing_start.get().split("/")
        start = datetime(int(d[2]), int(d[0]), int(d[1]))
        d = signing_end.get().split("/")
        end = datetime(int(d[2]), int(d[0]), int(d[1]))
        if start > end:
            messagebox.showerror("Invalid Data Entry",
                                 "Your starting signing date must be earlier than your "
                                 "ending signing date.",
                                 parent=frame)
            return
        to_add = "date_signed BETWEEN '{}' AND '{}'".format(start, end)
        conditions.append(to_add)
    if station.get() == "Select a Station":
        messagebox.showerror("Invalid Station",
                             "You must select a station.",
                             parent=frame)
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
    Label(wd[3], text="Informal C: Settlement Search Criteria", font=macadj("bold", "Helvetica 18")) \
        .grid(row=0, columnspan=6, sticky="w")
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
    Label(wd[3], text=" Station ", background=macadj("gray95", "grey"), fg=macadj("black", "white"),
          anchor="w", width=macadj(14, 12)).grid(row=2, column=0, columnspan=3, sticky="w")
    station_options = projvar.list_of_stations
    if "out of station" in station_options:
        station_options.remove("out of station")
    station_om = OptionMenu(wd[3], station, *station_options)
    station_om.config(width=macadj(38, 31))
    station_om.grid(row=2, column=3, columnspan=3, sticky="e")
    station.set("Select a Station")
    Label(wd[3], text="Search For", fg="grey").grid(row=3, column=0, columnspan=2, sticky="w")
    Label(wd[3], text="Category", fg="grey").grid(row=3, column=3)
    Label(wd[3], text="Start", fg="grey").grid(row=3, column=4)
    Label(wd[3], text="End", fg="grey").grid(row=3, column=5)
    # select for starting date
    Radiobutton(wd[3], text="yes", variable=incident_date, value='yes', width=macadj(2, 4)) \
        .grid(row=4, column=0, sticky="w")
    Radiobutton(wd[3], text="no", variable=incident_date, value='no', width=macadj(2, 4)) \
        .grid(row=4, column=1, sticky="w")
    Label(wd[3], text="", width=macadj(2, 4)).grid(row=4, column=2)
    Label(wd[3], text=" Incident Dates", background=macadj("gray95", "grey"), fg=macadj("black", "white"),
          anchor="w", width=14).grid(row=4, column=3, sticky="w")
    Entry(wd[3], textvariable=incident_start, width=macadj(12, 8), justify='right').grid(row=4, column=4)
    Entry(wd[3], textvariable=incident_end, width=macadj(12, 8), justify='right').grid(row=4, column=5)
    incident_date.set('no')
    # select for signing date
    Radiobutton(wd[3], text="yes", variable=signing_date, value='yes', width=macadj(2, 4)) \
        .grid(row=5, column=0, sticky="w")
    Radiobutton(wd[3], text="no", variable=signing_date, value='no', width=macadj(2, 4)) \
        .grid(row=5, column=1, sticky="w")
    Label(wd[3], text=" Signing Dates", background=macadj("gray95", "grey"), fg=macadj("black", "white"),
          anchor="w", width=14).grid(row=5, column=3, sticky="w")
    Entry(wd[3], textvariable=signing_start, width=macadj(12, 8), justify='right').grid(row=5, column=4)
    Entry(wd[3], textvariable=signing_end, width=macadj(12, 8), justify='right').grid(row=5, column=5)
    signing_date.set('no')
    # select for settlement level
    Radiobutton(wd[3], text="yes", variable=set_lvl, value='yes', width=macadj(2, 4)) \
        .grid(row=6, column=0, sticky="w")
    Radiobutton(wd[3], text="no", variable=set_lvl, value='no', width=macadj(2, 4)) \
        .grid(row=6, column=1, sticky="w")
    set_lvl.set("no")
    Label(wd[3], text=" Settlement Level ", background=macadj("gray95", "grey"), fg=macadj("black", "white"),
          anchor="w", width=14, height=1).grid(row=6, column=3, sticky="w")
    lvl_options = ("informal a", "formal a", "step b", "pre-arb", "arbitration")
    lvl_om = OptionMenu(wd[3], level, *lvl_options)
    lvl_om.config(width=macadj(20, 16))
    lvl_om.grid(row=6, column=4, columnspan=3, sticky="e")
    level.set("informal a")
    # select for gats number
    Radiobutton(wd[3], text="yes", variable=gats, value='yes', width=macadj(2, 4)) \
        .grid(row=7, column=0, sticky="w")
    Radiobutton(wd[3], text="no", variable=gats, value='no', width=macadj(2, 4)) \
        .grid(row=7, column=1, sticky="w")
    Label(wd[3], text=" GATS Number", background=macadj("gray95", "grey"), fg=macadj("black", "white"),
          anchor="w", width=14, height=1).grid(row=7, column=3, sticky="w")
    gats_options = ("no", "yes")
    gats_om = OptionMenu(wd[3], have_gats, *gats_options)
    gats_om.config(width=macadj(10, 8))
    gats_om.grid(row=7, column=4, columnspan=3, sticky="e")
    have_gats.set('no')
    gats.set('no')

    # select for documentation
    Radiobutton(wd[3], text="yes", variable=docs, value='yes', width=macadj(2, 4)) \
        .grid(row=9, column=0, sticky="w")
    Radiobutton(wd[3], text="no", variable=docs, value='no', width=macadj(2, 4)) \
        .grid(row=9, column=1, sticky="w")
    Label(wd[3], text=" Documentation", background=macadj("gray95", "grey"), fg=macadj("black", "white"),
          anchor="w", width=14, height=1).grid(row=9, column=3, sticky="w")
    doc_options = ("moot", "no", "partial", "yes", "incomplete", "verified")
    docs_om = OptionMenu(wd[3], have_docs, *doc_options)
    docs_om.config(width=macadj(10, 8))
    docs_om.grid(row=9, column=4, columnspan=3, sticky="e")
    have_docs.set('no')
    docs.set("no")
    Label(wd[3], text="").grid(row=13)
    # buttons
    Button(wd[4], text="Search", width=20,
           command=lambda: informalc_grvlist_apply(wd[0], incident_date, incident_start, incident_end,
                                                   signing_date, signing_start, signing_end, station, set_lvl, level,
                                                   gats, have_gats, docs, have_docs)).grid(row=0, column=1)
    Button(wd[4], text="Go Back", width=20, anchor="w", command=lambda: informalc(wd[0])).grid(row=0, column=0)
    rear_window(wd)


def informalc_new(frame, msg):
    wd = front_window(frame)  # F,S,C,FF,buttons
    Label(wd[3], text="New Settlement", font=macadj("bold", "Helvetica 18")).grid(row=0, column=0, sticky="w")
    Label(wd[3], text="").grid(row=1, column=0, sticky="w")
    Label(wd[3], text="Grievance Number: ", background=macadj("gray95", "grey"),
          fg=macadj("black", "white"), width=macadj(22, 20), anchor="w", height=macadj(1, 1)) \
        .grid(row=2, column=0, sticky="w")
    grv_no = StringVar(wd[0])
    Entry(wd[3], textvariable=grv_no, justify='right', width=macadj(20, 15)) \
        .grid(row=2, column=1, sticky="w")
    Label(wd[3], text="Incident Date").grid(row=3, column=0, sticky="w")
    Label(wd[3], text="  Start (mm/dd/yyyy): ", background=macadj("gray95", "grey"),
          fg=macadj("black", "white"), width=macadj(22, 20), anchor="w") \
        .grid(row=4, column=0, sticky="w")
    incident_start = StringVar(wd[0])
    Entry(wd[3], textvariable=incident_start, justify='right', width=macadj(20, 15)) \
        .grid(row=4, column=1, sticky="w")
    Label(wd[3], text="  End (mm/dd/yyyy): ", background=macadj("gray95", "grey"),
          fg=macadj("black", "white"), width=macadj(22, 20), anchor="w", height=macadj(1, 1)) \
        .grid(row=5, column=0, sticky="w")
    incident_end = StringVar(wd[0])
    Entry(wd[3], textvariable=incident_end, justify='right', width=macadj(20, 15)) \
        .grid(row=5, column=1, sticky="w")
    Label(wd[3], text="Date Signed (mm/dd/yyyy): ", background=macadj("gray95", "grey"),
          fg=macadj("black", "white"), width=macadj(22, 20), anchor="w", height=macadj(1, 1)) \
        .grid(row=6, column=0, sticky="w")
    date_signed = StringVar(wd[0])
    Entry(wd[3], textvariable=date_signed, justify='right', width=macadj(20, 15)) \
        .grid(row=6, column=1, sticky="w")
    # select level
    Label(wd[3], text="Settlement Level: ", background=macadj("gray95", "grey"),
          fg=macadj("black", "white"), width=macadj(22, 20), anchor="w", height=macadj(1, 1)) \
        .grid(row=7, column=0, sticky="w")  # select settlement level
    lvl = StringVar(wd[0])
    lvl_options = ("informal a", "formal a", "step b", "pre arb", "arbitration")
    lvl_om = OptionMenu(wd[3], lvl, *lvl_options)
    lvl_om.config(width=macadj(13, 13))
    lvl_om.grid(row=7, column=1)
    lvl.set("informal a")
    Label(wd[3], text="Station: ", background=macadj("gray95", "grey"),  # select a station
          fg=macadj("black", "white"), width=macadj(22, 20), anchor="w", height=macadj(1, 1)). \
        grid(row=8, column=0, sticky="w")
    Label(wd[3], text="", height=macadj(1, 2)).grid(row=8, column=1)
    station = StringVar(wd[0])
    station.set("Select a Station")
    station_options = projvar.list_of_stations
    if "out of station" in station_options:
        station_options.remove("out of station")
    station_om = OptionMenu(wd[3], station, *station_options)
    station_om.config(width=macadj(40, 34))
    station_om.grid(row=9, column=0, columnspan=2, sticky="e")
    Label(wd[3], text="GATS Number: ", background=macadj("gray95", "grey"),
          fg=macadj("black", "white"), width=macadj(22, 20), anchor="w", height=macadj(1, 1)) \
        .grid(row=10, column=0, sticky="w")  # enter gats number
    gats_number = StringVar(wd[0])
    Entry(wd[3], textvariable=gats_number, justify='right', width=macadj(20, 15)) \
        .grid(row=10, column=1, sticky="w")
    Label(wd[3], text="Documentation?: ", background=macadj("gray95", "grey"),
          fg=macadj("black", "white"), width=macadj(22, 20), anchor="w", height=macadj(1, 1)) \
        .grid(row=11, column=0, sticky="w")  # select documentation
    docs = StringVar(wd[0])
    doc_options = ("moot", "no", "partial", "yes", "incomplete", "verified")
    docs_om = OptionMenu(wd[3], docs, *doc_options)
    docs_om.config(width=macadj(13, 13))
    docs_om.grid(row=11, column=1)
    docs.set("no")
    Label(wd[3], text="Description: ", background=macadj("gray95", "grey"),
          fg=macadj("black", "white"), width=macadj(22, 20), anchor="w", height=macadj(1, 1)) \
        .grid(row=15, column=0, sticky="w")
    Label(wd[3], text="", height=macadj(1, 2)).grid(row=15, column=1)
    description = StringVar(wd[0])
    Entry(wd[3], textvariable=description, width=macadj(48, 36), justify='right') \
        .grid(row=16, column=0, sticky="w", columnspan=2)
    Label(wd[3], text="", height=macadj(1, 1)).grid(row=17, column=0)
    Label(wd[3], text=msg, fg="red", height=macadj(1, 1)).grid(row=18, column=0, columnspan=2, sticky="w")
    Button(wd[4], text="Go Back", width=macadj(19, 18), anchor="w",
           command=lambda: informalc(wd[0])).grid(row=0, column=0)
    Button(wd[4], text="Enter", width=macadj(19, 18),
           command=lambda: informalc_new_apply
           (wd[0], grv_no, incident_start, incident_end, date_signed, station, gats_number,
            docs, description, lvl)).grid(row=0, column=1)
    rear_window(wd)


def informalc_poe_apply_search(frame, year, station, backdate):
    if year.get().strip() == "":
        messagebox.showerror("Data Entry Error",
                             "You must enter a year.",
                             parent=frame)
        return
    if "." in year.get():
        messagebox.showerror("Data Entry Error",
                             "The year can not contain decimal points.",
                             parent=frame)
        return
    if not year.get().isnumeric():
        messagebox.showerror("Data Entry Error",
                             "The year must numeric without any letters or special characters.",
                             parent=frame)
        return
    if float(year.get()) > 9999 or float(year.get()) < 2:
        messagebox.showerror("Data Entry Error",
                             "The year must be between the year 2 and 9999.\nI think I'm being "
                             "reasonable.",
                             parent=frame)
        return
    if station.get() == "undefined":
        messagebox.showerror("Data Entry Error",
                             "You must select a station.",
                             parent=frame)
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
        messagebox.showerror("Data Entry Error",
                             "You must select a name.",
                             parent=frame)
        return
    for i in range(len(poe_add_pay_periods)):
        pp = poe_add_pay_periods[i].get().strip()
        hr = poe_add_hours[i].get().strip()
        rt = poe_add_rate[i].get().strip()
        amt = poe_add_amount[i].get().strip()
        if pp and not isint(pp):
            messagebox.showerror("Data Input Error",
                                 "Input error for {} in row {}. The pay period must be a number"
                                 .format(name, str(i + 1)),
                                 parent=frame)
            return
        if pp and int(pp) > 27:
            messagebox.showerror("Data Input Error",
                                 "Input error for {} in row {}. The pay period can not be greater "
                                 "than 27".format(name, str(i + 1)),
                                 parent=frame)
            return
        if hr and amt:
            messagebox.showerror("Data Input Error",
                                 "Input error for {} in row {}. You can not enter both hours and "
                                 "amount. You can only enter one or another, but not both. "
                                 "Awards can be in the form of "
                                 "hours at a given rate OR an amount.".format(name, str(i + 1)),
                                 parent=frame)
            return
        if rt and amt:
            messagebox.showerror("Data Input Error",
                                 "Input error for {} in row {}. You can not enter both rate and "
                                 "amount. You can only enter one or another, but not both. "
                                 "Awards can be in the form of "
                                 "hours at a given rate OR an amount.".format(name, str(i + 1)),
                                 parent=frame)
            return
        if hr and not rt:
            messagebox.showerror("Data Input Error",
                                 "Input error for {} in row {}. Hours must be a accompanied by a "
                                 "rate.".format(name, str(i + 1)),
                                 parent=frame)
            return
        if rt and not hr:
            messagebox.showerror("Data Input Error",
                                 "Input error for {} in row {}. Rate must be a accompanied by a "
                                 "hours.".format(name, str(i + 1)),
                                 parent=frame)
            return
        if hr and not isfloat(hr):
            messagebox.showerror("Data Input Error", "Input error for {} in row {}. Hours must be a number."
                                 .format(name, str(i + 1)),
                                 parent=frame)
            return
        if hr and '.' in hr:
            s_hrs = hr.split(".")
            if len(s_hrs[1]) > 2:
                messagebox.showerror("Data Input Error",
                                     "Input error for {} in row {}. Hours must have no "
                                     "more than 2 decimal places.".format(name, str(i + 1)),
                                     parent=frame)
                return
        if rt and amt:
            messagebox.showerror("Data Input Error",
                                 "Input error for {} in row {}. You can not enter both rate and "
                                 "amount. You can only enter one or the other, but not both. "
                                 "Awards can be in the form of "
                                 "hours at a given rate OR an amount.".format(name, str(i + 1)),
                                 parent=frame)
            return
        if rt and not isfloat(rt):
            messagebox.showerror("Data Input Error",
                                 "Input error for {} in row {}. Rate must be a number."
                                 .format(name, str(i + 1)),
                                 parent=frame)
            return
        if rt and '.' in rt:
            s_rate = rt.split(".")
            if len(s_rate[1]) > 2:
                messagebox.showerror("Data Input Error",
                                     "Input error for {} in row {}. Rates must have no "
                                     "more than 2 decimal places.".format(name, str(i + 1)),
                                     parent=frame)
                return
        if rt and float(rt) > 10:
            messagebox.showerror("Data Input Error",
                                 "Input error for {} in row {}. Values greater than 10 are not "
                                 "accepted. \n"
                                 "Note the following rates would be expressed as: \n "
                                 "additional %50         .50 or just .5 \n"
                                 "straight time rate     1.00 or just 1 \n"
                                 "overtime rate          1.50 or 1.5 \n"
                                 "penalty rate           2.00 or just 2".format(name, str(i + 1)),
                                 parent=frame)
            return
        if amt and not isfloat(amt):
            messagebox.showerror("Data Input Error",
                                 "Input error for {} in row {}. Amounts can only be expressed as "
                                 "numbers. No special characters, such as $ are allowed."
                                 .format(name, str(i + 1)),
                                 parent=frame)
            return
        if amt and '.' in amt:
            s_amt = amt.split(".")
            if len(s_amt[1]) > 2:
                messagebox.showerror("Data Input Error",
                                     "Input error for {} in row {}. Amounts must have no "
                                     "more than 2 decimal places.".format(name, str(i + 1)),
                                     parent=frame)
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
                dt = find_pp(int(year), pp)  # returns the starting date of the pp when given year and pay period
                dt += timedelta(days=20)
                paydays.append(dt)
                sql = "INSERT INTO informalc_payouts (year,pp,payday,carrier_name,hours,rate,amount) " \
                      "VALUES('%s','%s','%s','%s','%s','%s','%s')" \
                      % (year, poe_add_pay_periods[i].get().strip(), paydays[i], name,
                         poe_add_hours[i].get().strip(), poe_add_rate[i].get().strip(),
                         poe_add_amount[i].get().strip())
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
    Label(wd[3], text="Informal C: Payout Entry", font=macadj("bold", "Helvetica 18")) \
        .grid(row=0, column=0, sticky="w", columnspan=5)
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
    except TclError:
        pass
    informalc_poe_search(frame)


def informalc_poe_listbox(dt_year, station, dt_start, year):
    global informalc_poe_lbox  # initialize the global
    poe_root = Tk()
    informalc_poe_lbox = poe_root  # set the global
    poe_root.title("KLUSTERBOX")
    titlebar_icon(poe_root)  # place icon in titlebar
    x_position = projvar.root.winfo_x() + 450
    y_position = projvar.root.winfo_y() - 25
    poe_root.geometry("%dx%d+%d+%d" % (240, 600, x_position, y_position))
    n_f = Frame(poe_root)
    n_f.pack()
    n_buttons = Canvas(n_f)  # button bar
    n_buttons.pack(fill=BOTH, side=BOTTOM)
    Label(n_f, text="Carrier List", font=macadj("bold", "Helvetica 18")).pack(anchor="w")
    Label(n_f, text="{} Station:".format(station.get())).pack(anchor="w")
    Label(n_f, text="{} though {}".format(dt_year.strftime("%Y"), dt_start.strftime("%Y"))).pack(anchor="w")
    Label(n_f, text="").pack()
    scrollbar = Scrollbar(n_f, orient=VERTICAL)
    listbox = Listbox(n_f, selectmode="single", yscrollcommand=scrollbar.set)
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
    the_station = StringVar(wd[0])
    station_options = projvar.list_of_stations
    if "out of station" in station_options:
        station_options.remove("out of station")
    the_station.set("undefined")
    backdate = StringVar(wd[0])
    backdate.set("1")
    Label(wd[3], text="Informal C: Payout Entry Criteria", font=macadj("bold", "Helvetica 18")) \
        .grid(row=0, column=0, sticky="w", columnspan=4)
    Label(wd[3], text="").grid(row=1)
    Label(wd[3], text="Enter the year and the station to be updated.") \
        .grid(row=2, column=0, columnspan=4, sticky="w")
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


def informalc_por_all(frame, afterdate, beforedate, station, backdate):
    check = informalc_date_checker(frame, afterdate, "After Date")
    if check == "fail":
        return
    check = informalc_date_checker(frame, beforedate, "Before Date")
    if check == "fail":
        return
    start = informalc_date_converter(afterdate)
    end = informalc_date_converter(beforedate)
    if start > end:
        messagebox.showerror("Data Entry Error",
                             "The After Date can not be earlier than the Before Date",
                             parent=frame)
        return
    if station.get() == "undefined":
        messagebox.showerror("Data Entry Error",
                             "You must select a station. ",
                             parent=frame)
        return
    weeks = int(backdate.get()) * 52
    clist_start = start - timedelta(weeks=weeks)
    carrier_list = informalc_gen_clist(clist_start, end, station.get())

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = "infc_grv_list" + "_" + stamp + ".txt"
    report = open(dir_path('infc_grv') + filename, "w")
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
                if result[4]:
                    hour = float(result[4])
                if result[5]:
                    rate = float(result[5])
                if result[6]:
                    amt = float(result[6])
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
                report.write('    {:<5}{:>17}{:>9}{:>7}{:>10}{:>12}\n'.format(pp, payday, hours, rate, adj, amt))
            report.write("    --------------------------------------------------------------\n")
            report.write("    {:<40}{:>10}\n".format("Payouts adjusted to straight time", "{0:.2f}"
                                                     .format(float(payxadj))))
            report.write("    {:<38}{:>24}\n".format("Payouts as flat dollar amount", "{0:.2f}"
                                                     .format(float(payxamt))))
            report.write("\n\n\n")

    report.close()
    if sys.platform == "win32":
        os.startfile(dir_path('infc_grv') + filename)
    if sys.platform == "linux":
        subprocess.call(["xdg-open", 'kb_sub/infc_grv/' + filename])
    if sys.platform == "darwin":
        subprocess.call(["open", dir_path('infc_grv') + filename])


def informalc_por(frame):
    wd = front_window(frame)
    afterdate = StringVar(wd[0])
    beforedate = StringVar(wd[0])
    station = StringVar(wd[0])
    station_options = projvar.list_of_stations
    if "out of station" in station_options:
        station_options.remove("out of station")
    station.set("undefined")
    backdate = StringVar(wd[0])
    backdate.set("1")
    Label(wd[3], text="Informal C: Payout Report Search Criteria", font=macadj("bold", "Helvetica 18")) \
        .grid(row=0, column=0, columnspan=4, sticky="w")
    Label(wd[3], text="").grid(row=1)
    Label(wd[3], text="Enter range of dates and select station").grid(row=2, column=0, columnspan=4, sticky="w")
    Label(wd[3], text="\tProvide dates in mm/dd/yyyy format.", fg="grey") \
        .grid(row=3, column=0, columnspan=4, sticky="w")
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
           command=lambda: informalc_por_all(wd[3], afterdate, beforedate, station, backdate)) \
        .grid(row=0, column=2)
    Button(wd[4], text="By Carrier", width=16).grid(row=0, column=3)
    rear_window(wd)


def informalc(frame):
    if os.path.isdir(dir_path_check('infc_grv')):  # clear contents of temp folder
        shutil.rmtree(dir_path_check('infc_grv'))
    sql = 'CREATE table IF NOT EXISTS informalc_grv (grv_no varchar, indate_start varchar, indate_end varchar,' \
          'date_signed varchar, station varchar, gats_number varchar, ' \
          'docs varchar, description varchar, level varchar)'
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
    sql = 'CREATE table IF NOT EXISTS informalc_payouts(year varchar, pp varchar, ' \
          'payday varchar, carrier_name varchar,' \
          'hours varchar,rate varchar,amount varchar)'
    commit(sql)
    # put out of station back into the list of stations in case it has been removed.
    if "out of station" not in projvar.list_of_stations:
        projvar.list_of_stations.append("out of station")
    wd = front_window(frame)  # F,S,C,FF,buttons
    Label(wd[3], text="Informal C", font=macadj("bold", "Helvetica 18")).grid(row=0, sticky="w")
    Label(wd[3], text="The C is for Compliance").grid(row=1, sticky="w")
    Label(wd[3], text="").grid(row=2)
    Button(wd[3], text="New Settlement", width=30, command=lambda: informalc_new(wd[0], " ")).grid(row=3, pady=5)
    Button(wd[3], text="Settlement List", width=30, command=lambda: informalc_grvlist(wd[0])).grid(row=4, pady=5)
    Button(wd[3], text="Payout Entry", width=30, command=lambda: informalc_poe_search(wd[0])).grid(row=5, pady=5)
    Button(wd[3], text="Payout Report", width=30, command=lambda: informalc_por(wd[0])).grid(row=6, pady=5)
    Label(wd[3], text="", width=70).grid(row=7)
    button_back = Button(wd[4])
    button_back.config(text="Go Back", width=20, command=lambda: MainFrame().start(frame=wd[0]))
    if sys.platform == "win32":
        button_back.config(anchor="w")
    button_back.grid(row=0, column=0)
    rear_window(wd)


def wkly_avail(frame):  # creates a spreadsheet which shows weekly otdl availability
    path = dir_filedialog()
    file_path = filedialog.askopenfilename(initialdir=path, filetypes=[("Excel files", "*.csv *.xls")])
    if file_path[-4:].lower() == ".csv" or file_path[-4:].lower() == ".xls":
        pass
    else:
        messagebox.showerror("Report Generator",
                             "The file you have selected is not a .csv or .xls file.\n"
                             "You must select a file with a .csv or .xls extension.",
                             parent=frame)
        return
    with open(file_path, newline="") as file:
        a_file = csv.reader(file)
        cc = 0
        for line in a_file:
            if cc == 0 and line[0][:8] != "TAC500R3":
                messagebox.showwarning("File Selection Error",
                                       "The selected file does not appear to be an "
                                       "Employee Everything report.",
                                       parent=frame)
                return
            if cc == 3:
                tacs_pp = line[0]  # find the pay period
                tacs_station = line[2]  # find the station
                break
            cc += 1
        cc = 0
        range_days = ["Saturday", "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
        for line in a_file:  # find the range
            if line[18] in range_days:
                range_days.remove(line[18])
            if cc == 150: break  # survey 150 lines before breaking to anaylize results.
            cc += 1
        if len(range_days) > 5:
            messagebox.showwarning("File Selection Error",
                                   "Employee Everything Reports that cover only one day /n"
                                   "are not supported in version {} of Klusterbox.".format(version),
                                   parent=frame)
            return
        else:
            t_range = True
    year = int(tacs_pp[:-3])  # set the globals
    pp = tacs_pp[-3:]
    t_date = find_pp(year, pp)  # returns the starting date of the pp when given year and pay period
    s_year = t_date.strftime("%Y")
    s_mo = t_date.strftime("%m")
    s_day = t_date.strftime("%d")
    sql = "SELECT kb_station FROM station_index WHERE tacs_station = '%s'" % tacs_station
    station = inquire(sql)  # check to see if station has match in station index
    if not station:
        messagebox.showwarning("Error",
                               "This station has not been matched with Auto Data Entry.",
                               parent=frame)
        return
    set_globals(s_year, s_mo, s_day, t_range, station[0][0], "None")  # set the investigation range
    # get the otdl list from the carriers table
    sql = "SELECT carrier_name FROM carriers WHERE effective_date <= '%s' and station = '%s' and list_status = '%s'" \
          "ORDER BY carrier_name, effective_date desc" % (projvar.invran_date_week[6], projvar.invran_station, 'otdl')
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
        sql = "SELECT emp_id FROM name_index WHERE kb_name='%s'" % name
        results = inquire(sql)
        if results:  # record emp id to otdl carrier info
            ot_wkly.append(results[0][0])
        else:  # mark otdl carriers who don't have emp id available
            ot_wkly.append("no index")
        sql = "SELECT effective_date,list_status,station FROM carriers " \
              "WHERE carrier_name='%s' and effective_date<='%s'" \
              "ORDER BY effective_date desc" % (name, projvar.invran_date_week[6])
        results = inquire(sql)
        ot_wkly.append(name)
        for date in projvar.invran_date_week:  # loop for each day of the week
            for rec in results:  # loop for each record starting from the latest
                if rec[2] == projvar.invran_station:  # if there is a station match
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
        station_anchor = "no"  # reset
    not_indexed = []
    for name in wkly_list:  # check to see if there are any otdl carriers who do not have a rec in name index
        if name[0] == "no index":
            not_indexed.append(name[1])  # add any names who do not into an array
    if len(not_indexed) != 0:  # message box info that some otdl do not have a record in the name index
        messagebox.showwarning("Missing Data",
                               "There are {} name/s which have not been matched with their employee id."
                               " Please exit and run the Auto Data Entry Feature to ensure that all carriers have "
                               " employee ids entered into Klusterbox.".format(len(not_indexed)),
                               parent=frame)
    if len(otdl_list) == 0:
        messagebox.showwarning("Empty OTDL",
                               "Klusterbox has no records of any otdl carriers for {} station "
                               "for the week of {}. This could mean that: \n1. The carrier list is empty. Run the "
                               "Automatic Data Entry Feature, selecting the Employee Everything Report you used here "
                               " to remedy this. You do not have to enter the rings data at the final step "
                               " \n2. The Name Index which matches the carrier name to the employee id "
                               "empty. As in #1, run the Automatic Data Entry Feature to fix this.\n3. "
                               "The carrier list has no otdl carriers "
                               "designated. Use the Multi Input Feature to designate otdl carriers. \n"
                               "This Weekly Availability Report can not be generated without a list of otdl carriers. "
                               "Build the carrier list/otdl before re-running Weekly Availability."
                               .format(projvar.invran_station, projvar.invran_date_week[0].strftime("%b %d, %Y")),
                               parent=frame)
        MainFrame().start(frame=frame)
    else:  # if there is an otdl then build array holding hours for each day
        days = ("Saturday", "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday")
        extra_hour_codes = ("49", "52", "55", "56", "57", "58", "59", "60")
        running_total = 0
        with open(file_path, newline="") as file:
            a_file = csv.reader(file)
            cc = 0
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
                if cc != 0 and line[4].zfill(8) in otdl_list:  # if the emp_id matches ones we are looking for
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
                    # find first line of specific carrier
                    if line[18] == "Base" and line[19] in ("844", "134", "434"):
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
                cc += 1
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
        cell = wkly_total.cell(row=1, column=1)
        cell.value = "Weekly Availability Summary"
        cell.style = ws_header
        wkly_total.merge_cells('A1:E1')
        wkly_total['A3'] = "Date:  "  # create date/ pay period/ station header
        wkly_total['A3'].style = date_dov_title
        range_of_dates = format(projvar.invran_date_week[0], "%A  %m/%d/%y") + " - " + \
                         format(projvar.invran_date_week[6], "%A  %m/%d/%y")
        wkly_total['B3'] = range_of_dates
        wkly_total['B3'].style = date_dov
        wkly_total.merge_cells('B3:H3')
        date = datetime(int(projvar.invran_year), int(projvar.invran_month), int(projvar.invran_day))
        projvar.pay_period = pp_by_date(date)
        wkly_total['E4'] = "Pay Period:  "
        wkly_total['E4'].style = date_dov_title
        wkly_total.merge_cells('E4:F4')
        wkly_total['G4'] = projvar.pay_period
        wkly_total['G4'].style = date_dov
        wkly_total.merge_cells('G4:H4')
        wkly_total['A4'] = "Station:  "
        wkly_total['A4'].style = date_dov_title
        wkly_total['B4'] = projvar.invran_station
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
        xl_filename = "kb_wa" + str(format(projvar.invran_date_week[0], "_%y_%m_%d")) + ".xlsx"
        ok = messagebox.askokcancel("Spreadsheet generator",
                                    "Do you want to generate a spreadsheet?",
                                    parent=frame)
        if ok:
            try:
                wb.save(dir_path('weekly_availability') + xl_filename)
                messagebox.showinfo("Spreadsheet generator",
                                    "Your spreadsheet was successfully generated. \n"
                                    "File is named: {}".format(xl_filename),
                                    parent=frame)
                if sys.platform == "win32":
                    os.startfile(dir_path('weekly_availability') + xl_filename)
                if sys.platform == "linux":
                    subprocess.call(["xdg-open", 'kb_sub/weekly_availability/' + xl_filename])
                if sys.platform == "darwin":
                    subprocess.call(["open", dir_path('weekly_availability') + xl_filename])
            except PermissionError:
                messagebox.showerror("Spreadsheet generator",
                                     "The spreadsheet was not generated. \n"
                                     "Suggestion: "
                                     "Make sure that identically named spreadsheets are closed "
                                     "(the file can't be overwritten while open).",
                                     parent=frame)
        MainFrame().start(frame=frame)


def station_rec_del(frame, tacs, kb):
    sql = "DELETE FROM station_index WHERE tacs_station = '%s' and kb_station='%s'" % (tacs, kb)
    commit(sql)
    frame.destroy()
    station_index_mgmt("none")


def station_index_rename_apply(frame, tacs, newname):
    sql = "UPDATE station_index SET kb_station='%s' WHERE tacs_station='%s'" % (newname.get(), tacs)
    commit(sql)
    station_index_mgmt(frame)


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
        Button(frame, text="rename", command=lambda: station_index_rename_apply(self, tacs, newname)) \
            .grid(row=1, column=2)
    else:
        Label(frame, text="No Unassigned Stations Available").grid(row=1, column=0, columnspan=2, sticky="e")


def stationindexer_del_all():
    sql = "DELETE FROM station_index"
    commit(sql)
    station_index_mgmt("none")


def station_index_mgmt(frame):
    wd = front_window(frame)  # get window objects 0=F,1=S,2=C,3=FF,4=buttons
    g = 0
    Label(wd[3], text="Station Index Management", font=macadj("bold", "Helvetica 18")) \
        .grid(row=g, column=0, sticky="w")
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
        Label(header_frame, text="TACS Station Name", width=macadj(30, 25), anchor="w") \
            .grid(row=0, column=0, sticky="w")
        Label(header_frame, text="Klusterbox Station Name", width=macadj(30, 25), anchor="w") \
            .grid(row=0, column=1, sticky="w")
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
            Button(frame[f], text=record[0], width=macadj(30, 25), anchor="w").grid(row=0, column=0)
            Button(frame[f], text=record[1], width=macadj(30, 25), anchor="w").grid(row=0, column=1)
            to_add = Button(frame[f], text="rename", width=6)
            rename_button.append(to_add)
            rename_button[f]['command'] = lambda frame=frame[f], tacs=record[0], kb=record[1], newname=si_newname[f], \
                                                 button=rename_button[f]: station_index_rename\
                (wd[0], frame, tacs, kb, newname, button, all_stations)
            rename_button[f].grid(row=0, column=2)
            delete_button = Button(frame[f], text="delete", width=6,
                                   command=lambda tacs=record[0], kb=record[1]: station_rec_del(wd[0], tacs, kb))
            delete_button.grid(row=0, column=3)
            f += 1
            g += 1
        Label(wd[3], text="", height=1).grid(row=g)
        Button(wd[3], text="Delete All", width="15", command=lambda: (wd[0].destroy(), stationindexer_del_all())) \
            .grid(row=g + 1, column=0, columnspan=3, sticky="e")
    button = Button(wd[4])
    button.config(text="Go Back", width=20, command=lambda: MainFrame().start(frame=wd[0]))
    if sys.platform == "win32":
        button.config(anchor="w")
    button.pack(side=LEFT)
    rear_window(wd)


def apply_nameindexer_list(frame, x):
    sql = "DELETE FROM name_index WHERE emp_id = '%s'" % x
    commit(sql)
    frame.destroy()
    name_index_screen()


def del_all_nameindexer(frame):
    sql = "DELETE FROM name_index"
    commit(sql)
    frame.destroy()
    name_index_screen()


def name_index_screen():
    sql = "SELECT * FROM name_index ORDER BY tacs_name"
    results = inquire(sql)
    wd = front_window("none")  # get window objects
    x = 0
    if len(results) == 0:
        Label(wd[3], text="The Name Index is empty").grid(row=0, column=x)
    else:
        Label(wd[3], text="Name Index Management", font=macadj("bold", "Helvetica 18")) \
            .grid(row=x, column=0, sticky="w", columnspan=2)  # page header
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
            Button(wd[3], text="delete", anchor="w", width=5, relief=RIDGE, command=lambda xx=item[2]:
                apply_nameindexer_list(wd[0], xx)).grid(row=x, column=4)
            x += 1
        Button(wd[3], text="Delete All", width="15", command=lambda: del_all_nameindexer(wd[0])) \
            .grid(row=x, column=0, columnspan=5, sticky="e")
    Button(wd[4], text="Go Back", width=20, command=lambda: MainFrame().start(frame=wd[0])).pack(side=LEFT)
    wd[0].update()
    wd[2].config(scrollregion=wd[2].bbox("all"))
    mainloop()


def gen_ns_dict(file_path, to_addname):  # creates a dictionary of ns days
    days = ("Saturday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday")
    mv_codes = ("BT", "MV", "ET")
    good_jobs = ("134", "844", "434")
    results = []
    carrier = []
    id_bank = []
    aux_list = []
    for id in to_addname:
        id_bank.append(id[0].zfill(8))
        if id[3] in ("auxiliary", "part time flex"):
            aux_list.append(id[0].zfill(8))  # make an array of auxiliary carrier emp ids
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
        return results


class AutoDataEntry:
    def __init__(self):
        self.frame = None
        self.file_path = None
        self.a_file = None  # the openned csv file
        self.tacs_station = None  # the station from the ee report
        self.t_range = None  # false - one day/ true - weekly ee report
        self.t_date = None  # the starting date of the pp
        self.station_index = []  # create a list of klusterbox stations
        self.possible_stations = []  # array of all stations in stations table minus station index
        self.tacs_list = []  # Get the names from tacs report
        self.check_these = []
        self.new_carrier = []  # new carriers who have duplicate names send these to auto indexer 6
        self.name_sorter = []  # stores stringvar objects for name pairing in an array
        self.to_addname = []  # initialize array of names to be added.
        self.tried_names = []
        self.is_mac = macadj(False, True)  # returns True for mac, False if not mac
        self.csv_fix = None
        self.target_file = None

    def run(self, frame):
        self.frame = frame
        self.AutoSetUp(self).run(self.frame)

    def get_file(self):  # read the csv file and assign to self.a_file attribute
        self.target_file = open(self.file_path, newline="")
        self.a_file = csv.reader(self.target_file)

    def go_back(self, frame):
        """
        This first closes the opened csv file is being read
        Then destroys the temporary csv file created by CsvRepair() and referenced by self.file_path.
        Then the MainFrame() is called to return the user to the main screen.
        This is called with self.parent.go_back(frame)
        """
        self.target_file.close()
        self.csv_fix.destroy()
        MainFrame().start(frame=frame)

    class AutoSetUp:
        def __init__(self, parent):
            self.parent = parent
            self.frame = None
            self.tacs_pp = None  # pay period read from csv file
            self.tacs_index = []  # create a list of tacs station names
            self.kb_stations = []  # array of all stations in stations table

        def run(self, frame):
            self.frame = frame
            if not self.get_path():  # get the path to the employee everything report
                return  # return if invalid response
            self.parent.csv_fix = CsvRepair()  # create a CsvRepair object
            # returns a file path for a checked and, if needed, fixed csv file.
            self.parent.file_path = self.parent.csv_fix.run(self.parent.file_path)
            self.auto_precheck()  # delete recs from name index which don't have corresponding recs in carriers table
            self.parent.get_file()  # read the csv file and assign to self.a_file attribute
            if not self.check_file():  # check for invalid file, find station and pay period
                return  # return if invalid response
            if not self.check_range():  # check that file covers full service week
                return  # return if invalid response
            if not self.check_tacs_station():  # check that file has a station
                return  # return if invalid response
            self.get_tacs_date()  # get the date from tacs
            self.get_stations()  # build arrays of stations
            if self.parent.tacs_station not in self.tacs_index:
                self.parent.AutoIndexer1(self.parent).run(self.frame)
            else:
                self.parent.AutoIndexer2(self.parent).run(self.frame)

        def get_path(self):  # get the path to the employee everything report or return False
            path = dir_filedialog()
            self.parent.file_path = filedialog.askopenfilename(initialdir=path,
                                                               filetypes=[("Excel files", "*.csv *.xls")])
            if self.parent.file_path == "":  # if there is no selections - end
                return False
            elif self.parent.file_path[-4:].lower() == ".csv" or self.parent.file_path[-4:].lower() == ".xls":
                return True
            else:  # if an csv nor xls is selected - end
                messagebox.showerror("Report Generator",
                                     "The file you have selected is not a .csv or .xls file. "
                                     "You must select a file with a .csv or .xls extension.",
                                     parent=self.frame)
                return False

        @staticmethod
        def auto_precheck():
            # delete any records from name index which don't have corresponding records in carriers table
            sql = "SELECT kb_name FROM name_index"
            kb_name = inquire(sql)
            sql = "SELECT carrier_name FROM carriers"
            results = inquire(sql)
            carriers = []
            for item in results:
                if item not in carriers: carriers.append(item)
            # create progressbar
            pb = ProgressBarDe(title="Database Maintenance", label="Updating Changes: ")
            pb.max_count(len(kb_name))
            pb.start_up()
            i = 0
            for name in kb_name:
                pb.move_count(i)  # increment progress bar
                if name not in carriers:
                    sql = "DELETE FROM name_index WHERE kb_name = '%s'" % name
                    commit(sql)
                i += 1
            pb.stop()  # stop and destroy the progress bar

        def check_file(self):  # check for invalid file, find station and pay period
            self.parent.get_file()  # read the csv file
            cc = 0
            for line in self.parent.a_file:
                if cc == 0 and line[0][:8] != "TAC500R3":
                    messagebox.showwarning("File Selection Error",
                                           "The selected file does not appear to be an "
                                           "Employee Everything report.",
                                           parent=self.frame)
                    return False
                if cc == 3:
                    self.tacs_pp = line[0]  # find the pay period
                    self.parent.tacs_station = line[2]  # find the station
                    break
                cc += 1
            return True

        def check_range(self):  # check that file covers full service week
            """
            self.parent.a_file is not refreshed. So the loop will will start on line 4 of the csv file. The
            loop will read 150 lines of the code and pick up all the range_days to ensure that a full week is
            covered.
            """
            cc = 0
            range_days = ["Saturday", "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
            for line in self.parent.a_file:  # find the range
                if line[18] in range_days:
                    range_days.remove(line[18])
                if cc == 150: break  # survey 150 lines before breaking to anaylize results.
                cc += 1
            if len(range_days) > 5:
                self.parent.t_range = False  # set the range
                messagebox.showwarning("File Selection Error",
                                       "Employee Everything Reports that cover only one day /n"
                                       "are not supported in this version of Klusterbox.",
                                       parent=self.parent.frame)
                return False
            else:
                self.parent.t_range = True
            return True

        def check_tacs_station(self):  # make sure the csv has a stations
            if len(self.parent.tacs_station) == 0:
                messagebox.showwarning("Auto Data Entry Error",
                                       "The Employee Everything Report is corrupt. Data Entry will stop.  \n"
                                       "The Employee Everything Report does not include "
                                       "information about the station. This could be caused by an error of the pdf "
                                       "converter. If you can obtain an Employee Everything Report from management in "
                                       "csv format, you should have better results.",
                                       parent=self.frame)
                return False
            return True

        def get_tacs_date(self):  # get the tacs date expressed as pay period
            year = int(self.tacs_pp[:-3])
            pp = self.tacs_pp[-3:]
            self.parent.t_date = find_pp(year, pp)  # returns the starting date of the pp when given year and pay period

        def get_stations(self):
            sql = "SELECT tacs_station, kb_station, finance_num FROM station_index"
            results = inquire(sql)
            for line in results:
                self.parent.station_index.append(line[1])  # build station index
                self.tacs_index.append(line[0])  # build tacs_index
            sql = "SELECT station FROM stations"
            results = inquire(sql)
            for record in results:
                self.kb_stations.append(record[0])  # build kb_stations
            for item in self.kb_stations:
                self.parent.possible_stations.append(item)  # build possible stations
            self.parent.station_index.append("out of station")
            self.parent.possible_stations = \
                [x for x in self.parent.possible_stations if x not in self.parent.station_index]

    class AutoIndexer1:  # The station pairing screen
        """ This screen will only appear if station does not have a record in the station index table.
         If there is not a record in the staton index table, this screen will allow to pair the station with
         a station in the stations table which is not already paired in the station index. Or the screen will
         allow the user to enter a completely new station which will be added to the station table and the
         station index table. """
        def __init__(self, parent):
            self.parent = parent
            self.frame = None
            self.win = None  # creates a window object
            self.station_sorter = None
            self.station_new = None

        def run(self, frame):
            self.frame = frame
            self.get_window_object()
            self.station_screen()

        def get_window_object(self):
            self.win = MakeWindow()
            self.win.create(self.frame)

        def station_screen(self):  # pair station from tacs to correct station in klusterbox/ part 1
            Label(self.win.body, text="Station Pairing", font=macadj("bold", "Helvetica 18"), pady=10) \
                .grid(row=0, column=0, columnspan=4, sticky=W)  # page contents
            Label(self.win.body, text="Match the station detected from TACS with a pre-existing station\n "
                           "or use ADD STATION to add the station if there isn't a match.", justify=LEFT) \
                .grid(row=1, column=0, columnspan=4, sticky=W)
            Label(self.win.body, text="Detected Station: ", anchor="w").grid(row=2, column=0, sticky="w")
            Label(self.win.body, text=self.parent.tacs_station, fg="blue").grid(row=3, column=0, columnspan=4)
            Label(self.win.body, text="Select Station: ", anchor="w").grid(row=4, column=0, sticky=W)
            self.station_sorter = StringVar(self.win.body)
            station_options = ["select matching station"] + self.parent.possible_stations + ["ADD STATION"]
            self.station_sorter.set(station_options[0])
            option_menu = OptionMenu(self.win.body, self.station_sorter, *station_options)
            option_menu.config(width=30)
            option_menu.grid(row=5, column=0, columnspan=2, sticky=W)
            Label(self.win.body, text=" ", justify=LEFT).grid(row=6, column=0, sticky=W)
            Label(self.win.body, text="If the station is not present in the drop down menu, select  \n "
                           "ADD STATION from the menu and enter the new station name \n"
                           "below to pair it with the station originating the report", justify=LEFT) \
                .grid(row=7, column=0, columnspan=4, sticky=W)
            Label(self.win.body, text=" ", justify=LEFT).grid(row=8, column=0, sticky=W)
            Label(self.win.body, text="Enter New Station Name: ", anchor="w")\
                .grid(row=9, column=0, columnspan=4, sticky=W)
            # insert entry for station name
            self.station_new = StringVar(self.win.body)
            Entry(self.win.body, width=35, textvariable=self.station_new).grid(row=10, column=0, columnspan=4, sticky=W)
            button_cancel = Button(self.win.buttons)  # cancel button
            button_cancel.config(text="Go Back", width=20, command=lambda: self.parent.go_back(self.win.topframe))
            if sys.platform == "win32":
                button_cancel.config(anchor="w")
            button_cancel.pack(side=LEFT)
            button_apply = Button(self.win.buttons)  # apply button
            button_apply.config(text="Submit", width=20, command=lambda: self.apply())
            if sys.platform == "win32":
                button_apply.config(anchor="w")
            button_apply.pack(side=LEFT)
            self.win.fill(11, 30)  # add white space at bottom of page
            self.win.finish()  # close out the window function

        def apply(self):
            if self.check():  # if the user entered data passes all checks
                self.insert()  # insert the user entered data into the database
                self.parent.AutoIndexer2(self.parent).run(self.win.topframe)
            else:  # if the user entered data fails the checks
                frame = self.win.topframe  # store the frame object so __init__ does not destroy it
                self.__init__(self.parent)  # re initialize the class
                self.run(frame)  # re run the methods of the class

        def check(self):
            self.station_new = self.station_new.get()
            self.station_new = self.station_new.strip()
            """ user didn't select station from the option menu or didn't select ADD STATION and entered a station 
            in the entry widget """
            if self.station_sorter.get() == "select matching station":  # user selected the label and not a station
                messagebox.showerror("Data Entry Error",
                                     "You must select a station or ADD STATION",
                                     parent=self.win.topframe)
                return False
            """ user selected add station but gave no station """
            if self.station_sorter.get() == "ADD STATION" and self.station_new == "":
                messagebox.showerror("Data Entry Error",
                                     "You must provide a name for the new station.",
                                     parent=self.win.topframe)
                return False
            """ user selected station and added station - error """
            if self.station_sorter.get() != "ADD STATION" and self.station_new != "":
                messagebox.showerror("Data Entry Error",
                                     "You can not select a station from the drop down menu AND enter "
                                     "a station in the text field.",
                                     parent=self.win.topframe)
                return False
            return True

        def insert(self):
            if self.station_sorter.get() == "ADD STATION":
                """ if the user is using ADD STATION  to enter a new station not in the option menu """
                # add the new station to the stations table if it is not already there.
                if self.station_new not in projvar.list_of_stations:
                    sql = "INSERT INTO stations (station) VALUES('%s')" % self.station_new
                    commit(sql)
                    projvar.list_of_stations.append(self.station_new)
                # add the station to the station index
                sql = "INSERT INTO station_index (tacs_station, kb_station, finance_num) VALUES('%s','%s','%s')" \
                      % (self.parent.tacs_station, self.station_new, "")
                commit(sql)
                messagebox.showinfo("Database Updated",
                                    "The {} station has been added to the list of stations automatically "
                                    "recognized.".format(self.station_new),
                                    parent=self.win.topframe)
            else:
                """ if the carrier is selecting a station from the drop down menu. add the station to the 
                station index """
                sql = "INSERT INTO station_index (tacs_station, kb_station, finance_num) VALUES('%s','%s','%s')" \
                      % (self.parent.tacs_station, self.station_sorter.get(), "")
                commit(sql)
                messagebox.showinfo("Database Updated",
                                    "The {} station has been paired to the {} station. In the future, this association "
                                    "will be automatically recognized."
                                    .format(self.parent.tacs_station, self.station_sorter.get()),
                                    parent=self.win.topframe)

    class AutoIndexer2:  # Search for name matchs #1
        """ This screen will give the user the opportunity to pair carriers with records in the carrier table
        to new carriers from the employee everything report. Carriers with names that match exactly will be
        matched/paired automatically. Only carriers with no record in the name index will appear in this screen.
        If a new carrier has a name exactly matching an existing carrier, that carrier's employee id number
        will be added to the end of their name because the name is a unique identifier. """
        def __init__(self, parent):
            self.parent = parent
            self.frame = None
            self.name_index = []  # create a list of klusterbox names
            self.id_index = []  # create a list of emp ids
            self.c_list = []  # create a list of unique names from carrier list (a set)
            self.to_remove = []  # intialized array of names to be removed from tacs names
            self.win = None
            self.possible_names = []  # an array of possible matches of kb names and tacs names
            self.possible_match = False  # False if there are no possible matches ever

        def run(self, frame):  # namepairing_create
            self.frame = frame
            self.set_globals()
            self.get_carrier_indexes()
            self.get_carrier_list()
            self.parent.get_file()  # read the csv file and assign to self.a_file attribute
            self.get_tacslist()
            self.remove_tacs_duplicates()
            self.qualify_tacslist()
            self.get_new_carrier()
            self.limit_tacslist()
            self.get_name_index()
            self.namepairing_router()

        def set_globals(self):
            s_year = self.parent.t_date.strftime("%Y")
            s_mo = self.parent.t_date.strftime("%m")
            s_day = self.parent.t_date.strftime("%d")
            sql = "SELECT kb_station FROM station_index WHERE tacs_station = '%s'" % self.parent.tacs_station
            station = inquire(sql)
            set_globals(s_year, s_mo, s_day, self.parent.t_range, station[0][0], "None")

        def get_carrier_indexes(self):
            sql = "SELECT tacs_name, kb_name, emp_id FROM name_index ORDER BY kb_name"
            results = inquire(sql)
            for line in results:
                self.name_index.append(line[1])  # create a list of klusterbox names
                self.id_index.append(line[2].zfill(8))  # create a list of emp ids

        def get_carrier_list(self):
            """ this method creates a list of carriers who are currently at the station during the dates of the
            investigation range. This is stored in the self.c_list array"""
            carrier_list = gen_carrier_list()  # generate an in range carrier list
            for each in carrier_list:  # create a list of unique names from carrier list (a set)
                if each[1] not in self.c_list:
                    self.c_list.append(each[1])

        def get_tacslist(self):  # Get the names from tacs report and create tacs_list
            good_jobs = ("134", "844", "434")
            cc = 0
            for line in self.parent.a_file:
                if cc > 1 and line[19] in good_jobs:
                    # create a note for carrier's assignment - reg w/route, reg floater or aux
                    route = line[25].zfill(6)
                    lvl = line[23].zfill(2)
                    if line[19] == "134" and lvl == "01":
                        tac_route = route[1] + route[2] + route[3] + route[4] + route[5]
                        assignment = "reg " + Handler(tac_route).routes_adj()
                    elif line[19] == "134" and lvl == "02":
                        assignment = "reg " + "floater"
                    elif line[19] == "434":
                        assignment = "part time flex"
                    elif line[19] == "844":
                        assignment = "auxiliary"
                    else:
                        assignment = "undetected"
                    lastname = line[5].lower().replace("\'", " ")
                    add_to_list = [line[4].zfill(8), lastname, line[6].lower(),
                                   assignment]  # create list to insert in list
                    self.parent.tacs_list.append(add_to_list)
                cc += 1

        def remove_tacs_duplicates(self):
            holder = ["", "", "", ""]  # find the duplicates and remove them where there is both BASE and TEMP
            put_back = []
            for item in self.parent.tacs_list:  # crawler goes down the list to identify Temp entries
                if item[0] == holder[0]:
                    if item == holder:
                        self.to_remove.append(holder)  # remove both records
                    if item != holder:
                        self.to_remove.append(holder)
                        self.to_remove.append(item)
                    put_back.append(item)  # put the later record back in the list
                holder = item  # hold the record to compare in the next loop
            # remove the duplicates
            self.parent.tacs_list = [x for x in self.parent.tacs_list if x not in self.to_remove]  
            for record in put_back:  # put the Temp record back into the tacs_list
                self.parent.tacs_list.append(record)
            self.parent.tacs_list.sort(key=itemgetter(1))  # re-alphabetize the list of carriers

        def qualify_tacslist(self):
            sql = ""
            add = 0  # create tallies for reports
            rec = 0
            out = 0
            # carriers who are already or newly placed in name index - remove them from further processing
            self.to_remove = []
            pb = ProgressBarDe(title="Database Maintenance", label="Updating Changes: ")  # create progressbar
            pb.max_count(len(self.parent.tacs_list))  # set length of progress bar
            pb.start_up()
            i = 0
            for each in self.parent.tacs_list:
                pb.move_count(i)  # increment progress bar
                tac_str = "{}, {}".format(each[1], each[2])  # tac str is last name and first initial from tacs report
                # if there is an identical match between kb and tacs names:
                if tac_str in self.c_list and each[0] not in self.id_index:
                    # if there is a dup name / need a complete list of carrier names from index
                    if tac_str in self.name_index:
                        # maybe just pass information via new_carrier and add later
                        self.parent.new_carrier.append(each)  
                    else:  # go ahead and pair the emp id with the name in carriers
                        sql = "INSERT INTO name_index (tacs_name, kb_name, emp_id ) VALUES('%s','%s','%s')" \
                              % (tac_str, tac_str, each[0])
                        self.name_index.append(tac_str)
                        self.id_index.append(each[0])
                    add += 1
                    commit(sql)
                    self.to_remove.append(each[0])
                    self.name_index.append(tac_str)
                elif each[0] in self.id_index:  # RECOGNIZED -  the emp id is already in the name index
                    self.to_remove.append(each[0])
                    self.parent.check_these.append(each)
                    rec += 1
                else:
                    out += 1
                i += 1
            pb.stop()  # stop and destroy the progress bar

        def get_new_carrier(self):
            # find the carriers in name_index who have records w/ eff dates in the future
            dont_check = []  # remove items from check these if future carriers are found
            for name in self.parent.check_these:
                sql = "SELECT kb_name FROM name_index WHERE emp_id = '%s'" % name[0]
                result = inquire(sql)
                kb_name = result[0][0]
                sql = "SELECT effective_date,carrier_name FROM carriers " \
                      "WHERE carrier_name = '%s' AND effective_date <= '%s' " \
                      "ORDER BY effective_date DESC" % (kb_name, projvar.invran_date_week[0])
                result = inquire(sql)
                if not result:
                    self.parent.new_carrier.append(name)  # will add as new carrier in AI 3
                    dont_check.append(name[0])  # removes from check these array
                    self.to_remove.append(name[0])  # removes from tacs list
            # removes don't check from check these
            self.parent.check_these = [x for x in self.parent.check_these if x[0] not in dont_check]

        def limit_tacslist(self):
            self.parent.tacs_list = [x for x in self.parent.tacs_list if x[0] not in self.parent.new_carrier]
            self.parent.tacs_list = [x for x in self.parent.tacs_list if x[0] not in self.to_remove]

        def get_name_index(self):
            sql = "SELECT tacs_name, kb_name, emp_id FROM name_index ORDER BY kb_name"
            results = inquire(sql)
            for item in self.name_index:
                self.parent.tried_names.append(item)
            self.name_index = []  # create a list of klusterbox names
            for line in results:
                self.name_index.append(line[1])

        def namepairing_router(self):  # route to appropriate function based on array contents
            # all tacs list resolved/ nothing to check
            if len(self.parent.tacs_list) < 1 and len(self.parent.new_carrier) < 1 and len(self.parent.check_these) < 1:
                self.parent.AutoIndexer6(self.parent).run(self.frame)  # to straight to entering rings
            # all tacs list resolved/ new names unresolved
            elif len(self.parent.tacs_list) < 1 and len(self.parent.new_carrier) > 0:
                self.parent.AutoIndexer4(self.parent).run(self.frame)  # add new carriers in AI6
            # tacs and new carriers resolved/ carriers to check
            elif len(self.parent.tacs_list) < 1 and len(self.parent.new_carrier) < 1 and \
                    len(self.parent.check_these) > 0:
                # step to AI  to check discrepancies
                self.parent.AutoIndexer5(self.parent).run(self.frame)
            else:  # If there are candidates sort, generate PAIRING SCREEN 1
                self.namepairing_screen()

        def namepairing_screen(self):  # Pairing screen #1
            self.c_list = [x for x in self.c_list if x not in self.name_index]
            self.win = MakeWindow()
            self.win.create(self.frame)
            Label(self.win.body, text="Search for Name Matches #1", font=macadj("bold", "Helvetica 18"), pady=10) \
                .grid(row=0, column=0, sticky="w", columnspan=10)  # page contents
            wintext = "Look for possible matches for each unrecognized name. If the name has already been entered " \
                      "manually, you \n should be able to find it on this screen or the next. It is possible " \
                      "that the " \
                      "name has no match, if that is \n the case then select \"ADD NAME\" in the next screen. " \
                      "You can " \
                      "change the default between \"NOT FOUND\" and \n \"DISCARD\" using the buttons below. " \
                      "Information from TACS is shown in blue\n\n"
            mactext = \
                "Look for possible matches for each unrecognized name. If the name has already been \n" \
                "entered manually, you should be able to find it on this screen or the next. It is \n" \
                "possible that the name has no match, if that is the case then select \"ADD NAME\" in \n" \
                "the next screen. You can change the default between \"NOT FOUND\" and \"DISCARD\" \n" \
                "using the buttons below. Information from TACS is shown in blue\n\n"
            text = "Investigation Range: {0} through {1}\n\n".format(
                projvar.invran_date_week[0].strftime("%a - %b %d, %Y"),
                projvar.invran_date_week[6].strftime("%a - %b %d, %Y"))
            Label(self.win.body, text=macadj(wintext, mactext) + text, justify=LEFT) \
                .grid(row=1, column=0, columnspan=10, sticky="w")
            Button(self.win.body, text="DISCARD", width=10,
                   command=lambda: self.indexer_default(self.parent.name_sorter, i + 1, name_options, 1)) \
                .grid(row=2, column=3, sticky="w", columnspan=2)
            Label(self.win.body, text="switch default to DISCARD").grid(row=2, column=1, sticky="w", columnspan=2)
            Button(self.win.body, text="NOT FOUND", width=10,
                   command=lambda: self.indexer_default(self.parent.name_sorter, i + 1, name_options, 0)) \
                .grid(row=3, column=3, sticky="w", columnspan=2)
            Label(self.win.body, text="switch default to NOT FOUND").grid(row=3, column=1, sticky="w", columnspan=2)
            Label(self.win.body, text="").grid(row=4, column=0)
            Label(self.win.body, text="Name", fg="grey").grid(row="5", column="1", sticky="w")
            Label(self.win.body, text="Assignment", fg="grey").grid(row="5", column="2", sticky="w")
            Label(self.win.body, text="Candidates", fg="grey").grid(row="5", column="3", sticky="w")
            cc = 6
            i = 0
            color = "blue"
            for t_name in self.parent.tacs_list:  # for each name in the tacs report
                self.get_possible_names(t_name)  # fill the self.possible names array
                if self.possible_names:
                    Label(self.win.body, text=str(i + 1), anchor="w").grid(row=cc, column=0, sticky="w")
                    Label(self.win.body, text=t_name[1] + ", " + t_name[2], anchor="w", width=15, fg=color) \
                        .grid(row=cc, column=1, sticky="w")  # name
                    Label(self.win.body, text=t_name[3], anchor="w", width=10, fg=color) \
                        .grid(row=cc, column=2, sticky="w")  # assignment
                name_options = ["NOT FOUND", "DISCARD"] + self.possible_names
                self.parent.name_sorter.append(StringVar(self.win.body))
                option_menu = OptionMenu(self.win.body, self.parent.name_sorter[i], *name_options)
                self.parent.name_sorter[i].set(name_options[0])
                option_menu.config(width=15)
                if self.possible_names:
                    option_menu.grid(row=cc, column=3, sticky="w")  # possible matches
                    if len(self.possible_names) == 1:  # display indicator for possible matches
                        Label(self.win.body, text=str(len(self.possible_names)) + " name")\
                            .grid(row=cc, column=4, sticky="w")
                    elif len(self.possible_names) > 1:
                        Label(self.win.body, text=str(len(possible_names)) + " names")\
                            .grid(row=cc, column=4, sticky="w")
                    else:  # display indicator for possible matches
                        Label(self.win.body, text="no match", fg="grey").grid(row=cc, column=4, sticky="w")
                cc += 1
                i += 1
            Button(self.win.buttons, text="Continue", width=macadj(15, 16),
                   command=lambda: self.parent.AutoIndexer3(self.parent).run(self.win.topframe)).grid(row=0, column=0)
            Button(self.win.buttons, text="Cancel", width=macadj(15, 16),
                   command=lambda: self.parent.go_back(self.win.topframe)).grid(row=0, column=1)
            if not self.possible_match:  # if there are no possible matches to any carrier names
                self.parent.AutoIndexer3(self.parent).run(self.win.topframe)  # go to next screen
            else:
                self.win.finish()  # otherwise stay on this screen

        def get_possible_names(self, t_name):
            self.possible_names = []
            for c_name in self.c_list:
                """ if the first letter of a carrier name has no record in the name index and matches the 
                 carrier name from the tacs report - append the name to the possible names array so it cann 
                 be used in the option menu. """
                if c_name[0:3] == t_name[1][0:3]:
                    self.possible_names.append(c_name)
                    self.parent.tried_names.append(c_name)
                    self.possible_match = True

        @staticmethod
        def indexer_default(widget, count, options, choice):  # changes the default for the optionmenu widget
            for i in range(count - 1):
                widget[i].set(options[choice])
            
    class AutoIndexer3:  # Carrier pairing screen -
        # allows users to match new carrier entries to carriers already in klusterbox.
        def __init__(self, parent):
            self.parent = parent
            self.frame = None
            self.to_remove = []  # intialized array of names to be removed from tacs names
            self.to_nameindex = []  # initialize array of names to be be paired in name index
            self.c_list = []
            self.win = None
            self.n_index = []  # an array of all klusterbox names with records in the name index table
            self.to_chg = []  # changes for apply ai3
            self.new_name = []  # array of new names which have been modified with emp id

        def run(self, frame):
            self.frame = frame
            self.apply_namepairing_1()  # apply pairing screen
            # if empty tacs list and something in check these
            if len(self.parent.tacs_list) < 1 and len(self.parent.check_these) > 0:  
                self.parent.AutoIndexer5(self.parent).run(self.frame)
            elif len(self.parent.tacs_list) < 1 and len(self.parent.check_these) < 1:
                self.parent.AutoIndexer6(self.parent).run(self.frame)
            else:
                self.build_namepairing_options()
                self.namepairing_screen_2()  # create pairing screen #2
    
        def apply_namepairing_1(self):  # apply pairing screen #1 / AutoIndexer 2
            i = 0  # count iterations of loops
            dis = 0  # count of discarded items
            out = 0  # count of unresolved items
            pair = 0  # count of added items
            self.to_remove = []  # intialized array of names to be removed from tacs names
            not_found = []  # initialize array of names to be futher analyzed.
            for item in self.parent.name_sorter:
                if item.get() == "DISCARD":
                    self.to_remove.append(self.parent.tacs_list[i][0])
                    dis += 1
                elif item.get() == "NOT FOUND":
                    not_found.append(self.parent.tacs_list[i])
                    out += 1
                else:
                    to_add = [self.parent.tacs_list[i], item.get()]
                    self.to_nameindex.append(to_add)
                    self.to_remove.append(self.parent.tacs_list[i][0])
                    self.parent.check_these.append(self.parent.tacs_list[i])
                    pair += 1
                i += 1
            self.parent.tacs_list = [x for x in self.parent.tacs_list if x[0] not in self.to_remove]
            for item in self.to_nameindex:
                # tac str is last name and first initial from tacs report
                tac_str = "{}, {}".format(item[0][1], item[0][2])  
                sql = "INSERT INTO name_index (tacs_name, kb_name, emp_id) VALUES('%s','%s','%s')" \
                      % (tac_str, item[1], item[0][0])
                commit(sql)
            """
        # message screens to summerize output
            messagebox.showinfo("Processing Carriers", 
            "{} Carrier names were paired to names in klusterbox\n"
            "{} Carrier names were discarded.\n"
            "{} Carrier names have not been handled."
            .format(pair, dis, out), parent=frame)
            """
    
        def build_namepairing_options(self):  # build possible names for option menus
            sql = "SELECT kb_name FROM name_index"
            results = inquire(sql)
            name_result = []  # create a list of klusterbox names
            for line in results:
                name_result.append(line[0])
            sql = "SELECT carrier_name FROM carriers ORDER BY carrier_name"  # get all names from the carrier list
            results = inquire(sql)  # call function to access database
            for item in results:
                if item[0] not in self.c_list and item[0] not in self.parent.tried_names and item[0] not in name_result:
                    self.c_list.append(item[0])
    
        def namepairing_screen_2(self):  # create pairing screen #2
            self.win = MakeWindow()
            self.win.create(self.frame)
            self.parent.name_sorter = []  # page contents
            Label(self.win.body, text="Search for Name Matches #2", font=macadj("bold", "Helvetica 18"), pady=10) \
                .grid(row=0, column=0, sticky="w", columnspan=10)  # page contents
            wintext = \
                "Look for possible matches for each unrecognized name. If the name has already been entered " \
                "manually, \n" \
                " you should be able to find it on this screen. It is possible that the name has no match, if " \
                "that is \n" \
                "the case then select \"ADD NAME\" in this screen. You can change the default between \"ADD NAME\" " \
                "and \n" \
                "\"DISCARD\" using the buttons below. Information from TACS is shown in blue\n\n"
            mactext = \
                "Look for possible matches for each unrecognized name. If the name has already been \n" \
                "entered manually, you should be able to find it on this screen. It is possible that \n" \
                "the name has no match, if that is the case then select \"ADD NAME\" in this screen. \n" \
                "You can change the default between \"ADD NAME\" and \"DISCARD\" using the buttons \n" \
                "below. Information from TACS is shown in blue\n\n"
            text = "Investigation Range: {0} through {1}\n\n".format(
                projvar.invran_date_week[0].strftime("%a - %b %d, %Y"),
                projvar.invran_date_week[6].strftime("%a - %b %d, %Y"))
            Label(self.win.body, text=macadj(wintext, mactext) + text, justify=LEFT)\
                .grid(row=1, column=0, columnspan=10, sticky="w")
            Button(self.win.body, text="DISCARD", width=10,
                   command=lambda: self.indexer_default(self.parent.name_sorter, i + 1, name_options, 1)) \
                .grid(row=2, column=3, sticky="w", columnspan=2)
            Label(self.win.body, text="switch default to DISCARD").grid(row=2, column=1, sticky="w", columnspan=2)
            Button(self.win.body, text="ADD NAME", width=10,
                   command=lambda: self.indexer_default(self.parent.name_sorter, i + 1, name_options, 0)) \
                .grid(row=3, column=3, sticky="w", columnspan=2)
            Label(self.win.body, text="switch default to ADD NAME").grid(row=3, column=1, sticky="w", columnspan=2)
            Label(self.win.body, text="").grid(row=4, column=0)
            Label(self.win.body, text="Name", fg="grey").grid(row="5", column="1", sticky="w")
            Label(self.win.body, text="Assignment", fg="grey").grid(row="5", column="2", sticky="w")
            Label(self.win.body, text="Candidates", fg="grey").grid(row="5", column="3", sticky="w")
            cc = 6  # item and grid row counter
            i = 0  # count iterations of the loop
            color = "blue"
            for t_name in self.parent.tacs_list:
                possible_names = []
                Label(self.win.body, text=str(i + 1), anchor="w").grid(row=cc, column=0)
                Label(self.win.body, text=t_name[1] + ", " + t_name[2], anchor="w", width=15, fg=color)\
                    .grid(row=cc, column=1)  # name
                Label(self.win.body, text=t_name[3], anchor="w", width=10, fg=color)\
                    .grid(row=cc, column=2)  # assignment
                # build option menu for unmatched tacs names
                for c_name in self.c_list:
                    if c_name[0] == t_name[1][0]:
                        possible_names.append(c_name)
                name_options = ["ADD NAME", "DISCARD"] + possible_names
                self.parent.name_sorter.append(StringVar(self.win.body))
                option_menu = OptionMenu(self.win.body, self.parent.name_sorter[i], *name_options)
                self.parent.name_sorter[i].set(name_options[0])
                option_menu.config(width=15)
                option_menu.grid(row=cc, column=3)  # possible matches
                if len(possible_names) == 1:  # display indicator for possible matches
                    Label(self.win.body, text=str(len(possible_names)) + " name").grid(row=cc, column=4)
                if len(possible_names) > 1:
                    Label(self.win.body, text=str(len(possible_names)) + " names").grid(row=cc, column=4)
                cc += 1
                i += 1
            Button(self.win.buttons, text="Continue", width=macadj(15, 16),
                   command=lambda: self.ai3_apply()) \
                .grid(row=0, column=0)
            Button(self.win.buttons, text="Cancel", width=macadj(15, 16),
                   command=lambda: self.parent.go_back(self.win.topframe)).grid(row=0, column=1)
            self.win.finish()
    
        @staticmethod
        def indexer_default(widget, count, options, choice):  # changes the default for the optionmenu widget
            for i in range(count - 1):
                widget[i].set(options[choice])
    
        def ai3_apply(self):  # apply pairing screen 2
            self.build_n_index()
            self.ai3_apply_sort()  # discard, add or pair name
            self.insert_to_nameindex()  # add names to name index
            self.insert_to_addname()  # add names to name index
            # self.apply_ai3_report()  # message screens to summerize output
            self.build_addname()  # build to_addname array
            if len(self.parent.to_addname) > 0:  # route conditional on arrays
                self.parent.AutoIndexer4(self.parent).run(self.win.topframe)
            elif len(self.parent.check_these) > 0:
                self.parent.AutoIndexer5(self.parent).run(self.win.topframe)
            else:
                self.parent.AutoIndexer6(self.parent).run(self.win.topframe)
    
        def build_n_index(self):
            """ creates an array of all klusterbox names with records in the name index table. """
            sql = "SELECT tacs_name, kb_name, emp_id FROM name_index ORDER BY kb_name"
            results = inquire(sql)
            self.n_index = []  # create a list of klusterbox names
            for line in results:  # loop to fill arrays
                self.n_index.append(line[1])
    
        def ai3_apply_sort(self):  # discard, add or pair name
            i = 0  # count iterations of the loops..
            for item in self.parent.name_sorter:  # sort passed data from auto index 4
                if item.get() == "DISCARD":
                    self.to_remove.append(self.parent.tacs_list[i][0])
                    # dis += 1  # count of discarded items
                elif item.get() == "ADD NAME":
                    self.parent.to_addname.append(self.parent.tacs_list[i])
                    # add += 1
                else:
                    to_add = [self.parent.tacs_list[i], item.get()]
                    self.to_nameindex.append(to_add)
                    self.to_remove.append(self.parent.tacs_list[i][0])
                    self.parent.check_these.append(self.parent.tacs_list[i])
                    # pair += 1  # count of paired items
                i += 1
    
        def insert_to_nameindex(self):  # add names to name index
            pb_label = Label(self.win.buttons, text="Updating Changes: ")  # make label for progress bar
            pb_label.grid(row=0, column=2)
            pb = ttk.Progressbar(self.win.buttons, length=250, mode="determinate")  # create progress bar
            pb.grid(row=0, column=3)
            pb["maximum"] = len(self.to_nameindex)  # set length of progress bar
            pb.start()
            i = 0
            for item in self.to_nameindex:  # when a name from the optionmenu was selected
                if self.no_record(item[0][0]):  # check for a record in name index by employee id #
                    # tac str is last name and first initial from tacs report
                    tac_str = "{}, {}".format(item[0][1], item[0][2])
                    sql = "INSERT INTO name_index (tacs_name, kb_name, emp_id) VALUES('%s','%s','%s')" \
                          % (tac_str, item[1], item[0][0])
                    commit(sql)
                pb["value"] = i  # increment progress bar
                self.win.buttons.update()  # update the progress bar
                i += 1
            pb.stop()  # stop and destroy the progress bar
            pb_label.destroy()  # destroy the label for the progress bar
            pb.destroy()

        @staticmethod
        def no_record(empid):  # check for a record in name index by employee id #
            sql = "SELECT emp_id FROM name_index WHERE emp_id = '%s'" % empid
            result = inquire(sql)
            if result:
                return False  # if there is a record
            return True  # if there is no record
    
        def insert_to_addname(self):  # add names to name index
            self.to_chg = []  # array of items from to_addname where the name needs to be modified with emp id
            self.new_name = []  # array of new names which have been modified with emp id
            for name in self.parent.new_carrier:
                self.parent.to_addname.append(name)  # add new carriers in list to be added to carrier table
            pb_label = Label(self.win.buttons, text="Updating Changes: ")  # make label for progress bar
            pb_label.grid(row=0, column=2)
            pb = ttk.Progressbar(self.win.buttons, length=250, mode="determinate")  # create progress bar
            pb.grid(row=0, column=3)
            pb["maximum"] = len(self.parent.to_addname)  # set length of progress bar
            pb.start()
            i = 0
            for item in self.parent.to_addname:  # when add name was selected from option menu
                pb["value"] = i  # increment progress bar
                tacs_str = "{}, {}".format(item[1], item[2])  # tacs str is last name and first initial from tacs report
                kb_str = "{}, {}".format(item[1], item[2])  # kb str is last name and first initial from tacs report
                if kb_str in self.n_index or kb_str in self.c_list:  # detect matches with name index
                    sql = "SELECT emp_id, kb_name FROM name_index WHERE emp_id = '%s'" % item[0]
                    result = inquire(sql)
                    if not result:
                        kb_str = "{} {}".format(kb_str, item[0])
                        self.to_chg.append(item)
                        mod_name = "{} {}".format(item[2], item[0])
                        self.new_name.append(mod_name)
                    if result:  # if the carrier is in the name index
                        # if the kb name is not the same in the name index record - change name
                        if result[0][1] != kb_str:  
                            self.to_chg.append(item)
                            mod_name = result[0][1].split(",")
                            mod_name = mod_name[1].strip()
                            self.new_name.append(mod_name)
                self.n_index.append(kb_str)  # add to n_index array so dups can be detected
                sql = "SELECT emp_id FROM name_index WHERE emp_id = '%s'" % item[0]
                result = inquire(sql)
                if not result:
                    sql = "INSERT INTO name_index (tacs_name, kb_name, emp_id) VALUES('%s','%s','%s')" \
                          % (tacs_str, str(kb_str), item[0])
                    commit(sql)
                self.win.buttons.update()  # update the progress bar
                i += 1
            pb.stop()  # stop and destroy the progress bar
            pb_label.destroy()  # destroy the label for the progress bar
            pb.destroy()
    
        def apply_ai3_report(self):  # message screens to summerize output
            messagebox.showinfo("Processing Carriers", "{} Carrier names were added to the database\n"
                                                       "{} Carrier names were paired to names in klusterbox\n"
                                                       "{} Carrier names were discarded.\n"
                                                       .format(len(self.parent.to_addname), len(self.to_nameindex),
                                                       len(self.to_remove)), parent=self.win.topframe)
    
        def build_addname(self):  # build to_addname array
            count = 0  # swap out the names which have been modified in self.parent.to_addname
            for item in self.to_chg:  # for each item to be swapped
                self.parent.to_addname.remove(item)  # clear out the old one
                # create a modified array with modified name
                mod_str = [item[0], item[1], self.new_name[count], item[3]]
                self.parent.to_addname.append(mod_str)  # put in the new one
                count += 1

    class AutoIndexer4:  # input new carrier information after a check
        def __init__(self, parent):
            self.parent = parent
            self.frame = None
            self.opt_nsday = []  # make an array of "day / color" options for option menu
            self.full_ns_dict = {}
            self.ns_dict = {}  # create dictionary for ns day data
            self.eff_date = None  # effective date for apply
            self.station = None  # station as stringvar for apply
            self.changecount = None
            self.win = None
            self.ai4_carrier_name = []  # create array for carrier names
            self.ai4_l_s = []  # create array for list status
            self.ai4_l_ns = []  # create array for ns days
            self.ai4_route = []  # create array for route/s

        def run(self, frame):  # add new carriers to carrier table / pairing screen #3
            self.frame = frame
            self.ai4_opt_nsday()
            self.ai4_full_ns_dict()
            self.ai4_ns_dict()
            self.ai4_screen()

        def ai4_opt_nsday(self):  # get ns structure preference from database
            sql = "SELECT tolerance FROM tolerances WHERE category='%s'" % "ns_auto_pref"
            result = inquire(sql)
            ns_toggle = result[0][0]  # modify available ns days per ns_toggle
            if ns_toggle == "rotation":
                remove_array = ("sat", "mon", "tue", "wed", "thu", "fri")
            else:
                remove_array = ("green", "brown", "red", "black", "yellow", "blue")
            ns_code_mod = dict()  # copy the projvar.ns_code dict to ns_code_mod using dict()
            for key in projvar.ns_code:
                ns_code_mod[key] = projvar.ns_code[key]
            for key in remove_array:
                if key in ns_code_mod:
                    del ns_code_mod[key]  # modify available ns days per ns_toggle
            for each in ns_code_mod:  #
                ns_option = ns_code_mod[each] + " - " + each  # make a string for each day/color
                if each == "none":
                    ns_option = "       " + " - " + each  # if the ns day is "none" - make a special string
                self.opt_nsday.append(ns_option)

        def ai4_full_ns_dict(self):
            days = ("Saturday", "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday")
            for each in self.opt_nsday:  # Make a dictionary to match full days and option menu options
                for day in days:
                    if day[:3] == each[:3]:
                        self.full_ns_dict[day] = each  # creates full_ns_dict
                if each[-4:] == "none":
                    ns_option = "       " + " - " + "none"  # if the ns day is "none" - make a special string
                    self.full_ns_dict["None"] = ns_option  # creates full_ns_dict None option

        def ai4_ns_dict(self):
            results = gen_ns_dict(self.parent.file_path, self.parent.to_addname)  # returns id and name
            for ids in results:  # loop to fill dictionary with ns day info
                self.ns_dict[ids[0]] = ids[1]
            return self.ns_dict

        def ai4_screen(self):
            self.win = MakeWindow()
            self.win.create(self.frame)
            Label(self.win.body, text="Input New Carriers", font=macadj("bold", "Helvetica 18"), pady=10) \
                .grid(row=0, column=0, sticky="w", columnspan=6)  # Pairing Screen #3
            wintext = \
                "Enter in information for carriers not already recorded in the Klusterbox database. You can use " \
                "the TACS \n" \
                "information (shown in blue),as a guide if it is accurate. As OTDL/WAL information is not in TACS, " \
                "it is \n" \
                "not shown and this information will have to requested from management. Routes must be only 4 " \
                "digits \n" \
                "long. In cases were there are multiple routes, the routes must be separated by a \"/\" backslash.\n\n"
            mactext = \
                "Enter in information for carriers not already recorded in the Klusterbox database. You can \n" \
                "use the TACS information (shown in blue),as a guide if it is accurate. As OTDL/WAL \n" \
                "information is not in TACS, it is not shown and this information will have to requested \n" \
                "from management. Routes must be only 4 digits long. In cases were there are multiple \n" \
                "routes, the routes must be separated by a \"/\" backslash.\n\n"
            text = "Investigation Range: {0} through {1}\n\n"\
                .format(projvar.invran_date_week[0].strftime("%a - %b %d, %Y"),
                        projvar.invran_date_week[6].strftime("%a - %b %d, %Y"))
            # is_mac = macadj(False, True)
            Label(self.win.body, text=macadj(wintext, mactext) + text, justify=LEFT)\
                .grid(row=1, column=0, sticky="w", columnspan=6)
            y = 2  # count for the row
            Label(self.win.body, text="Name", fg="Grey").grid(row=y, column=0, sticky="w")
            Label(self.win.body, text=macadj("List Status", "List"), fg="Grey").grid(row=y, column=1, sticky="w")
            Label(self.win.body, text="NS Day", fg="Grey").grid(row=y, column=2, sticky="w")
            Label(self.win.body, text="Route_s", fg="Grey").grid(row=y, column=3, sticky="w")
            if not self.parent.is_mac:
                Label(self.win.body, text="Station", fg="Grey").grid(row=y, column=4, sticky="w")
                Label(self.win.body, text="              ", fg="Grey").grid(row=y, column=5, sticky="w")
            y += 1
            i = 0  # count the instances of the array
            color = "blue"
            for name in self.parent.to_addname:
                Label(self.win.body, text=name[1] + ", " + name[2], fg=color).grid(row=y, column=0, sticky="w")
                self.ai4_carrier_name.append(str(name[1] + ", " + name[2]))
                Label(self.win.body, text=macadj("not in record", "unknown"), fg=color)\
                    .grid(row=y, column=1, sticky="w")
                Label(self.win.body, text=str(self.ns_dict[name[0]]), fg=color).grid(row=y, column=2, sticky="w")
                Label(self.win.body, text=name[3], fg=color).grid(row=y, column=3, sticky="w")
                if not self.parent.is_mac:
                    Label(self.win.body, text=projvar.invran_station, fg=color).grid(row=y, column=4, sticky="w")
                y += 1
                list_options = ("otdl", "wal", "nl", "ptf", "aux")  # create optionmenu for list status
                if name[3] == "auxiliary":
                    lx = 4  # configure defaults for list status
                elif name[3] == "part time flex":
                    lx = 3  # set as ptf
                else:
                    lx = 2  # set as 'nl' if not 'aux'
                self.ai4_l_s.append(StringVar(self.win.body))
                self.ai4_l_s[i].set(list_options[lx])  # set the list status
                list_status = OptionMenu(self.win.body, self.ai4_l_s[i], *list_options)
                list_status.config(width=macadj(5, 4))
                list_status.grid(row=y, column=1, sticky="w")
                self.ai4_l_ns.append(StringVar(self.win.body))  # create optionmenu for ns days
                self.ai4_l_ns[i].set(self.full_ns_dict[str(self.ns_dict[name[0]])])  # set ns day default
                ns_day = OptionMenu(self.win.body, self.ai4_l_ns[i], *self.opt_nsday)
                ns_day.config(width=macadj(12, 10))
                ns_day.grid(row=y, column=2, sticky="w")
                self.ai4_route.append(StringVar(self.win.body))  # create entry field for route
                # create entry for routes
                Entry(self.win.body, width=24, textvariable=self.ai4_route[i]).grid(row=y, column=3, sticky="w")
                if "reg " in name[3] and name[3] != "reg floater":
                    rte = name[3].replace("reg ", "")
                else:
                    rte = ""
                self.ai4_route[i].set(rte)
                y += 1
                i += 1
                Label(self.win.body, text="").grid(row=y, column=0, sticky="w")
                y += 1
            Button(self.win.buttons, text="Continue", width=macadj(15, 16),
                   command=lambda: self.ai4_apply()).pack(side=LEFT)
            Button(self.win.buttons, text="Cancel", width=macadj(15, 16),
                   command=lambda: self.parent.go_back(self.win.topframe)).pack(side=LEFT)
            self.win.finish()

        def ai4_apply(self):  # adds new carriers to the carriers table
            self.ai4_date()  # get the effective date
            self.ai4_station()  # get the station as a stringvar (apply2 reads station as stringvar)
            if self.ai4_check():
                self.ai4_count_change()
                # route conditional to arrays
                if len(self.changecount) >= len(self.ai4_carrier_name) and len(self.parent.check_these) > 0:
                    self.parent.AutoIndexer5(self.parent).run(self.win.topframe)
                elif len(self.changecount) >= len(self.ai4_carrier_name):
                    self.parent.AutoIndexer6(self.parent).run(self.win.topframe)
                else:
                    return
            else:
                frame = self.win.topframe  # prevent the object from being obliterated by rerunning __init__
                self.__init__(self.parent)  # re initialize the child class
                # self.re_init(self.parent)  # re initialize the child class
                # self.run(self.win.topframe)
                self.run(frame)

        def ai4_date(self):  # get the effective date
            self.eff_date = projvar.invran_date_week[0]  # if investigation range is weekly
            if not projvar.invran_weekly_span:  # if investigation range is daily
                self.eff_date = projvar.invran_date

        def ai4_station(self):  # get the station as a stringvar (apply2 reads station as stringvar)
            self.station = StringVar(self.win.body)  # put station var in a StringVar object
            self.station.set(projvar.invran_station)

        def ai4_check(self):  # check and enter carrier info
            pb_label = Label(self.win.buttons, text="Updating Changes: ")  # make label for progress bar
            pb_label.pack(side=LEFT)
            pb = ttk.Progressbar(self.win.buttons, length=250, mode="determinate")  # create progress bar
            pb.pack(side=LEFT)
            pb["maximum"] = len(self.ai4_carrier_name)  # set length of progress bar
            pb.start()
            for i in range(len(self.ai4_carrier_name)):
                pb["value"] = i  # increment progress bar
                passed_ns = self.ai4_l_ns[i].get().split(" - ")  # clean the passed ns day data
                clean_ns = StringVar(self.win.body)  # put ns day var in StringVar object
                clean_ns.set(passed_ns[1])
                # check moves/route and enter data into rings table
                if not apply_2(self.eff_date, self.ai4_carrier_name[i], self.ai4_l_s[i], clean_ns,
                               self.ai4_route[i], self.station, self.win.body):
                    return False
                self.win.buttons.update()
            pb.stop()  # stop and destroy the progress bar
            pb_label.destroy()  # destroy the label for the progress bar
            pb.destroy()
            return True

        def ai4_count_change(self):  # get count of carrier changes for current day
            self.changecount = []
            for name in self.ai4_carrier_name:
                sql = "SELECT * FROM carriers WHERE carrier_name == '%s' and effective_date == '%s'" \
                      % (name, self.eff_date)
                result = inquire(sql)
                if result:
                    self.changecount.append(result)

    class AutoIndexer5:  # discrepancy resolution screen
        def __init__(self, parent):
            self.parent = parent
            self.frame = None
            self.opt_nsday = []  # make an array of "day / color" options for option menu
            self.ns_opt_dict = {}  # creates a dictionary of ns colors/ options for menu
            self.full_ns_dict = {}
            self.ns_dict = {}  # create dictionary for ns day data
            self.name_dict = {}  # generate dictionary for emp id to kb_name
            self.carriers_names_list = []  # generate list of only names from 'in range carrier list'
            self.ai5_carrier_list = None
            self.code_ns = None
            self.win = None
            self.y = 1  # count for the row
            self.i = 0  # count the instances of the array for the screen
            self.carrier_name = []  # create array for carrier names  # attributes for screen
            self.l_s = []  # create array for list status
            self.l_ns = []  # create array for ns days
            self.e_route = []  # create array for routes
            self.l_station = []  # create array for stations
            self.aux_list_tuple = ("aux", "ptf")
            self.reg_list_tuple = ("nl", "wal", "otdl")
            self.skip_this_screen = True
            self.color = "blue"  # the display color of information from tacs

        def run(self, frame):
            self.frame = frame
            if len(self.parent.check_these) == 0:
                self.parent.AutoIndexer6(self.parent).run(self.frame)
            else:
                self.parent.check_these.sort(key=itemgetter(1))  # sort the incoming tacs information
                self.ai5_opt_nsday()  # creates the option menu options for ns day menu
                self.ai5_ns_dict()  # create dictionary for ns day data
                self.ai5_nameindex_dict()  # generate dictionary for emp id to kb_name
                self.ai5_carrierlist()  # generate list of only names from 'in range carrier list'
                self.ai5_nscode()  # generate reverse ns code dict
                self.ai5_screen()

        def ai5_opt_nsday(self):
            for each in projvar.ns_code:  # creates the option menu options for ns day menu
                ns_option = projvar.ns_code[each] + " - " + each  # make a string for each day/color
                self.ns_opt_dict[each] = ns_option
                if each == "none":
                    ns_option = "       " + " - " + each  # if the ns day is "none" - make a special string
                    self.ns_opt_dict[each] = ns_option
                self.opt_nsday.append(ns_option)

        def ai5_ns_dict(self):  # create dictionary for ns day data
            results = gen_ns_dict(self.parent.file_path, self.parent.check_these)  # returns id and name
            for id in results:  # loop to fill dictionary with ns day info
                self.ns_dict[id[0]] = id[1]
                
        def ai5_nameindex_dict(self):  # generate dictionary for emp id to kb_name
            sql = "SELECT tacs_name, kb_name, emp_id FROM name_index ORDER BY kb_name"
            results = inquire(sql)
            for line in results:  # loop to fill arrays
                self.name_dict[line[2]] = line[1]

        def ai5_carrierlist(self):  # generate list of only names from 'in range carrier list'
            self.ai5_carrier_list = gen_carrier_list()  # generate an in range carrier list
            for name in self.ai5_carrier_list:
                self.carriers_names_list.append(name[1])
            remainders = []  # find carriers in 'check these' but not in 'in range carrier list' aka 'remainders'
            for name in self.parent.check_these:
                if self.name_dict[name[0]] not in self.carriers_names_list:
                    remainders.append(name)
            for name in remainders:  # get carriers data from carriers for remainders
                sql = "SELECT * FROM carriers WHERE carrier_name = '%s' and effective_date <= '%s'" \
                      "ORDER BY effective_date desc" % (self.name_dict[name[0]], projvar.invran_date_week[0])
                result = inquire(sql)
                self.ai5_carrier_list.append(list(result[0]))
            self.ai5_carrier_list.sort(key=itemgetter(1))  # resort carrier list after additions
            
        def ai5_nscode(self):
            self.code_ns = NsDayDict(projvar.invran_date_week[0]).gen_rev_ns_dict()  # generate reverse ns code dict
            
        def ai5_screen(self):
            self.win = MakeWindow()
            self.win.create(self.frame)
            self.ai5_screen_header()
            self.ai5_screen_labels()
            self.ai5_find_discrepancies()
            self.ai5_screen_buttons()
            
        def ai5_screen_header(self):
            header = Frame(self.win.body)
            header.grid(row=0, columnspan=6, sticky="w")
            Label(header, text="Discrepancy Resolution Screen", font=macadj("bold", "Helvetica 18"), pady=10) \
                .grid(row=0, sticky="w")
            Label(header, text="Correct "
            "any discrepancies and inconsistencies that exist between the incoming TACS data (in blue) \n"
            "and the information currently recorded in the Klusterbox database (below in the entry fields and \n"
            "option menus)to reflect the carrier's status accurately. This will update the Klusterbox database. \n"
            "Routes must 4  or 5 digits long. In cases where there are multiple routes, the routes must be \n"
            "separated by a \"/\" backslash.\n\n"
            "Investigation Range: {0} through {1}\n\n"
                  .format(projvar.invran_date_week[0].strftime("%a - %b %d, %Y"),
                          projvar.invran_date_week[6].strftime("%a - %b %d, %Y")), justify=LEFT) \
                .grid(row=1, sticky="w")
            
        def ai5_screen_labels(self):
            if not self.parent.is_mac:  # skip labels if the os is mac
                Label(self.win.body, text="    ", fg="Grey").grid(row=self.y, column=0, sticky="w")
                Label(self.win.body, text=macadj("List Status", "List"), fg="Grey")\
                    .grid(row=self.y, column=1, sticky="w")
                Label(self.win.body, text="NS Day", fg="Grey").grid(row=self.y, column=2, sticky="w")
                Label(self.win.body, text="Route_s", fg="Grey").grid(row=self.y, column=3, sticky="w")
                Label(self.win.body, text="Station", fg="Grey").grid(row=self.y, column=4, sticky="w")
                Label(self.win.body, text=macadj("             ", ""), fg="Grey")\
                    .grid(row=self.y, column=5, sticky="w")
                self.y += 1
                
        def ai5_find_discrepancies(self):  # look for any discrepancies in carrier list
            tlist = ()
            tnsday = "none"
            troute = ""
            for name in self.parent.check_these:
                for k_name in self.ai5_carrier_list:
                    if self.name_dict[name[0]] == k_name[1]:  # if the names match
                        if name[3] == "auxiliary":  # parse assignments from tacs list
                            tlist = self.aux_list_tuple
                            tnsday = "none"
                            troute = ""
                        if name[3] == "part time flex":  # parse assignments from tacs list
                            tlist = self.aux_list_tuple
                            tnsday = "none"
                            troute = ""
                        if name[3][-4:].isnumeric():
                            tlist = self.reg_list_tuple
                            tnsday = self.code_ns[str(self.ns_dict[name[0]])]
                            troute = name[3][-4:]
                        if name[3][-7:] == "floater":
                            tlist = self.reg_list_tuple
                            tnsday = self.code_ns[str(self.ns_dict[name[0]])]
                            troute = "floater"
                        if name[3] == "undetected":
                            tlist = "undetected"
                            tnsday = self.code_ns[str(self.ns_dict[name[0]])]
                            troute = "undetected"
                        discrepancy = False
                        # check tacs data against data in carriers table/ klusterbox
                        if k_name[2] not in tlist:  # check list status
                            discrepancy = True
                        if k_name[3] != tnsday:  # check nsday
                            discrepancy = True
                        k_rte_len = len(k_name[4].split('/'))  # check route
                        if k_rte_len == 0:  # check if route is aux
                            if troute != "":
                                discrepancy = True
                        if k_rte_len == 1:  # check if route is regular
                            if troute != k_name[4]:
                                discrepancy = True
                        if k_rte_len == 5:  # check if route is floater
                            if troute != "floater":
                                discrepancy = True
                        if projvar.invran_station != k_name[5]:  # check if station is correct
                            discrepancy = True
                        if discrepancy:  # if there are no discrepancies, then skip the screen
                            self.skip_this_screen = False
                            self.ai5_display_discrepancies(name, k_name)
        
        def ai5_display_discrepancies(self, name, k_name):
            name_f = Frame(self.win.body)  # create separate frame for names
            name_f.grid(row=self.y, columnspan=6, sticky="w")
            Label(name_f, text="Name: ", fg="Grey").grid(row=0, column=0, sticky="w")
            Label(name_f, text=name[1] + ", " + name[2], fg=self.color).grid(row=0, column=1, sticky="w")
            Label(name_f, text=" / " + k_name[1]).grid(row=0, column=2, sticky="w")
            self.y += 1
            if not self.parent.is_mac:
                Label(self.win.body, text="    ", fg=self.color).grid(row=self.y, column=0, sticky="w")
            Label(self.win.body, text=macadj("not in record", "unknown"), fg=self.color) \
                .grid(row=self.y, column=1, sticky="w")
            Label(self.win.body, text=str(self.ns_dict[name[0]]), fg=self.color).grid(row=self.y, column=2, sticky="w")
            Label(self.win.body, text=name[3], fg=self.color).grid(row=self.y, column=3, sticky="w")
            Label(self.win.body, text=projvar.invran_station, fg=self.color).grid(row=self.y, column=4, sticky="w")
            self.y += 1
            self.carrier_name.append(k_name[1])  # add kb name to the array
            list_options = ("otdl", "wal", "nl", "ptf", "aux")  # create optionmenu for list status
            self.l_s.append(StringVar(self.win.body))
            self.l_s[self.i].set(k_name[2])  # set the list status
            list_status = OptionMenu(self.win.body, self.l_s[self.i], *list_options)
            list_status.config(width=macadj(6, 4))
            list_status.grid(row=self.y, column=1, sticky="w")
            self.l_ns.append(StringVar(self.win.body))  # create optionmenu for ns days
            self.l_ns[self.i].set(self.ns_opt_dict[k_name[3]])  # set ns day default
            ns_day = OptionMenu(self.win.body, self.l_ns[self.i], *self.opt_nsday)
            ns_day.config(width=macadj(12, 8))
            ns_day.grid(row=self.y, column=2, sticky="w")
            self.e_route.append(StringVar(self.win.body))  # create entry field for route
            Entry(self.win.body, width=25, textvariable=self.e_route[self.i]) \
                .grid(row=self.y, column=3, sticky="w")  # create entry for routes
            self.e_route[self.i].set(k_name[4])
            self.l_station.append(StringVar(self.win.body))
            self.l_station[self.i].set(k_name[5])
            list_station = OptionMenu(self.win.body, self.l_station[self.i], *projvar.list_of_stations)
            list_station.config(width=macadj(25, 18))
            list_station.grid(row=self.y, column=4, sticky="w")
            self.y += 1
            Label(self.win.body, text="").grid(row=self.y, column=1)
            self.y += 1
            self.i += 1

        def ai5_screen_buttons(self):
            Button(self.win.buttons, text="Continue", width=macadj(15, 16),
                   command=lambda: self.ai5_apply()).pack(side=LEFT)
            Button(self.win.buttons, text="Cancel", width=macadj(15, 16),
                   command=lambda: self.parent.go_back(self.win.topframe)).pack(side=LEFT)
            if self.skip_this_screen:
                self.parent.AutoIndexer6(self.parent).run(self.win.topframe)
            else:
                self.win.finish()  # get rear window objects

        def ai5_apply(self):  # generate progressbar - sends data to be checked
            eff_date = projvar.invran_date_week[0]  # if investigation range is weekly
            if not projvar.invran_weekly_span:  # if investigation range is daily
                eff_date = projvar.invran_date
            pb_label = Label(self.win.buttons, text="Updating Changes: ")  # make label for progress bar
            pb_label.pack(side=LEFT)
            pb = ttk.Progressbar(self.win.buttons, length=250, mode="determinate")  # create progress bar
            pb.pack(side=LEFT)
            pb["maximum"] = len(self.carrier_name)  # set length of progress bar
            pb.start()
            for i in range(len(self.carrier_name)):
                pb["value"] = i  # increment progress bar
                passed_ns = self.l_ns[i].get().split(" - ")  # clean the passed ns day data
                clean_ns = StringVar(self.win.topframe)  # put ns day var in StringVar object
                clean_ns.set(passed_ns[1])
                if not self.check_and_apply(self.win.topframe, eff_date, self.carrier_name[i],
                                          self.l_s[i], clean_ns, self.e_route[i], self.l_station[i]):
                    frame = self.win.topframe  # prevent the object from being obliterated by rerunning __init__
                    self.__init__(self.parent)  # re initialize the child class
                    self.run(frame)
                    return
                self.win.buttons.update()
            pb.stop()  # stop and destroy the progress bar
            pb_label.destroy()  # destroy the label for the progress bar
            pb.destroy()
            self.parent.AutoIndexer6(self.parent).run(self.win.topframe)

        @staticmethod
        def check_and_apply(frame, date, carrier, ls, ns, route, station):  # adds new carriers to the carriers table
            if len(route.get()) > 29:
                messagebox.showerror("Route number input error",
                                     "There can be no more than five routes per carrier "
                                     "(for T6 carriers).\n Routes numbers must be four or five digits long.\n"
                                     "If there are multiple routes, route numbers must be separated by "
                                     "the \'/\' character. For example: 1001/1015/10124/10224/0972. Do not use "
                                     "commas or empty spaces", parent=frame)
                return False
            route_list = route.get().split("/")
            for item in route_list:
                item = item.strip()
                if item != "":
                    if len(item) < 4 or len(item) > 5:
                        messagebox.showerror("Route number input error",
                                             "Routes numbers must be four or five digits long.\n"
                                             "If there are multiple routes, route numbers must be separated by "
                                             "the \'/\' character. For example: 1001/1015/10124/10224/0972. "
                                             "Do not use commas or empty spaces",
                                             parent=frame)
                        return False
                if item.isdigit() == FALSE and item != "":
                    messagebox.showerror("Route number input error",
                                         "Route numbers must be numbers and can not contain "
                                         "letters",
                                         parent=frame)
                    return False
            route_input = Handler(route.get()).routes_adj()
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
            return True

    class AutoIndexer6:  # detect carriers who are no longer in station
        def __init__(self, parent):
            self.parent = parent
            self.frame = None
            self.names_list = []  # list of carriers in investigation range.
            self.filtered_ids = []  # filter the tacs ids to only good jobs
            self.t_names = []  # matches emp id to the kb name
            self.ex_carrier = []  # carriers in carrier list but not tacs data
            self.win = None
            self.y = 1  # count for the row
            self.carrier_name = []
            self.list_status = []
            self.ns_day = []
            self.route = []
            self.station = []
            self.new_station = []
            self.cc = 0

        def run(self, frame):
            self.frame = frame
            self.ai6_nameslist()  # create the names list array
            self.ai6_filtered_ids()
            self.ai6_t_names()
            self.ai6_ex_carriers()
            if len(self.ex_carrier) == 0:
                self.parent.AutoSkimmer(self.parent).run(self.frame)
            else:
                self.ai6_screen()  # create the 'carriers no longer in station' screen
                self.ai6_screen_header()
                self.ai6_screen_labels()
                self.ai6_screen_loop()
                self.ai6_screen_buttons()
                self.win.finish()

        def ai6_nameslist(self):  # list who are not in the TACS list
            carrier_list = gen_carrier_list()  # create names_list array
            for name in carrier_list:  # eliminate duplicate names
                if name[1] not in self.names_list:
                    self.names_list.append(name[1])

        def ai6_filtered_ids(self):  # filter the tacs ids to get the good jobs
            self.parent.get_file()  # read the csv file
            tacs_ids = []  # generate tacs list
            good_jobs = ("844", "134", "434")
            to_add = ("x", "x")  # create placeholder for
            for line in self.parent.a_file:
                if len(line) > 19:  # if there are enough items in the line
                    if line[18] == "Temp":
                        to_add = (line[4].zfill(8), line[19])
                    elif line[19] != "Temp" or line[19] != "Base":
                        if to_add != ("x", "x"):  # if not placeholder
                            tacs_ids.append(to_add)  # add tacs data to the array
                            to_add = ("x", "x")  # reset placeholder
                    if line[18] == "Base":
                        to_add = (line[4].zfill(8), line[19])
            self.filtered_ids = []  # filter the tacs ids to only good jobs
            for item in tacs_ids:
                if item[1] in good_jobs:
                    self.filtered_ids.append(item)
            del tacs_ids

        def ai6_t_names(self):
            for name in self.filtered_ids:  #
                sql = "SELECT kb_name FROM name_index WHERE emp_id = '%s'" % (name[0])
                result = inquire(sql)  # check dbase for a match
                if result:  # if there is a match in the dbase, then add data to array
                    self.t_names.append(result[0][0])

        def ai6_ex_carriers(self):  # get a list of carriers no longer in the station
            for name in self.names_list:  # for each name in carrier list
                if name not in self.t_names:  # if they are not also in the tacs data
                    self.ex_carrier.append(name)  # then add them to the array

        def ai6_screen(self):
            self.win = MakeWindow()
            self.win.create(self.frame)

        def ai6_screen_header(self):
            header = Frame(self.win.body)
            header.grid(row=0, columnspan=5, sticky="w")
            Label(header, text="Carriers No Longer At Station", font=macadj("bold", "Helvetica 18"), pady=10) \
                .grid(row=0, sticky="w")
            wintext = "Klusterbox has detected that the following carriers may no longer be at the station. " \
                      "If they are no longer at the\n station, then please use the option menu below to move " \
                      "them to the correct station (if listed). If the correct \nis not listed or the carrier " \
                      "is no longer working for the post office, then select \"out of station\".\n\n"
            mactext = \
                "Klusterbox has detected that the following carriers may no longer be at the station. If they \n" \
                "are no longer at the station, then please use the option menu below to move them to the \n" \
                "correct station (if listed). If the correct is not listed or the carrier is no longer working \n" \
                "for the post office, then select \"out of station\".\n\n"
            text = "Investigation Range: {0} through {1}\n\n".format(
                projvar.invran_date_week[0].strftime("%a - %b %d, %Y"),
                projvar.invran_date_week[6].strftime("%a - %b %d, %Y"))
            Label(header, text=macadj(wintext, mactext) + text, justify=LEFT).grid(row=1, sticky="w")

        def ai6_screen_labels(self):
            Label(self.win.body, text="Name", fg="Grey").grid(row=self.y, column=0, sticky="w")
            Label(self.win.body, text=macadj("List Status", "List"), fg="Grey").grid(row=self.y, column=1, sticky="w")
            if sys.platform != "darwin":
                Label(self.win.body, text="Route_s", fg="Grey").grid(row=self.y, column=2, sticky="w")
            Label(self.win.body, text="Station", fg="Grey").grid(row=self.y, column=3, sticky="w")
            Label(self.win.body, text="             ", fg="Grey").grid(row=self.y, column=4, sticky="w")
            self.y += 1

        def ai6_screen_loop(self):
            for name in self.ex_carrier:
                sql = "SELECT * FROM carriers WHERE carrier_name = '%s' and effective_date <= '%s' " \
                      "ORDER BY effective_date DESC" \
                      % (name, projvar.invran_date_week[0])
                result = inquire(sql)
                self.carrier_name.append(StringVar(self.win.body))  # store name
                self.carrier_name[self.cc].set(result[0][1])
                Button(self.win.body, text=result[0][1], relief=RIDGE, width=25, anchor="w") \
                    .grid(row=self.y, column=0, sticky="w")  # name
                self.list_status.append(StringVar(self.win.body))  # store list status
                self.list_status[self.cc].set(result[0][2])
                Button(self.win.body, text=result[0][2], relief=RIDGE, width=7, anchor="w") \
                    .grid(row=self.y, column=1, sticky="w")  # list
                self.ns_day.append(StringVar(self.win.body))  # store ns day
                self.ns_day[self.cc].set(result[0][3])
                self.route.append(StringVar(self.win.body))  # store route
                self.route[self.cc].set(result[0][4])
                if sys.platform != "darwin":
                    Button(self.win.body, text=result[0][4], relief=RIDGE, width=20, anchor="w") \
                        .grid(row=self.y, column=2, sticky="w")  # route
                self.station.append(StringVar(self.win.body))  # store station
                self.station[self.cc].set(result[0][5])
                self.new_station.append(StringVar(self.win.body))
                self.new_station[self.cc].set("out of station")
                stat_om = OptionMenu(self.win.body, self.new_station[self.cc], *projvar.list_of_stations)  # station
                if sys.platform != "darwin":
                    stat_om.config(width=25, anchor="w")
                else:
                    stat_om.config(width=25)
                stat_om.grid(row=self.y, column=3, sticky="w")
                Label(self.win.body, text="                     ").grid(row=self.y, column=4)
                self.cc += 1
                self.y += 1

        def ai6_screen_buttons(self):
            Button(self.win.buttons, text="Continue", width=macadj(15, 16),
                   command=lambda: self.ai6_apply()).pack(side=LEFT)
            Button(self.win.buttons, text="Cancel", width=macadj(15, 16),
                   command=lambda: self.parent.go_back(self.win.topframe)).pack(side=LEFT)

        def ai6_apply(self):
            date = projvar.invran_date_week[0]
            pb_label = Label(self.win.buttons, text="Updating Changes: ")  # make label for progress bar
            pb_label.pack(side=LEFT)
            pb = ttk.Progressbar(self.win.buttons, length=250, mode="determinate")  # create progress bar
            pb.pack(side=LEFT)
            pb["maximum"] = len(self.carrier_name)  # set length of progress bar
            pb.start()
            for i in range(len(self.carrier_name)):
                pb["value"] = i  # increment progress bar
                if self.station[i].get() != self.new_station[i].get():  # if there is a change of station
                    self.parent.AutoIndexer5(self.parent).check_and_apply(self.win.topframe, date,
                    self.carrier_name[i].get(), self.list_status[i], self.ns_day[i], self.route[i], self.new_station[i])
                self.win.buttons.update()
            pb.stop()  # stop and destroy the progress bar
            pb_label.destroy()  # destroy the label for the progress bar
            pb.destroy()
            # self.auto_skimmer(self.win.topframe, self.file_path)
            self.parent.AutoSkimmer(self.parent).run(self.win.topframe)

    class AutoSkimmer:
        """
        This class enters in the clock rings by reading the employee everything report csv. While the above
        classes focused on the Base and Temp lines, this class focus on the lines dealing with hours worked,
        paid leave, unpaid leave, begin tour, moves an end tour.
        """
        def __init__(self, parent):
            self.parent = parent
            self.frame = None
            self.allow_zero_top = None
            self.allow_zero_bottom = None
            self.skippers = None
            self.days = ("Saturday", "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday")
            self.mv_codes = ("BT", "MV", "ET")
            self.day_dict = {}  # make a dictionary for each day in the week
            self.carrier_lines = []
            self.weekly_protoarray = []  # an array of daily_protoarrays
            self.daily_protoarray = []  # returned from daily analysis
            self.newest_carrier = []
            self.kb_name = ""  # carrier name from nameindex table
            self.routes = []  # get route/s
            self.c_code = "none"  # notes ns day hit
            self.mv_triad = []  # triad is route#, start time off route, end time off route
            self.mv_str = ""  # moves
            self.hr_52 = ""  # paid leave
            self.rs = ""  # return to station time
            self.lv_type = ""  # 5200 leave type
            self.lv_time = ""  # 5200 leave time
            self.current_array = []  # product of skim weekly - formatted data to dbase input
            # variables for build_protoarray()
            self.daily_rings = []
            self.daily_line = []
            self.day_name = ""
            self.day_hr_52 = 0.0  # work hours
            self.day_hr_55 = 0.0  # annual leave
            self.day_hr_56 = 0.0  # sick leave
            self.day_hr_58 = 0.0  # holiday leave
            self.day_hr_62 = 0.0  # guaranteed time
            self.day_hr_86 = 0.0  # other paid leave
            self.day_rs = 0
            self.day_code = ""
            self.day_moves = []
            self.day_leave_type = []
            self.day_leave_time = []
            self.day_final_leave_type = ""
            self.day_final_leave_time = 0.0
            self.day_dayofweek = None
            # variables for fix carrier lines
            self.new_order = []

        def run(self, frame):
            self.frame = frame
            self.skim_configs()  # get configuration settings
            carrier_list_cleaning_for_auto_skimmer(self.frame)
            self.skim_day_dict()  # make a dictionary for each day in the week
            if not self.skim_check_csv():  # checks for employee everything report
                self.parent.go_back(self.frame)  # quit and return to main screen
            else:
                if not messagebox.askokcancel("Automatic Rings Entry",
                                          "Do you want to automatically enter the rings?",
                                          parent=self.frame):
                    self.parent.go_back(self.frame)  # quit and return to main screen
                else:
                    self.skim_enter_rings()
                    messagebox.showinfo("Automatic Rings Entry",
                                        "The Employee Everything Report has been sucessfully inputed into the database",
                                        parent=self.frame)
                    self.parent.go_back(self.frame)  # quit and return to main screen

        def skim_configs(self):  # get configuration settings
            sql = "SELECT tolerance FROM tolerances WHERE category='%s'" % "allow_zero_top"
            result = inquire(sql)
            self.allow_zero_top = result[0][0]
            sql = "SELECT tolerance FROM tolerances WHERE category='%s'" % "allow_zero_bottom"
            result = inquire(sql)
            self.allow_zero_bottom = result[0][0]
            sql = "SELECT code FROM skippers"  # get skippers data from dbase
            results = inquire(sql)
            self.skippers = []  # fill the array for skippers
            for item in results:
                self.skippers.append(item[0])

        def skim_day_dict(self):
            x = 0
            for item in self.days:  # make a dictionary for each day in the week
                self.day_dict[item] = projvar.invran_date_week[x]
                x += 1

        def skim_check_csv(self):  # checks for employee everything report
            self.parent.get_file()  # read the csv file
            for line in self.parent.a_file:
                if line[0][:8] == "TAC500R3":
                    return True
                else:
                    messagebox.showwarning("File Selection Error",
                                           "The selected file does not appear to be an "
                                           "Employee Everything report.", parent=self.frame)
                    return False

        def skim_enter_rings(self):
            """
            Takes the entire csv file, goes line by line and breaks it up into one chunk per carrier
            and sends it to the skim weekly for further breakdown by day
            """
            self.parent.get_file()  # read the csv file
            row_count = sum(1 for row in self.parent.a_file)  # get number of rows in csv file
            self.parent.get_file()  # read the csv file
            pb = ProgressBarDe(title="Entering Carrier Rings", label="Updating Rings: ", text="Stand by...")
            pb.max_count(int(row_count))
            pb.start_up()
            i = 0
            cc = 0
            good_id = "no"
            for line in self.parent.a_file:
                pb.move_count(i)
                if cc != 0:
                    if good_id != line[4] and good_id != "no":  # if new carrier_lines or employee
                        self.skim_weekly()  # trigger analysis
                        del self.carrier_lines[:]  # empty array
                        good_id = "no"  # reset trigger
                    # find first line of specific carrier_lines
                    if line[18] == "Base" and line[19] in ("844", "134", "434"):
                        good_id = line[4]  # set trigger to id of carriers who are FT or aux carriers
                        self.carrier_lines.append(line)  # gather times and moves for anaylsis
                        pb.change_text("Entering rings for {}".format(line[5]))
                    if good_id == line[4] and line[18] != "Base":
                        if line[18] in self.days:  # get the hours for each day
                            self.carrier_lines.append(line)  # gather times and moves for anaylsis
                        if line[19] in self.mv_codes and line[32] != "(W)Ring Deleted From PC":
                            self.carrier_lines.append(line)  # gather times and moves for anaylsis
                        pb.change_text("Entering rings for {}".format(line[5]))
                cc += 1
                i += 1
            self.skim_weekly()  # when loop ends, run final analysis
            del self.carrier_lines[:]  # empty array
            pb.stop()  # stop and destroy the progress bar

        def skim_weekly(self):
            """
            Takes the carrier lines sent by enter rings method and sends it to input rings to convert
            it into an array of proto arrays - one for each day collected in input rings
            """
            self.fix_carrierlines()
            del self.weekly_protoarray[:]  # delete prior input rings
            self.skim_input_rings()  # build an array of protoarrays
            if self.weekly_protoarray[0] is not None:
                result = self.skim_check_nameindex()  # get the carriers employee id number
                if result:  # if there is an employee id number in the name index, then continue
                    if self.skim_check_carriers(result):  # get the kb name which correlates to the emp id
                        self.skim_detect_nsday()  # find the ns day for the carrier
                        self.skim_get_routes()  # create an array of the carrier's routes for self.routes
                        for i in range(len(self.weekly_protoarray)):  # loop for each day of carrier information
                            self.daily_protoarray = self.weekly_protoarray[i]
                            """ should be dealing with input rings and not protoarray as input rings is a storage 
                            array for the daily protoarrays"""
                            self.skim_detect_moves()  # find the moves if any
                            if not self.allow_zero_bottom:
                                self.allow_zero_bottom()
                            if not self.allow_zero_top:
                                self.allow_zero_top()
                            self.skim_get_movestring()
                            if self.skim_get_hour52():
                                self.skim_returntostation()
                                self.skim_get_leavetime()
                                self.skim_current_array()
                                self.skim_input_update()

        def fix_carrierlines(self):
            """
            This method solves a problem of lines in the employee everything report being slightly out of
            sequence such as a move coming before a begin tour or a move coming after an end tour. This method
            rearranges the begin tour, moves and end tour lines to make sure begin tour rings come first, moves
            are in the middle and end tours are last. Lines which are not BT, Moves or ET keep their original
            positions.
            """
            self.new_order = []  # carrier lines restructured
            moves_holder = []
            for line in self.carrier_lines:
                if line[19] in self.mv_codes:  # mv_codes is ("BT", "MV", "ET")
                    moves_holder.append(line)  # captures the BT, MV or ET lines
                else:
                    if moves_holder:  # if there are BT, MV or ET lines in the move holder
                        self.fix_carrierline_moves(moves_holder)  # call a method to put them in proper order
                        del moves_holder[:]  # delete the contents of the array
                    self.new_order.append(line)  # non-BT, MV or ET lines go straight to the new order array.
            if moves_holder:  # at the end of the loop check if there are BT, MV or ET lines in the move holder
                self.fix_carrierline_moves(moves_holder)  # call a method to put them in proper order
            self.carrier_lines = self.new_order[:]  # carrier lines is over written with correctly order array

        def fix_carrierline_moves(self, moves_holder):  # puts the BT, MV and ET lines in proper order
            bt_array = []  # holds begin tour lines
            mv_array = []  # hold moves lines
            et_array = []  # holds end tour lines
            for move in moves_holder:  # loop through the BT, MV or ET lines
                if move[19] == "BT":
                    bt_array.append(move)  # capture begin tours in an array
                if move[19] == "MV":
                    mv_array.append(move)  # capture moves in an array
                if move[19] == "ET":
                    et_array.append(move)  # captures end tours in an array
            for move in (bt_array + mv_array + et_array):  # with the lines in the proper order...
                self.new_order.append(move)  # put the BT, MV or ET lines into the new order array

        def skim_input_rings(self):
            """
            Takes The carrier lines from enter rings method and creates a daily protoarrays for the
            investigation range.
            """
            rings = []
            good_day = "no"
            for line in self.carrier_lines:
                if line[18] in self.days and line[18] != good_day and good_day != "no":
                    to_input = self.build_protoarray(rings)  # returns the protoarray for one day
                    self.weekly_protoarray.append(to_input)
                    del rings[:]
                    good_day = line[18]
                if line[18] == "Base" and line[19] in ("844", "134", "434"):  # find first line of specific carrier
                    continue  # gather base line data
                elif line[18] == "Temp" and line[19] in ("844", "134", "434"):  # find first line of specific carrier
                    continue  # gather base line data
                else:
                    if line[18] in self.days and line[18] == good_day:
                        rings.append(line)
                    if line[18] in self.days and good_day == "no":  # day change triggers
                        good_day = line[18]
                        rings.append(line)
                    if line[18] not in self.days:
                        rings.append(line)
            to_input = self.build_protoarray(rings)  # call function for last line  # returns the protoarray for one day
            self.weekly_protoarray.append(to_input)  # add the proto array for an array

        def skim_check_nameindex(self):
            sql = "SELECT kb_name FROM name_index WHERE emp_id = '%s'" % self.weekly_protoarray[0][1]
            result = inquire(sql)  # check to verify that they are in the name index
            return result  # if there is a match in the name index, then continue
                
        def skim_check_carriers(self, result):    
            self.kb_name = result[0][0]  # get the kb name which correlates to the emp id
            for line in self.weekly_protoarray:
                self.daily_protoarray = line
                sql = "SELECT effective_date, carrier_name, list_status, ns_day, route_s FROM" \
                      " carriers WHERE carrier_name = '%s' and effective_date <= '%s' " \
                      "ORDER BY effective_date DESC" % (self.kb_name, self.day_dict[self.daily_protoarray[0]])
                result = inquire(sql)
                for array in result:  # find the most recent carrier record
                    eff_date = datetime.strptime(array[0], '%Y-%m-%d %H:%M:%S')
                    if eff_date <= self.day_dict[self.daily_protoarray[0]]:
                        self.newest_carrier = array
                        break  # stop. we only need the most recent record
                if result:
                    return True
                return False
                
        def skim_detect_nsday(self):
            # find the code, if any  / as of version 4.003 otdl carriers are allowed ns day code
            if self.newest_carrier[2] in ("nl", "wal", "otdl"):
                if self.day_dict[self.daily_protoarray[0]].strftime("%a") == projvar.ns_code[self.newest_carrier[3]] and \
                        float(self.daily_protoarray[2]) > 0:
                    self.c_code = "ns day"
                else:
                    self.c_code = "none"
            elif self.newest_carrier[2] in ("otdl", "ptf", "aux"):
                if self.daily_protoarray[4] == "":
                    self.c_code = "none"  # self.daily_protoarray[4] is the code from proto-array
                else:
                    self.c_code = self.daily_protoarray[4]  # can be sick or annual
            else:
                self.c_code = "none"
            
        def skim_get_routes(self):
            self.routes = []  # create an array for self.routes
            if self.newest_carrier[4] != "":
                self.routes = self.newest_carrier[4].split("/")

        def skim_detect_moves(self):  # find the moves if any
            self.mv_triad = []  # triad is route number, start time off route, end time off route
            route_holder = ""
            if len(self.routes) > 0:  # if the route is in kb
                pair = "closed"  # trigger opens when a move set needs to be closed
                for m in self.daily_protoarray[5]:  # loop through all the rings
                    mv_time = Convert(m[1]).zero_or_hundredths()  # assign move time variable and format
                    if m[3] not in self.routes and pair == "closed":
                        if m[3] == "0000" and m[2] in self.skippers:  # sometimes off route is not off route
                            continue
                        else:
                            route_holder = m[3]  # hold route to put at end of triad
                            self.mv_triad.append(mv_time)  # add start time to second place of triad
                            pair = "open"
                    if m[3] in self.routes and pair == "open":
                        self.mv_triad.append(mv_time)  # add end time to third place of triad
                        self.mv_triad.append(route_holder)
                        pair = "closed"
                if pair == "open":  # if open at end, then close it with the last ring
                    # assign move time variable and format for the last move if pair == 'open'
                    mv_time = Convert(self.daily_protoarray[5][len(self.daily_protoarray[5]) - 1][1]).zero_or_hundredths()
                    self.mv_triad.append(mv_time)
                    self.mv_triad.append(route_holder)
            
        def allow_zero_bottom(self):
            if len(self.mv_triad) > 0:  # find and remove duplicate ET rings at end
                # if the last 2 are the same
                if self.mv_triad[int(len(self.mv_triad) - 3)] == self.mv_triad[int(len(self.mv_triad) - 2)]:
                    self.mv_triad.pop()  # pop out the last triad
                    self.mv_triad.pop()
                    self.mv_triad.pop()
        
        def allow_zero_top(self):
            if len(self.mv_triad) > 0:  # find and remove rings in the front
                if self.mv_triad[0] == self.mv_triad[1]:
                    self.mv_triad.pop(0)  # pop out the triad
                    self.mv_triad.pop(0)
                    self.mv_triad.pop(0)
        
        def skim_get_movestring(self):                
            self.mv_str = ','.join(self.mv_triad)  # format array as string to fit in dbase
            
        def skim_get_hour52(self):  # get paid leave
            # if hours worked > 0 or there is a code or a leave type
            if float(self.daily_protoarray[2]) > 0 or self.c_code != "none" or self.daily_protoarray[6] != "":
                hr_52 = self.daily_protoarray[2]  # assign 5200 hours variable
                if RingTimeChecker(hr_52).check_for_zeros():  # adjust hr_52to version 4 record standards
                    self.hr_52 = ""
                else:
                    self.hr_52 = Convert(hr_52).hundredths()
                return True
            return False
        
        def skim_returntostation(self):            
            rs = self.daily_protoarray[3]  # assign return to station variable
            if RingTimeChecker(rs).check_for_zeros():  # adjust rs to version 4 record standards
                self.rs = ""
            else:
                self.rs = Convert(rs).hundredths()
              
        def skim_get_leavetime(self):
            lv_time = float(self.daily_protoarray[7])  # assign leave time variable
            self.lv_type = Convert(self.daily_protoarray[6]).none_not_empty()  # adjust lv type to version 4 standards
            if RingTimeChecker(lv_time).check_for_zeros():  # adjust lv time to version 4 record standards
                self.lv_time = ""
            else:
                self.lv_time = Convert(lv_time).hundredths()
       
        def skim_current_array(self):         
            self.current_array = [str(self.day_dict[self.daily_protoarray[0]]), self.kb_name, self.hr_52, self.rs,
                                  self.c_code, self.mv_str, self.lv_type, self.lv_time]
            
        def skim_input_update(self):    
            # check rings table to see if record already exist.
            sql = "SELECT * FROM rings3 WHERE carrier_name = '%s' and rings_date = '%s'" % (
                self.kb_name, self.day_dict[self.daily_protoarray[0]])
            result = inquire(sql)
            if len(result) == 0:
                sql = "INSERT INTO rings3 (rings_date, carrier_name, total, " \
                      "rs, code, moves, leave_type,leave_time) " \
                      "VALUES('%s','%s','%s','%s','%s','%s','%s','%s')" % \
                      (self.current_array[0], self.current_array[1], self.current_array[2], self.current_array[3],
                       self.current_array[4],
                       self.current_array[5], self.current_array[6], self.current_array[7])
                commit(sql)
            else:
                sql = "UPDATE rings3 SET total='%s',rs='%s' ,code='%s',moves='%s'," \
                      "leave_type ='%s',leave_time = '%s'" \
                      "WHERE rings_date = '%s' and carrier_name = '%s'" \
                      % (
                          self.current_array[2], self.current_array[3], self.current_array[4], self.current_array[5],
                          self.current_array[6], self.current_array[7],
                          self.current_array[0], self.current_array[1])
                commit(sql)

        def build_protoarray(self, rings):
            self.daily_rings = rings
            self.skim_daily_initialize()  # zero out all daily values for each iteration
            if len(self.daily_rings) > 0:
                self.skim_name()  # get the carrier id from the tacs data
                for line in self.daily_rings:
                    self.daily_line = line
                    # get 5200 or non 5200 times for TOTAL, code, leave_type and leave_time
                    if self.daily_line[18] in self.days:
                        self.skim_dayofweek()  # get the day of the week from the tacs data line as self.day_dayofweek
                        self.skim_get_hours()  # get hours as for the day as self.day_hr_52, etc
                        if float(self.day_hr_55) > 0 or float(self.day_hr_56) > 0 or float(self.day_hr_58) > 0 or \
                                float(self.day_hr_62) > 0 or float(self.day_hr_86) > 0:
                            self.skim_daily_leavetime()  # fill day leave type and time variables
                        self.skim_get_code()  # detects annual or sick leave for day_code variable
                    # get the RETURN TO OFFICE time
                    if self.daily_line[19] == "MV" and self.daily_line[23][:3] == "722":
                        self.skim_get_returntostation()  # get return to station time and fill day_rs variable
                    if self.daily_line[19] in self.mv_codes:  # get the MOVES
                        self.skim_get_moves()  # build an array of moves for the day
                proto_array = [self.day_dayofweek, self.day_name, self.day_hr_52, self.day_rs, self.day_code,
                               self.day_moves, self.day_final_leave_type, self.day_final_leave_time]
                return proto_array  # send it back to auto weekly analysis()

        def skim_daily_initialize(self):  # initialize variables for build_protoarray()
            self.day_hr_52 = 0.0  # work hours
            self.day_hr_55 = 0.0  # annual leave
            self.day_hr_56 = 0.0  # sick leave
            self.day_hr_58 = 0.0  # holiday leave
            self.day_hr_62 = 0.0  # guaranteed time
            self.day_hr_86 = 0.0  # other paid leave
            self.day_rs = 0
            self.day_code = ""
            self.day_moves = []
            self.day_leave_type = []
            self.day_leave_time = []
            self.day_final_leave_type = ""
            self.day_final_leave_time = 0.0
            self.day_dayofweek = None

        def skim_name(self):  # get the carrier id from the tacs data
            self.day_name = self.daily_rings[0][4].zfill(8)  # Get NAME

        def skim_dayofweek(self):  # get the day of the week from the tacs data line
            self.day_dayofweek = self.daily_line[18]

        def skim_get_hours(self):
            spt_20 = self.daily_line[20].split(':')  # split to get code and hours
            # get second and third digits of the of the split line 20 or spt_20
            spt_20_mod = "".join([spt_20[0][1], spt_20[0][2]])
            if spt_20_mod == "52":
                self.day_hr_52 = spt_20[1]  # get the total hours worked
            if spt_20_mod == "55":
                self.day_hr_55 = spt_20[1]  # get the annual leave hours
            if spt_20_mod == "56":
                self.day_hr_56 = spt_20[1]  # get the sick leave hours
            if spt_20_mod == "58":
                self.day_hr_58 = spt_20[1]  # get the holiday leave hours
            if spt_20_mod == "62":
                self.day_hr_62 = spt_20[1]  # get the guaranteed time hours
            if spt_20_mod == "86":
                self.day_hr_86 = spt_20[1]  # get other leave hours

        def skim_daily_leavetime(self):  # fill day leave type and time variables
            if float(self.day_hr_55) > 0:
                self.day_leave_type.append("annual")
                self.day_leave_time.append(self.day_hr_55)
            if float(self.day_hr_56) > 0:
                self.day_leave_type.append("sick")
                self.day_leave_time.append(self.day_hr_56)
            if float(self.day_hr_58) > 0:
                self.day_leave_type.append("holiday")
                self.day_leave_time.append(self.day_hr_58)
            if float(self.day_hr_62) > 0:
                self.day_leave_type.append("guaranteed")
                self.day_leave_time.append(self.day_hr_62)
            if float(self.day_hr_86) > 0:
                self.day_leave_type.append("other")
                self.day_leave_time.append(self.day_hr_86)
            if len(self.day_leave_type) > 1:
                self.day_final_leave_type = "combo"
                self.day_final_leave_time = float(self.day_hr_55) + float(self.day_hr_56) + float(self.day_hr_58) + \
                                            float(self.day_hr_62) + float(self.day_hr_86)
            elif len(self.day_leave_type) == 1:
                self.day_final_leave_type = self.day_leave_type[0]
                self.day_final_leave_time = self.day_leave_time[0]
            else:
                self.day_final_leave_type = ""
                self.day_final_leave_time = 0.0

        def skim_get_code(self):  # detects annual or sick leave for day_code variable
            if float(self.day_hr_55) > 1:
                self.day_code = "annual"  # alter CODE if annual leave was used
            if float(self.day_hr_56) > 1:
                self.day_code = "sick"  # alter code if sick leave was used

        def skim_get_returntostation(self):  # get return to station time and fill day_rs variable
            self.day_rs = self.daily_line[21]  # save the last occurrence.

        def skim_get_moves(self):  # build an array of moves for the day
            route_z = self.daily_line[24].zfill(6)  # because some reports omit leading zeros
            # reformat route to 5 digit format
            route = route_z[1] + route_z[2] + route_z[3] + route_z[4] + route_z[5]  # build 5 digit route number
            route = Handler(route).routes_adj()  # convert to 4 digits if route < 100
            # MV code, time off, time on, route
            mv_data = [self.daily_line[19], self.daily_line[21], self.daily_line[23][:3], route]
            self.day_moves.append(mv_data)


def save_all(frame):
    messagebox.showinfo("For Your Information ",
                        "All data has already been saved. Data is saved to the\n"
                        "database whenever an apply or submit button is pressed.\n"
                        "This button does nothing. :)",
                        parent=frame)


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
                if spt_20[0] == "05200":
                    hr_52 = spt_20[1]
                if spt_20[0] == "05300":
                    hr_53 = spt_20[1]
                if spt_20[0] == "04300":
                    hr_43 = spt_20[1]
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
        # find first line of specific carrier
        if line[18] == "Base" and line[19] == "844" \
                or line[18] == "Base" and line[19] == "134" \
                or line[18] == "Base" and line[19] == "434":
            if line[19] == "844":
                list = "aux"
                route = ""
                ns_day = ""
            elif line[19] == "434":
                list = "ptf"
                route = ""
                ns_day = ""
            else:
                list = "FT"
                ns_day = ee_ns_detect(array)  # call function to find the ns day
                if line[23].zfill(2) == "01":
                    route = line[25].zfill(6)
                    route = route[1] + route[2] + route[4] + route[5]
                    route = Handler(route).routes_adj()
                if line[23].zfill(2) == "02":
                    route = "floater"
            report.write("================================================\n")
            report.write(line[5].lower() + ", " + line[6].lower() + "\n")  # write name
            report.write(list + "\n")
            if list == "FT":
                report.write("route:" + route + "\n")
                if ns_day is None:
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


def ee_skimmer(frame):
    days = ("Saturday", "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday")
    mv_codes = ("BT", "MV", "ET")
    carrier = []
    path = dir_filedialog()
    file_path = filedialog.askopenfilename(initialdir=path, filetypes=[("Excel files", "*.csv *.xls")])
    if file_path[-4:].lower() == ".csv" or file_path[-4:].lower() == ".xls":
        with open(file_path, newline="") as file:
            a_file = csv.reader(file)
            cc = 0
            good_id = "no"
            for line in a_file:
                if cc == 0:
                    if line[0][:8] != "TAC500R3":
                        messagebox.showwarning("File Selection Error",
                                               "The selected file does not appear to be an "
                                               "Employee Everything report.",
                                               parent=frame)
                        return
                if cc == 2:
                    pp = line[0]  # find the pay period
                    filename = "ee_reader" + "_" + pp + ".txt"
                    try:
                        report = open(dir_path('ee_reader') + filename, "w")
                    except (PermissionError, FileNotFoundError):
                        messagebox.showwarning("Report Generator",
                                               "The Employee Everything Report Reader "
                                               "was not generated.",
                                               parent=frame)
                        return
                    report.write("\nEmployee Everything Report Reader\n")
                    report.write(
                        "pay period: " + pp[:-3] + " " + pp[4] + pp[5] + "-" + pp[6] + "\n\n")  # printe pay period
                if cc != 0:
                    if good_id != line[4] and good_id != "no":  # if new carrier or employee
                        ee_analysis(carrier, report)  # trigger analysis
                        del carrier[:]  # empty array
                        good_id = "no"  # reset trigger
                    # find first line of specific carrier
                    if line[18] == "Base" and line[19] in ("844", "134", "434"):
                        good_id = line[4]  # set trigger to id of carriers who are FT or aux carriers
                        carrier.append(line)  # gather times and moves for anaylsis
                    if good_id == line[4] and line[18] != "Base":
                        if line[18] in days:  # get the hours for each day
                            carrier.append(line)  # gather times and moves for anaylsis
                        if line[19] in mv_codes and line[32] != "(W)Ring Deleted From PC":
                            carrier.append(line)  # gather times and moves for anaylsis
                cc += 1
            ee_analysis(carrier, report)  # when loop ends, run final analysis
            del carrier[:]  # empty array
            report.close()
            if sys.platform == "win32":
                os.startfile(dir_path('ee_reader') + filename)
            if sys.platform == "linux":
                subprocess.call(["xdg-open", 'kb_sub/ee_reader/' + filename])
            if sys.platform == "darwin":
                subprocess.call(["open", dir_path('ee_reader') + filename])
    else:
        messagebox.showerror("Report Generator",
                             "The file you have selected is not a .csv or .xls file.\n"
                             "You must select a file with a .csv or .xls extension.",
                             parent=frame)
        return


def pp_by_date(sat_range):  # returns a formatted pay period when given the starting date
    year = sat_range.strftime("%Y")
    pp_end = find_pp(int(year) + 1, "011")  # returns the starting date of the pp when given year and pay period
    if sat_range >= pp_end:
        year = int(year) + 1
        year = str(year)
    firstday = find_pp(int(year), "011")  # returns the starting date of the pp when given year and pay period
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


def max_hr(frame):  # generates a report for 12/60 hour violations
    path = dir_filedialog()
    file_path = filedialog.askopenfilename(initialdir=path, filetypes=[("Excel files", "*.csv *.xls")])
    day_xlr = {"Saturday": "sat", "Sunday": "sun", "Monday": "mon", "Tuesday": "tue", "Wednesday": "wed",
               "Thursday": "thr", "Friday": "fri"}
    leave_xlr = {"49": "owcp   ", "55": "annual ", "56": "sick   ", "58": "holiday", "59": "lwop   ", "60": "lwop   "}
    maxhour = []
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
            cc = 0
            good_id = "no"
            for line in a_file:
                if cc == 0:
                    if line[0][:8] != "TAC500R3":
                        messagebox.showwarning("File Selection Error",
                                               "The selected file does not appear to be an "
                                               "Employee Everything report.",
                                               parent=frame)
                        return
                if cc == 2:  # on the second line
                    pp = line[0]  # find the pay period
                    pp = pp.strip()  # strip whitespace out of pay period information
                if cc != 0:  # on all but the first line
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
                            maxhour.append(add_maxhr)
                            for item in extra_hours:  # get any extra hours codes for non-5200 hours list
                                all_extra.append(item)
                            # find the all adjustments
                            if ft:
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
                    # find first line of specific carrier
                    if line[18] == "Base" and line[19] in ("844", "134", "434"):
                        good_id = line[4]  # remember id of carriers who are FT or aux carriers
                        if line[19] in ("844", "434"):
                            ft = False
                        else:
                            ft = True
                    if good_id == line[4] and line[18] != "Base":
                        if line[18] in days:  # get the hours for each day
                            spt_20 = line[20].split(':')  # split to get code and hours
                            hr_type = spt_20[0][1] + spt_20[0][2]  # parse hour code to 2 digits
                            # if hr_type in hr_codes:  # compare to array of hour codes
                            if hr_type == "52":  # compare to array of hour codes
                                if float(spt_20[1]) > 11.5 and not ft:
                                    add_max_aux = (line[5].lower(), line[6].lower(), line[18], spt_20[1])
                                    max_aux_day.append(add_max_aux)
                                if float(spt_20[1]) > 12 and ft:
                                    add_max_ft = (line[5].lower(), line[6].lower(), line[18], spt_20[1])
                                    max_ft_day.append(add_max_ft)
                                if ft:  # increment daily totals to find weekly total
                                    add_day_hours = (line[5].lower(), line[6].lower(), spt_20[1], line[18])
                                    day_hours.append(add_day_hours)
                            extra_hour_codes = ("49", "55", "56", "58")  # paid leave types only , (lwop "59", "60")
                            if hr_type in extra_hour_codes and ft:  # if there is holiday pay
                                add_day_hours = (line[5].lower(), line[6].lower(), spt_20[1], line[18])
                                day_hours.append(add_day_hours)
                                add_extra_hours = (line[5].lower(), line[6].lower(), line[18], hr_type, spt_20[1])
                                extra_hours.append(add_extra_hours)  # track non 5200 hours
                cc += 1
    elif file_path == "":
        return
    else:
        messagebox.showerror("Report Generator",
                             "The file you have selected is not a .csv or .xls file.\n"
                             "You must select a file with a .csv or .xls extension.",
                             parent=frame)
        return
    # find the weekly total by adding daily totals for last carrier
    if len(day_hours) > 0:
        wkly_total = 0
        for t in day_hours:
            wkly_total += float(t[2])
        if wkly_total > 60:
            add_maxhr = (day_hours[0][0].lower(), day_hours[0][1].lower(), wkly_total)
            maxhour.append(add_maxhr)
            for item in extra_hours:  # get any extra hours codes for non-5200 hours list
                all_extra.append(item)
        del day_hours[:]
        del extra_hours[:]

    if len(maxhour) == 0 and len(max_ft_day) == 0 and len(max_aux_day) == 0:
        messagebox.showwarning("Report Generator",
                               "No violations were found. "
                               "The report was not generated.",
                               parent=frame)
        return
    weekly_max = []  # array hold each carrier's hours for the week
    daily_max = []  # array hold each carrier's sum of maximum daily hours for the week
    if len(maxhour) > 0 or len(max_ft_day) > 0 or len(max_aux_day) > 0:
        pp_str = pp[:-3] + "_" + pp[4] + pp[5] + "_" + pp[6]
        filename = "max" + "_" + pp_str + ".txt"
        report = open(dir_path('over_max') + filename, "w")
        report.write("12 and 60 Hour Violations Report\n\n")
        report.write("pay period: " + pp[:-3] + " " + pp[4] + pp[5] + "-" + pp[6] + "\n")  # printe pay period
        pp_date = find_pp(int(pp[:-3]), pp[-3:])  # send year and pp to get the date
        pp_date_end = pp_date + timedelta(days=6)  # add six days to get the last part of the range
        report.write(
            "week of: " + pp_date.strftime("%x") + " - " + pp_date_end.strftime("%x") + "\n")  # printe date
        report.write("\n60 hour violations \n\n")
        report.write("name                              total   over\n")
        report.write("-----------------------------------------------\n")
        if len(maxhour) == 0:
            report.write("no violations" + "\n")
        else:
            diff_total = 0
            maxhour.sort(key=itemgetter(0))
            for item in maxhour:
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
        if len(all_extra) == 0:
            report.write("no contributions" + "\n")
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
        if len(max_ft_day) == 0:
            report.write("no violations" + "\n")
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
        if len(max_aux_day) == 0:
            report.write("no violations" + "\n")
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
        if len(adjustment) == 0:
            report.write("no adjustments" + "\n")
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
                if w_max[0] + w_max[1] == d_max[0] + d_max[1]:  # look for names with both weekly and daily violations
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
        try:
            if sys.platform == "win32":
                os.startfile(dir_path('over_max') + filename)
            if sys.platform == "linux":
                subprocess.call(["xdg-open", 'kb_sub/over_max/' + filename])
            if sys.platform == "darwin":
                subprocess.call(["open", dir_path('over_max') + filename])
        except PermissionError:
            messagebox.showerror("Report Generator",
                                 "The report was not generated.",
                                 parent=frame)


def file_dialogue(folder):  # opens file folders to access generated reports
    if not os.path.isdir(folder):
        os.makedirs(folder)
    if projvar.platform == "py":
        file_path = filedialog.askopenfilename(initialdir=os.getcwd() + "/" + folder)
    else:
        file_path = filedialog.askopenfilename(initialdir=folder)
    if file_path:
        if sys.platform == "win32":
            os.startfile(file_path)
        if sys.platform == "linux":
            subprocess.call(["xdg-open", file_path])
        if sys.platform == "darwin":
            subprocess.call(["open", file_path])


def remove_file(folder):  # removes a file and all contents
    if os.path.isdir(folder):
        shutil.rmtree(folder)


def remove_file_var(frame, folder):  # removes a file and all contents
    if sys.platform == "win32":
        folder_name = folder.split("\\")
    else:
        folder_name = folder.split("/")
    folder_name = folder_name[-2]
    if os.path.isdir(folder):
        if messagebox.askokcancel("Delete Folder Contents",
                                  "This will delete all the files in the {} archive. "
                                          .format(folder_name),
                                  parent=frame):
            try:
                shutil.rmtree(folder)
                if not os.path.isdir(folder):
                    messagebox.showinfo("Delete Folder Contents",
                                        "Success! All the files in the {} archive have been deleted."
                                        .format(folder_name),
                                        parent=frame)
            except PermissionError:
                messagebox.showerror("Delete Folder Contents",
                                     "Failure! {} can not be deleted because it is being used by another program."
                                     .format(folder_name),
                                     parent=frame)
    else:
        messagebox.showwarning("Delete Folder Contents",
                               "The {} folder is already empty".format(folder_name),
                               parent=frame)


class AboutKlusterbox:
    def __init__(self):
        self.win = None
        self.frame = None
        self.photo = None
        
    def start(self, frame):
        self.frame = frame
        self.win = MakeWindow()
        self.win.create(self.frame)
        self.build()
        self.button_frame()
        self.win.finish()
        
    def build(self):
        r = 0  # set row counter
        if projvar.platform == "macapp":
            path = os.path.join(os.path.sep, 'Applications', 'klusterbox.app', 'Contents', 'Resources', 'kb_about.jpg')
        elif projvar.platform == "winapp":
            path = os.path.join(os.path.sep, os.getcwd(), 'kb_about.jpg')
        else:
            path = os.path.join(os.path.sep, os.getcwd(), 'kb_sub', 'kb_images', 'kb_about.jpg')
        try:
            self.photo = ImageTk.PhotoImage(Image.open(path))
            Label(self.win.body, image=self.photo).grid(row=r, column=0, columnspan=10, sticky="w")
        except (TclError, FileNotFoundError):
            pass
        r += 1
        Label(self.win.body, text="").grid(row=r)
        r += 1
        Label(self.win.body, text="Klusterbox", font=macadj("bold", "Helvetica 18"), fg="red", anchor=W) \
            .grid(row=r, column=0, sticky="w", columnspan=6)
        r += 1
        Label(self.win.body, text="").grid(row=r)
        r += 1
        Label(self.win.body, text="version: {}".format(version), anchor=W)\
            .grid(row=r, column=0, sticky="w", columnspan=6)
        r += 1
        Label(self.win.body, text="release date: {}".format(release_date), anchor=W)\
            .grid(row=r, column=0, sticky="w", columnspan=6)
        r += 1
        Label(self.win.body, text="created by Thomas Weeks", anchor=W).grid(row=r, column=0, sticky="w", columnspan=6)
        r += 1
        Label(self.win.body, text="Original release: October 2018", anchor=W)\
            .grid(row=r, column=0, sticky="w", columnspan=6)
        r += 1
        Label(self.win.body, text=" ", anchor=W).grid(row=r, column=0, sticky="w", columnspan=6)
        r += 1
        Label(self.win.body, text="comments and criticisms are welcome", anchor=W, fg="red") \
            .grid(row=r, column=0, sticky="w", columnspan=6)
        r += 1
        Label(self.win.body, text=" ", anchor=W).grid(row=r, column=0, sticky="w", columnspan=6)
        r += 1
        Label(self.win.body, text="contact information: ", anchor=W).grid(row=r, column=0, sticky="w", columnspan=6)
        r += 1
        Label(self.win.body, text="Thomas Weeks", anchor=W).grid(row=r, column=0, sticky="w", columnspan=6)
        r += 1
        Label(self.win.body, text="    tomandsusan4ever@msn.com", anchor=W)\
            .grid(row=r, column=0, sticky="w", columnspan=6)
        r += 1
        Label(self.win.body, text="    (please put \"klusterbox\" in the subject line)", anchor=W) \
            .grid(row=r, column=0, sticky="w", columnspan=6)
        r += 1
        Label(self.win.body, text="I've found that some emails get filtered out by the junk folder so", anchor=W) \
            .grid(row=r, column=0, sticky="w", columnspan=6)
        r += 1
        Label(self.win.body, text="Message me on Facebook Messenger:", anchor=W) \
            .grid(row=r, column=0, sticky="w", columnspan=6)
        r += 1
        kb_link = Label(self.win.body, text="    facebook.com/thomas.weeks.artist", fg="blue", cursor="hand2")
        kb_link.grid(row=r, columnspan=6, sticky="w")
        kb_link.bind("<Button-1>", lambda e: self.callback("http://www.facebook.com/thomas.weeks.artist"))
        r += 1
        Label(self.win.body, text="    720.280.0415", anchor=W).grid(row=r, column=0, sticky="w", columnspan=6)
        r += 1
        Label(self.win.body, text="").grid(row=r)
        r += 1
        Label(self.win.body, text="For the lastest updates on Klusterbox check out the official Klusterbox") \
            .grid(row=r, columnspan=6, sticky="w")
        r += 1
        Label(self.win.body, text="website at:").grid(row=r, columnspan=6, sticky="w")
        r += 1
        kb_link = Label(self.win.body, text="    www.klusterbox.com", fg="blue", cursor="hand2")
        kb_link.grid(row=r, columnspan=6, sticky="w")
        kb_link.bind("<Button-1>", lambda e: self.callback("http://klusterbox.com"))
        r += 1
        Label(self.win.body, text="Also look on Facebook for Klusterbox - Software for NALC Stewards at:") \
            .grid(row=r, columnspan=6, sticky="w")
        r += 1
        fb_link = Label(self.win.body, text="    www.facebook.com/klusterbox", fg="blue", cursor="hand2")
        fb_link.grid(row=r, columnspan=6, sticky="w")
        fb_link.bind("<Button-1>", lambda e: self.callback("http://www.facebook.com/klusterbox"))
        r += 1
        Label(self.win.body, text="Like, Follow and Share!").grid(row=r, columnspan=6, sticky="w")
        r += 1
        Label(self.win.body, text="").grid(row=r)
        r += 1
        Label(self.win.body, text="Project Documentation", font=macadj("bold", "Helvetica 16"), anchor=W) \
            .grid(row=r, column=0, sticky="w", columnspan=3)
        Label(self.win.body, text="                                             ").grid(row=r, column=3)
        Label(self.win.body, text="                                             ").grid(row=r, column=4)
        r += 1
        Label(self.win.body, text="").grid(row=r)
        r += 1
        Button(self.win.body, text="read", width=macadj(7, 7), command=lambda: self.open_docs("readme.txt")) \
            .grid(row=r, column=0, sticky="w")
        Label(self.win.body, text="Read Me", anchor=E).grid(row=r, column=1, sticky="w")
        r += 1
        Label(self.win.body, text="").grid(row=r)
        r += 1
        Button(self.win.body, text="read", width=macadj(7, 7), command=lambda: self.open_docs("history.txt")) \
            .grid(row=r, column=0, sticky="w")
        Label(self.win.body, text="History", anchor=E).grid(row=r, column=1, sticky="w")
        r += 1
        Label(self.win.body, text="").grid(row=r)
        r += 1
        Button(self.win.body, text="read", width=macadj(7, 7), command=lambda: self.open_docs("LICENSE.txt")) \
            .grid(row=r, column=0, sticky="w")
        Label(self.win.body, text="License", anchor=E).grid(row=r, column=1, sticky="w")
        r += 1
        Label(self.win.body, text="").grid(row=r)
        r += 1
        """
        Enter all modules imported by klusterbox below as part of the sourcecode tuple. All modules must be in the 
        klusterbox project folder.
        """
        sourcecode = ("klusterbox.py",
                      "projvar.py",
                      "kbtoolbox.py",
                      "kbdatabase.py",
                      "kbreports.py",
                      "kbspreadsheets.py",
                      "kbspeedsheets.py",
                      "kbequitability.py",
                      )
        for i in range(len(sourcecode)):
            Button(self.win.body, text="read", width=macadj(7, 7),
                   command=lambda source=sourcecode[i]: self.open_docs(source)).grid(row=r, column=0, sticky="w")
            Label(self.win.body, text="Source Code - {}".format(sourcecode[i]), anchor=E)\
                .grid(row=r, column=1, sticky="w")
            r += 1
            Label(self.win.body, text="").grid(row=r)
            r += 1
        Button(self.win.body, text="read", width=macadj(7, 7), command=lambda: self.open_docs("requirements.txt")) \
            .grid(row=r, column=0, sticky="w")
        Label(self.win.body, text="python requirements", anchor=E).grid(row=r, column=1, sticky="w")

    def button_frame(self):
        button = Button(self.win.buttons)
        button.config(text="Go Back", width=20, command=lambda: MainFrame().start(frame=self.win.topframe))
        if sys.platform == "win32":
            button.config(anchor="w")
        button.pack(side=LEFT)

    def open_docs(self, doc):  # opens docs in the about_klusterbox() function
        try:
            if sys.platform == "win32":
                if projvar.platform == "py":
                    try:
                        path = doc
                        os.startfile(path)  # in IDE the files are in the project folder
                    except FileNotFoundError:
                        path = os.path.join(os.path.sep, os.getcwd(), 'kb_sub', doc)
                        os.startfile(path)  # in KB legacy the files are in the kb_sub folder
                if projvar.platform == "winapp":
                    path = os.path.join(os.path.sep, os.getcwd(), doc)
                    os.startfile(path)
            if sys.platform == "linux":
                subprocess.call(doc)
            if sys.platform == "darwin":
                if projvar.platform == "macapp":
                    path = os.path.join(os.path.sep, 'Applications', 'klusterbox.app', 'Contents', 'Resources', doc)
                    subprocess.call(["open", path])
                if projvar.platform == "py":
                    subprocess.call(["open", doc])
        except FileNotFoundError:
            messagebox.showerror("Project Documents",
                                 "The document was not opened or found.",
                                 parent=self.win.body)

    @staticmethod
    def callback(url):  # open hyperlinks at about_klusterbox()
        webbrowser.open_new(url)


class StartUp:
    def __init__(self):
        self.win = None
        self.new_station = None

    def start(self):
        self.win = MakeWindow()
        self.win.create(None)
        self.new_station = StringVar(self.win.body)
        self.build()
        self.win.fill(7, 20)
        self.buttons_frame()
        self.win.finish()

    def build(self):
        Label(self.win.body, text="Welcome to Klusterbox", font=macadj("bold", "Helvetica 18")) \
            .grid(row=0, columnspan=2, sticky="w")
        Label(self.win.body, text="version: {}".format(version)).grid(row=1, columnspan=2, sticky="w")
        Label(self.win.body, text="", pady=20).grid(row=2, column=0)
        # enter new stations
        Label(self.win.body, text="To get started, please enter your station name:", pady=5) \
            .grid(row=3, columnspan=2, sticky="w")
        e = Entry(self.win.body, width=35, textvariable=self.new_station)
        e.grid(row=4, column=0, sticky="w")
        self.new_station.set("")
        Button(self.win.body, width=5, anchor="w", text="ENTER",
               command=lambda: self.apply_startup()).grid(row=4, column=1, sticky="w")
        Label(self.win.body, text="", pady=20).grid(row=5, columnspan=2, sticky="w")
        Label(self.win.body, text="Or you can exit to the main screen and enter your\n"
                       "station by going to Management > list of stations.").grid(row=6, columnspan=2, sticky="w")
        Button(self.win.body, width=5, text="EXIT",
               command=lambda: MainFrame().start(frame=self.win.topframe)).grid(row=7, columnspan=2, sticky="e")

    def buttons_frame(self):
        Label(self.win.buttons, text="").pack()

    def apply_startup(self):
        if not self.new_station.get().strip():
            messagebox.showerror("Prohibited Action",
                                 "You can not enter a blank entry for a station.",
                                 parent=self.win.body)
            return
        sql = "INSERT INTO stations (station) VALUES('%s')" % (self.new_station.get().strip())
        commit(sql)
        projvar.list_of_stations.append(self.new_station.get().strip())
        # access list of stations from database
        sql = "SELECT * FROM stations ORDER BY station"
        results = inquire(sql)
        # define and populate list of stations variable
        del projvar.list_of_stations[:]
        for stat in results:
            projvar.list_of_stations.append(stat[0])
        MainFrame().start(frame=self.win.topframe)  # load new frame


def carrier_list_cleaning_for_auto_skimmer(frame):  # cleans the database of duplicate records
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
        titlebar_icon(pb_root)  # place icon in titlebar
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
        messagebox.showinfo("Database Maintenance",
                            "All redundancies have been eliminated from the carrier list.",
                            parent=frame)
    del duplicates[:]


def carrier_list_cleaning(frame):  # cleans the database of duplicate records
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
        ok = messagebox.askokcancel("Database Maintenance",
                                    "Did you want to eliminate database redundancies? \n"
                                    "{} redundancies have been found in the database \n"
                                    "This is recommended maintenance.".format(len(duplicates)),
                                    parent=frame)
    if ok:
        pb_root = Tk()  # create a window for the progress bar
        pb_root.title("Database Maintenance")
        titlebar_icon(pb_root)  # place icon in titlebar
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
        messagebox.showinfo("Database Maintenance",
                            "All redundancies have been eliminated from the carrier list.",
                            parent=frame)
        MainFrame().start(frame=frame)
    if not ok:
        messagebox.showinfo("Database Maintenance",
                            "No redundancies have been found in the carrier list.",
                            parent=frame)
    del duplicates[:]


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
        messagebox.showerror("Data Entry Error",
                             "It is prohibited to exclude code {}"
                             .format(code.get(),
                                     parent=frame))
        return
    if code.get() in existing_codes:
        messagebox.showerror("Data Entry Error",
                             "This code had already been entered.",
                             parent=frame)
        return
    if code.get().isdigit() == FALSE:
        messagebox.showerror("Data Entry Error",
                             "TACS code must contain only numbers.",
                             parent=frame)
        return
    if len(code.get()) > 3 or len(code.get()) < 3:
        messagebox.showerror("Data Entry Error",
                             "TACS code must be 3 digits long.",
                             parent=frame)
        return
    if len(description.get()) > 39:
        messagebox.showerror("Data Enty Error",
                             "Please limit description to less than 40 characters.",
                             parent=frame)
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
    messagebox.showinfo("Settings Updated",
                        "Auto Data Entry settings have been updated.",
                        parent=frame)


def data_entry_permit_zero(frame, top, bottom):
    sql = "UPDATE tolerances SET tolerance='%s'WHERE category='%s'" % (top.get(), "allow_zero_top")
    commit(sql)
    sql = "UPDATE tolerances SET tolerance='%s'WHERE category='%s'" % (bottom.get(), "allow_zero_bottom")
    commit(sql)
    messagebox.showinfo("Settings Updated",
                        "Auto Data Entry settings have been updated.",
                        parent=frame)


def auto_data_entry_settings(frame):
    wd = front_window(frame)  # F,S,C,FF,buttons
    r = 0
    Label(wd[3], text="Auto Data Entry Settings", font=macadj("bold", "Helvetica 18")) \
        .grid(row=r, column=0, sticky="w", columnspan=4)
    r += 1
    Label(wd[3], text="").grid(row=r, column=1)
    r += 1
    Label(wd[3], text="NS Day Structure Preference", font=macadj("bold", "Helvetica 18")) \
        .grid(row=r, column=0, columnspan=4, sticky="w")
    r += 1
    ns_structure = StringVar(wd[3])
    sql = "SELECT tolerance FROM tolerances WHERE category='%s'" % "ns_auto_pref"
    result = inquire(sql)
    Radiobutton(wd[3], text="rotation", variable=ns_structure, value="rotation") \
        .grid(row=r, column=1, sticky="e")
    Radiobutton(wd[3], text="fixed", variable=ns_structure, value="fixed") \
        .grid(row=r, column=2, sticky="w")
    ns_structure.set(result[0][0])
    r += 1
    Button(wd[3], text="Set", width=5, command=lambda: apply_auto_ns_structure(wd[0], ns_structure)) \
        .grid(row=r, column=3)
    r += 1
    Label(wd[3], text="List of TACS MODS Codes", font=macadj("bold", "Helvetica 18")) \
        .grid(row=r, column=0, columnspan=4, sticky="w")
    r += 1
    Label(wd[3], text="(to exclude from Auto Data Entry moves).") \
        .grid(row=r, column=0, columnspan=4, sticky="w")
    r += 1
    Label(wd[3], text="code", fg="grey", anchor="w") \
        .grid(row=r, column=0)
    Label(wd[3], text="description", fg="grey", anchor="w") \
        .grid(row=r, column=1, columnspan=2)
    sql = "SELECT * FROM skippers"
    results = inquire(sql)
    r += 1
    if len(results) > 0:
        for i in range(len(results)):
            Button(wd[3], text=results[i][0], anchor="w", width=5) \
                .grid(row=i + r, column=0)  # display code
            Button(wd[3], text=results[i][1], anchor="w", width=30) \
                .grid(row=i + r, column=1, columnspan=2)  # display description
            Button(wd[3], text="delete", command=lambda x=i: data_mods_codes_delete(wd[0], results[x])) \
                .grid(row=i + r, column=3)
    else:
        Label(wd[3], text="No Exceptions Listed.", anchor="w") \
            .grid(row=r, column=0, sticky="w", columnspan=3)
        i = 1
    r = r + i
    r += 1
    Label(wd[3], text="").grid(row=r, column=2)
    r += 1
    Label(wd[3], text="Add New Code", font=macadj("bold", "Helvetica 18")) \
        .grid(row=r, column=0, columnspan=3, sticky="w")  # add new code labels
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
    Label(wd[3], text="Permit Zero Sums", font=macadj("bold", "Helvetica 18")) \
        .grid(row=r, column=0, columnspan=2, sticky="w")
    text = "Selecting 'allow' will permit entries into moves where the MOVE OFF and MOVE ON " \
           "times are the same. While these entries do not add to the total for Overtime Worked " \
           "Off route, they might indicate something that would merit further investigation. " \
           "You can always delete them manually. Selecting 'don't allow' will hide these entries." \
           "\n'Top' refers to the start of the workday and 'Bottom' refers to the end of the workday."
    Button(wd[3], text="info", width=5,
           command=lambda: messagebox.showinfo("For Your Information",
                                               text,
                                               parent=wd[0])) \
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

    Button(wd[4], text="Go Back", width=20, command=lambda: (MainFrame().start(frame=wd[0]))) \
        .grid(row=0, column=0, sticky="w")
    rear_window(wd)


class SpreadsheetConfig:
    def __init__(self):
        self.frame = None
        self.win = None
        self.minrows_limit = 100  # hardcoded limit of min rows
        self.min_nl = 0.0
        self.min_wal = 0.0
        self.min_otdl = 0.0
        self.min_aux = 0.0
        self.min_overmax = 0.0
        self.pb_nl_wal = True  # page break between no list and work assignment
        self.pb_wal_otdl = True  # page break between work assignment and otdl
        self.pb_otdl_aux = True  # page break between otdl and auxiliary
        self.min_ot_equit = None  # minimum rows for ot equitability spreadsheet
        self.ot_calc_pref = None  # overtime calcuations preference for otdl equitability
        self.min_ot_dist = None  # minimum rows for ot distribution spreadsheet
        self.ot_calc_pref_dist = None  # overtime calcuations preference for otdl distribution
        self.min_nl_var = None
        self.min_wal_var = None
        self.min_otdl_var = None
        self.min_aux_var = None
        self.min_overmax_var = None
        self.pb_nl_wal_var = None  # page break between no list and work assignment
        self.pb_wal_otdl_var = None  # page break between work assignment and otdl
        self.pb_otdl_aux_var = None  # page break between otdl and auxiliary
        self.min_ot_equit_var = None  # minimum rows for ot equitability spreadsheet
        self.ot_calc_pref_var = None  # overtime calcuations preference for otdl equitability
        self.min_ot_dist_var = None  # minimum rows for ot distribution spreadsheet
        self.ot_calc_pref_dist_var = None  # overtime calcuations preference for otdl distribution
        self.status_update = None  # Label(self.win.buttons, text="", fg="red")
        self.report_counter = 0
        self.check_i = 0  # the iteration of the apply/check method
        self.add_min_nl = 0.0  # prep values to be entered into database
        self.add_min_wal = 0.0
        self.add_min_otdl = 0.0
        self.add_min_aux = 0.0
        self.add_min_overmax = 0.0
        self.add_pb_nl_wal = True  # page break between no list and work assignment
        self.add_pb_wal_otdl = True  # page break between work assignment and otdl
        self.add_pb_otdl_aux = True  # page break between otdl and auxiliary
        self.add_min_ot_equit = None
        self.add_ot_calc_pref = None
        self.add_min_ot_dist = None  # minimum rows for ot distribution spreadsheet
        self.add_ot_calc_pref_dist = None  # overtime calcuations preference for otdl distribution

    def start(self, frame):
        self.frame = frame
        self.win = MakeWindow()
        self.win.create(self.frame)
        self.get_settings()
        self.build_stringvars()
        self.set_stringvars()
        self.build()
        self.buttons_frame()
        self.win.finish()

    def get_settings(self):
        sql = "SELECT tolerance FROM tolerances"
        results = inquire(sql)  # get spreadsheet settings from database
        self.min_nl = results[3][0]
        self.min_wal = results[4][0]
        self.min_otdl = results[5][0]
        self.min_aux = results[6][0]
        self.min_overmax = results[14][0]
        self.pb_nl_wal = results[21][0]  # page break between no list and work assignment
        self.pb_wal_otdl = results[22][0]  # page break between work assignment and otdl
        self.pb_otdl_aux = results[23][0]  # page break between otdl and auxiliary
        # convert bool to "on" or "off"
        self.pb_nl_wal = Convert(self.pb_nl_wal).strbool_to_onoff()
        self.pb_wal_otdl = Convert(self.pb_wal_otdl).strbool_to_onoff()
        self.pb_otdl_aux = Convert(self.pb_otdl_aux).strbool_to_onoff()
        # otdl equitability vars
        self.min_ot_equit = results[25][0]  # minimum rows
        self.ot_calc_pref = results[26][0]  # ot calculation preference
        # overtime distribution vars
        self.min_ot_dist = results[27][0]  # minimum rows
        self.ot_calc_pref_dist = results[28][0]  # ot calculations preference

    def build_stringvars(self):    # create stringvars
        self.min_nl_var = StringVar(self.win.body)
        self.min_wal_var = StringVar(self.win.body)
        self.min_otdl_var = StringVar(self.win.body)
        self.min_aux_var = StringVar(self.win.body)
        self.min_overmax_var = StringVar(self.win.body)
        self.pb_nl_wal_var = StringVar(self.win.body)
        self.pb_wal_otdl_var = StringVar(self.win.body)
        self.pb_otdl_aux_var = StringVar(self.win.body)
        self.min_ot_equit_var = StringVar(self.win.body)
        self.ot_calc_pref_var = StringVar(self.win.body)
        self.min_ot_dist_var = StringVar(self.win.body)
        self.ot_calc_pref_dist_var = StringVar(self.win.body)

    def set_stringvars(self):  # set stringvar values
        self.min_nl_var.set(self.min_nl)
        self.min_wal_var.set(self.min_wal)
        self.min_otdl_var.set(self.min_otdl)
        self.min_aux_var.set(self.min_aux)
        self.min_overmax_var.set(self.min_overmax)
        self.pb_nl_wal_var.set(self.pb_nl_wal)
        self.pb_wal_otdl_var.set(self.pb_wal_otdl)
        self.pb_otdl_aux_var.set(self.pb_otdl_aux)
        self.min_ot_equit_var.set(self.min_ot_equit)
        self.ot_calc_pref_var.set(self.ot_calc_pref)
        self.min_ot_dist_var.set(self.min_ot_dist)
        self.ot_calc_pref_dist_var.set(self.ot_calc_pref_dist)

    def build(self):
        row = 0
        Label(self.win.body, text="Improper Mandate Spreadsheet Configuration", 
              font=macadj("bold", "Helvetica 18"), anchor="w").grid(row=row, sticky="w", columnspan=4)
        row += 1
        Label(self.win.body, text="").grid(row=row, column=0)
        row += 1
        Label(self.win.body, text="Minimum rows for No List Carriers", width=30, anchor="w") \
            .grid(row=row, column=0, ipady=5, sticky="w")
        Entry(self.win.body, width=5, textvariable=self.min_nl_var).grid(row=row, column=1, padx=4, sticky="e")
        Button(self.win.body, width=5, text="info",
               command=lambda: Messenger(self.win.topframe).tolerance_info("min_nl")).grid(row=row, column=2, padx=4)
        row += 1
        Label(self.win.body, text="Minimum rows for Work Assignment", width=30, anchor="w") \
            .grid(row=row, column=0, ipady=5, sticky="w")
        Entry(self.win.body, width=5, textvariable=self.min_wal_var).grid(row=row, column=1, padx=4, sticky="e")
        Button(self.win.body, width=5, text="info",
               command=lambda: Messenger(self.win.topframe).tolerance_info("min_wal")).grid(row=row, column=2, padx=4)
        row += 1
        Label(self.win.body, text="Minimum rows for OT Desired", width=30, anchor="w") \
            .grid(row=row, column=0, ipady=5, sticky="w")
        Entry(self.win.body, width=5, textvariable=self.min_otdl_var).grid(row=row, column=1, padx=4, sticky="e")
        Button(self.win.body, width=5, text="info",
               command=lambda: Messenger(self.win.topframe).tolerance_info("min_otdl")).grid(row=row, column=2, padx=4)
        row += 1
        Label(self.win.body, text="Minimum rows for Auxiliary", width=30, anchor="w") \
            .grid(row=row, column=0, ipady=5, sticky="w")
        Entry(self.win.body, width=5, textvariable=self.min_aux_var).grid(row=row, column=1, padx=4, sticky="e")
        Button(self.win.body, width=5, text="info",
               command=lambda: Messenger(self.win.topframe).tolerance_info("min_aux")).grid(row=row, column=2, padx=4)
        row += 1
        Label(self.win.body, text="").grid(row=row, column=0)
        row += 1
        Label(self.win.body, text="Page Breaks Between List:", anchor="w").grid(row=row, column=0, sticky="w")
        row += 1
        # Page break between no list and work assignment
        Label(self.win.body, text="  No List and Work Assignment", width=30, anchor="w")\
            .grid(row=row, column=0, ipady=5, sticky="w")
        om_pb_1 = OptionMenu(self.win.body, self.pb_nl_wal_var, "on", "off")
        om_pb_1.config(width=3)
        om_pb_1.grid(row=row, column=1, padx=4, sticky="e")
        Button(self.win.body, width=5, text="info", 
               command=lambda: Messenger(self.win.topframe).tolerance_info("pb_nl_wal"))\
            .grid(row=row, column=2, padx=4)
        row += 1
        # Page break between no list and work assignment
        Label(self.win.body, text="  Work Assignment and OT Desired", width=30, anchor="w") \
            .grid(row=row, column=0, ipady=5, sticky="w")
        om_pb_2 = OptionMenu(self.win.body, self.pb_wal_otdl_var, "on", "off")
        om_pb_2.config(width=3)
        om_pb_2.grid(row=row, column=1, padx=4, sticky="e")
        Button(self.win.body, width=5, text="info",
               command=lambda: Messenger(self.win.topframe).tolerance_info("pb_wal_otdl"))\
            .grid(row=row, column=2, padx=4)
        row += 1
        # Page break between no list and work assignment
        Label(self.win.body, text="  OT Desired and Auxiliary", width=30, anchor="w") \
            .grid(row=row, column=0, ipady=5, sticky="w")
        om_pb_3 = OptionMenu(self.win.body, self.pb_otdl_aux_var, "on", "off")
        om_pb_3.config(width=3)
        om_pb_3.grid(row=row, column=1, padx=4, sticky="e")
        Button(self.win.body, width=5, text="info",
               command=lambda: Messenger(self.win.topframe).tolerance_info("pb_otdl_aux"))\
            .grid(row=row, column=2, padx=4)
        row += 1
        # Display header for 12 and 60 Hour Violations Spread Sheet
        Label(self.win.body, text="").grid(row=row, column=0)
        row += 1
        Label(self.win.body, text="12 and 60 Hour Violations Spreadsheet Settings",
              font=macadj("bold", "Helvetica 18")) \
            .grid(row=row, column=0, sticky="w", columnspan=4)
        row += 1
        Label(self.win.body, text="").grid(row=row, column=0)
        row += 1
        # Display widgets for 12 and 60 Hour Violations Spread Sheet
        Label(self.win.body, text="Minimum rows for Over Max", width=30, anchor="w") \
            .grid(row=row, column=0, ipady=5, sticky="w")
        Entry(self.win.body, width=5, textvariable=self.min_overmax_var).grid(row=row, column=1, padx=4, sticky="e")
        Button(self.win.body, width=5, text="info",
               command=lambda: Messenger(self.win.topframe).tolerance_info("min_overmax"))\
            .grid(row=row, column=2, padx=4)
        row += 1
        Label(self.win.body, text="").grid(row=row, column=0)
        row += 1

        # Display header for OTDL Equitability Spread Sheet
        Label(self.win.body, text="OTDL Equitability Spreadsheet Settings",
              font=macadj("bold", "Helvetica 18")) \
            .grid(row=row, column=0, sticky="w", columnspan=4)
        row += 1
        Label(self.win.body, text="").grid(row=row, column=0)
        row += 1
        # Display widgets for OTDL Equitability Spread Sheet
        Label(self.win.body, text="Minimum rows for OTDL Equitability", width=30, anchor="w") \
            .grid(row=row, column=0, ipady=5, sticky="w")
        Entry(self.win.body, width=5, textvariable=self.min_ot_equit_var).grid(row=row, column=1, padx=4, sticky="e")
        Button(self.win.body, width=5, text="info",
               command=lambda: Messenger(self.win.topframe).tolerance_info("min_ot_equit")) \
            .grid(row=row, column=2, padx=4)
        row += 1
        Label(self.win.body, text="Overtime Calculation Preference", width=30, anchor="w") \
            .grid(row=row, column=0, ipady=5, sticky="w")
        om_ot_equit = OptionMenu(self.win.body, self.ot_calc_pref_var, "all", "off_route")
        om_ot_equit.config(width=7)
        om_ot_equit.grid(row=row, column=1, padx=4, sticky="e")
        Button(self.win.body, width=5, text="info",
               command=lambda: Messenger(self.win.topframe).tolerance_info("ot_calc_pref")) \
            .grid(row=row, column=2, padx=4)
        row += 1
        Label(self.win.body, text="").grid(row=row, column=0)
        row += 1

        # Display header for Overtime Distribution Spread Sheet
        Label(self.win.body, text="Overtime Distribution Spreadsheet Settings",
              font=macadj("bold", "Helvetica 18")) \
            .grid(row=row, column=0, sticky="w", columnspan=4)
        row += 1
        Label(self.win.body, text="").grid(row=row, column=0)
        row += 1
        # Display widgets for Overtime Distribution Spread Sheet
        Label(self.win.body, text="Minimum rows for Overtime Distribution", width=30, anchor="w") \
            .grid(row=row, column=0, ipady=5, sticky="w")
        Entry(self.win.body, width=5, textvariable=self.min_ot_dist_var).grid(row=row, column=1, padx=4, sticky="e")
        Button(self.win.body, width=5, text="info",
               command=lambda: Messenger(self.win.topframe).tolerance_info("min_ot_dist")) \
            .grid(row=row, column=2, padx=4)
        row += 1
        Label(self.win.body, text="Overtime Calculation Preference", width=30, anchor="w") \
            .grid(row=row, column=0, ipady=5, sticky="w")
        om_ot_equit = OptionMenu(self.win.body, self.ot_calc_pref_dist_var, "all", "off_route")
        om_ot_equit.config(width=7)
        om_ot_equit.grid(row=row, column=1, padx=4, sticky="e")
        Button(self.win.body, width=5, text="info",
               command=lambda: Messenger(self.win.topframe).tolerance_info("ot_calc_pref_dist")) \
            .grid(row=row, column=2, padx=4)
        row += 1
        Label(self.win.body, text="").grid(row=row, column=0)

        row += 1
        dashes = ""
        dashcount = 71
        if sys.platform == "darwin":
            dashcount = 55
        for i in range(dashcount):
            dashes = dashes + "_"
        Label(self.win.body, text=dashes, pady=5).grid(row=row, columnspan=4, sticky="w")
        row += 1
        Label(self.win.body, text="Restore Defaults").grid(row=row, column=0, ipady=5, sticky="w")
        Button(self.win.body, width=5, text="set", command=lambda: self.min_ss_presets("default")) \
            .grid(row=row, column=2)
        row += 1
        Label(self.win.body, text="Set rows to one").grid(row=row, column=0, ipady=5, sticky="w")
        Button(self.win.body, width=5, text="set", command=lambda: self.min_ss_presets("one")) \
            .grid(row=row, column=2)
        self.win.fill(row + 1, 15)

    def buttons_frame(self):
        button_submit = Button(self.win.buttons)
        button_apply = Button(self.win.buttons)
        button_back = Button(self.win.buttons)
        button_submit.config(text="Submit", command=lambda: self.apply(True))
        button_apply.config(text="Apply", command=lambda: self.apply(False))
        button_back.config(text="Go Back", command=lambda: MainFrame().start(frame=self.win.topframe))
        if sys.platform == "win32":
            button_submit.config(width=15, anchor="w")
            button_apply.config(width=15, anchor="w")
            button_back.config(width=15, anchor="w")
        else:
            button_submit.config(width=9)
            button_apply.config(width=9)
            button_back.config(width=9)
        button_submit.pack(side=LEFT)
        button_apply.pack(side=LEFT)
        button_back.pack(side=LEFT)
        self.status_update = Label(self.win.buttons, text="", fg="red")
        self.status_update.pack(side=LEFT)

    def min_ss_presets(self, order):
        num = "25"
        over_num = "30"
        ot_num = "19"  # default for otdl equitability minimum rows
        ot_dist_num = "25"  # default for ot distribution minimum rows
        msg = "Minimum rows reset to default. "
        if order == "one":
            num = "1"
            over_num = "1"
            ot_num = "1"
            ot_dist_num = "1"
            msg = "Minimum rows set to one. "
        self.status_update.config(text="{}".format(msg))
        types = ("min_ss_nl", "min_ss_wal", "min_ss_otdl", "min_ss_aux")
        for t in types:  # set minimum row values for improper mandate spreadsheet
            sql = "UPDATE tolerances SET tolerance ='%s' WHERE category = '%s'" % (num, t)
            commit(sql)
        # set minimum row value for overmax spreadsheet
        sql = "UPDATE tolerances SET tolerance ='%s' WHERE category = '%s'" % (over_num, "min_ss_overmax")
        commit(sql)
        # set minimum row value for otdl equitability
        sql = "UPDATE tolerances SET tolerance ='%s' WHERE category = '%s'" % (ot_num, "min_ot_equit")
        commit(sql)
        # set minimum row value for ot distribution
        sql = "UPDATE tolerances SET tolerance ='%s' WHERE category = '%s'" % (ot_dist_num, "min_ot_dist")
        commit(sql)
        pagebreaks = ("pb_nl_wal", "pb_wal_otdl", "pb_otdl_aux")
        if order == "default":
            for pb in pagebreaks:
                sql = "UPDATE tolerances SET tolerance ='%s' WHERE category = '%s'" % ("True", pb)
                commit(sql)
            sql = "UPDATE tolerances SET tolerance ='%s' WHERE category = '%s'" % ("off_route", "ot_calc_pref")
            commit(sql)
            sql = "UPDATE tolerances SET tolerance ='%s' WHERE category = '%s'" % ("off_route", "ot_calc_pref_dist")
            commit(sql)
        self.get_settings()
        self.set_stringvars()

    def check(self, var):
        current_var = ("No List minimum rows", "Work Assignment minimum rows", "OT Desired minimum rows",
                       "Auxiliary minimum rows", "Over Max minimum rows", "OTDL Equitability minimum rows")
        if MinrowsChecker(var).is_empty():
            return True
        if not MinrowsChecker(var).is_numeric():
            text = "The value must be a number for {}".format(current_var[self.check_i])
            messagebox.showerror("Minimum Row Value Entry Error", text, parent=self.win.body)
            return False
        if not MinrowsChecker(var).no_decimals():
            text = "Numbers with decimals are not allowed for {}".format(current_var[self.check_i])
            messagebox.showerror("Minimum Row Value Entry Error", text, parent=self.win.body)
            return False
        if not MinrowsChecker(var).not_negative():
            text = "Numbers less than zero are not allowed for {}".format(current_var[self.check_i])
            messagebox.showerror("Minimum Row Value Entry Error", text, parent=self.win.body)
            return False
        if not MinrowsChecker(var).not_zero():
            text = "Numbers less than one are not allowed for {}".format(current_var[self.check_i])
            messagebox.showerror("Minimum Row Value Entry Error", text, parent=self.win.body)
            return False
        if not MinrowsChecker(var).within_limit(self.minrows_limit):
            text = "Numbers greater than {} are not allowed for {}"\
                .format(self.minrows_limit, current_var[self.check_i])
            messagebox.showerror("Minimum Row Value Entry Error", text, parent=self.win.body)
            return False
        return True

    def apply(self, go_home):
        onrecs_min = (self.min_nl, self.min_wal, self.min_otdl, self.min_aux, self.min_overmax, self.min_ot_equit,
                      self.min_ot_dist)
        onrecs_breaks = (self.pb_nl_wal, self.pb_wal_otdl, self.pb_otdl_aux)
        onrecs_misc = (self.ot_calc_pref, self.ot_calc_pref_dist)
        check_these = (self.min_nl_var.get(), self.min_wal_var.get(), self.min_otdl_var.get(), self.min_aux_var.get(),
                       self.min_overmax_var.get(), self.min_ot_equit_var.get(), self.min_ot_dist_var.get())
        add_these = [self.add_min_nl, self.add_min_wal, self.add_min_otdl, self.add_min_aux, self.add_min_overmax,
                     self.add_min_ot_equit, self.add_min_ot_dist]
        categories = ("min_ss_nl", "min_ss_wal", "min_ss_otdl", "min_ss_aux", "min_ss_overmax", "min_ot_equit",
                      "min_ot_dist")
        pbs = (self.pb_nl_wal_var.get(), self.pb_wal_otdl_var.get(), self.pb_otdl_aux_var.get())
        add_pbs = [self.add_pb_nl_wal, self.add_pb_wal_otdl, self.add_pb_otdl_aux]
        pb_categories = ("pb_nl_wal", "pb_wal_otdl", "pb_otdl_aux")
        misc = (self.ot_calc_pref_var.get(), self.ot_calc_pref_dist_var.get())  # misc stringvars
        add_misc = [self.add_ot_calc_pref, self.add_ot_calc_pref_dist]  # misc values to update to database
        misc_categories = ("ot_calc_pref", "ot_calc_pref_dist")  # list of records in the tolerance table.
        self.check_i = 0
        for var in check_these:  # check each of the minimum rows stringvars
            if not self.check(var):  # if any fail
                return  # stop the method
            self.check_i += 1
        for i in range(len(check_these)):
            add_this = Convert(check_these[i]).zero_not_empty()  # replace empty strings with a zero
            add_these[i] = Handler(add_this).format_str_as_int()  # format the string as an int
            if onrecs_min[i] != add_these[i]:
                sql = "UPDATE tolerances SET tolerance ='%s' WHERE category = '%s'" % (add_these[i], categories[i])
                commit(sql)
                self.report_counter += 1
        for i in range(len(pbs)):  # loop through pagebreak stringvars
            add_pbs[i] = Convert(pbs[i]).onoff_to_bool()
            if onrecs_breaks[i] != str(pbs[i]):
                sql = "UPDATE tolerances SET tolerance ='%s' WHERE category = '%s'" % (add_pbs[i], pb_categories[i])
                commit(sql)
                self.report_counter += 1
        for i in range(len(misc)):  # loop through misc/otdl calculation preferences stringvar
            add_misc[i] = str(misc[i])
            if onrecs_misc[i] != str(misc[i]):
                sql = "UPDATE tolerances SET tolerance ='%s' WHERE category = '%s'" % (add_misc[i], misc_categories[i])
                commit(sql)
                self.report_counter += 1
        if go_home:
            MainFrame().start(frame=self.win.topframe)
        else:
            self.write_report()
            self.get_settings()
            self.set_stringvars()

    def write_report(self):
        text = "No Records Updated"
        if self.report_counter:
            text = "{} Record{} Updated"\
                .format(self.report_counter, Handler(self.report_counter).plurals())
        self.status_update.config(text=text)
        self.report_counter = 0


def apply_tolerance(frame, tolerance, type):
    if not isfloat(tolerance):
        text = "You must enter a number."
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
    if float(tolerance) > 1:
        text = "You must enter a value less than one."
        messagebox.showerror("Tolerance value entry error", text, parent=frame)
        return
    if float(tolerance) < 1:
        number = tolerance.split('.')
        if len(number) == 2:
            if len(number[1]) > 2:
                text = "Value cannot exceed two decimal places."
                messagebox.showerror("Tolerance value entry error", text, parent=frame)
        else:
            if len(number[0]) > 2:
                text = "Value cannot exceed two decimal places."
                messagebox.showerror("Tolerance value entry error", text, parent=frame)
    sql = "UPDATE tolerances SET tolerance ='%s' WHERE category = '%s'" % (tolerance, type)
    commit(sql)
    tolerances(frame)


def tolerance_presets(frame, order):
    if order == "default":
        num = ".25"
    if order == "zero":
        num = "0"
    types = ("ot_own_rt", "ot_tol", "av_tol")
    for t in types:
        sql = "UPDATE tolerances SET tolerance ='%s' WHERE category = '%s'" % (num, t)
        commit(sql)
    tolerances(frame)


def tolerances(frame):
    frame.destroy()
    f = Frame(projvar.root)
    f.pack(fill=BOTH, side=LEFT)
    c1 = Canvas(f)
    c1.pack(fill=BOTH, side=BOTTOM)
    # apply and close buttons
    button = Button(c1)
    button.config(text="Go Back", width=20, command=lambda: MainFrame().start(frame=f))
    if sys.platform == "win32":
        button.config(anchor="w")
    button.pack(side=LEFT)
    # link up the canvas and scrollbar
    s = Scrollbar(f)
    c = Canvas(f, width=1600)
    s.pack(side=RIGHT, fill=BOTH)
    c.pack(side=LEFT, fill=BOTH, pady=10, padx=20)
    s.configure(command=c.yview, orient="vertical")
    c.configure(yscrollcommand=s.set)
    if sys.platform == "win32":
        c.bind_all('<MouseWheel>', lambda event: c.yview_scroll(int(projvar.mousewheel * (event.delta / 120)), "units"))
    elif sys.platform == "darwin":
        c.bind_all('<MouseWheel>', lambda event: c.yview_scroll(int(projvar.mousewheel * event.delta), "units"))
    elif sys.platform == "linux":
        c.bind_all('<Button-4>', lambda event: c.yview('scroll', -1, 'units'))
        c.bind_all('<Button-5>', lambda event: c.yview('scroll', 1, 'units'))
    # create the frame inside the canvas
    ff = Frame(c)
    c.create_window((0, 0), window=ff, anchor=NW)
    # page contents
    sql = "SELECT * FROM tolerances"
    results = inquire(sql)
    ot_own_rt = StringVar(ff)
    ot_tol = StringVar(ff)
    av_tol = StringVar(ff)
    Label(ff, text="Tolerances", font=macadj("bold", "Helvetica 18"), anchor="w") \
        .grid(row=0, column=0, columnspan=4, sticky="w")
    Label(ff, text=" ").grid(row=1, column=0, columnspan=4, sticky="w")
    Label(ff, text="Overtime on own route", width=20, anchor="w") \
        .grid(row=2, column=0, ipady=5, sticky="w")
    Entry(ff, width=5, textvariable=ot_own_rt).grid(row=2, column=1, padx=4)
    Button(ff, width=5, text="change", command=lambda: apply_tolerance(f, ot_own_rt.get(), "ot_own_rt")) \
        .grid(row=2, column=2, padx=4)
    Button(ff, width=5, text="info", command=lambda: Messenger(f).tolerance_info("OT_own_route")) \
        .grid(row=2, column=3, padx=4)
    Label(ff, text="Overtime off own route").grid(row=3, column=0, ipady=5, sticky="w")
    Entry(ff, width=5, textvariable=ot_tol).grid(row=3, column=1)
    Button(ff, width=5, text="change", command=lambda: apply_tolerance(f, ot_tol.get(), "ot_tol")) \
        .grid(row=3, column=2)
    Button(ff, width=5, text="info", command=lambda: Messenger(f).tolerance_info("OT_off_route")) \
        .grid(row=3, column=3)
    Label(ff, text="Availability tolerance").grid(row=4, column=0, ipady=5, sticky="w")
    Entry(ff, width=5, textvariable=av_tol).grid(row=4, column=1)
    Button(ff, width=5, text="change", command=lambda: apply_tolerance(f, av_tol.get(), "av_tol")) \
        .grid(row=4, column=2)
    Button(ff, width=5, text="info", command=lambda: Messenger(f).tolerance_info("availability")) \
        .grid(row=4, column=3)
    dashes = ""
    dashcount = 59
    if sys.platform == "darwin":
        dashcount = 47
    for _ in range(dashcount):
        dashes = dashes + "_"
    Label(ff, text=dashes, pady=5).grid(row=5, columnspan=4, sticky="w")
    Label(ff, text="Recommended settings").grid(row=6, column=0, ipady=5, sticky="w")
    Button(ff, width=5, text="set", command=lambda: tolerance_presets(f, "default")) \
        .grid(row=6, column=2)
    Label(ff, text="Set tolerances to zero").grid(row=7, column=0, ipady=5, sticky="w")
    Button(ff, width=5, text="set", command=lambda: tolerance_presets(f, "zero")) \
        .grid(row=7, column=2)
    ot_own_rt.set(results[0][2])
    ot_tol.set(results[1][2])
    av_tol.set(results[2][2])
    projvar.root.update()
    c.config(scrollregion=c.bbox("all"))


def apply_station(switch, station, frame):
    if switch == "enter":
        if station.get().strip() == "" or station.get().strip() == "x":
            messagebox.showerror("Prohibited Action",
                                 "You can not enter a blank entry for a station.",
                                 parent=frame)
            return
        if station.get() in projvar.list_of_stations:
            messagebox.showerror("Prohibited Action",
                                 "That station is already in the list of stations.",
                                 parent=frame)
            return
    if switch == "enter":
        sql = "INSERT INTO stations (station) VALUES('%s')" % (station.get().strip())
        commit(sql)
        projvar.list_of_stations.append(station.get())
    if switch == "delete":
        if station == "out of station":
            text = "You can not delete the \"out of station\" listing."
            messagebox.showerror("Action not allowed", text, parent=frame)
            return
        if messagebox.askokcancel("Delete Station",
                                  "Are you sure you want to delete {}? \n"
                                  "The station will be deleted and maintenance actions will\n"
                                  "clean any orphan carriers, clock rings and indexes from\n"
                                  "database. This can not be reversed.".format(station),
                                  parent=frame):
            sql = "DELETE FROM stations WHERE station='%s'" % station
            commit(sql)
            database_clean_carriers()
            database_clean_rings()
            if projvar.invran_station == station:
                reset("none")  # reset initial value of globals
    # access list of stations from database
    sql = "SELECT * FROM stations ORDER BY station"
    results = inquire(sql)
    # define and populate list of stations variable
    del projvar.list_of_stations[:]
    for stat in results:
        projvar.list_of_stations.append(stat[0])
    station_list(frame)


def station_update_apply(frame, old_station, new_station):
    if old_station.get() == "select a station":
        messagebox.showerror("Prohibited Action",
                             "Please select a station.",
                             parent=frame)
        return
    if new_station.get().strip() == "" or \
            new_station.get() == "enter a new station name" or \
            new_station.get().strip() == "x":
        messagebox.showerror("Prohibited Action",
                             "You can not enter a blank entry for a station.",
                             parent=frame)
        return
    if projvar.invran_station == old_station.get():
        reset("none")  # reset initial value of globals
    go_ahead = True
    duplicate = False
    if new_station.get() in projvar.list_of_stations:
        go_ahead = messagebox.askokcancel("Duplicate Detected",
                                          "This station already exist in the list of stations. "
                                          "If you proceed, all records for {} will be merged with "
                                          "records from {}. Do you want to proceed?"
                                          .format(old_station.get(), new_station.get()),
                                          parent=frame)
        duplicate = True
    if duplicate and go_ahead:
        sql = "DELETE FROM stations WHERE station='%s'" % old_station.get()
        commit(sql)
        projvar.list_of_stations.remove(new_station.get())
    if go_ahead:
        sql = "UPDATE stations SET station='%s' WHERE station='%s'" % (new_station.get(), old_station.get())
        commit(sql)
        sql = "UPDATE carriers SET station='%s' WHERE station='%s'" % (new_station.get(), old_station.get())
        commit(sql)
        sql = "UPDATE station_index SET kb_station='%s' WHERE kb_station='%s'" % (new_station.get(), old_station.get())
        commit(sql)
        projvar.list_of_stations.append(new_station.get())
        projvar.list_of_stations.remove(old_station.get())
        station_list(frame)
    if not go_ahead:
        return


def station_list(frame):
    frame.destroy()
    f = Frame(projvar.root)
    f.pack(fill=BOTH, side=LEFT)
    c1 = Canvas(f)
    c1.pack(fill=BOTH, side=BOTTOM)
    button = Button(c1)
    button.config(text="Go Back", width=20, command=lambda: MainFrame().start(frame=f))
    if sys.platform == "win32":
        button.config(anchor="w")
    button.pack(side=LEFT)
    # link up the canvas and scrollbar
    s = Scrollbar(f)
    c = Canvas(f, width=1600)
    s.pack(side=RIGHT, fill=BOTH)
    c.pack(side=LEFT, fill=BOTH, pady=10, padx=20)
    s.configure(command=c.yview, orient="vertical")
    c.configure(yscrollcommand=s.set)
    if sys.platform == "win32":
        c.bind_all('<MouseWheel>', lambda event: c.yview_scroll(int(projvar.mousewheel * (event.delta / 120)), "units"))
    elif sys.platform == "darwin":
        c.bind_all('<MouseWheel>', lambda event: c.yview_scroll(int(projvar.mousewheel * event.delta), "units"))
    elif sys.platform == "linux":
        c.bind_all('<Button-4>', lambda event: c.yview('scroll', -1, 'units'))
        c.bind_all('<Button-5>', lambda event: c.yview('scroll', 1, 'units'))
    # create the frame inside the canvas
    ff = Frame(c)
    c.create_window((0, 0), window=ff, anchor=NW)
    # page title
    row = 0
    Label(ff, text="Manage Station List", font=macadj("bold", "Helvetica 18")) \
        .grid(row=row, columnspan=2, sticky="w")
    row += 1
    Label(ff, text="____________________________________________________", pady=5). \
        grid(row=row, columnspan=2, sticky="w")
    row += 1
    # enter new stations
    new_name = StringVar(ff)
    Label(ff, text="Enter New Station", pady=5, font=macadj("bold", "Helvetica 18")) \
        .grid(row=row, columnspan=2, sticky="w")
    row += 1
    e = Entry(ff, width=35, textvariable=new_name)
    e.grid(row=row, column=0, sticky="w")
    new_name.set("")
    Button(ff, width=5, anchor="w", text="ENTER", command=lambda: apply_station("enter", new_name, f)). \
        grid(row=row, column=1, sticky="w")
    row += 1
    Label(ff, text="____________________________________________________", pady=5). \
        grid(row=row, columnspan=2, sticky="w")
    row += 1
    # list current list of stations and delete buttons.
    sql = "SELECT * FROM stations ORDER BY station"
    results = inquire(sql)
    Label(ff, text="List Of Stations", font=macadj("bold", "Helvetica 18"), pady=5) \
        .grid(row=row, columnspan=2, sticky="w")
    row += 1
    for record in results:
        Button(ff, text=record[0], width=30, anchor="w").grid(row=row, column=0, sticky="w")
        Button(ff, text="delete", command=lambda x=record[0]: apply_station("delete", x, f)) \
            .grid(row=row, column=1, sticky="w")
        row += 1
    Label(ff, text="____________________________________________________", pady=5). \
        grid(row=row, columnspan=2, sticky="w")
    row += 1
    if len(results) > 1:
        # change names of stations
        Label(ff, text="Change Station Name", font=macadj("bold", "Helvetica 18")) \
            .grid(row=row, column=0, sticky="w")
        row += 1
        all_stations = []
        for rec in results:
            all_stations.append(rec[0])
        if "out of station" in all_stations:
            all_stations.remove("out of station")
        old_station = StringVar(ff)
        om = OptionMenu(ff, old_station, *all_stations)
        om.config(width="35")
        om.grid(row=row, column=0, sticky="w", columnspan=2)
        row += 1
        old_station.set("select a station")
        Label(ff, text="enter a new name:").grid(row=row, column=0, sticky="w")
        row += 1
        new_station = StringVar(ff)
        Entry(ff, textvariable=new_station, width="30").grid(row=row, column=0, sticky="w")
        new_station.set("enter a new station name")
        Button(ff, text="update", command=lambda: station_update_apply(f, old_station, new_station)) \
            .grid(row=row, column=1, sticky="w")
        row += 1
        Label(ff, text="____________________________________________________", pady=5). \
            grid(row=row, columnspan=2, sticky="w")
        row += 1
    # find and display list of unique stations
    Label(ff, text="List Of Stations", pady=5, font=macadj("bold", "Helvetica 18")) \
        .grid(row=row, columnspan=3, sticky="w")
    row += 1
    Label(ff, text="(referenced in carrier database)", pady=5) \
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
    for ss in unique_station:
        Label(ff, text="{}.  {}".format(count, ss)).grid(row=row, columnspan=2, sticky="w")
        count += 1
        row += 1
    projvar.root.update()
    c.config(scrollregion=c.bbox("all"))


def apply_mi(frame, array_var, ls, ns, station, route, date):  # enter changes from multiple input into database
    x = date.get()
    year = IntVar()
    month = IntVar()
    day = IntVar()
    y = projvar.invran_date_week[x].strftime("%Y").lstrip("0")
    m = projvar.invran_date_week[x].strftime("%m").lstrip("0")
    d = projvar.invran_date_week[x].strftime("%d").lstrip("0")
    year.set(y)
    month.set(m)
    day.set(d)
    sql = "SELECT * FROM ns_configuration"
    ns_results = inquire(sql)
    ns_dict = {}  # build dictionary for ns days
    for r in ns_results:  # build dictionary for rotating ns days
        ns_dict[r[2]] = r[0]
    ns_dict["none"] = "none"  # add "none" to dictionary
    for i in range(len(array_var)):  # loop through all received data
        if "fixed: " not in ns[i].get():
            passed_ns = ns[i].get().split("  ")  # break apart the day/color_code
            ns[i].set(ns_dict[passed_ns[1]])  # match color_code to proper color_code in dict and set
        else:
            passed_ns = ns[i].get().split("  ")  # do not subject the fixed to the dictionary
            ns[i].set(passed_ns[1])
        # if there is a differance, then put the new record in the database
        if array_var[i][2] != ls[i].get() or array_var[i][3] != ns[i].get() or array_var[i][5] != station[i].get():
            apply(year, month, day, array_var[i][1], ls[i], ns[i], route[i], station[i], frame)


def mass_input(frame, day, sort):
    sql = ""
    frame.destroy()
    switch_f7 = Frame(projvar.root)
    switch_f7.pack()
    c1 = Canvas(switch_f7)
    c1.pack(fill=BOTH, side=BOTTOM)
    button_submit = Button(c1)  # apply and close buttons
    button_apply = Button(c1)
    button_back = Button(c1)
    button_submit.config(text="Submit", width=10, command=lambda:
        [switch_f7.destroy(), apply_mi(switch_f7, array_var, mi_list, mi_nsday, mi_station, mi_route,
                                       pass_date), MainFrame().start()])
    button_apply.config(text="Apply", width=10,
           command=lambda: [apply_mi(switch_f7, array_var, mi_list, mi_nsday, mi_station, mi_route, pass_date),
                            mass_input(switch_f7, day, sort)])
    button_back.config(text="Go Back", width=10,
           command=lambda: MainFrame().start(frame=switch_f7))
    if sys.platform == "win32":
        button_submit.config(anchor="w")
        button_apply.config(anchor="w")
        button_back.config(anchor="w")
    button_submit.pack(side=LEFT)
    button_apply.pack(side=LEFT)
    button_back.pack(side=LEFT)
    # link up the canvas and scrollbar
    s = Scrollbar(switch_f7)
    c = Canvas(switch_f7, height=800, width=1600)
    s.pack(side=RIGHT, fill=BOTH)
    c.pack(side=LEFT, fill=BOTH, pady=10, padx=20)
    s.configure(command=c.yview, orient="vertical")
    c.configure(yscrollcommand=s.set)
    if sys.platform == "win32":
        c.bind_all('<MouseWheel>', lambda event: c.yview_scroll(int(projvar.mousewheel * (event.delta / 120)), "units"))
    elif sys.platform == "darwin":
        c.bind_all('<MouseWheel>', lambda event: c.yview_scroll(int(projvar.mousewheel * event.delta), "units"))
    elif sys.platform == "linux":
        c.bind_all('<Button-4>', lambda event: c.yview('scroll', -1, 'units'))
        c.bind_all('<Button-5>', lambda event: c.yview('scroll', 1, 'units'))
    # create the frame inside the canvas
    head_f = Frame(c)
    c.create_window((0, 0), window=head_f, anchor=NW)
    f = Frame(c)
    c.create_window((0, 50), window=f, anchor=NW)
    # set up the option menus to order results by day and sort criteria.
    mi_date = StringVar()
    mi_sort = StringVar()
    opt_day = ["Sat", "Sun", "Mon", "Tue", "Wed", "Thu", "Fri"]
    opt_sort = ["name", "list", "ns day"]
    mi_date.set(day)
    if projvar.invran_weekly_span:  # if investigation range is daily
        mi_date.set(day)
        om1 = OptionMenu(head_f, mi_date, *opt_day)
        om1.config(width="5")
        om1.grid(row=0, column=0)
    mi_sort.set(sort)
    om2 = OptionMenu(head_f, mi_sort, *opt_sort)
    om2.grid(row=0, column=1)
    om2.config(width="8")
    Button(head_f, text="set", width=6, command=lambda: mass_input(switch_f7, mi_date.get(), mi_sort.get())) \
        .grid(row=0, column=2)
    # figure out the day and display
    pass_date = IntVar()
    if projvar.invran_weekly_span:   # if investigation range is weekly
        for i in range(len(projvar.invran_date_week)):
            if opt_day[i] == day:
                f_date = projvar.invran_date_week[i].strftime("%a - %b %d, %Y")
                pass_date.set(i)
                Label(f, text="Showing results for {}"
                      .format(f_date), font=macadj("bold", "Helvetica 18"), justify=LEFT) \
                    .grid(row=0, column=0, columnspan=4, sticky=W)
    if not projvar.invran_weekly_span:  # if investigation range is daily
        for i in range(len(opt_day)):
            if projvar.invran_date.strftime("%a") == opt_day[i]:
                f_date = projvar.invran_date.strftime("%a - %b %d, %Y")
                pass_date.set(i)
                Label(f, text="Showing results for {}"
                      .format(f_date), font=macadj("bold", "Helvetica 18"), justify=LEFT) \
                    .grid(row=0, column=0, columnspan=4, sticky=W)
    # access database
    for i in range(len(projvar.invran_date_week)):
        if opt_day[i] == day:
            if projvar.invran_weekly_span:  # if investigation range is weekly
                sql = "SELECT effective_date, carrier_name,list_status, ns_day,route_s, station, rowid" \
                      " FROM carriers WHERE effective_date <= '%s'" \
                      "ORDER BY carrier_name, effective_date" % (projvar.invran_date_week[i])
            else:
                sql = "SELECT effective_date, carrier_name,list_status, ns_day,route_s, station, rowid" \
                      " FROM carriers WHERE effective_date <= '%s'" \
                      "ORDER BY carrier_name, effective_date" % projvar.invran_date
    results = inquire(sql)
    # initialize arrays for data sorting
    carrier_list = []
    candidates = []
    otdl_array = []
    wal_array = []
    nl_array = []
    ptf_array = []
    aux_array = []
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
            if winner[5] == projvar.invran_station:  # if that record matches the current station...
                carrier_list.append(winner)  # then insert that record in the carrier list
                if sort == "list":  # sort carrier list by ot list if selected
                    if winner[2] == "otdl":
                        otdl_array.append(winner)
                    if winner[2] == "wal":
                        wal_array.append(winner)
                    if winner[2] == "nl":
                        nl_array.append(winner)
                    if winner[2] == "ptf":
                        ptf_array.append(winner)
                    if winner[2] == "aux":
                        aux_array.append(winner)
                if sort == "ns day":  # sort carrier list by ns day if selected
                    if winner[3] == "yellow":
                        yellow_array.append(winner)
                    if winner[3] == "blue":
                        blue_array.append(winner)
                    if winner[3] == "green":
                        green_array.append(winner)
                    if winner[3] == "brown":
                        brown_array.append(winner)
                    if winner[3] == "red":
                        red_array.append(winner)
                    if winner[3] == "black":
                        black_array.append(winner)
                    if winner[3] == "none":
                        none_array.append(winner)
        del candidates[:]
    # Display results XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
    i = 1
    array_var = []
    list_header = ""
    # set up first header
    if sort == "name":
        for car in carrier_list:
            array_var.append(car)
        list_header = "carrier list"
    if sort == "list":
        array_var = nl_array + wal_array + otdl_array + ptf_array + aux_array
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
    Label(f, text=list_header).grid(row=i, column=0)
    i += 1
    sql = "SELECT * FROM ns_configuration"
    ns_results = inquire(sql)
    ns_dict = {}  # build dictionary for ns days
    days = ("sat", "mon", "tue", "wed", "thu", "fri")
    for r in ns_results:  # build dictionary for rotating ns days
        ns_dict[r[0]] = r[2]
    for d in days:  # expand dictionary for fixed days
        ns_dict[d] = "fixed: " + d
    ns_dict["none"] = "none"  # add "none" to dictionary
    # intialize arrays for option menus
    mi_list = []
    opt_list = "nl", "wal", "otdl", "aux", "ptf"
    mi_nsday = []
    nsk = []
    days = ("sat", "mon", "tue", "wed", "thu", "fri")
    for each in projvar.ns_code.keys():
        nsk.append(each)  # make an array of projvar.ns_code keys
    opt_nsday = []  # make an array of "day / color" options for option menu
    for each in projvar.ns_code:
        ns_option = projvar.ns_code[each] + "  " + ns_dict[each]  # make a string for each day/color
        if each in days:
            ns_option = "fixed:" + "  " + each  # if the ns day is fixed - make a special string
        if each == "none":
            ns_option = "---" + "  " + each  # if the ns day is "none" - make a special string
        opt_nsday.append(ns_option)
    mi_station = []
    mi_route = []
    count = 0
    for record in array_var:  # loop to put information on to window
        # set up color
        if i & 1:
            color = "light yellow"
        else:
            color = "white"
        if sort == "list":
            if list_header != record[2]:
                list_header = record[2]
                Label(f, text=list_header).grid(row=i, column=0)
                i += 1
        if sort == "ns day":
            if list_header != record[3]:
                list_header = record[3]
                Label(f, text=list_header).grid(row=i, column=0)
                i += 1
        # set up carrier name button and variable
        Button(f, text=record[1], width=macadj(24, 20), anchor="w", bg=color, bd=0).grid(row=i, column=0)
        # set up list status option menu and variable
        mi_list.append(StringVar(f))
        om_list = OptionMenu(f, mi_list[count], *opt_list)  # configuration below
        om_list.grid(row=i, column=1, ipadx=0)
        mi_list[count].set(record[2])
        # set up ns day option menu and variable
        mi_nsday.append(StringVar(f))
        om_nsday = OptionMenu(f, mi_nsday[count], *opt_nsday)  # configuration below
        om_nsday.grid(row=i, column=2)
        ns_index = nsk.index(record[3])
        mi_nsday[count].set(opt_nsday[ns_index])
        # set up station option menu and variable
        mi_station.append(StringVar(f))
        om_station = OptionMenu(f, mi_station[count], *projvar.list_of_stations)  # configuration below
        om_station.grid(row=i, column=3)
        mi_station[count].set(record[5])
        # adjust optionmenu configuration by platform
        if sys.platform == "darwin":
            om_list.config(width=4, bg=color)
            om_nsday.config(width=9, bg=color)
            om_station.config(width=18, bg=color)
        else:
            om_list.config(width=5, anchor="w", bg=color, relief='ridge', bd=0)
            om_nsday.config(width=10, anchor="w", bg=color, relief='ridge', bd=0)
            om_station.config(width=28, anchor="w", bg=color, relief='ridge', bd=0)
        # set up route variable - not visible but passed along with other variables
        mi_route.append(StringVar(f))
        mi_route[count].set(record[4])
        count += 1
        i += 1
    del carrier_list[:]
    projvar.root.update()
    c.config(scrollregion=c.bbox("all"))


def tab_selected(t):  # attach notebook tab for
    global current_tab
    current_tab = t


def output_tab(frame, list_carrier):
    frame.destroy()
    switch_f5 = Frame(projvar.root, bg="white")
    switch_f5.pack(fill=BOTH, side=LEFT)
    c1 = Canvas(switch_f5)
    c1.pack(fill=BOTH, side=BOTTOM)
    Button(c1, text="spreadsheet", width=15, anchor="w",
           command=lambda: ImpManSpreadsheet().create(switch_f5)).pack(side=LEFT)
    Button(c1, text="Go Back", width=15, anchor="w",
           command=lambda: MainFrame().start(frame=switch_f5)).pack(side=LEFT)
    dates = []  # array containing days
    if projvar.invran_weekly_span:  # if investigation range is weekly
        dates = projvar.invran_date_week
    if not projvar.invran_weekly_span:  # if investigation range is daily
        dates.append(projvar.invran_date)
    if projvar.invran_weekly_span:  # if investigation range is weekly
        sql = "SELECT * FROM rings3 WHERE rings_date BETWEEN '%s' AND '%s' ORDER BY rings_date, carrier_name" \
              % (projvar.invran_date_week[0], projvar.invran_date_week[6])
    else:
        sql = "SELECT * FROM rings3 WHERE rings_date = '%s' ORDER BY rings_date, " \
              "carrier_name" % projvar.invran_date
    r_rings = inquire(sql)
    sql = "SELECT * FROM tolerances"  # get tolerances
    tol_results = inquire(sql)
    ot_own_rt = tol_results[0][2]
    ot_tol = tol_results[1][2]
    av_tol = tol_results[2][2]
    daily_list = []  # array
    candidates = []
    dl_nl = []
    dl_wal = []
    dl_otdl = []
    dl_aux = []
    # list the names of the tabs
    tab = ["saturday", "sunday", "monday", "tuesday", "wednesday", "thursday", "friday"]
    c = ["C0", "C1", "C2", "C3", "C4", "C5", "C6"]
    global current_tab
    current_tab = 0
    tab_control = ttk.Notebook(switch_f5)  # Create Tab Control
    tab_control.pack(expand=1, fill="both")
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
                if winner[5] == projvar.invran_station:
                    daily_list.append(winner)  # add the record if it matches the station
                del candidates[:]  # empty out the candidates array.
        for item in daily_list:  # sort carriers in daily list by the list they are in
            if item[2] == "nl":
                dl_nl.append(item)
            if item[2] == "wal":
                dl_wal.append(item)
            if item[2] == "otdl":
                dl_otdl.append(item)
            if item[2] in ("aux", "ptf"):
                dl_aux.append(item)
        tabs = Frame(tab_control)  # put frame in notebook
        tabs.pack(fill=BOTH, side=LEFT)
        if projvar.invran_weekly_span:  # if investigation range is weekly
            tab_control.add(tabs, text="{}".format(tab[t]))  # Add the tab
        c[t] = Canvas(tabs, width=1600, bg="white")  # put canvas inside notebook frame
        s = Scrollbar(tabs, command=c[t].yview)  # define and bind the scrollbar with the canvas
        c[t].config(yscrollcommand=s.set, scrollregion=(0, 0, 100, 5000))  # bind the canvas with the scrollbar
        #   Enable mousewheel
        c[t].bind("<Map>", lambda event, t=t: tab_selected(t))
        if sys.platform == "win32":
            c[current_tab].bind_all('<MouseWheel>',
                                    lambda event: c[current_tab].yview_scroll
                                    (int(projvar.mousewheel * (event.delta / 120)), "units"))
        elif sys.platform == "darwin":
            c[current_tab].bind_all('<MouseWheel>',
                                    lambda event: c[current_tab].yview_scroll
                                    (int(projvar.mousewheel * event.delta), "units"))
        elif sys.platform == "linux":
            c[current_tab].bind_all('<Button-4>', lambda event: c[current_tab].yview('scroll', -1, 'units'))
            c[current_tab].bind_all('<Button-5>', lambda event: c[current_tab].yview('scroll', 1, 'units'))

        s.pack(side=RIGHT, fill=BOTH)
        c[t].pack(side=LEFT, fill=BOTH, expand=True)
        f = Frame(c[t], bg="white")  # put a frame in the canvas
        f.pack()
        c[t].create_window((0, 0), window=f, anchor=NW)  # create window with frame
        oi = 0
        Label(f, text=day.strftime("%A  %m/%d/%y"), justify=LEFT, anchor=W, font=macadj("bold", "Helvetica 18"),
              pady=5, bg="white").grid(row=oi, column=0, columnspan=10, sticky=W)
        in_color = "white"
        out_color = "light goldenrod yellow"
        oi += 1
        #  no list xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
        Label(f, text="no list", justify=LEFT, bg="white",
              font=macadj('Helvetica 10 bold', 'Futura 16 bold')) \
            .grid(sticky=W, row=oi, column=0, columnspan=10)
        oi += 1
        Label(f, text=" moves", bg="gray90", width=macadj(24, 16), anchor="w") \
            .grid(row=oi, column=4, columnspan=4)  # top of move total
        Label(f, text="off", bg="white").grid(row=oi, column=9)  # top of off route
        Label(f, text="ot off", bg="white").grid(row=oi, column=10)  # top of ot off route
        oi += 1
        Label(f, text="Carrier", bg="white").grid(row=oi, column=0, sticky=W)
        Label(f, text="note", bg="white").grid(row=oi, column=1)
        Label(f, text="5200", bg="white").grid(row=oi, column=2)
        Label(f, text="RS", bg="white").grid(row=oi, column=3)
        Label(f, text=macadj("MV off", "off"), bg="white").grid(row=oi, column=4)
        Label(f, text=macadj("MV on", "on"), bg="white").grid(row=oi, column=5)
        Label(f, text="Rte", bg="white").grid(row=oi, column=6)
        Label(f, text="total", bg="white").grid(row=oi, column=7)
        Label(f, text="OT", bg="white").grid(row=oi, column=8)
        Label(f, text="route", bg="white").grid(row=oi, column=9)
        Label(f, text="route", bg="white").grid(row=oi, column=10)
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
                        cc = 0
                        for i in range(int(len(s_moves) / 3)):
                            total = float(s_moves[cc + 1]) - float(s_moves[cc])  # calc off time off route
                            cc = cc + 3
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
                        Label(f, text=each[1], anchor=W, width=macadj(21, 16), relief=RIDGE, bg=in_color) \
                            .grid(row=oi, column=0)  # name
                        if each[4] == "none":
                            code = ""
                        else:
                            code = each[4]
                        Label(f, text=code, justify=LEFT, width=macadj(6, 4), relief=RIDGE, bg=in_color) \
                            .grid(row=oi, column=1)  # code
                        if each[2] == "" or each[2] == " ":  # handle empty 5200 strings
                            t_hrs = ""
                        else:
                            t_hrs = format(float(each[2]), '.2f')
                        Label(f, text=t_hrs, justify=LEFT, width=macadj(6, 4), relief=RIDGE, bg=in_color) \
                            .grid(row=oi, column=2)  # 5200
                        if each[3] == "" or each[3] == " ":
                            rs = ""  # handle empty RS strings
                        else:
                            rs = format(float(each[3]), '.2f')
                        Label(f, text=rs, justify=LEFT, width=macadj(6, 4), relief=RIDGE, bg=in_color) \
                            .grid(row=oi, column=3)  # return to station
                        count = 0
                        if move_count == 0:  # if there are no moves, fill in with empty cells.
                            for i in range(4, 8):
                                if i < 7:
                                    color = in_color
                                else:
                                    color = out_color
                                if i == 6:
                                    ml = 5
                                else:
                                    ml = 4
                                Label(f, text="", justify=LEFT, width=macadj(6, ml),
                                      relief=RIDGE, bg=color).grid(row=oi, column=i)
                        for i in range(move_count):  # if there are moves, create + populate cells
                            Label(f, text=format(float(s_moves[count]), '.2f'), justify=LEFT, width=macadj(6, 4),
                                  relief=RIDGE, bg=in_color).grid(row=oi, column=4)  # move off
                            count += 1
                            Label(f, text=format(float(s_moves[count]), '.2f'), justify=LEFT, width=macadj(6, 4),
                                  relief=RIDGE, bg=in_color).grid(row=oi, column=5)  # move on
                            count += 1
                            Label(f, text=s_moves[count], justify=LEFT, width=macadj(6, 5),
                                  relief=RIDGE, bg=in_color).grid(row=oi, column=6)  # route
                            count += 1
                            Label(f, text=format(move_totals[i], '.2f'), justify=LEFT, width=macadj(6, 4),
                                  relief=RIDGE, bg=out_color).grid(row=oi, column=7)  # move total
                            if i < move_count - 1:
                                oi += 1
                        Label(f, text=format(ot, '.2f'), justify=LEFT, width=macadj(6, 4),
                              relief=RIDGE, bg=out_color) \
                            .grid(row=oi, column=8)  # overtime
                        Label(f, text=format(off_route, '.2f'), justify=LEFT, width=macadj(6, 4),
                              relief=RIDGE, bg=out_color) \
                            .grid(row=oi, column=9)  # off route
                        Label(f, text=format(ot_off_route, '.2f'), justify=LEFT, width=macadj(6, 4),
                              relief=RIDGE, bg=out_color) \
                            .grid(row=oi, column=10)  # OT off route
                        oi += 1
            #  if there is no match, then just printe the name.
            if match == "miss":
                Label(f, text=line[1], anchor=W, width=macadj(21, 16), relief=RIDGE, bg=in_color) \
                    .grid(row=oi, column=0)  # name
                for i in range(10):
                    if i < 6:
                        color = in_color
                    else:
                        color = out_color
                    if i == 5:
                        ml = 5
                    else:
                        ml = 4
                    Label(f, text="", width=macadj(6, ml), relief=RIDGE, bg=color) \
                        .grid(row=oi, column=i + 1)  # generate blank cells
                oi += 1
        oi += 1
        Label(f, text="", height=2, bg="white").grid(row=oi, column=0)
        Label(f, text=format(ot_total, '.2f'), justify=LEFT, width=macadj(6, 4), relief=RIDGE, bg=out_color) \
            .grid(row=oi, column=8)  # overtime
        Label(f, text=format(ot_off_total, '.2f'), justify=LEFT, width=macadj(6, 4), relief=RIDGE, bg=out_color) \
            .grid(row=oi, column=10)  # OT off route
        oi += 2
        # work assignment list xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
        Label(f, text="work assignment list", justify=LEFT,
              font=macadj('Helvetica 10 bold', 'Futura 16 bold'), bg="white") \
            .grid(sticky=W, row=oi, column=0, columnspan=10)
        oi += 1
        Label(f, text=" moves", bg="gray90", width=macadj(24, 16), anchor="w") \
            .grid(row=oi, column=4, columnspan=4)  # top of move total
        Label(f, text="off", bg="white").grid(row=oi, column=9)  # top of off route
        Label(f, text="ot off", bg="white").grid(row=oi, column=10)  # top of ot off route
        oi += 1
        Label(f, text="Carrier", bg="white").grid(row=oi, column=0, sticky=W)
        Label(f, text="note", bg="white").grid(row=oi, column=1)
        Label(f, text="5200", bg="white").grid(row=oi, column=2)
        Label(f, text="RS", bg="white").grid(row=oi, column=3)
        Label(f, text="off", bg="white").grid(row=oi, column=4)
        Label(f, text="on", bg="white").grid(row=oi, column=5)
        Label(f, text="Rte", bg="white").grid(row=oi, column=6)
        Label(f, text="total", bg="white").grid(row=oi, column=7)
        Label(f, text="OT", bg="white").grid(row=oi, column=8)
        Label(f, text="route", bg="white").grid(row=oi, column=9)
        Label(f, text="route", bg="white").grid(row=oi, column=10)
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
                        cc = 0
                        for i in range(int(len(s_moves) / 3)):
                            total = float(s_moves[cc + 1]) - float(s_moves[cc])  # calc off time off route
                            cc = cc + 3
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
                        Label(f, text=each[1], anchor=W, width=macadj(21, 16), relief=RIDGE, bg=in_color) \
                            .grid(row=oi, column=0)  # name
                        if each[4] == "none":
                            code = ""
                        else:
                            code = each[4]
                        Label(f, text=code, justify=LEFT, width=macadj(6, 4), relief=RIDGE, bg=in_color) \
                            .grid(row=oi, column=1)  # code
                        if each[2] == "" or each[2] == " ":  # handle empty 5200 strings
                            t_hrs = ""
                        else:
                            t_hrs = format(float(each[2]), '.2f')
                        Label(f, text=t_hrs, justify=LEFT, width=macadj(6, 4), relief=RIDGE, bg=in_color) \
                            .grid(row=oi, column=2)  # 5200
                        if each[3] == "" or each[3] == " ":
                            rs = ""  # handle empty RS strings
                        else:
                            rs = format(float(each[3]), '.2f')
                        Label(f, text=rs, justify=LEFT, width=macadj(6, 4), relief=RIDGE, bg=in_color) \
                            .grid(row=oi, column=3)  # return to station
                        count = 0
                        if move_count == 0:  # if there are no moves, fill in with empty cells.
                            for i in range(4, 8):
                                if i < 7:
                                    color = in_color
                                else:
                                    color = out_color
                                if i == 6:
                                    ml = 5
                                else:
                                    ml = 4
                                Label(f, text="", justify=LEFT, width=macadj(6, ml),
                                      relief=RIDGE, bg=color).grid(row=oi, column=i)
                        for i in range(move_count):  # if there are moves, create + populate cells
                            Label(f, text=format(float(s_moves[count]), '.2f'), justify=LEFT, width=macadj(6, 4),
                                  relief=RIDGE, bg=in_color).grid(row=oi, column=4)  # move off
                            count += 1
                            Label(f, text=format(float(s_moves[count]), '.2f'), justify=LEFT, width=macadj(6, 4),
                                  relief=RIDGE, bg=in_color).grid(row=oi, column=5)  # move on
                            count += 1
                            Label(f, text=s_moves[count], justify=LEFT, width=macadj(6, 5),
                                  relief=RIDGE, bg=in_color).grid(row=oi, column=6)  # route
                            count += 1
                            Label(f, text=format(move_totals[i], '.2f'), justify=LEFT, width=macadj(6, 4),
                                  relief=RIDGE, bg=out_color).grid(row=oi, column=7)  # move total
                            if i < move_count - 1:
                                oi += 1
                        Label(f, text=format(ot, '.2f'), justify=LEFT, width=macadj(6, 4),
                              relief=RIDGE, bg=out_color) \
                            .grid(row=oi, column=8)  # overtime
                        Label(f, text=format(off_route, '.2f'), justify=LEFT, width=macadj(6, 4),
                              relief=RIDGE, bg=out_color) \
                            .grid(row=oi, column=9)  # off route
                        Label(f, text=format(ot_off_route, '.2f'), justify=LEFT, width=macadj(6, 4),
                              relief=RIDGE, bg=out_color) \
                            .grid(row=oi, column=10)  # OT off route
                        oi += 1
            # if there is no match, then just printe the name.
            if match == "miss":
                Label(f, text=line[1], anchor=W, width=macadj(21, 16), relief=RIDGE, bg=in_color) \
                    .grid(row=oi, column=0)  # name
                for i in range(10):
                    if i < 6:
                        color = in_color
                    else:
                        color = out_color
                    if i == 5:
                        ml = 5
                    else:
                        ml = 4
                    Label(f, text="", width=macadj(6, ml), relief=RIDGE, bg=color) \
                        .grid(row=oi, column=i + 1)  # generate blank cells
                oi += 1
        oi += 1
        Label(f, text="", height=2, bg="white").grid(row=oi, column=0)
        Label(f, text=format(ot_total, '.2f'), justify=LEFT, width=macadj(6, 4), relief=RIDGE, bg=out_color) \
            .grid(row=oi, column=8)  # overtime
        Label(f, text=format(ot_off_total, '.2f'), justify=LEFT, width=macadj(6, 4), relief=RIDGE, bg=out_color) \
            .grid(row=oi, column=10)  # OT off route
        oi += 2
        #  overtime desired list xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
        Label(f, text="overtime desired list", justify=LEFT,
              font=macadj('Helvetica 10 bold', 'Futura 16 bold'), bg="white") \
            .grid(sticky=W, row=oi, column=0, columnspan=10)
        oi += 1
        Label(f, text="Availability to:", bg="white") \
            .grid(row=oi, column=4, columnspan=macadj(3, 3), sticky=W)
        oi += 1
        Label(f, text="Carrier", bg="white").grid(row=oi, column=0, sticky=W)
        Label(f, text="note", bg="white").grid(row=oi, column=1)
        Label(f, text="5200", bg="white").grid(row=oi, column=2)
        Label(f, text="RS", bg="white").grid(row=oi, column=3)
        Label(f, text="10", bg="white").grid(row=oi, column=4)
        Label(f, text="12", bg="white").grid(row=oi, column=5)
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
                        Label(f, text=each[1], anchor=W, width=macadj(21, 16), relief=RIDGE, bg=in_color) \
                            .grid(row=oi, column=0)  # name
                        if each[4] == "none":
                            code = ""
                        else:
                            code = each[4]
                        Label(f, text=code, justify=LEFT, width=macadj(6, 4), relief=RIDGE, bg=in_color) \
                            .grid(row=oi, column=1)  # code
                        if each[2] == "" or each[2] == " ":  # handle empty 5200 strings
                            t_hrs = ""
                        else:
                            t_hrs = format(float(each[2]), '.2f')
                        Label(f, text=t_hrs, justify=LEFT, width=macadj(6, 4), relief=RIDGE, bg=in_color) \
                            .grid(row=oi, column=2)  # 5200
                        if each[3] == "" or each[3] == " ":  # handle empty RS strings
                            rs = ""
                        else:
                            rs = format(float(each[3]), '.2f')
                        Label(f, text=rs, justify=LEFT, width=macadj(6, 4), relief=RIDGE, bg=in_color) \
                            .grid(row=oi, column=3)  # return to station
                        Label(f, text=format(float(aval_10), '.2f'), justify=LEFT, width=macadj(6, 4),
                              relief=RIDGE, bg=out_color) \
                            .grid(row=oi, column=4)  # availability to 10
                        Label(f, text=format(float(aval_12), '.2f'), justify=LEFT, width=macadj(6, 4),
                              relief=RIDGE, bg=out_color) \
                            .grid(row=oi, column=5)  # availability to 12
                        oi += 1
                    # if there is no match, then just printe the name.
            if match == "miss":
                Label(f, text=line[1], anchor=W, width=macadj(21, 16), relief=RIDGE, bg=in_color) \
                    .grid(row=oi, column=0)  # name
                for i in range(5):
                    if i < 3:
                        color = in_color
                    else:
                        color = out_color
                    Label(f, text="", width=macadj(6, 4), relief=RIDGE, bg=color) \
                        .grid(row=oi, column=i + 1)  # generate blank cells
                oi += 1
        oi += 1
        Label(f, text="", height=2, bg="white").grid(row=oi, column=0)
        Label(f, text=format(aval_10_total, '.2f'), justify=LEFT, width=macadj(6, 4), relief=RIDGE, bg=out_color) \
            .grid(row=oi, column=4)  # availability to 10 total
        Label(f, text=format(aval_12_total, '.2f'), justify=LEFT, width=macadj(6, 4), relief=RIDGE, bg=out_color) \
            .grid(row=oi, column=5)  # availability to 12 total
        oi += 2
        # auxiliary assistance xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
        Label(f, text="auxiliary assistance", justify=LEFT,
              font=macadj('Helvetica 10 bold', 'Futura 16 bold'), bg="white") \
            .grid(sticky=W, row=oi, column=0, columnspan=10)
        oi += 1
        Label(f, text="Availability to:", bg="white").grid(row=oi, column=4, columnspan=macadj(3, 3), sticky=W)
        oi += 1
        Label(f, text="Carrier", bg="white").grid(row=oi, column=0, sticky=W)
        Label(f, text="note", bg="white").grid(row=oi, column=1)
        Label(f, text="5200", bg="white").grid(row=oi, column=2)
        Label(f, text="RS", bg="white").grid(row=oi, column=3)
        Label(f, text="10", bg="white").grid(row=oi, column=4)
        Label(f, text="11.5", bg="white").grid(row=oi, column=5)
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
                        Label(f, text=each[1], anchor=W, width=macadj(21, 16), relief=RIDGE, bg=in_color) \
                            .grid(row=oi, column=0)  # name
                        if each[4] == "none":
                            code = ""
                        else:
                            code = each[4]
                        Label(f, text=code, justify=LEFT, width=macadj(6, 4), relief=RIDGE, bg=in_color) \
                            .grid(row=oi, column=1)  # code
                        if each[2] == "" or each[2] == " ":  # handle empty 5200 strings
                            t_hrs = ""
                        else:
                            t_hrs = format(float(each[2]), '.2f')
                        Label(f, text=t_hrs, justify=LEFT, width=macadj(6, 4), relief=RIDGE, bg=in_color) \
                            .grid(row=oi, column=2)  # 5200
                        if each[3] == "" or each[3] == " ":  # handle empty RS strings
                            rs = ""
                        else:
                            rs = format(float(each[3]), '.2f')
                        Label(f, text=rs, justify=LEFT, width=macadj(6, 4), relief=RIDGE, bg=in_color) \
                            .grid(row=oi, column=3)  # return to station
                        Label(f, text=format(float(aval_10), '.2f'), justify=LEFT, width=macadj(6, 4),
                              relief=RIDGE, bg=out_color) \
                            .grid(row=oi, column=4)  # availability to 10
                        Label(f, text=format(float(aval_115), '.2f'), justify=LEFT, width=macadj(6, 4),
                              relief=RIDGE, bg=out_color) \
                            .grid(row=oi, column=5)  # availability to 12
                        oi += 1
            # if there is no match, then just printe the name.
            if match == "miss":
                Label(f, text=line[1], anchor=W, width=macadj(21, 16), relief=RIDGE, bg=in_color) \
                    .grid(row=oi, column=0)  # name
                for i in range(5):
                    if i < 3:
                        color = in_color
                    else:
                        color = out_color
                    Label(f, text="", width=macadj(6, 4), relief=RIDGE, bg=color) \
                        .grid(row=oi, column=i + 1)  # generate blank cells
            oi += 1
        oi += 1
        Label(f, text="", height=2, bg="white").grid(row=oi, column=0)
        Label(f, text=format(aval_10_total, '.2f'), justify=LEFT, width=macadj(6, 4), relief=RIDGE, bg=out_color) \
            .grid(row=oi, column=4)  # availability to 10 total
        Label(f, text=format(aval_115_total, '.2f'), justify=LEFT, width=macadj(6, 4), relief=RIDGE, bg=out_color) \
            .grid(row=oi, column=5)  # availability to 11.5 total
        oi += 2
        t += 1  # t increaments tabs
    projvar.root.mainloop()


class EnterRings:
    def __init__(self, carrier):
        self.frame = None
        self.origin_frame = None  # defunct
        self.win = None
        self.carrier = carrier
        self.carrecs = []  # get the carrier rec set
        self.ringrecs = []  # get the rings for the week
        self.dates = []  # get a datetime object for each day in the investigation range
        self.daily_carrecs = []  # get the carrier record for each day
        self.daily_ringrecs = []  # get the rings record for each day
        self.totals = []  # arrays holding stringvars
        self.rss = []
        self.moves = []
        self.codes = []
        self.lvtypes = []
        self.lvtimes = []
        self.now_moves = ""  # default values of the stringvars
        self.sat_mm = []  # holds daily stringvars for moves
        self.sun_mm = []
        self.mon_mm = []
        self.tue_mm = []
        self.wed_mm = []
        self.thu_mm = []
        self.fri_mm = []
        self.move_string = ""
        self.ot_rings_limiter = None
        self.chg_these = []
        self.addrings = []
        if projvar.invran_weekly_span:
            for i in range(7):
                self.addrings.append([])
        self.status_update = ""
        self.delete_report = 0
        self.update_report = 0
        self.insert_report = 0
        self.day = ("Saturday", "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday")

    def start(self, frame):
        self.frame = frame
        self.win = MakeWindow()
        self.win.create(self.frame)
        self.re_initialize()
        self.get_carrecs()
        self.get_ringrecs()
        self.get_dates()
        self.get_daily_carrecs()
        self.get_daily_ringrecs()
        self.get_rings_limiter()
        self.build_page()
        self.write_report()
        self.buttons_frame()
        self.zero_report_vars()
        self.win.finish()

    def re_initialize(self):
        self.carrecs = []  # get the carrier rec set
        self.ringrecs = []  # get the rings for the week
        self.dates = []  # get a datetime object for each day in the investigation range
        self.daily_carrecs = []  # get the carrier record for each day
        self.daily_ringrecs = []  # get the rings record for each day
        self.totals = []  # arrays holding stringvars
        self.rss = []
        self.moves = []
        self.codes = []
        self.lvtypes = []
        self.lvtimes = []
        self.now_moves = ""  # default values of the stringvars
        self.sat_mm = []  # holds daily stringvars for moves
        self.sun_mm = []
        self.mon_mm = []
        self.tue_mm = []
        self.wed_mm = []
        self.thu_mm = []
        self.fri_mm = []
        self.move_string = ""
        self.chg_these = []
        self.addrings = []
        if projvar.invran_weekly_span:
            for i in range(7):
                self.addrings.append([])

    def get_carrecs(self):  # get the carrier's carrier rec set
        if projvar.invran_weekly_span:
            self.carrecs = CarrierRecSet(self.carrier, projvar.invran_date_week[0], projvar.invran_date_week[6],
                                         projvar.invran_station).get()
        else:
            self.carrecs = CarrierRecSet(self.carrier, projvar.invran_date, projvar.invran_date,
                                         projvar.invran_station).get()

    def get_ringrecs(self):  # get the ring recs for the invran
        if projvar.invran_weekly_span:
            self.ringrecs = Rings(self.carrier, projvar.invran_date).get_for_week()
        else:
            self.ringrecs = Rings(self.carrier, projvar.invran_date).get_for_day()

    def get_dates(self):  # get a datetime object for each day in the investigation range
        if projvar.invran_weekly_span:
            self.dates = projvar.invran_date_week
        else:
            self.dates = [projvar.invran_date, ]

    def get_daily_carrecs(self):  # make a list of carrecs for each day
        for d in self.dates:
            for rec in self.carrecs:
                if rec[0] <= str(d):
                    self.daily_carrecs.append(rec)
                    break

    def get_daily_ringrecs(self):  # make list of ringrecs for each day, insert empty rec if there is no rec
        match = False
        for d in self.dates:  # for each day in self.dates
            for rr in self.ringrecs:
                if rr:  # if there is a ring rec
                    if rr[0] == str(d):  # when the dates match
                        self.daily_ringrecs.append(list(rr))  # creates the daily_ringrecs array
                        match = True
            if not match:  # if there is no match
                add_this = [d, self.carrier, "", "", "none", "", "none", ""]  # insert an empty record
                self.daily_ringrecs.append(add_this)  # creates the daily_ringrecs array
            match = False
        # convert the time item from string to datetime object
        for i in range(len(self.daily_ringrecs)):
            if type(self.daily_ringrecs[i][0]) == str:
                self.daily_ringrecs[i][0] = Convert(self.daily_ringrecs[i][0]).dt_converter()

    def get_rings_limiter(self):
        sql = "SELECT tolerance FROM tolerances WHERE category = '%s'" % "ot_rings_limiter"
        results = inquire(sql)
        self.ot_rings_limiter = int(results[0][0])

    def build_page(self):
        now_total = None
        now_rs = None
        now_code = None
        now_moves = None
        now_lv_type = None
        now_lv_time = None
        day = ("sat", "sun", "mon", "tue", "wed", "thr", "fri")
        frame = ["F0", "F1", "F2", "F3", "F4", "F5", "F6"]
        color = ["red", "light blue", "yellow", "green", "brown", "gold", "purple", "grey", "light grey"]
        nolist_codes = ("none", "ns day")
        ot_codes = ("none", "ns day", "no call", "light", "sch chg", "annual", "sick", "excused")
        aux_codes = ("none", "no call", "light", "sch chg", "annual", "sick", "excused")
        lv_options = ("none", "annual", "sick", "holiday", "other", "combo")
        option_menu = ["om0", "om1", "om2", "om3", "om4", "om5", "om6"]
        lv_option_menu = ["lom0", "lom1", "lom2", "lom3", "lom4", "lom5", "lom6"]
        total_widget = ["tw0", "tw1", "tw2", "tw3", "tw4", "tw5", "tw6"]
        frame_i = 0  # counter for the frame
        header_frame = Frame(self.win.body, width=500)  # header  frame
        header_frame.grid(row=frame_i, padx=5, sticky="w")
        # Header at top of window: name
        Label(header_frame, text="carrier name: ", fg="Grey", font=macadj("bold", "Helvetica 18")) \
            .grid(row=0, column=0, sticky="w")
        Label(header_frame, text="{}".format(self.carrecs[0][1]), font=macadj("bold", "Helvetica 18")) \
            .grid(row=0, column=1, sticky="w")
        Label(header_frame, text="list status: {}".format(self.carrecs[0][2])) \
            .grid(row=1, sticky="w", columnspan=2)
        if self.carrecs[0][4] != "":
            Label(header_frame, text="route/s: {}".format(self.carrecs[0][4])) \
                .grid(row=2, sticky="w", columnspan=2)
        frame_i += 2
        if projvar.invran_weekly_span:  # if investigation range is weekly
            i_range = 7  # loop 7 times for week or once for day
        else:
            i_range = 1
        for i in range(i_range):
            # for ring in self.daily_ringrecs:  # assign the values for each rings attribute
            now_total = Convert(self.daily_ringrecs[i][2]).empty_not_zero()
            now_rs = Convert(self.daily_ringrecs[i][3]).empty_not_zero()
            now_code = Convert(self.daily_ringrecs[i][4]).none_not_empty()
            self.now_moves = self.daily_ringrecs[i][5]
            now_lv_type = Convert(self.daily_ringrecs[i][6]).none_not_empty()
            now_lv_time = Convert(self.daily_ringrecs[i][7]).empty_not_zero()
            grid_i = 0  # counter for the grid within the frame
            frame[i] = Frame(self.win.body, width=500)
            frame[i].grid(row=frame_i, padx=5, sticky="w")
            # Display the day and date
            if projvar.ns_code[self.carrecs[0][3]] == self.dates[i].strftime("%a"):
                Label(frame[i], text="{} NS DAY".format(self.dates[i].strftime("%a %b %d, %Y")), fg="red") \
                    .grid(row=grid_i, column=0, columnspan=5, sticky="w")
            else:
                Label(frame[i], text=self.dates[i].strftime("%a %b %d, %Y"), fg="blue") \
                    .grid(row=grid_i, column=0, columnspan=5, sticky="w")
            grid_i += 1
            column = 6  # if the ot rings limiter is off/false - column = 6
            if self.daily_carrecs[i][2] in ("aux", "ptf"):  # don't show moves for aux or ptf carriers
                column = 3
            elif self.daily_carrecs[i][2] in ("otdl", ):  # show moves for otdl unless ot rings limiter is on
                if self.ot_rings_limiter:  # if ot rings limiter is on/true - colummn = 3
                    column = 3
            if self.daily_carrecs[i][5] == projvar.invran_station:
                Label(frame[i], text="5200", fg=color[7]).grid(row=grid_i, column=0)  # Display all labels
                Label(frame[i], text="RS", fg=color[7]).grid(row=grid_i, column=1)
                if column == 6:  # don't show moves for aux, ptf and (maybe) otdl
                    Label(frame[i], text="MV off", fg=color[7]).grid(row=grid_i, column=2)
                    Label(frame[i], text="MV on", fg=color[7]).grid(row=grid_i, column=3)
                    Label(frame[i], text="Route", fg=color[7]).grid(row=grid_i, column=4)
                Label(frame[i], text="code", fg=color[7]).grid(row=grid_i, column=column)
                Label(frame[i], text="LV type", fg=color[7]).grid(row=grid_i, column=column+1)
                Label(frame[i], text="LV time", fg=color[7]).grid(row=grid_i, column=column+2)
                grid_i += 1
                # Display the entry widgets
                # 5200 time
                self.totals.append(StringVar(frame[i]))  # append stringvar to totals array
                total_widget[i] = Entry(frame[i], width=macadj(8, 4), textvariable=self.totals[i])
                total_widget[i].grid(row=grid_i, column=0)
                self.totals[i].set(now_total)  # set the starting value for total
                # Return to Station (rs)
                self.rss.append(StringVar(frame[i]))  # RS entry widget
                Entry(frame[i], width=macadj(8, 4), textvariable=self.rss[i]).grid(row=grid_i, column=1)
                self.rss[i].set(now_rs)  # set the starting value for RS
                # Moves
                if column == 6:  # don't show moves for aux, ptf and (maybe) otdl
                    self.new_entry(frame[i], day[i])  # MOVES on, off and route entry widgets
                    Button(frame[i], text="more moves", command=lambda x=i: self.new_entry(frame[x], day[x])) \
                        .grid(row=grid_i, column=5)
                self.now_moves = ""  # zero out self.now_moves so more moves button works properly
                # Codes/Notes
                self.codes.append(StringVar(frame[i]))  # code entry widget
                if self.daily_carrecs[i][2] == "wal" or self.daily_carrecs[i][2] == "nl":
                    option_menu[i] = OptionMenu(frame[i], self.codes[i], *nolist_codes)
                elif self.daily_carrecs[i][2] == "otdl":
                    option_menu[i] = OptionMenu(frame[i], self.codes[i], *ot_codes)
                else:
                    option_menu[i] = OptionMenu(frame[i], self.codes[i], *aux_codes)
                self.codes[i].set(now_code)
                option_menu[i].configure(width=macadj(7, 6))
                option_menu[i].grid(row=grid_i, column=column)  # code widget
                # Leave Type
                self.lvtypes.append(StringVar(frame[i]))  # leave type entry widget
                lv_option_menu[i] = OptionMenu(frame[i], self.lvtypes[i], *lv_options)
                lv_option_menu[i].configure(width=macadj(7, 6))
                lv_option_menu[i].grid(row=grid_i, column=column+1)  # leave type widget
                # Leave Time
                self.lvtimes.append(StringVar(frame[i]))  # leave time entry widget
                self.lvtypes[i].set(now_lv_type)  # set the starting value for leave type
                self.lvtimes[i].set(now_lv_time)  # set the starting value for leave type
                Entry(frame[i], width=macadj(8, 4), textvariable=self.lvtimes[i]) \
                    .grid(row=grid_i, column=column+2)  # leave time widget
            else:
                self.totals.append(StringVar(frame[i]))  # 5200 entry widget
                self.rss.append(StringVar(frame[i]))  # RS entry
                if self.daily_carrecs[i][5] != "no record":  # display for records that are out of station
                    Label(frame[i], text="out of station: {}".format(self.daily_carrecs[i][5]),
                          fg="white", bg="grey", width=55, height=2, anchor="w").grid(row=grid_i, column=0)
                else:  # display for when there is no record relevant for that day.
                    Label(frame[i], text="no record", fg="white", bg="grey", width=55, height=2, anchor="w")\
                        .grid(row=grid_i, column=0)
            frame_i += 1
        f7 = Frame(self.win.body)
        f7.grid(row=frame_i)
        Label(f7, height=50).grid(row=1, column=0)  # extra white space on bottom of form to facilitate moves

    @staticmethod
    def triad_row_finder(index):  # finds the row of the moves entry widget or button
        if index % 3 == 0:
            return int(index / 3)
        elif (index - 1) % 3 == 0:
            return int((index - 1) / 3)
        elif (index - 2) % 3 == 0:
            return int((index - 2) / 3)

    @staticmethod
    def triad_col_finder(index):  # finds the column of the moves widget
        if index % 3 == 0:  # first column
            return int(0)
        elif (index - 1) % 3 == 0:  # second column
            return int(1)
        elif (index - 2) % 3 == 0:  # third column
            return int(2)

    def new_entry(self, frame, day):  # creates new entry fields for "more move functionality"
        mm = []
        if day == "sat":
            mm = self.sat_mm  # find the day in question and use the correlating  array
        elif day == "sun":
            mm = self.sun_mm
        elif day == "mon":
            mm = self.mon_mm
        elif day == "tue":
            mm = self.tue_mm
        elif day == "wed":
            mm = self.wed_mm
        elif day == "thr":
            mm = self.thu_mm
        elif day == "fri":
            mm = self.fri_mm
        # what to do depending on the moves
        if self.now_moves == "":  # if there are no moves sent to the function
            mm.append(StringVar(frame))  # create first entry field for new entries
            Entry(frame, width=macadj(8, 4), textvariable=mm[len(mm) - 1]) \
                .grid(row=self.triad_row_finder(len(mm) - 1) + 2,
                      column=self.triad_col_finder(len(mm) - 1) + 2)  # route
            mm.append(StringVar(frame))  # create second entry field for new entries
            Entry(frame, width=macadj(8, 4), textvariable=mm[len(mm) - 1]) \
                .grid(row=self.triad_row_finder(len(mm) - 1) + 2,
                      column=self.triad_col_finder(len(mm) - 1) + 2)  # move off
            mm.append(StringVar(frame))  # create second entry field for new entries
            Entry(frame, width=macadj(8, 5), textvariable=mm[len(mm) - 1]) \
                .grid(row=self.triad_row_finder(len(mm) - 1) + 2,
                      column=self.triad_col_finder(len(mm) - 1) + 2)  # move on
        else:  # if there are moves which need to be set
            moves = self.now_moves.split(",")  # turn now_moves into an array
            iterations = len(moves)  # get the number of items in moves array
            for i in range(int(iterations)):  # loop through all items in moves array
                mm.append(StringVar(frame))  # create entry field for moves from database
                mm[i].set(moves[i])  # set values for the StringVars
                if (i + 1) % 3 == 0:  # adjust the lenght of the route widget pending os
                    ml = 5  # on mac, the route widget lenght is 5
                else:
                    ml = 4  # on mac, the rings widget lenght is 4
                # build the widget
                Entry(frame, width=macadj(8, ml), textvariable=mm[i]) \
                    .grid(row=self.triad_row_finder(i) + 2, column=self.triad_col_finder(i) + 2)

    def write_report(self):  # build the report to appear on bottom of screen
        if not self.status_update:
            return
        if self.delete_report + self.update_report + self.insert_report == 0:
            self.status_update = "No records changed. "  # if there are no changes
            return
        status_update = ""
        if self.insert_report:  # new records
            status_update += str(self.insert_report) + " new record{} added. "\
                .format(Handler(self.insert_report).plurals())  # make "record" plural if necessary
        if self.update_report:  # updated records
            status_update += str(self.update_report) + " record{} updated. "\
                .format(Handler(self.update_report).plurals())  # make "record" plural if necessary
        if self.delete_report:  # deleted records
            status_update += str(self.delete_report) + " record{} deleted. "\
                .format(Handler(self.delete_report).plurals())  # make "record" plural if necessary
        self.status_update = status_update

    def buttons_frame(self):
        Button(self.win.buttons, text="Submit", width=10, anchor="w",
               command=lambda: self.apply_rings(True)).pack(side=LEFT)
        Button(self.win.buttons, text= "Apply", width=10, anchor="w",
               command=lambda: self.apply_rings(False)).pack(side=LEFT)
        Button(self.win.buttons, text="Go Back", width=10, anchor="w",
               command=lambda: MainFrame().start(frame=self.win.topframe)).pack(side=LEFT)
        Label(self.win.buttons, text="{}".format(self.status_update), fg="red").pack(side=LEFT)

    def zero_report_vars(self):
        self.status_update = "No records changed."
        self.delete_report = 0
        self.update_report = 0
        self.insert_report = 0

    def apply_rings(self, go_home):
        self.empty_addrings()
        self.add_date()
        if not self.check_5200():
            return  # abort if there is an error
        if not self.check_rs():
            return  # abort if there is an error
        self.add_codes()
        if not self.check_moves():
            return  # abort if there is an error
        self.add_leavetype()
        if not self.check_leave():
            return  # abort if there is an error
        self.addrecs()  # insert rings into the database
        if go_home:  # if True, then exit screen to main screen
            MainFrame().start(frame=self.win.topframe)
        else:  # if False, then rebuild the Enter Rings screen
            self.start(self.win.topframe)

    def empty_addrings(self):  # empty out addring arrays
        for i in range(len(self.addrings)):
            self.addrings[i] = []

    def add_date(self):  # start the addrings array
        for i in range(len(self.dates)):  # loop for each day in the investigation
            self.addrings[i].append(self.dates[i])  # add the date
            self.addrings[i].append(self.carrier)  # add the carrier name

    def check_5200(self):
        for i in range(len(self.totals)):
            total = self.totals[i].get().strip()
            if RingTimeChecker(total).check_for_zeros():
                self.addrings[i].append("")  # if variable is zero or empty, add an empty string to addrings
                continue  # skip other checks
            if not RingTimeChecker(total).check_numeric():
                text = "You must enter a numeric value in 5200 for {}.".format(self.day[i])
                messagebox.showerror("5200 Error", text, parent=self.win.topframe)
                return False
            if not RingTimeChecker(total).over_24():
                text = "Values greater than 24 are not accepted in 5200 for {}.".format(self.day[i])
                messagebox.showerror("5200 Error", text, parent=self.win.topframe)
                return False
            if not RingTimeChecker(total).less_than_zero():
                text = "Values less than or equal to 0 are not accepted in 5200 for {}.".format(self.day[i])
                messagebox.showerror("5200 Error", text, parent=self.win.topframe)
                return False
            if not RingTimeChecker(total).count_decimals_place():
                text = "Values with more than 2 decimal places are not accepted in 5200 for {}.".format(self.day[i])
                messagebox.showerror("5200 Error", text, parent=self.win.topframe)
                return False
            total = Convert(total).hundredths()  # format it as a number with 2 decimal places
            self.addrings[i].append(total)  # if all checks pass, add to addrings
        return True

    def check_rs(self):
        for i in range(len(self.rss)):
            rs = str(self.rss[i].get()).strip()
            if RingTimeChecker(rs).check_for_zeros():
                self.addrings[i].append("")  # if variable is zero or empty, add an empty string to addrings
                continue  # skip other checks
            if not RingTimeChecker(rs).check_numeric():
                text = "You must enter a numeric value in RS for {}.".format(self.day[i])
                messagebox.showerror("RS Error", text, parent=self.win.topframe)
                return False
            if not RingTimeChecker(rs).over_24():
                text = "Values greater than 24 are not accepted in RS for {}.".format(self.day[i])
                messagebox.showerror("RS Error", text, parent=self.win.topframe)
                return False
            if not RingTimeChecker(rs).less_than_zero():
                text = "Values less than or equal to 0 are not accepted in RS for {}.".format(self.day[i])
                messagebox.showerror("RS Error", text, parent=self.win.topframe)
                return False
            if not RingTimeChecker(rs).count_decimals_place():
                text = "Values with more than 2 decimal places are not accepted in RS for {}.".format(self.day[i])
                messagebox.showerror("RS Error", text, parent=self.win.topframe)
                return False
            rs = Convert(rs).hundredths()  # format it as a number with 2 decimal places
            self.addrings[i].append(rs)  # if all checks pass, add to addrings
        return True

    def add_codes(self):
        for i in range(len(self.codes)):
            self.addrings[i].append(self.codes[i].get())

    def bypass_moves(self):  # keep existing moves if otdl rings limiter is on/True
        if projvar.invran_weekly_span:  # if investigation range is weekly
            i_range = 7  # investigation range is seven days
        else:
            i_range = 1  # investigation range is one day
        for i in range(i_range):  # loop for each day in investigation
            moves = self.daily_ringrecs[i][5]  # get the preexisting record for that day
            self.addrings[i].append(moves)  # add that record to addrings array

    def move_string_constructor(self, first, second, third):
        if self.move_string and first and second:
            self.move_string += ","
        if first and second:
            self.move_string += first + "," + second + "," + third

    def check_moves(self):
        if self.ot_rings_limiter:  # if the otdl rings limiter is on/True
            self.bypass_moves()  # bypass all checks and put preexisting moves into addrings
            return True  # mission accomplished
        first_move = None
        second_move = None
        route = None
        days = (self.sat_mm, self.sun_mm, self.mon_mm, self.tue_mm, self.wed_mm, self.thu_mm, self.fri_mm)
        cc = 0  # increments one for each day
        for d in days:  # check for bad inputs in moves
            self.move_string = ""  # emtpy out string where moves data is passed
            x = len(d)
            for i in range(x):
                if self.triad_col_finder(i) == 0:  # find the first of the triad
                    first_move = d[i].get().strip()
                    second_move = d[i + 1].get().strip()
                    if MovesChecker(first_move).check_for_zeros() or MovesChecker(second_move).check_for_zeros():
                        if MovesChecker(first_move).check_for_zeros() and \
                                MovesChecker(second_move).check_for_zeros():  # if both are zeros
                            continue  # skip the rest of the checks
                        text = "You must provide two values on moves for {}.".format(self.day[cc])
                        messagebox.showerror("Move entry error", text, parent=self.win.topframe)
                        return False
                    if not RingTimeChecker(first_move).check_numeric() or \
                            not RingTimeChecker(second_move).check_numeric():
                        text = "You must enter a numeric value on moves for {}.".format(self.day[cc])
                        messagebox.showerror("Move entry error", text, parent=self.win.topframe)
                        return False
                    if not MovesChecker(first_move).compare(second_move):
                        text = "The earlier value can not be greater than the later value on moves for {}." \
                            .format(self.day[cc])
                        messagebox.showerror("Move entry error", text, parent=self.win.topframe)
                        return False
                    if not RingTimeChecker(first_move).over_24() or not RingTimeChecker(second_move).over_24():
                        text = "Values greater than 24 are not accepted on moves for {}.".format(self.day[cc])
                        messagebox.showerror("Move entry error", text, parent=self.win.topframe)
                        return False
                    if not RingTimeChecker(first_move).less_than_zero() or \
                            not RingTimeChecker(second_move).less_than_zero():
                        text = "Values less than 0 are not accepted on moves for {}.".format(self.day[cc])
                        messagebox.showerror("Move entry error", text, parent=self.win.topframe)
                        return False
                    if not RingTimeChecker(first_move).count_decimals_place() or \
                            not RingTimeChecker(second_move).count_decimals_place():
                        text = "Moves can not have more than two decimal places on moves for {}.".format(self.day[cc])
                        messagebox.showerror("Move entry error", text, parent=self.win.topframe)
                        return False
                    first_move = Convert(first_move).hundredths()
                    second_move = Convert(second_move).hundredths()
                if self.triad_col_finder(i) == 2:  # find the third of the triad
                    route = d[i].get().strip()
                    if RouteChecker(route).is_empty():  # if the route is an empty string
                        self.move_string_constructor(first_move, second_move, "")
                        continue  # skip the rest of the checks
                    if not RouteChecker(route).check_numeric():
                        text = "You must enter a numeric value on route for {}.".format(self.day[cc])
                        messagebox.showerror("Move entry error", text, parent=self.win.topframe)
                        return False
                    if not RouteChecker(route).only_one():
                        text = "Only one route is allowed in route field on moves for {}.".format(self.day[cc])
                        messagebox.showerror("Move entry error", text, parent=self.win.topframe)
                        return False
                    if RouteChecker(route).only_numbers():
                        text = "Only numbers are allowed in route field on moves for {}.".format(self.day[cc])
                        messagebox.showerror("Move entry error", text, parent=self.win.topframe)
                        return False
                    if not RouteChecker(route).check_length():
                        text = "The route number for {} must be four or five digits long.".format(self.day[cc])
                        messagebox.showerror("Move entry error", text, parent=self.win.topframe)
                        return False
                    self.move_string_constructor(first_move, second_move, route)
            self.addrings[cc].append(self.move_string)
            cc += 1
        return True

    def add_leavetype(self):
        for i in range(len(self.lvtypes)):
            self.addrings[i].append(self.lvtypes[i].get())

    def check_leave(self):
        for i in range(len(self.lvtimes)):
            lvtime = str(self.lvtimes[i].get()).strip()
            if RingTimeChecker(lvtime).check_for_zeros():
                self.addrings[i].append("")  # if variable is zero or empty, add an empty string to addrings
                continue  # skip other checks
            if not RingTimeChecker(lvtime).check_numeric():
                text = "You must enter a numeric value in Leave Time for {}.".format(self.day[i])
                messagebox.showerror("Leave Time Error", text, parent=self.win.topframe)
                return False
            if not RingTimeChecker(lvtime).over_8():
                text = "Values greater than 8 are not accepted in Leave Time for {}.".format(self.day[i])
                messagebox.showerror("Leave Time Error", text, parent=self.win.topframe)
                return False
            if not RingTimeChecker(lvtime).less_than_zero():
                text = "Values less than or equal to 0 are not accepted in Leave Time for {}.".format(self.day[i])
                messagebox.showerror("Leave Time Error", text, parent=self.win.topframe)
                return False
            if not RingTimeChecker(lvtime).count_decimals_place():
                text = "Values with more than 2 decimal places are not accepted in Leave Time for {}."\
                    .format(self.day[i])
                messagebox.showerror("Leave Time Error", text, parent=self.win.topframe)
                return False
            # lvtime = format(float(lvtime), '.2f')  # format it as a float with 2 decimal places
            lvtime = Convert(lvtime).hundredths()  # format it as a number with 2 decimal places
            self.addrings[i].append(lvtime)  # if all checks pass, add to addrings
        return True

    def addrecs(self):  # add records to database
        sql = ""
        for i in range(len(self.dates)):
            empty_rec = [self.dates[i], self.carrier, "", "", "none", "", "none", ""]
            if self.addrings[i] == self.daily_ringrecs[i]:
                sql = ""  # if new and old are a match, take no action
            elif not self.addrings[i][2] and not self.addrings[i][7] and self.addrings[i][4] != "no call" \
                    and self.daily_ringrecs[i] == empty_rec:
                sql = ""  # if old is empty and new is not qualified as a legit record, take no action
            elif not self.addrings[i][2] and not self.addrings[i][7] and self.addrings[i][4] != "no call":
                # if new record has no total or lvtime
                sql = "DELETE FROM rings3 WHERE rings_date = '%s' and carrier_name = '%s'" \
                      % (self.dates[i], self.carrier)
                self.delete_report += 1
            elif self.daily_ringrecs[i] != empty_rec and self.addrings[i] != empty_rec:
                # if a record exist but is different from the new record
                sql = "UPDATE rings3 SET total='%s',rs='%s',code='%s',moves='%s',leave_type = '%s'," \
                      "leave_time = '%s' WHERE rings_date = '%s' and carrier_name = '%s'" \
                      % (self.addrings[i][2], self.addrings[i][3], self.addrings[i][4],
                         self.addrings[i][5], self.addrings[i][6], self.addrings[i][7],
                         self.dates[i], self.carrier)
                self.update_report += 1
            elif self.daily_ringrecs[i] == empty_rec and self.addrings[i] != empty_rec:
                # if a record doesn't exist and the new record is not empty
                sql = "INSERT INTO rings3 (rings_date, carrier_name, total, rs, code, moves, leave_type, leave_time )" \
                      "VALUES('%s','%s','%s','%s','%s','%s','%s','%s') " \
                      % (self.dates[i], self.carrier, self.addrings[i][2], self.addrings[i][3],
                         self.addrings[i][4], self.addrings[i][5], self.addrings[i][6], self.addrings[i][7])
                self.insert_report += 1
            if sql:
                commit(sql)


def apply_update_carrier(year, month, day, name, ls, ns, route, station, rowid, frame):
    if year.get() > 9999:
        messagebox.showerror("Year Input Error", "Year must be between 1 and 9999", parent=frame)
        return
    if year.get() < 1:
        messagebox.showerror("Year Input Error", "Year must be between 1 and 9999", parent=frame)
        return
    try:
        date = datetime(year.get(), month.get(), day.get())
    except ValueError:
        messagebox.showerror("Invalid Date", "Date entered is not valid", parent=frame)
        return
    route_list = route.get().split("/")
    if len(route.get()) > 29:
        messagebox.showerror("Route number input error",
                             "There can be no more than five routes per carrier "
                             "(for T6 carriers).\n Routes numbers must be 4 or 5 digits long.\n"
                             "If there are multiple routes, route numbers must be separated by "
                             "the \'/\' character. For example: 1001/1015/10124/10224/0972. Do not use "
                             "commas or empty spaces",
                             parent=frame)
        return
    for item in route_list:
        item = item.strip()
        if item != "":
            if len(item) < 4 or len(item) > 5:
                messagebox.showerror("Route number input error",
                                     'Routes numbers must be four or five digits long.\n'
                                     'If there are multiple routes, route numbers must be separated by '
                                     'the \'/\' character. For example: 1001/1015/10124/10224/0972. Do not use '
                                     'commas or empty spaces',
                                     parent=frame)
                return
        if item.isdigit() == FALSE and item != "":
            messagebox.showerror("Route number input error",
                                 "Route numbers must be numbers and can not contain "
                                 "letters",
                                 parent=frame)
            return
    route_input = Handler(route.get()).routes_adj()  # call routes adj to shorten routes that don't need 5 digits
    if route_input == "0000":
        route_input = ""
    sql = "UPDATE carriers SET effective_date='%s',list_status='%s',ns_day='%s',route_s='%s',station='%s' " \
          "WHERE rowid = '%s'" % \
          (date, ls.get(), ns.get(), route_input, station.get(), rowid)
    commit(sql)
    frame.destroy()
    edit_carrier(name)


def delete_carrier(name):
    sql = "DELETE FROM carriers WHERE rowid = '%s'" % name[6]
    commit(sql)
    sql = "SELECT carrier_name FROM carriers WHERE carrier_name = '%s'" % name[1]
    results = inquire(sql)
    if len(results) > 0:
        edit_carrier(name[1])
    else:
        MainFrame().start()


def apply(year, month, day, c_name, ls, ns, route, station, frame):
    if year.get() > 9999:
        messagebox.showerror("Year Input Error", "Year must be between 1 and 9999", parent=frame)
        return
    if year.get() < 1:
        messagebox.showerror("Year Input Error", "Year must be between 1 and 9999", parent=frame)
        return

    try:
        date = datetime(year.get(), month.get(), day.get())
    except ValueError:
        messagebox.showerror("Invalid Date", "Date entered is not valid", parent=frame)
        return
    carrier = c_name.strip().lower()
    if len(carrier) > 30:
        messagebox.showerror("Name input error",
                             "Names must not exceed 30 characters.", parent=frame)
        return
    if len(carrier) < 1:
        messagebox.showerror("Name input error", "You must enter a name.", parent=frame)
        return
    if not apply_2(date, carrier, ls, ns, route, station, frame):
        return


def apply_2(date, carrier, ls, ns, route, station, frame):
    route_list = route.get().split("/")
    if len(route.get()) > 29:
        messagebox.showerror("Route number input error",
                             "There can be no more than five routes per carrier "
                             "(for T6 carriers).\n Routes numbers must be four or five digits long.\n"
                             "If there are multiple routes, route numbers must be separated by "
                             "the \'/\' character. For example: 1001/1015/10124/10224/0972. Do not use "
                             "commas or empty spaces",
                             parent=frame)
        return False
    for item in route_list:
        item = item.strip()
        if item != "":
            if len(item) < 4 or len(item) > 5:
                messagebox.showerror("Route number input error",
                                     'Routes numbers must be four or five digits long.\n'
                                     'If there are multiple routes, route numbers must be separated by '
                                     'the \'/\' character. For example: 1001/1015/1024/1036/1072. Do not use '
                                     'commas or empty spaces',
                                     parent=frame)
                return False
        if item.isdigit() == FALSE and item != "":
            messagebox.showerror("Route number input error",
                                 "Route numbers must be numbers and can not contain "
                                 "letters",
                                 parent=frame)
            return False
    # find all matches for date and name
    route_input = Handler(route.get()).routes_adj()  # call routes adj to shorten routes that don't need 5 digits
    if route_input == "0000":  # do not enter route for unassigned regulars
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
    return True


def name_change(name, c_name, frame):
    c_name = c_name.get().strip().lower()
    if messagebox.askokcancel("Name Change",
                              "This will change the name {} to {} in all records. "
                              "Are you sure?".format(name, c_name),
                              parent=frame):
        if len(c_name) > 42:
            messagebox.showerror("Name input error", "Names must not exceed 42 characters.", parent=frame)
            return
        if len(c_name) < 1:
            messagebox.showerror("Name input error", "You must enter a name.", parent=frame)
            return
        sql = "SELECT kb_name FROM name_index WHERE kb_name = '%s'" % c_name
        result = inquire(sql)
        if result:
            messagebox.showerror("Name input error", "This name is already being used for another carrier.",
                                 parent=frame)
            return
        sql = "SELECT carrier_name FROM carriers WHERE carrier_name = '%s'" % c_name
        result = inquire(sql)
        if result:
            messagebox.showerror("Name input error", "This name is already being used for another carrier.",
                                 parent=frame)
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
        MainFrame().start(frame=frame)


def purge_carrier(frame, carrier):
    if not messagebox.askokcancel("Delete Carrier",
                                  "This will delete the carrier and all records associated with "
                                  "this carrier, including rings and name index.\n\n"
                                  "If this carrier has left the station, quit, been fired or retired "
                                  "you should change station to \"out of station\" and not delete. \n\n"
                                  "This can not be reversed.",
                                  parent=frame):
        return
    sql = "DELETE FROM carriers WHERE carrier_name = '%s'" % carrier
    commit(sql)
    sql = "DELETE FROM rings3 WHERE carrier_name= '%s'" % carrier
    commit(sql)
    sql = "DELETE FROM name_index WHERE kb_name = '%s'" % carrier
    commit(sql)
    MainFrame().start(frame=frame)


def update_carrier(a):
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
    switch_f4 = Frame(projvar.root)
    switch_f4.pack(fill=BOTH, side=LEFT)
    c1 = Canvas(switch_f4)
    c1.pack(fill=BOTH, side=BOTTOM)
    # define scrollbar and canvas
    s = Scrollbar(switch_f4)
    c = Canvas(switch_f4, width=1600)
    # link up the canvas and scrollbar
    s.pack(side=RIGHT, fill=BOTH)
    c.pack(side=LEFT, fill=BOTH, pady=10, padx=20)
    s.configure(command=c.yview, orient="vertical")
    c.configure(yscrollcommand=s.set)
    if sys.platform == "win32":
        c.bind_all('<MouseWheel>', lambda event: c.yview_scroll(int(projvar.mousewheel * (event.delta / 120)), "units"))
    elif sys.platform == "darwin":
        c.bind_all('<MouseWheel>', lambda event: c.yview_scroll(int(projvar.mousewheel * event.delta), "units"))
    elif sys.platform == "linux":
        c.bind_all('<Button-4>', lambda event: c.yview('scroll', -1, 'units'))
        c.bind_all('<Button-5>', lambda event: c.yview('scroll', 1, 'units'))
    # create the frame inside the canvas
    f = Frame(c)
    c.create_window((0, 0), window=f, anchor=NW)
    # page title
    title_f = Frame(f)
    Label(title_f, text="Update Carrier Information", font=macadj("bold", "Helvetica 18")) \
        .grid(row=0, column=0, columnspan=4)
    title_f.grid(row=0, sticky=W, pady=5)  # put frame on grid
    # date
    date_frame = Frame(f)  # define frame
    year = IntVar(date_frame)  # define variables for date
    month = IntVar(date_frame)
    day = IntVar(date_frame)
    # pre set values for date
    month.set(int(a[0][5:7]))
    day.set(int(a[0][8:10]))
    year.set(int(a[0][:4]))
    Label(date_frame, text=" Date (month/day/year):", background=macadj("gray95", "grey"),
          fg=macadj("black", "white"), width=30, anchor="w") \
        .grid(row=0, column=0, sticky=W, columnspan=30)  # date label
    om_month = OptionMenu(date_frame, month, "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12")
    om_month.config(width=2)
    om_month.grid(row=1, column=0, sticky=W)
    om_day = OptionMenu(date_frame, day, "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14",
                        "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28", "29",
                        "30", "31")
    om_day.config(width=2)
    om_day.grid(row=1, column=1, sticky=W)
    Entry(date_frame, width=6, textvariable=year).grid(row=1, column=2, sticky=W)
    date_frame.grid(row=1, sticky=W, pady=5)  # put frame on grid
    # carrier name
    name_frame = Frame(f, pady=2)
    name = StringVar(name_frame)
    name = a[1]  # name value if name is not changed
    Label(name_frame, text=" Carrier Name: ", anchor="w", background=macadj("gray95", "grey"),
          fg=macadj("black", "white"), width=30).grid(row=0, column=0, sticky=W)
    Label(name_frame, text="{}".format(a[1].lower()), anchor="w", width=37).grid(row=1, column=0, sticky=W)
    name_frame.grid(row=2, sticky=W, pady=5)
    # list status
    list_frame = Frame(f, bd=1, relief=RIDGE, pady=2)
    Label(list_frame, text=" List Status", anchor="w", background=macadj("gray95", "grey"),
          fg=macadj("black", "white"), width=30).grid(row=0, column=0, sticky=W, columnspan=2)
    ls = StringVar(list_frame)
    ls.set(value=a[2])
    Radiobutton(list_frame, text="OTDL", variable=ls, value='otdl', justify=LEFT) \
        .grid(row=1, column=0, sticky=W)
    Radiobutton(list_frame, text="Work Assignment", variable=ls, value='wal', justify=LEFT) \
        .grid(row=1, column=1, sticky=W)
    Radiobutton(list_frame, text="No List", variable=ls, value='nl', justify=LEFT) \
        .grid(row=2, column=0, sticky=W)
    Radiobutton(list_frame, text="Auxiliary", variable=ls, value='aux', justify=LEFT) \
        .grid(row=2, column=1, sticky=W)
    Radiobutton(list_frame, text="Part Time Flex", variable=ls, value="ptf", justify=LEFT) \
        .grid(row=3, column=1, sticky=W)
    list_frame.grid(row=3, sticky=W, pady=5)
    # set non scheduled day
    ns_frame = Frame(f, pady=2)
    Label(ns_frame, text=" Non Scheduled Day", anchor="w", background=macadj("gray95", "grey"),
          fg=macadj("black", "white"), width=30).grid(row=0, column=0, sticky=W, columnspan=2)
    ns = StringVar(ns_frame)
    ns.set(a[3])
    Radiobutton(ns_frame, text="{}:   {}"
                .format(projvar.ns_code['yellow'], ns_results[0][2]), variable=ns, value="yellow",
                indicatoron=macadj(0, 1), width=15, anchor="w",
                bg=macadj("grey", "white"), fg=macadj("white", "black"), selectcolor=ns_color_dict["yellow"])\
        .grid(row=1, column=0)
    Radiobutton(ns_frame, text="{}:   {}"
                .format(projvar.ns_code['blue'], ns_results[1][2]), variable=ns, value="blue",
                indicatoron=macadj(0, 1), width=15, anchor="w",
                bg=macadj("grey", "white"), fg=macadj("white", "black"), selectcolor=ns_color_dict["blue"]) \
        .grid(row=1, column=1)
    Radiobutton(ns_frame, text="{}:   {}"
                .format(projvar.ns_code['green'], ns_results[2][2]), variable=ns, value="green",
                indicatoron=macadj(0, 1), width=15, anchor="w",
                bg=macadj("grey", "white"), fg=macadj("white", "black"), selectcolor=ns_color_dict["green"]) \
        .grid(row=2, column=0)
    Radiobutton(ns_frame, text="{}:   {}"
                .format(projvar.ns_code['brown'], ns_results[3][2]), variable=ns, value="brown",
                indicatoron=macadj(0, 1), width=15, anchor="w",
                bg=macadj("grey", "white"), fg=macadj("white", "black"), selectcolor=ns_color_dict["brown"]) \
        .grid(row=2, column=1)
    Radiobutton(ns_frame, text="{}:   {}".format(projvar.ns_code['red'], ns_results[4][2]), variable=ns, value="red",
                indicatoron=macadj(0, 1), width=15, anchor="w",
                bg=macadj("grey", "white"), fg=macadj("white", "black"), selectcolor=ns_color_dict["red"]) \
        .grid(row=3, column=0)
    Radiobutton(ns_frame, text="{}:   {}"
                .format(projvar.ns_code['black'], ns_results[5][2]), variable=ns, value="black",
                indicatoron=macadj(0, 1), width=15, anchor="w",
                bg=macadj("grey", "white"), fg=macadj("white", "black"), selectcolor=ns_color_dict["black"]) \
        .grid(row=3, column=1)
    Label(ns_frame, text=" Fixed:", anchor="w").grid(row=4, column=0, sticky="w")
    Radiobutton(ns_frame, text="none", variable=ns, value="none", indicatoron=macadj(0, 1), width=15,
                bg=macadj("grey", "white"), fg=macadj("white", "black"),
                selectcolor=ns_color_dict["none"], anchor="w") \
        .grid(row=4, column=1)
    Radiobutton(ns_frame, text="Sat:   fixed", variable=ns, value="sat",
                bg=macadj("grey", "white"), fg=macadj("white", "black"), selectcolor=ns_color_dict["sat"],
                indicatoron=macadj(0, 1), width=15, anchor="w") \
        .grid(row=5, column=0)
    Radiobutton(ns_frame, text="Mon:   fixed", variable=ns, value="mon",
                bg=macadj("grey", "white"), fg=macadj("white", "black"), selectcolor=ns_color_dict["mon"],
                indicatoron=macadj(0, 1), width=15, anchor="w") \
        .grid(row=5, column=1)
    Radiobutton(ns_frame, text="Tue:   fixed", variable=ns, value="tue",
                bg=macadj("grey", "white"), fg=macadj("white", "black"), selectcolor=ns_color_dict["tue"],
                indicatoron=macadj(0, 1), width=15, anchor="w") \
        .grid(row=6, column=0)
    Radiobutton(ns_frame, text="Wed:   fixed", variable=ns, value="wed",
                bg=macadj("grey", "white"), fg=macadj("white", "black"), selectcolor=ns_color_dict["wed"],
                indicatoron=macadj(0, 1), width=15, anchor="w") \
        .grid(row=6, column=1)
    Radiobutton(ns_frame, text="Thu:   fixed", variable=ns, value="thu",
                bg=macadj("grey", "white"), fg=macadj("white", "black"), selectcolor=ns_color_dict["thu"],
                indicatoron=macadj(0, 1), width=15, anchor="w") \
        .grid(row=7, column=0)
    Radiobutton(ns_frame, text="Fri:   fixed", variable=ns, value="fri",
                bg=macadj("grey", "white"), fg=macadj("white", "black"), selectcolor=ns_color_dict["fri"],
                indicatoron=macadj(0, 1), width=15, anchor="w") \
        .grid(row=7, column=1)
    ns_frame.grid(row=4, sticky=W, pady=5)
    # set route entry field
    route_frame = Frame(f, bd=1, relief=RIDGE, pady=2)
    Label(route_frame, text=" Route/s", anchor="w", background=macadj("gray95", "grey"),
          fg=macadj("black", "white"), width=30).grid(row=0, column=0, sticky=W)
    route = StringVar(route_frame)
    route.set(a[4])
    Entry(route_frame, width=macadj(37, 29), textvariable=route).grid(row=1, column=0, sticky=W)
    route_frame.grid(row=5, sticky=W, pady=5)
    # set station option menu
    station_frame = Frame(f, pady=2)
    Label(station_frame, text="Station", anchor="w", background=macadj("gray95", "grey"),
          fg=macadj("black", "white"), width=5).grid(row=0, column=0, sticky=W)
    station = StringVar(station_frame)
    station.set(a[5])  # default value
    om_stat = OptionMenu(station_frame, station, *projvar.list_of_stations)
    om_stat.config(width=macadj("24", "22"))
    om_stat.grid(row=0, column=1, sticky=W)
    station_frame.grid(row=6, sticky=W, pady=5)
    # set rowid
    rowid = StringVar(f)
    rowid = a[6]
    projvar.root.update()
    c.config(scrollregion=c.bbox("all"))
    # apply and close buttons
    button_apply = Button(c1)  # buttons at bottom of screen
    button_back = Button(c1)
    button_apply.config(text="Apply", command=lambda:
        apply_update_carrier(year, month, day, name, ls, ns, route, station, rowid, switch_f4))
    button_back.config(text="Go Back", command=lambda: MainFrame().start(frame=switch_f4))
    if sys.platform == "win32":
        button_apply.config(anchor="w", width=15)
        button_back.config(anchor="w", width=15)
    else:
        button_apply.config(width=16)
        button_back.config(width=16)
    button_apply.pack(side=LEFT)
    button_back.pack(side=LEFT)


def edit_carrier(e_name):
    sql = "SELECT effective_date, carrier_name, list_status, ns_day,route_s, station, rowid" \
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
    switch_f3 = Frame(projvar.root)
    switch_f3.pack(fill=BOTH, side=LEFT)
    c1 = Canvas(switch_f3)
    c1.pack(fill=BOTH, side=BOTTOM)
    # define scrollbar and canvas
    s = Scrollbar(switch_f3)
    c = Canvas(switch_f3, width=1600)
    # link up the canvas and scrollbar
    s.pack(side=RIGHT, fill=BOTH)
    c.pack(side=LEFT, fill=BOTH, pady=10, padx=20)
    s.configure(command=c.yview, orient="vertical")
    c.configure(yscrollcommand=s.set)
    if sys.platform == "win32":
        c.bind_all('<MouseWheel>', lambda event: c.yview_scroll(int(projvar.mousewheel * (event.delta / 120)), "units"))
    elif sys.platform == "darwin":
        c.bind_all('<MouseWheel>', lambda event: c.yview_scroll(int(projvar.mousewheel * event.delta), "units"))
    elif sys.platform == "linux":
        c.bind_all('<Button-4>', lambda event: c.yview('scroll', -1, 'units'))
        c.bind_all('<Button-5>', lambda event: c.yview('scroll', 1, 'units'))
    # create the frame inside the canvas
    f = Frame(c)
    c.create_window((0, 0), window=f, anchor=NW)
    # page title
    title_f = Frame(f)
    Label(title_f, text="Edit Carrier Information", font=macadj("bold", "Helvetica 18")) \
        .grid(row=0, column=0, columnspan=4)
    title_f.grid(row=0, sticky=W, pady=5)  # put frame on grid
    # current date
    year = IntVar(f)
    month = IntVar(f)
    day = IntVar(f)
    # pre set values for date
    month.set(projvar.invran_month)
    day.set(projvar.invran_day)
    year.set(projvar.invran_year)
    # define frame
    date_frame = Frame(f)
    Label(date_frame, text=" Date (month/day/year):", background=macadj("gray95", "grey"), fg=macadj("black", "white"),
          width=30, anchor="w").grid(row=0, column=0, sticky=W, columnspan=30)  # date label
    om_month = OptionMenu(date_frame, month, "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12")
    om_month.config(width=2)
    om_month.grid(row=1, column=0, sticky=W)  # option menu for month
    om_day = OptionMenu(date_frame, day, "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14",
                        "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28", "29", "30",
                        "31")
    om_day.config(width=2)
    om_day.grid(row=1, column=1, sticky=W)  # option menu for day
    Entry(date_frame, width=6, textvariable=year).grid(row=1, column=2, sticky=W)  # entry field for year
    date_frame.grid(row=1, column=0, sticky=W, pady=5)  # put frame on grid
    # carrier name
    name_frame = Frame(f, pady=2)
    c_name = StringVar(name_frame)
    name = StringVar(name_frame)
    name = e_name  # name value if name is not changed
    c_name.set(e_name)  # name value for name changes
    Label(name_frame, text=" Carrier Name: {}".format(e_name), anchor="w", background=macadj("gray95", "grey"),
          fg=macadj("black", "white"), width=30).grid(row=0, column=0, columnspan=4, sticky=W)
    Entry(name_frame, width=macadj(37, 29), textvariable=c_name).grid(row=1, column=0, columnspan=4, sticky=W)
    Label(name_frame, text="Change Name: ").grid(row=2, column=0, sticky=W)
    Button(name_frame, width=7, text="update", command=lambda: name_change(name, c_name, switch_f3)) \
        .grid(row=2, column=1, sticky=W, pady=6)
    name_frame.grid(row=2, sticky=W, pady=5)
    # list status
    list_frame = Frame(f, bd=1, relief=RIDGE, pady=2)
    Label(list_frame, text=" List Status", anchor="w", background=macadj("gray95", "grey"),
          fg=macadj("black", "white"), width=30).grid(row=0, column=0, sticky=W, columnspan=2)
    ls = StringVar(list_frame)
    ls.set(results[0][2])
    Radiobutton(list_frame, text="OTDL", variable=ls, value='otdl', justify=LEFT) \
        .grid(row=1, column=0, sticky=W)
    Radiobutton(list_frame, text="Work Assignment", variable=ls, value='wal', justify=LEFT) \
        .grid(row=1, column=1, sticky=W)
    Radiobutton(list_frame, text="No List", variable=ls, value='nl', justify=LEFT) \
        .grid(row=2, column=0, sticky=W)
    Radiobutton(list_frame, text="Auxiliary", variable=ls, value='aux', justify=LEFT) \
        .grid(row=2, column=1, sticky=W)
    Radiobutton(list_frame, text="Part Time Flex", variable=ls, value="ptf", justify=LEFT) \
        .grid(row=3, column=1, sticky=W)
    list_frame.grid(row=3, sticky=W, pady=5)
    # set non scheduled day
    ns_frame = Frame(f, pady=2)
    Label(ns_frame, text=" Non Scheduled Day", anchor="w", background=macadj("gray95", "grey"),
          fg=macadj("black", "white"),
          width=30).grid(row=0, column=0, sticky=W, columnspan=2)
    ns = StringVar(ns_frame)
    ns.set(results[0][3])
    Radiobutton(ns_frame, text="{}:   {}"
                .format(projvar.ns_code['yellow'], ns_results[0][2]), variable=ns, value="yellow",
                indicatoron=macadj(0, 1), width=15, anchor="w",
                bg=macadj("grey", "white"), fg=macadj("white", "black"), selectcolor=ns_color_dict["yellow"]) \
        .grid(row=1, column=0)
    Radiobutton(ns_frame, text="{}:   {}"
                .format(projvar.ns_code['blue'], ns_results[1][2]), variable=ns, value="blue",
                indicatoron=macadj(0, 1), width=15, anchor="w",
                bg=macadj("grey", "white"), fg=macadj("white", "black"), selectcolor=ns_color_dict["blue"]) \
        .grid(row=1, column=1)
    Radiobutton(ns_frame, text="{}:   {}"
                .format(projvar.ns_code['green'], ns_results[2][2]), variable=ns, value="green",
                indicatoron=macadj(0, 1), width=15, anchor="w",
                bg=macadj("grey", "white"), fg=macadj("white", "black"), selectcolor=ns_color_dict["green"]) \
        .grid(row=2, column=0)
    Radiobutton(ns_frame, text="{}:   {}"
                .format(projvar.ns_code['brown'], ns_results[3][2]), variable=ns, value="brown",
                indicatoron=macadj(0, 1), width=15, anchor="w",
                bg=macadj("grey", "white"), fg=macadj("white", "black"), selectcolor=ns_color_dict["brown"]) \
        .grid(row=2, column=1)
    Radiobutton(ns_frame, text="{}:   {}"
                .format(projvar.ns_code['red'], ns_results[4][2]), variable=ns, value="red",
                indicatoron=macadj(0, 1), width=15, anchor="w",
                bg=macadj("grey", "white"), fg=macadj("white", "black"), selectcolor=ns_color_dict["red"]) \
        .grid(row=3, column=0)
    Radiobutton(ns_frame, text="{}:   {}"
                .format(projvar.ns_code['black'], ns_results[5][2]), variable=ns, value="black",
                indicatoron=macadj(0, 1), width=15, anchor="w",
                bg=macadj("grey", "white"), fg=macadj("white", "black"), selectcolor=ns_color_dict["black"]) \
        .grid(row=3, column=1)
    Label(ns_frame, text=" Fixed:", anchor="w").grid(row=4, column=0, sticky="w")
    Radiobutton(ns_frame, text="none", variable=ns, value="none", indicatoron=macadj(0, 1), width=15,
                bg=macadj("grey", "white"), fg=macadj("white", "black"),
                selectcolor=ns_color_dict["none"], anchor="w") \
        .grid(row=4, column=1)
    Radiobutton(ns_frame, text="Sat:   fixed", variable=ns, value="sat",
                bg=macadj("grey", "white"), fg=macadj("white", "black"), selectcolor=ns_color_dict["sat"],
                indicatoron=macadj(0, 1), width=15, anchor="w") \
        .grid(row=5, column=0)
    Radiobutton(ns_frame, text="Mon:   fixed", variable=ns, value="mon",
                bg=macadj("grey", "white"), fg=macadj("white", "black"), selectcolor=ns_color_dict["mon"],
                indicatoron=macadj(0, 1), width=15, anchor="w") \
        .grid(row=5, column=1)
    Radiobutton(ns_frame, text="Tue:   fixed", variable=ns, value="tue",
                bg=macadj("grey", "white"), fg=macadj("white", "black"), selectcolor=ns_color_dict["tue"],
                indicatoron=macadj(0, 1), width=15, anchor="w") \
        .grid(row=6, column=0)
    Radiobutton(ns_frame, text="Wed:   fixed", variable=ns, value="wed",
                bg=macadj("grey", "white"), fg=macadj("white", "black"), selectcolor=ns_color_dict["wed"],
                indicatoron=macadj(0, 1), width=15, anchor="w") \
        .grid(row=6, column=1)
    Radiobutton(ns_frame, text="Thu:   fixed", variable=ns, value="thu",
                bg=macadj("grey", "white"), fg=macadj("white", "black"), selectcolor=ns_color_dict["thu"],
                indicatoron=macadj(0, 1), width=15, anchor="w") \
        .grid(row=7, column=0)
    Radiobutton(ns_frame, text="Fri:   fixed", variable=ns, value="fri",
                bg=macadj("grey", "white"), fg=macadj("white", "black"), selectcolor=ns_color_dict["fri"],
                indicatoron=macadj(0, 1), width=15, anchor="w") \
        .grid(row=7, column=1)
    ns_frame.grid(row=4, sticky=W, pady=5)
    # set route entry field
    route_frame = Frame(f, bd=1, relief=RIDGE, pady=2)
    Label(route_frame, text=" Route/s", anchor="w", background=macadj("gray95", "grey"), fg=macadj("black", "white"),
          width=30).grid(row=0, column=0, sticky=W)
    route = StringVar(route_frame)
    route.set(results[0][4])
    Entry(route_frame, width=macadj(37, 29), textvariable=route).grid(row=1, column=0, sticky=W)
    route_frame.grid(row=5, sticky=W, pady=5)
    # set station option menu
    station_frame = Frame(f, pady=2)
    Label(station_frame, text="Station", anchor="w", background=macadj("gray95", "grey"), fg=macadj("black", "white"),
          width=5).grid(row=0, column=0, sticky=W)
    station = StringVar(station_frame)
    station.set(results[0][5])  # default value
    om_stat = OptionMenu(station_frame, station, *projvar.list_of_stations)
    om_stat.config(width=macadj("24", "22"))
    om_stat.grid(row=0, column=1, sticky=W)
    # Label(station_frame, text=" ").grid(row=1)
    station_frame.grid(row=6, sticky=W, pady=5)
    #  delete button
    delete_frame = Frame(f, bd=1, relief=RIDGE, pady=2)
    Label(delete_frame, text=" Delete All", anchor="w", background=macadj("gray95", "grey"),
          fg=macadj("black", "white"),
          width=macadj(8, 10)).grid(row=0, column=0, sticky=W)
    Label(delete_frame, text="Delete carrier and all associated records. ", anchor="w") \
        .grid(row=1, column=0, sticky=W)
    Button(delete_frame, text="Delete", width=15,
           bg=macadj("red3", "white"), fg=macadj("white", "red"),
           command=lambda: purge_carrier(switch_f3, e_name)).grid(row=3, column=0, sticky=W, padx=8)
    delete_frame.grid(row=7, sticky=W, pady=5)
    report_frame = Frame(f, padx=2, )
    Label(report_frame, text="Status Change Report: ", anchor="w").grid(row=0, column=0, sticky=W, columnspan=4)
    Label(report_frame, text="Generate Report: ", anchor="w").grid(row=1, column=0, sticky=W)
    Button(report_frame, text="Report", width=10, command=lambda: Reports(switch_f3).rpt_carrier_history(e_name)) \
        .grid(row=1, column=1, sticky=W, padx=10)
    report_frame.grid(row=8, sticky=W, pady=5)
    Label(f, text="").grid(row=9)
    #   History of status changes
    history_frame = Frame(f, pady=2)
    row_line = 0
    Label(history_frame, text=" Status Change History", anchor="w", font=macadj("bold", "Helvetica 18"),
          background=macadj("gray95", "grey"), fg=macadj("black", "white"), width=30) \
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
        Label(history_frame, width=35, text="route: {}".format(line[4]), anchor="w") \
            .grid(row=row_line, column=0, sticky=W, columnspan=4)
        row_line += 1
        Label(history_frame, width=25, text="station: {}".format(line[5]), anchor="w") \
            .grid(row=row_line, column=0, sticky=W, columnspan=4)
        row_line += 1
        Button(history_frame, width=14, text="edit", anchor="w",
               command=lambda x=line: [switch_f3.destroy(), update_carrier(x)]) \
            .grid(row=row_line, column=0, sticky=W, )
        Button(history_frame, width=14, text="delete", anchor="w",
               command=lambda x=line: [switch_f3.destroy(), delete_carrier(x)]) \
            .grid(row=row_line, column=1, sticky=W)
        Label(history_frame, text="                             ").grid(row=row_line, column=2, sticky=W)
        row_line += 1
    history_frame.grid(row=9, sticky=W, pady=5)
    projvar.root.update()
    c.config(scrollregion=c.bbox("all"))
    button_apply = Button(c1)  # buttons at bottom of screen
    button_back = Button(c1)
    button_apply.config(text="Apply", command=lambda: [apply(year, month, day, name, ls, ns, route, station, switch_f3),
                            MainFrame().start(frame=switch_f3)])
    button_back.config(text="Go Back", command=lambda: MainFrame().start(frame=switch_f3))
    if sys.platform == "win32":
        button_apply.config(anchor="w", width=15)
        button_back.config(anchor="w", width=15)
    else:
        button_apply.config(width=16)
        button_back.config(width=16)
    button_apply.pack(side=LEFT)
    button_back.pack(side=LEFT)


def nc_apply(year, month, day, nc_name, nc_fname, nc_ls, nc_ns, nc_route, nc_station, frame):
    if year.get() > 9999 or year.get() < 1000:
        messagebox.showerror("Year Input Error", "Year must be between 1000 and 9999", parent=frame)
        return
    try:
        date = datetime(year.get(), month.get(), day.get())
    except ValueError:
        messagebox.showerror("Invalid Date",
                             "Date entered is not valid",
                             parent=frame)
        return
    carrier = nc_name.get().strip().lower() + ", " + nc_fname.get().strip().lower()
    if len(nc_name.get()) > 30 or len(nc_fname.get()) > 12:
        messagebox.showerror("Name input error",
                             "Names must not exceed 30 characters."
                             "First names must not exceed 12 characters",
                             parent=frame)
        return
    if len(nc_name.get()) < 1:
        messagebox.showerror("Name input error",
                             "You must enter a name.",
                             parent=frame)
        return
    if len(nc_fname.get()) < 1:
        messagebox.showerror("Name input error",
                             "You must enter a first initial or name.",
                             parent=frame)
        return
    if len(nc_fname.get()) > 1:
        answer = messagebox.askyesno("Caution",
                                     "It is recommended that you use only the first initial of the first"
                                     "name unless it is necessary to create a unique identifier, such as"
                                     "when you have two identical names that must be distinguished."
                                     "Do you want to proceed?",
                                     parent=frame)
        if not answer:
            return
    nc_route_list = nc_route.get().split("/")
    if len(nc_route.get()) > 29:
        messagebox.showerror("Route number input error",
                             "There can be no more than five routes per carrier "
                             "(for T6 carriers).\n Routes numbers four or five digits long.\n"
                             "If there are multiple routes, route numbers must be separated by "
                             "the \'/\' character. For example: 1001/1015/10124/10224/0972. Do not use "
                             "commas or empty spaces",
                             parent=frame)
        return
    for item in nc_route_list:
        item = item.strip()
        if item != "":
            if len(item) < 4 or len(item) > 5:
                messagebox.showerror("Route number input error",
                                     "Routes numbers must be four or five digits long.\n"
                                     "If there are multiple routes, route numbers must be separated by "
                                     "the \"/\" character. For example: 1001/1015/10124/10224/0972. Do not use "
                                     "commas or empty spaces",
                                     parent=frame)
                return
        if item.isdigit() == FALSE and item != "":
            messagebox.showerror("Route number input error",
                                 "Route numbers must be numbers and can not contain "
                                 "letters",
                                 parent=frame)
            return
    route_input = Handler(nc_route.get()).routes_adj()  # call routes adj to shorten routes that don't need 5 digits
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
        ok = messagebox.askokcancel("New Carrier Input Warning",
                                    "This carrier name is already in the database.\n"
                                    "Did you want to proceed?",
                                    parent=frame)
        if ok:
            for pair in results:
                if pair[0] == carrier and pair[1] == str(datetime(year.get(), month.get(), day.get(), 00, 00, 00)):
                    messagebox.showwarning("New Carrier - Prohibited Action",
                                           "There is a pre existing record for this carrier on this day.\n"
                                           "You can not update that record using this window.\n"
                                           "To edit/ delete this record, return to the main page and press\n"
                                           "\"edit\" to the right of the carrier's name. ",
                                           parent=frame)
                    match = True
        if not ok:
            match = True
    if not match:
        commit(sql)
    MainFrame().start(frame=frame)


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
    switch_f6 = Frame(projvar.root)
    switch_f6.pack(fill=BOTH, side=LEFT)
    c1 = Canvas(switch_f6)
    c1.pack(fill=BOTH, side=BOTTOM)
    button_apply = Button(c1)  # buttons at bottom of screen
    button_back = Button(c1)
    button_apply.config(text="Apply", command=lambda:
        (nc_apply(year, month, day, nc_name, nc_fname, nc_ls, nc_ns, nc_route, nc_station, switch_f6)))
    button_back.config(text="Go Back", command=lambda: MainFrame().start(frame=switch_f6))
    if sys.platform == "win32":
        button_apply.config(anchor="w", width=15)
        button_back.config(anchor="w", width=15)
    else:
        button_apply.config(width=16)
        button_back.config(width=16)
    button_apply.pack(side=LEFT)
    button_back.pack(side=LEFT)
    # set up variable for scrollbar and canvas
    s = Scrollbar(switch_f6)
    c = Canvas(switch_f6, width=1600)
    # link up the canvas and scrollbar
    s.pack(side=RIGHT, fill=BOTH)
    c.pack(side=LEFT, fill=BOTH, pady=10, padx=20)
    s.configure(command=c.yview, orient="vertical")
    c.configure(yscrollcommand=s.set)
    if sys.platform == "win32":
        c.bind_all('<MouseWheel>', lambda event: c.yview_scroll(int(projvar.mousewheel * (event.delta / 120)), "units"))
    elif sys.platform == "darwin":
        c.bind_all('<MouseWheel>', lambda event: c.yview_scroll(int(projvar.mousewheel * event.delta), "units"))
    elif sys.platform == "linux":
        c.bind_all('<Button-4>', lambda event: c.yview('scroll', -1, 'units'))
        c.bind_all('<Button-5>', lambda event: c.yview('scroll', 1, 'units'))
    # create the frame inside the canvas
    nc_f = Frame(c)
    c.create_window((0, 0), window=nc_f, anchor=NW)
    # page title
    title_f = Frame(nc_f)
    Label(title_f, text="Enter New Carrier", font=macadj("bold", "Helvetica 18")) \
        .grid(row=0, column=0, columnspan=4)
    title_f.grid(row=0, sticky=W, pady=5)  # put frame on grid
    # date
    date_frame = Frame(nc_f)  # define frame
    year = IntVar(date_frame)  # define variables for date
    month = IntVar(date_frame)
    day = IntVar(date_frame)
    month.set(projvar.invran_month)  # set values for variables
    day.set(projvar.invran_day)
    year.set(projvar.invran_year)
    Label(date_frame, text=" Date (month/day/year):", background=macadj("gray95", "grey"),
          fg=macadj("black", "white"), width=30,
          anchor="w").grid(row=0, column=0, sticky=W, columnspan=30)  # date label
    om_month = OptionMenu(date_frame, month, "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12")
    om_month.config(width=2)
    om_month.grid(row=1, column=0, sticky=W)
    om_day = OptionMenu(date_frame, day, "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14",
                        "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28", "29",
                        "30", "31")
    om_day.config(width=2)
    om_day.grid(row=1, column=1, sticky=W)
    Entry(date_frame, width=6, textvariable=year).grid(row=1, column=2, sticky=W)
    date_frame.grid(row=1, sticky=W, pady=5)  # put frame on grid
    # carrier name:
    name_frame = Frame(nc_f, pady=2)
    Label(name_frame, text=" Last Name: ", width=22, anchor="w", background=macadj("gray95", "grey"),
          fg=macadj("black", "white")).grid(row=0, column=0, sticky=W)
    Label(name_frame, text=" 1st Initial ", width=7, anchor="w", background=macadj("gray95", "grey"),
          fg=macadj("black", "white")).grid(row=0, column=1, sticky=W)
    nc_name = StringVar(nc_f)
    nc_fname = StringVar(nc_f)
    Entry(name_frame, width=macadj(27, 22), textvariable=nc_name).grid(row=1, column=0, sticky=W)
    Entry(name_frame, width=macadj(8, 6), textvariable=nc_fname).grid(row=1, column=1, sticky=W)
    name_frame.grid(row=2, sticky=W, pady=5)
    # list status
    list_frame = Frame(nc_f, bd=1, relief=RIDGE, pady=5)
    Label(list_frame, width=30, text=" List Status", anchor="w", background=macadj("gray95", "grey"),
          fg=macadj("black", "white")).grid(row=0, column=0, sticky=W, columnspan=2)
    nc_ls = StringVar(list_frame)
    nc_ls.set(value="nl")
    Radiobutton(list_frame, text="OTDL", variable=nc_ls, value='otdl', justify=LEFT) \
        .grid(row=1, column=0, sticky=W)
    Radiobutton(list_frame, text="Work Assignment", variable=nc_ls, value='wal', justify=LEFT) \
        .grid(row=1, column=1, sticky=W)
    Radiobutton(list_frame, text="No List", variable=nc_ls, value='nl', justify=LEFT) \
        .grid(row=2, column=0, sticky=W)
    Radiobutton(list_frame, text="Auxiliary", variable=nc_ls, value='aux', justify=LEFT) \
        .grid(row=2, column=1, sticky=W)
    Radiobutton(list_frame, text="Part Time Flex", variable=nc_ls, value='ptf', justify=LEFT) \
        .grid(row=3, column=1, sticky=W)
    list_frame.grid(row=3, sticky=W, pady=5)
    # set non scheduled day
    ns_frame = Frame(nc_f, pady=5)
    Label(ns_frame, width=30, text=" Non Scheduled Day", anchor="w", background=macadj("gray95", "grey"),
          fg=macadj("black", "white")).grid(row=0, column=0, sticky=W, columnspan=2)
    nc_ns = StringVar(ns_frame)
    nc_ns.set("none")
    Radiobutton(ns_frame, text="{}:   yellow".format(projvar.ns_code['yellow']), variable=nc_ns, value="yellow",
                indicatoron=macadj(0, 1), width=15, anchor="w", bg=macadj("grey", "white"), fg=macadj("white", "black"),
                selectcolor=ns_color_dict["yellow"]).grid(row=1, column=0)
    Radiobutton(ns_frame, text="{}:   blue".format(projvar.ns_code['blue']), variable=nc_ns, value="blue",
                indicatoron=macadj(0, 1), width=15, anchor="w", bg=macadj("grey", "white"), fg=macadj("white", "black"),
                selectcolor=ns_color_dict["blue"]).grid(row=2, column=0)
    Radiobutton(ns_frame, text="{}:   green".format(projvar.ns_code['green']), variable=nc_ns, value="green",
                indicatoron=macadj(0, 1),
                width=15, anchor="w", bg=macadj("grey", "white"), fg=macadj("white", "black"),
                selectcolor=ns_color_dict["green"]).grid(row=3, column=0)
    Radiobutton(ns_frame, text="{}:   brown".format(projvar.ns_code['brown']), variable=nc_ns, value="brown",
                indicatoron=macadj(0, 1),
                width=15, anchor="w", bg=macadj("grey", "white"), fg=macadj("white", "black"),
                selectcolor=ns_color_dict["brown"]).grid(row=1, column=1)
    Radiobutton(ns_frame, text="{}:   red".format(projvar.ns_code['red']), variable=nc_ns, value="red",
                indicatoron=macadj(0, 1), width=15,
                anchor="w", bg=macadj("grey", "white"), fg=macadj("white", "black"),
                selectcolor=ns_color_dict["red"]).grid(row=2, column=1)
    Radiobutton(ns_frame, text="{}:   black".format(projvar.ns_code['black']), variable=nc_ns, value="black",
                indicatoron=macadj(0, 1),
                width=15, anchor="w", bg=macadj("grey", "white"), fg=macadj("white", "black"),
                selectcolor=ns_color_dict["black"]).grid(row=3, column=1)
    Label(ns_frame, text=" Fixed:", anchor="w").grid(row=4, column=0, sticky="w")
    Radiobutton(ns_frame, text="none", variable=nc_ns, value="none", indicatoron=macadj(0, 1),
                width=15, anchor="w") \
        .grid(row=4, column=1)
    Radiobutton(ns_frame, text="none", variable=nc_ns, value="none",
                indicatoron=macadj(0, 1), width=15, bg=macadj("grey", "white"), fg=macadj("white", "black"),
                selectcolor=ns_color_dict["none"], anchor="w").grid(row=4, column=1)
    Radiobutton(ns_frame, text="Sat:   fixed", variable=nc_ns, value="sat", bg=macadj("grey", "white"),
                fg=macadj("white", "black"),
                selectcolor=ns_color_dict["sat"], indicatoron=macadj(0, 1),
                width=15, anchor="w").grid(row=5, column=0)
    Radiobutton(ns_frame, text="Mon:   fixed", variable=nc_ns, value="mon", bg=macadj("grey", "white"),
                fg=macadj("white", "black"),
                selectcolor=ns_color_dict["mon"], indicatoron=macadj(0, 1),
                width=15, anchor="w").grid(row=5, column=1)
    Radiobutton(ns_frame, text="Tue:   fixed", variable=nc_ns, value="tue", bg=macadj("grey", "white"),
                fg=macadj("white", "black"),
                selectcolor=ns_color_dict["tue"], indicatoron=macadj(0, 1),
                width=15, anchor="w").grid(row=6, column=0)
    Radiobutton(ns_frame, text="Wed:   fixed", variable=nc_ns, value="wed", bg=macadj("grey", "white"),
                fg=macadj("white", "black"),
                selectcolor=ns_color_dict["wed"], indicatoron=macadj(0, 1),
                width=15, anchor="w").grid(row=6, column=1)
    Radiobutton(ns_frame, text="Thu:   fixed", variable=nc_ns, value="thu", bg=macadj("grey", "white"),
                fg=macadj("white", "black"),
                selectcolor=ns_color_dict["thu"], indicatoron=macadj(0, 1),
                width=15, anchor="w").grid(row=7, column=0)
    Radiobutton(ns_frame, text="Fri:   fixed", variable=nc_ns, value="fri", bg=macadj("grey", "white"),
                fg=macadj("white", "black"),
                selectcolor=ns_color_dict["fri"], indicatoron=macadj(0, 1),
                width=15, anchor="w").grid(row=7, column=1)
    ns_frame.grid(row=4, sticky=W, pady=5)
    # set route entry field
    route_frame = Frame(nc_f, bd=1, relief=RIDGE, pady=2)
    Label(route_frame, text=" Route/s", width=30, anchor="w", background=macadj("gray95", "grey"),
          fg=macadj("black", "white")).grid(row=0, column=0, sticky=W)
    nc_route = StringVar(route_frame)
    nc_route.set("")
    Entry(route_frame, width=macadj(37, 29), textvariable=nc_route).grid(row=1, column=0, sticky=W)
    route_frame.grid(row=5, sticky=W)
    # set station option menu
    station_frame = Frame(nc_f, pady=5)
    Label(station_frame, text="Station", width=5, anchor="w", background=macadj("gray95", "grey"),
          fg=macadj("black", "white")) \
        .grid(row=0, column=0, sticky=W)
    nc_station = StringVar(station_frame)
    nc_station.set(projvar.invran_station)  # default value
    om_stat = OptionMenu(station_frame, nc_station, *projvar.list_of_stations)
    om_stat.config(width=macadj(24, 22))
    om_stat.grid(row=0, column=1, sticky=W)
    station_frame.grid(row=6, sticky=W, pady=5)
    projvar.root.update()
    c.config(scrollregion=c.bbox("all"))


def reset(frame):  # reset initial value of globals
    projvar.invran_year = None
    projvar.invran_month = None
    projvar.invran_day = None
    projvar.invran_weekly_span = None  # default is weekly investigation range
    projvar.invran_station = None
    projvar.invran_date_week = []
    projvar.invran_date = None
    projvar.ns_code = {}
    if frame != "none":
        MainFrame().start(frame=frame)


def set_globals(s_year, s_mo, s_day, i_range, station, frame):
    projvar.invran_weekly_span = i_range
    if station == "undefined":
        messagebox.showerror("Investigation station setting",
                             'Please select a station.',
                             parent=frame)
        return
    # error check for valid date
    date = ""  # reference before assignment
    try:
        date = datetime(int(s_year), int(s_mo), int(s_day))
    except ValueError:
        messagebox.showerror("Investigation date/range",
                             'The date entered is not valid.',
                             parent=frame)
        return
    projvar.invran_date = date
    wkdy_name = date.strftime("%a")
    while wkdy_name != "Sat":  # while date enter is not a saturday
        date -= timedelta(days=1)  # walk back the date until it is a saturday
        wkdy_name = date.strftime("%a")
    sat_range = date  # sat range = sat or the sat most prior
    projvar.pay_period = pp_by_date(sat_range)
    projvar.invran_year = int(date.strftime("%Y"))  # format that sat to form the global
    projvar.invran_month = int(date.strftime("%m"))
    projvar.invran_day = int(date.strftime("%d"))
    del projvar.invran_date_week[:]  # empty out the array for the global date variable
    d = datetime(int(projvar.invran_year), int(projvar.invran_month), int(projvar.invran_day))
    # set the projvar.invran_date_week variable
    projvar.invran_date_week.append(d)
    for i in range(6):
        d += timedelta(days=1)
        projvar.invran_date_week.append(d)
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
    projvar.ns_code = {}
    for i in range(7):
        if i == 0:
            projvar.ns_code[pat[x]] = date.strftime("%a")
            date += timedelta(days=1)
        elif i == 1:
            date += timedelta(days=1)
            if x > 4:
                x = 0
            else:
                x += 1
        else:
            projvar.ns_code[pat[x]] = date.strftime("%a")
            date += timedelta(days=1)
            if x > 4:
                x = 0
            else:
                x += 1
    projvar.ns_code["none"] = "  "
    if not i_range:  # if investigation range is one day
        projvar.invran_year = int(s_year)
        projvar.invran_month = int(s_mo)
        projvar.invran_day = int(s_day)
        projvar.invran_day = int(s_day)
    projvar.ns_code["sat"] = "Sat"
    projvar.ns_code["mon"] = "Mon"
    projvar.ns_code["tue"] = "Tue"
    projvar.ns_code["wed"] = "Wed"
    projvar.ns_code["thu"] = "Thu"
    projvar.ns_code["fri"] = "Fri"
    projvar.invran_station = station
    if frame != "None":
        MainFrame().start(frame=frame)


class MainFrame:
    def __init__(self):
        self.win = None
        self.invest_frame = None
        self.main_frame = None
        self.start_year = None
        self.start_month = None  # stringvars
        self.start_day = None
        self.i_range = None  # investigation range boolean
        self.invran = None  # investigation range stringvar
        self.start_date = None
        self.end_date = None
        self.station = None
        self.carrier_list = []
        self.invran_date = None  # investigation range date
        self.stations_minus_outofstation = []  # list of stations
        self.invran_result = None

    def start(self, frame=None):  # master method for controlling methods in class
        self.win = MakeWindow()
        self.win.create(frame)  # create the window
        self.invest_frame = Frame(self.win.body)
        self.main_frame = Frame(self.win.body)
        self.invest_frame.pack()  # put the investigation frame in the window
        self.main_frame.pack()  # puts the mainframe in the window
        self.set_dates()
        self.make_stringvars()
        self.get_carrierlist()  # call CarrierList to get Carrier Rec Set
        self.pulldown_menu()  # create a pulldown menu, and add it to the menu bar
        self.set_investigation_vars()  # set the stringvars for the investigation range
        self.get_stations_list()  # get a list of stations for station optionmenu
        self.get_invran_mode()  # get the investigation range mode. alternate widget layouts for investigation range
        if self.invran_result in ("simple", "no labels"):
            self.investigation_range_simple()  # configure widgets for setting investigation range
        else:
            self.investigation_range()  # configure widgets for setting investigation range
        self.investigation_status()  # provide message on status of investigation range
        if projvar.invran_station is None:  # if the investigation range is not set
            self.invran_not_set()  # investigation range not set screen
        else:
            if self.carrier_list:  # is the carrier is has contents
                self.show_carrierlist()  # show the carrier list
            else:  # if the carrier list is empty
                self.empty_carrierlist()  # the carrier list is empty screen
        self.bottom_of_frame()  # place necessary code to mainloop the window
        self.win.finish()  # close the window

    def set_dates(self):
        self.start_date = projvar.invran_date
        self.end_date = projvar.invran_date
        if projvar.invran_weekly_span:
            self.start_date = projvar.invran_date_week[0]
            self.end_date = projvar.invran_date_week[6]

    def make_stringvars(self):  # create stringvars
        self.start_year = StringVar(self.win.body)
        self.start_month = StringVar(self.win.body)
        self.start_day = StringVar(self.win.body)
        self.invran_date = StringVar(self.win.body)
        self.i_range = BooleanVar(self.win.body)
        self.invran = StringVar(self.win.body)
        self.station = StringVar(self.invest_frame)

    def get_carrierlist(self):  # call CarrierList to get Carrier Rec Set
        # get carrier list
        self.carrier_list = CarrierList(self.start_date, self.end_date, projvar.invran_station).get()

    def set_investigation_vars(self):  # set the stringvars for the investigation range
        now = datetime.now()
        self.start_month.set(now.month)  # default setting is now
        self.start_day.set(now.day)
        self.start_year.set(now.year)
        self.invran_date.set(now.strftime("%m/%d/%Y"))
        self.station.set("undefined")  # default value
        if projvar.invran_month:  # set month if a month is set
            self.start_month.set(projvar.invran_month)
        if projvar.invran_day:  # set day if a day is set
            self.start_day.set(projvar.invran_day)
        if projvar.invran_year:  # set year if a year is set
            self.start_year.set(projvar.invran_year)
        if projvar.invran_weekly_span:
            self.invran_date.set(projvar.invran_date_week[0].strftime("%m/%d/%Y"))
        elif projvar.invran_weekly_span is False:
            self.invran_date.set(projvar.invran_date.strftime("%m/%d/%Y"))
        if projvar.invran_station:
            self.station.set(projvar.invran_station)
        if projvar.invran_weekly_span is None:  # investigation range weekly/true or daily/false or none
            self.i_range.set(True)
            self.invran.set("week")
        elif not projvar.invran_weekly_span:  # if investigation range is daily
            self.i_range.set(False)
            self.invran.set("day")
        else:  # if investigation range is weekly
            self.i_range.set(True)
            self.invran.set("week")

    def get_stations_list(self):  # get a list of stations for station optionmenu
        self.stations_minus_outofstation = projvar.list_of_stations[:]
        self.stations_minus_outofstation.remove("out of station")
        if len(self.stations_minus_outofstation) == 0:
            self.stations_minus_outofstation.append("undefined")

    def get_invran_mode(self):  # get the investigation range mode
        sql = "SELECT tolerance FROM tolerances WHERE category = '%s'" % "invran_mode"
        results = inquire(sql)
        self.invran_result = results[0][0]

    def investigation_range_simple(self):
        Label(self.invest_frame, text="INVESTIGATION RANGE").grid(row=0, column=0, columnspan=2, sticky=W)
        if self.invran_result != "no labels":  # create a label row
            Label(self.invest_frame, text="Date: ", fg="grey").grid(row=1, column=0, sticky=W)
            Label(self.invest_frame, text="Range: ", fg="grey").grid(row=1, column=1, sticky=W)
            Label(self.invest_frame, text="Station: ", fg="grey").grid(row=1, column=2, sticky=W)
            Label(self.invest_frame, text="Set/Reset: ", fg="grey").grid(row=1, column=3, columnspan=2, sticky=W)
        # create widget row
        Entry(self.invest_frame, textvariable=self.invran_date, width=macadj(14, 9), justify='center')\
            .grid(row=2, column=0, padx=2)
        om_range = OptionMenu(self.invest_frame, self.invran, "week", "day")
        om_range.config(width=4)
        om_range.grid(row=2, column=1, sticky=W, padx=2)
        om_station = OptionMenu(self.invest_frame, self.station, *self.stations_minus_outofstation)
        om_station.config(width=macadj(31, 29))
        om_station.grid(row=2, column=2, sticky=W, padx=2)
        # set and reset buttons for investigation range
        Button(self.invest_frame, text="Set", width=macadj(5, 6), bg=macadj("green", "SystemButtonFace"),
               fg=macadj("white", "green"), command=lambda: self.call_globals()).grid(row=2, column=3, padx=2)
        Button(self.invest_frame, text="Reset", width=macadj(5, 6), bg=macadj("red", "SystemButtonFace"),
               fg=macadj("white", "red"), command=lambda: reset(self.win.topframe)).grid(row=2, column=4, padx=2)

    def investigation_range(self):  # configure widgets for setting investigation range
        Label(self.invest_frame, text="INVESTIGATION RANGE").grid(row=1, column=0, columnspan=2)
        om_month = OptionMenu(self.invest_frame, self.start_month, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12)
        om_month.config(width=2)
        om_month.grid(row=1, column=2)
        om_day = OptionMenu(self.invest_frame, self.start_day, "1", "2", "3", "4", "5", "6", "7", "8",
                            "9", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20",
                            "21", "22", "23", "24", "25", "26", "27", "28", "29", "30", "31")
        om_day.config(width=2)
        om_day.grid(row=1, column=3)
        date_year = Entry(self.invest_frame, width=6, textvariable=self.start_year)
        date_year.grid(row=1, column=4)
        Label(self.invest_frame, text="RANGE", width=macadj(6, 8)).grid(row=1, column=5)
        if projvar.invran_weekly_span is None:
            self.i_range.set(True)
        elif not projvar.invran_weekly_span:  # if investigation range is daily
            self.i_range.set(False)
        else:  # if investigation range is weekly
            self.i_range.set(True)
        Radiobutton(self.invest_frame, text="weekly", variable=self.i_range, value=True,
                    width=macadj(6, 7), anchor="w").grid(row=1, column=6)
        Radiobutton(self.invest_frame, text="daily", variable=self.i_range, value=False,
                    width=macadj(6, 7), anchor="w").grid(row=1, column=7)
        # set station option menu
        Label(self.invest_frame, text="STATION", anchor="w").grid(row=2, column=0, sticky=W)
        om = OptionMenu(self.invest_frame, self.station, *self.stations_minus_outofstation)
        om.config(width=macadj(40, 34))
        om.grid(row=2, column=1, columnspan=5, sticky=W)
        # set and reset buttons for investigation range
        Button(self.invest_frame, text="Set", width=macadj(8, 9), bg=macadj("green", "SystemButtonFace"),
               fg=macadj("white", "green"), command=lambda: set_globals(self.start_year.get(),
               self.start_month.get(), self.start_day.get(), self.i_range.get(), self.station.get(),
               self.win.topframe)).grid(row=2, column=6)
        Button(self.invest_frame, text="Reset", width=macadj(8, 9), bg=macadj("red", "SystemButtonFace"),
               fg=macadj("white", "red"), command=lambda: reset(self.win.topframe)).grid(row=2, column=7)

    def call_globals(self):
        msg_rear = "\n Dates must be formatted as \"mm/dd/yyyy\".\n" \
                   "Month must be expressed as number between 1 and 12.\n" \
                   "Day must be expressed as a number between 1 and 31.\n" \
                   "Year must be have four digits and be above 0010. "
        breakdown = BackSlashDateChecker(self.invran_date.get())
        if not breakdown.count_backslashes():
            msg = "The date must have 2 backslashes. " + msg_rear
            messagebox.showerror("Set Investigation Range", msg, parent=self.invest_frame)
            return
        breakdown.breaker()  # fully form the backslashdatechecker object
        if not breakdown.check_numeric():
            msg = "All month, day and year must be numbers. " + msg_rear
            messagebox.showerror("Set Investigation Range", msg, parent=self.invest_frame)
            return
        if not breakdown.check_minimums():
            msg = "All month, day and year must be greater than zero. " + msg_rear
            messagebox.showerror("Set Investigation Range", msg, parent=self.invest_frame)
            return
        if not breakdown.check_month():
            msg = "The value provided for the month is not acceptable. " + msg_rear
            messagebox.showerror("Set Investigation Range", msg, parent=self.invest_frame)
            return
        if not breakdown.check_day():
            msg = "The value provided for the day is not acceptable. " + msg_rear
            messagebox.showerror("Set Investigation Range", msg, parent=self.invest_frame)
            return
        if not breakdown.check_year():
            msg = "The value provided for the year is not acceptable. " + msg_rear
            messagebox.showerror("Set Investigation Range", msg, parent=self.invest_frame)
            return
        if not breakdown.valid_date():
            msg = "The investigation date is not valid. " + msg_rear
            messagebox.showerror("Set Investigation Range", msg, parent=self.invest_frame)
            return
        invest_range = True
        if self.invran.get() == "day":
            invest_range = False
        set_globals(breakdown.year, breakdown.month, breakdown.day, invest_range, self.station.get(), self.win.topframe)

    def investigation_status(self):  # provide message on status of investigation range
        # Investigation date SET/NOT SET notification
        if projvar.invran_weekly_span is None:
            Label(self.invest_frame, text="----> Investigation date/range not set", foreground="red") \
                .grid(row=3, column=0, columnspan=8, sticky="w")
        elif projvar.invran_weekly_span == 0:  # if the investigation range is one day
            f_date = projvar.invran_date.strftime("%a - %b %d, %Y")
            Label(self.invest_frame, text="---> Day Set: {} --> Pay Period: {}".format(f_date, projvar.pay_period),
                  foreground="red").grid(row=3, column=0, columnspan=8, sticky="w")
        else:
            # if the investigation range is weekly
            f_date = projvar.invran_date_week[0].strftime("%a - %b %d, %Y")
            end_f_date = projvar.invran_date_week[6].strftime("%a - %b %d, %Y")
            Label(self.invest_frame, text="---> Range Set: {0} through {1} --> Pay Period: {2}"
                  .format(f_date, end_f_date, projvar.pay_period),
                  foreground="red").grid(row=3, column=0, columnspan=8, sticky="w")

    def invran_not_set(self):  #investigation range is not set
        # Button(self.main_frame, text="Automatic Data Entry", width=30,
        #        command=lambda: call_indexers(self.win.topframe)).grid(row=0, column=1, pady=5)
        Button(self.main_frame, text="Automatic Data Entry", width=30,
               command=lambda: AutoDataEntry().run(self.win.topframe)).grid(row=0, column=1, pady=5)
        Button(self.main_frame, text="Informal C", width=30,
               command=lambda: informalc(self.win.topframe)).grid(row=1, column=1, pady=5)
        Button(self.main_frame, text="Quit", width=30, command=lambda: projvar.root.destroy())\
            .grid(row=2, column=1, pady=5)
        for i in range(25):
            Label(self.main_frame, text="").grid(row=4 + i, column=1)

    def empty_carrierlist(self):  # the carrier list is empty
        Label(self.main_frame, text="").grid(row=0, column=0)
        Label(self.main_frame, text="The carrier list is empty. ", font=macadj("bold", "Helvetica 18")) \
            .grid(row=1, column=0, sticky="w")
        Label(self.main_frame, text="").grid(row=2, column=0)
        Label(self.main_frame, text="Build the carrier list with the New Carrier feature\nor by running "
                                    "the Automatic Data Entry Feature.").grid(row=3, column=0)

    def show_carrierlist(self):  # investigation range is set and carrier list is not empty
        Label(self.main_frame, text="Name (click for Rings)", fg="grey").grid(row=0, column=1, sticky="w")
        Label(self.main_frame, text="List", fg="grey").grid(row=0, column=2, sticky="w")
        Label(self.main_frame, text="N/S", fg="grey").grid(row=0, column=3, sticky="w")
        Label(self.main_frame, text="Route", fg="grey").grid(row=0, column=4, sticky="w")
        Label(self.main_frame, text="Edit", fg="grey").grid(row=0, column=5, sticky="w")
        r = 1
        i = 0
        ii = 1
        for line in self.carrier_list:
            rec_count = 0
            # detect any out of station records and modify recset - function returns arrays with (startdate, carrier)
            line = CarrierRecFilter(line, self.start_date).detect_outofstation(projvar.invran_station)
            # if the row is even, then choose a color for it
            if i & 1:
                color = "light yellow"
            else:
                color = "white"
            for rec in line:
                if rec_count == 0:  # display the first row of carrier recs
                    Label(self.main_frame, text=ii).grid(row=r, column=0)  # display count
                    Button(self.main_frame, text=rec[1], width=macadj(25, 23), bg=color, anchor="w",
                           command=lambda x=rec: EnterRings(x[1]).start(self.win.topframe)).grid(row=r, column=1)
                    Button(self.main_frame, text="edit", width=4, bg=color, anchor="w",
                           command=lambda x=rec[1]: [self.win.topframe.destroy(), edit_carrier(x)]) \
                        .grid(row=r, column=5)
                    ii += 1
                else:  # display non first rows of carrier recs
                    dt = datetime.strptime(rec[0], "%Y-%m-%d %H:%M:%S")
                    Button(self.main_frame, text=dt.strftime("%a"), width=macadj(25, 23), bg=color, anchor="e")\
                        .grid(row=r, column=1)
                    Button(self.main_frame, text="", width=4, bg=color) \
                        .grid(row=r, column=5)
                if len(rec) > 2:  # because "out of station" recs only have two items
                    # list
                    Button(self.main_frame, text=rec[2], width=macadj(3, 4), bg=color, anchor="w").grid(row=r, column=2)
                    day_off = projvar.ns_code[rec[3]].lower()
                    Button(self.main_frame, text=day_off, width=4, bg=color, anchor="w").grid(row=r, column=3)  # nsday
                    Button(self.main_frame, text=rec[4], width=25, bg=color, anchor="w")\
                        .grid(row=r, column=4)  # route
                    rec_count += 1
                else:
                    Button(self.main_frame, text="out of station", width=35, bg=color)\
                        .grid(row=r, column=2, columnspan=3)
                r += 1
                rec_count += 1
            i += 1
            r += 1

    def pulldown_menu(self):  # create a pulldown menu, and add it to the menu bar
        menubar = Menu(self.win.topframe)
        # file menu
        basic_menu = Menu(menubar, tearoff=0)
        basic_menu.add_command(label="Save All", command=lambda: save_all(self.win.topframe))
        basic_menu.add_separator()
        basic_menu.add_command(label="New Carrier", command=lambda: input_carriers(self.win.topframe))
        basic_menu.add_command(label="Multiple Input", 
                               command=lambda dd="Sat", ss="name": mass_input(self.win.topframe, dd, ss))
        # basic_menu.add_command(label="Report Summary",
        # command=lambda: output_tab(self.win.topframe, self.carrier_list))

        basic_menu.add_command(label="Mandates Spreadsheet",
                               command=lambda r_rings="x": ImpManSpreadsheet().create(self.win.topframe))
        basic_menu.add_command(label="Over Max Spreadsheet",
                               command=lambda r_rings="x": OvermaxSpreadsheet().create(self.win.topframe))
        ot_date = projvar.invran_date  # build argument for ot equitability spreadsheet
        if projvar.invran_weekly_span:  # if the investigation range is weekly
            ot_date = projvar.invran_date_week[6]   # pass the last day of the investigation range as datetime
        basic_menu.add_command(label="OT Equitability Spreadsheet",
                               command=lambda: OTEquitSpreadsheet().create(self.win.topframe,
                                                                           ot_date, self.station.get()))
        listoptions = ("wal", "nl")  # ot distribution spreadsheet will show only work assignment and no list carriers
        basic_menu.add_command(label="OT Distribution Spreadsheet", command=lambda: OTDistriSpreadsheet()
                               .create(self.win.topframe, projvar.invran_date_week[0], self.station.get(),
                                       "weekly", listoptions))
        basic_menu.add_separator()
        basic_menu.add_command(label="OT Preferences", command=lambda: OtEquitability().create(self.win.topframe))
        basic_menu.add_command(label="OT Distribution", command=lambda: OtDistribution().create(self.win.topframe))
        basic_menu.add_separator()
        basic_menu.add_command(label="Informal C", command=lambda: informalc(self.win.topframe))
        basic_menu.add_separator()
        basic_menu.add_command(label="Location", command=lambda: Messenger(self.win.topframe).location_klusterbox())
        basic_menu.add_command(label="About Klusterbox", command=lambda: AboutKlusterbox().start(self.win.topframe))
        basic_menu.add_separator()
        basic_menu.add_command(label="View Out of Station",
                               command=lambda: set_globals(self.start_year.get(), self.start_month.get(),
                                                           self.start_day.get(), self.i_range.get(),
                                                           "out of station", self.win.topframe))
        basic_menu.add_separator()
        basic_menu.add_command(label="Quit", command=lambda: projvar.root.destroy())
        # gray out options if no investigation range is set
        if projvar.invran_day is None:
            basic_menu.entryconfig(2, state=DISABLED)
            basic_menu.entryconfig(3, state=DISABLED)
            basic_menu.entryconfig(4, state=DISABLED)
            basic_menu.entryconfig(5, state=DISABLED)
            basic_menu.entryconfig(6, state=DISABLED)
            basic_menu.entryconfig(7, state=DISABLED)
            basic_menu.entryconfig(10, state=DISABLED)
        menubar.add_cascade(label="Basic", menu=basic_menu)
        # automated menu
        automated_menu = Menu(menubar, tearoff=0)
        # automated_menu.add_command(label="Automatic Data Entry", command=lambda: call_indexers(self.win.topframe))
        automated_menu.add_command(label="Automatic Data Entry",
                                   command=lambda: AutoDataEntry().run(self.win.topframe))
        automated_menu.add_separator()
        automated_menu.add_command(label=" Auto Over Max Finder", command=lambda: max_hr(self.win.topframe))
        automated_menu.add_command(label="Everything Report Reader", command=lambda: ee_skimmer(self.win.topframe))
        automated_menu.add_command(label="Weekly Availability", command=lambda: wkly_avail(self.win.topframe))
        automated_menu.add_separator()
        automated_menu.add_command(label="PDF Converter", command=lambda: pdf_converter(self.win.topframe))
        automated_menu.add_command(label="PDF Splitter", command=lambda: pdf_splitter(self.win.topframe))
        menubar.add_cascade(label="Readers", menu=automated_menu)
        # reports menu
        reports_menu = Menu(menubar, tearoff=0)
        reports_menu.add_command(label="Carrier Route and NS Day", 
                                 command=lambda: Reports(self.win.topframe).rpt_carrier())
        reports_menu.add_command(label="Carrier Route", 
                                 command=lambda: Reports(self.win.topframe).rpt_carrier_route())
        reports_menu.add_command(label="Carrier NS Day", 
                                 command=lambda: Reports(self.win.topframe).rpt_carrier_nsday())
        reports_menu.add_command(label="Carrier by List", 
                                 command=lambda: Reports(self.win.topframe).rpt_carrier_by_list())
        reports_menu.add_command(label="Carrier Status History", 
                                 command=lambda: RptWin(self.win.topframe).rpt_find_carriers(projvar.invran_station))
        reports_menu.add_separator()
        reports_menu.add_command(label="Clock Rings Summary", 
                                 command=lambda: database_rings_report(self.win.topframe, projvar.invran_station))
        reports_menu.add_separator()
        reports_menu.add_command(label="Pay Period Guide Generator", 
                                 command=lambda: Reports(self.win.topframe).pay_period_guide())
        if projvar.invran_day is None:
            reports_menu.entryconfig(0, state=DISABLED)
            reports_menu.entryconfig(1, state=DISABLED)
            reports_menu.entryconfig(2, state=DISABLED)
            reports_menu.entryconfig(3, state=DISABLED)
            reports_menu.entryconfig(6, state=DISABLED)
        menubar.add_cascade(label="Reports", menu=reports_menu)
        # speedsheeet menu
        speed_menu = Menu(menubar, tearoff=0)
        speed_menu.add_command(label="Generate All Inclusive",
                               command=lambda: SpeedSheetGen(self.win.topframe, True).gen())
        speed_menu.add_command(label="Generate Carrier",
                               command=lambda: SpeedSheetGen(self.win.topframe, False).gen())
        speed_menu.add_command(label="Pre-check",
                               command=lambda: SpeedWorkBookGet().open_file(self.win.topframe, False))
        speed_menu.add_command(label="Input to Database",
                               command=lambda: SpeedWorkBookGet().open_file(self.win.topframe, True))
        if projvar.invran_day is None:
            speed_menu.entryconfig(0, state=DISABLED)
            speed_menu.entryconfig(1, state=DISABLED)
        speed_menu.add_separator()
        speed_menu.add_command(label="Cheatsheet",
                               command=lambda: OpenText().open_docs(self.win.body, 'cheatsheet.txt'))
        speed_menu.add_command(label="Instructions",
                               command=lambda: OpenText().open_docs(self.win.body, 'speedsheet_instructions.txt'))
        speed_menu.add_command(label="Speedsheet Archive", command=lambda: file_dialogue(dir_path('speedsheets')))
        speed_menu.add_command(label="Clear Archive",
                                command=lambda: remove_file_var(self.win.topframe, dir_path('speedsheets')))
        menubar.add_cascade(label="Speedsheet", menu=speed_menu)
        # archive menu
        reportsarchive_menu = Menu(menubar, tearoff=0)
        reportsarchive_menu.add_command(label="Mandates Spreadsheet",
                                        command=lambda: file_dialogue(dir_path('spreadsheets')))
        reportsarchive_menu.add_command(label="Over Max Spreadsheet",
                                        command=lambda: file_dialogue(dir_path('over_max_spreadsheet')))
        reportsarchive_menu.add_command(label="Speedsheets",
                                        command=lambda: file_dialogue(dir_path('speedsheets')))
        reportsarchive_menu.add_command(label="Over Max Finder",
                                        command=lambda: file_dialogue(dir_path('over_max')))
        reportsarchive_menu.add_command(label="OT Equitability",
                                        command=lambda: file_dialogue(dir_path('ot_equitability')))
        reportsarchive_menu.add_command(label="OT Distribution",
                                        command=lambda: file_dialogue(dir_path('ot_distribution')))
        reportsarchive_menu.add_command(label="Everything Report",
                                        command=lambda: file_dialogue(dir_path('ee_reader')))
        reportsarchive_menu.add_command(label="Weekly Availability",
                                        command=lambda: file_dialogue(dir_path('weekly_availability')))
        reportsarchive_menu.add_command(label="Pay Period Guide",
                                        command=lambda: file_dialogue(dir_path('pp_guide')))
        reportsarchive_menu.add_separator()
        cleararchive = Menu(reportsarchive_menu, tearoff=0)
        cleararchive.add_command(label="Mandates Spreadsheet",
                                 command=lambda: remove_file_var(self.win.topframe, dir_path('spreadsheets')))
        cleararchive.add_command(label="Over Max Spreadsheet",
                                 command=lambda: remove_file_var(self.win.topframe, dir_path('over_max_spreadsheet')))
        cleararchive.add_command(label="Speedsheets",
                                 command=lambda: remove_file_var(self.win.topframe, dir_path('speedsheets')))
        cleararchive.add_command(label="Over Max Finder",
                                 command=lambda: remove_file_var(self.win.topframe, dir_path('over_max')))
        cleararchive.add_command(label="OT Equitability",
                                 command=lambda: remove_file_var(self.win.topframe, dir_path('ot_equitability')))
        cleararchive.add_command(label="OT Distribution",
                                 command=lambda: remove_file_var(self.win.topframe, dir_path('ot_distribution')))
        cleararchive.add_command(label="Everything Report",
                                 command=lambda: remove_file_var(self.win.topframe, dir_path('ee_reader')))
        cleararchive.add_command(label="Weekly Availability",
                                 command=lambda: remove_file_var(self.win.topframe, dir_path('weekly_availability')))
        cleararchive.add_command(label="Pay Period Guide",
                                 command=lambda: remove_file_var(self.win.topframe, dir_path('pp_guide')))
        reportsarchive_menu.add_cascade(label="Clear Archive", menu=cleararchive)
        menubar.add_cascade(label="Archive", menu=reportsarchive_menu)
        # management menu
        management_menu = Menu(menubar, tearoff=0)
        # management_menu.add_command(label="GUI Configuration", command=lambda: gui_config(self.win.topframe))
        management_menu.add_command(label="GUI Configuration", command=lambda: GuiConfig(self.win.topframe).create())
        management_menu.add_separator()
        management_menu.add_command(label="List of Stations", command=lambda: station_list(self.win.topframe))
        management_menu.add_command(label="Tolerances", command=lambda: tolerances(self.win.topframe))
        management_menu.add_command(label="Spreadsheet Settings",
                                    command=lambda: SpreadsheetConfig().start(self.win.topframe))
        management_menu.add_command(label="NS Day Configurations", command=lambda: ns_config(self.win.topframe))
        if projvar.invran_day is None:
            management_menu.entryconfig(5, state=DISABLED)
        management_menu.add_command(label="Speedsheet Settings", 
                                    command=lambda: SpeedConfigGui(self.win.topframe).create())
        management_menu.add_separator()
        management_menu.add_command(label="Auto Data Entry Settings", 
                                    command=lambda: auto_data_entry_settings(self.win.topframe))
        management_menu.add_command(label="PDF Converter Settings", 
                                    command=lambda: pdf_converter_settings(self.win.topframe))
        management_menu.add_separator()
        management_menu.add_command(label="Database", 
                                    command=lambda: (self.win.topframe.destroy(), 
                                                     database_maintenance(self.win.topframe)))
        management_menu.add_command(label="Delete Carriers", 
                                    command=lambda: database_delete_carriers(self.win.topframe, projvar.invran_station))
        management_menu.add_command(label="Clean Carrier List", 
                                    command=lambda: carrier_list_cleaning(self.win.topframe))
        management_menu.add_command(label="Clean Rings", 
                                    command=lambda: clean_rings3_table())
        management_menu.add_separator()
        management_menu.add_command(label="Name Index", 
                                    command=lambda: (self.win.topframe.destroy(), name_index_screen()))
        management_menu.add_command(label="Station Index", command=lambda: station_index_mgmt(self.win.topframe))
        menubar.add_cascade(label="Management", menu=management_menu)
        projvar.root.config(menu=menubar)
        
    def bottom_of_frame(self):  # configure buttons on the bottom of the frame
        if projvar.invran_day is not None:
            Button(self.win.buttons, text="New Carrier", command=lambda: input_carriers(self.win.topframe),
                   width=macadj(13, 13)).pack(side=LEFT)
            Button(self.win.buttons, text="Multi Input",
                   command=lambda dd="Sat", ss="name": mass_input(self.win.topframe, dd, ss),
                   width=macadj(13, 13)).pack(side=LEFT)
            Button(self.win.buttons, text="Auto Data Entry", command=lambda: AutoDataEntry().run(self.win.topframe),
                   width=macadj(12, 12)).pack(side=LEFT)
            Button(self.win.buttons, text="Spreadsheet", width=macadj(13, 13),
                   command=lambda: ImpManSpreadsheet().create(self.win.topframe)).pack(side=LEFT)
            Button(self.win.buttons, text="Quit", width=macadj(13, 13), command=projvar.root.destroy).pack(side=LEFT)
        else:
            Label(self.win.buttons, text="").pack(side=LEFT)


if __name__ == "__main__":
    # declare all global variables
    global informalc_newroot
    global informalc_addframe
    global poe_add_pay_periods
    global poe_add_hours
    global poe_add_rate
    global poe_add_amount
    global informalc_poe_gadd
    global informalc_poe_lbox
    global allow_zero_top
    global allow_zero_bottom
    global skippers
    global current_tab
    global pb_flag
    global ade_flag
    setup_plaformvar()   # set up platform variable
    setup_dirs_by_platformvar()  # create directories if they don't exist
    DataBase().setup()  # set up the database
    projvar.root = Tk()  # initialize root window
    position_x = 100  # initialize position and size for root window
    position_y = 50
    size_x = 625
    size_y = 600
    projvar.root.title("KLUSTERBOX version {}".format(version))
    titlebar_icon(projvar.root)  # place icon in titlebar
    projvar.root.geometry("%dx%d+%d+%d" % (size_x, size_y, position_x, position_y))
    if len(projvar.list_of_stations) < 2:  # if there are no stations in the stations list
        StartUp().start()  # a start up screen for first time use
    else:
        remove_file(dir_path_check('report'))  # empty out folders
        remove_file(dir_path_check('infc_grv'))
        MainFrame().start()  # get the show on the road
