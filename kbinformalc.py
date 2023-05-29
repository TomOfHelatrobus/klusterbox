"""
a klusterbox module 
This module runs the Informal C, a program which allows users to record and track grievance settlements. It keeps 
track of grievance numbers, dates of the violation, at what level the settlement was signed, date of the signing, etc.
It also tracks which carrier was award what amount and if that settlement was paid. Report are available by carrier, 
by grievance as well as summaries. 
"""

# custom modules
import projvar  # defines project variables used in all modules.
from kbtoolbox import commit, dir_path, dir_path_check, dt_converter, find_pp, inquire, isfloat, macadj, \
    isint, NewWindow, titlebar_icon, informalc_date_checker, ProgressBarDe, ReportName, Handler, NameChecker, \
    GrievanceChecker, BackSlashDateChecker, Convert, DateTimeChecker
# standard libraries
from tkinter import messagebox, ttk, BOTH, BOTTOM, Button, Canvas, END, Entry, Frame, Label, LEFT, \
    Listbox, mainloop, NW, OptionMenu, Radiobutton, RIDGE, RIGHT, Scrollbar, StringVar, TclError, \
    Tk, VERTICAL, Y, Menu, filedialog, DISABLED
from kbreports import Archive
from datetime import datetime, timedelta
from shutil import rmtree
import os
import sys
import subprocess
import re
import time
from threading import Thread  # run load workbook while progress bar runs
# non standard libraries
from openpyxl import Workbook, load_workbook
from openpyxl.styles import NamedStyle, Font, Border, Side, Alignment
# define globals
global root  # used to hold the Tk() root for the new window used by all Informal C windows.
global pb_flag  #

""" this module has its own MakeWindow() class since it uses a different root. 
So it is not imported from kbtoolbar. """


def informalc_gen_clist(start, end, station):
    """ generates carrier list for informal c. """
    rec = None
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


def informalc_date_converter(date):
    """ be sure to run informalc date checker before using this """
    sd = date.get().split("/")
    return datetime(int(sd[2]), int(sd[0]), int(sd[1]))


class InformalC:
    """
    This is the home page of the Informal C program. It will open as a new window when launched in klusterbox.
    """
    def __init__(self):
        self.win = NewWindow(title="Informal C")
        global root
        root = self.win.root
        self.stationvar = None  # this is the stringvar for the station.
        self.station = None  # the station
        self.station_options = []  # the list of station options

    def informalc(self, frame):
        """ a master method for running the other methods in proper sequence. """
        self.clear_tempfolders()  # clear contents of temp folder
        self.get_station()  # this uses the investigation range station as the default
        self.get_station_options()  # this gets the list of stations
        self.build_tables()  # build needed tables if they do not exist.
        if not self.station_screen_autorouting():
            if not self.station:
                self.station_screen(frame)  # this allows the user to change/select the station
            else:
                self.menu_screen(frame)  # this fills the screen with widgets.

    def get_station(self):
        """ this sets the station to what was used for the klusterbox investigation range. """
        if projvar.invran_station:
            self.station = projvar.invran_station

    def get_station_options(self):
        """ this will get the station options ona place them in self.station_options"""
        for station in projvar.list_of_stations:
            self.station_options.append(station)
        if "out of station" in self.station_options:
            self.station_options.remove("out of station")

    def pulldown_menu(self):
        """ create a pulldown menu, and add it to the menu bar """
        menubar = Menu(self.win.topframe)
        # speedsheeet menu
        speed_menu = Menu(menubar, tearoff=0)
        speed_menu.add_command(label="Open Archive",
                               command=lambda: Archive().file_dialogue(dir_path('informalc_speedsheets')))
        speed_menu.add_command(label="Clear Archive",
                               command=lambda: Archive().remove_file_var(self.win.topframe, 'informalc_speedsheets'))
        speed_menu.add_command(label="Generate New Grievances",
                               command=lambda: SpeedSheetGen(self.win.topframe, self.station, "new").new())
        speed_menu.add_command(label="Generate Selected Grievances",
                               command=lambda: SpeedSheetGen(self.win.topframe, self.station, "selected").selected())
        speed_menu.add_command(label="Generate All Grievances",
                               command=lambda: SpeedSheetGen(self.win.topframe, self.station, "all").all())
        speed_menu.add_command(label="Pre-check",
                               command=lambda: SpeedWorkBookGet().open_file(self.win.topframe, False))
        speed_menu.add_command(label="Input to Database",
                               command=lambda: SpeedWorkBookGet().open_file(self.win.topframe, True))
        #  reports_menu.entryconfig(2, state=DISABLED)
        speed_menu.entryconfig(3, state=DISABLED)
        menubar.add_cascade(label="Speedsheet", menu=speed_menu)
        root.config(menu=menubar)

    @staticmethod
    def build_tables():
        """ build tables needed if they do no exist. """
        if os.path.isdir(dir_path_check('infc_grv')):  # clear contents of temp folder
            rmtree(dir_path_check('infc_grv'))
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

    @staticmethod
    def clear_tempfolders():
        """ clear contents of temp folder """
        if os.path.isdir(dir_path_check('infc_grv')):
            rmtree(dir_path_check('infc_grv'))

    def station_screen(self, frame):
        """ this allows the user to change/ select the station """
        self.win.create(frame)  # creates the screen object
        self.stationvar = StringVar(self.win.body)
        row = 0
        Label(self.win.body, text="Informal C", font=macadj("bold", "Helvetica 18")).grid(row=row, sticky="w")
        row += 1
        Label(self.win.body, text="The C is for Compliance").grid(row=row, sticky="w")
        row += 1
        Label(self.win.body, text="").grid(row=row)
        row += 1
        Label(self.win.body, text="Please select a Station: ", background=macadj("gray95", "white"),
              fg=macadj("black", "black"), width=macadj(22, 20), anchor="w", height=macadj(1, 1)). \
            grid(row=row, column=0, sticky="w")
        Label(self.win.body, text="", height=macadj(1, 2)).grid(row=row, column=1)
        row += 1
        self.stationvar.set("Select a Station")
        station_om = OptionMenu(self.win.body, self.stationvar, *self.station_options)
        station_om.config(width=macadj(40, 34))
        station_om.grid(row=row, column=0, columnspan=2, sticky="e")
        # self.station_screen_autorouting()
        # configure the submit button
        button_submit = Button(self.win.buttons)
        button_submit.config(text="Submit", width=20, command=lambda: self.station_screen_submit())
        if sys.platform == "win32":
            button_submit.config(anchor="w")
        button_submit.grid(row=0, column=1)
        # configure the "quit" button
        button_back = Button(self.win.buttons)
        button_back.config(text="Quit Informal C", width=20, command=lambda: self.win.root.destroy())
        if sys.platform == "win32":
            button_back.config(anchor="w")
        button_back.grid(row=0, column=0)
        self.win.finish()  # this commands the window to loop and persist.

    def station_screen_autorouting(self):
        """ this will automatically route the user depending on the amount of station options.
        One station option will automatically chose that option,
        Zero station options will show an error message and exit informal c. """
        if not self.station_options:
            messagebox.showerror("No Stations in Database",
                                 "There are no stations in the Klusterbox Database./n"
                                 "Proper function of Informal C requires at least one "
                                 "station to be entered into the Klusterbox Database. \n"
                                 "Please return to Klusterbox and enter a station.\n\n"
                                 "Informal C will end now. ",
                                 parent=self.win.body)
            self.win.root.destroy()  # terminate informal c
            return True
        if len(self.station_options) == 1:
            self.station = self.station_options[0]
            self.menu_screen(self.win.topframe)
            return True
        return False

    def station_screen_submit(self):
        """ this will update the station and route the user to the main menu
        or if no selection is made, there will be an error message and the screen will refresh. """

        if self.stationvar.get() == "Select a Station":
            messagebox.showerror("Prohibited Action",
                                 "Please select a station.",
                                 parent=self.win.body)
            self.station_screen(self.win.topframe)  # return to and refresh the station screen
        else:
            self.station = self.stationvar.get()
            self.menu_screen(self.win.topframe)

    def menu_screen(self, frame):
        """ the main screen for informal c. """
        self.win.create(frame)  # creates the screen object
        self.pulldown_menu()
        Label(self.win.body, text="Informal C", font=macadj("bold", "Helvetica 18")).grid(row=0, sticky="w")
        Label(self.win.body, text="The C is for Compliance").grid(row=1, sticky="w")
        Label(self.win.body, text="").grid(row=2)
        row = 3
        Button(self.win.body, text=" Enter New Grievance", width=30,
               command=lambda: self.NewGrievances(self).informalc_new(self.win.topframe)).grid(row=row, pady=5)
        row += 1
        Button(self.win.body, text="Enter New Settlement", width=30,
               command=lambda: self.New(self).informalc_new(self.win.topframe)).grid(row=row, pady=5)
        row += 1
        Button(self.win.body, text="Grievance Tracker", width=30,
               command=lambda: self.GrvList(self).grvlist_search(self.win.topframe)).grid(row=row, pady=5)
        row += 1
        Button(self.win.body, text="Tracker Settlement", width=30,
               command=lambda: self.GrvList(self).grvlist_search(self.win.topframe)).grid(row=row, pady=5)
        row += 1
        Button(self.win.body, text="Payout Entry", width=30,
               command=lambda: self.PayoutEntry(self).poe_search(self.win.topframe)).grid(row=row, pady=5)
        row += 1
        Button(self.win.body, text="Payout Report", width=30,
               command=lambda: self.PayoutReport(self).informalc_por(self.win.topframe)).grid(row=row, pady=5)
        row += 1
        Label(self.win.body, text="", width=70).grid(row=row)
        # configure the "quit" button
        button_back = Button(self.win.buttons)
        button_back.config(text="Quit Informal C", width=20, command=lambda: self.win.root.destroy())
        if sys.platform == "win32":
            button_back.config(anchor="w")
        button_back.grid(row=0, column=0)
        self.win.finish()  # this commands the window to loop and persist.

    class NewGrievances:
        """
        Allows the user to create new records of grievances.
        """

        def __init__(self, parent):
            self.parent = parent
            self.win = None
            self.msg = ""
            #  define the stringvars
            self.grievant = None  # 1
            self.station = None  # 2
            self.grv_no = None  # 3
            self.incident_start = None  # 4
            self.incident_end = None  # 5
            self.meeting_date = None  # 6
            self.issue = None  # 7
            self.article = None
            self.non_c = None  # 8  is the grievance a non compliance grievance

        def informalc_new(self, frame):
            """ master method for running other methods in proper order."""
            self.win = MakeWindow()
            self.get_stringvars()
            self.win.create(frame)
            self.build_screen()
            self.win.finish()

        def get_stringvars(self):
            """ initialize the stringvars """
            self.grievant = StringVar(self.win.body)
            self.station = StringVar(self.win.body)
            self.grv_no = StringVar(self.win.body)
            self.incident_start = StringVar(self.win.body)
            self.incident_end = StringVar(self.win.body)
            self.meeting_date = StringVar(self.win.body)
            self.issue = StringVar(self.win.body)
            self.article = StringVar(self.win.body)
            self.non_c = StringVar(self.win.body)

        def build_screen(self):
            """ screen for entering in new settlements. """
            row = 0
            Label(self.win.body, text="Enter New Grievance", font=macadj("bold", "Helvetica 18")) \
                .grid(row=row, column=0, columnspan=2, sticky="w")
            row += 1
            Label(self.win.body, text="").grid(row=row, column=0, sticky="w")
            row += 1

            Label(self.win.body, text="Grievant: ", background=macadj("gray95", "white"),
                  fg=macadj("black", "black"), width=macadj(22, 20), anchor="w", height=macadj(1, 1)) \
                .grid(row=row, column=0, sticky="w")
            Entry(self.win.body, textvariable=self.grievant, justify='right', width=macadj(20, 15)) \
                .grid(row=row, column=1, sticky="w")
            row += 1

            Label(self.win.body, text="Grievance Number: ", background=macadj("gray95", "white"),
                  fg=macadj("black", "black"), width=macadj(22, 20), anchor="w", height=macadj(1, 1)) \
                .grid(row=row, column=0, sticky="w")
            Entry(self.win.body, textvariable=self.grv_no, justify='right', width=macadj(20, 15)) \
                .grid(row=row, column=1, sticky="w")
            row += 1
            # start and end dates
            Label(self.win.body, text="Incident Date").grid(row=row, column=0, sticky="w")
            row += 1
            # start date
            Label(self.win.body, text="  Start (mm/dd/yyyy): ", background=macadj("gray95", "white"),
                  fg=macadj("black", "black"), width=macadj(22, 20), anchor="w") \
                .grid(row=row, column=0, sticky="w")
            Entry(self.win.body, textvariable=self.incident_start, justify='right', width=macadj(20, 15)) \
                .grid(row=row, column=1, sticky="w")
            row += 1
            # end date
            Label(self.win.body, text="  End (mm/dd/yyyy): ", background=macadj("gray95", "white"),
                  fg=macadj("black", "black"), width=macadj(22, 20), anchor="w", height=macadj(1, 1)) \
                .grid(row=row, column=0, sticky="w")
            Entry(self.win.body, textvariable=self.incident_end, justify='right', width=macadj(20, 15)) \
                .grid(row=row, column=1, sticky="w")
            row += 1

            Label(self.win.body, text="Meeting Date (mm/dd/yyyy): ", background=macadj("gray95", "white"),
                  fg=macadj("black", "black"), width=macadj(22, 20), anchor="w", height=macadj(1, 1)) \
                .grid(row=row, column=0, sticky="w")
            Entry(self.win.body, textvariable=self.meeting_date, justify='right', width=macadj(20, 15)) \
                .grid(row=row, column=1, sticky="w")
            row += 1
            Label(self.win.body, text="Non compliance", background=macadj("gray95", "white"),
                  fg=macadj("black", "black"), width=macadj(22, 20), anchor="w", height=macadj(1, 1)) \
                .grid(row=row, column=0, sticky="w")

            # issue
            Label(self.win.body, text="Issue: ", background=macadj("gray95", "white"),
                  fg=macadj("black", "black"), width=macadj(22, 20), anchor="w", height=macadj(1, 1)) \
                .grid(row=row, column=0, sticky="w")
            Label(self.win.body, text="", height=macadj(1, 2)).grid(row=row, column=1)
            row += 1
            Entry(self.win.body, textvariable=self.issue, width=macadj(48, 36), justify='right') \
                .grid(row=row, column=0, sticky="w", columnspan=2)
            row += 1

            Label(self.win.body, text="Article: ", background=macadj("gray95", "white"),
                  fg=macadj("black", "black"), width=macadj(22, 20), anchor="w", height=macadj(1, 1)) \
                .grid(row=row, column=0, sticky="w")
            Entry(self.win.body, textvariable=self.article, justify='right', width=macadj(20, 15)) \
                .grid(row=row, column=1, sticky="w")
            row += 1

            Label(self.win.body, text="", height=macadj(1, 1)).grid(row=row, column=0)
            row += 1

            Label(self.win.body, text=self.msg, fg="red", height=macadj(1, 1)) \
                .grid(row=row, column=0, columnspan=2, sticky="w")
            row += 1

            # configure buttons on the bottom of the screen
            button_alignment = macadj("w", "center")
            Button(self.win.buttons, text="Go Back", width=macadj(19, 18), anchor=button_alignment,
                   command=lambda: self.parent.informalc(self.win.topframe)).grid(row=0, column=0)
            Button(self.win.buttons, text="Enter", width=macadj(19, 18), anchor=button_alignment,
                   command=lambda: self.informalc_new_apply()).grid(row=0, column=1)

        def informalc_new_apply(self):
            """ applies changes to settlement information. """
            check = self.informalc_check_grv()
            if check:
                dates = [self.incident_start, self.incident_end, self.meeting_date]
                in_start = datetime(1, 1, 1)
                in_end = datetime(1, 1, 1)
                m_date = datetime(1, 1, 1)
                dt_dates = [in_start, in_end, m_date]
                i = 0
                for date in dates:
                    date = date.get()  # get the data from the stringvar
                    date = date.strip()  # strip out white space
                    d = date.split("/")  # convert the date into an array
                    new_date = datetime(int(d[2].lstrip("0")), int(d[0].lstrip("0")), int(d[1].lstrip("0")))
                    dt_dates[i] = new_date
                    i += 1
                if dt_dates[0] > dt_dates[1]:
                    messagebox.showerror("Data Entry Error",
                                         "The Incident Start Date can not be later that the Incident End "
                                         "Date.",
                                         parent=self.win.topframe)
                    return
                if dt_dates[0] > dt_dates[2]:
                    messagebox.showerror("Data Entry Error",
                                         "The Incident Start Date can not be later that the Date Signed.",
                                         parent=self.win.topframe)
                    return
                sql = "SELECT grv_no FROM informalc_grv"
                results = inquire(sql)
                existing_grv = []
                for result in results:
                    for grv in result:
                        existing_grv.append(grv)
                grv_no = self.grv_no.get()  # get the value from the string var
                grv_no = grv_no.strip()  # strip out the white space
                grv_no = grv_no.lower()  # convert to all lowercase
                issue = self.issue.get()
                issue = issue.strip()
                issue = issue.lower()
                if grv_no in existing_grv:
                    messagebox.showerror("Data Entry Error",
                                         "The Grievance Number {} is already present in the database. You can not "
                                         "create a duplicate.".format(grv_no),
                                         parent=self.win.topframe)
                    return
                sql = "INSERT INTO informalc_grievance (grievant, grv_no, startdate, enddate, meetingdate, " \
                      "station, issue, article) " \
                      "VALUES('%s','%s','%s','%s','%s','%s','%s','%s')" % \
                      (self.grievant, grv_no, dt_dates[0], dt_dates[1], dt_dates[2], self.station.get(), issue,
                       self.article)
                commit(sql)
                self.msg = "Grievance Settlement Added: #{}.".format(grv_no)
                self.informalc_new(self.win.topframe)

        def informalc_check_grv(self):
            """ checks the grievance number. """
            if self.station.get() == "Select a Station":
                messagebox.showerror("Invalid Data Entry",
                                     "You must select a station.",
                                     parent=self.win.topframe)
                return False
            grv_no = self.grv_no.get()  # get the value from the stringvar
            grv_no = grv_no.strip()  # strip out white space
            grv_no = grv_no.lower()
            if grv_no == "":
                messagebox.showerror("Invalid Data Entry",
                                     "You must enter a grievance number",
                                     parent=self.win.topframe)
                return False
            if re.search('[^1234567890abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ]', grv_no):
                messagebox.showerror("Invalid Data Entry",
                                     "The grievance number can only contain numbers and letters. No other "
                                     "characters are allowed",
                                     parent=self.win.topframe)
                return False
            if len(grv_no) < 4:
                messagebox.showerror("Invalid Data Entry",
                                     "The grievance number must be at least four characters long",
                                     parent=self.win.topframe)
                return False
            if len(grv_no) > 20:
                messagebox.showerror("Invalid Data Entry",
                                     "The grievance number must not exceed 20 characters in length.",
                                     parent=self.win.topframe)
                return False
            return self.informalc_check_grv_2()

        def informalc_check_grv_2(self):
            """ checks the information for informalc grievances. """
            dates = [self.incident_start.get(), self.incident_end.get(), self.meeting_date.get()]
            date_ids = ("starting incident date", "ending incident date", "date signed")
            i = 0
            for date in dates:
                date = date.strip()
                d = date.split("/")
                if len(d) != 3:
                    messagebox.showerror("Invalid Data Entry",
                                         "The date for the {} is not properly formatted.".format(date_ids[i]),
                                         parent=self.win.topframe)
                    return False
                for num in d:
                    if not num.isnumeric():
                        messagebox.showerror("Invalid Data Entry",
                                             "The month, day and year for the {} "
                                             "must be numeric.".format(date_ids[i]),
                                             parent=self.win.topframe)
                        return False
                if len(d[0]) > 2:
                    messagebox.showerror("Invalid Data Entry",
                                         "The month for the {} must be no more than two digits"
                                         " long.".format(date_ids[i]),
                                         parent=self.win.topframe)
                    return False
                if len(d[1]) > 2:
                    messagebox.showerror("Invalid Data Entry",
                                         "The day for the {} must be no more than two digits"
                                         " long.".format(date_ids[i]),
                                         parent=self.win.topframe)
                    return False
                if len(d[2]) != 4:
                    messagebox.showerror("Invalid Data Entry",
                                         "The year for the {} must be four digits long."
                                         .format(date_ids[i]),
                                         parent=self.win.topframe)
                    return False
                try:
                    date = datetime(int(d[2]), int(d[0]), int(d[1]))
                    valid_date = True
                    if date:
                        # use project variable to absorb error from unused try/except statement.
                        projvar.try_absorber = True
                except ValueError:
                    valid_date = False
                if not valid_date:
                    messagebox.showerror("Invalid Data Entry",
                                         "The date entered for {} is not a valid date."
                                         .format(date_ids[i]),
                                         parent=self.win.topframe)
                    return False
                i += 1
            if self.issue.get().strip() != "":
                if not all(x.isalnum() or x.isspace() for x in self.issue.get()):
                    messagebox.showerror("Invalid Data Entry",
                                         "The Description can only contain letters and numbers. No "
                                         "special characters are allowed.",
                                         parent=self.win.topframe)
                    return False
                if len(self.issue.get()) > 40:
                    messagebox.showerror("Invalid Data Entry",
                                         "The Description is limited to no more than 40 characters. ",
                                         parent=self.win.topframe)
                    return False
            return True

    class New:
        """
        Allows the user to create new records of settlements.
        """
        def __init__(self, parent):
            self.parent = parent
            self.win = None
            self.msg = ""
            #  define the stringvars 
            self.grv_no = None
            self.incident_start = None
            self.incident_end = None
            self.date_signed = None
            self.station = None
            self.gats_number = None
            self.docs = None
            self.description = None
            self.lvl = None

        def informalc_new(self, frame):
            """ master method for running other methods in proper order."""
            self.win = MakeWindow()
            self.get_stringvars()
            self.win.create(frame)
            self.build_screen()
            self.win.finish()
            
        def get_stringvars(self):
            """ initialize the stringvars """
            self.grv_no = StringVar(self.win.body)
            self.incident_start = StringVar(self.win.body)
            self.incident_end = StringVar(self.win.body)
            self.date_signed = StringVar(self.win.body)
            self.lvl = StringVar(self.win.body)
            self.station = StringVar(self.win.body)
            self.gats_number = StringVar(self.win.body)
            self.docs = StringVar(self.win.body)
            self.description = StringVar(self.win.body)
            
        def build_screen(self):    
            """ screen for entering in new settlements. """
            Label(self.win.body, text="New Settlement", font=macadj("bold", "Helvetica 18"))\
                .grid(row=0, column=0, sticky="w")
            Label(self.win.body, text="").grid(row=1, column=0, sticky="w")
            Label(self.win.body, text="Grievance Number: ", background=macadj("gray95", "white"),
                  fg=macadj("black", "black"), width=macadj(22, 20), anchor="w", height=macadj(1, 1)) \
                .grid(row=2, column=0, sticky="w")
            
            Entry(self.win.body, textvariable=self.grv_no, justify='right', width=macadj(20, 15)) \
                .grid(row=2, column=1, sticky="w")
            Label(self.win.body, text="Incident Date").grid(row=3, column=0, sticky="w")
            Label(self.win.body, text="  Start (mm/dd/yyyy): ", background=macadj("gray95", "white"),
                  fg=macadj("black", "black"), width=macadj(22, 20), anchor="w") \
                .grid(row=4, column=0, sticky="w")
            
            Entry(self.win.body, textvariable=self.incident_start, justify='right', width=macadj(20, 15)) \
                .grid(row=4, column=1, sticky="w")
            Label(self.win.body, text="  End (mm/dd/yyyy): ", background=macadj("gray95", "white"),
                  fg=macadj("black", "black"), width=macadj(22, 20), anchor="w", height=macadj(1, 1)) \
                .grid(row=5, column=0, sticky="w")
            
            Entry(self.win.body, textvariable=self.incident_end, justify='right', width=macadj(20, 15)) \
                .grid(row=5, column=1, sticky="w")
            Label(self.win.body, text="Date Signed (mm/dd/yyyy): ", background=macadj("gray95", "white"),
                  fg=macadj("black", "black"), width=macadj(22, 20), anchor="w", height=macadj(1, 1)) \
                .grid(row=6, column=0, sticky="w")
            
            Entry(self.win.body, textvariable=self.date_signed, justify='right', width=macadj(20, 15)) \
                .grid(row=6, column=1, sticky="w")
            # select level
            Label(self.win.body, text="Settlement Level: ", background=macadj("gray95", "white"),
                  fg=macadj("black", "black"), width=macadj(22, 20), anchor="w", height=macadj(1, 1)) \
                .grid(row=7, column=0, sticky="w")  # select settlement level
            
            lvl_options = ("informal a", "formal a", "step b", "pre arb", "arbitration")
            lvl_om = OptionMenu(self.win.body, self.lvl, *lvl_options)
            lvl_om.config(width=macadj(13, 13))
            lvl_om.grid(row=7, column=1)
            self.lvl.set("informal a")
            Label(self.win.body, text="Station: ", background=macadj("gray95", "white"),  # select a station
                  fg=macadj("black", "black"), width=macadj(22, 20), anchor="w", height=macadj(1, 1)). \
                grid(row=8, column=0, sticky="w")
            Label(self.win.body, text="", height=macadj(1, 2)).grid(row=8, column=1)
            
            self.station.set("Select a Station")
            station_options = projvar.list_of_stations
            if "out of station" in station_options:
                station_options.remove("out of station")
            station_om = OptionMenu(self.win.body, self.station, *station_options)
            station_om.config(width=macadj(40, 34))
            station_om.grid(row=9, column=0, columnspan=2, sticky="e")
            Label(self.win.body, text="GATS Number: ", background=macadj("gray95", "white"),
                  fg=macadj("black", "black"), width=macadj(22, 20), anchor="w", height=macadj(1, 1)) \
                .grid(row=10, column=0, sticky="w")  # enter gats number
            
            Entry(self.win.body, textvariable=self.gats_number, justify='right', width=macadj(20, 15)) \
                .grid(row=10, column=1, sticky="w")
            Label(self.win.body, text="Documentation?: ", background=macadj("gray95", "white"),
                  fg=macadj("black", "black"), width=macadj(22, 20), anchor="w", height=macadj(1, 1)) \
                .grid(row=11, column=0, sticky="w")  # select documentation
            
            doc_options = ("moot", "no", "partial", "yes", "incomplete", "verified")
            docs_om = OptionMenu(self.win.body, self.docs, *doc_options)
            docs_om.config(width=macadj(13, 13))
            docs_om.grid(row=11, column=1)
            self.docs.set("no")
            Label(self.win.body, text="Description: ", background=macadj("gray95", "white"),
                  fg=macadj("black", "black"), width=macadj(22, 20), anchor="w", height=macadj(1, 1)) \
                .grid(row=15, column=0, sticky="w")
            Label(self.win.body, text="", height=macadj(1, 2)).grid(row=15, column=1)
            
            Entry(self.win.body, textvariable=self.description, width=macadj(48, 36), justify='right') \
                .grid(row=16, column=0, sticky="w", columnspan=2)
            Label(self.win.body, text="", height=macadj(1, 1)).grid(row=17, column=0)
            Label(self.win.body, text=self.msg, fg="red", height=macadj(1, 1))\
                .grid(row=18, column=0, columnspan=2, sticky="w")
            button_alignment = macadj("w", "center")
            Button(self.win.buttons, text="Go Back", width=macadj(19, 18), anchor=button_alignment,
                   command=lambda: self.parent.informalc(self.win.topframe)).grid(row=0, column=0)
            Button(self.win.buttons, text="Enter", width=macadj(19, 18), anchor=button_alignment,
                   command=lambda: self.informalc_new_apply()).grid(row=0, column=1)

        def informalc_new_apply(self):
            """ applies changes to settlement information. """
            check = self.informalc_check_grv()
            if check:
                dates = [self.incident_start, self.incident_end, self.date_signed]
                in_start = datetime(1, 1, 1)
                in_end = datetime(1, 1, 1)
                d_sign = datetime(1, 1, 1)
                dt_dates = [in_start, in_end, d_sign]
                i = 0
                for date in dates:
                    date = date.get()  # get the data from the stringvar
                    date = date.strip()  # strip out white space
                    d = date.split("/")  # convert the date into an array
                    new_date = datetime(int(d[2].lstrip("0")), int(d[0].lstrip("0")), int(d[1].lstrip("0")))
                    dt_dates[i] = new_date
                    i += 1
                if dt_dates[0] > dt_dates[1]:
                    messagebox.showerror("Data Entry Error",
                                         "The Incident Start Date can not be later that the Incident End "
                                         "Date.",
                                         parent=self.win.topframe)
                    return
                if dt_dates[0] > dt_dates[2]:
                    messagebox.showerror("Data Entry Error",
                                         "The Incident Start Date can not be later that the Date Signed.",
                                         parent=self.win.topframe)
                    return
                sql = "SELECT grv_no FROM informalc_grv"
                results = inquire(sql)
                existing_grv = []
                for result in results:
                    for grv in result:
                        existing_grv.append(grv)
                grv_no = self.grv_no.get()  # get the value from the string var
                grv_no = grv_no.strip()  # strip out the white space
                grv_no = grv_no.lower()  # convert to all lowercase
                description = self.description.get()
                description = description.strip()
                description = description.lower()
                if grv_no in existing_grv:
                    messagebox.showerror("Data Entry Error",
                                         "The Grievance Number {} is already present in the database. You can not "
                                         "create a duplicate.".format(grv_no),
                                         parent=self.win.topframe)
                    return
                sql = "INSERT INTO informalc_grv (grv_no, indate_start, indate_end, date_signed, station, " \
                      "gats_number, docs, description, level) " \
                      "VALUES('%s','%s','%s','%s','%s','%s','%s','%s','%s')" % \
                      (grv_no, dt_dates[0], dt_dates[1], dt_dates[2], self.station.get(),
                       self.gats_number.get().strip(), self.docs.get(), description, self.lvl.get())
                commit(sql)
                self.msg = "Grievance Settlement Added: #{}.".format(grv_no)
                self.informalc_new(self.win.topframe)

        def informalc_check_grv(self):
            """ checks the grievance number. """
            if self.station.get() == "Select a Station":
                messagebox.showerror("Invalid Data Entry",
                                     "You must select a station.",
                                     parent=self.win.topframe)
                return False
            grv_no = self.grv_no.get()  # get the value from the stringvar
            grv_no = grv_no.strip()  # strip out white space
            grv_no = grv_no.lower()
            if grv_no == "":
                messagebox.showerror("Invalid Data Entry",
                                     "You must enter a grievance number",
                                     parent=self.win.topframe)
                return False
            if re.search('[^1234567890abcdefghijklmnopqrstuvwxyz:ABCDEFGHIJKLMNOPQRSTUVWXYZ,]', grv_no):
                messagebox.showerror("Invalid Data Entry",
                                     "The grievance number can only contain numbers and letters. No other "
                                     "characters are allowed",
                                     parent=self.win.topframe)
                return False
            if len(grv_no) < 8:
                messagebox.showerror("Invalid Data Entry",
                                     "The grievance number must be at least eight characters long",
                                     parent=self.win.topframe)
                return False
            if len(grv_no) > 20:
                messagebox.showerror("Invalid Data Entry",
                                     "The grievance number must not exceed 20 characters in length.",
                                     parent=self.win.topframe)
                return False
            return self.informalc_check_grv_2()

        def informalc_check_grv_2(self):
            """ checks the information for informalc grievances. """
            dates = [self.incident_start.get(), self.incident_end.get(), self.date_signed.get()]
            date_ids = ("starting incident date", "ending incident date", "date signed")
            i = 0
            for date in dates:
                date = date.strip()
                d = date.split("/")
                if len(d) != 3:
                    messagebox.showerror("Invalid Data Entry",
                                         "The date for the {} is not properly formatted.".format(date_ids[i]),
                                         parent=self.win.topframe)
                    return False
                for num in d:
                    if not num.isnumeric():
                        messagebox.showerror("Invalid Data Entry",
                                             "The month, day and year for the {} "
                                             "must be numeric.".format(date_ids[i]),
                                             parent=self.win.topframe)
                        return False
                if len(d[0]) > 2:
                    messagebox.showerror("Invalid Data Entry",
                                         "The month for the {} must be no more than two digits"
                                         " long.".format(date_ids[i]),
                                         parent=self.win.topframe)
                    return False
                if len(d[1]) > 2:
                    messagebox.showerror("Invalid Data Entry",
                                         "The day for the {} must be no more than two digits"
                                         " long.".format(date_ids[i]),
                                         parent=self.win.topframe)
                    return False
                if len(d[2]) != 4:
                    messagebox.showerror("Invalid Data Entry",
                                         "The year for the {} must be four digits long."
                                         .format(date_ids[i]),
                                         parent=self.win.topframe)
                    return False
                try:
                    date = datetime(int(d[2]), int(d[0]), int(d[1]))
                    valid_date = True
                    if date:
                        # use project variable to absorb error from unused try/except statement.
                        projvar.try_absorber = True  
                except ValueError:
                    valid_date = False
                if not valid_date:
                    messagebox.showerror("Invalid Data Entry",
                                         "The date entered for {} is not a valid date."
                                         .format(date_ids[i]),
                                         parent=self.win.topframe)
                    return False
                i += 1
            if len(self.gats_number.get()) > 50:
                messagebox.showerror("Invalid Data Entry",
                                     "The GATS number is limited to no more than 20 characters. ",
                                     parent=self.win.topframe)
                return False
            if self.gats_number.get().strip() != "":
                if not all(x.isalnum() or x.isspace() for x in self.gats_number.get()):
                    messagebox.showerror("Invalid Data Entry",
                                         "The GATS number can only contain letters and numbers. No "
                                         "special characters are allowed.",
                                         parent=self.win.topframe)
                    return False
            if self.description.get().strip() != "":
                if not all(x.isalnum() or x.isspace() for x in self.description.get()):
                    messagebox.showerror("Invalid Data Entry",
                                         "The Description can only contain letters and numbers. No "
                                         "special characters are allowed.",
                                         parent=self.win.topframe)
                    return False
                if len(self.description.get()) > 40:
                    messagebox.showerror("Invalid Data Entry",
                                         "The Description is limited to no more than 40 characters. ",
                                         parent=self.win.topframe)
                    return False
            return True
    
    class GrvList:
        """ 
        creates a display where users can access all settlements in the database. From there users can generate reports
        and make edits. 
        """
        def __init__(self, parent):
            self.parent = parent
            self.win = None
            # initialized the stringvars used for the search criteria.
            self.incident_date = None
            self.incident_start = None
            self.incident_end = None
            self.signing_date = None
            self.signing_start = None
            self.signing_end = None
            self.station = None
            self.set_lvl = None
            self.level = None
            self.gats = None
            self.have_gats = None
            self.docs = None
            self.have_docs = None
            self.sql = None  # var for sql query. hold in variable so search can be duplicated.
            self.search_result = None  # var for the search result
            self.companion_root = None  # a companion root window for entering carrier awards
            self.companion_frame = None  # a companion frame for the addframe root.
            #  vars for the edit methods
            self.grv_num = None
            self.msg = ""
            # vars for the edit stringvars
            self.grv_no = None
            self.edit_incident_start = None
            self.edit_incident_end = None
            self.date_signed = None
            self.lvl = None
            self.station = None
            self.gats_number = None
            self.edit_docs = None
            self.description = None
            # vars for add award
            self.var_id = None
            self.var_name = None
            self.var_hours = None
            self.var_rate = None
            self.var_amount = None

        def grvlist_search(self, frame):
            """ master method for running other methods in proper order. """
            self.win = MakeWindow()
            self.win.create(frame)
            self.get_stringvars()
            self.build_screen()
            self.build_buttons()
            self.win.finish()
            
        def get_stringvars(self):
            """ initialize varibles """
            self.station = StringVar(self.win.topframe)
            self.incident_date = StringVar(self.win.topframe)
            self.incident_start = StringVar(self.win.topframe)
            self.incident_end = StringVar(self.win.topframe)
            self.signing_date = StringVar(self.win.topframe)
            self.signing_start = StringVar(self.win.topframe)
            self.signing_end = StringVar(self.win.topframe)
            self.set_lvl = StringVar(self.win.topframe)
            self.level = StringVar(self.win.topframe)
            self.gats = StringVar(self.win.topframe)
            self.have_gats = StringVar(self.win.topframe)
            self.docs = StringVar(self.win.topframe)
            self.have_docs = StringVar(self.win.topframe)
            
        def build_screen(self):
            """ builds page for searching grievance settlements. """
            Label(self.win.body, text="Informal C: Settlement Search Criteria", font=macadj("bold", "Helvetica 18")) \
                .grid(row=0, columnspan=6, sticky="w")
            Label(self.win.body, text=" ").grid(row=1, columnspan=6)
            # select station
            Label(self.win.body, text=" Station ", background=macadj("gray95", "white"), fg=macadj("black", "black"),
                  anchor="w", width=macadj(14, 12)).grid(row=2, column=0, columnspan=3, sticky="w")
            station_options = projvar.list_of_stations
            if "out of station" in station_options:
                station_options.remove("out of station")
            station_om = OptionMenu(self.win.body, self.station, *station_options)
            station_om.config(width=macadj(38, 31))
            station_om.grid(row=2, column=3, columnspan=3, sticky="e")
            self.station.set("Select a Station")
            Label(self.win.body, text="Search For", fg="grey").grid(row=3, column=0, columnspan=2, sticky="w")
            Label(self.win.body, text="Category", fg="grey").grid(row=3, column=3)
            Label(self.win.body, text="Start", fg="grey").grid(row=3, column=4)
            Label(self.win.body, text="End", fg="grey").grid(row=3, column=5)
            # select for starting date
            Radiobutton(self.win.body, text="yes", variable=self.incident_date, value='yes', width=macadj(2, 4)) \
                .grid(row=4, column=0, sticky="w")
            Radiobutton(self.win.body, text="no", variable=self.incident_date, value='no', width=macadj(2, 4)) \
                .grid(row=4, column=1, sticky="w")
            Label(self.win.body, text="", width=macadj(2, 4)).grid(row=4, column=2)
            Label(self.win.body, text=" Incident Dates", background=macadj("gray95", "white"),
                  fg=macadj("black", "black"), anchor="w", width=14).grid(row=4, column=3, sticky="w")
            Entry(self.win.body, textvariable=self.incident_start, width=macadj(12, 8), justify='right')\
                .grid(row=4, column=4)
            Entry(self.win.body, textvariable=self.incident_end, width=macadj(12, 8), justify='right')\
                .grid(row=4, column=5)
            self.incident_date.set('no')
            # select for signing date
            Radiobutton(self.win.body, text="yes", variable=self.signing_date, value='yes', width=macadj(2, 4)) \
                .grid(row=5, column=0, sticky="w")
            Radiobutton(self.win.body, text="no", variable=self.signing_date, value='no', width=macadj(2, 4)) \
                .grid(row=5, column=1, sticky="w")
            Label(self.win.body, text=" Signing Dates", background=macadj("gray95", "white"),
                  fg=macadj("black", "black"), anchor="w", width=14).grid(row=5, column=3, sticky="w")
            Entry(self.win.body, textvariable=self.signing_start, width=macadj(12, 8), justify='right')\
                .grid(row=5, column=4)
            Entry(self.win.body, textvariable=self.signing_end, width=macadj(12, 8), justify='right')\
                .grid(row=5, column=5)
            self.signing_date.set('no')
            # select for settlement level
            Radiobutton(self.win.body, text="yes", variable=self.set_lvl, value='yes', width=macadj(2, 4)) \
                .grid(row=6, column=0, sticky="w")
            Radiobutton(self.win.body, text="no", variable=self.set_lvl, value='no', width=macadj(2, 4)) \
                .grid(row=6, column=1, sticky="w")
            self.set_lvl.set("no")
            Label(self.win.body, text=" Settlement Level ", background=macadj("gray95", "white"),
                  fg=macadj("black", "black"), anchor="w", width=14, height=1).grid(row=6, column=3, sticky="w")
            lvl_options = ("informal a", "formal a", "step b", "pre-arb", "arbitration")
            lvl_om = OptionMenu(self.win.body, self.level, *lvl_options)
            lvl_om.config(width=macadj(20, 16))
            lvl_om.grid(row=6, column=4, columnspan=3, sticky="e")
            self.level.set("informal a")
            # select for gats number
            Radiobutton(self.win.body, text="yes", variable=self.gats, value='yes', width=macadj(2, 4)) \
                .grid(row=7, column=0, sticky="w")
            Radiobutton(self.win.body, text="no", variable=self.gats, value='no', width=macadj(2, 4)) \
                .grid(row=7, column=1, sticky="w")
            Label(self.win.body, text=" GATS Number", background=macadj("gray95", "white"), fg=macadj("black", "black"),
                  anchor="w", width=14, height=1).grid(row=7, column=3, sticky="w")
            gats_options = ("no", "yes")
            gats_om = OptionMenu(self.win.body, self.have_gats, *gats_options)
            gats_om.config(width=macadj(10, 8))
            gats_om.grid(row=7, column=4, columnspan=3, sticky="e")
            self.have_gats.set('no')
            self.gats.set('no')
            # select for documentation
            Radiobutton(self.win.body, text="yes", variable=self.docs, value='yes', width=macadj(2, 4)) \
                .grid(row=9, column=0, sticky="w")
            Radiobutton(self.win.body, text="no", variable=self.docs, value='no', width=macadj(2, 4)) \
                .grid(row=9, column=1, sticky="w")
            Label(self.win.body, text=" Documentation", background=macadj("gray95", "white"),
                  fg=macadj("black", "black"), anchor="w", width=14, height=1).grid(row=9, column=3, sticky="w")
            doc_options = ("moot", "no", "partial", "yes", "incomplete", "verified")
            docs_om = OptionMenu(self.win.body, self.have_docs, *doc_options)
            docs_om.config(width=macadj(10, 8))
            docs_om.grid(row=9, column=4, columnspan=3, sticky="e")
            self.have_docs.set('no')
            self.docs.set("no")
            Label(self.win.body, text="").grid(row=13)
        
        def build_buttons(self):
            """ build the buttons on the bottom of the screen. """
            button_alignment = macadj("w", "center")
            Button(self.win.buttons, text="Search", width=20, anchor=button_alignment,
                   command=lambda: self.grvlist_apply()).grid(row=0, column=1)
            Button(self.win.buttons, text="Go Back", width=20, anchor=button_alignment,
                   command=lambda: self.parent.informalc(self.win.topframe)).grid(row=0, column=0)

        def grvlist_apply(self):
            """ applies changes to the grievance list after a check. """
            conditions = []
            if self.incident_date.get() == "yes":
                if not informalc_date_checker(self.win.topframe, self.incident_start, "starting incident date"):
                    return
                if not informalc_date_checker(self.win.topframe, self.incident_end, "ending incident date"):
                    return
                d = self.incident_start.get().split("/")
                start = datetime(int(d[2]), int(d[0]), int(d[1]))
                d = self.incident_end.get().split("/")
                end = datetime(int(d[2]), int(d[0]), int(d[1]))
                if start > end:
                    messagebox.showerror("Invalid Data Entry",
                                         "Your starting incident date must be earlier than your "
                                         "ending incident date.",
                                         parent=self.win.topframe)
                    return
                to_add = "indate_start > '{}' and indate_end < '{}'".format(start, end)
                conditions.append(to_add)
            if self.signing_date.get() == "yes":
                if not informalc_date_checker(self.win.topframe, self.signing_start, "starting signing date"):
                    return
                if not informalc_date_checker(self.win.topframe, self.signing_end, "ending signing date"):
                    return
                d = self.signing_start.get().split("/")
                start = datetime(int(d[2]), int(d[0]), int(d[1]))
                d = self.signing_end.get().split("/")
                end = datetime(int(d[2]), int(d[0]), int(d[1]))
                if start > end:
                    messagebox.showerror("Invalid Data Entry",
                                         "Your starting signing date must be earlier than your "
                                         "ending signing date.",
                                         parent=self.win.topframe)
                    return
                to_add = "date_signed BETWEEN '{}' AND '{}'".format(start, end)
                conditions.append(to_add)
            if self.station.get() == "Select a Station":
                messagebox.showerror("Invalid Station",
                                     "You must select a station.",
                                     parent=self.win.topframe)
                return
            to_add = "station = '{}'".format(self.station.get())
            conditions.append(to_add)
            if self.set_lvl.get() == "yes":
                to_add = "level = '{}'".format(self.level.get())
                conditions.append(to_add)

            if self.gats.get() == "yes":
                if self.have_gats.get() == "yes":
                    to_add = "gats_number IS NOT ''"
                    conditions.append(to_add)
                if self.have_gats.get() == "no":
                    to_add = "gats_number IS ''"
                    conditions.append(to_add)
            if self.docs.get() == "yes":
                to_add = "docs = '{}'".format(self.have_docs.get())
                conditions.append(to_add)
            where_str = ""
            for i in range(len(conditions)):
                where_str += "{}".format(conditions[i])
                if i + 1 < len(conditions):
                    where_str += " and "
            self.sql = "SELECT * FROM informalc_grv WHERE {} ORDER BY date_signed DESC".format(where_str)
            self.search_result = inquire(self.sql)
            self.grvlist_result()

        def grvlist_result(self):
            """ shows the results for the specified range."""
            frame = self.win  # preserve the window object from being destroyed by the next line.
            self.win = MakeWindow()
            self.win.create(frame.topframe)
            Label(self.win.body, text="Informal C: Search Results", font=macadj("bold", "Helvetica 18")) \
                .grid(row=0, column=0, columnspan=4, sticky="w")
            Label(self.win.body, text="").grid(row=1)
            if len(self.search_result) == 0:
                Label(self.win.body, text="The search has no results.").grid(row=2, column=0, columnspan=4)
            else:
                Label(self.win.body, text="Grievance Number", fg="grey", anchor="w").grid(row=2, column=1, sticky="w")
                Label(self.win.body, text="Incident Start", fg="grey", anchor="w").grid(row=2, column=2, sticky="w")
                Label(self.win.body, text="Incident End", fg="grey", anchor="w").grid(row=2, column=3, sticky="w")
                Label(self.win.body, text="Date Signed", fg="grey", anchor="w").grid(row=2, column=4, sticky="w")
            row = 3
            ii = 1
            for r in self.search_result:
                """ 
                Show search results. loop once for each settlement. 
                """
                Label(self.win.body, text=str(ii), anchor="w", width=macadj(4, 2)).grid(row=row, column=0)
                Button(self.win.body, text=" " + r[0], anchor="w", width=macadj(14, 12), relief=RIDGE)\
                    .grid(row=row, column=1)
                in_start = datetime.strptime(r[1], '%Y-%m-%d %H:%M:%S')
                in_end = datetime.strptime(r[2], '%Y-%m-%d %H:%M:%S')
                sign_date = datetime.strptime(r[3], '%Y-%m-%d %H:%M:%S')
                Button(self.win.body, text=in_start.strftime("%b %d, %Y"), width=macadj(11, 10),
                       anchor="w", relief=RIDGE) \
                    .grid(row=row, column=2)
                Button(self.win.body, text=in_end.strftime("%b %d, %Y"), width=macadj(11, 10),
                       anchor="w", relief=RIDGE) \
                    .grid(row=row, column=3)
                Button(self.win.body, text=sign_date.strftime("%b %d, %Y"), width=macadj(11, 10),
                       anchor="w", relief=RIDGE) \
                    .grid(row=row, column=4)
                Button(self.win.body, text="Edit", width=macadj(6, 5), relief=RIDGE,
                       command=lambda x=r[0]: self.edit(x)) \
                    .grid(row=row, column=5)
                Button(self.win.body, text="Report", width=macadj(6, 5), relief=RIDGE,
                       command=lambda x=r: self.rptbygrv(x)).grid(row=row, column=6)
                Button(self.win.body, text=macadj("Enter Awards", "Awards"), width=macadj(10, 6), relief=RIDGE,
                       command=lambda x=r[0]: self.addaward(x))\
                    .grid(row=row, column=7)
                row += 1
                Label(self.win.body, text="         {}".format(r[7]), anchor="w", fg="grey") \
                    .grid(row=row, column=1, columnspan=5, sticky="w")
                row += 1
                ii += 1
            """ 
            define the buttons at the bottom of the page: 
            """
            Button(self.win.buttons, text="Go Back", width=macadj(16, 13),
                   command=lambda: self.grvlist_search(self.win.topframe)) \
                .grid(row=0, column=0)
            Label(self.win.buttons, text="Report: ", width=macadj(16, 11)).grid(row=0, column=1)
            Button(self.win.buttons, text="By Settlements", width=macadj(16, 13),
                   command=lambda: self.rptgrvsum()) \
                .grid(row=0, column=2)
            Button(self.win.buttons, text="By Carriers", width=macadj(16, 13),
                   command=lambda: self.bycarriers()) \
                .grid(row=0, column=3)
            Button(self.win.buttons, text="By Carrier", width=macadj(16, 13),
                   command=lambda: self.bycarrier()) \
                .grid(row=0, column=4)
            Label(self.win.buttons, text="Summary: ", width=macadj(16, 11)).grid(row=1, column=1)
            Button(self.win.buttons, text="By Settlements", width=macadj(16, 13),
                   command=lambda: self.grvlist_setsum()).grid(row=1, column=2)
            Button(self.win.buttons, text="Carrier List", width=macadj(16, 13),
                   command=lambda: self.RptCarrierId(self).run()).grid(row=1, column=3)
            self.win.finish()
            
        def edit(self, grv_num):
            """ screen for editing informalc grievances. """
            frame = self.win
            # self.result = self.search_result
            self.grv_num = grv_num
            self.win = MakeWindow()
            self.win.create(frame.topframe)
            self.get_edit_stringvars()
            self.set_edit_stringvars()
            self.build_edit()
            self.win.finish()
            
        def get_edit_stringvars(self):
            """ define the stringvars for the edit """
            self.grv_no = StringVar(self.win.topframe)
            self.edit_incident_start = StringVar(self.win.topframe)
            self.edit_incident_end = StringVar(self.win.topframe)
            self.date_signed = StringVar(self.win.topframe)
            self.lvl = StringVar(self.win.topframe)
            self.station = StringVar(self.win.topframe)
            self.gats_number = StringVar(self.win.topframe)
            self.edit_docs = StringVar(self.win.topframe)
            self.description = StringVar(self.win.topframe)

        def set_edit_stringvars(self):
            """ set the values to the stringvars with values from the database. """
            sql = "SELECT * FROM informalc_grv WHERE grv_no='%s'" % self.grv_num
            search = inquire(sql)
            if search:
                in_start = datetime.strptime(search[0][1], '%Y-%m-%d %H:%M:%S')
                in_end = datetime.strptime(search[0][2], '%Y-%m-%d %H:%M:%S')
                sign_date = datetime.strptime(search[0][3], '%Y-%m-%d %H:%M:%S')
                self.edit_incident_start.set(in_start.strftime("%m/%d/%Y"))
                self.edit_incident_end.set(in_end.strftime("%m/%d/%Y"))
                self.date_signed.set(sign_date.strftime("%m/%d/%Y"))
                self.station.set(search[0][4])
                self.gats_number.set(search[0][5])
                self.edit_docs.set(search[0][6])
                self.description.set(search[0][7])
                if search[0][8] is None:
                    self.lvl.set("unknown")
                else:
                    self.lvl.set(search[0][8])
            
        def build_edit(self):
            """ build the body of the edit screen for settlements in the grievance list results."""
            Label(self.win.body, text="Informal C: Edit Grievance", font=macadj("bold", "Helvetica 18"))\
                .grid(row=0, columnspan=2, sticky="w")
            Label(self.win.body, text="").grid(row=1)
            Label(self.win.body, text="Grievance Number: ", background=macadj("gray95", "grey"),
                  fg=macadj("black", "white"), width=macadj(22, 20), anchor="w", height=macadj(1, 1)) \
                .grid(row=2, column=0, sticky="w")
            
            Entry(self.win.body, textvariable=self.grv_no, justify='right', width=macadj(20, 15)) \
                .grid(row=2, column=1, sticky="w")
            Button(self.win.body, width=9, text="update",
                   command=lambda: self.grvchange(self.grv_num, self.grv_no)).grid(row=3, column=1, sticky="e")
            self.grv_no.set(self.grv_num)
            Label(self.win.body, text="Incident Date", background=macadj("gray95", "grey"),
                  fg=macadj("black", "white"), width=macadj(22, 20), anchor="w", height=macadj(1, 1)) \
                .grid(row=4, column=0, sticky="w")
            Label(self.win.body, text="  Start (mm/dd/yyyy): ", background=macadj("gray95", "grey"),
                  fg=macadj("black", "white"), width=macadj(22, 20), anchor="w", height=macadj(1, 1)) \
                .grid(row=5, column=0, sticky="w")
            
            Entry(self.win.body, textvariable=self.edit_incident_start, justify='right', width=macadj(20, 15)) \
                .grid(row=5, column=1, sticky="w")
            Label(self.win.body, text="  End (mm/dd/yyyy): ", background=macadj("gray95", "grey"),
                  fg=macadj("black", "white"), width=macadj(22, 20), anchor="w", height=macadj(1, 1)) \
                .grid(row=6, column=0, sticky="w")
            
            Entry(self.win.body, textvariable=self.edit_incident_end, justify='right', width=macadj(20, 15)) \
                .grid(row=6, column=1, sticky="w")
            Label(self.win.body, text="Date Signed (mm/dd/yyyy): ", background=macadj("gray95", "grey"),
                  fg=macadj("black", "white"), width=macadj(22, 20), anchor="w", height=macadj(1, 1)) \
                .grid(row=7, column=0, sticky="w")
            
            Entry(self.win.body, textvariable=self.date_signed, justify='right', width=macadj(20, 15)) \
                .grid(row=7, column=1, sticky="w")
            Label(self.win.body, text="Settlement Level: ", background=macadj("gray95", "grey"),
                  fg=macadj("black", "white"), width=macadj(22, 20), anchor="w", height=macadj(1, 1)) \
                .grid(row=8, column=0, sticky="w")  # select settlement level
            
            lvl_options = ("informal a", "formal a", "step b", "pre arb", "arbitration")
            lvl_om = OptionMenu(self.win.body, self.lvl, *lvl_options)
            lvl_om.config(width=13)
            lvl_om.grid(row=8, column=1)
            Label(self.win.body, text="Station: ", background=macadj("gray95", "grey"),
                  fg=macadj("black", "white"), width=macadj(22, 20), anchor="w", height=macadj(1, 1)) \
                .grid(row=9, column=0, sticky="w")  # select a station
            
            station_options = projvar.list_of_stations
            if "out of station" in station_options:
                station_options.remove("out of station")
            station_om = OptionMenu(self.win.body, self.station, *station_options)
            station_om.config(width=macadj(40, 34))
            station_om.grid(row=10, column=0, columnspan=2, sticky="e")
            Label(self.win.body, text="GATS Number: ", background=macadj("gray95", "grey"),
                  fg=macadj("black", "white"), width=macadj(22, 20), anchor="w", height=macadj(1, 1)) \
                .grid(row=11, column=0, sticky="w")
            
            Entry(self.win.body, textvariable=self.gats_number, justify='right', width=macadj(20, 15)) \
                .grid(row=11, column=1, sticky="w")
            Label(self.win.body, text="Documentation: ", background=macadj("gray95", "grey"),
                  fg=macadj("black", "white"), width=macadj(22, 20), anchor="w", height=macadj(1, 1)) \
                .grid(row=12, column=0, sticky="w")
            
            doc_options = ("moot", "no", "partial", "yes", "incomplete", "verified")
            docs_om = OptionMenu(self.win.body, self.edit_docs, *doc_options)
            docs_om.config(width=13)
            docs_om.grid(row=12, column=1)
            Label(self.win.body, text="Description: ", background=macadj("gray95", "grey"),
                  fg=macadj("black", "white"), width=macadj(22, 20), anchor="w", height=macadj(1, 1)) \
                .grid(row=16, column=0, sticky="w")
            
            Entry(self.win.body, textvariable=self.description, width=macadj(47, 36), justify='right') \
                .grid(row=17, column=0, sticky="e", columnspan=2)
            Label(self.win.body, text="").grid(row=18, column=0)

            Label(self.win.body, text=" ").grid(row=20)
            Label(self.win.body, text="Delete Grievance", background=macadj("gray95", "grey"),
                  fg=macadj("black", "white"), width=macadj(22, 20), anchor="w", height=macadj(1, 1)) \
                .grid(row=21, column=0, sticky="w")
            Button(self.win.body, text="Delete", width=9,
                   command=lambda: self.delete(self.win.topframe, self.grv_no)) \
                .grid(row=21, column=1, sticky="e")
            Label(self.win.body, text=" ").grid(row=22)
            Label(self.win.body, text=self.msg, fg="red", anchor="w").grid(row=23, column=0, columnspan=5, sticky="w")
            self.msg = ""  # reset the message to empty string
            Button(self.win.buttons, text="Go Back", width=macadj(19, 18),
                   command=lambda: self.grvlist_result()).grid(row=0, column=0)
            Button(self.win.buttons, text="Enter", width=macadj(19, 18),
                   command=lambda: self.edit_apply()).grid(row=0, column=1)

        def edit_apply(self):
            """  check then edit informalc peticulars. """
            dates = [self.edit_incident_start.get(), self.edit_incident_end.get(), self.date_signed.get()]
            date_ids = ("starting incident date", "ending incident date", "date signed")
            i = 0
            for date in dates:
                date = date.strip()
                d = date.split("/")
                if len(d) != 3:
                    messagebox.showerror("Invalid Data Entry",
                                         "The date for the {} is not properly formatted.".format(date_ids[i]),
                                         parent=self.win.topframe)
                    return
                for num in d:
                    if not num.isnumeric():
                        messagebox.showerror("Invalid Data Entry",
                                             "The month, day and year for the {} "
                                             "must be numeric.".format(date_ids[i]),
                                             parent=self.win.topframe)
                        return
                if len(d[0]) > 2:
                    messagebox.showerror("Invalid Data Entry",
                                         "The month for the {} must be no more than two digits"
                                         " long.".format(date_ids[i]),
                                         parent=self.win.topframe)
                    return
                if len(d[1]) > 2:
                    messagebox.showerror("Invalid Data Entry",
                                         "The day for the {} must be no more than two digits"
                                         " long.".format(date_ids[i]),
                                         parent=self.win.topframe)
                    return
                if len(d[2]) != 4:
                    messagebox.showerror("Invalid Data Entry",
                                         "The year for the {} must be four digits long."
                                         .format(date_ids[i]),
                                         parent=self.win.topframe)
                    return
                try:
                    date = datetime(int(d[2]), int(d[0]), int(d[1]))
                    valid_date = True
                    if date:
                        # use project variable to absorb error from unused try/except statement.
                        projvar.try_absorber = True
                except ValueError:
                    valid_date = False
                if not valid_date:
                    messagebox.showerror("Invalid Data Entry",
                                         "The date entered for {} is not a valid date."
                                         .format(date_ids[i]),
                                         parent=self.win.topframe)
                    return
                i += 1
            if len(self.gats_number.get()) > 50:
                messagebox.showerror("Invalid Data Entry",
                                     "The GATS number is limited to no more than 20 characters. ",
                                     parent=self.win.topframe)
                return
            if self.gats_number.get().strip() != "":
                if not all(x.isalnum() or x.isspace() for x in self.gats_number.get()):
                    messagebox.showerror("Invalid Data Entry",
                                         "The GATS number can only contain letters and numbers. No "
                                         "special characters are allowed.",
                                         parent=self.win.topframe)
                    return
            if self.description.get().strip() != "":
                if not all(x.isalnum() or x.isspace() for x in self.description.get()):
                    messagebox.showerror("Invalid Data Entry",
                                         "The Description can only contain letters and numbers. No "
                                         "special characters are allowed.",
                                         parent=self.win.topframe)
                    return
                if len(self.description.get()) > 40:
                    messagebox.showerror("Invalid Data Entry",
                                         "The Description is limited to no more than 40 characters. ",
                                         parent=self.win.topframe)
                    return
            dates = [self.edit_incident_start.get(), self.edit_incident_end.get(), self.date_signed.get()]
            in_start = datetime(1, 1, 1)
            in_end = datetime(1, 1, 1)
            d_sign = datetime(1, 1, 1)
            dt_dates = [in_start, in_end, d_sign]
            i = 0
            for date in dates:
                date = date.strip()
                d = date.split("/")
                new_date = datetime(int(d[2].lstrip("0")), int(d[0].lstrip("0")), int(d[1].lstrip("0")))
                dt_dates[i] = new_date
                i += 1
            if dt_dates[0] > dt_dates[1]:
                messagebox.showerror("Data Entry Error",
                                     "The Incident Start Date can not be later that the Incident End "
                                     "Date.",
                                     parent=self.win.topframe)
                return
            if dt_dates[0] > dt_dates[2]:
                messagebox.showerror("Data Entry Error",
                                     "The Incident Start Date can not be later that the Date Signed.",
                                     parent=self.win.topframe)
                return
            description = self.description.get()
            description = description.strip()
            description = description.lower()
            sql = "UPDATE informalc_grv SET indate_start='%s',indate_end='%s',date_signed='%s',station='%s'," \
                  "gats_number='%s', docs='%s',description='%s', level='%s' WHERE grv_no='%s'" % \
                  (dt_dates[0], dt_dates[1], dt_dates[2], self.station.get(), self.gats_number.get().strip(),
                   self.edit_docs.get(), description, self.lvl.get(), self.grv_no.get())
            commit(sql)
            messagebox.showerror("Sucessful Update",
                                 "Grievance number: {} succesfully updated.".format(self.grv_no.get()),
                                 parent=self.win.topframe)
            self.update_search_results()  # update the search results
            # self.grvlist_search(self.win.topframe)
            self.grvlist_result()  # return to the grievance list results with previous search criteria

        def update_search_results(self):
            """ update the search results of the grievance list """
            self.search_result = inquire(self.sql)

        def delete(self, frame, grv_no):
            """ deletes a record and associated records for a grievance. """
            check = messagebox.askokcancel("Delete Grievance",
                                           "Are you sure you want to delete his grievance and all the "
                                           "data associated with it?",
                                           parent=self.win.topframe)
            if not check:
                return
            else:
                sql = "DELETE FROM informalc_grv WHERE grv_no='%s'" % grv_no.get()
                commit(sql)
                self.grvlist_search(frame)

        def grvchange(self, old_num, new_num):
            """ change the grievance number. check grv number and input it into the informalc_grv table. """
            l_passed_result = [list(x) for x in self.search_result]  # chg tuple of tuples to list of lists
            if messagebox.askokcancel("Grievance Number Change",
                                      "This will change the grievance number from {} to {} in all "
                                      "records. Are you sure you want to proceed?".format(old_num, new_num.get()),
                                      parent=self.win.topframe):
                new_number = new_num.get()  # get the value from the passed new_num stringvar
                new_number = new_number.strip()  # strip out all whitespace in front and back
                new_number = new_number.lower()  # change all upper case to lower case
                if new_number == "":
                    messagebox.showerror("Invalid Data Entry",
                                         "You must enter a grievance number",
                                         parent=self.win.topframe)
                    return "fail"
                if not new_number.isalnum():
                    messagebox.showerror("Invalid Data Entry",
                                         "The grievance number can only contain numbers and letters. No other "
                                         "characters are allowed",
                                         parent=self.win.topframe)
                    return "fail"
                if len(new_number) < 8:
                    messagebox.showerror("Invalid Data Entry",
                                         "The grievance number must be at least eight characters long",
                                         parent=self.win.topframe)
                    return "fail"
                if len(new_number) > 16:
                    messagebox.showerror("Invalid Data Entry",
                                         "The grievance number must not exceed 16 characters in length.",
                                         parent=self.win.topframe)
                    return "fail"
                sql = "SELECT grv_no FROM informalc_grv WHERE grv_no = '%s'" % new_number
                result = inquire(sql)
                if result:
                    messagebox.showerror("Grievance Number Error",
                                         "This number is already being used for another grievance.",
                                         parent=self.win.topframe)
                    return "fail"

                sql = "UPDATE informalc_grv SET grv_no = '%s' WHERE grv_no = '%s'" % (new_number, old_num)
                commit(sql)
                sql = "UPDATE informalc_awards SET grv_no = '%s' WHERE grv_no = '%s'" % (new_number, old_num)
                commit(sql)
                for record in l_passed_result:
                    if record[0] == old_num:
                        record[0] = new_number
                self.msg = "The grievance number has been changed."
                self.search_result = l_passed_result[:]
                self.edit(new_number)

        def addaward(self, grv_no):
            """ adds carrier to the add award screen from the companion screen. """
            self.informalc_root(grv_no)
            self.addaward2(grv_no)

        def informalc_root(self, grv_no):
            """ creates a companion window for selecting carrier names. """
            start = None
            end = None
            station = None
            self.companion_root = Tk()
            self.companion_root.title("KLUSTERBOX")
            titlebar_icon(self.companion_root)  # place icon in titlebar
            x_position = projvar.root.winfo_x() + 450
            y_position = projvar.root.winfo_y() - 25
            self.companion_root.geometry("%dx%d+%d+%d" % (240, 600, x_position, y_position))
            topframe = Frame(self.companion_root)
            topframe.pack()
            buttons = Canvas(topframe)  # button bar
            buttons.pack(fill=BOTH, side=BOTTOM)
            Label(topframe, text="Add Carriers", font=macadj("bold", "Helvetica 18")).pack(anchor="w")
            Label(topframe, text="").pack()
            scrollbar = Scrollbar(topframe, orient=VERTICAL)
            listbox = Listbox(topframe, selectmode="multiple", yscrollcommand=scrollbar.set)
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
            Button(buttons, text="Add Carrier", width=10,
                   command=lambda: (self.addnames(grv_no, c_list, listbox.curselection()),
                                    self.addaward2(grv_no))) \
                .pack(side=LEFT, anchor="w")
            Button(buttons, text="Clear", width=10,
                   command=lambda: (self.companion_root.destroy(), self.informalc_root(grv_no))) \
                .pack(side=LEFT, anchor="w")
            Button(buttons, text="Close", width=10,
                   command=lambda: (self.companion_root.destroy())).pack(side=LEFT, anchor="w")

        @staticmethod
        def addnames(grv_no, c_list, listbox):
            """ inserts names into informal c awards table. """
            for index in listbox:
                sql = "INSERT INTO informalc_awards (grv_no,carrier_name,hours,rate,amount) " \
                      "VALUES('%s','%s','%s','%s','%s')" \
                      % (grv_no, c_list[int(index)], '', '', '')
                commit(sql)

        def addaward2(self, grv_no):
            """ creates a screen which allows a user to adds the awards to a settlement. """
            frame = self.win  # preserve the window object from being destroy by the next line. 
            self.win = MakeWindow()
            self.win.create(frame.topframe)
            self.companion_frame = self.win.topframe
            Label(self.win.body, text="Add/Update Settlement Awards", font=macadj("bold", "Helvetica 18")) \
                .grid(row=0, column=0, sticky="w", columnspan=4)
            Label(self.win.body, text=" ".format(self.companion_frame)).grid(row=1, column=0)
            Label(self.win.body, text="   Grievance Number: {}".format(grv_no), fg="blue") \
                .grid(row=2, column=0, sticky="w", columnspan=4)
            sql = "SELECT grv_no,rowid,carrier_name,hours,rate,amount FROM informalc_awards WHERE grv_no ='%s' " \
                  "ORDER BY carrier_name" % grv_no
            result = inquire(sql)
            # initialize arrays for names
            self.var_id = []
            self.var_name = []
            self.var_hours = []
            self.var_rate = []
            self.var_amount = []
            if len(result) == 0:
                Label(self.win.body, text="No records in database").grid(row=3)
            else:
                Label(self.win.body, text="Carrier", fg="grey", padx=10).grid(row=3, column=0, sticky="w")
                Label(self.win.body, text="Hours", fg="grey", padx=10).grid(row=3, column=1, sticky="w")
                Label(self.win.body, text="Rate", fg="grey", padx=10).grid(row=3, column=2, sticky="w")
                Label(self.win.body, text="Amount", fg="grey", padx=10).grid(row=3, column=3, sticky="w")
                i = 0
                r = 4
                for res in result:
                    self.var_id.append(StringVar(self.win.topframe))  # add to arrays
                    self.var_name.append(StringVar(self.win.topframe))
                    self.var_hours.append(StringVar(self.win.topframe))
                    self.var_rate.append(StringVar(self.win.topframe))
                    self.var_amount.append(StringVar(self.win.topframe))
                    Label(self.win.body, text=res[2], anchor="w", width=16)\
                        .grid(row=r, column=0, sticky="w", padx=10)  # display name widget
                    Entry(self.win.body, textvariable=self.var_hours[i], width=8)\
                        .grid(row=r, column=1, padx=10)  # display hours widget
                    Entry(self.win.body, textvariable=self.var_rate[i], width=8)\
                        .grid(row=r, column=2, padx=10)  # display rate widget
                    Entry(self.win.body, textvariable=self.var_amount[i], width=8)\
                        .grid(row=r, column=3, padx=10)  # display amount widget
                    Button(self.win.body, text="delete",
                           command=lambda ident=res[1]: self.deletename(grv_no, ident)) \
                        .grid(row=r, column=4, padx=10)  # display the delete button
                    self.var_id[i].set(res[1])  # set the textvariables
                    self.var_name[i].set(res[2])
                    self.var_hours[i].set(res[3])
                    self.var_rate[i].set(res[4])
                    self.var_amount[i].set(res[5])
                    r += 1
                    i += 1
            Button(self.win.buttons, text="Go Back", width=15,
                   command=lambda: self.call_grvlist_result()) \
                .grid(row=0, column=0)
            Button(self.win.buttons, text="Apply", width=15,
                   command=lambda: self.addaward_apply(grv_no)).grid(row=0, column=1)
            self.win.finish()

        def addaward_apply(self, grv_no):
            """ checks and adds records to the informal c add awards table. """
            pb_label = Label(self.win.buttons, text="Updating Changes: ")  # make label for progress bar
            pb_label.grid(row=0, column=2)
            pb = ttk.Progressbar(self.win.buttons, length=200, mode="determinate")  # create progress bar
            pb.grid(row=0, column=3)
            pb["maximum"] = len(self.var_id)  # set length of progress bar
            pb.start()
            ii = 0
            for i in range(len(self.var_id)):
                pb["value"] = ii  # increment progress bar
                id_no = self.var_id[i].get()  # simplify variable names
                name = self.var_name[i].get()
                hours = self.var_hours[i].get().strip()
                rate = self.var_rate[i].get().strip()
                amount = self.var_amount[i].get().strip()
                if hours and amount:
                    messagebox.showerror("Data Input Error",
                                         "Input error for {} in row {}. You can not enter both hours and "
                                         "amount. You can only enter one or another, but not both. "
                                         "Awards can be in the form of "
                                         "hours at a given rate OR an amount.".format(name, str(i + 1)),
                                         parent=self.win.topframe)
                    pb_label.destroy()  # destroy the label for the progress bar
                    pb.destroy()  # destroy the progress bar
                    return
                if rate and amount:
                    messagebox.showerror("Data Input Error",
                                         "Input error for {} in row {}. You can not enter both rate and "
                                         "amount. You can only enter one or another, but not both. "
                                         "Awards can be in the form of "
                                         "hours at a given rate OR an amount.".format(name, str(i + 1)),
                                         parent=self.win.topframe)
                    pb_label.destroy()  # destroy the label for the progress bar
                    pb.destroy()  # destroy the progress bar
                    return
                if hours and not rate:
                    messagebox.showerror("Data Input Error",
                                         "Input error for {} in row {}. Hours must be a accompanied by a "
                                         "rate.".format(name, str(i + 1)),
                                         parent=self.win.topframe)
                    pb_label.destroy()  # destroy the label for the progress bar
                    pb.destroy()  # destroy the progress bar
                    return
                if rate and not hours:
                    messagebox.showerror("Data Input Error",
                                         "Input error for {} in row {}. Rate must be a accompanied by a "
                                         "hours.".format(name, str(i + 1)),
                                         parent=self.win.topframe)
                    pb_label.destroy()  # destroy the label for the progress bar
                    pb.destroy()  # destroy the progress bar
                    return
                if hours and not isfloat(hours):
                    messagebox.showerror("Data Input Error",
                                         "Input error for {} in row {}. Hours must be a number."
                                         .format(name, str(i + 1)),
                                         parent=self.win.topframe)
                    pb_label.destroy()  # destroy the label for the progress bar
                    pb.destroy()  # destroy the progress bar
                    return
                if hours and '.' in hours:
                    s_hrs = hours.split(".")
                    if len(s_hrs[1]) > 2:
                        messagebox.showerror("Data Input Error",
                                             "Input error for {} in row {}. Hours must have no "
                                             "more than 2 decimal places.".format(name, str(i + 1)),
                                             parent=self.win.topframe)
                        pb_label.destroy()  # destroy the label for the progress bar
                        pb.destroy()  # destroy the progress bar
                        return
                if rate and amount:
                    messagebox.showerror("Data Input Error",
                                         "Input error for {} in row {}. You can not enter both rate and "
                                         "amount. You can only enter one or another, but not both. "
                                         "Awards can be in the form of "
                                         "hours at a given rate OR an amount.".format(name, str(i + 1)),
                                         parent=self.win.topframe)
                    pb_label.destroy()  # destroy the label for the progress bar
                    pb.destroy()  # destroy the progress bar
                    return
                if rate and amount:
                    messagebox.showerror("Data Input Error",
                                         "Input error for {} in row {}. You can not enter both rate and "
                                         "amount. You can only enter one or another, but not both. "
                                         "Awards can be in the form of "
                                         "hours at a given rate OR an amount.".format(name, str(i + 1)),
                                         parent=self.win.topframe)
                    pb_label.destroy()  # destroy the label for the progress bar
                    pb.destroy()  # destroy the progress bar
                    return
                if rate and not isfloat(rate):
                    messagebox.showerror("Data Input Error", "Input error for {} in row {}. Rates must be a number."
                                         .format(name, str(i + 1)),
                                         parent=self.win.topframe)
                    pb_label.destroy()  # destroy the label for the progress bar
                    pb.destroy()  # destroy the progress bar
                    return
                if rate and '.' in rate:
                    s_rate = rate.split(".")
                    if len(s_rate[1]) > 2:
                        messagebox.showerror("Data Input Error",
                                             "Input error for {} in row {}. Rates must have no "
                                             "more than 2 decimal places.".format(name, str(i + 1)),
                                             parent=self.win.topframe)
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
                                         parent=self.win.topframe)
                    pb_label.destroy()  # destroy the label for the progress bar
                    pb.destroy()  # destroy the progress bar
                    return
                if amount and not isfloat(amount):
                    messagebox.showerror("Data Input Error",
                                         "Input error for {} in row {}. Amounts can only be expressed as "
                                         "numbers. No special characters, such as $ are allowed."
                                         .format(name, str(i + 1)),
                                         parent=self.win.topframe)
                    pb_label.destroy()  # destroy the label for the progress bar
                    pb.destroy()  # destroy the progress bar
                    return
                if amount and '.' in amount:
                    s_amt = amount.split(".")
                    if len(s_amt[1]) > 2:
                        messagebox.showerror("Data Input Error",
                                             "Input error for {} in row {}. "
                                             "Amounts must have no more than 2 decimal places."
                                             .format(name, str(i + 1)),
                                             parent=self.win.topframe)
                        pb_label.destroy()  # destroy the label for the progress bar
                        pb.destroy()  # destroy the progress bar
                        return
                sql = "UPDATE informalc_awards SET hours='%s',rate='%s',amount='%s' WHERE rowid='%s'" % (
                    hours, rate, amount, id_no)
                commit(sql)
                self.win.buttons.update()  # update the progress bar
                ii += 1
            pb.stop()  # stop and destroy the progress bar
            pb_label.destroy()  # destroy the label for the progress bar
            pb.destroy()
            self.addaward2(grv_no)

        def call_grvlist_result(self):
            """ exit out fo the Add Award screen. Destroy the companion window if it still exist. """
            try:
                self.companion_root.destroy()
            except TclError:
                pass
            self.grvlist_result()

        def deletename(self, grv_no, ids):
            """ deletes records from informal c awards. """
            sql = "DELETE FROM informalc_awards WHERE rowid='%s'" % ids
            commit(sql)
            self.addaward2(grv_no)

        def grvlist_setsum(self):
            """ generates text report for settlement list summary showing all grievance settlements. """
            if len(self.search_result) > 0:
                stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                filename = "infc_grv_list" + "_" + stamp + ".txt"
                report = open(dir_path('infc_grv') + filename, "w")
                report.write("   Settlement List Summary\n")
                report.write("   (ordered by date signed)\n\n")
                report.write('  {:<18}{:<12}{:>9}{:>11}{:>12}{:>12}{:>12}\n'
                             .format("    Grievance #", "Date Signed", "GATS #", "Docs?", "Level", "Hours", "Dollars"))
                report.write(
                    "      ----------------------------------------------------------------------------------\n")
                total_hour = 0
                total_amt = 0
                i = 1
                for sett in self.search_result:
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
                            awardxhour += hour * rate
                        if amt:
                            awardxamt += amt
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
                                                 "{0:.2f}".format(float(awardxhour)),
                                                 "{0:.2f}".format(float(awardxamt))))
                        if gi != 0:
                            report.write('{:<34}{:<12}\n'.format("", s_gats[gi]))
                    if i % 3 == 0:
                        report.write(
                            "      ----------------------------------------------------------------------"
                            "------------\n")
                    i += 1
                report.write(
                    "      ----------------------------------------------------------------------------------\n")
                report.write("{:<20}{:>58}\n".format("      Total Hours", "{0:.2f}".format(total_hour)))
                report.write("{:<20}{:>70}\n".format("      Total Dollars", "{0:.2f}".format(total_amt)))
                report.close()
                if sys.platform == "win32":
                    os.startfile(dir_path('infc_grv') + filename)
                if sys.platform == "linux":
                    subprocess.call(["xdg-open", 'kb_sub/infc_grv/' + filename])
                if sys.platform == "darwin":
                    subprocess.call(["open", dir_path('infc_grv') + filename])

        def rptbygrv(self, grv_info):
            """ generates a text report for a specific grievance number. """
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
                    awardxhour += hour * rate
                if amt:
                    awardxamt += amt
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
                messagebox.showerror("Report Generator", "The report was not generated.", parent=self.win.topframe)

        def bycarrier(self):
            """ builds a screen that allows a user to select a carrier and generate a text report of settlements. """
            unique_carrier = self.uniquecarrier()
            frame = self.win  # preserve the window object which is destroyed in the next line
            self.win = MakeWindow()
            self.win.create(frame.topframe)
            Label(self.win.body, text="Informal C: Select Carrier", font=macadj("bold", "Helvetica 18"))\
                .pack(anchor="w")
            Label(self.win.body, text="").pack()
            scrollbar = Scrollbar(self.win.body, orient=VERTICAL)
            listbox = Listbox(self.win.body, selectmode="single", yscrollcommand=scrollbar.set)
            listbox.config(height=30, width=50)
            for name in unique_carrier:
                listbox.insert(END, name)
            scrollbar.config(command=listbox.yview)
            scrollbar.pack(side=RIGHT, fill=Y)
            listbox.pack(side=LEFT, expand=1)
            Button(self.win.buttons, text="Go Back", width=20,
                   command=lambda: self.grvlist_result()).pack(side=LEFT)
            Button(self.win.buttons, text="Report", width=20,
                   command=lambda: self.bycarrier_apply
                   (unique_carrier, listbox.curselection())).pack(side=LEFT)
            self.win.finish()

        def bycarrier_apply(self, names, cursor):
            """ generates a text report for a specified carrier. """
            if len(cursor) == 0:
                return
            unique_grv = []  # get a list of all grv numbers in search range
            for grv in self.search_result:
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
                    total_adj += float(r[1]) * float(r[2])
                else:
                    adj = "---"
                if r[3]:
                    amt = "{0:.2f}".format(float(r[3]))
                    total_amt += float(r[3])
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
                messagebox.showerror("Report Generator", "The report was not generated.", parent=self.win.topframe)

        def uniquecarrier(self):
            """ gets the awards for a carrier from the informalc awards table. """
            unique_grv = []
            for grv in self.search_result:
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

        def bycarriers(self):
            """ generates a text report for settlements by carriers. """
            unique_carrier = self.uniquecarrier()
            unique_grv = []  # get a list of all grv numbers in search range
            for grv in self.search_result:
                if grv[0] not in unique_grv:
                    unique_grv.append(grv[0])  # put these in "unique_grv"
            stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = "infc_grv_list" + "_" + stamp + ".txt"
            report = open(dir_path('infc_grv') + filename, "w")
            report.write("Settlement Report By Carriers\n\n")
            for name in unique_carrier:
                report.write("{:<30}\n\n".format(name))
                report.write(
                    "        Grievance Number    Hours    Rate    Adjusted      Amount       docs       level\n")
                report.write(
                    "    ------------------------------------------------------------------------------------\n")
                results = []
                for ug in unique_grv:  # do search for each grievance in list of unique grievances
                    sql = "SELECT informalc_awards.grv_no, informalc_awards.hours, informalc_awards.rate, " \
                          "informalc_awards.amount, informalc_grv.docs, informalc_grv.level " \
                          "FROM informalc_awards, informalc_grv " \
                          "WHERE informalc_awards.grv_no = informalc_grv.grv_no and " \
                          "informalc_awards.carrier_name='%s'" \
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
                        total_adj += float(r[1]) * float(r[2])
                    else:
                        adj = "---"
                    if r[3]:
                        amt = "{0:.2f}".format(float(r[3]))
                        total_amt += float(r[3])
                    else:
                        amt = "---"
                    if r[5] is None or r[5] == "unknown":
                        r[5] = "---"
                    report.write("    {:<4}{:<17}{:>8}{:>8}{:>12}{:>12}{:>11}{:>12}\n"
                                 .format(str(i), r[0], hours, rate, adj, amt, r[4], r[5]))
                    i += 1
                report.write(
                    "    ------------------------------------------------------------------------------------\n")
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
                messagebox.showerror("Report Generator", "The report was not generated.", parent=self.win.topframe)

        def rptgrvsum(self):
            """ generates a text report for grievance summary. """
            if len(self.search_result) > 0:
                result = list(self.search_result)
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
                            awardxhour += hour * rate
                        if amt:
                            awardxamt += amt
                    space = " "
                    space += num_space * " "
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
                    messagebox.showerror("Report Generator", "The report was not generated.", parent=self.win.topframe)

        def rptcarrierandid(self):
            """ generates a text report with only carrier name and employee id number. """
            if len(self.search_result) == 0:
                return
            stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = "infc_grv_list" + "_" + stamp + ".txt"
            report = open(dir_path('infc_grv') + filename, "w")
            report.write("Carrier List\n\n")
            carriers = self.uniquecarrier()  # get a list of carrier names
            i = 1
            for carrier in carriers:
                emp_id = ""
                sql = "SELECT emp_id FROM name_index WHERE kb_name = '%s'" % carrier
                result = inquire(sql)
                if result:
                    emp_id = result[0][0]
                report.write("{:>4} {:<25}{:>8}\n".format(str(i), carrier, emp_id))
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
                messagebox.showerror("Report Generator", "The report was not generated.", parent=self.win.topframe)

        class RptCarrierId:
            """
            Generate a spread sheet with the carrier's name and employee id for all carriers in the search criteria.
            """

            def __init__(self, parent):
                self.parent = parent
                self.wb = None  # workbook object
                self.carrierlist = None  # workbook name
                self.ws_header = None  # style
                self.input_name = None  # style
                self.input_s = None  # style
                self.col_header = None  # style
                self.i = 0  # this counts the rows/ number of carriers.
                self.no_empid = []  # an array for carriers with no employee id

            def run(self):
                """ this method is the master method for running all other methods in proper order """
                self.get_styles()
                self.build_workbook()
                self.set_dimensions()
                self.build_header()
                self.fill_body()
                self.show_noempid()
                self.save_open()

            def get_styles(self):
                """ Named styles for workbook """
                bd = Side(style='thin', color="80808080")  # defines borders
                self.ws_header = NamedStyle(name="ws_header", font=Font(bold=True, name='Arial', size=12))
                self.input_name = NamedStyle(name="input_name", font=Font(name='Arial', size=8),
                                             border=Border(left=bd, top=bd, right=bd, bottom=bd))
                self.input_s = NamedStyle(name="input_s", font=Font(name='Arial', size=8),
                                          border=Border(left=bd, top=bd, right=bd, bottom=bd),
                                          alignment=Alignment(horizontal='right'))
                self.col_header = NamedStyle(name="col_header", font=Font(bold=True, name='Arial', size=8),
                                             alignment=Alignment(horizontal='left'))

            def build_workbook(self):
                """ creates the workbook object """
                self.wb = Workbook()  # define the workbook
                self.carrierlist = self.wb.active  # create first worksheet
                self.carrierlist.title = "carrier list"  # title first worksheet
                self.carrierlist.oddFooter.center.text = "&A"

            def set_dimensions(self):
                """ adjust the height and width on the violations/ instructions page """
                self.carrierlist.column_dimensions["A"].width = 5
                self.carrierlist.column_dimensions["B"].width = 20
                self.carrierlist.column_dimensions["C"].width = 10

            def build_header(self):
                """ build the header of the spreadsheet """
                self.carrierlist.merge_cells('A1:R1')
                self.carrierlist['A1'] = "Carrier List with Employee ID Numbers"
                self.carrierlist['A1'].style = self.ws_header
                cell = self.carrierlist.cell(row=3, column=2)
                cell.value = "Carrier Name"
                cell.style = self.col_header
                cell = self.carrierlist.cell(row=3, column=3)
                cell.value = "Employee ID"
                cell.style = self.col_header

            def fill_body(self):
                """ this loop will fill the body of the spreadsheet with the carrier list """
                carriers = self.parent.uniquecarrier()  # get a list of carrier names
                self.i = 1
                for carrier in carriers:
                    sql = "SELECT emp_id FROM name_index WHERE kb_name = '%s'" % carrier
                    result = inquire(sql)
                    if result:
                        emp_id = result[0][0]
                        cell = self.carrierlist.cell(row=self.i + 3, column=1)
                        cell.value = str(self.i)
                        cell.style = self.input_name
                        cell = self.carrierlist.cell(row=self.i + 3, column=2)
                        cell.value = carrier
                        cell.style = self.input_name
                        cell = self.carrierlist.cell(row=self.i + 3, column=3)
                        cell.value = emp_id
                        cell.style = self.input_s
                        self.i += 1
                    else:
                        self.no_empid.append(carrier)

            def show_noempid(self):
                """ this will display the a list of carriers with no employee id. """
                if len(self.no_empid) == 0:
                    return
                self.i += 4
                cell = self.carrierlist.cell(row=self.i, column=2)
                cell.value = "Carriers without Employee ID"
                cell.style = self.col_header
                i = 1
                self.i += 1
                for carrier in self.no_empid:
                    cell = self.carrierlist.cell(row=self.i, column=1)
                    cell.value = str(i)
                    cell.style = self.input_name
                    cell = self.carrierlist.cell(row=self.i, column=2)
                    cell.value = carrier
                    cell.style = self.input_name
                    self.i += 1
                    i += 1

            def save_open(self):
                """ save the spreadsheet and open """
                stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                xl_filename = "infc_grv_list" + "_" + stamp + ".xlsx"
                try:
                    self.wb.save(dir_path('infc_grv') + xl_filename)
                    messagebox.showinfo("Spreadsheet generator",
                                        "Your spreadsheet was successfully generated. \n"
                                        "File is named: {}".format(xl_filename),
                                        parent=self.parent.win.topframe)
                    if sys.platform == "win32":  # open the text document
                        os.startfile(dir_path('infc_grv') + xl_filename)
                    if sys.platform == "linux":
                        subprocess.call(["xdg-open", 'kb_sub/infc_grv/' + xl_filename])
                    if sys.platform == "darwin":
                        subprocess.call(["open", dir_path('infc_grv') + xl_filename])
                except PermissionError:
                    messagebox.showerror("Spreadsheet generator",
                                         "The spreadsheet was not generated. \n"
                                         "Suggestion: "
                                         "Make sure that identically named spreadsheets are closed "
                                         "(the file can't be overwritten while open).",
                                         parent=self.parent.win.topframe)

    class PayoutEntry:
        """
        this class allows users to enter carrier payout 
        """
        def __init__(self, parent):
            self.parent = parent
            self.win = None
            self.add_pay_periods = None
            self.add_hours = None
            self.add_rate = None
            self.add_amount = None
            self.poe_listbox_ = None  # holds the list box object
            
        def poe_search(self, frame):
            """ creates a screen that allows user to update payouts for carriers. """
            self.win = MakeWindow()
            self.win.create(frame)
            the_year = StringVar(self.win.topframe)
            the_station = StringVar(self.win.topframe)
            station_options = projvar.list_of_stations
            if "out of station" in station_options:
                station_options.remove("out of station")
            the_station.set("undefined")
            backdate = StringVar(self.win.topframe)
            backdate.set("1")
            Label(self.win.body, text="Informal C: Payout Entry Criteria", font=macadj("bold", "Helvetica 18")) \
                .grid(row=0, column=0, sticky="w", columnspan=4)
            Label(self.win.body, text="").grid(row=1)
            Label(self.win.body, text="Enter the year and the station to be updated.") \
                .grid(row=2, column=0, columnspan=4, sticky="w")
            Label(self.win.body, text="\t\t\tYear: ").grid(row=3, column=1, sticky="e")
            Entry(self.win.body, textvariable=the_year, width=12).grid(row=3, column=2, sticky="w")
            Label(self.win.body, text="Station").grid(row=4, column=1, sticky="e")
            om_station = OptionMenu(self.win.body, the_station, *station_options)
            om_station.config(width=28)
            om_station.grid(row=4, column=2, columnspan=2)
            Label(self.win.body, text="Build the carrier list by going back how many years?") \
                .grid(row=5, column=0, columnspan=4, sticky="w")
            Label(self.win.body, text="Back Date: ").grid(row=6, column=1, sticky="w")
            om_backdate = OptionMenu(self.win.body, backdate, "1", "2", "3", "4", "5", "6", "7", "8", "9", "10")
            om_backdate.config(width=5)
            om_backdate.grid(row=6, column=2, sticky="w")
            button_alignment = macadj("w", "center")
            Button(self.win.buttons, text="Go Back", width=20, anchor=button_alignment,
                   command=lambda: self.parent.informalc(self.win.topframe))\
                .grid(row=0, column=1, sticky="w")
            Button(self.win.buttons, text="Apply", width=20, anchor=button_alignment,
                   command=lambda: self.apply_search(self.win.topframe, the_year,
                                                     the_station, backdate)) \
                .grid(row=0, column=2, sticky="w")
            self.win.finish()

        def apply_search(self, frame, year, station, backdate):
            """ pay out entry - search the args for acceptable values. """
            if year.get().strip() == "":
                messagebox.showerror("Data Entry Error",
                                     "You must enter a year.",
                                     parent=self.win.topframe)
                return
            if "." in year.get():
                messagebox.showerror("Data Entry Error",
                                     "The year can not contain decimal points.",
                                     parent=self.win.topframe)
                return
            if not year.get().isnumeric():
                messagebox.showerror("Data Entry Error",
                                     "The year must numeric without any letters or special characters.",
                                     parent=self.win.topframe)
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
            self.poe_listbox(dt_year, station, dt_start, year)
            self.add(frame, array, selection, year, msg)

        def apply_add(self, frame, name, year, buttons):
            """ payout entry - apply changes """
            if name == "none":
                messagebox.showerror("Data Entry Error",
                                     "You must select a name.",
                                     parent=self.win.topframe)
                return
            for i in range(len(self.add_pay_periods)):
                pp = self.add_pay_periods[i].get().strip()
                hr = self.add_hours[i].get().strip()
                rt = self.add_rate[i].get().strip()
                amt = self.add_amount[i].get().strip()
                if pp and not isint(pp):
                    messagebox.showerror("Data Input Error",
                                         "Input error for {} in row {}. The pay period must be a number"
                                         .format(name, str(i + 1)),
                                         parent=self.win.topframe)
                    return
                if pp and int(pp) > 27:
                    messagebox.showerror("Data Input Error",
                                         "Input error for {} in row {}. The pay period can not be greater "
                                         "than 27".format(name, str(i + 1)),
                                         parent=self.win.topframe)
                    return
                if hr and amt:
                    messagebox.showerror("Data Input Error",
                                         "Input error for {} in row {}. You can not enter both hours and "
                                         "amount. You can only enter one or another, but not both. "
                                         "Awards can be in the form of "
                                         "hours at a given rate OR an amount.".format(name, str(i + 1)),
                                         parent=self.win.topframe)
                    return
                if rt and amt:
                    messagebox.showerror("Data Input Error",
                                         "Input error for {} in row {}. You can not enter both rate and "
                                         "amount. You can only enter one or another, but not both. "
                                         "Awards can be in the form of "
                                         "hours at a given rate OR an amount.".format(name, str(i + 1)),
                                         parent=self.win.topframe)
                    return
                if hr and not rt:
                    messagebox.showerror("Data Input Error",
                                         "Input error for {} in row {}. Hours must be a accompanied by a "
                                         "rate.".format(name, str(i + 1)),
                                         parent=self.win.topframe)
                    return
                if rt and not hr:
                    messagebox.showerror("Data Input Error",
                                         "Input error for {} in row {}. Rate must be a accompanied by a "
                                         "hours.".format(name, str(i + 1)),
                                         parent=self.win.topframe)
                    return
                if hr and not isfloat(hr):
                    messagebox.showerror("Data Input Error", "Input error for {} in row {}. Hours must be a number."
                                         .format(name, str(i + 1)),
                                         parent=self.win.topframe)
                    return
                if hr and '.' in hr:
                    s_hrs = hr.split(".")
                    if len(s_hrs[1]) > 2:
                        messagebox.showerror("Data Input Error",
                                             "Input error for {} in row {}. Hours must have no "
                                             "more than 2 decimal places.".format(name, str(i + 1)),
                                             parent=self.win.topframe)
                        return
                if rt and amt:
                    messagebox.showerror("Data Input Error",
                                         "Input error for {} in row {}. You can not enter both rate and "
                                         "amount. You can only enter one or the other, but not both. "
                                         "Awards can be in the form of "
                                         "hours at a given rate OR an amount.".format(name, str(i + 1)),
                                         parent=self.win.topframe)
                    return
                if rt and not isfloat(rt):
                    messagebox.showerror("Data Input Error",
                                         "Input error for {} in row {}. Rate must be a number."
                                         .format(name, str(i + 1)),
                                         parent=self.win.topframe)
                    return
                if rt and '.' in rt:
                    s_rate = rt.split(".")
                    if len(s_rate[1]) > 2:
                        messagebox.showerror("Data Input Error",
                                             "Input error for {} in row {}. Rates must have no "
                                             "more than 2 decimal places.".format(name, str(i + 1)),
                                             parent=self.win.topframe)
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
                                         parent=self.win.topframe)
                    return
                if amt and not isfloat(amt):
                    messagebox.showerror("Data Input Error",
                                         "Input error for {} in row {}. Amounts can only be expressed as "
                                         "numbers. No special characters, such as $ are allowed."
                                         .format(name, str(i + 1)),
                                         parent=self.win.topframe)
                    return
                if amt and '.' in amt:
                    s_amt = amt.split(".")
                    if len(s_amt[1]) > 2:
                        messagebox.showerror("Data Input Error",
                                             "Input error for {} in row {}. Amounts must have no "
                                             "more than 2 decimal places.".format(name, str(i + 1)),
                                             parent=self.win.topframe)
                        return
            pb_label = Label(buttons, text="Updating Changes: ")  # make label for progress bar
            pb_label.grid(row=1, column=2)
            pb = ttk.Progressbar(buttons, length=200, mode="determinate")  # create progress bar
            pb.grid(row=1, column=3)
            pb["maximum"] = len(self.add_pay_periods) * 2  # set length of progress bar
            pb.start()
            sql = "DELETE FROM informalc_payouts WHERE year='%s' and carrier_name='%s'" % (year, name)
            pb["value"] = len(self.add_pay_periods)  # increment progress bar
            buttons.update()
            commit(sql)
            ii = len(self.add_pay_periods)
            count = 0
            paydays = []
            for i in range(len(self.add_pay_periods)):
                if self.add_pay_periods[i].get().strip() != "":
                    if self.add_hours[i].get().strip() != "" and self.add_rate[i].get().strip() != "" \
                            or self.add_amount[i].get().strip() != "":
                        pp = self.add_pay_periods[i].get().zfill(2)
                        one = "1"
                        pp += one  # format pp so it can fit in find_pp()
                        dt = find_pp(int(year),
                                     pp)  # returns the starting date of the pp when given year and pay period
                        dt += timedelta(days=20)
                        paydays.append(dt)
                        sql = "INSERT INTO informalc_payouts (year,pp,payday,carrier_name,hours,rate,amount) " \
                              "VALUES('%s','%s','%s','%s','%s','%s','%s')" \
                              % (year, self.add_pay_periods[i].get().strip(), paydays[i], name,
                                 self.add_hours[i].get().strip(), self.add_rate[i].get().strip(),
                                 self.add_amount[i].get().strip())
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
            self.add(frame, array, selection, year, msg)

        def add_plus(self, frame, payouts):
            """ pay out entry """
            if len(payouts) == 0:
                self.add_pay_periods.append(StringVar(frame))  # set up array of stringvars for hours,rate,amount
                self.add_hours.append(StringVar(frame))
                self.add_rate.append(StringVar(frame))
                self.add_amount.append(StringVar(frame))
                Entry(frame, textvariable=self.add_pay_periods[len(self.add_pay_periods) - 1], width=10) \
                    .grid(row=len(self.add_pay_periods) + 6, column=0, pady=5, padx=5, sticky="w")
                Entry(frame, textvariable=self.add_hours[len(self.add_hours) - 1], width=10) \
                    .grid(row=len(self.add_hours) + 6, column=1, pady=5, padx=5)
                Entry(frame, textvariable=self.add_rate[len(self.add_rate) - 1], width=10) \
                    .grid(row=len(self.add_rate) + 6, column=2, pady=5, padx=5)
                Entry(frame, textvariable=self.add_amount[len(self.add_amount) - 1], width=10) \
                    .grid(row=len(self.add_amount) + 6, column=3, pady=5, padx=5)
            else:
                for i in range(len(payouts)):
                    self.add_pay_periods.append(StringVar(frame))  # set up array of stringvars for hours,rate,amount
                    self.add_hours.append(StringVar(frame))
                    self.add_rate.append(StringVar(frame))
                    self.add_amount.append(StringVar(frame))
                    self.add_pay_periods[i].set(payouts[i][1])
                    self.add_hours[i].set(payouts[i][4])
                    self.add_rate[i].set(payouts[i][5])
                    self.add_amount[i].set(payouts[i][6])
                    Entry(frame, textvariable=self.add_pay_periods[i], width=10) \
                        .grid(row=len(self.add_pay_periods) + 6, column=0, sticky="w")
                    Entry(frame, textvariable=self.add_hours[i], width=10) \
                        .grid(row=len(self.add_hours) + 6, column=1, pady=5, padx=5)
                    Entry(frame, textvariable=self.add_rate[i], width=10) \
                        .grid(row=len(self.add_rate) + 6, column=2, pady=5, padx=5)
                    Entry(frame, textvariable=self.add_amount[i], width=10) \
                        .grid(row=len(self.add_amount) + 6, column=3, pady=5, padx=5)

        def add(self, frame, array, selection, year, msg):
            """ pay out entry - add payout. """
            empty_array = []
            self.add_pay_periods = []
            self.add_hours = []
            self.add_rate = []
            self.add_amount = []
            self.win = MakeWindow()
            self.win.create(frame)
            Label(self.win.body, text="Informal C: Payout Entry", font=macadj("bold", "Helvetica 18")) \
                .grid(row=0, column=0, sticky="w", columnspan=5)
            Label(self.win.body, text="").grid(row=1)
            if selection != "none":
                Label(self.win.body, text=array[int(selection[0])], font="bold")\
                    .grid(row=2, column=0, sticky="w", columnspan=5)
                name = array[int(selection[0])]
                Label(self.win.body, text="Year: {}".format(year)).grid(row=3, column=0, sticky="w")
                Label(self.win.body, text="").grid(row=4)
                Label(self.win.body, text="PP", width=10, fg="grey").grid(row=5, column=0, sticky="w")
                Label(self.win.body, text="Hours", width=10, fg="grey").grid(row=5, column=1, sticky="w")
                Label(self.win.body, text="Rate", width=10, fg="grey").grid(row=5, column=2, sticky="w")
                Label(self.win.body, text="Amount", width=10, fg="grey").grid(row=5, column=3, sticky="w")
                Button(self.win.body, text="Add Payouts", width=10,
                       command=lambda: self.add_plus(self.win.body, empty_array))\
                    .grid(row=5, column=4, sticky="w")
                sql = "SELECT * FROM informalc_payouts WHERE year ='%s' and carrier_name='%s'ORDER BY pp" \
                      % (year, name)
                payouts = inquire(sql)
                self.add_plus(self.win.body, payouts)
            else:
                Label(self.win.body, text="Select a carrier from the carrier list.").grid(row=2, column=0, sticky="w",
                                                                                          columnspan=5)
                name = "none"
            if msg != "":  # display a message when there is a message
                Label(self.win.buttons, text=msg, fg="red", width=60, anchor="w")\
                    .grid(row=0, column=0, columnspan=4, sticky="w")
            Button(self.win.buttons, text="Go Back", width=20,
                   command=lambda: self.goback(self.win.topframe)) \
                .grid(row=1, column=0, sticky="w")
            Button(self.win.buttons, text="Apply", width=20,
                   command=lambda: self.apply_add(self.win.topframe, name, year, self.win.buttons)) \
                .grid(row=1, column=1, sticky="w")
            Label(self.win.buttons, text="", width=10).grid(row=1, column=2)
            Label(self.win.buttons, text="", width=10).grid(row=1, column=3)
            self.win.finish()

        def goback(self, frame):
            """ pay out entry - go back and destroy the companion window if it still exist. """
            try:
                self.poe_listbox_.destroy()
            except TclError:
                pass
            self.poe_search(frame)

        def poe_listbox(self, dt_year, station, dt_start, year):
            """ pay out entry - create a listbox which allows the user to add carriers. """
            poe_root = Tk()
            self.poe_listbox_ = poe_root  # set the value
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
                   command=lambda: self.add(self.win.topframe, c_list, listbox.curselection(), year, msg)) \
                .pack(side=LEFT, anchor="w")

            Button(n_buttons, text="Close", width=10,
                   command=lambda: (poe_root.destroy())).pack(side=LEFT, anchor="w")

    class PayoutReport:
        """
        this generates reports of pay outs for carriers using information generated by the Payout Entry class. 
        """
        def __init__(self, parent):
            self.parent = parent
            self.win = None

        def informalc_por(self, frame):
            """ pay out report - allows user to conduct a search. """
            self.win = MakeWindow()
            self.win.create(frame)
            afterdate = StringVar(self.win.topframe)
            beforedate = StringVar(self.win.topframe)
            station = StringVar(self.win.topframe)
            station_options = projvar.list_of_stations
            if "out of station" in station_options:
                station_options.remove("out of station")
            station.set("undefined")
            backdate = StringVar(self.win.topframe)
            backdate.set("1")
            Label(self.win.body, text="Informal C: Payout Report Search Criteria",
                  font=macadj("bold", "Helvetica 18")) \
                .grid(row=0, column=0, columnspan=4, sticky="w")
            Label(self.win.body, text="").grid(row=1)
            Label(self.win.body, text="Enter range of dates and select station")\
                .grid(row=2, column=0, columnspan=4, sticky="w")
            Label(self.win.body, text="\tProvide dates in mm/dd/yyyy format.", fg="grey") \
                .grid(row=3, column=0, columnspan=4, sticky="w")
            Label(self.win.body, text="", width=20).grid(row=4, column=0)
            Label(self.win.body, text="After Date: ").grid(row=4, column=1, sticky="w")
            Entry(self.win.body, textvariable=afterdate, width=16).grid(row=4, column=2, sticky="w")
            Label(self.win.body, text="Before Date: ").grid(row=5, column=1, sticky="w")
            Entry(self.win.body, textvariable=beforedate, width=16).grid(row=5, column=2, sticky="w")
            Label(self.win.body, text="Station: ").grid(row=6, column=1, sticky="w")
            om_station = OptionMenu(self.win.body, station, *station_options)
            om_station.config(width=28)
            om_station.grid(row=6, column=2, columnspan=2)
            Label(self.win.body, text="Build the carrier list by going back how many years?") \
                .grid(row=7, column=0, columnspan=4, sticky="w")
            Label(self.win.body, text="Back Date: ").grid(row=8, column=1, sticky="w")
            om_backdate = OptionMenu(self.win.body, backdate, "1", "2", "3", "4", "5", "6", "7", "8", "9", "10")
            om_backdate.config(width=5)
            om_backdate.grid(row=8, column=2, sticky="w")
            button_alignment = macadj("w", "center")
            Button(self.win.buttons, text="Go Back", width=16, anchor=button_alignment,
                   command=lambda: self.parent.informalc(self.win.topframe)).grid(row=0, column=0)
            Label(self.win.buttons, text="Report: ", width=16).grid(row=0, column=1)
            Button(self.win.buttons, text="All Carriers", width=16, anchor=button_alignment,
                   command=lambda: self.por_all(afterdate, beforedate, station, backdate)) \
                .grid(row=0, column=2)
            self.win.finish()

        def por_all(self, afterdate, beforedate, station, backdate):
            """ pay out report. generates text report for all. """
            if not informalc_date_checker(self.win.topframe, afterdate, "After Date"):
                return
            if not informalc_date_checker(self.win.topframe, beforedate, "Before Date"):
                return
            start = informalc_date_converter(afterdate)
            end = informalc_date_converter(beforedate)
            if start > end:
                messagebox.showerror("Data Entry Error",
                                     "The After Date can not be earlier than the Before Date",
                                     parent=self.win.topframe)
                return
            if station.get() == "undefined":
                messagebox.showerror("Data Entry Error",
                                     "You must select a station. ",
                                     parent=self.win.topframe)
                return
            weeks = int(backdate.get()) * 52
            clist_start = start - timedelta(weeks=weeks)
            carrier_list = informalc_gen_clist(clist_start, end, station.get())

            stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = "infc_grv_list" + "_" + stamp + ".txt"
            report = open(dir_path('infc_grv') + filename, "w")
            report.write("  Payouts Report\n\n")
            report.write(
                "  Range of Dates: " + start.strftime("%b %d, %Y") + " - " + end.strftime("%b %d, %Y") + "\n\n")

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
                            payxadj += hour * rate
                        if amt:
                            payxamt += amt
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
                        report.write(
                            '    {:<5}{:>17}{:>9}{:>7}{:>10}{:>12}\n'.format(pp, payday, hours, rate, adj, amt))
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

 
class MakeWindow:
    """
    creates a window with a scrollbar and a frame for buttons on the bottom
    """
    def __init__(self):
        self.topframe = Frame(root)
        self.s = Scrollbar(self.topframe)
        self.c = Canvas(self.topframe, width=1600)
        self.body = Frame(self.c)
        self.buttons = Canvas(self.topframe)  # button bar

    def create(self, frame):
        """ call this method to build the window. If a frame is passed, it will be destroyed """
        if frame is not None:
            frame.destroy()  # close out the previous frame
        self.topframe.pack(fill=BOTH, side=LEFT)
        self.buttons.pack(fill=BOTH, side=BOTTOM)
        # link up the canvas and scrollbar
        self.s.pack(side=RIGHT, fill=BOTH)
        self.c.pack(side=LEFT, fill=BOTH)
        self.s.configure(command=self.c.yview, orient="vertical")
        self.c.configure(yscrollcommand=self.s.set)
        # link the mousewheel - implementation varies by platform
        if sys.platform == "win32":
            self.c.bind_all('<MouseWheel>',
                            lambda event: self.c.yview_scroll(int(projvar.mousewheel * (event.delta / 120)), "units"))
        elif sys.platform == "darwin":
            self.c.bind_all('<MouseWheel>',
                            lambda event: self.c.yview_scroll(int(projvar.mousewheel * event.delta), "units"))
        elif sys.platform == "linux":
            self.c.bind_all('<Button-4>', lambda event: self.c.yview('scroll', -1, 'units'))
            self.c.bind_all('<Button-5>', lambda event: self.c.yview('scroll', 1, 'units'))
        self.c.create_window((0, 0), window=self.body, anchor=NW)

    def finish(self):
        """ This closes the window created by front_window() """
        root.update()
        self.c.config(scrollregion=self.c.bbox("all"))
        try:
            mainloop()  # the window object will loop if it exist.
        except (KeyboardInterrupt, AttributeError):
            try:  # if the object has already been destroyed
                if root:
                    root.destroy()  # destroy it.
            except TclError:
                pass  # else do no nothing.

    def fill(self, last, count):
        """ fill bottom of screen to for scrolling. """
        for i in range(count):
            Label(self.body, text="").grid(row=last + i)
        Label(self.body, text="kb", fg="lightgrey", anchor="w").grid(row=last + count + 1, sticky="w")


class SpeedSheetGen:
    """ this generates and reads a speedsheet for the informal c grievance tracker """

    def __init__(self, frame, station, selection_range):
        self.frame = frame
        self.selection_range = selection_range
        self.station = station
        self.titles = ""
        self.filename = ""
        self.ws_titles = ["grievances", "settlements", "non compliance", "batch settlements", "remanded"]
        # get sql results from the tables.
        self.grievance_onrecs = []
        self.settlement_onrecs = []
        self.nonc_onrecs = []
        self.batch_onrecs = []
        self.remand_onrecs = []
        self.file_result = []
        self.ws_list = []
        self.wb = Workbook()  # define the workbook
        self.ws = None  # the worksheet of the workbook
        self.ws_header = None  # styles for workbook
        self.list_header = None  # styles for workbook
        self.date_dov = None  # styles for workbook
        self.date_dov_title = None  # styles for workbook
        self.col_header = None  # styles for workbook
        self.input_s = None  # styles for workbook
        self.input_ns = None  # styles for workbook
        self.index_columns = [
            ["settlement", "follow up"],  # non compliance index
            ["main", "sub"],  # batch settlement index
            ["remanded", "follow up"]  # remanded index
        ]

    def run(self):
        pass

    def new(self):
        """ this generates a blank speedsheet for new greivances"""
        self.name_styles()
        self.get_titles()  # generate the title and filename
        self.make_workbook_object()  # make the workbook object
        self.create_ws_headers()
        self.create_grievance_headers()
        self.create_settlement_headers()
        self.create_index_headers()
        self.column_formatting_grievances()  # format sheet column widths, fonts, numbers
        self.column_formatting_settlements()  # format sheet column widths, fonts, numbers
        self.column_formatting_indexes()  # format sheet column widths, fonts, numbers
        self.stopsaveopen()

    def selected(self):
        """ this generates a blank speedsheet for selected range of greivances"""
        self.name_styles()
        self.get_titles()  # generate the title and filename
        self.make_workbook_object()  # make the workbook object
        self.create_ws_headers()
        self.create_grievance_headers()
        self.create_settlement_headers()
        self.create_index_headers()
        self.column_formatting_grievances()  # format sheet column widths, fonts, numbers
        self.column_formatting_settlements()  # format sheet column widths, fonts, numbers
        self.column_formatting_indexes()  # format sheet column widths, fonts, numbers
        self.stopsaveopen()

    def all(self):
        """ this generates a blank speedsheet for all greivances"""
        self.name_styles()
        self.get_titles()  # generate the title and filename
        self.get_onrecs()  # get data from all tables to fill speedsheets
        self.make_workbook_object()  # make the workbook object
        self.create_ws_headers()
        self.create_grievance_headers()
        self.create_settlement_headers()
        self.create_index_headers()
        self.column_formatting_grievances()  # format sheet column widths, fonts, numbers
        self.column_formatting_settlements()  # format sheet column widths, fonts, numbers
        self.column_formatting_indexes()  # format sheet column widths, fonts, numbers
        self.insert_grievance_onrecs()  # fills the grievance speedsheet with data from informalc grievances table
        self.insert_settlement_onrecs()  # fills the settlement speedsheet with data from the informalc settlements
        self.stopsaveopen()

    def name_styles(self):
        """ Named styles for workbook """
        bd = Side(style='thin', color="80808080")  # defines borders
        self.ws_header = NamedStyle(name="ws_header", font=Font(bold=True, name='Arial', size=12))
        self.list_header = NamedStyle(name="list_header", font=Font(bold=True, name='Arial', size=9))
        self.date_dov = NamedStyle(name="date_dov", font=Font(name='Arial', size=9))
        self.date_dov_title = NamedStyle(name="date_dov_title", font=Font(bold=True, name='Arial', size=9),
                                         alignment=Alignment(horizontal='right'))
        self.col_header = NamedStyle(name="col_header", font=Font(bold=True, name='Arial', size=9),
                                     border=Border(left=bd, top=bd, right=bd, bottom=bd),
                                     alignment=Alignment(horizontal='left'))

    def get_titles(self):
        """ generate title and filename. The titles and file names vary depending on the selection
        range - new, selected, or all inclusive. This is passed in the command calling SpeedSheetGen. """
        text = "New"
        filetext = "new"
        if self.selection_range == "selected":
            text = "Selected"
            filetext = "selected"
        if self.selection_range == "all":
            text = "All"
            filetext = "all"
        self.titles = (
            "Speedsheet - {} Grievances".format(text),
            "Speedsheet - {} Settlements ".format(text),
            "Speedsheet - {} Non Compliance Index".format(text),
            "Speedsheet - {} Batch Settlement Index".format(text),
            "Speedsheet - {} Remanded Index".format(text)
        )
        self.filename = "{}_grievances_speedsheet".format(filetext) + ".xlsx"

    def get_onrecs(self):
        """ get data from tables """
        sql = "SELECT * FROM 'informalc_grievances' WHERE station = '%s'" % self.station
        self.grievance_onrecs = inquire(sql)
        grv_list = []  # array to hold all grievance numbers
        for grv in self.grievance_onrecs:
            grv_list.append(grv[2])
        # use arrays and loops to get search results for all the grievances in the grv_list array.
        # search these tables
        tables_array = ("informalc_settlements", "informalc_noncindex", "informalc_batchindex",
                        "informalc_remandindex")
        # search these columns in the tables
        search_criteria_array = ("grv_no", "settlement", "main", "remanded")
        # store the results in these arrays
        results_array = [self.settlement_onrecs, self.nonc_onrecs, self.batch_onrecs,
                         self.remand_onrecs]
        # loop for grievance in each table
        for i in range(len(results_array)):
            for ii in range(len(grv_list)):
                sql = "SELECT * FROM '%s' WHERE %s = '%s'" % (tables_array[i], search_criteria_array[i], grv_list[ii])
                result = inquire(sql)
                # get the onrecs for informalc settlements
                if tables_array[i] == "informalc_settlements":
                    if result:
                        self.settlement_onrecs.append(result[0])
                # get the onrecs for informalc non compliance index
                if tables_array[i] == "informalc_noncindex":
                    if result:
                        self.nonc_onrecs.append(result[0])
                # get the onrecs for informalc_batchindex
                if tables_array[i] == "informalc_batchindex":
                    if result:
                        self.batch_onrecs.append(result[0])
                # get the onrecs for informalc_remandindex
                if tables_array[i] == "informalc_remandindex":
                    if result:
                        self.remand_onrecs.append(result[0])

    def make_workbook_object(self):
        """ make the workbook object """
        self.ws_list = ["grievances", "settlements", "non compliance", "batch settlements", "remanded"]
        self.ws_list[0] = self.wb.active  # create first worksheet - this will be for grievances
        self.ws_list[0].title = self.ws_titles[0]  # title first worksheet - this is for grievances
        for i in range(1, len(self.ws_list)):  # loop to create all other worksheets
            self.ws_list[i] = self.wb.create_sheet(self.ws_titles[i])

    def create_ws_headers(self):
        """ use a loop to create headers for all the worksheets """
        for i in range(5):  # there are five worksheets
            cell = self.ws_list[i].cell(column=1, row=1)
            cell.value = self.titles[i]
            cell.style = self.ws_header
            self.ws_list[i].merge_cells('A1:G1')
            cell = self.ws_list[i].cell(column=1, row=3)
            cell.value = "Station: "
            cell.style = self.date_dov_title
            cell = self.ws_list[i].cell(column=2, row=3)
            cell.value = self.station
            cell.style = self.date_dov
            self.ws_list[i].merge_cells('B3:C3')

    def create_grievance_headers(self):
        """ create the grievance worksheet. all worksheets must be formatted separately since they all have
        distinct information. """
        cell = self.ws_list[0].cell(column=1, row=5)
        cell.value = "grievant"
        cell.style = self.col_header
        cell = self.ws_list[0].cell(column=2, row=5)
        cell.value = "grievance number"
        cell.style = self.col_header
        cell = self.ws_list[0].cell(column=3, row=5)
        cell.value = "start incident"
        cell.style = self.col_header
        cell = self.ws_list[0].cell(column=4, row=5)
        cell.value = "end incident"
        cell.style = self.col_header
        cell = self.ws_list[0].cell(column=5, row=5)
        cell.value = "meeting date"
        cell.style = self.col_header
        cell = self.ws_list[0].cell(column=6, row=5)
        cell.value = "issue"
        cell.style = self.col_header
        cell = self.ws_list[0].cell(column=7, row=5)
        cell.value = "article"
        cell.style = self.col_header

    def create_settlement_headers(self):
        """ create the grievance worksheet. all worksheets must be formatted separately since they all have
        distinct information. """
        cell = self.ws_list[1].cell(column=1, row=5)
        cell.value = "grievance number"
        cell.style = self.col_header
        cell = self.ws_list[1].cell(column=2, row=5)
        cell.value = "level"
        cell.style = self.col_header
        cell = self.ws_list[1].cell(column=3, row=5)
        cell.value = "date signed"
        cell.style = self.col_header
        cell = self.ws_list[1].cell(column=4, row=5)
        cell.value = "decision"
        cell.style = self.col_header
        cell = self.ws_list[1].cell(column=5, row=5)
        cell.value = "proof due"
        cell.style = self.col_header
        cell = self.ws_list[1].cell(column=6, row=5)
        cell.value = "docs"
        cell.style = self.col_header
        cell = self.ws_list[1].cell(column=7, row=5)
        cell.value = "gats number"
        cell.style = self.col_header

    def create_index_headers(self):
        """ use a loop to fill in the index headers using self.index_columns """
        for i in range(3):
            cell = self.ws_list[i+2].cell(column=1, row=5)
            cell.value = self.index_columns[i][0]
            cell.style = self.col_header
            cell = self.ws_list[i+2].cell(column=2, row=5)
            cell.value = self.index_columns[i][1]
            cell.style = self.col_header

    def column_formatting_grievances(self):
        """ format the columns. this can be overridden by individually formating the cells. """
        self.ws_list[0].oddFooter.center.text = "&A"
        col = self.ws_list[0].column_dimensions["A"]
        col.width = 25
        col.font = Font(size=9, name="Arial")
        col = self.ws_list[0].column_dimensions["B"]
        col.width = 20
        col.font = Font(size=9, name="Arial")
        col.number_format = '@'
        col = self.ws_list[0].column_dimensions["C"]
        col.width = 12
        col.font = Font(size=9, name="Arial")
        col.number_format = 'MM/DD/YYYY'
        col = self.ws_list[0].column_dimensions["D"]
        col.width = 12
        col.font = Font(size=9, name="Arial")
        col.number_format = 'MM/DD/YYYY'
        col = self.ws_list[0].column_dimensions["E"]
        col.width = 12
        col.font = Font(size=9, name="Arial")
        col.number_format = 'MM/DD/YYYY'
        col = self.ws_list[0].column_dimensions["F"]
        col.width = 25
        col.font = Font(size=9, name="Arial")
        col = self.ws_list[0].column_dimensions["G"]
        col.width = 6
        col.font = Font(size=9, name="Arial")
        
    def column_formatting_settlements(self):
        """ format the columns. this can be overridden by individually formating the cells. """
        self.ws_list[1].oddFooter.center.text = "&A"
        col = self.ws_list[1].column_dimensions["A"]  # grievance number
        col.width = 18
        col.font = Font(size=9, name="Arial")
        col = self.ws_list[1].column_dimensions["B"]  # level
        col.width = 10
        col.font = Font(size=9, name="Arial")
        col = self.ws_list[1].column_dimensions["C"]  # date signed
        col.width = 10
        col.font = Font(size=9, name="Arial")
        col.number_format = 'MM/DD/YYYY'
        col = self.ws_list[1].column_dimensions["D"]  # decision
        col.width = 20
        col.font = Font(size=9, name="Arial")
        col = self.ws_list[1].column_dimensions["E"]  # proof due
        col.width = 10
        col.font = Font(size=9, name="Arial")
        col.number_format = 'MM/DD/YYYY'
        col = self.ws_list[1].column_dimensions["F"]  # docs
        col.width = 15
        col.font = Font(size=9, name="Arial")
        col = self.ws_list[1].column_dimensions["G"]  # gats_number
        col.width = 12
        col.font = Font(size=9, name="Arial")

    def column_formatting_indexes(self):
        """ format the columns of all index worksheets - non compliance, batch settlements and remanded"""
        for i in range(3):
            self.ws_list[i+2].oddFooter.center.text = "&A"
            col = self.ws_list[i+2].column_dimensions["A"]  # settlement/main/remanded
            col.width = 20
            col.font = Font(size=9, name="Arial")
            col = self.ws_list[i+2].column_dimensions["B"]  # followup/sub/followup
            col.width = 20
            col.font = Font(size=9, name="Arial")

    def insert_grievance_onrecs(self):
        """ loop for each grievance on record to fill the grievance speedsheet which is ws.list[0] """
        row = 6  # start on row 6 to make room for headers
        for grv in self.grievance_onrecs:
            grievant = grv[0]
            grievance_number = grv[2]
            start_incident = Convert(grv[3]).dtstr_to_backslashstr()
            end_incident = Convert(grv[4]).dtstr_to_backslashstr()
            meeting_date = Convert(grv[5]).dtstr_to_backslashstr()
            issue = grv[6]
            article = grv[7]
            values_array = [grievant, grievance_number, start_incident, end_incident, meeting_date, 
                            issue, article]
            for i in range(len(values_array)):
                cell = self.ws_list[0].cell(row=row, column=i+1)  # carrier effective date
                cell.value = values_array[i]
                if i in (2, 3, 4):
                    cell.number_format = 'MM/DD/YYYY'
            row += 1

    def insert_settlement_onrecs(self):
        """ loop for each grievance on record to fill the grievance speedsheet which is ws.list[0] """
        row = 6  # start on row 6 to make room for headers
        for sett in self.settlement_onrecs:  # loop for each row
            grievance_number = sett[0]  # define all the fields
            level = sett[1]
            date_signed = Convert(sett[2]).dtstr_to_backslashstr()
            decision = sett[3]
            proofdue = Convert(sett[4]).dtstr_to_backslashstr()
            docs = sett[5]
            gats_number = sett[6]
            values_array = [grievance_number, level, date_signed, decision,
                            proofdue, docs, gats_number]
            for i in range(len(values_array)):  # loop for each column
                cell = self.ws_list[1].cell(row=row, column=i+1)  # define the cell by sheet and cell coordinates
                cell.value = values_array[i]  # insert the appropriate element
                if i in (2, 4):  # for date signed and proof due, format the cell as a date.
                    cell.number_format = 'MM/DD/YYYY'
                    cell.style = self.date_dov
            row += 1

    def stopsaveopen(self):
        """ save and open the speedsheet. """
        try:
            self.wb.save(dir_path('informalc_speedsheets') + self.filename)
            messagebox.showinfo("Speedsheet Generator",
                                "Your speedsheet was successfully generated. \n"
                                "File is named: {}".format(self.filename),
                                parent=self.frame)
            if sys.platform == "win32":
                os.startfile(dir_path('informalc_speedsheets') + self.filename)
            if sys.platform == "linux":
                subprocess.call(["xdg-open", 'kb_sub/informalc_speedsheets/' + self.filename])
            if sys.platform == "darwin":
                subprocess.call(["open", dir_path('informalc_speedsheets') + self.filename])
        except PermissionError:
            messagebox.showerror("Speedsheet generator",
                                 "The speedsheet was not generated. \n"
                                 "Suggestion: \n"
                                 "Make sure that identically named informalc_speedsheets are closed \n"
                                 "(the file can't be overwritten while open).\n",
                                 parent=self.frame)


class SpeedWorkBookGet:
    """
    this class gets the speedsheet and opens it.
    """

    def __init__(self):
        pass

    @staticmethod
    def get_filepath():
        """ get the file path"""
        if projvar.platform == "macapp" or projvar.platform == "winapp":
            return os.path.join(os.path.sep,
                                os.path.expanduser("~"), 'Documents', 'klusterbox', 'informalc_speedsheets')
        else:
            return 'kb_sub/informalc_speedsheets'

    def get_file(self):
        """ returns the file path if there is one. else no selection/invalid selection. """
        path_ = self.get_filepath()
        file_path = filedialog.askopenfilename(initialdir=path_, filetypes=[("Excel files", "*.xlsx")])
        if file_path[-5:].lower() == ".xlsx":
            return file_path
        elif file_path == "":
            return "no selection"
        else:
            return "invalid selection"

    def open_file(self, frame, interject):
        """ gets the file and calls the speedsheet check and progress bar. """
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


class SpeedLoadThread(Thread):
    """ use multithreading to load workbook while progress bar runs """

    def __init__(self, path_):
        Thread.__init__(self)
        self.path_ = path_
        self.workbook = ""

    def run(self):
        """ runs the speedsheet loading. """
        global pb_flag  # this will signal when the thread has ended to end the progress bar
        wb = load_workbook(self.path_)  # load xlsx doc with openpyxl
        self.workbook = wb
        pb_flag = False


class SpeedSheetCheck:
    """ a class for checking the informal c grievance speedsheets. """
    def __init__(self, frame, wb, path_, interject):
        self.frame = frame
        self.station = None
        self.wb = wb
        self.ws = None  # this hold the worksheet
        self.path_ = path_
        self.interject = interject
        self.input_type = None
        self.sheets = None
        self.sheet_count = None
        self.grievance_count = 0  # count of how many grievances have been checked.
        self.fatal_rpt = 0
        self.add_rpt = 0
        self.fyi_rpt = 0
        self.settlement_count = 0  # count of how many settlements have been checked
        self.settlement_fatal_rpt = 0
        self.settlement_add_rpt = 0
        self.settlement_fyi_rpt = 0
        self.sheet_rowcount = []
        self.row_counter = 0  # get the total amount of rows in the worksheet
        self.start_row = 6  # the row where after the headers
        self.pb = None
        # self.pb = ProgressBarDe(label="SpeedSheet Checking")
        self.pb_counter = 0
        self.filename = ReportName("speedsheet_precheck").create()  # generate a name for the report
        self.report = open(dir_path('report') + self.filename, "w")  # open the report
        self.grv_mentioned = False  # keeps grievance numbers from being repeated in reports
        self.worksheet = ("grievances", "settlements", "non compliance", "batch", "remanded")
        self.index_columns = [
            ["settlement", "follow up"],  # non compliance index
            ["main", "sub"],  # batch settlement index
            ["remanded", "follow up"]  # remanded index
        ]
        self.allowaddrecs = True
        self.fullreport = True
        self.name_mentioned = False
        self.issue_index = []  # get the speedsheet issue index number for issue categories
        self.issue_description = []  # get the issue description for issue categories
        self.issue_article = []  # get the article of the issue for issue catergories
        self.decision_index = []  # get the speedsheet decision index number for decision categories
        self.decision_description = []  # get the decision description for the decision categories

    def check(self):
        """ master method for running other methods and returns to the mainframe. """
        try:
            self.pb = ProgressBarDe(label="SpeedSheet Checking")
            self.get_issuecats()  # fetch the issue categories from the informalc_issuescategories table
            self.get_decisioncats()  # fetch the decision categories from the informalc_decisioncategories table
            self.set_sheet_facts()
            self.set_station()
            self.start_reporter()
            self.checking()
            self.reporter()
            self.pb.stop()
        except KeyError:  # if wrong type of file is selected, there will be an error
            self.pb.delete()  # stop and destroy progress bar
            self.showerror()

    def get_issuecats(self):
        """ fetch the issue categories from the informalc_issuescategories table of the db and place them in arrays. """
        sql = "SELECT * FROM informalc_issuescategories"
        results = inquire(sql)
        for r in results:
            self.issue_index.append(r[0])
            self.issue_description.append(r[2])
            self.issue_article.append(r[1])

    def get_decisioncats(self):
        """ fetch the decision categories from the informalc_decisioncategories table of the db and place them in
         arrays """
        sql = "SELECT * FROM informalc_decisioncategories"
        results = inquire(sql)
        for r in results:
            self.decision_index.append(r[0])
            self.decision_description.append(r[2])

    def set_sheet_facts(self):
        """ get the worksheet names and number worksheets. """
        # there are three input types: new, selected, or all inclusive
        self.input_type = "new"
        self.sheets = self.wb.sheetnames  # get the names of the worksheets as a list
        self.sheet_count = len(self.sheets)  # get the number of worksheets

    def set_station(self):
        """ gets the station from the speedsheet. """
        self.station = self.wb[self.sheets[0]].cell(row=3, column=2).value  # get the station.

    def start_reporter(self):
        """ starts the report. """
        self.report.write("\nSpeedSheet Pre-Check Report \n")
        self.report.write(">>> {}\n".format(self.path_))

    def row_count(self):
        """ get a count of all rows for all sheets - need for progress bar """
        total_rows = 0
        for i in range(self.sheet_count):
            ws = self.wb[self.sheets[i]]  # assign the worksheet object
            row_count = ws.max_row  # get the total amount of rows in the worksheet
            self.sheet_rowcount.append(row_count)
            total_rows += row_count
        return total_rows

    def showerror(self):
        """ message box for showing errors. """
        messagebox.showerror("Klusterbox SpeedSheets",
                             "SpeedSheets Precheck or Input has failed. \n"
                             "Either you have selected a spreadsheet that is not \n"
                             "a SpeedSheet or your Speedsheet is corrupted. \n"
                             "Suggestion: Verify that the file you are selecting \n "
                             "is a SpeedSheet. \n"
                             "Suggestion: Try re-generating the SpeedSheet.",
                             parent=self.frame)

    def checking(self):
        """ reads rows and send to SpeedCarrierCheck or SpeedRingCheck. """
        count_diff = self.sheet_count * (self.start_row - 1)  # subtract top five/six rows from the row count
        self.pb.max_count(self.row_count() - count_diff)  # get total count of rows for the progress bar
        self.pb.start_up()  # start up the progress bar
        self.pb_counter = 0  # initialize the progress bar counter
        for i in range(self.sheet_count):  # loop once for each worksheet in the workbook
            self.ws = self.wb[self.sheets[i]]  # assign the worksheet object
            self.row_counter = self.ws.max_row  # get the total amount of rows in the worksheet
            if self.worksheet[i] == "grievances":  # execute for grievance speedsheet
                self.scan_grievances(i)
            if self.worksheet[i] == "settlements":  # execute for settlements speedsheet
                self.scan_settlements(i)
        # self.pb.stop()

    def scan_grievances(self, i):
        """ scan the values of the grievances worksheet, line by line. """
        # loop through all rows, start with row 5 or 6 until the end
        for ii in range(self.start_row, self.row_counter + 1):
            self.pb.move_count(self.pb_counter)
            self.grv_mentioned = False  # keeps names from being repeated in reports
            self.grievance_count += 1  # get a count of the carriers for reports
            grievant = Handler(self.ws.cell(row=ii, column=1).value).nonetype()
            grv_no = Handler(self.ws.cell(row=ii, column=2).value).nonetype()
            startdate = Handler(self.ws.cell(row=ii, column=3).value).nonetype()
            enddate = Handler(self.ws.cell(row=ii, column=4).value).nonetype()
            meetingdate = Handler(self.ws.cell(row=ii, column=5).value).nonetype()
            issue = Handler(self.ws.cell(row=ii, column=6).value).nonetype()
            article = Handler(self.ws.cell(row=ii, column=7).value).nonetype()
            self.pb.change_text("Reading Speedcell: {}".format(grv_no))  # update text for progress bar
            SpeedGrvCheck(self, self.sheets[i], ii, grievant, grv_no, startdate, enddate, meetingdate,
                          issue, article).check_all()
        self.pb_counter += 1

    def scan_settlements(self, i):
        """ scan the values of the grievances worksheet, line by line. """
        for ii in range(self.start_row,
                        self.row_counter + 1):  # loop through all rows, start with row 5 or 6 until the end
            if self.ws.cell(row=ii, column=1).value is not None:  # if there is a grievance number
                self.pb.move_count(self.pb_counter)
                self.grv_mentioned = False  # keeps names from being repeated in reports
                self.settlement_count += 1  # get a count of the carriers for reports
                grv_no = Handler(self.ws.cell(row=ii, column=1).value).nonetype()
                level = Handler(self.ws.cell(row=ii, column=2).value).nonetype()
                datesigned = Handler(self.ws.cell(row=ii, column=3).value).nonetype()
                decision = Handler(self.ws.cell(row=ii, column=4).value).nonetype()
                proofdue = Handler(self.ws.cell(row=ii, column=5).value).nonetype()
                docs = Handler(self.ws.cell(row=ii, column=6).value).nonetype()
                gatsnumber = Handler(self.ws.cell(row=ii, column=7).value).nonetype()
                self.pb.change_text("Reading Speedcell: {}".format(grv_no))  # update text for progress bar
                SpeedSetCheck(self, self.sheets[i], ii, grv_no, level, datesigned, decision, proofdue, docs,
                              gatsnumber).check_all()
            else:
                self.pb.change_text("Detected empty Speedcell.")  # update text for progress bar
        self.pb_counter += 1

    def scan_indexes(self, i):
        """ scan the values of the grievances worksheet, line by line. """
        # loop through all rows, start with row 5 or 6 until the end
        for ii in range(self.start_row, self.row_counter + 1):
            # if there is a grievance number for both columns
            if self.ws.cell(row=ii, column=1).value and self.ws.cell(row=ii, column=2).value is not None:
                self.pb.move_count(self.pb_counter)
                self.grv_mentioned = False  # keeps names from being repeated in reports
                self.settlement_count += 1  # get a count of the carriers for reports
                self.index_columns[i-2][0] = Handler(self.ws.cell(row=ii, column=1).value).nonetype()
                self.index_columns[i-2][1] = Handler(self.ws.cell(row=ii, column=2).value).nonetype()
                # update text for progress bar
                self.pb.change_text("Reading Speedcell: {}".format(self.index_columns[i-2][0]))
                SpeedIndexCheck(self, self.sheets[i], ii, self.index_columns[i-2][0], self.index_columns[i-2][1])\
                    .check_all()
            else:
                self.pb.change_text("Detected empty Speedcell.")  # update text for progress bar
        self.pb_counter += 1

    def reporter(self):
        """ writes the report """
        self.report.write("\n\n----------------------------------")
        # build report summary for carrier checks
        self.report.write("\n\nGrievance SpeedSheet Check Complete.\n\n")
        msg = "grievance{} checked".format(Handler(self.grievance_count).plurals())
        self.report.write('{:>6}  {:<40}\n'.format(self.grievance_count, msg))
        msg = "fatal error{} found".format(Handler(self.fatal_rpt).plurals())
        self.report.write('{:>6}  {:<40}\n'.format(self.fatal_rpt, msg))
        if self.interject:
            msg = "addition{} made".format(Handler(self.add_rpt).plurals())
            self.report.write('{:>6}  {:<40}\n'.format(self.add_rpt, msg))
        else:
            msg = "fyi notification{}".format(Handler(self.fyi_rpt).plurals())
            self.report.write('{:>6}  {:<40}\n'.format(self.fyi_rpt, msg))
        # build report summary for rings checks
        self.report.write("\n\nSettlements SpeedSheet Check Complete.\n\n")
        msg = "settlement{} checked".format(Handler(self.settlement_count).plurals())
        self.report.write('{:>6}  {:<40}\n'.format(self.settlement_count, msg))
        msg = "fatal error{} found".format(Handler(self.settlement_fatal_rpt).plurals())
        self.report.write('{:>6}  {:<40}\n'.format(self.settlement_fatal_rpt, msg))
        if self.interject:
            msg = "addition{} made".format(Handler(self.settlement_add_rpt).plurals())
            self.report.write('{:>6}  {:<40}\n'.format(self.settlement_add_rpt, msg))
        else:
            msg = "fyi notification{}".format(Handler(self.settlement_fyi_rpt).plurals())
            self.report.write('{:>6}  {:<40}\n'.format(self.settlement_fyi_rpt, msg))
        # close out the report and open in notepad
        self.report.close()
        if sys.platform == "win32":  # open the text document
            os.startfile(dir_path('report') + self.filename)
        if sys.platform == "linux":
            subprocess.call(["xdg-open", 'kb_sub/report/' + self.filename])
        if sys.platform == "darwin":
            subprocess.call(["open", dir_path('report') + self.filename])


class SpeedGrvCheck:
    """ checks one line of the grievance speedsheet when it is called by the SpeedSheetCheck class. """
    def __init__(self, parent, sheet, row, grievant, grv_no, startdate, enddate, meetingdate, issue, article):
        self.parent = parent
        self.sheet = sheet  # input here is coming directly from the speedcell
        self.row = str(row)
        self.grievant = grievant
        self.grv_no = grv_no
        self.startdate = startdate
        self.enddate = enddate
        self.meetingdate = meetingdate
        self.input_date = []  # array to hold startdate, enddate and meetingdate - form in check_dates()
        self.issue = issue
        self.article = article
        # onrec variables - these hold the values of the record currently in the database.
        self.onrec = False  # this value is True if a sql search shows that there is a rec of the grv_no in the db.
        self.onrec_grievant = ""
        # skip station as that is held in self.parent.station
        # skip grievance number as that is self.grv_no
        self.onrec_startdate = ""
        self.onrec_enddate = ""
        self.onrec_meetingdate = ""
        self.onrec_issue = ""
        self.onrec_article = ""
        self.error_array = []  # gives a report of failed checks
        self.attn_array = []  # gives a report of issues to bring to the attention of users
        self.add_array = []  # gives a report of records to add to the database
        self.fyi_array = []  # gives a report of useful information for the user
        self.parent.name_mentioned = False  # reset this so that name is not repeated on reports
        self.addday = []  # checked input formatted for entry into database
        self.addgrievant = "empty"
        self.addstartdate = "empty"
        self.addenddate = "empty"
        self.addmeetingdate = "empty"
        self.adddate = [self.addstartdate, self.addenddate, self.addmeetingdate]
        self.addissue = "empty"
        self.addarticle = "empty"

    def check_all(self):
        """ master method to run other methods. """
        self.reformat_grv_no()  # reformat the grievance number to all lowercase, no whitespaces, no dashes.
        if self.check_grv_number():  # first check the grievance number. if that is good, then proceed.
            self.get_onrecs()  # 'on record' - get the record currently in the database if it exist
            self.check_grievant()
            self.check_dates()
            self.check_issue()
            self.add_recs()  # write changes to the db
        self.generate_report()

    def check_grv_number(self):
        """ check the grievance number input """
        if not GrievanceChecker(self.grv_no).has_value():
            error = "     ERROR: The grievance number must not be blank. \n"
            self.error_array.append(error)
            self.parent.allowaddrecs = False  # do not allow this speedcell be be input into database
            return False
        if not GrievanceChecker(self.grv_no).check_characters():
            error = "     ERROR: The grievance number can only contain numbers and letters. \n"
            self.error_array.append(error)
            self.parent.allowaddrecs = False  # do not allow this speedcell be be input into database
            return False
        if not GrievanceChecker(self.grv_no).min_lenght():
            error = "     ERROR: The grievance number must contain at least 4 characters. \n"
            self.error_array.append(error)
            self.parent.allowaddrecs = False  # do not allow this speedcell be be input into database
            return False
        if not GrievanceChecker(self.grv_no).max_lenght():
            error = "     ERROR: The grievance number can not contain more than 20 characters. \n"
            self.error_array.append(error)
            self.parent.allowaddrecs = False  # do not allow this speedcell be be input into database
            return False
        return True

    def reformat_grv_no(self):
        """ reformat the grievance number to all lowercase, no whitespaces, no dashes. """
        self.grv_no = self.grv_no.lower()  # convert grievance number to lowercas
        self.grv_no = self.grv_no.strip()  # strip whitespace from start and end of the string.
        self.grv_no = self.grv_no.replace('-', '')  # remove any dashes
        self.grv_no = self.grv_no.replace(' ', '')  # remove any whitespace

    def reformat_grievant(self):
        """ reformat the grievant to all lowercase, no whitespaces """
        self.grievant = self.grievant.lower()
        self.grievant = self.grievant.strip()

    def get_onrecs(self):
        """ check if there is an existing record for the grievance number in the informalc grievances table.
        if so, store the values in the self.onrec variables. if not, the self.onrec variables default to empty. """
        sql = "SELECT * FROM informalc_grievances WHERE grv_no = '%s' and station = '%s'" \
              % (self.grv_no, self.parent.station)
        results = inquire(sql)
        if results:
            self.onrec = True  # this value is True if a sql search shows that there is a rec in the db.
            self.onrec_grievant = results[0][0]
            # skip station as that is held in self.parent.station and is part of the search criteria
            # skip grievance number as that is self.grv_no and is part of the search criteria
            self.onrec_startdate = results[0][3]
            self.onrec_enddate = results[0][4]
            self.onrec_meetingdate = results[0][5]
            self.onrec_issue = results[0][6]
            self.onrec_article = results[0][7]

    def check_grievant(self):
        """ check the grievant input. this is either 'class action' or a carrier name. it can be blank. """
        self.reformat_grievant()  # remove external whitespace and convert to lower case
        not_names = ("class action", "")
        if self.grievant in not_names:  # "class action" is a standard entry
            self.add_grievant()
            return
        if not NameChecker(self.grievant).check_characters():
            error = "     ERROR: Grievant name can not contain numbers or most special characters\n"
            self.error_array.append(error)
            self.parent.allowaddrecs = False  # do not allow this speedcell be be input into database
            return
        if not NameChecker(self.grievant).check_length():
            error = "     ERROR: Grievant name must not exceed 42 characters\n"
            self.error_array.append(error)
            self.parent.allowaddrecs = False  # do not allow this speedcell be be input into database
            return
        if not NameChecker(self.grievant).check_comma():
            error = "     ERROR: Grievant name must contain one comma to separate last name and first initial\n"
            self.error_array.append(error)
            self.parent.allowaddrecs = False  # do not allow this speedcell be be input into database
            return
        if not NameChecker(self.grievant).check_initial():
            attn = "     ATTENTION: Grievant name should must contain one initial ideally, \n" \
                   "                unless more are needed to create a distinct carrier name.\n"
            self.attn_array.append(attn)
        self.add_grievant()

    def add_grievant(self):
        """ add the grievant to add_grivant variable """
        if self.grievant == self.onrec_grievant:
            pass  # retain "empty" value for grievant variable
        else:
            fyi = "     FYI: New or updated grievant: {}\n".format(self.grievant)
            self.fyi_array.append(fyi)
            self.addgrievant = self.grievant  # save to input to dbase

    def check_dates(self):
        """ check the startdate, enddate and meetingdate.
         since these are all dates with similiar criteria, use a loop to check them.
         sometimes, openpyxl sends the dates as strings of datetime objects, instead of the mm/dd/yyyy formated dates,
         the DateTimeChecker() will identify these and skip the checks. """
        self.input_date = [self.startdate, self.enddate, self.meetingdate]
        for i in range(3):
            self.check_date_loop(i)

    def check_date_loop(self, i):
        """ loop from check dates """
        _type = ("start", "end", "meeting")
        if self.input_date[i].strip() == "":  # if the value is blank, skip all the checks
            self.add_date(i)
            return
        # if the value is a valid dt object, skip all the checks
        if DateTimeChecker().check_dtstring(self.input_date[i]):
            self.add_date(i)
            return
        date_object = BackSlashDateChecker(self.input_date[i])  # first create the date_object
        if not date_object.count_backslashes():  # this checks that there are 2 backslashes in the date
            error = "     ERROR: The date for the {} date must have two backslashes. Got instead: {}\n"\
                .format(_type[i], self.input_date[i])
            self.error_array.append(error)
            self.parent.allowaddrecs = False  # do not allow this speedcell be be input into database
            return
        date_object.breaker()  # this breaks the object into month, day and year elements.
        if not date_object.check_numeric():  # check each element in the date to ensure they are numeric
            error = "     ERROR: The month, day and year for the {} date must be numeric\n".format(_type[i])
            self.error_array.append(error)
            self.parent.allowaddrecs = False  # do not allow this speedcell be be input into database
            return
        if not date_object.check_minimums():  # check each element in the date to ensure they are greater than zero
            error = "     ERROR: The month, day and year for the {} date must be greater than zero.\n"\
                .format(_type[i])
            self.error_array.append(error)
            self.parent.allowaddrecs = False  # do not allow this speedcell be be input into database
            return
        if not date_object.check_month():  # returns False if the month is greater than 12.
            error = "     ERROR: The month for the {} date must less than 13.\n".format(_type[i])
            self.error_array.append(error)
            self.parent.allowaddrecs = False  # do not allow this speedcell be be input into database
            return
        if not date_object.check_day():  # return False if the day is greater than 31.
            error = "     ERROR: The day entered for the {} date is must be less than 32.\n".format(_type[i])
            self.error_array.append(error)
            self.parent.allowaddrecs = False  # do not allow this speedcell be be input into database
            return
        if not date_object.check_year():  # returns False if the year does not have 4 digits.
            error = "     ERROR: The year entered for the {} date must have 4 digits.\n".format(_type[i])
            self.error_array.append(error)
            self.parent.allowaddrecs = False  # do not allow this speedcell be be input into database
            return
        if not date_object.valid_date():  # returns False if the date is not a valid date
            error = "     ERROR: The date entered for the {} date is not a valid date.\n".format(_type[i])
            self.error_array.append(error)
            self.parent.allowaddrecs = False  # do not allow this speedcell be be input into database
            return
        # this removes white space from the date and each element of the date.
        self.input_date[i] = self.reformat_date(i)
        # convert the input date into a string of a datetime object.
        self.input_date[i] = Convert(self.input_date[i]).backslashdate_to_dtstring()
        self.add_date(i)  # add the dates to add_date variables

    def reformat_date(self, i):
        """ this removes white space from the date and each element of the date. """
        breakdown = self.input_date[i].strip()
        breakdown = breakdown.split("/")
        month = breakdown[0].strip()
        day = breakdown[1].strip()
        year = breakdown[2].strip()
        return "{}/{}/{}".format(month, day, year)

    def add_date(self, i):
        """ add the dates to add_date variables
         this is self.addstartdate, self.addenddate and self.addmeetingdate
         a counter is passed from the self.check_date method above. """
        onrec_date = [self.onrec_startdate, self.onrec_enddate, self.onrec_meetingdate]
        _type = ("start", "end", "meeting")
        if self.input_date[i] == onrec_date[i]:  # if the new input and the old record are the same - do nothing
            pass  # retain "empty" value for grievant variable
        else:
            fyi = "     FYI: New or updated {} date: {}\n".format(_type[i], self.input_date[i])
            self.fyi_array.append(fyi)
            self.adddate[i] = self.input_date[i]  # save to input to dbase

    def check_issue(self):
        """ check the issue input """
        self.issue = self.issue.strip()  # strip out any whitespace before or after the string
        if self.issue == "":  # accept blank entries
            return
        if isint(self.issue):  # identify issue index entries and execute as valid - this also update the article
            self.check_issue_index()
            return
        self.check_issue_description()

    def check_issue_index(self):
        """ check that the issue index provided by the user is valid.
        use arrays of issue categories and articles collected in the SpeedSheetCheck class"""
        if self.issue in self.parent.issue_index:
            self.addissue = self.parent.issue_description[int(self.issue)-1]
            self.addarticle = self.parent.issue_article[int(self.issue)-1]
            fyi = "     FYI: New or updated issue and article (issue index entry): {} Article: {}\n"\
                .format(self.addissue, self.addarticle)
            self.fyi_array.append(fyi)
            return
        error = "     ERROR: The number for issue is in the index of issues. Got: {}\n".format(self.issue)
        self.error_array.append(error)
        self.parent.allowaddrecs = False  # do not allow this speedcell be be input into database

    def check_issue_description(self):
        """ check if the issue description is already in the list of issues. If so, update article. """
        if self.issue in self.parent.issue_description:
            index = self.parent.issue_description.index(self.issue)
            self.addarticle = self.parent.issue_article[index]
            fyi = "     FYI: New or updated issue and article (issue description entry): {} Article: {}\n" \
                .format(self.addissue, self.addarticle)
            self.add_issue(fyi)
            return
        fyi = "     FYI: New or updated issue: {}\n" \
            .format(self.addissue)
        self.add_issue(fyi)

    def add_issue(self, msg):
        """ add the issue to the add issue var """
        if self.issue == self.onrec_issue:
            pass
        else:
            self.addissue = self.issue
            self.fyi_array.append(msg)

    def check_article(self):
        """ check the article input """
        self.article = self.article.strip()
        if not self.article:
            self.add_article()
            return
        if not isint(self.issue):
            error = "     ERROR: The number the article must be a whole number. Got: {}\n".format(self.issue)
            self.error_array.append(error)
            self.parent.allowaddrecs = False  # do not allow this speedcell be be input into database
            return

    def add_article(self):
        """ add the article to the add_article var """
        if self.article == self.onrec_article:
            return
        else:
            fyi = "     FYI: New or updated article: {}\n".format(self.article)
            self.fyi_array.append(fyi)
            self.addarticle = self.article

    def add_recs(self):
        """ add records using the add___ vars. """
        chg_these = []
        if not self.onrec:  # if there is no record of the grievance number in the db informalc_grievance table
            add = "     INPUT: New Grievance Number added to database >>{}\n" \
                .format(self.grv_no)  # report
            self.add_array.append(add)
            chg_these.append('grv_no')
        if not self.parent.allowaddrecs:  # if all checks passed
            return
        # get grievant place
        if self.addgrievant != "empty":
            add = "     INPUT: Grievant added or updated to database >>{}\n" \
                .format(self.addgrievant)  # report
            self.add_array.append(add)
            chg_these.append("grievant")
            grievant_place = self.addgrievant
        else:
            grievant_place = self.onrec_grievant
        # get date places using loop
        onrec_date = [self.onrec_startdate, self.onrec_enddate, self.onrec_meetingdate]
        startdate_place = None
        enddate_place = None
        meetingdate_place = None
        date_place = [startdate_place, enddate_place, meetingdate_place]
        chg_notation = ("startdate", "enddate", "meetingdate")
        _type = ("Start", "End", "Meeting")
        for i in range(3):
            if self.adddate[i] != "empty":
                add = "     INPUT: {} Date added or updated to database >>{}\n".format(_type[i], self.adddate[i])
                self.add_array.append(add)
                chg_these.append(chg_notation[i])
                date_place[i] = self.adddate[i]
            else:
                date_place[i] = onrec_date[i]
        # get issue place
        if self.addissue != "empty":
            add = "     INPUT: Issue added or updated to database >>{}\n".format(self.addissue)  # report
            self.add_array.append(add)
            chg_these.append("issue")
            issue_place = self.addissue
        else:
            issue_place = self.onrec_issue
        # get article place
        # the addarticle might be assigned a value in self.check_issue_description() so check against onrec
        if self.addarticle == self.onrec_article:
            article_place = self.onrec_article
        elif self.addarticle != "empty":
            add = "     INPUT: Article added or updated to database >>{}\n".format(self.addarticle)  # report
            self.add_array.append(add)
            chg_these.append("article")
            article_place = self.addarticle
        else:
            article_place = self.onrec_article
        # if any values have changed - form sql statements using _place vars and commit to db.
        if len(chg_these) != 0:  # if change these is empty, then there is no need to insert/update records
            if not self.onrec:  # if there is no rec on file for the grievance, insert the first rec
                sql = "INSERT INTO informalc_grievances(grievant, station, grv_no, startdate, enddate, " \
                      "meetingdate, issue, article) VALUES('%s','%s','%s','%s','%s','%s','%s','%s')" \
                      % (grievant_place, self.parent.station, self.grv_no, date_place[0], date_place[1],
                         date_place[2], issue_place, article_place)
            else:  # update the first rec to replace pre existing record.
                sql = "UPDATE informalc_grievances SET grievant='%s', startdate='%s', enddate ='%s', " \
                      "meetingdate='%s', issue='%s', article='%s' WHERE grv_no='%s' and station='%s'" \
                      % (grievant_place, date_place[0], date_place[1], date_place[2], issue_place, article_place,
                         self.grv_no, self.parent.station)
            commit(sql)

    def generate_report(self):
        """ generate a report """
        self.parent.fatal_rpt += len(self.error_array)
        if len(self.add_array):  # if there is anything in the add array - increment the add report by 1
            self.parent.add_rpt += 1
        if len(self.fyi_array):  # if there is anything in the fyi array - increment the add report by 1
            self.parent.fyi_rpt += 1
        if not self.parent.interject:
            master_array = self.error_array + self.attn_array  # use these reports for precheck
            if self.parent.fullreport:  # if the full report option is selected...
                master_array += self.fyi_array   # include the fyi messages.
        else:
            master_array = self.error_array + self.attn_array  # use these reports for input
            if self.parent.fullreport:  # if the full report option is selected...
                master_array += self.add_array  # include the adds messages.
        if len(master_array) > 0:
            if not self.parent.name_mentioned:
                self.parent.report.write("\nGrievance Number: {}\n".format(self.grv_no))
                self.parent.name_mentioned = True
            self.parent.report.write("   >>> sheet: \"{}\" --> row: \"{}\"  <<<\n".format(self.sheet, self.row))
            if not self.parent.allowaddrecs:
                self.parent.report.write("     GRIEVANCE RECORD ENTRY PROHIBITED: Correct errors!\n")
            for rpt in master_array:  # write all reports that have been keep in arrays.
                self.parent.report.write(rpt)


class SpeedSetCheck:
    """ checks one line of the settlement speedsheet when it is called by the SpeedSheetCheck class. """
    def __init__(self, parent, sheet, row, grv_no, level, datesigned, decision, proofdue, docs, gatsnumber):
        self.parent = parent
        self.sheet = sheet  # input here is coming directly from the speedcell
        self.row = str(row)
        self.grv_no = grv_no
        self.level = level
        self.datesigned = datesigned
        self.decision = decision
        self.proofdue = proofdue
        self.input_date = []  # array to hold datesigned and proofdue - form in check_dates()
        self.docs = docs
        self.gatsnumber = gatsnumber
        self.onrec = False  # this value is True if a sql search shows that there is a rec of the grv_no in the db.
        self.onrec_grv_no = None
        self.onrec_level = None
        self.onrec_datesigned = None
        self.onrec_decision = None
        self.onrec_proofdue = None
        self.onrec_docs = None
        self.onrec_gatsnumber = None
        self.addlevel = "empty"  # post checked values
        self.adddatesigned = "empty"
        self.adddecision = "empty"
        self.addproofdue = "empty"
        self.adddate = [self.adddatesigned, self.addproofdue]  # holds date values for self.add_date() loop
        self.adddocs = "empty"
        self.addgatsnumber = "empty"
        self.error_array = []  # gives a report of failed checks
        self.attn_array = []  # gives a report of issues to bring to the attention of users
        self.add_array = []  # gives a report of records to add to the database
        self.fyi_array = []  # gives a report of useful information for the user
        self.parent.name_mentioned = False  # reset this so that name is not repeated on reports
        self.levelarray = ("informal a", "formal a", "step b", "pre arb", "arbitration")
        self.docsarray = ("non-applicable", "no", "yes", "unknown", "yes - not paid", "yes - in part", 
                          "yes - verified", "no - moot", "no - ignore")

    def check_all(self):
        """ master method to run other methods. """
        if self.check_grv_number():  # check the grievance number input
            self.get_onrecs()
            self.check_level()
            self.check_dates()
            self.check_decision()
            self.check_docs()
            self.check_gatsnumber()
            self.add_recs()
        self.generate_report()

    def check_grv_number(self):
        """ check the grievance number input """
        if not GrievanceChecker(self.grv_no).has_value():
            error = "     ERROR: The grievance number must not be blank. \n"
            self.error_array.append(error)
            self.parent.allowaddrecs = False  # do not allow this speedcell be be input into database
            return False
        # check that there is a record of the grievance in informalc_grievances
        sql = "SELECT * FROM informalc_grievances WHERE grv_no = '%s' and station = '%s'" \
              % (self.grv_no, self.parent.station)
        result = inquire(sql)
        if not result:
            error = "     ERROR: There is no prior record of the grievance. \n"
            self.error_array.append(error)
            self.parent.allowaddrecs = False  # do not allow this speedcell be be input into database
            return False
        if not GrievanceChecker(self.grv_no).check_characters():
            error = "     ERROR: The grievance number can only contain numbers and letters. \n"
            self.error_array.append(error)
            self.parent.allowaddrecs = False  # do not allow this speedcell be be input into database
            return False
        if not GrievanceChecker(self.grv_no).min_lenght():
            error = "     ERROR: The grievance number must contain at least 4 characters. \n"
            self.error_array.append(error)
            self.parent.allowaddrecs = False  # do not allow this speedcell be be input into database
            return False
        if not GrievanceChecker(self.grv_no).max_lenght():
            error = "     ERROR: The grievance number can not contain more than 20 characters. \n"
            self.error_array.append(error)
            self.parent.allowaddrecs = False  # do not allow this speedcell be be input into database
            return False
        return True

    def get_onrecs(self):
        """ check if there is an existing record for the grievance number in the informalc grievances table.
        if so, store the values in the self.onrec variables. if not, the self.onrec variables default to empty. """
        sql = "SELECT * FROM informalc_settlements WHERE grv_no = '%s'" % self.grv_no
        results = inquire(sql)
        if results:
            self.onrec = True  # this value is True if a sql search shows that there is a rec in the db.
            # skip grievance number as that is self.grv_no and is part of the search criteria
            self.onrec_level = results[0][1]
            self.onrec_datesigned = results[0][2]
            self.onrec_decision = results[0][3]
            self.onrec_proofdue = results[0][4]
            self.onrec_docs = results[0][5]
            self.onrec_gatsnumber = results[0][6]

    def check_level(self):
        """ check the grievance number input """
        self.level = self.level.strip()
        self.level = self.level.lower()
        if not self.level:
            pass
        if self.level not in self.levelarray:
            error = "     ERROR: The level must be either 'informal a', 'formal a', 'step b', 'pre arb' or \n" \
                    "            'arbitration'. No other values are allowed. \n"
            self.error_array.append(error)
            self.parent.allowaddrecs = False  # do not allow this speedcell be be input into database
            return False
        self.add_level()

    def add_level(self):
        """ add level to the self.addlevel var """
        if self.level == self.onrec_level:
            pass
        else:
            fyi = "     FYI: New or updated level: {}\n".format(self.level)
            self.fyi_array.append(fyi)
            self.addlevel = self.level

    def check_dates(self):
        """ check the startdate, enddate and meetingdate.
         since these are all dates with similiar criteria, use a loop to check them.
         sometimes, openpyxl sends the dates as strings of datetime objects, instead of the mm/dd/yyyy formated dates,
         the DateTimeChecker() will identify these and skip the checks. """
        self.input_date = [self.datesigned, self.proofdue]
        for i in range(2):
            self.check_date_loop(i)

    def check_date_loop(self, i):
        """ loop from check dates """
        _type = ("date signed", "proof due")
        if self.input_date[i].strip() == "":  # if the value is blank, skip all the checks
            self.add_date(i)
            return
        # if the value is a valid dt object, skip all the checks
        if DateTimeChecker().check_dtstring(self.input_date[i]):
            self.add_date(i)
            return
        date_object = BackSlashDateChecker(self.input_date[i])  # first create the date_object
        if not date_object.count_backslashes():  # this checks that there are 2 backslashes in the date
            error = "     ERROR: The date for the {} date must have two backslashes. Got instead: {}\n"\
                .format(_type[i], self.input_date[i])
            self.error_array.append(error)
            self.parent.allowaddrecs = False  # do not allow this speedcell be be input into database
            return
        date_object.breaker()  # this breaks the object into month, day and year elements.
        if not date_object.check_numeric():  # check each element in the date to ensure they are numeric
            error = "     ERROR: The month, day and year for the {} date must be numeric\n".format(_type[i])
            self.error_array.append(error)
            self.parent.allowaddrecs = False  # do not allow this speedcell be be input into database
            return
        if not date_object.check_minimums():  # check each element in the date to ensure they are greater than zero
            error = "     ERROR: The month, day and year for the {} date must be greater than zero.\n"\
                .format(_type[i])
            self.error_array.append(error)
            self.parent.allowaddrecs = False  # do not allow this speedcell be be input into database
            return
        if not date_object.check_month():  # returns False if the month is greater than 12.
            error = "     ERROR: The month for the {} date must less than 13.\n".format(_type[i])
            self.error_array.append(error)
            self.parent.allowaddrecs = False  # do not allow this speedcell be be input into database
            return
        if not date_object.check_day():  # return False if the day is greater than 31.
            error = "     ERROR: The day entered for the {} date is must be less than 32.\n".format(_type[i])
            self.error_array.append(error)
            self.parent.allowaddrecs = False  # do not allow this speedcell be be input into database
            return
        if not date_object.check_year():  # returns False if the year does not have 4 digits.
            error = "     ERROR: The year entered for the {} date must have 4 digits.\n".format(_type[i])
            self.error_array.append(error)
            self.parent.allowaddrecs = False  # do not allow this speedcell be be input into database
            return
        if not date_object.valid_date():  # returns False if the date is not a valid date
            error = "     ERROR: The date entered for the {} date is not a valid date.\n".format(_type[i])
            self.error_array.append(error)
            self.parent.allowaddrecs = False  # do not allow this speedcell be be input into database
            return
        # this removes white space from the date and each element of the date.
        self.input_date[i] = self.reformat_date(i)
        # convert the input date into a string of a datetime object.
        self.input_date[i] = Convert(self.input_date[i]).backslashdate_to_dtstring()
        self.add_date(i)  # add the dates to add_date variables

    def reformat_date(self, i):
        """ this removes white space from the date and each element of the date. """
        breakdown = self.input_date[i].strip()
        breakdown = breakdown.split("/")
        month = breakdown[0].strip()
        day = breakdown[1].strip()
        year = breakdown[2].strip()
        return "{}/{}/{}".format(month, day, year)

    def add_date(self, i):
        """ add the dates to add_date variables
         this is self.addstartdate, self.addenddate and self.addmeetingdate
         a counter is passed from the self.check_date method above. """
        onrec_date = [self.onrec_datesigned, self.onrec_proofdue]
        _type = ("date signed", "proof due")
        if self.input_date[i] == onrec_date[i]:  # if the new input and the old record are the same - do nothing
            pass  # retain "empty" value for grievant variable
        else:
            fyi = "     FYI: New or updated {} date: {}\n".format(_type[i], self.input_date[i])
            self.fyi_array.append(fyi)
            self.adddate[i] = self.input_date[i]  # save to input to dbase

    def check_decision(self):
        """ check the decision input """
        self.decision = self.decision.strip()  # strip out any whitespace before or after the string
        if self.decision == "":  # accept blank entries
            msg = ""
            self.add_decision(msg)
        elif isint(self.decision):  # identify decision index entries and execute as valid - this also updates article
            self.check_decision_index()
            return
        self.check_decision_description()

    def check_decision_index(self):
        """ check that the decision index provided by the user is valid.
        use arrays of decision categories and articles collected in the SpeedSheetCheck class"""
        if self.decision in self.parent.decision_index:
            self.adddecision = self.parent.decision_description[int(self.decision)-1]
            fyi = "     FYI: New or updated decision (decision index entry): {}\n"\
                .format(self.adddecision)
            self.fyi_array.append(fyi)
            return
        error = "     ERROR: The number for decision is in the index of decisions. Got: {}\n".format(self.decision)
        self.error_array.append(error)
        self.parent.allowaddrecs = False  # do not allow this speedcell be be input into database

    def check_decision_description(self):
        """ check if the decision description is already in the list of decisions. If so, update article. """
        if self.decision in self.parent.decision_description:
            fyi = "     FYI: New or updated decision and article (decision description entry):\n"\
                .format(self.adddecision)
            self.add_decision(fyi)
            return
        fyi = "     FYI: New or updated decision: {}\n".format(self.adddecision)
        self.add_decision(fyi)

    def add_decision(self, msg):
        """ add the decision to the add decision var """
        if self.decision == self.onrec_decision:
            pass
        else:
            self.adddecision = self.decision
            if msg:
                self.fyi_array.append(msg)

    def check_docs(self):
        """ check the grievance number input """
        self.docs = self.docs.strip()
        self.docs = self.docs.lower()
        if not self.docs:
            pass
        elif self.docs in self.docsarray:
            pass
        else:
            print(self.docs, type(self.docs))
            error = "     ERROR: The docs input must be either 'non-applicable', 'no', 'yes', 'unknown', \n" \
                    "            'yes - not paid', 'yes - in part', 'yes - verified', 'no - moot' or \n" \
                    "            'no - ignore'. No other values are allowed. Got: {}\n".format(self.docs)
            self.error_array.append(error)
            self.parent.allowaddrecs = False  # do not allow this speedcell be be input into database
            return False
        self.add_docs()

    def add_docs(self):
        """ add docs to the self.adddocs var """
        if self.docs == self.onrec_docs:
            pass
        else:
            fyi = "     FYI: New or updated docs: {}\n".format(self.docs)
            self.fyi_array.append(fyi)
            self.adddocs = self.docs

    def check_gatsnumber(self):
        """ check the article input - this is an open field that takes almost anything with no limits or indexes. """
        self.gatsnumber = self.gatsnumber.strip()
        self.gatsnumber = self.gatsnumber.lower()
        if not self.gatsnumber:
            pass
        if len(self.gatsnumber) < 30:
            pass
        else:
            error = "     ERROR: The gats number must not be longer than 30 characters.  \n"
            self.error_array.append(error)
            self.parent.allowaddrecs = False  # do not allow this speedcell be be input into database
            return False
        self.add_gatsnumber()

    def add_gatsnumber(self):
        """ add gats number to the self.addgatsnumber var """
        if self.gatsnumber == self.onrec_gatsnumber:
            pass
        else:
            fyi = "     FYI: New or updated gats number: {}\n".format(self.gatsnumber)
            self.fyi_array.append(fyi)
            self.addgatsnumber = self.gatsnumber

    def add_recs(self):
        """ add records using the add___ vars. """
        chg_these = []
        if not self.parent.allowaddrecs:  # if all checks passed
            return
        if not self.onrec:  # if there is no record of the grievance number in the db informalc_grievance table
            add = "     INPUT: New Grievance Number added to database >>{}\n" \
                .format(self.grv_no)  # report
            self.add_array.append(add)
            chg_these.append('grv_no')
        # get level place
        if self.addlevel != "empty":
            add = "     INPUT: Level added or updated to database >>{}\n" \
                .format(self.addlevel)  # report
            self.add_array.append(add)
            chg_these.append("level")
            level_place = self.addlevel
        else:
            level_place = self.onrec_level
        # get date places using loop
        onrec_date = [self.onrec_datesigned, self.onrec_proofdue]
        datesigned_place = None  # aka date_place[0]
        proofdue_place = None  # aka date_place[1]
        date_place = [datesigned_place, proofdue_place]
        chg_notation = ("datesigned", "proofdue")
        _type = ("Date Signed", "Proof Due Date")
        for i in range(2):
            if self.adddate[i] != "empty":
                add = "     INPUT: {} added or updated to database >>{}\n".format(_type[i], self.adddate[i])
                self.add_array.append(add)
                chg_these.append(chg_notation[i])
                date_place[i] = self.adddate[i]
            else:
                date_place[i] = onrec_date[i]

        # get decision place
        if self.adddecision != "empty":
            add = "     INPUT: Decision added or updated to database >>{}\n".format(self.adddecision)  # report
            self.add_array.append(add)
            chg_these.append("decision")
            decision_place = self.adddecision
        else:
            decision_place = self.onrec_decision

        # get docs place
        if self.adddocs != "empty":
            add = "     INPUT: Docs added or updated to database >>{}\n".format(self.adddocs)  # report
            self.add_array.append(add)
            chg_these.append("docs")
            docs_place = self.adddocs
        else:
            docs_place = self.onrec_docs

        # get gats place
        if self.addgatsnumber != "empty":
            add = "     INPUT: Gats Number added or updated to database >>{}\n".format(self.addgatsnumber)  # report
            self.add_array.append(add)
            chg_these.append("gatsnumber")
            gats_place = self.addgatsnumber
        else:
            gats_place = self.onrec_gatsnumber
        # if any values have changed - form sql statements using _place vars and commit to db.
        if len(chg_these) != 0:  # if change these is empty, then there is no need to insert/update records
            if not self.onrec:  # if there is no rec on file for the grievance, insert the first rec
                sql = "INSERT INTO informalc_settlements(grv_no, level, date_signed, decision, proofdue, " \
                      "docs, gats_number) VALUES('%s','%s','%s','%s','%s','%s','%s')" \
                      % (self.grv_no, level_place, date_place[0], decision_place, date_place[1], docs_place, gats_place)
            else:  # update the first rec to replace pre existing record.
                sql = "UPDATE informalc_settlements SET level='%s', date_signed='%s', decision ='%s', " \
                      "proofdue='%s', docs='%s', gats_number='%s' WHERE grv_no='%s'" \
                      % (level_place, date_place[0], decision_place, date_place[1], docs_place, gats_place,
                         self.grv_no)
            commit(sql)

    def generate_report(self):
        """ generate a report
        """
        self.parent.settlement_fatal_rpt += len(self.error_array)
        if len(self.add_array):  # if there is anything in the add array - increment the add report by 1
            self.parent.settlement_add_rpt += 1
        if len(self.fyi_array):  # if there is anything in the fyi array - increment the add report by 1
            self.parent.settlement_fyi_rpt += 1
        if not self.parent.interject:
            master_array = self.error_array + self.attn_array  # use these reports for precheck
            if self.parent.fullreport:  # if the full report option is selected...
                master_array += self.fyi_array   # include the fyi messages.
        else:
            master_array = self.error_array + self.attn_array  # use these reports for input
            if self.parent.fullreport:  # if the full report option is selected...
                master_array += self.add_array  # include the adds messages.
        if len(master_array) > 0:
            if not self.parent.name_mentioned:
                self.parent.report.write("\nGrievance Number: {}\n".format(self.grv_no))
                self.parent.name_mentioned = True
            self.parent.report.write("   >>> sheet: \"{}\" --> row: \"{}\"  <<<\n".format(self.sheet, self.row))
            if not self.parent.allowaddrecs:
                self.parent.report.write("     SETTLEMENT RECORD ENTRY PROHIBITED: Correct errors!\n")
            for rpt in master_array:  # write all reports that have been keep in arrays.
                self.parent.report.write(rpt)


class SpeedIndexCheck:
    """ checks one line of the settlement speedsheet when it is called by the SpeedSheetCheck class. """
    def __init__(self, parent, sheet, row, first, second):
        self.parent = parent
        self.sheet = sheet  # input here is coming directly from the speedcell
        self.row = str(row)
        self.first = first
        self.second = second
        self.onrec_first = None
        self.onrec_second = None
        self.error_array = []
        self.grv_no = ""

    def check_all(self):
        """ master method to run other methods. """
        self.generate_report()

    def check_first(self):
        """ check the grievant input """
        if not NameChecker(self.grv_no).check_characters():
            error = "     ERROR: Carrier name can not contain numbers or most special characters\n"
            self.error_array.append(error)
            self.parent.allowaddrecs = False  # do not allow this speedcell be be input into database

    def check_second(self):
        """ check the grievance number input """
        pass

    def generate_report(self):
        """" generate the text report """
        pass


class ProgressBarIn:
    """ Indeterminate Progress Bar """

    def __init__(self, title="", label="", text=""):
        self.title = title
        self.label = label
        self.text = text
        self.pb_root = Tk()  # create a window for the progress bar
        self.pb_label = Label(self.pb_root, text=self.label)  # make label for progress bar
        self.pb = ttk.Progressbar(self.pb_root, length=400, mode="indeterminate")  # create progress bar
        self.pb_text = Label(self.pb_root, text=self.text, anchor="w")

    def start_up(self):
        """ starts up the progress bar. """
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
        """ stops and destroys the progress bar. """
        self.pb.stop()  # stop and destroy the progress bar
        self.pb_text.destroy()
        self.pb_label.destroy()  # destroy the label for the progress bar
        self.pb.destroy()
        self.pb_root.destroy()


if __name__ == "__main__":
    """ this is where the program starts if not launched from another app. """
    InformalC().informalc(None)
