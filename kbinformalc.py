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
    isint, NewWindow, titlebar_icon, informalc_date_checker
from tkinter import *
from tkinter import messagebox, ttk
from datetime import datetime, timedelta
import os
import shutil
import sys
import subprocess
# define globals
global root  # used to hold the Tk() root for the new window used by all Informal C windows.


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

    def informalc(self, frame):
        """ a master method for running the other methods in proper sequence. """
        self.win.create(frame)
        self.build_tables()  # build needed tables if they do not exist.
        self.build_screen()  # this fills the screen with widgets.
        self.win.finish()  # this commands the window to loop and persist.
            
    @staticmethod
    def build_tables():
        """ build tables needed if they do no exist. """
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
            
    def build_screen(self):
        """ the main screen for informal c. """
        Label(self.win.body, text="Informal C", font=macadj("bold", "Helvetica 18")).grid(row=0, sticky="w")
        Label(self.win.body, text="The C is for Compliance").grid(row=1, sticky="w")
        Label(self.win.body, text="").grid(row=2)
        Button(self.win.body, text="New Settlement", width=30,
               command=lambda: self.New(self).informalc_new(self.win.topframe)).grid(row=3, pady=5)
        Button(self.win.body, text="Settlement List", width=30,
               command=lambda: self.GrvList(self).grvlist_search(self.win.topframe)).grid(row=4, pady=5)
        Button(self.win.body, text="Payout Entry", width=30,
               command=lambda: self.PayoutEntry(self).poe_search(self.win.topframe)).grid(row=5, pady=5)
        Button(self.win.body, text="Payout Report", width=30,
               command=lambda: self.PayoutReport(self).informalc_por(self.win.topframe)).grid(row=6, pady=5)
        Label(self.win.body, text="", width=70).grid(row=7)
        button_back = Button(self.win.buttons)
        button_back.config(text="Quit Informal C", width=20, command=lambda: self.win.root.destroy())
        if sys.platform == "win32":
            button_back.config(anchor="w")
        button_back.grid(row=0, column=0)

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
            Label(self.win.body, text="Grievance Number: ", background=macadj("gray95", "grey"),
                  fg=macadj("black", "white"), width=macadj(22, 20), anchor="w", height=macadj(1, 1)) \
                .grid(row=2, column=0, sticky="w")
            
            Entry(self.win.body, textvariable=self.grv_no, justify='right', width=macadj(20, 15)) \
                .grid(row=2, column=1, sticky="w")
            Label(self.win.body, text="Incident Date").grid(row=3, column=0, sticky="w")
            Label(self.win.body, text="  Start (mm/dd/yyyy): ", background=macadj("gray95", "grey"),
                  fg=macadj("black", "white"), width=macadj(22, 20), anchor="w") \
                .grid(row=4, column=0, sticky="w")
            
            Entry(self.win.body, textvariable=self.incident_start, justify='right', width=macadj(20, 15)) \
                .grid(row=4, column=1, sticky="w")
            Label(self.win.body, text="  End (mm/dd/yyyy): ", background=macadj("gray95", "grey"),
                  fg=macadj("black", "white"), width=macadj(22, 20), anchor="w", height=macadj(1, 1)) \
                .grid(row=5, column=0, sticky="w")
            
            Entry(self.win.body, textvariable=self.incident_end, justify='right', width=macadj(20, 15)) \
                .grid(row=5, column=1, sticky="w")
            Label(self.win.body, text="Date Signed (mm/dd/yyyy): ", background=macadj("gray95", "grey"),
                  fg=macadj("black", "white"), width=macadj(22, 20), anchor="w", height=macadj(1, 1)) \
                .grid(row=6, column=0, sticky="w")
            
            Entry(self.win.body, textvariable=self.date_signed, justify='right', width=macadj(20, 15)) \
                .grid(row=6, column=1, sticky="w")
            # select level
            Label(self.win.body, text="Settlement Level: ", background=macadj("gray95", "grey"),
                  fg=macadj("black", "white"), width=macadj(22, 20), anchor="w", height=macadj(1, 1)) \
                .grid(row=7, column=0, sticky="w")  # select settlement level
            
            lvl_options = ("informal a", "formal a", "step b", "pre arb", "arbitration")
            lvl_om = OptionMenu(self.win.body, self.lvl, *lvl_options)
            lvl_om.config(width=macadj(13, 13))
            lvl_om.grid(row=7, column=1)
            self.lvl.set("informal a")
            Label(self.win.body, text="Station: ", background=macadj("gray95", "grey"),  # select a station
                  fg=macadj("black", "white"), width=macadj(22, 20), anchor="w", height=macadj(1, 1)). \
                grid(row=8, column=0, sticky="w")
            Label(self.win.body, text="", height=macadj(1, 2)).grid(row=8, column=1)
            
            self.station.set("Select a Station")
            station_options = projvar.list_of_stations
            if "out of station" in station_options:
                station_options.remove("out of station")
            station_om = OptionMenu(self.win.body, self.station, *station_options)
            station_om.config(width=macadj(40, 34))
            station_om.grid(row=9, column=0, columnspan=2, sticky="e")
            Label(self.win.body, text="GATS Number: ", background=macadj("gray95", "grey"),
                  fg=macadj("black", "white"), width=macadj(22, 20), anchor="w", height=macadj(1, 1)) \
                .grid(row=10, column=0, sticky="w")  # enter gats number
            
            Entry(self.win.body, textvariable=self.gats_number, justify='right', width=macadj(20, 15)) \
                .grid(row=10, column=1, sticky="w")
            Label(self.win.body, text="Documentation?: ", background=macadj("gray95", "grey"),
                  fg=macadj("black", "white"), width=macadj(22, 20), anchor="w", height=macadj(1, 1)) \
                .grid(row=11, column=0, sticky="w")  # select documentation
            
            doc_options = ("moot", "no", "partial", "yes", "incomplete", "verified")
            docs_om = OptionMenu(self.win.body, self.docs, *doc_options)
            docs_om.config(width=macadj(13, 13))
            docs_om.grid(row=11, column=1)
            self.docs.set("no")
            Label(self.win.body, text="Description: ", background=macadj("gray95", "grey"),
                  fg=macadj("black", "white"), width=macadj(22, 20), anchor="w", height=macadj(1, 1)) \
                .grid(row=15, column=0, sticky="w")
            Label(self.win.body, text="", height=macadj(1, 2)).grid(row=15, column=1)
            
            Entry(self.win.body, textvariable=self.description, width=macadj(48, 36), justify='right') \
                .grid(row=16, column=0, sticky="w", columnspan=2)
            Label(self.win.body, text="", height=macadj(1, 1)).grid(row=17, column=0)
            Label(self.win.body, text=self.msg, fg="red", height=macadj(1, 1))\
                .grid(row=18, column=0, columnspan=2, sticky="w")
            Button(self.win.buttons, text="Go Back", width=macadj(19, 18), anchor="w",
                   command=lambda: self.parent.informalc(self.win.topframe)).grid(row=0, column=0)
            Button(self.win.buttons, text="Enter", width=macadj(19, 18),
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
                    d = date.get().split("/")
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
                if self.grv_no.get() in existing_grv:
                    messagebox.showerror("Data Entry Error",
                                         "The Grievance Number {} is already present in the database. You can not "
                                         "create a duplicate.".format(self.grv_no.get()),
                                         parent=self.win.topframe)
                    return
                sql = "INSERT INTO informalc_grv (grv_no, indate_start, indate_end, date_signed, station, " \
                      "gats_number, docs, description, level) " \
                      "VALUES('%s','%s','%s','%s','%s','%s','%s','%s','%s')" % \
                      (self.grv_no.get().lower(), dt_dates[0], dt_dates[1], dt_dates[2], self.station.get(),
                      self.gats_number.get().strip(), self.docs.get(), self.description.get(), self.lvl.get())
                commit(sql)
                self.msg = "Grievance Settlement Added: #{}.".format(self.grv_no.get().lower())
                self.informalc_new(self.win.topframe)

        def informalc_check_grv(self):
            """ checks the grievance number. """
            if self.station.get() == "Select a Station":
                messagebox.showerror("Invalid Data Entry",
                                     "You must select a station.",
                                     parent=self.win.topframe)
                return False
            if self.grv_no.get().strip() == "":
                messagebox.showerror("Invalid Data Entry",
                                     "You must enter a grievance number",
                                     parent=self.win.topframe)
                return False
            if re.search('[^1234567890abcdefghijklmnopqrstuvwxyz:ABCDEFGHIJKLMNOPQRSTUVWXYZ,]', self.grv_no.get()):
                messagebox.showerror("Invalid Data Entry",
                                     "The grievance number can only contain numbers and letters. No other "
                                     "characters are allowed",
                                     parent=self.win.topframe)
                return False
            if len(self.grv_no.get()) < 8:
                messagebox.showerror("Invalid Data Entry",
                                     "The grievance number must be at least eight characters long",
                                     parent=self.win.topframe)
                return False
            if len(self.grv_no.get()) > 20:
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
            Label(self.win.body, text=" Station ", background=macadj("gray95", "grey"), fg=macadj("black", "white"),
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
            Label(self.win.body, text=" Incident Dates", background=macadj("gray95", "grey"), 
                  fg=macadj("black", "white"), anchor="w", width=14).grid(row=4, column=3, sticky="w")
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
            Label(self.win.body, text=" Signing Dates", background=macadj("gray95", "grey"), 
                  fg=macadj("black", "white"), anchor="w", width=14).grid(row=5, column=3, sticky="w")
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
            Label(self.win.body, text=" Settlement Level ", background=macadj("gray95", "grey"), 
                  fg=macadj("black", "white"), anchor="w", width=14, height=1).grid(row=6, column=3, sticky="w")
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
            Label(self.win.body, text=" GATS Number", background=macadj("gray95", "grey"), fg=macadj("black", "white"),
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
            Label(self.win.body, text=" Documentation", background=macadj("gray95", "grey"), 
                  fg=macadj("black", "white"), anchor="w", width=14, height=1).grid(row=9, column=3, sticky="w")
            doc_options = ("moot", "no", "partial", "yes", "incomplete", "verified")
            docs_om = OptionMenu(self.win.body, self.have_docs, *doc_options)
            docs_om.config(width=macadj(10, 8))
            docs_om.grid(row=9, column=4, columnspan=3, sticky="e")
            self.have_docs.set('no')
            self.docs.set("no")
            Label(self.win.body, text="").grid(row=13)
        
        def build_buttons(self):
            """ build the buttons on the bottom of the screen. """
            Button(self.win.buttons, text="Search", width=20,
                   command=lambda: self.grvlist_apply()).grid(row=0, column=1)
            Button(self.win.buttons, text="Go Back", width=20, anchor="w", 
                   command=lambda: self.parent.informalc(self.win.topframe)).grid(row=0, column=0)

        def grvlist_apply(self):
            """ applies changes to the grievance list after a check. """
            conditions = []
            if self.incident_date.get() == "yes":
                check = informalc_date_checker(self.win.topframe, self.incident_start, "starting incident date")
                if check == "fail":
                    return
                check = informalc_date_checker(self.win.topframe, self.incident_end, "ending incident date")
                if check == "fail":
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
                check = informalc_date_checker(self.win.topframe, self.signing_start, "starting signing date")
                if check == "fail":
                    return
                check = informalc_date_checker(self.win.topframe, self.signing_end, "ending signing date")
                if check == "fail":
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
            Button(self.win.body, width=9, text="update", command=lambda:
            self.grvchange(self.grv_num, self.grv_no)).grid(row=3, column=1, sticky="e")
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
            sql = "UPDATE informalc_grv SET indate_start='%s',indate_end='%s',date_signed='%s',station='%s'," \
                  "gats_number='%s', docs='%s',description='%s', level='%s' WHERE grv_no='%s'" % \
                  (dt_dates[0], dt_dates[1], dt_dates[2], self.station.get(), self.gats_number.get().strip(),
                   self.edit_docs.get(), self.description.get(), self.lvl.get(), self.grv_no.get())
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
                if new_num.get().strip() == "":
                    messagebox.showerror("Invalid Data Entry",
                                         "You must enter a grievance number",
                                         parent=self.win.topframe)
                    return "fail"
                if not new_num.get().isalnum():
                    messagebox.showerror("Invalid Data Entry",
                                         "The grievance number can only contain numbers and letters. No other "
                                         "characters are allowed",
                                         parent=self.win.topframe)
                    return "fail"
                if len(new_num.get()) < 8:
                    messagebox.showerror("Invalid Data Entry",
                                         "The grievance number must be at least eight characters long",
                                         parent=self.win.topframe)
                    return "fail"
                if len(new_num.get()) > 16:
                    messagebox.showerror("Invalid Data Entry",
                                         "The grievance number must not exceed 16 characters in length.",
                                         parent=self.win.topframe)
                    return "fail"
                sql = "SELECT grv_no FROM informalc_grv WHERE grv_no = '%s'" % new_num.get().lower()
                result = inquire(sql)
                if result:
                    messagebox.showerror("Grievance Number Error",
                                         "This number is already being used for another grievance.",
                                         parent=self.win.topframe)
                    return "fail"

                sql = "UPDATE informalc_grv SET grv_no = '%s' WHERE grv_no = '%s'" % (new_num.get().lower(), old_num)
                commit(sql)
                sql = "UPDATE informalc_awards SET grv_no = '%s' WHERE grv_no = '%s'" % (new_num.get().lower(), old_num)
                commit(sql)
                for record in l_passed_result:
                    if record[0] == old_num:
                        record[0] = new_num.get().lower()
                self.msg = "The grievance number has been changed."
                self.search_result = l_passed_result[:]
                self.edit(new_num.get().lower())

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
            self.poe_listbox = None  # holds the list box object
            
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
            Button(self.win.buttons, text="Go Back", width=20, 
                   command=lambda: self.parent.informalc(self.win.topframe))\
                .grid(row=0, column=1, sticky="w")
            Button(self.win.buttons, text="Apply", width=20,
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
                self.poe_listbox.destroy()
            except TclError:
                pass
            self.poe_search(frame)

        def poe_listbox(self, dt_year, station, dt_start, year):
            """ pay out entry - create a listbox which allows the user to add carriers. """
            poe_root = Tk()
            self.poe_listbox = poe_root  # set the value
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
            Button(self.win.buttons, text="Go Back", width=16,
                   command=lambda: self.parent.informalc(self.win.topframe)).grid(row=0, column=0)
            Label(self.win.buttons, text="Report: ", width=16).grid(row=0, column=1)
            Button(self.win.buttons, text="All Carriers", width=16,
                   command=lambda: self.por_all(afterdate, beforedate, station, backdate)) \
                .grid(row=0, column=2)
            # Button(self.win.buttons, text="By Carrier", width=16).grid(row=0, column=3)
            self.win.finish()

        def por_all(self, afterdate, beforedate, station, backdate):
            """ pay out report. generates text report for all. """
            check = informalc_date_checker(self.win.topframe, afterdate, "After Date")
            if check == "fail":
                return
            check = informalc_date_checker(self.win.topframe, beforedate, "Before Date")
            if check == "fail":
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
            self.c.bind_all('<MouseWheel>', lambda event: self.c.yview_scroll
            (int(projvar.mousewheel * (event.delta / 120)), "units"))
        elif sys.platform == "darwin":
            self.c.bind_all('<MouseWheel>', lambda event: self.c.yview_scroll
            (int(projvar.mousewheel * event.delta), "units"))
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
        except KeyboardInterrupt:
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


if __name__ == "__main__":
    """ this is where the program starts if not launched from another app. """
    InformalC().informalc(None)
