from tkinter import Tk, ttk, Frame, Scrollbar, Canvas, BOTH, LEFT, BOTTOM, RIGHT, NW, Label, mainloop, \
    messagebox, TclError, PhotoImage
import projvar
import os
import sys
import sqlite3
from datetime import datetime, timedelta


def inquire(sql):
    # query the database
    if projvar.platform == "macapp":
        path = os.path.expanduser("~") + '/Documents/.klusterbox/mandates.sqlite'
    elif projvar.platform == "winapp":
        path = os.path.expanduser("~") + '\\Documents\\.klusterbox\\mandates.sqlite'
    elif projvar.platform == "py":
        path = "kb_sub/mandates.sqlite"
    else:
        path = "kb_sub/mandates.sqlite"
    db = sqlite3.connect(path)
    cursor = db.cursor()
    try:
        cursor.execute(sql)
        results = cursor.fetchall()
        return results
    except sqlite3.OperationalError:
        messagebox.showerror("Database Error",
                             "Unable to access database.\n"
                             "\n Attempted Query: {}".format(sql))
    db.close()


# write to the database
def commit(sql):
    if projvar.platform == "macapp":
        path = os.path.expanduser("~") + '/Documents/.klusterbox/mandates.sqlite'
    elif projvar.platform == "winapp":
        path = os.path.expanduser("~") + '\\Documents\\.klusterbox\\mandates.sqlite'
    elif projvar.platform == "py":
        path = "kb_sub/mandates.sqlite"
    else:
        path = "kb_sub/mandates.sqlite"
    db = sqlite3.connect(path)
    cursor = db.cursor()
    try:
        cursor.execute(sql)
        db.commit()
        db.close()
    except sqlite3.OperationalError:
        messagebox.showerror("Database Error",
                             "Unable to access database.\n"
                             "\n Attempted Query: {}".format(sql))


def titlebar_icon(root):  # place icon in titlebar
    if sys.platform == "win32" and projvar.platform == "py":
        try:
            root.iconbitmap(r'kb_sub/kb_images/kb_icon2.ico')
        except TclError:
            pass
    if sys.platform == "win32" and projvar.platform == "winapp":
        try:
            root.iconbitmap(os.getcwd() + "\\" + "kb_icon2.ico")
        except TclError:
            pass
    if sys.platform == "darwin" and projvar.platform == "py":
        try:
            root.iconbitmap('kb_sub/kb_images/kb_icon1.icns')
        except TclError:
            pass
    if sys.platform == "darwin" and projvar.platform == "macapp":
        try:
            path = os.path.join(os.path.sep, 'Applications', 'klusterbox.app', 'Contents', 'Resources', 'kb_icon2.jpg')
            root.iconphoto(False, PhotoImage(file=path))
        except TclError:
            pass
    if sys.platform == "linux":
        try:
            img = PhotoImage(file='kb_sub/kb_images/kb_icon2.gif')
            # root.tk.call('wm', 'iconphoto', root._w, img)
            root.tk.call('wm', 'iconphoto', root.w, img)
        except TclError:
            pass


def dt_converter(string):  # converts a string of a datetime to an actual datetime
    dt = datetime.strptime(string, '%Y-%m-%d %H:%M:%S')
    return dt


def macadj(win, mac):  # switch between variables depending on platform
    if sys.platform == "darwin":
        arg = mac
    else:
        arg = win
    return arg


class MakeWindow:
    def __init__(self):
        self.topframe = Frame(projvar.root)
        self.s = Scrollbar(self.topframe)
        self.c = Canvas(self.topframe, width=1600)
        self.body = Frame(self.c)
        self.buttons = Canvas(self.topframe)  # button bar

    def create(self, frame):
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

    def finish(self):  # This closes the window created by front_window()
        projvar.root.update()
        self.c.config(scrollregion=self.c.bbox("all"))
        try:
            mainloop()
        except KeyboardInterrupt:
            projvar.root.destroy()

    def fill(self, last, count):  # fill bottom of screen to for scrolling.
        for i in range(count):
            Label(self.body, text="").grid(row=last + i)
        Label(self.body, text="kb", fg="lightgrey", anchor="w").grid(row=last + count + 1, sticky="w")


def front_window(frame):  # Sets up a tkinter page with buttons on the bottom
    if frame != "none":
        frame.destroy()  # close out the previous frame
    f = Frame(projvar.root)  # create new frame
    f.pack(fill=BOTH, side=LEFT)
    buttons = Canvas(f)  # button bar
    buttons.pack(fill=BOTH, side=BOTTOM)
    # link up the canvas and scrollbar
    s = Scrollbar(f)
    c = Canvas(f, width=1600)
    s.pack(side=RIGHT, fill=BOTH)
    c.pack(side=LEFT, fill=BOTH)
    s.configure(command=c.yview, orient="vertical")
    c.configure(yscrollcommand=s.set)
    # link the mousewheel - implementation varies by platform
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
    return f, s, c, ff, buttons
    # page contents - then call rear_window(wd)


def rear_window(wd):  # This closes the window created by front_window()
    projvar.root.update()
    wd[2].config(scrollregion=wd[2].bbox("all"))
    try:
        mainloop()
    except KeyboardInterrupt:
        projvar.root.destroy()


def set_globals(s_year, s_mo, s_day, i_range, station, frame):
    projvar.invran_weekly_span = i_range
    if station == "undefined":
        messagebox.showerror("Investigation station setting",
                             'Please select a station.',
                             parent=frame)
        return
    # error check for valid date
    # date = ""  # reference before assignment
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
    # if frame != "None":
    #     MainFrame().start(frame=frame)


class CarrierRecSet:
    def __init__(self, carrier, start, end, station):
        self.carrier = carrier
        self.start = start
        self.end = end
        self.station = station
        if self.start == self.end:
            self.range = "day"
        else:
            self.range = "week"

    def get(self):  # returns carrier records for one day or a week
        if self.range == "day":
            sql = "SELECT MAX(effective_date), carrier_name, list_status, ns_day, route_s, station " \
                  "FROM carriers WHERE carrier_name = '%s' and effective_date <= '%s' " \
                  % (self.carrier, self.start)
            daily_rec = inquire(sql)
            rec_set = []
            if daily_rec[0][5] == self.station:
                rec_set.append(daily_rec[0])  # since all weekly rings are in a rec_set, be consistant
                return rec_set
            else:
                return
        else:  # if the investigation range is weekly
            # get all records in the investigation range - the entire service week
            sql = "SELECT * FROM carriers WHERE carrier_name = '%s' and effective_date BETWEEN '%s' AND '%s' " \
                  "ORDER BY effective_date DESC" \
                  % (self.carrier, self.start, self.end)
            rec = inquire(sql)
            # get the relevant previous record (RPR)
            sql = "SELECT MAX(effective_date), carrier_name, list_status, ns_day, route_s, station " \
                  "FROM carriers WHERE carrier_name = '%s' and effective_date <= '%s' " \
                  "ORDER BY effective_date DESC" \
                  % (self.carrier, self.start - timedelta(days=1))
            before_range = inquire(sql)
            #  append before_range if there is no record for saturday or invest range is daily
            add_it = True  # indicates that the RPR needs to be added to carrier records
            if rec:
                for r in rec:  # loop through all records in carrier records
                    if r[0] == str(self.start):  # if a record is for the saturday in range
                        add_it = False  # do not add the RPR.
            if add_it:  # add the RPR if there is not sat range record
                if before_range[0] != (None, None, None, None, None, None):
                    rec.append(before_range[0])
            #  filter out record sets with no station matches
            station_anchor = False
            for r in rec:  # loop through all carrier records
                if r[5] == self.station:  # check if at least one record matchs to the station
                    station_anchor = True  # indicates that the carrecs are good for the current station.
            rec_set = []  # initialize array to put record sets into carrier list
            #  filter out any consecutive duplicate records
            if station_anchor:
                last_rec = ["xxx", "xxx", "xxx", "xxx", "xxx", "xxx"]
                for r in reversed(rec):
                    if r[2] != last_rec[2] or r[3] != last_rec[3] or r[4] != last_rec[4] or r[5] != last_rec[5]:
                        last_rec = r
                        rec_set.insert(0, r)  # add to the front of the list
                return rec_set
            else:
                return  # will return None - NoneType


class CarrierList:
    def __init__(self, start, end, station):
        self.start = start
        self.end = end
        self.station = station

    def get(self):  # get a weekly or daily carrier list
        c_list = []
        sql = "SELECT DISTINCT carrier_name FROM carriers WHERE station = '%s' AND effective_date <= '%s' " \
              "ORDER BY carrier_name, effective_date desc" \
              % (self.station, self.end)
        distinct = inquire(sql)  # call function to access database
        for carrier in distinct:
            rec_set = CarrierRecSet(carrier[0], self.start, self.end, self.station).get()  # get rec set per carrier
            if rec_set is not None:  # only add rec sets if there is something there
                c_list.append(rec_set)
        return c_list

    def get_distinct(self):  # get a list of distinct carrier names
        sql = "SELECT DISTINCT carrier_name FROM carriers WHERE station = '%s' AND effective_date <= '%s' " \
              "ORDER BY carrier_name, effective_date desc" \
              % (self.station, self.end)
        return inquire(sql)  # call function to access database


def gen_carrier_list():
    sql = ""
    # generate in range carrier list
    if projvar.invran_weekly_span:  # select sql dependant on range
        sql = "SELECT effective_date, carrier_name, list_status, ns_day, route_s, station, rowid" \
              " FROM carriers WHERE effective_date <= '%s' " \
              "ORDER BY carrier_name, effective_date desc" % projvar.invran_date_week[6]
    if not projvar.invran_weekly_span:   # if investigation range is weekly
        sql = "SELECT effective_date, carrier_name,list_status, ns_day,route_s, station, rowid" \
              " FROM carriers WHERE effective_date <= '%s' " \
              "ORDER BY carrier_name, effective_date desc" % projvar.invran_date
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
                # if record falls in investigation range - add it to more rows array
                if projvar.invran_weekly_span:  # if investigation range is weekly
                    if str(projvar.invran_date_week[1]) <= record[0] <= str(projvar.invran_date_week[6]):
                        more_rows.append(record)
                    if record[0] <= str(projvar.invran_date_week[0]) and len(pre_invest) == 0:
                        pre_invest.append(record)
                if not projvar.invran_weekly_span:  # if investigation range is daily...
                    # if date match and no pre_investigation
                    if record[0] <= str(projvar.invran_date) and len(pre_invest) == 0:
                        pre_invest.append(record)  # add rec to pre_invest array
            # find carriers who start in the middle of the investigation range CATEGORY ONE
            if len(more_rows) > 0 and len(pre_invest) == 0:
                station_anchor = "no"
                for each in more_rows:  # check if any records place the carrier in the selected station
                    if each[5] == projvar.invran_station:
                        station_anchor = "yes"  # if so, set the station anchor
                if station_anchor == "yes":
                    list(more_rows)
                    for each in more_rows:
                        x = list(each)  # convert the tuple to a list
                        carrier_list.append(x)  # add it to the list
            # find carriers with records before and during the investigation range CATEGORY TWO
            if len(more_rows) > 0 and len(pre_invest) > 0:
                station_anchor = "no"
                for each in more_rows + pre_invest:
                    if each[5] == projvar.invran_station:
                        station_anchor = "yes"
                if station_anchor == "yes":
                    xx = list(pre_invest[0])
                    carrier_list.append(xx)
            # find carrier with records from only before investigation range.CATEGORY THREE
            if len(more_rows) == 0 and len(pre_invest) == 1:
                for each in pre_invest:
                    if each[5] == projvar.invran_station:
                        x = list(pre_invest[0])
                        carrier_list.append(x)
            del more_rows[:]
            del pre_invest[:]
            del candidates[:]
    return carrier_list


class NsDayDict:
    def __init__(self, date):
        self.date = date  # is a datetime object
        self.pat = ("blue", "green", "brown", "red", "black", "yellow")  # define color sequence tuple

    def get_sat_range(self):  # calculate the n/s day of sat/first day of investigation range
        sat_range = self.date  # saturday range, first day of the investigation range
        wkdy_name = self.date.strftime("%a")
        while wkdy_name != "Sat":  # while date enter is not a saturday
            sat_range -= timedelta(days=1)  # walk back the date until it is a saturday
            wkdy_name = sat_range.strftime("%a")
        return sat_range

    def get(self):  # Dictionary of NS days
        sat_range = self.get_sat_range()  # calculate the n/s day of sat/first day of investigation range
        end_date = sat_range + timedelta(days=-1)
        cdate = datetime(2017, 1, 7)  #
        x = 0  # x is the index of self.pattern
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
        ns_xlate = {}  # ns translate dictionary
        for i in range(7):
            if i == 0:
                ns_xlate[self.pat[x]] = date.strftime("%a")
                date += timedelta(days=1)
            elif i == 1:
                date += timedelta(days=1)
                if x > 4:
                    x = 0
                else:
                    x += 1
            else:
                ns_xlate[self.pat[x]] = date.strftime("%a")
                date += timedelta(days=1)
                if x > 4:
                    x = 0
                else:
                    x += 1
        ns_xlate["none"] = "  "  # if there is no ns day, such as auxiliary assistance
        ns_xlate["sat"] = "Sat"  # if there are fixed ns days
        ns_xlate["mon"] = "Mon"
        ns_xlate["tue"] = "Tue"
        ns_xlate["wed"] = "Wed"
        ns_xlate["thu"] = "Thu"
        ns_xlate["fri"] = "Fri"
        return ns_xlate

    def ssn_ns(self, rotation):  # SpreadSheet Notation NS Day dictionary
        ssn_ns_code = {}
        # rotation is boolean -
        dic = self.get()
        if rotation:
            for p in self.pat:  # if rotation is True, annotate fixed ns days
                ssn_ns_code[p] = dic[p]
            ssn_ns_code["none"] = "  "  # if there is no ns day, such as auxiliary assistance
            ssn_ns_code["sat"] = "fSat"  # if there are fixed ns days
            ssn_ns_code["mon"] = "fMon"
            ssn_ns_code["tue"] = "fTue"
            ssn_ns_code["wed"] = "fWed"
            ssn_ns_code["thu"] = "fThu"
            ssn_ns_code["fri"] = "fFri"
        else:
            for p in self.pat:  # if rotation is false, annotate rotating nsdays
                ssn_ns_code[p] = "r{}".format(dic[p])
            ssn_ns_code["none"] = "  "  # if there is no ns day, such as auxiliary assistance
            ssn_ns_code["sat"] = "Sat"  # if there are fixed ns days
            ssn_ns_code["mon"] = "Mon"
            ssn_ns_code["tue"] = "Tue"
            ssn_ns_code["wed"] = "Wed"
            ssn_ns_code["thu"] = "Thu"
            ssn_ns_code["fri"] = "Fri"
        return ssn_ns_code

    def get_rev(self, rotation):  # Dictionary NS days - keys/values reversed
        dic = self.get()
        rev_rotate_dic = {}
        rev_fixed_dic = {}
        for (key, value) in dic.items():
            if key in self.pat:
                rev_rotate_dic[value.lower()] = key
            else:
                rev_fixed_dic[value.lower()] = key
        if rotation:
            return rev_rotate_dic
        else:
            return rev_fixed_dic

    @staticmethod
    def custom_config():  # shows custom ns day configurations for  printout / reports
        sql = "SELECT * FROM ns_configuration"
        ns_results = inquire(sql)
        custom_ns_dict = {}  # build dictionary for ns days
        days = ("sat", "mon", "tue", "wed", "thu", "fri")
        for r in ns_results:  # build dictionary for rotating ns days
            custom_ns_dict[r[0]] = "rotating: " + r[2]
        for d in days:  # expand dictionary for fixed days
            custom_ns_dict[d] = "fixed: " + d
        custom_ns_dict["none"] = "none"  # add "none" to dictionary
        return custom_ns_dict

    @staticmethod
    def get_custom_nsday():  # get ns day color configurations from dbase and make dictionary
        sql = "SELECT * FROM ns_configuration"
        ns_results = inquire(sql)
        ns_dict = {}  # build dictionary for ns days
        days = ("sat", "mon", "tue", "wed", "thu", "fri")
        for r in ns_results:  # build dictionary for rotating ns days
            ns_dict[r[0]] = r[2]  # build dictionary for ns fill colors
        for d in days:  # expand dictionary for fixed days
            ns_dict[d] = "fixed: " + d
        ns_dict["none"] = "none"  # add "none" to dictionary
        return ns_dict

    @staticmethod
    def gen_rev_ns_dict():  # creates full day/color ns day dictionary
        days = ("Saturday", "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday")
        color_pat = ("blue", "green", "brown", "red", "black", "yellow")
        code_ns = {}
        for d in days:
            for c in color_pat:
                if d[:3] == projvar.ns_code[c]:
                    code_ns[d] = c
        code_ns["None"] = "none"
        return code_ns


def dir_path(dirr):  # create needed directories if they don't exist and return the appropriate path
    path = ""
    if sys.platform == "darwin":
        if projvar.platform == "macapp":
            if not os.path.isdir(os.path.expanduser("~") + '/Documents'):
                os.makedirs(os.path.expanduser("~") + '/Documents')
            if not os.path.isdir(os.path.expanduser("~") + '/Documents/klusterbox'):
                os.makedirs(os.path.expanduser("~") + '/Documents/klusterbox')
            if not os.path.isdir(os.path.expanduser("~") + '/Documents/klusterbox/' + dirr):
                os.makedirs(os.path.expanduser("~") + '/Documents/klusterbox/' + dirr)
            path = os.path.expanduser("~") + '/Documents/klusterbox/' + dirr + '/'
        else:
            if not os.path.isdir('kb_sub/' + dirr):
                os.makedirs(('kb_sub/' + dirr))
            path = 'kb_sub/' + dirr + '/'
    if sys.platform == "win32":
        if projvar.platform == "winapp":
            if not os.path.isdir(os.path.expanduser("~") + '\\Documents'):
                os.makedirs(os.path.expanduser("~") + '\\Documents')
            if not os.path.isdir(os.path.expanduser("~") + '\\Documents\\klusterbox'):
                os.makedirs(os.path.expanduser("~") + '\\Documents\\klusterbox')
            if not os.path.isdir(os.path.expanduser("~") + '\\Documents\\klusterbox\\' + dirr):
                os.makedirs(os.path.expanduser("~") + '\\Documents\\klusterbox\\' + dirr)
            path = os.path.expanduser("~") + '\\Documents\\klusterbox\\' + dirr + '\\'
        else:
            if not os.path.isdir('kb_sub\\' + dirr):
                os.makedirs(('kb_sub\\' + dirr))
            path = 'kb_sub\\' + dirr + '\\'
    return path


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


def gen_ns_dict(file_path, to_addname):  # creates a dictionary of ns days
    days = ("Saturday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday")
    mv_codes = ("BT", "MV", "ET")
    good_jobs = ("134", "844", "434")
    results = []
    carrier = []
    id_bank = []
    aux_list = []
    for ident in to_addname:
        id_bank.append(ident[0].zfill(8))
        if ident[3] in ("auxiliary", "part time flex"):
            aux_list.append(ident[0].zfill(8))  # make an array of auxiliary carrier emp ids
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
                    good_id = line[4].zfill(8)  # set trigger to ident of carriers who are FT or aux carriers
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


class BuildPath:
    def __init__(self):
        self.delimiter = ""
        self.newpath = ""

    def get_delimiter(self):
        if sys.platform == "darwin":
            self.delimiter = "/"
        else:
            self.delimiter = "\\"

    def build(self, path_array):
        """
        Takes an array and of directories and a file and converts it into a path suitable for the operating system.
        """
        self.get_delimiter()
        for i in range(len(path_array)):
            self.newpath += path_array[i]
            if i < len(path_array)-1:
                self.newpath += self.delimiter
        return self.newpath


class SaturdayInRange:  # recieves a datetime object
    def __init__(self, dt):
        self.dt = dt

    def get(self):  # returns the sat range
        wkdy_name = self.dt.strftime("%a")
        while wkdy_name != "Sat":  # while date enter is not a saturday
            self.dt -= timedelta(days=1)  # walk back the date until it is a saturday
            wkdy_name = self.dt.strftime("%a")
        return self.dt


class ReportName:  # returns a file name which is stamped with the datetime
    def __init__(self, filename):
        self.filename = filename

    def create(self):
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")  # create a file name
        return self.filename + "_" + stamp + ".txt"


class Convert:
    def __init__(self, data):
        self.data = data

    def datetime_separation(self):  # converts a datetime object into an array with year, month and day
        year = self.data.strftime("%Y")
        month = self.data.strftime("%m")
        day = self.data.strftime("%d")
        date = [year, month, day]
        return date

    def str_to_bool(self):  # change a string into a boolean variable type
        if self.data == 'True':
            return True
        return False

    def bool_to_onoff(self):  # takes a boolean and returns on for true, off for false
        if int(self.data):
            return "on"
        return "off"

    def strbool_to_onoff(self):  # take a boolean in the form of a string and returns "on" or "off"
        if self.data == "True":
            return "on"
        return "off"

    def onoff_to_bool(self):
        if self.data == "on":
            return True
        return False

    def backslashdate_to_datetime(self):  # convert a date with backslashes into a datetime
        date = self.data.split("/")
        string = date[2] + "-" + date[0] + "-" + date[1] + " 00:00:00"
        return dt_converter(string)

    def array_to_string(self):  # make an array into a string (with commas)
        string = ""
        for i in range(len(self.data)):
            string += self.data[i]
            if i != len(self.data) - 1:
                string += ","
        return string

    def string_to_array(self):  # make string into array, remove whitespace
        new_array = []
        array = self.data.split(",")
        for a in array:
            a = a.strip()
            new_array.append(a)
        return new_array

    # takes day (eg "mon","wed") and converts to datetime. needs saturday in range
    def day_to_datetime_str(self, sat_range):
        if self.data == sat_range.strftime("%a").lower():  # saturday
            return str(sat_range)
        sat_range += timedelta(days=1)
        if self.data == sat_range.strftime("%a").lower():  # sunday
            return str(sat_range)
        sat_range += timedelta(days=1)
        if self.data == sat_range.strftime("%a").lower():  # monday
            return str(sat_range)
        sat_range += timedelta(days=1)
        if self.data == sat_range.strftime("%a").lower():  # tueday
            return str(sat_range)
        sat_range += timedelta(days=1)
        if self.data == sat_range.strftime("%a").lower():  # wednesday
            return str(sat_range)
        sat_range += timedelta(days=1)
        if self.data == sat_range.strftime("%a").lower():  # thursday
            return str(sat_range)
        sat_range += timedelta(days=1)
        if self.data == sat_range.strftime("%a").lower():  # friday
            return str(sat_range)

    def dt_converter(self):  # converts a string of a datetime to an actual datetime
        dt = datetime.strptime(self.data, '%Y-%m-%d %H:%M:%S')
        return dt

    def empty_not_zero(self):  # returns an empty string for any value equal to zero
        if self.data == "0":
            return ""
        if self.data == "0.0":
            return ""
        return self.data

    def empty_not_zerofloat(self):
        if self.data == 0.0:
            return ""
        if self.data == 0:
            return ""
        return self.data

    def str_to_floatoremptystr(self):
        if self.data == "*":
            return "*"
        if self.data == "":
            return ""
        if self.data == "0.0":
            return ""
        if self.data == "0":
            return ""
        if isfloat(self.data):
            return float(self.data)
        return self.data

    def none_not_empty(self):  # returns none instead of empty string for option menus
        if self.data == "":
            return "none"
        return self.data

    def empty_not_none(self):  # returns empty string instead of "none" for spreadsheets
        if self.data == "none":
            return ""
        return self.data

    def hundredths(self):  # returns a number (as a string) into a number with 2 decimal places
        number = float(self.data)  # convert the number to a float
        return "{:.2f}".format(number)  # return the number as a string with 2 decimal places

    def zero_or_hundredths(self):
        try:
            if float(self.data) == 0:
                number = 0.00  # convert the number to a float
                return "{:.2f}".format(number)  # return the number as a string with 2 decimal places
            else:
                number = float(self.data)  # convert the number to a float
                return "{:.2f}".format(number)  # return the number as a string with 2 decimal places
        except ValueError:
            number = 0.00  # convert the number to a float
            return "{:.2f}".format(number)  # return the number as a string with 2 decimal places

    def empty_or_hunredths(self):
        if self.data.strip() in ("0", "0.0", "0.00", ".0", ".00", ".", ""):
            return ""
        else:
            number = float(self.data)  # convert the number to a float
            return "{:.2f}".format(number)  # return the number as a string with 2 decimal places

    def zero_not_empty(self):
        if self.data == "":
            return 0
        return self.data


class Handler:
    def __init__(self, data):
        self.data = data

    def nonetype(self):
        if self.data is None:
            return str("")
        else:
            self.data = str(self.data)
            self.data = self.data.strip()
            return self.data

    def ns_nonetype(self):
        if self.data is None:
            return str("  ")
        else:
            return self.data

    def nsblank2none(self):
        if self.data.strip() == "":
            return str("none")
        else:
            return self.data

    def plurals(self):  # put an "s" on the end of words to make them plural
        if self.data == 1:
            return ""
        else:
            return "s"

    def format_str_as_int(self):
        num = int(self.data)
        return str(num)

    def str_to_int_or_str(self):
        if self.data.isnumeric():
            return int(self.data)
        else:
            return self.data

    def str_to_float_or_str(self):
        try:
            self.data = float(self.data)
            return self.data
        except ValueError:
            return self.data

    @staticmethod
    def route_adj(route):  # convert five digit route numbers to four when the route number > 99
        if len(route) == 5:  # when the route number is five digits
            if route[2] == "0":  # and the third digit is a zero
                return route[0] + route[1] + route[3] + route[4]  # rewrite the string, deleting the third digit
            else:
                return route  # if the route number is > 99, return it without change
        if len(route) == 4:
            return route  # if the route number is 4 digits, return it without change

    def routes_adj(self):  # only allow five digit route numbers in chains where route number > 99
        if self.data.strip() == "":
            return ""  # return empty strings with an empty string
        routes = self.data.split("/")  # convert andy chains into an array
        new_array = []
        for r in routes:
            new_array.append(self.route_adj(r))
        separator = "/"  # convert the array into a string
        return separator.join(new_array)  # and return


def isfloat(value):
    try:
        float(value)
        return True
    except ValueError:
        return False


def isint(value):  # checks if the argument is an integer
    try:
        int(value)
        return True
    except ValueError:
        return False


def dir_path_check(dirr):  # return appropriate path depending on platorm
    path = ""
    if sys.platform == "darwin":
        if projvar.platform == "macapp":
            path = os.path.expanduser("~") + '/Documents/klusterbox/' + dirr
        else:
            path = 'kb_sub/' + dirr
    if sys.platform == "win32":
        if projvar.platform == "winapp":
            path = os.path.expanduser("~") + '\\Documents\\klusterbox\\' + dirr
        else:
            path = 'kb_sub\\' + dirr
    return path


def dir_filedialog():
    # determine where the file dialog opens
    if sys.platform == "darwin":
        path = os.path.join(os.path.sep, os.path.expanduser("~"), 'Documents')
    elif sys.platform == "win32":
        path = os.path.expanduser("~") + '\\Documents'
    else:
        path = os.path.expanduser("~")
    return path


class Quarter:
    def __init__(self, data):
        self.data = data

    def find(self):  # pass month in (as number) as argument - quarter is returned
        if int(self.data) in (1, 2, 3):
            return 1
        if int(self.data) in (4, 5, 6):
            return 2
        if int(self.data) in (7, 8, 9):
            return 3
        if int(self.data) in (10, 11, 12):
            return 4


class Rings:
    def __init__(self, name, date):
        self.name = name
        self.date = date  # provide any date in investigation range
        self.ring_recs = []  # put all results in an array

    def get(self, day):
        sql = "SELECT * FROM rings3 WHERE carrier_name = '%s' and rings_date = '%s'" % (self.name, day)
        return inquire(sql)

    def get_for_day(self):
        ring = self.get(self.date)
        if not ring:  # if the results are empty
            self.ring_recs.append(ring)  # return empty list
        else:  # if results are not empty
            self.ring_recs.append(ring[0])  # return first result of list
        return self.ring_recs

    def get_for_week(self):
        sat_range = SaturdayInRange(self.date).get()
        for i in range(7):
            ring = self.get(sat_range)
            if not ring:  # if the results are empty
                self.ring_recs.append(ring)  # return empty list
            else:  # if results are not empty
                self.ring_recs.append(ring[0])  # return first result of list
            sat_range += timedelta(days=1)
        return self.ring_recs


class Moves:
    def __init__(self):
        self.moves = None
        self.timeoff = 0

    def checkempty(self):
        if not self.moves:  # fail if self moves is empty
            return False
        return True

    def checklenght(self):
        if len(self.moves) % 3 == 0:  # fail if number of elements are not a multiple of 3
            return True
        return False

    def checksforzero(self):
        if self.timeoff <= 0:  # return empty string if result is 0
            return ""
        self.timeoff = round(self.timeoff, 2)  # round the time off route to 2 decimal places
        return str(self.timeoff)

    def timeoffroute(self, moves):  # gives the time off route given a moves set
        self.moves = moves
        if not self.checkempty():  # check if len(moves) is multiple of 3 and not 0.
            return ""
        self.moves = Convert(self.moves).string_to_array()
        if not self.checklenght():
            return ""
        self.timeoff = 0
        for i in range(0, len(self.moves), 3):
            self.timeoff += (float(self.moves[i+1]) - float(self.moves[i]))
        return self.checksforzero()


class Overtime:
    def __init__(self):
        self.total = None  # total hours worked or daily 5200 time
        self.timeoff = None  # this is the ot off route calcuated from the moves in Moves() class
        self.code = None  # looking for ns day
        self.overtime = None

    def check_empty_total(self):
        if not self.total:
            return False
        return True

    def check_total(self):
        if self.code != "ns day":
            if float(self.total) <= 8.00:
                return False
        else:
            if float(self.total) <= 0:
                return False
        return True

    def checks(self):
        if not self.check_empty_total():
            return False
        if not self.check_total():
            return False
        return True

    def check_empty_timeoff(self):  # if there was no time worked off route, return False
        if not self.timeoff:
            return False
        return True

    def straight_overtime(self, total, code):
        self.total = total
        self.code = code
        if not self.checks():
            return ""
        if self.code != "ns day":  # if it is not the ns day,
            total = float(self.total) - 8  # calculate overtime by subtracting the 8 hour day.
            return format(total, '.2f')  # return a formated string
        return self.total  # if it is the ns day, the overtime is all hours worked that day.

    def proper_overtime(self, total, timeoff, code):
        self.total = total
        self.timeoff = timeoff
        self.code = code
        if not self.checks():  # if the total is empty or less than 8.00 - return empty string
            return ""
        if self.code == "ns day":  # if it is the ns day, the overtime is all hours worked that day.
            return self.total
        overtime = float(self.total) - 8
        if not self.check_empty_timeoff():
            return format(overtime, '.2f')
        offroute = min(overtime, float(self.timeoff))
        return format(offroute, '.2f')


class SpeedSettings:
    def __init__(self):
        sql = "SELECT tolerance FROM tolerances"
        results = inquire(sql)  # get spreadsheet settings from database
        self.abc_breakdown = Convert(results[15][0]).str_to_bool()
        self.min_empid = int(results[16][0])
        self.min_alpha = int(results[17][0])
        self.min_abc = int(results[18][0])
        self.speedcell_ns_rotate_mode = Convert(results[19][0]).str_to_bool()


class NameChecker:
    def __init__(self, name):
        self.name = name.lower()

    def check_characters(self):  # checks if characters in name are in approved tuple
        for char in self.name:
            if char in ("a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q",
                        "r", "s", "t", "u", "v", "w", "x", "y", "z", " ", "-", "'", ".", ","):
                pass
            else:
                return False
        return True

    def check_length(self):  # checks that the name is not too long
        if len(self.name) < 29:
            return True
        else:
            return False

    def check_comma(self):  # checks if there is a comma in the name
        s_name = self.name.split(",")
        if len(s_name) == 2:
            return True
        else:
            return False

    def check_initial(self):  # checks if theres is an initial in the variable
        s_name = self.name.split(",")
        if len(s_name) > 1:
            if len(s_name[1].strip()) == 1:
                return True
        else:
            return False
    """
        def add_comma_spacing(self):  # make sure there is a space between the comma and the first initial
        s_name = self.name.split(",")
        if s_name[1][0] != " ":
            return s_name[0] + ", " + s_name[1]
        return self.name
    """


class RouteChecker:
    def __init__(self, route):
        self.route = route
        self.routearray = self.route.split("/")

    def is_empty(self):
        if self.route == "":
            return True
        return False

    def check_numeric(self):  # is the route numeric?
        if self.route == "":
            return True
        for r in self.routearray:
            if not r.isnumeric():
                return False
        return True

    def check_array(self):  # are there 1 or 5 items in the route string
        if len(self.routearray) == 1:
            return True
        elif len(self.routearray) == 5:
            return True
        else:
            return False

    def check_length(self):  # are the routes 4 or 5 digits long
        if self.route == "":
            return True
        for r in self.routearray:
            if len(r) < 4 or len(r) > 5:
                return False
        return True

    def only_one(self):  # returns False if there is more than one route given
        if len(self.routearray) > 1:
            return False
        return True

    def only_numbers(self):  # returns True if variable is empty string or contains only numbers
        if self.route == "":
            return True
        try:
            self.route = int(self.route)
        except ValueError:
            return False

    def check_all(self):  # do all checks, return False if any fail.
        if not self.check_numeric():
            return False
        if not self.check_array():
            return False
        if not self.check_length():
            return False
        return True


class RingTimeChecker:
    def __init__(self, ring):
        self.ring = ring

    def make_float(self):  # converts self.ring to a floatType or returns False
        try:
            self.ring = float(self.ring)
            return self.ring
        except ValueError:
            return False

    def check_for_zeros(self):  # returns True if zero or empty
        try:
            if float(self.ring) == 0:
                return True
        except ValueError:
            pass
        if self.ring == "":
            return True
        return False

    def check_numeric(self):  # is the ring numeric?
        try:
            float(self.ring)
            return True
        except ValueError:
            return False

    def over_24(self):  # is the time greater than 24 hours
        if float(self.ring) > 24:
            return False
        return True

    def over_8(self):  # is the time greater than 8 hours
        if float(self.ring) > 8:
            return False
        return True

    def over_5000(self):  # if the time is greater than 5000 hours - upper limit on make ups for OT Equitability
        if float(self.ring) > 5000:
            return False
        return True

    def less_than_zero(self):  # disappear here
        if float(self.ring) < 0:
            return False
        return True

    def count_decimals_place(self):  # limit time to two decimal places
        return round(float(self.ring), 2) == float(self.ring)  # returns False if self.ring has > two decimal places


class MovesChecker:
    def __init__(self, moves):
        self.moves = moves

    def length(self):  # return False if not a multiple of three
        return len(self.moves) % 3 == 0

    def check_for_zeros(self):  # returns True if zero or empty
        if self.moves == "":
            return True
        try:
            if float(self.moves) == 0:
                return True
        except ValueError:
            pass
        return False

    def compare(self, second):  # return False if first move is greater than second move
        if float(self.moves) > float(second):
            return False
        return True


class MinrowsChecker:
    def __init__(self, data):
        self.data = data

    def is_empty(self):  # is the data an empty string?
        if self.data == "":
            return True
        return False

    def is_numeric(self):  # is the data a number?
        try:
            self.data = float(self.data)
            return True
        except ValueError:
            return False

    def no_decimals(self):  # does the data have no decimal places?
        if "." in self.data:
            return False
        return True

    def not_negative(self):  # is the data not a negative?
        if "-" in self.data:
            return False
        return True

    def not_zero(self):
        if float(self.data) == 0:
            return False
        return True

    def within_limit(self, limit):  # is the data not exceed a given limit?
        if int(self.data) <= limit:
            return True
        return False


class RefusalTypeChecker:
    def __init__(self, data):
        self.data = data

    def is_empty(self):
        if not self.data:
            return True
        return False

    def is_one(self):
        if len(self.data) > 1:
            return False
        return True

    def is_letter(self):
        if not self.data.isalpha():
            return False
        return True


class BackSlashDateChecker:
    """
    Create a backslashdate object by calling the class.
    Next use count_backslashes next to make sure the date can be broken down into 3 parts.
    Next call breakdown method to fully create the backslashdate object.
    Next call the following methods using the backslashdate instance to ensure the date is correctly fomatted.
    """
    def __init__(self, data):
        self.data = data
        self.breakdown = []  # break down the date into an array of 3 items
        self.month = ""
        self.day = ""
        self.year = ""

    def count_backslashes(self):
        if self.data.count("/") != 2:
            return False
        return True

    def breaker(self):  # this will fully create the instance of the object
        self.breakdown = self.data.split("/")
        self.month = self.breakdown[0]
        self.day = self.breakdown[1]
        self.year = self.breakdown[2]

    def check_numeric(self):  # check each element in the date to ensure they are numeric
        for date_element in self.breakdown:
            if not isint(date_element):
                return False
        return True

    def check_minimums(self):  # check each element in the date to ensure they are greater than zero
        for date_element in self.breakdown:
            if int(date_element) <= 0:
                return False
        return True

    def check_month(self):
        if int(self.month) > 12:
            return False
        return True

    def check_day(self):
        if int(self.day) > 31:
            return False
        return True

    def check_year(self):
        if len(self.year) != 4:
            return False
        return True

    def valid_date(self):
        try:
            datetime(int(self.year), int(self.month), int(self.day))
            return True
        except ValueError:
            return False


class CarrierRecFilter:  # accepts carrier records from CarrierList().get()
    def __init__(self, recset, startdate):
        self.recset = []  # initialize vars as empty for new carriers
        self.startdate = ""
        self.carrier = ""
        self.nsday = ""
        self.route = ""
        self.station = ""
        lastrec = None
        if recset is not None:  # handle carriers who are not new carriers
            if len(recset) != 0:  # new carriers can appear as NoneType or an empty list
                self.recset = recset
                for r in reversed(recset):  # get the earliest record in the recset. use reversed()
                    lastrec = r
                    break
                self.startdate = startdate
                self.date = lastrec[0]
                self.carrier = lastrec[1]
                self.nsday = lastrec[3]
                self.route = lastrec[4]
                self.station = lastrec[5]

    def filter_nonlist_recs(self):  # filters out any records were the list status hasn't changed.
        filtered_set = []
        last_rec = ["xxx", "xxx", "xxx", "xxx", "xxx", "xxx"]
        for r in reversed(self.recset):
            if r[2] != last_rec[2]:
                last_rec = r
                filtered_set.insert(0, r)  # add to the front of the list
        return filtered_set

    def condense_recs(self, ns_rotate_mode):  # condense multiple recs into format used by speedsheets
        ns_dic = NsDayDict(self.startdate).ssn_ns(ns_rotate_mode)  # get speedsheet notation for nsdays
        date_str = ""
        list_str = ""
        i = 1
        for rec in reversed(self.recset):
            if i == 1:
                date_str = ""
            else:
                date_str = date_str + dt_converter(rec[0]).strftime('%a').lower()
            list_str = list_str + rec[2]
            if i != len(self.recset):
                if i == 1:
                    date_str = ""
                else:
                    date_str = date_str + ","
                list_str = list_str + ","
            i += 1
        ns = ns_dic[self.nsday]  # ns day is given with speedsheet notation for nsdays
        return date_str, self.carrier, list_str, ns, self.route, self.station

    def condense_recs_ns(self):  # condense multiple recs into format used by speedsheets
        date_str = ""
        list_str = ""
        i = 1
        for rec in reversed(self.recset):
            if i == 1:
                date_str = ""
            else:
                date_str = date_str + dt_converter(rec[0]).strftime('%a').lower()
            list_str = list_str + rec[2]
            if i != len(self.recset):
                if i == 1:
                    date_str = ""
                else:
                    date_str = date_str + ","
                list_str = list_str + ","
            i += 1
        return date_str, self.carrier, list_str, self.nsday, self.route, self.station

    def detect_outofstation(self, station):  # returns a rec with only the date and name
        record_set = []
        if Convert(self.date).dt_converter() > self.startdate:
            to_add = [self.startdate, self.carrier]  # out of station records only have one item
            record_set.append(to_add)
        for rec in reversed(self.recset):
            if rec[5] == station:
                record_set.append(rec)
            else:
                to_add = [str(self.startdate), self.carrier]  # out of station records only date and name
                record_set.append(to_add)
        return record_set


class PdfConverterFix:
    def __init__(self, data):
        self.data = data

    """
    pass the array of routes as data to the class, 
    pass the count (an integer) to the method. 
    method will add "000000" to the array until its length matches the count.
    """
    def route_filler(self, count):
        if len(self.data) < count:
            while len(self.data) < count:
                self.data.append("000000")
        return self.data


class ProgressBarDe:  # determinate Progress Bar
    def __init__(self, title="Klusterbox", label="working", text="Stand by..."):
        self.title = title
        self.label = label
        self.text = text
        self.pb_root = Tk()  # create a window for the progress bar
        self.pb_label = Label(self.pb_root, text=self.label)  # make label for progress bar
        self.pb = ttk.Progressbar(self.pb_root, length=400, mode="determinate")  # create progress bar
        self.pb_text = Label(self.pb_root, text=self.text, anchor="w")

    def delete(self):  # self destruct the progress bar object for keyerror exceptions
        self.pb_root.update_idletasks()
        self.pb_root.update()
        self.pb_root.destroy()
        del self.pb_root

    def max_count(self, maxx):  # set length of progress bar
        self.pb["maximum"] = maxx

    def start_up(self):
        titlebar_icon(self.pb_root)  # place icon in titlebar
        self.pb_root.title(self.title)
        self.pb_label.grid(row=0, column=0, sticky="w")
        self.pb.grid(row=1, column=0, sticky="w")
        self.pb_text.grid(row=2, column=0, sticky="w")

    def move_count(self, count):  # changes the count of the progress bar
        self.pb['value'] = count
        self.pb_root.update()

    def change_text(self, text):  # changes the text of the progress bar
        self.pb_text.config(text="{}".format(text))
        self.pb_root.update()
        # projvar.root.update()

    def stop(self):
        self.pb.stop()  # stop and destroy the progress bar
        self.pb_text.destroy()
        self.pb_label.destroy()  # destroy the label for the progress bar
        self.pb.destroy()
        self.pb_root.destroy()
