"""
This is a module in the klusterbox library. It creates a window where the user can enter in clock rings and leaves
information.
"""

import projvar
from kbtoolbox import commit, inquire, CarrierRecSet, Convert, Handler, macadj, MovesChecker, NewWindow, Rings, \
    RingTimeChecker, RouteChecker
# Standard Libraries
from tkinter import Button, Entry, OptionMenu, StringVar, Frame, Label, LEFT
from tkinter import messagebox


class EnterRings:
    """
    A Screen for entering in carrier clock rings
    """

    def __init__(self, carrier):
        self.win = NewWindow(title="Enter Clock Rings")
        self.frame = None
        self.carrier = carrier
        self.carrecs = []  # get the carrier rec set
        self.ringrecs = []  # get the rings for the week
        self.dates = []  # get a datetime object for each day in the investigation range
        self.daily_carrecs = []  # get the carrier record for each day
        self.daily_ringrecs = []  # get the rings record for each day
        self.onrec_totals = []  # ring records currently in the database
        self.onrec_begintour = []
        self.onrec_moves = []
        self.onrec_rss = []
        self.onrec_endtour = []
        self.onrec_codes = []
        self.onrec_lvtype = []
        self.onrec_lvtime = []
        self.totals = []  # arrays holding stringvars
        self.moves = []
        self.rss = []
        self.codes = []
        self.lvtypes = []
        self.lvtimes = []
        self.refusals = []
        self.begintour = []
        self.endtour = []
        self.now_moves = []  # the current elements of the onrec_moves array
        self.sat_mm = []  # holds daily stringvars for moves
        self.sun_mm = []
        self.mon_mm = []
        self.tue_mm = []
        self.wed_mm = []
        self.thu_mm = []
        self.fri_mm = []
        self.daily_mm = []  # holds daily stringvars for moves when invran is daily
        if projvar.invran_weekly_span:
            self.moves_array = [self.sat_mm, self.sun_mm, self.mon_mm, self.tue_mm, self.wed_mm, self.thu_mm,
                                self.fri_mm]
        else:
            self.moves_array = [self.daily_mm, ]
        self.daily_moves = []
        self.ot_rings_limiter = None
        self.tourrings = None  # True if user wants to display the BT (begin tour) and ET (end tour)
        self.chg_these = []
        self.addrings = []
        if projvar.invran_weekly_span:
            for _ in range(7):
                self.addrings.append([])
        else:
            self.addrings.append([])
        self.status_update = ""
        self.delete_report = 0
        self.update_report = 0
        self.insert_report = 0
        self.day = ("Saturday", "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday")
        self.days = ("sat", "sun", "mon", "tue", "wed", "thu", "fri")
        self.i_range = 0  # loop 7 times for week or once for day
        self.frame = []

    def start(self, frame=None):
        """ a master method for running the other methods in proper sequence """
        if frame:
            self.win.create(frame)
        else:
            self.win.create(None)
        self.re_initialize()  # initialize all variables
        self.get_carrecs()
        self.get_ringrecs()
        self.get_dates()
        self.get_i_range()
        self.get_daily_carrecs()
        self.get_daily_ringrecs()
        self.get_rings_limiter()
        self.get_tourrings()
        self.build_frames()
        self.get_onrecs()
        self.get_stringvars()
        self.set_stringvars()
        self.build_page()
        self.write_report()
        self.buttons_frame()
        self.zero_report_vars()
        self.win.finish()

    def re_initialize(self):
        """ a method for re initializing all variables after Apply is pressed or when first started. """
        self.carrecs = []  # get the carrier rec set
        self.ringrecs = []  # get the rings for the week
        self.dates = []  # get a datetime object for each day in the investigation range
        self.daily_carrecs = []  # get the carrier record for each day
        self.daily_ringrecs = []  # get the rings record for each day
        self.onrec_totals = []  # ring records currently in the database
        self.onrec_begintour = []
        self.onrec_moves = []
        self.onrec_rss = []
        self.onrec_endtour = []
        self.onrec_codes = []
        self.onrec_lvtype = []
        self.onrec_lvtime = []
        self.totals = []  # arrays holding stringvars
        self.moves = []
        self.rss = []
        self.codes = []
        self.lvtypes = []
        self.lvtimes = []
        self.refusals = []
        self.begintour = []
        self.endtour = []
        self.now_moves = []  # the current elements of the onrec_moves array
        self.sat_mm = []  # holds daily stringvars for moves
        self.sun_mm = []
        self.mon_mm = []
        self.tue_mm = []
        self.wed_mm = []
        self.thu_mm = []
        self.fri_mm = []
        self.daily_mm = []  # holds daily stringvars for moves when invran is daily
        if projvar.invran_weekly_span:
            self.moves_array = [self.sat_mm, self.sun_mm, self.mon_mm, self.tue_mm, self.wed_mm, self.thu_mm,
                                self.fri_mm]
        else:
            self.moves_array = [self.daily_mm, ]
        self.daily_moves = []
        self.chg_these = []
        self.addrings = []
        if projvar.invran_weekly_span:
            for _ in range(7):
                self.addrings.append([])
        else:
            self.addrings.append([])
        self.i_range = 0  # loop 7 times for week or once for day
        self.frame = []

    def get_carrecs(self):
        """ get the carrier's carrier rec set """
        if projvar.invran_weekly_span:  # get the records for the full service week
            self.carrecs = CarrierRecSet(self.carrier, projvar.invran_date_week[0], projvar.invran_date_week[6],
                                         projvar.invran_station).get()
        else:  # get the records for the day
            self.carrecs = CarrierRecSet(self.carrier, projvar.invran_date, projvar.invran_date,
                                         projvar.invran_station).get()

    def get_ringrecs(self):
        """ get the ring recs for the invran """
        if projvar.invran_weekly_span:  # get the records for the full service week
            self.ringrecs = Rings(self.carrier, projvar.invran_date).get_for_week()
        else:  # get the records for the day
            self.ringrecs = Rings(self.carrier, projvar.invran_date).get_for_day()

    def get_dates(self):
        """ get a datetime object for each day in the investigation range """
        if projvar.invran_weekly_span:
            self.dates = projvar.invran_date_week
        else:
            self.dates = [projvar.invran_date, ]
            
    def get_i_range(self):
        """ i range is a count of days in the investigation range. """
        if projvar.invran_weekly_span:  # if investigation range is weekly
            self.i_range = 7  # loop 7 times for week or once for day
        else:
            self.i_range = 1

    def get_daily_carrecs(self):
        """ make a list of carrier records for each day """
        for d in self.dates:
            for rec in self.carrecs:
                if rec[0] <= str(d):  # if the dates match
                    self.daily_carrecs.append(rec)  # append the record
                    break

    def get_daily_ringrecs(self):
        """ make list of ringrecs for each day, insert empty rec if there is no rec """
        match = False
        for d in self.dates:  # for each day in self.dates
            for rr in self.ringrecs:
                if rr:  # if there is a ring rec
                    if rr[0] == str(d):  # when the dates match
                        self.daily_ringrecs.append(list(rr))  # creates the daily_ringrecs array
                        match = True
            if not match:  # if there is no match
                add_this = [d, self.carrier, "", "", "none", "", "none", "", None, "", ""]  # insert an empty record
                self.daily_ringrecs.append(add_this)  # creates the daily_ringrecs array
            match = False
        # convert the time item from string to datetime object
        for i in range(len(self.daily_ringrecs)):
            if type(self.daily_ringrecs[i][0]) == str:
                self.daily_ringrecs[i][0] = Convert(self.daily_ringrecs[i][0]).dt_converter()

    def get_rings_limiter(self):
        """ get the status of rings limiter which limits the widgets in the screen """
        sql = "SELECT tolerance FROM tolerances WHERE category = '%s'" % "ot_rings_limiter"
        results = inquire(sql)
        self.ot_rings_limiter = int(results[0][0])

    def get_tourrings(self):
        """ get tourrings from database which allow user to show BT (begin tour) and ET (end tour) """
        sql = "SELECT tolerance FROM tolerances WHERE category = '%s'" % "tourrings"
        results = inquire(sql)
        self.tourrings = int(results[0][0])

    def get_widgetlist(self, i):
        """ returns a list with moves and/or tourrings or an empty list. """
        widgetlist = []
        if self.tourrings:
            widgetlist.append("tourrings")
        if self.daily_carrecs[i][2] in ("otdl",) and not self.ot_rings_limiter:
            widgetlist.append("moves")
        if self.daily_carrecs[i][2] in ("nl", "wal"):
            widgetlist.append("moves")
        return widgetlist

    @staticmethod
    def get_ww(widgetlist):
        """ get the widget of widgets in windows """
        if len(widgetlist) > 1:  # if there is more than one thing in the widgetlist
            return 6  # shorten the widget
        return 8
    
    def build_frames(self):
        """ build the frames that nest inside the body frame. There is one frame per day. """
        for _ in range(self.i_range):
            self.frame.append(Frame(self.win.body, width=500))

    def get_onrecs(self):
        """ get rings currently in the database. """
        for i in range(self.i_range):
            now_total = Convert(self.daily_ringrecs[i][2]).empty_not_zero()
            now_bt = Convert(self.daily_ringrecs[i][9]).empty_not_zero()
            now_rs = Convert(self.daily_ringrecs[i][3]).empty_not_zero()
            now_et = Convert(self.daily_ringrecs[i][10]).auto_not_zero()
            now_code = Convert(self.daily_ringrecs[i][4]).none_not_empty()
            now_moves = self.daily_ringrecs[i][5]
            now_lv_type = Convert(self.daily_ringrecs[i][6]).none_not_empty()
            now_lv_time = Convert(self.daily_ringrecs[i][7]).empty_not_zero()
            self.onrec_totals.append(now_total)
            self.onrec_begintour.append(now_bt)
            self.onrec_moves.append(now_moves)
            self.onrec_rss.append(now_rs)
            self.onrec_endtour.append(now_et)
            self.onrec_codes.append(now_code)
            self.onrec_lvtype.append(now_lv_type)
            self.onrec_lvtime.append(now_lv_time)

    def get_now_moves(self, i):
        """ convert the on record moves array from a string into an array. """
        self.now_moves = []
        if not self.onrec_moves[i]:
            return
        moves = self.onrec_moves[i].split(",")
        for ring in moves:
            self.now_moves.append(ring)
    
    def get_stringvars(self):
        """ make the stringvars for the page """
        for i in range(self.i_range):
            self.totals.append(StringVar(self.frame[i]))  # append stringvar to totals array
            self.begintour.append(StringVar(self.frame[i]))  # append stringvar to bt array
            self.rss.append(StringVar(self.frame[i]))  # RS entry widget
            self.endtour.append(StringVar(self.frame[i]))  # append stringvar to et array
            self.codes.append(StringVar(self.frame[i]))  # code entry widget
            self.lvtypes.append(StringVar(self.frame[i]))  # leave type entry widget
            self.lvtimes.append(StringVar(self.frame[i]))  # leave time entry widget
            self.refusals.append(None)  # refusals column is not used.

    def set_stringvars(self):
        """ set the values for the stringvars """
        for i in range(self.i_range):
            self.totals[i].set(self.onrec_totals[i])  # set the starting value for total
            self.begintour[i].set(self.onrec_begintour[i])  # set the starting value for BT
            self.rss[i].set(self.onrec_rss[i])  # set the starting value for RS
            self.endtour[i].set(self.onrec_endtour[i])  # set the starting value for ET
            self.codes[i].set(self.onrec_codes[i])
            self.lvtypes[i].set(self.onrec_lvtype[i])  # set the starting value for leave type
            self.lvtimes[i].set(self.onrec_lvtime[i])  # set the starting value for leave type

    def build_page(self):
        """ builds the screen """
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
        frame_i += 2  # the row where daily frames are placed on the grid. 
        for i in range(self.i_range):
            # for ring in self.daily_ringrecs:  # assign the values for each rings attribute
            grid_i = 0  # counter for the grid within the frame
            self.frame[i].grid(row=frame_i, padx=5, sticky="w")
            # Display the day and date
            if projvar.ns_code[self.carrecs[0][3]] == self.dates[i].strftime("%a"):
                Label(self.frame[i], text="{} NS DAY".format(self.dates[i].strftime("%a %b %d, %Y")), fg="red") \
                    .grid(row=grid_i, column=0, columnspan=5, sticky="w")
            else:
                Label(self.frame[i], text=self.dates[i].strftime("%a %b %d, %Y"), fg="blue") \
                    .grid(row=grid_i, column=0, columnspan=5, sticky="w")
            grid_i += 1
            widgetlist = self.get_widgetlist(i)  # returns a list with moves and/or tourrings or an empty list.
            ww = self.get_ww(widgetlist)
            colcount = 0
            if self.daily_carrecs[i][5] == projvar.invran_station:
                Label(self.frame[i], text="5200", fg=color[7]).grid(row=grid_i, column=colcount)  # Display 5200 label
                colcount += 1
                if "tourrings" in widgetlist:
                    Label(self.frame[i], text="BT", fg=color[7]).grid(row=grid_i, column=colcount)
                    colcount += 1
                if "moves" in widgetlist:
                    # display move off, move on and route triad.
                    Label(self.frame[i], text="MV off", fg=color[7]).grid(row=grid_i, column=colcount)
                    Label(self.frame[i], text="MV on", fg=color[7]).grid(row=grid_i, column=colcount + 1)
                    Label(self.frame[i], text="Route", fg=color[7]).grid(row=grid_i, column=colcount + 2)
                    colcount += 4
                Label(self.frame[i], text="RS", fg=color[7]).grid(row=grid_i, column=colcount)  # Display RS label
                colcount += 1
                if "tourrings" in widgetlist:
                    Label(self.frame[i], text="ET", fg=color[7]).grid(row=grid_i, column=colcount)  # Display ET label
                    colcount += 1
                Label(self.frame[i], text="code", fg=color[7]).grid(row=grid_i, column=colcount)  # Display code label
                colcount += 1
                # Display LV type label
                Label(self.frame[i], text="LV type", fg=color[7]).grid(row=grid_i, column=colcount)
                colcount += 1
                # Display LV time label
                Label(self.frame[i], text="LV time", fg=color[7]).grid(row=grid_i, column=colcount)
                grid_i += 1  # increment the grid to add a line
                colcount = 0  # reset the column counter to zero
                # Display the entry widgets
                # 5200 time
                total_widget[i] = Entry(self.frame[i], width=macadj(ww, 4), textvariable=self.totals[i])
                total_widget[i].grid(row=grid_i, column=colcount)
                colcount += 1
                # BT - begin tour
                if "tourrings" in widgetlist:  # only display if show BT/ET is configured.
                    Entry(self.frame[i], width=macadj(ww, 4), textvariable=self.begintour[i]) \
                        .grid(row=grid_i, column=colcount)
                    colcount += 1
                # Moves
                # build three entry widgets and a button for more moves.
                self.get_now_moves(i)  # converts onrec_moves to a daily array
                self.more_moves(i, ww, widgetlist)
                # More Moves Button
                if "moves" in widgetlist:  # don't show moves for aux, ptf and (maybe) otdl
                    colcount += 3
                mactext = "move moves"  # write text for move moves buttons for mac
                if "tourrings" in widgetlist:  # move moves button text is conditional on if tourrings
                    mactext = "+mv"  # if tourrings use a shorter version
                mv_button = Button(self.frame[i], text=macadj("more moves", mactext), fg=macadj("black", "grey"),
                       command=lambda x=i: self.more_moves(x, ww, widgetlist))
                if "moves" in widgetlist:  # don't show moves for aux, ptf and (maybe) otdl
                    mv_button.grid(row=grid_i, column=colcount)
                    colcount += 1
                # Return to Station (rs)
                Entry(self.frame[i], width=macadj(ww, 4), textvariable=self.rss[i]).grid(row=grid_i, column=colcount)
                colcount += 1
                # ET - end tour
                if "tourrings" in widgetlist:  # only display if show BT/ET is configured.
                    Entry(self.frame[i], width=macadj(ww, 4), textvariable=self.endtour[i]) \
                        .grid(row=grid_i, column=colcount)
                    colcount += 1
                # Codes/Notes
                if self.daily_carrecs[i][2] == "wal" or self.daily_carrecs[i][2] == "nl":
                    option_menu[i] = OptionMenu(self.frame[i], self.codes[i], *nolist_codes)
                elif self.daily_carrecs[i][2] == "otdl":
                    option_menu[i] = OptionMenu(self.frame[i], self.codes[i], *ot_codes)
                else:
                    option_menu[i] = OptionMenu(self.frame[i], self.codes[i], *aux_codes)
                option_menu[i].configure(width=macadj(7, 6))

                option_menu[i].grid(row=grid_i, column=colcount)  # code widget
                colcount += 1  # increment the column
                # Leave Type
                lv_option_menu[i] = OptionMenu(self.frame[i], self.lvtypes[i], *lv_options)
                lv_option_menu[i].configure(width=macadj(7, 6))
                lv_option_menu[i].grid(row=grid_i, column=colcount)  # leave type widget
                colcount += 1  # increment the column
                # Leave Time
                Entry(self.frame[i], width=macadj(ww, 4), textvariable=self.lvtimes[i]) \
                    .grid(row=grid_i, column=colcount)  # leave time widget
                colcount += 1  # increment the column
                # Refusals are not used
            else:
                self.totals.append(StringVar(self.frame[i]))  # 5200 entry widget
                self.rss.append(StringVar(self.frame[i]))  # RS entry
                if self.daily_carrecs[i][5] != "no record":  # display for records that are out of station
                    Label(self.frame[i], text="out of station: {}".format(self.daily_carrecs[i][5]),
                          fg="white", bg="grey", width=55, height=2, anchor="w").grid(row=grid_i, column=0)
                else:  # display for when there is no record relevant for that day.
                    Label(self.frame[i], text="no record", fg="white", bg="grey", width=55, height=2, anchor="w") \
                        .grid(row=grid_i, column=0)
            frame_i += 1
        f7 = Frame(self.win.body)
        f7.grid(row=frame_i)
        Label(f7, height=50).grid(row=1, column=0)  # extra white space on bottom of form to facilitate moves

    @staticmethod
    def triad_row_finder(index):
        """ finds the row of the moves entry widget or button """
        if index % 3 == 0:
            return int(index / 3)
        elif (index - 1) % 3 == 0:
            return int((index - 1) / 3)
        elif (index - 2) % 3 == 0:
            return int((index - 2) / 3)

    @staticmethod
    def triad_col_finder(index):
        """ finds the column of the moves widget """
        if index % 3 == 0:  # first column
            return int(0)
        elif (index - 1) % 3 == 0:  # second column
            return int(1)
        elif (index - 2) % 3 == 0:  # third column
            return int(2)

    def moves_col_increment(self):
        """ if the bt/et is being displayed, move the moves columns over one. """
        if self.tourrings:
            return 2
        return 1

    def more_moves(self, i, ww, widgetlist):
        """" this executes if the user presses the more moves button. it will generate more move triad entry
        widgets. """
        mm = []
        if self.days[i] == "sat":
            mm = self.sat_mm  # find the day in question and use the correlating  array
        elif self.days[i] == "sun":
            mm = self.sun_mm
        elif self.days[i] == "mon":
            mm = self.mon_mm
        elif self.days[i] == "tue":
            mm = self.tue_mm
        elif self.days[i] == "wed":
            mm = self.wed_mm
        elif self.days[i] == "thu":
            mm = self.thu_mm
        elif self.days[i] == "fri":
            mm = self.fri_mm
        if not projvar.invran_weekly_span:  # if the invran is daily, not weekly - use daily multi moves holder
            mm = self.daily_mm  # holds daily stringvars for moves when invran is daily
        if not self.now_moves:  # if first time or button pushed.
            for _ in range(3):
                self.now_moves.append("")
        mci = self.moves_col_increment()  # moves counter increment: if bt/et is displayed, increment the column by 1
        mwc = len(mm)  # moves widget counter: count the number of widgets
        for ii in range(len(self.now_moves)):  # loop through all items in array of stringvars for the day.
            mm.append(StringVar(self.frame[i]))  # create entry field for moves from database
            mm[ii + mwc].set(self.now_moves[ii])  # set values for the StringVars
            if (ii + 1) % 3 == 0:  # adjust the lenght of the route widget pending os
                ml = 5  # on mac, the route widget lenght is 5
            else:
                ml = 4  # on mac, the rings widget lenght is 4
            # build the widget use triad col/row finder to locate right place on grid
            mv_widget = Entry(self.frame[i], width=macadj(ww, ml), textvariable=mm[ii + mwc])
            if "moves" in widgetlist:  # don't show moves for aux, ptf and (maybe) otdl
                mv_widget.grid(column=self.triad_col_finder(ii + mwc) + mci, row=self.triad_row_finder(ii + mwc) + 2)
        self.now_moves = []  # zero out self.now_moves so more moves button works properly

    def write_report(self):
        """ build the report to appear on bottom of screen """
        if not self.status_update:
            return
        if self.delete_report + self.update_report + self.insert_report == 0:
            self.status_update = "No records changed. "  # if there are no changes
            return
        status_update = ""
        if self.insert_report:  # new records
            status_update += str(self.insert_report) + " new record{} added. " \
                .format(Handler(self.insert_report).plurals())  # make "record" plural if necessary
        if self.update_report:  # updated records
            status_update += str(self.update_report) + " record{} updated. " \
                .format(Handler(self.update_report).plurals())  # make "record" plural if necessary
        if self.delete_report:  # deleted records
            status_update += str(self.delete_report) + " record{} deleted. " \
                .format(Handler(self.delete_report).plurals())  # make "record" plural if necessary
        self.status_update = status_update

    def buttons_frame(self):
        """ build the buttons for the bottom of the screen """
        button_alignment = macadj("w", "center")
        Button(self.win.buttons, text="Submit", width=10, anchor=button_alignment,
               command=lambda: self.apply_rings(True)).pack(side=LEFT)
        Button(self.win.buttons, text="Apply", width=10, anchor=button_alignment,
               command=lambda: self.apply_rings(False)).pack(side=LEFT)
        Button(self.win.buttons, text="Go Back", width=10, anchor=button_alignment,
               command=lambda: self.win.root.destroy()).pack(side=LEFT)
        Label(self.win.buttons, text="{}".format(self.status_update), fg="red").pack(side=LEFT)

    def zero_report_vars(self):
        """ initializes the report variables. """
        self.status_update = "No records changed."
        self.delete_report = 0
        self.update_report = 0
        self.insert_report = 0

    def apply_rings(self, go_home):
        """ execute when apply or submit is pressed """
        self.empty_addrings()  # make sure addrings are empty before building them
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
        self.add_refusals()
        if not self.check_bt():
            return  # abort if there is an error
        if not self.check_et():
            return  # abort if there is an error
        self.addrecs()  # insert rings into the database
        if go_home:  # if True, then exit screen to main screen
            self.win.root.destroy()
        else:  # if False, then rebuild the Enter Rings screen
            self.start(self.win.topframe)

    def empty_addrings(self):
        """ empty out addring arrays """
        for i in range(len(self.addrings)):
            self.addrings[i] = []

    def add_date(self):
        """ start the addrings array """
        for i in range(len(self.dates)):  # loop for each day in the investigation
            self.addrings[i].append(self.dates[i])  # add the date
            self.addrings[i].append(self.carrier)  # add the carrier name

    def check_5200(self):
        """ a check for the 5200 time """
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
        """ a check for return to station time """
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
        """ adds the code to the an array of values to be entered into the database """
        for i in range(len(self.codes)):
            self.addrings[i].append(self.codes[i].get())

    def daily_moves_builder(self, first, second, third):
        """ build an array of daily moves """
        if first.strip() and second.strip():  # if the first and second values are not empty
            self.daily_moves.append(first)
            self.daily_moves.append(second)
            self.daily_moves.append(third)

    def check_moves(self):
        """ checks the moves for errors """
        cc = 0  # index: increments one for each day
        for d in self.moves_array:  # check for bad inputs in moves
            if not self.check_moves_daily(d, cc):
                return False
            cc += 1
        return True

    def check_moves_daily(self, d, cc):
        """  checks the moves for errors by the day  """
        first_move = None
        second_move = None
        self.daily_moves = []  # empty out array where moves data is passed
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
                    # self.move_string_constructor(first_move, second_move, "")
                    self.daily_moves_builder(first_move, second_move, "")
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
                self.daily_moves_builder(first_move, second_move, route)
        self.addrings[cc].append(self.daily_moves)
        return True

    def add_leavetype(self):
        """ adds the leave type into an array to be entered into the database. """
        for i in range(len(self.lvtypes)):
            self.addrings[i].append(self.lvtypes[i].get())

    def check_leave(self):
        """ checks then adds the leave time into an array to be entered into the database. """
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
                text = "Values with more than 2 decimal places are not accepted in Leave Time for {}." \
                    .format(self.day[i])
                messagebox.showerror("Leave Time Error", text, parent=self.win.topframe)
                return False
            lvtime = Convert(lvtime).hundredths()  # format it as a number with 2 decimal places
            self.addrings[i].append(lvtime)  # if all checks pass, add to addrings
        return True

    def add_refusals(self):
        """ adds the refusals into an array to be entered into the database. """
        for i in range(len(self.refusals)):
            self.addrings[i].append(self.refusals[i])

    def check_bt(self):
        """ a check for begin tour time """
        for i in range(len(self.begintour)):
            bt = str(self.begintour[i].get()).strip()
            if RingTimeChecker(bt).check_for_zeros():
                self.addrings[i].append("")  # if variable is zero or empty, add an empty string to addrings
                continue  # skip other checks
            if not RingTimeChecker(bt).check_numeric():
                text = "You must enter a numeric value in BT for {}.".format(self.day[i])
                messagebox.showerror("BT Error", text, parent=self.win.topframe)
                return False
            if not RingTimeChecker(bt).over_24():
                text = "Values greater than 24 are not accepted in BT for {}.".format(self.day[i])
                messagebox.showerror("BT Error", text, parent=self.win.topframe)
                return False
            if not RingTimeChecker(bt).less_than_zero():
                text = "Values less than or equal to 0 are not accepted in BT for {}.".format(self.day[i])
                messagebox.showerror("BT Error", text, parent=self.win.topframe)
                return False
            if not RingTimeChecker(bt).count_decimals_place():
                text = "Values with more than 2 decimal places are not accepted in BT for {}.".format(self.day[i])
                messagebox.showerror("BT Error", text, parent=self.win.topframe)
                return False
            bt = Convert(bt).hundredths()  # format it as a number with 2 decimal places
            self.addrings[i].append(bt)  # if all checks pass, add to addrings
        return True

    def check_et(self):
        """ a check for end tour time """
        for i in range(len(self.endtour)):
            et = str(self.endtour[i].get()).strip()
            if et == "" or et == "auto":
                total = self.totals[i].get().strip()
                bt = str(self.begintour[i].get()).strip()
                autotime = ""  # convert empty string or "auto" to empty string.
                if total and bt:  # if 5200 and bt both exist, then calculate the automated end tour.
                    autotime = self.auto_endtour(total, bt)  # returns the automated end tour time.
                self.addrings[i].append(autotime)  # if variable is zero or empty, add an empty string to addrings
                continue  # skip other checks
            if RingTimeChecker(et).check_for_zeros():
                self.addrings[i].append("")  # if variable is zero or empty, add an empty string to addrings
                continue  # skip other checks
            if not RingTimeChecker(et).check_numeric():
                text = "You must enter a numeric value in ET for {}.".format(self.day[i])
                messagebox.showerror("ET Error", text, parent=self.win.topframe)
                return False
            if not RingTimeChecker(et).over_24():
                text = "Values greater than 24 are not accepted in ET for {}.".format(self.day[i])
                messagebox.showerror("ET Error", text, parent=self.win.topframe)
                return False
            if not RingTimeChecker(et).less_than_zero():
                text = "Values less than or equal to 0 are not accepted in ET for {}.".format(self.day[i])
                messagebox.showerror("ET Error", text, parent=self.win.topframe)
                return False
            if not RingTimeChecker(et).count_decimals_place():
                text = "Values with more than 2 decimal places are not accepted in ET for {}.".format(self.day[i])
                messagebox.showerror("ET Error", text, parent=self.win.topframe)
                return False
            et = Convert(et).hundredths()  # format it as a number with 2 decimal places
            self.addrings[i].append(et)  # if all checks pass, add to addrings
        return True

    @staticmethod
    def auto_endtour(total, bt):
        """ add 50 clicks to the begin tour an 5200 time """
        if float(total) >= 6:  # if the 5200 time is 6 hours or more
            auto_et = float(bt) + float(total) + .50  # ET is automatically calculated: ET = 5200 + BT + lunch
        else:  # if 5200 time is less than 6 hours
            auto_et = float(bt) + float(total)  # ET is automatically calculated: ET = 5200 + BT (no lunch added)
        if auto_et >= 24:  # if the calculated ET goes over midnight
            auto_et -= 24  # make a correct AM time.
        return "{:.2f}".format(auto_et)  # return as a string with two decimal places.

    def addrecs(self):
        """ add records to database """
        sql = ""
        for i in range(len(self.dates)):
            # convert the daily moves from an array into a string.
            self.addrings[i][5] = Convert(self.addrings[i][5]).array_to_string()
            # empty_rec = [self.dates[i], self.carrier, "", "", "none", "", "none", "", "", "", ""]
            empty_rec = [self.dates[i], self.carrier, "", "", "none", "", "none", "", None, "", ""]
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
                sql = "UPDATE rings3 SET total='%s',rs='%s',code='%s',moves='%s',leave_type ='%s'," \
                      "leave_time='%s', bt='%s', et='%s' WHERE rings_date='%s' and carrier_name='%s'" \
                      % (self.addrings[i][2], self.addrings[i][3], self.addrings[i][4],
                         self.addrings[i][5], self.addrings[i][6], self.addrings[i][7],
                         self.addrings[i][9], self.addrings[i][10],
                         self.dates[i], self.carrier)
                self.update_report += 1
            elif self.daily_ringrecs[i] == empty_rec and self.addrings[i] != empty_rec:
                # if a record doesn't exist and the new record is not empty
                sql = "INSERT INTO rings3 (rings_date, carrier_name, total, rs, code, moves, leave_type, leave_time, " \
                      "bt, et )" \
                      "VALUES('%s','%s','%s','%s','%s','%s','%s','%s','%s','%s') " \
                      % (self.dates[i], self.carrier, self.addrings[i][2], self.addrings[i][3],
                         self.addrings[i][4], self.addrings[i][5], self.addrings[i][6], self.addrings[i][7],
                         self.addrings[i][9], self.addrings[i][10])
                self.insert_report += 1
            if sql:
                commit(sql)
