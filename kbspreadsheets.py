"""
a klusterbox module: Klusterbox Improper Mandates and 12 and 60 Hour Violations Spreadsheets Generator.
klusterbox classes for spreadsheets: the improper mandate worksheet and the 12 and 60 hour violations spreadsheets
"""
import projvar  # custom libraries
from kbtoolbox import inquire, CarrierList, dir_path, isfloat, Convert, Rings, ProgressBarDe, Moves
# standard libraries
from tkinter import messagebox
import os
import sys
import subprocess
from datetime import timedelta
# non standard libraries
from openpyxl import Workbook
from openpyxl.worksheet.pagebreak import Break
from openpyxl.styles import NamedStyle, Font, Border, Side, Alignment, PatternFill


class ImpManSpreadsheet:
    """
    This generates the famous klusterbox spreadsheets breaking down availability and off route mandates.
    """
    def __init__(self):
        self.frame = None  # the frame of parent
        self.pb = None  # progress bar object
        self.pbi = 0  # progress bar count index
        self.startdate = None  # start date of the investigation
        self.enddate = None  # ending date of the investigation
        self.dates = []  # all days of the investigation
        self.carrierlist = []  # all carriers in carrier list
        self.carrier_breakdown = []  # all carriers in carrier list broken down into appropiate list
        self.wb = None  # the workbook object
        self.ws_list = []  # "saturday", "sunday", "monday", "tuesday", "wednesday", "thursday", "friday"
        self.summary = None  # worksheet for summary page
        self.reference = None  # worksheet for reference page
        self.ws_header = None  # style
        self.list_header = None  # style
        self.date_dov = None  # style
        self.date_dov_title = None  # style
        self.col_header = None  # style
        self.input_name = None  # style
        self.input_s = None  # style
        self.calcs = None  # style
        self.min_ss_nl = 0  # minimum rows for "no list"
        self.min_ss_wal = 0  # minimum rows for work assignment list
        self.min_ss_otdl = 0  # minimum rows for overtime desired list
        self.min_ss_aux = 0  # minimum rows for auxiliary
        self.day = None  # build worksheet - loop once for each day
        self.i = 0  # build worksheet loop iteration
        self.lsi = 0  # list loop iteration
        self.pref = ("nl", "wal", "otdl", "aux")
        self.ot_list = ("No List Carriers", "Work Assignment Carriers", "Overtime Desired List Carriers",
                        "Auxiliary Assistance")  # list loop iteration
        self.row = 0  # list loop iteration/ the row placement
        self.mod_carrierlist = []  # carrier list with empty recs added to reach minimum row quantity
        self.carrier = []  # carrier name
        self.rings = []  # carrier rings queried from database
        self.totalhours = ""  # carrier rings - 5200 time
        self.codes = ""  # carrier rings - code/note
        self.rs = ""  # carrier rings - return to station
        self.moves = ""  # carrier rings - moves on and off route with route
        self.lvtype = ""  # carrier rings - leave type
        self.lvtime = ""  # carrier rings - leave time
        self.movesarray = []
        self.move_i = 0  # increments rows for multiple move functionality
        self.tol_ot_ownroute = 0.0  # tolerance for ot on own route
        self.tol_ot_offroute = 0.0  # tolerance for ot off own route
        self.tol_availability = 0.0  # tolerance for availability
        self.pb_nl_wal = True  # page break between no list and work assignment
        self.pb_wal_otdl = True  # page break between work assignment and otdl
        self.pb_otdl_aux = True  # page break between otdl and auxiliary
        self.day_of_week = []  # seven day array for weekly investigations/ one day array for daily investigations
        self.mandates_own_route = []  # stores cell location on each sheet for no list own route overtime
        self.mandates_all = []  # stores cell location for total mandates for each sheet
        self.availability_10 = []  # stores cell location for total availability to 10 hrs for each sheet
        self.availability_max = []  # stores cell location for total maximum availabilityfor each sheet
        self.first_row = 0  # stores the first row for each list, re initialized at end of list
        self.last_row = 0  # stores the last row for each list, re initialized at end of list
        self.subtotal_loc_holder = []  # stores the cell location of a subtotal for total mandates/ availability

    def create(self, frame):
        """ master method for calling all methods in class """
        self.frame = frame
        if not self.ask_ok():  # abort if user selects cancel from askokcancel
            return
        self.pb = ProgressBarDe(label="Building Improper Mandates Spreadsheet")
        self.pb.max_count(100)  # set length of progress bar
        self.pb.start_up()  # start the progress bar
        self.pbi = 1
        self.pb.move_count(self.pbi)  # increment progress bar
        self.pb.change_text("Gathering Data... ")
        self.get_dates()
        self.get_pb_max_count()
        self.get_carrierlist()
        self.get_carrier_breakdown()  # breakdown carrier list into no list, wal, otdl, aux
        self.get_tolerances()  # get tolerances, minimum rows and page break preferences from tolerances table
        self.get_styles()
        self.build_workbook()
        self.set_dimensions()
        self.build_refs()
        self.build_ws_loop()  # calls list loop and carrier loop
        self.build_summary_header()
        self.build_summary()
        self.save_open()

    def ask_ok(self):
        """ ends process if user cancels """
        if messagebox.askokcancel("Spreadsheet generator",
                                  "Do you want to generate an Improper Mandates Spreadsheet?",
                                  parent=self.frame):
            return True
        return False

    def get_dates(self):
        """ get the dates from the project variables """
        self.startdate = projvar.invran_date  # set daily investigation range as default - get start date
        self.enddate = projvar.invran_date  # get end date
        self.dates = [projvar.invran_date, ]  # create an array of days - only one day if daily investigation range
        if projvar.invran_weekly_span:  # if the investigation range is weekly
            date = projvar.invran_date_week[0]
            self.startdate = projvar.invran_date_week[0]
            self.enddate = projvar.invran_date_week[6]
            self.dates = []
            for _ in range(7):  # create an array with all the days in the weekly investigation range
                self.dates.append(date)
                date += timedelta(days=1)

    def get_pb_max_count(self):
        """ set length of progress bar """
        self.pb.max_count((len(self.dates)*4)+3)  # once for each list in each day, plus reference, summary and saving

    def get_carrierlist(self):
        """ get record sets for all carriers """
        self.carrierlist = CarrierList(self.startdate, self.enddate, projvar.invran_station).get()

    def get_carrier_breakdown(self):
        """ breakdown carrier list into no list, wal, otdl, aux """
        timely_rec = []
        for day in self.dates:
            nl_array = []
            wal_array = []
            otdl_array = []
            aux_array = []
            for carrier in self.carrierlist:
                for rec in reversed(carrier):
                    if Convert(rec[0]).dt_converter() <= day:
                        timely_rec = rec
                if timely_rec[2] == "nl":
                    nl_array.append(timely_rec)
                if timely_rec[2] == "wal":
                    wal_array.append(timely_rec)
                if timely_rec[2] == "otdl":
                    otdl_array.append(timely_rec)
                if timely_rec[2] == "aux" or timely_rec[2] == "ptf":
                    aux_array.append(timely_rec)
            daily_breakdown = [nl_array, wal_array, otdl_array, aux_array]
            self.carrier_breakdown.append(daily_breakdown)

    def get_tolerances(self):
        """ get spreadsheet tolerances, row minimums and page break prefs from tolerance table """
        sql = "SELECT tolerance FROM tolerances"
        result = inquire(sql)
        self.tol_ot_ownroute = float(result[0][0])  # overtime on own route tolerance
        self.tol_ot_offroute = float(result[1][0])  # overtime off own route tolerance
        self.tol_availability = float(result[2][0])  # availability tolerance
        self.min_ss_nl = int(result[3][0])  # minimum rows for no list
        self.min_ss_wal = int(result[4][0])  # mimimum rows for work assignment
        self.min_ss_otdl = int(result[5][0])  # minimum rows for otdl
        self.min_ss_aux = int(result[6][0])  # minimum rows for auxiliary
        self.pb_nl_wal = Convert(result[21][0]).str_to_bool()  # page break between no list and work assignment
        self.pb_wal_otdl = Convert(result[22][0]).str_to_bool()  # page break between work assignment and otdl
        self.pb_otdl_aux = Convert(result[23][0]).str_to_bool()  # page break between otdl and auxiliary

    def get_styles(self):
        """ Named styles for workbook """
        bd = Side(style='thin', color="80808080")  # defines borders
        self.ws_header = NamedStyle(name="ws_header", font=Font(bold=True, name='Arial', size=12))
        self.list_header = NamedStyle(name="list_header", font=Font(bold=True, name='Arial', size=10))
        self.date_dov = NamedStyle(name="date_dov", font=Font(name='Arial', size=8))
        self.date_dov_title = NamedStyle(name="date_dov_title", font=Font(bold=True, name='Arial', size=8),
                                         alignment=Alignment(horizontal='right'))
        self.col_header = NamedStyle(name="col_header", font=Font(bold=True, name='Arial', size=8),
                                     alignment=Alignment(horizontal='right'))
        self.input_name = NamedStyle(name="input_name", font=Font(name='Arial', size=8),
                                     border=Border(left=bd, top=bd, right=bd, bottom=bd))
        self.input_s = NamedStyle(name="input_s", font=Font(name='Arial', size=8),
                                  border=Border(left=bd, top=bd, right=bd, bottom=bd),
                                  alignment=Alignment(horizontal='right'))
        self.calcs = NamedStyle(name="calcs", font=Font(name='Arial', size=8),
                                border=Border(left=bd, top=bd, right=bd, bottom=bd),
                                fill=PatternFill(fgColor='e5e4e2', fill_type='solid'),
                                alignment=Alignment(horizontal='right'))

    def build_workbook(self):
        """ build the workbook object """
        day_finder = ["Sat", "Sun", "Mon", "Tue", "Wed", "Thu", "Fri"]
        day_of_week = ["saturday", "sunday", "monday", "tuesday", "wednesday", "thursday", "friday"]
        i = 0
        self.wb = Workbook()  # define the workbook
        if not projvar.invran_weekly_span:  # if investigation range is daily
            for ii in range(len(day_finder)):
                if projvar.invran_date.strftime("%a") == day_finder[ii]:  # find the correct day
                    i = ii
            self.ws_list.append(self.wb.active)  # create first worksheet
            self.ws_list[0].title = day_of_week[i]  # title first worksheet
            self.day_of_week.append(day_of_week[i])  # create self.day_of_week array with one day
        if projvar.invran_weekly_span:  # if investigation range is weekly
            for day in day_of_week:
                self.day_of_week.append(day)  # create self.day_of_week array with seven days
            self.ws_list.append(self.wb.active)  # create first worksheet
            self.ws_list[0].title = "saturday"  # title first worksheet
            for i in range(1, 7):  # create worksheet for remaining six days
                self.ws_list.append(self.wb.create_sheet(day_of_week[i]))  # create subsequent worksheets
                self.ws_list[i].title = day_of_week[i]  # title subsequent worksheets
        self.summary = self.wb.create_sheet("summary")
        self.reference = self.wb.create_sheet("reference")

    def set_dimensions(self):
        """ set the dimensions of the workbook """
        for i in range(len(self.dates)):
            self.ws_list[i].oddFooter.center.text = "&A"
            self.ws_list[i].column_dimensions["A"].width = 20
            self.ws_list[i].column_dimensions["B"].width = 5
            self.ws_list[i].column_dimensions["C"].width = 6
            self.ws_list[i].column_dimensions["D"].width = 6
            self.ws_list[i].column_dimensions["E"].width = 6
            self.ws_list[i].column_dimensions["F"].width = 6
            self.ws_list[i].column_dimensions["G"].width = 6
            self.ws_list[i].column_dimensions["H"].width = 7
            self.ws_list[i].column_dimensions["I"].width = 6
            self.ws_list[i].column_dimensions["J"].width = 6
            self.ws_list[i].column_dimensions["K"].width = 7
        self.summary.column_dimensions["A"].width = 14
        self.summary.column_dimensions["B"].width = 9
        self.summary.column_dimensions["C"].width = 9
        self.summary.column_dimensions["D"].width = 9
        self.summary.column_dimensions["E"].width = 2
        self.summary.column_dimensions["F"].width = 9
        self.summary.column_dimensions["G"].width = 9
        self.summary.column_dimensions["H"].width = 9
        self.reference.column_dimensions["A"].width = 14
        self.reference.column_dimensions["B"].width = 8
        self.reference.column_dimensions["C"].width = 8
        self.reference.column_dimensions["D"].width = 2
        self.reference.column_dimensions["E"].width = 50

    def build_refs(self):
        """ build the references page. This shows tolerances and defines labels. """
        self.pbi += 1
        self.pb.move_count(self.pbi)  # increment progress bar
        self.pb.change_text("Building Reference Page")
        # tolerances
        self.reference['B2'].style = self.list_header
        self.reference['B2'] = "Tolerances"
        self.reference['C3'] = self.tol_ot_ownroute  # overtime on own route tolerance
        self.reference['C3'].style = self.input_s
        self.reference['C3'].number_format = "#,###.00;[RED]-#,###.00"
        self.reference['E3'] = "overtime on own route"
        self.reference['C4'] = self.tol_ot_offroute  # overtime off own route tolerance
        self.reference['C4'].style = self.input_s
        self.reference['C4'].number_format = "#,###.00;[RED]-#,###.00"
        self.reference['E4'] = "overtime off own route"
        self.reference['C5'] = self.tol_availability  # availability tolerance
        self.reference['C5'].style = self.input_s
        self.reference['C5'].number_format = "#,###.00;[RED]-#,###.00"
        self.reference['E5'] = "availability tolerance"
        # note guide
        self.reference['B7'].style = self.list_header
        self.reference['B7'] = "Note Guide"
        self.reference['C8'] = "ns day"
        self.reference['C8'].style = self.input_s
        self.reference['E8'] = "Carrier worked on their non scheduled day"
        self.reference['C10'] = "no call"
        self.reference['C10'].style = self.input_s
        self.reference['E10'] = "Carrier was not scheduled for overtime"
        self.reference['C11'] = "light"
        self.reference['C11'].style = self.input_s
        self.reference['E11'] = "Carrier on light duty and unavailable for overtime"
        self.reference['C12'] = "sch chg"
        self.reference['C12'].style = self.input_s
        self.reference['E12'] = "Schedule change: unavailable for overtime"
        self.reference['C13'] = "annual"
        self.reference['C13'].style = self.input_s
        self.reference['E13'] = "Annual leave"
        self.reference['C14'] = "sick"
        self.reference['C14'].style = self.input_s
        self.reference['E14'] = "Sick leave"
        self.reference['C15'] = "excused"
        self.reference['C15'].style = self.input_s
        self.reference['E15'] = "Carrier excused from mandatory overtime"
        # column headers
        self.reference['B17'].style = self.list_header
        self.reference['B17'] = "Column Headers"
        self.reference['C18'] = "Name"
        self.reference['C18'].style = self.input_s
        self.reference['E18'] = "The name of the carrier. "
        self.reference['C19'] = "Note"
        self.reference['C19'].style = self.input_s
        self.reference['E19'] = "Special circumstances. See note guide above."
        self.reference['C20'] = "5200"
        self.reference['C20'].style = self.input_s
        self.reference['E20'] = "Total hours worked"
        self.reference['C21'] = "RS"
        self.reference['C21'].style = self.input_s
        self.reference['E21'] = "Return to station time."
        self.reference['C22'] = "MV off"
        self.reference['C22'].style = self.input_s
        self.reference['E22'] = "Time moved off own route"
        self.reference['C23'] = "MV on"
        self.reference['C23'].style = self.input_s
        self.reference['E23'] = "Time moved on/returned to own route"
        self.reference['C24'] = "Route"
        self.reference['C24'].style = self.input_s
        self.reference['E24'] = "Route of overtime/pivot"
        self.reference['C25'] = "MV Total"
        self.reference['C25'].style = self.input_s
        self.reference['E25'] = "Time spent on overtime/pivot off route"
        self.reference['C26'] = "OT"
        self.reference['C26'].style = self.input_s
        self.reference['E26'] = "Daily overtime"
        self.reference['C27'] = "Off rte"
        self.reference['C27'].style = self.input_s
        self.reference['E27'] = "Total daily time spent off route"
        self.reference['C28'] = "OT off"
        self.reference['C28'].style = self.input_s
        self.reference['E28'] = "Daily overtime off route"
        self.reference['C30'] = "to 10"
        self.reference['C30'].style = self.input_s
        self.reference['E30'] = "Total availability to 10 hours"
        self.reference['C31'] = "to 12"
        self.reference['C31'].style = self.input_s
        self.reference['E31'] = "Total availability to 12 hours"
        
    def build_ws_loop(self):
        """ this loops once for each list. """
        self.i = 0
        for day in self.dates:
            self.day = day
            self.build_ws_headers()
            self.list_loop()  # loops four times. once for each list.
            self.i += 1

    def build_ws_headers(self):
        """ worksheet headers """
        cell = self.ws_list[self.i].cell(row=1, column=1)
        cell.value = "Improper Mandate Worksheet"
        cell.style = self.ws_header
        self.ws_list[self.i].merge_cells('A1:E1')
        cell = self.ws_list[self.i].cell(row=3, column=1)
        cell.value = "Date:  "  # create date/ pay period/ station header
        cell.style = self.date_dov_title
        cell = self.ws_list[self.i].cell(row=3, column=2)
        cell.value = format(self.day, "%A  %m/%d/%y")
        cell.style = self.date_dov
        self.ws_list[self.i].merge_cells('B3:D3')
        cell = self.ws_list[self.i].cell(row=3, column=5)
        cell.value = "Pay Period:  "
        cell.style = self.date_dov_title
        self.ws_list[self.i].merge_cells('E3:F3')
        cell = self.ws_list[self.i].cell(row=3, column=7)
        cell.value = projvar.pay_period
        cell.style = self.date_dov
        self.ws_list[self.i].merge_cells('G3:H3')
        cell = self.ws_list[self.i].cell(row=4, column=1)
        cell.value = "Station:  "
        cell.style = self.date_dov_title
        cell = self.ws_list[self.i].cell(row=4, column=2)
        cell.value = projvar.invran_station
        cell.style = self.date_dov
        self.ws_list[self.i].merge_cells('B4:D4')

    def increment_progbar(self):
        """ move the progress bar, update with info on what is being done """
        lst = ("No List", "Work Assignment", "Overtime Desired", "Auxiliary")
        self.pbi += 1
        self.pb.move_count(self.pbi)  # increment progress bar
        self.pb.change_text("Building day {}: list: {}".format(self.day.strftime("%A"), lst[self.lsi]))

    def list_loop(self):
        """ loops four times. once for each list. """
        self.lsi = 0  # iterations of the list loop method
        self.row = 6
        for _ in self.ot_list:  # loops for nl, wal, otdl and aux
            self.list_and_column_headers()  # builds headers for list and column
            self.carrierlist_mod()
            self.get_first_row()
            self.carrierloop()
            self.build_footer()
            self.pagebreak()
            self.increment_progbar()
            self.lsi += 1

    def list_and_column_headers(self):
        """ builds headers for list and column """
        cell = self.ws_list[self.i].cell(row=self.row, column=1)
        cell.value = self.ot_list[self.lsi]  # "No List Carriers",
        cell.style = self.list_header
        if self.pref[self.lsi] in ("nl", "wal"):
            self.row += 1
        else:
            self.row += 2
        cell = self.ws_list[self.i].cell(row=self.row, column=1)  # column headers for any list
        cell.value = "Name"
        cell.style = self.col_header
        cell = self.ws_list[self.i].cell(row=self.row, column=2)
        cell.value = "note"
        cell.style = self.col_header
        cell = self.ws_list[self.i].cell(row=self.row, column=3)
        cell.value = "5200"
        cell.style = self.col_header
        cell = self.ws_list[self.i].cell(row=self.row, column=4)
        cell.value = "RS"
        cell.style = self.col_header
        if self.pref[self.lsi] in ("nl", "wal"):
            self.column_header_non()  # column headers specific for non otdl
        else:
            self.column_header_ot()  # column headers specific for otdl or aux

    def column_header_non(self):
        """ column headers specific for non otdl """
        cell = self.ws_list[self.i].cell(row=self.row, column=5)
        cell.value = "MV off"
        cell.style = self.col_header
        cell = self.ws_list[self.i].cell(row=self.row, column=6)
        cell.value = "MV on"
        cell.style = self.col_header
        cell = self.ws_list[self.i].cell(row=self.row, column=7)
        cell.value = "Route"
        cell.style = self.col_header
        cell = self.ws_list[self.i].cell(row=self.row, column=8)
        cell.value = "MV total"
        cell.style = self.col_header
        cell = self.ws_list[self.i].cell(row=self.row, column=9)
        cell.value = "OT"
        cell.style = self.col_header
        cell = self.ws_list[self.i].cell(row=self.row, column=10)
        cell.value = "off rt"
        cell.style = self.col_header
        cell = self.ws_list[self.i].cell(row=self.row, column=11)
        cell.value = "OT off"
        cell.style = self.col_header
        self.row += 1

    def column_header_ot(self):
        """ column headers specific for otdl or aux """
        cell = self.ws_list[self.i].cell(row=self.row - 1, column=5)
        cell.value = "Availability to:"
        cell.style = self.col_header
        to_what = "to 11.5"
        if self.pref[self.lsi] == "otdl":
            to_what = "to 12"
        cell = self.ws_list[self.i].cell(row=self.row, column=5)
        cell.value = "to 10"
        cell.style = self.col_header
        cell = self.ws_list[self.i].cell(row=self.row, column=6)
        cell.value = to_what
        cell.style = self.col_header
        self.row += 1

    def carrierlist_mod(self):
        """ add empty carrier records to carrier list until quantity matches minrows preference """
        self.mod_carrierlist = self.carrier_breakdown[self.i][self.lsi]
        if self.pref[self.lsi] in ("nl",):  # if "no list"
            minrows = self.min_ss_nl
        elif self.pref[self.lsi] in ("wal",):  # if "work assignment list"
            minrows = self.min_ss_wal
        elif self.pref[self.lsi] in ("otdl",):  # if "overtime desired list"
            minrows = self.min_ss_otdl
        else:  # if "auxiliary"
            minrows = self.min_ss_aux
        while len(self.mod_carrierlist) < minrows:  # until carrier list quantity matches minrows
            add_this = ('', '', '', '', '', '')
            self.mod_carrierlist.append(add_this)  # append empty recs to carrier list

    def get_first_row(self):
        """ record the number of the first row for totals formulas in footers """
        self.first_row = self.row

    def carrierloop(self):
        """ loop for each carrier """
        for carrier in self.mod_carrierlist:
            self.get_last_row()  # record the number of the last row for total formulas in footers
            self.carrier = carrier[1]  # current iteration of carrier list is assigned self.carrier
            self.get_rings()  # get individual carrier rings for the day
            self.display_recs()  # put the carrier and the first part of rings into the spreadsheet
            if self.pref[self.lsi] in ("nl", "wal"):  # if the list is no list or work assignment
                self.get_movesarray()  # get the moves
                self.display_moves()  # display the moves
                self.display_formulas_non()  # display the formulas
            else:  # if otdl or aux
                self.display_formulas_ot()  # display formulas for otdl/aux
            self.increment_rows()

    def get_last_row(self):
        """ record the number of the last row for totals formulas in footers """
        self.last_row = self.row

    def increment_rows(self):
        """ increment the rows counter """
        self.row += 1
        self.row += self.move_i  # add 1 plus any the added rows from multiple moves
        self.move_i = 0  # reset the row incrementor for multiple move functionality

    def get_rings(self):
        """ get individual carrier rings for the day """
        self.rings = Rings(self.carrier, self.dates[self.i]).get_for_day()  # assign as self.rings
        self.totalhours = ""  # set default as an empty string
        self.rs = ""
        self.codes = ""
        self.moves = ""
        self.lvtype = ""
        self.lvtime = ""
        if self.rings[0]:  # if rings record is not blank
            self.totalhours = self.rings[0][2]
            self.rs = self.rings[0][3]
            self.codes = self.rings[0][4]
            self.moves = self.rings[0][5]
            self.lvtype = self.rings[0][6]
            self.lvtime = self.rings[0][7]

    def display_recs(self):
        """ put the carrier and the first part of rings into the spreadsheet """
        cell = self.ws_list[self.i].cell(row=self.row, column=1)  # name
        cell.value = self.carrier
        cell.style = self.input_name
        cell = self.ws_list[self.i].cell(row=self.row, column=2)  # code
        cell.value = Convert(self.codes).empty_not_none()
        cell.style = self.input_s
        cell = self.ws_list[self.i].cell(row=self.row, column=3)  # 5200
        cell.value = Convert(self.totalhours).str_to_floatoremptystr()
        cell.style = self.input_s
        cell.number_format = "#,###.00;[RED]-#,###.00"
        cell = self.ws_list[self.i].cell(row=self.row, column=4)  # return to station
        cell.value = Convert(self.rs).str_to_floatoremptystr()
        cell.style = self.input_s
        cell.number_format = "#,###.00;[RED]-#,###.00"

    def get_movesarray(self):
        """ builds sets of moves for each triad """
        multiple_sets = False  # is there more than one triad?
        self.movesarray = []  # re initialized - a list of tuples of move sets
        moves_array = []  # initialized - the moves string converted into an array
        move_off = ""  # if empty set, use default values
        move_back = ""
        move_route = ""
        formula = "=SUM(%s!F%s - %s!E%s)" % (self.day_of_week[self.i], self.row, self.day_of_week[self.i], self.row)
        if not self.moves:  # if string is empty
            pass  # use default values
        else:  # if the string is not empty
            moves_array = Convert(self.moves).string_to_array()
            if len(moves_array)/3 == 1:  # if there is only one set of moves
                move_off = moves_array[0]
                move_back = moves_array[1]
                move_route = moves_array[2]
            else:  # if there are multiple move sets
                multiple_sets = True
                move_off = "*"
                move_back = "*"
                move_route = "*"
                formula = "=SUM(%s!H%s:H%s)" % \
                          (self.day_of_week[self.i], self.row + 1, int(self.row + len(moves_array) / 3))
        add_this = (move_off, move_back, move_route, formula)
        self.movesarray.append(add_this)
        if multiple_sets:  # if multiple sets are detected
            i = 0
            formula_row_i = 1  # increment the row in the formula
            for move in moves_array:
                if (i + 3) % 3 == 0:
                    move_off = move
                if (i + 2) % 3 == 0:
                    move_back = move
                if (i + 1) % 3 == 0:
                    move_route = move
                    formula = "=SUM(%s!F%s - %s!E%s)" % (self.day_of_week[self.i], self.row + formula_row_i,
                                                         self.day_of_week[self.i], self.row + formula_row_i)
                    add_this = (move_off, move_back, move_route, formula)
                    self.movesarray.append(add_this)
                    formula_row_i += 1  # increment the row in the formula after each moves_set
                i += 1  # increment i

    def display_moves(self):
        """ fill the mv off, mv on and route columns. """
        for move_set in self.movesarray:
            for move_cell in range(4):
                move = move_set[move_cell]
                cell = self.ws_list[self.i].cell(row=self.row + self.move_i, column=5 + move_cell)
                if move_cell in (0, 1):  # format move times as floats or empty strings
                    cell.value = Convert(move).empty_not_zerofloat()  # insert an iteration of self.movesarray
                else:  # do not alter route or formula elements of move sets
                    cell.value = move  # insert an iteration of self.movesarray
                cell.style = self.input_s  # assign worksheet style for MV off, MV on and Route
                if move_cell == 3:
                    cell.style = self.calcs  # use alternate style for Moves Total
                if move_cell != 2:  # do not apply to routes column
                    cell.number_format = "#,###.00;[RED]-#,###.00"
            self.move_i += 1
        self.move_i -= 1  # correction

    def display_formulas_non(self):
        """ fill the formulas columns for non list carriers. """
        ot_formula = "=IF(%s!B%s =\"ns day\", %s!C%s, MAX(%s!C%s - 8, 0))" \
                     % (self.day_of_week[self.i], str(self.row), self.day_of_week[self.i], str(self.row),
                        self.day_of_week[self.i], str(self.row))
        if self.pref[self.lsi] == "nl":  # use alternate formula for non list carriers
            ot_formula = "=IF(%s!B%s =\"ns day\",%s!C%s," \
                         "IF(%s!C%s <= 8 + reference!C$3, 0, MAX(%s!C%s - 8, 0)))" \
                         % (self.day_of_week[self.i], str(self.row), self.day_of_week[self.i], str(self.row),
                            self.day_of_week[self.i], str(self.row), self.day_of_week[self.i], str(self.row))
        off_rt_formula = "=%s!H%s" % (self.day_of_week[self.i], str(self.row))  # copy data from column H/ MV total
        ot_off_rt_formula = "=IF(%s!C%s=\"\",0, " \
                            "IF(OR(%s!B%s=\"ns day\",%s!J%s>=%s!C%s),%s!C%s, " \
                            "IF(%s!C%s<=8+reference!C4,0, " \
                            "MIN(MAX(%s!C%s-8,0), " \
                            "IF(%s!J%s<=reference!C4,0,%s!J%s)))))" \
                            % (self.day_of_week[self.i], str(self.row), self.day_of_week[self.i], str(self.row),
                               self.day_of_week[self.i], str(self.row), self.day_of_week[self.i], str(self.row),
                               self.day_of_week[self.i], str(self.row), self.day_of_week[self.i], str(self.row),
                               self.day_of_week[self.i], str(self.row), self.day_of_week[self.i], str(self.row),
                               self.day_of_week[self.i], str(self.row))
        formulas = (ot_formula, off_rt_formula, ot_off_rt_formula)
        column_i = 0
        for formula in formulas:
            cell = self.ws_list[self.i].cell(row=self.row, column=9 + column_i)
            cell.value = formula  # insert an iteration of formulas
            cell.style = self.calcs  # assign worksheet style
            cell.number_format = "#,###.00;[RED]-#,###.00"
            column_i += 1

    def display_formulas_ot(self):
        """ fill the formula carrier for otdl carriers. """
        max_hrs = 12  # maximum hours for otdl carriers
        if self.pref[self.lsi] == "aux":  # alter formula by list preference
            max_hrs = 11.5  # maximux hours for auxiliary carriers
        formula_ten = "=IF(OR(%s!B%s = \"light\", %s!B%s = \"excused\", %s!B%s = \"sch chg\", " \
                      "%s!B%s = \"annual\", %s!B%s = \"sick\", %s!C%s >= 10 - reference!C5), 0, " \
                      "IF(%s!B%s = \"no call\", 10, " \
                      "IF(%s!C%s = 0, 0, MAX(10 - %s!C%s, 0))))" % \
                      (self.day_of_week[self.i], str(self.row), self.day_of_week[self.i], str(self.row),
                       self.day_of_week[self.i], str(self.row), self.day_of_week[self.i], str(self.row),
                       self.day_of_week[self.i], str(self.row), self.day_of_week[self.i], str(self.row),
                       self.day_of_week[self.i], str(self.row), self.day_of_week[self.i], str(self.row),
                       self.day_of_week[self.i], str(self.row))
        formula_max = "=IF(OR(%s!B%s = \"light\",%s!B%s = \"excused\", %s!B%s = \"sch chg\", %s!B%s = \"annual\", " \
                      "%s!B%s = \"sick\", %s!C%s >= %s - reference!C5), 0, IF(%s!B%s = \"no call\", %s, " \
                      "IF(%s!C%s = 0, 0, MAX(%s - %s!C%s, 0))))" % \
                      (self.day_of_week[self.i], str(self.row), self.day_of_week[self.i], str(self.row),
                       self.day_of_week[self.i], str(self.row), self.day_of_week[self.i], str(self.row),
                       self.day_of_week[self.i], str(self.row), self.day_of_week[self.i], str(self.row),
                       max_hrs, self.day_of_week[self.i], str(self.row),
                       max_hrs, self.day_of_week[self.i], str(self.row),
                       max_hrs, self.day_of_week[self.i], str(self.row))
        formulas = (formula_ten, formula_max)
        column_i = 0
        for formula in formulas:
            cell = self.ws_list[self.i].cell(row=self.row, column=5 + column_i)
            cell.value = formula  # insert an iteration formulas
            cell.style = self.calcs  # assign worksheet style
            cell.number_format = "#,###.00;[RED]-#,###.00"
            column_i += 1

    def build_footer(self):
        """ call the footer depending on the list. """
        if self.pref[self.lsi] == "nl":
            self.nl_footer()
        elif self.pref[self.lsi] == "wal":
            self.wal_footer()
        elif self.pref[self.lsi] == "otdl":
            self.otdl_footer()
        else:
            self.aux_footer()
            
    def nl_footer(self):
        """ build the non list footer. """
        self.row += 1
        cell = self.ws_list[self.i].cell(row=self.row, column=8)  # totals for no list overtime
        cell.value = "Total NL Overtime"
        cell.style = self.col_header
        cell = self.ws_list[self.i].cell(row=self.row, column=9)  # OT
        formula = "=SUM(%s!I%s:I%s)" % (self.day_of_week[self.i], self.first_row, self.last_row)
        cell.value = formula
        cell.style = self.calcs
        cell.number_format = "#,###.00;[RED]-#,###.00"
        location_nl_totals = (self.day_of_week[self.i], "I", self.row)  # save location for totals after wal
        self.mandates_own_route.append(location_nl_totals)  # collect totals for summary
        self.row += 2
        cell = self.ws_list[self.i].cell(row=self.row, column=10)  # totals for no list mandates
        cell.value = "Total NL Mandates"
        cell.style = self.col_header
        cell = self.ws_list[self.i].cell(row=self.row, column=11)  # OT off route
        formula = "=SUM(%s!K%s:K%s)" % (self.day_of_week[self.i], self.first_row, self.last_row)
        cell.value = formula
        cell.style = self.calcs
        cell.number_format = "#,###.00;[RED]-#,###.00"
        self.subtotal_loc_holder.append(self.row)  # collect subtotal location for total after wal
        self.row += 1

    def wal_footer(self):
        """ build the work assignment list footer. """
        self.row += 1
        cell = self.ws_list[self.i].cell(row=self.row, column=10)
        cell.value = "Total WAL Mandates"
        cell.style = self.col_header
        formula = "=SUM(%s!K%s:K%s)" % (self.day_of_week[self.i], self.first_row, self.last_row)
        cell = self.ws_list[self.i].cell(row=self.row, column=11)
        cell.value = formula  # OT off route
        cell.style = self.calcs
        cell.number_format = "#,###.00;[RED]-#,###.00"
        self.subtotal_loc_holder.append(self.row)  # collect subtotal location for total after wal
        self.row += 2
        formula = "=SUM(%s!K%s + %s!K%s)" % (self.day_of_week[self.i], self.subtotal_loc_holder[0],
                                             self.day_of_week[self.i], self.subtotal_loc_holder[1])
        cell = self.ws_list[self.i].cell(row=self.row, column=10)
        cell.value = "Total Mandates"
        cell.style = self.col_header
        cell = self.ws_list[self.i].cell(row=self.row, column=11)
        cell.value = formula  # total ot off route for nl and wal
        cell.style = self.calcs
        cell.number_format = "#,###.00;[RED]-#,###.00"
        add_this = (self.day_of_week[self.i], "K", self.row)
        self.mandates_all.append(add_this)
        self.subtotal_loc_holder = []  # empty out the subtotal location holder for future use with otdl/aux
        self.row += 1

    def otdl_footer(self):
        """ build the ot desired list footer. """
        self.row += 1
        cell = self.ws_list[self.i].cell(row=self.row, column=4)  # header
        cell.value = "Total OTDL Availability"
        cell.style = self.col_header
        cell = self.ws_list[self.i].cell(row=self.row, column=5)  # availability to 10
        formula = "=SUM(%s!E%s:E%s)" % (self.day_of_week[self.i], self.first_row, self.last_row)
        cell.value = formula
        cell.style = self.calcs
        cell.number_format = "#,###.00;[RED]-#,###.00"
        cell = self.ws_list[self.i].cell(row=self.row, column=6)  # availability to 12
        formula = "=SUM(%s!F%s:F%s)" % (self.day_of_week[self.i], self.first_row, self.last_row)
        cell.value = formula
        cell.style = self.calcs
        cell.number_format = "#,###.00;[RED]-#,###.00"
        self.subtotal_loc_holder.append(self.row)  # collect subtotal location for total after aux
        self.row += 1

    def aux_footer(self):
        """ build the auxiliary list footer. """
        self.row += 1
        cell = self.ws_list[self.i].cell(row=self.row, column=4)
        cell.value = "Total AUX Availability"
        cell.style = self.col_header
        formula = "=SUM(%s!E%s:E%s)" % (self.day_of_week[self.i], self.first_row, self.last_row)
        cell = self.ws_list[self.i].cell(row=self.row, column=5)
        cell.value = formula  # availability to 10
        cell.style = self.calcs
        cell.number_format = "#,###.00;[RED]-#,###.00"
        formula = "=SUM(%s!F%s:F%s)" % (self.day_of_week[self.i], self.first_row, self.last_row)
        cell = self.ws_list[self.i].cell(row=self.row, column=6)
        cell.value = formula  # availability to 11.5
        cell.style = self.calcs
        cell.number_format = "#,###.00;[RED]-#,###.00"
        self.subtotal_loc_holder.append(self.row)  # collect subtotal location for total after aux
        self.row += 2
        cell = self.ws_list[self.i].cell(row=self.row, column=4)
        cell.value = "Total Availability"
        cell.style = self.col_header
        formula = "=SUM(%s!E%s + %s!E%s)" % (self.day_of_week[self.i], self.subtotal_loc_holder[0],
                                             self.day_of_week[self.i], self.subtotal_loc_holder[1])
        cell = self.ws_list[self.i].cell(row=self.row, column=5)
        cell.value = formula  # availability to 10
        cell.style = self.calcs
        cell.number_format = "#,###.00;[RED]-#,###.00"
        formula = "=SUM(%s!F%s + %s!F%s)" % (self.day_of_week[self.i], self.subtotal_loc_holder[0],
                                             self.day_of_week[self.i], self.subtotal_loc_holder[1])
        cell = self.ws_list[self.i].cell(row=self.row, column=6)
        cell.value = formula  # availability to 11.5
        cell.style = self.calcs
        cell.number_format = "#,###.00;[RED]-#,###.00"
        add_this = (self.day_of_week[self.i], "E", self.row)  # location of total availability to 10
        self.availability_10.append(add_this)  # collect location of totals for summary, put in array
        add_this = (self.day_of_week[self.i], "F", self.row)  # location of total availability to max
        self.availability_max.append(add_this)  # collect location of totals for summary, put in array
        self.subtotal_loc_holder = []  # empty out the subtotal location holder for future use with otdl/aux
        self.row += 1

    def pagebreak(self):
        """ create a page break if consistant with user preferences """
        if self.pref[self.lsi] == "nl" and not self.pb_nl_wal:
            self.row += 1
            return
        if self.pref[self.lsi] == "wal" and not self.pb_wal_otdl:
            self.row += 1
            return
        if self.pref[self.lsi] == "otdl" and not self.pb_otdl_aux:
            self.row += 1
            return
        if self.pref[self.lsi] == "aux":
            self.row += 1
            return
        try:
            self.ws_list[self.i].page_breaks.append(Break(id=self.row))
        except AttributeError:
            self.ws_list[self.i].row_breaks.append(Break(id=self.row))  # effective for windows
        self.row += 1

    def build_summary_header(self):
        """ summary headers """
        self.pbi += 1
        self.pb.move_count(self.pbi)  # increment progress bar
        self.pb.change_text("Building day Summary...")
        self.summary['A1'] = "Improper Mandate Worksheet"
        self.summary['A1'].style = self.ws_header
        self.summary.merge_cells('A1:E1')
        self.summary['B3'] = "Summary Sheet"
        self.summary['B3'].style = self.date_dov_title
        self.summary['A5'] = "Pay Period:  "
        self.summary['A5'].style = self.date_dov_title
        self.summary['B5'] = projvar.pay_period
        self.summary['B5'].style = self.date_dov
        self.summary.merge_cells('B5:D5')
        self.summary['A6'] = "Station:  "
        self.summary['A6'].style = self.date_dov_title
        self.summary['B6'] = projvar.invran_station
        self.summary['B6'].style = self.date_dov
        # reference page has no header
        
    def build_summary(self):
        """ build the summary page. """
        self.summary['A1'] = "Improper Mandate Worksheet"
        self.summary['A1'].style = self.ws_header
        self.summary.merge_cells('A1:E1')
        self.summary['B3'] = "Summary Sheet"
        self.summary['B3'].style = self.date_dov_title
        self.summary['A5'] = "Pay Period:  "
        self.summary['A5'].style = self.date_dov_title
        self.summary['B5'] = projvar.pay_period
        self.summary['B5'].style = self.date_dov
        self.summary.merge_cells('B5:D5')
        self.summary['A6'] = "Station:  "
        self.summary['A6'].style = self.date_dov_title
        self.summary['B6'] = projvar.invran_station
        self.summary['B6'].style = self.date_dov
        self.summary.merge_cells('B6:D6')
        self.summary['B8'] = "Availability"
        self.summary['B8'].style = self.date_dov_title
        self.summary['B9'] = "to 10"
        self.summary['B9'].style = self.date_dov_title
        self.summary['C8'] = "No list"
        self.summary['C8'].style = self.date_dov_title
        self.summary['C9'] = "overtime"
        self.summary['C9'].style = self.date_dov_title
        self.summary['D9'] = "violations"
        self.summary['D9'].style = self.date_dov_title
        self.summary['F8'] = "Availability"
        self.summary['F8'].style = self.date_dov_title
        self.summary['F9'] = "to 12"
        self.summary['F9'].style = self.date_dov_title
        self.summary['G8'] = "Off route"
        self.summary['G8'].style = self.date_dov_title
        self.summary['G9'] = "mandates"
        self.summary['G9'].style = self.date_dov_title
        self.summary['H9'] = "violations"
        self.summary['H9'].style = self.date_dov_title
        row = 10
        for i in range(len(self.dates)):
            self.summary['A' + str(row)].value = format(self.dates[i], "%m/%d/%y %a")
            self.summary['A' + str(row)].style = self.date_dov_title
            self.summary['A' + str(row)].number_format = "#,###.00;[RED]-#,###.00"
            location = self.availability_10[i]  # get the location of the total from the worksheet from the array
            formula = "=%s!%s%s" % (location[0], location[1], location[2])
            self.summary['B' + str(row)] = formula
            self.summary['B' + str(row)].style = self.input_s
            self.summary['B' + str(row)].number_format = "#,###.00;[RED]-#,###.00"
            location = self.mandates_own_route[i]  # get the location of the total from the worksheet from the array
            formula = "=%s!%s%s" % (location[0], location[1], location[2])
            self.summary['C' + str(row)] = formula
            self.summary['C' + str(row)].style = self.input_s
            self.summary['C' + str(row)].number_format = "#,###.00;[RED]-#,###.00"
            formula = "=IF(%s!B%s<%s!C%s,%s!B%s,%s!C%s)" \
                      % ('summary', str(row), 'summary', str(row), 'summary',
                         str(row), 'summary', str(row))
            self.summary['D' + str(row)] = formula
            self.summary['D' + str(row)].style = self.calcs
            self.summary['D' + str(row)].number_format = "#,###.00;[RED]-#,###.00"
            location = self.availability_max[i]  # get the location of the total from the worksheet from the array
            formula = "=%s!%s%s" % (location[0], location[1], location[2])
            self.summary['F' + str(row)] = formula
            self.summary['F' + str(row)].style = self.input_s
            self.summary['F' + str(row)].number_format = "#,###.00;[RED]-#,###.00"
            location = self.mandates_all[i]  # get location of total mandates from worksheet from the array
            formula = "=%s!%s%s" % (location[0], location[1], location[2])  # total mandates
            self.summary['G' + str(row)] = formula
            self.summary['G' + str(row)].style = self.input_s
            self.summary['G' + str(row)].number_format = "#,###.00;[RED]-#,###.00"
            formula = "=IF(%s!F%s<%s!G%s,%s!F%s,%s!G%s)" \
                      % ('summary', str(row), 'summary', str(row), 'summary',
                         str(row), 'summary', str(row))
            self.summary['H' + str(row)] = formula
            self.summary['H' + str(row)].style = self.calcs
            self.summary['H' + str(row)].number_format = "#,###.00;[RED]-#,###.00"
            row += 2
        
    def save_open(self):
        """ name and open the excel file """
        self.pbi += 1
        self.pb.move_count(self.pbi)  # increment progress bar
        self.pb.change_text("Saving...")
        self.pb.stop()
        r = "_w"
        if not projvar.invran_weekly_span:  # if investigation range is daily
            r = "_d"
        xl_filename = "kb" + str(format(self.dates[0], "_%y_%m_%d")) + r + ".xlsx"
        try:
            self.wb.save(dir_path('spreadsheets') + xl_filename)
            messagebox.showinfo("Spreadsheet generator",
                                "Your spreadsheet was successfully generated. \n"
                                "File is named: {}".format(xl_filename),
                                parent=self.frame)
            if sys.platform == "win32":
                os.startfile(dir_path('spreadsheets') + xl_filename)
            if sys.platform == "linux":
                subprocess.call(["xdg-open", 'kb_sub/spreadsheets/' + xl_filename])
            if sys.platform == "darwin":
                subprocess.call(["open", dir_path('spreadsheets') + xl_filename])
        except PermissionError:
            messagebox.showerror("Spreadsheet generator",
                                 "The spreadsheet was not opened. \n"
                                 "Suggestion: "
                                 "Make sure that identically named spreadsheets are closed "
                                 "(the file can't be overwritten while open).",
                                 parent=self.frame)


class OvermaxSpreadsheet:
    """
    This generates the 12 and 60 hour violations worksheet. This spreadsheeet is a klusterbox original and is the
    most comprehensive spreadsheet of its kind.
    """
    def __init__(self):
        self.frame = None
        self.pb = None  # progress bar object
        self.pbi = 0  # progress bar count index
        self.carrier_list = []
        self.wb = None  # workbook object
        self.violations = None  # workbook object sheet
        self.instructions = None  # workbook object sheet
        self.summary = None  # workbook object sheet
        self.startdate = None  # start of the investigation range
        self.enddate = None  # ending day of the investigation range
        self.dates = []  # all the dates of the investigation range
        self.rings = []  # all rings for all carriers in the carrier list
        self.min_rows = 0  # the minimum number of rows set by user preferences
        self.non_otdl_violation = 11.5  # Hours that non-otdl carriers work before 12/60 violation
        self.wal_12hour = None
        self.wal_12hr_mod = ""  # text inserted into formulas which varies depending on wal_12hour setting
        self.wal_dec_exempt = None  # work assignment list december exemption - true or false
        self.wal_dec_exempt_mod = ""  # text inserted into formulas which varies depending on wal_dec_exempt setting
        self.ws_header = None  # style
        self.date_dov = None  # style
        self.date_dov_title = None  # style
        self.col_header = None  # style
        self.col_center_header = None  # style
        self.vert_header = None  # style
        self.input_name = None  # style
        self.input_s = None  # style
        self.calcs = None  # style
        self.vert_calcs = None  # style
        self.instruct_text = None  # style
        self.violation_recsets = []  # carrier info, daily hours, leavetypes and leavetimes

    def create(self, frame):
        """ master method for calling methods"""
        self.frame = frame
        if not self.ask_ok():
            return
        self.pb = ProgressBarDe(label="Building Improper Mandates Spreadsheet")
        self.pb.max_count(100)  # set length of progress bar
        self.pb.start_up()  # start the progress bar
        self.pbi = 1
        self.pb.move_count(self.pbi)  # increment progress bar
        self.pb.change_text("Gathering Data... ")
        self.get_dates()
        self.get_carrierlist()
        self.get_rings()
        self.get_minrows()
        self.set_wal12hrmod()
        self.set_waldecexempt()
        self.get_styles()
        self.build_workbook()
        self.set_dimensions()
        self.build_summary()
        self.build_violations()
        self.build_instructions()
        self.violated_recs()
        self.get_pb_len()
        self.show_violations()
        self.save_open()

    def ask_ok(self):
        """ continue if user selects ok. """
        if messagebox.askokcancel("Spreadsheet generator",
                                  "Do you want to generate an Over Max Spreadsheet for violations "
                                  "of the 12 and 60 Hour Rule?",
                                  parent=self.frame):
            return True
        return False

    def get_dates(self):
        """ get the dates of the investigation range from the project variables. """
        date = projvar.invran_date_week[0]
        self.startdate = projvar.invran_date_week[0]
        self.enddate = projvar.invran_date_week[6]
        for _ in range(7):
            self.dates.append(date)
            date += timedelta(days=1)

    def get_carrierlist(self):
        """ call the carrierlist class from kbtoolbox module to get the carrier list """
        carrierlist = CarrierList(self.startdate, self.enddate, projvar.invran_station).get()
        for carrier in carrierlist:
            self.carrier_list.append(carrier[0])  # add the first record for each carrier in rec set

    def get_rings(self):
        """ get clock rings from the rings table """
        sql = "SELECT * FROM rings3 WHERE rings_date BETWEEN '%s' AND '%s' ORDER BY rings_date, carrier_name" \
              % (projvar.invran_date_week[0], projvar.invran_date_week[6])
        self.rings = inquire(sql)

    def get_minrows(self):
        """ get minimum rows """
        sql = "SELECT tolerance FROM tolerances"
        result = inquire(sql)
        self.min_rows = int(result[14][0])
        #  get the work assignment list 12 hour violation setting -
        #  if True: violation occurs after 12 hours.
        #  if False: violation occurs after 11.50 hours
        self.wal_12hour = Convert(result[44][0]).str_to_bool()
        #  get the work assignment list december exemption setting -
        #  if True: treat wal carriers same as otdl carriers for december exemption.
        self.wal_dec_exempt = Convert(result[45][0]).str_to_bool()

    def set_wal12hrmod(self):
        """ if the wal_12hour setting is True, don't include 'wal' in formula"""
        self.wal_12hr_mod = "nl"  #
        if not self.wal_12hour:
            self.wal_12hr_mod = "wal"

    def set_waldecexempt(self):
        """ if the wal_dec_exempt is True, include 'wal' in formula for total violations.
        This will treat wal carriers same as otdl carriers durning december."""
        self.wal_dec_exempt_mod = "otdl"
        if self.wal_dec_exempt:
            self.wal_dec_exempt_mod = "wal"

    def get_styles(self):
        """ Named styles for workbook """
        bd = Side(style='thin', color="80808080")  # defines borders
        self.ws_header = NamedStyle(name="ws_header", font=Font(bold=True, name='Arial', size=12))
        self.date_dov = NamedStyle(name="date_dov", font=Font(name='Arial', size=8))
        self.date_dov_title = NamedStyle(name="date_dov_title", font=Font(bold=True, name='Arial', size=8),
                                         alignment=Alignment(horizontal='right'))
        self.col_header = NamedStyle(name="col_header", font=Font(bold=True, name='Arial', size=8),
                                     border=Border(left=bd, right=bd, top=bd, bottom=bd),
                                     alignment=Alignment(horizontal='left'))
        self.col_center_header = NamedStyle(name="col_center_header", font=Font(bold=True, name='Arial', size=8),
                                            alignment=Alignment(horizontal='center'),
                                            border=Border(left=bd, right=bd, top=bd, bottom=bd))
        self.vert_header = NamedStyle(name="vert_header", font=Font(bold=True, name='Arial', size=8),
                                      border=Border(left=bd, right=bd, top=bd, bottom=bd),
                                      alignment=Alignment(horizontal='right', text_rotation=90))
        self.input_name = NamedStyle(name="input_name", font=Font(name='Arial', size=8),
                                     border=Border(left=bd, right=bd, top=bd, bottom=bd),
                                     alignment=Alignment(horizontal='left'))
        self.input_s = NamedStyle(name="input_s", font=Font(name='Arial', size=8),
                                  border=Border(left=bd, right=bd, top=bd, bottom=bd),
                                  alignment=Alignment(horizontal='right'))
        self.calcs = NamedStyle(name="calcs", font=Font(name='Arial', size=8),
                                border=Border(left=bd, right=bd, top=bd, bottom=bd),
                                fill=PatternFill(fgColor='e5e4e2', fill_type='solid'),
                                alignment=Alignment(horizontal='right'))
        self.vert_calcs = NamedStyle(name="vert_calcs", font=Font(name='Arial', size=8),
                                     border=Border(left=bd, right=bd, top=bd, bottom=bd),
                                     fill=PatternFill(fgColor='e5e4e2', fill_type='solid'),
                                     alignment=Alignment(horizontal='right', text_rotation=90))
        self.instruct_text = NamedStyle(name="instruct_text", font=Font(name='Arial', size=8),
                                        alignment=Alignment(horizontal='left', vertical='top'))

    def build_workbook(self):
        """ creates the workbook object """
        self.pb.change_text("Building workbook...")
        self.wb = Workbook()  # define the workbook
        self.violations = self.wb.active  # create first worksheet
        self.violations.title = "violations"  # title first worksheet
        self.violations.oddFooter.center.text = "&A"
        self.summary = self.wb.create_sheet("summary")
        self.summary.oddFooter.center.text = "&A"
        self.instructions = self.wb.create_sheet("instructions")

    def set_dimensions(self):
        """ adjust the height and width on the violations/ instructions page """
        for x in range(2, 10):
            self.violations.row_dimensions[x].height = 10  # adjust all row height
        sheets = (self.violations, self.instructions)
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
            sheet.column_dimensions["X"].width = 6

    def build_summary(self):
        """ summary worksheet - format cells """
        self.pb.change_text("Building summary")
        self.summary.merge_cells('A1:R1')
        self.summary['A1'] = "12 and 60 Hour Violations Summary"
        self.summary['A1'].style = self.ws_header
        self.summary.column_dimensions["A"].width = 15
        self.summary.column_dimensions["B"].width = 8
        self.summary['A3'] = "Date:"
        self.summary['A3'].style = self.date_dov_title
        self.summary.merge_cells('B3:D3')  # blank field for date
        self.summary['B3'] = self.dates[0].strftime("%x") + " - " + self.dates[6].strftime("%x")
        self.summary['B3'].style = self.date_dov
        self.summary.merge_cells('K3:N3')
        self.summary['F3'] = "Pay Period:"  # Pay Period Header
        self.summary['F3'].style = self.date_dov_title
        self.summary.merge_cells('G3:I3')  # blank field for pay period
        self.summary['G3'] = projvar.pay_period
        self.summary['G3'].style = self.date_dov
        self.summary['A4'] = "Station:"  # Station Header
        self.summary['A4'].style = self.date_dov_title
        self.summary.merge_cells('B4:D4')  # blank field for station
        self.summary['B4'] = projvar.invran_station
        self.summary['B4'].style = self.date_dov
        self.summary['A6'] = "name"
        self.summary['A6'].style = self.col_center_header
        self.summary['B6'] = "violation"
        self.summary['B6'].style = self.col_center_header

    def build_violations(self):
        """ self.violations worksheet - format cells """
        self.pb.change_text("Building violations...")
        self.violations.merge_cells('A1:R1')
        self.violations['A1'] = "12 and 60 Hour Violations Worksheet"
        self.violations['A1'].style = self.ws_header
        self.violations['A3'] = "Date:"
        self.violations['A3'].style = self.date_dov_title
        self.violations.merge_cells('B3:J3')  # blank field for date
        self.violations['B3'] = self.dates[0].strftime("%x") + " - " + self.dates[6].strftime("%x")
        self.violations['B3'].style = self.date_dov
        self.violations.merge_cells('K3:N3')
        self.violations['K3'] = "Pay Period:"
        self.violations['k3'].style = self.date_dov_title
        self.violations.merge_cells('O3:S3')  # blank field for pay period
        self.violations['O3'] = projvar.pay_period
        self.violations['O3'].style = self.date_dov
        self.violations['A4'] = "Station:"
        self.violations['A4'].style = self.date_dov_title
        self.violations.merge_cells('B4:J4')  # blank field for station
        self.violations['B4'] = projvar.invran_station
        self.violations['B4'].style = self.date_dov
        self.violations.merge_cells('D7:Q7')
        self.violations.merge_cells('K4:N4')
        self.violations['K4'] = "Dec Exception:"  # December exception
        self.violations['k4'].style = self.date_dov_title
        self.violations.merge_cells('O4:S4')  # blank field for pay period
        self.violations['O4'] = "no"  # enter yes or no
        self.violations['O4'].style = self.date_dov
        self.violations['D7'] = "Daily Paid Leave times with type"
        self.violations['D7'].style = self.col_center_header
        self.violations.merge_cells('D8:Q8')
        self.violations['D8'] = "Daily 5200 times"
        self.violations['D8'].style = self.col_center_header
        self.violations['A9'] = "name"
        self.violations['A9'].style = self.col_header
        self.violations['B9'] = "list"
        self.violations['B9'].style = self.col_header
        self.violations.merge_cells('C6:C9')
        self.violations['C6'] = "Weekly\n5200"
        self.violations['C6'].style = self.vert_header
        self.violations.merge_cells('D9:E9')
        self.violations['D9'] = "sat"
        self.violations['D9'].style = self.col_center_header
        self.violations.merge_cells('F9:G9')
        self.violations['F9'] = "sun"
        self.violations['F9'].style = self.col_center_header
        self.violations.merge_cells('H9:I9')
        self.violations['H9'] = "mon"
        self.violations['H9'].style = self.col_center_header
        self.violations.merge_cells('J9:K9')
        self.violations['J9'] = "tue"
        self.violations['J9'].style = self.col_center_header
        self.violations.merge_cells('L9:M9')
        self.violations['L9'] = "wed"
        self.violations['L9'].style = self.col_center_header
        self.violations.merge_cells('N9:O9')
        self.violations['N9'] = "thr"
        self.violations['N9'].style = self.col_center_header
        self.violations.merge_cells('P9:Q9')
        self.violations['P9'] = "fri"
        self.violations['P9'].style = self.col_center_header
        self.violations.merge_cells('S5:S9')
        self.violations['S5'] = " Weekly\nViolation"
        self.violations['S5'].style = self.vert_header
        self.violations.merge_cells('T5:T9')
        self.violations['T5'] = "Daily\nViolation"
        self.violations['T5'].style = self.vert_header
        self.violations.merge_cells('U5:U9')
        self.violations['U5'] = "Wed Adj"
        self.violations['U5'].style = self.vert_header
        self.violations.merge_cells('V5:V9')
        self.violations['V5'] = "Thr Adj"
        self.violations['V5'].style = self.vert_header
        self.violations.merge_cells('W5:W9')
        self.violations['W5'] = "Fri Adj"
        self.violations['W5'].style = self.vert_header
        self.violations.merge_cells('X5:X9')
        self.violations['X5'] = "Total\nViolation"
        self.violations['X5'].style = self.vert_header

    def build_instructions(self):
        """ format the instructions cells """
        self.pb.change_text("Building instructions")
        self.instructions.merge_cells('A1:R1')
        self.instructions['A1'] = "12 and 60 Hour Violations Instructions"
        self.instructions['A1'].style = self.ws_header
        self.instructions.row_dimensions[3].height = 300
        self.instructions['A3'].style = self.instruct_text
        self.instructions.merge_cells('A3:X3')
        self.instructions['A3'] = \
            "Caution for Mac Users: \n" \
            "Using the Apple Numbers Spreadsheet program is not recommended. Apple Numbers " \
            "does not support vertical text or hidden fields, both of which are used in the " \
            "12 and 60 Hour Violations Spreadsheet. If you are using Mac, you can download " \
            "Libre Office Calc, which is recommended, for free. Microsoft Excel or Google Docs " \
            "will also work properly. \n\n" \
            "December Exemption Setting: \n" \
            "Enter \"yes\" in this cell (use lowercase only) to exempt otdl carriers from " \
            "violations during the month of December. The default is \"no\". " \
            "Turning WAL December Exemption to \'on\' in Spreadsheet Setting in Klusterbox " \
            "will modify the formulas to include \"wal\" carriers in the exemption.\n" \
            "\tWAL 12 Hour Violation Setting is {}\n\n" \
            "Instructions: \n" \
            "1. Fill in the name \n" \
            "2. Fill in the list. Enter either \"otdl\",\"wal\",\"nl\",\"aux\" or \"ptf\" in list " \
            "columns. Use only lowercase. \n" \
            "   If you do not enter anything, the default is \"otdl\". \n" \
            "\totdl = overtime desired list\n" \
            "\twal = work assignment list\n" \
            "\tnl = no list \n" \
            "\taux = auxiliary (this would be a cca or city carrier assistant).\n" \
            "\tptf = part time flexible \n" \
            "3. Fill in the weekly 5200 time in field C if it exceeds 60 hours " \
            "or if the sum of all daily non 5200 times (all fields D) plus \n" \
            "   the weekly 5200 time (field C) will  exceed 60 hours.\n" \
            "4. Fill in any daily non 5200 times and types in fields D and E. " \
            "Enter only paid leave types such as sick leave, annual\n" \
            "   leave and holiday leave. Do not enter unpaid leave types such as LWOP " \
            "(leave without pay) or AWOL (absent \n" \
            "   without leave).\n" \
            "5. Fill in any daily 5200 times which exceed 12 hours for otdl carriers " \
            "or 11.50 hours for any other carrier in fields F.\n" \
            "   Failing to fill out the daily values for Wednesday, Thursday and Friday " \
            "could cause errors in calculating the adjustments,\n" \
            "   so fill those in.\n" \
            "6. The gray fields will fill automatically. Do not enter an information in " \
            "these fields as it will delete the formulas.\n" \
            "7. Field O will show the violation in hours which you should seek a remedy " \
            "for. \n".format(Convert(self.wal_12hour).bool_to_onoff())
        self.instructions['A3'].alignment = Alignment(wrap_text=True, vertical='top')
        for x in range(4, 20):
            self.instructions.row_dimensions[x].height = 10  # adjust all row height
        self.instructions.merge_cells('D6:Q6')
        self.instructions['D6'] = "Daily Paid Leave times with type"
        self.instructions['D6'].style = self.col_center_header
        self.instructions.merge_cells('D7:Q7')
        self.instructions['D7'] = "Daily 5200 times"
        self.instructions['D7'].style = self.col_center_header
        self.instructions['A8'] = "name"
        self.instructions['A8'].style = self.col_header
        self.instructions['B8'] = "list"
        self.instructions['B8'].style = self.col_header
        self.instructions.merge_cells('C5:C8')
        self.instructions['C5'] = "Weekly\n5200"
        self.instructions['C5'].style = self.vert_header
        self.instructions.merge_cells('D8:E8')
        self.instructions['D8'] = "sat"
        self.instructions['D8'].style = self.col_center_header
        self.instructions.merge_cells('F8:G8')
        self.instructions['F8'] = "sun"
        self.instructions['F8'].style = self.col_center_header
        self.instructions.merge_cells('H8:I8')
        self.instructions['H8'] = "mon"
        self.instructions['H8'].style = self.col_center_header
        self.instructions.merge_cells('J8:K8')
        self.instructions['J8'] = "tue"
        self.instructions['J8'].style = self.col_center_header
        self.instructions.merge_cells('L8:M8')
        self.instructions['L8'] = "wed"
        self.instructions['L8'].style = self.col_center_header
        self.instructions.merge_cells('N8:O8')
        self.instructions['N8'] = "thr"
        self.instructions['N8'].style = self.col_center_header
        self.instructions.merge_cells('P8:Q8')
        self.instructions['P8'] = "fri"
        self.instructions['P8'].style = self.col_center_header
        self.instructions.merge_cells('S4:S8')
        self.instructions['S4'] = " Weekly\nViolation"
        self.instructions['S4'].style = self.vert_header
        self.instructions.merge_cells('T4:T8')
        self.instructions['T4'] = "Daily\nViolation"
        self.instructions['T4'].style = self.vert_header
        self.instructions.merge_cells('U4:U8')
        self.instructions['U4'] = "Wed Adj"
        self.instructions['U4'].style = self.vert_header
        self.instructions.merge_cells('V4:V8')
        self.instructions['V4'] = "Thr Adj"
        self.instructions['V4'].style = self.vert_header
        self.instructions.merge_cells('W4:W8')
        self.instructions['W4'] = "Fri Adj"
        self.instructions['W4'].style = self.vert_header
        self.instructions.merge_cells('X4:X8')
        self.instructions['X4'] = "Total\nViolation"
        self.instructions['X4'].style = self.vert_header
        self.instructions['A9'] = "A"
        self.instructions['A9'].style = self.col_center_header
        self.instructions['B9'] = "B"
        self.instructions['B9'].style = self.col_center_header
        self.instructions['C9'] = "C"
        self.instructions['C9'].style = self.col_center_header
        self.instructions['D9'] = "D"
        self.instructions['D9'].style = self.col_center_header
        self.instructions['E9'] = "E"
        self.instructions['E9'].style = self.col_center_header
        self.instructions['F9'] = "G"
        self.instructions['F9'].style = self.col_center_header
        self.instructions.merge_cells('F9:G9')
        self.instructions['H9'] = "D"
        self.instructions['H9'].style = self.col_center_header
        self.instructions['I9'] = "E"
        self.instructions['I9'].style = self.col_center_header
        self.instructions['J9'] = "D"
        self.instructions['J9'].style = self.col_center_header
        self.instructions['K9'] = "E"
        self.instructions['K9'].style = self.col_center_header
        self.instructions['L9'] = "D"
        self.instructions['L9'].style = self.col_center_header
        self.instructions['M9'] = "E"
        self.instructions['M9'].style = self.col_center_header
        self.instructions['N9'] = "D"
        self.instructions['N9'].style = self.col_center_header
        self.instructions['O9'] = "E"
        self.instructions['O9'].style = self.col_center_header
        self.instructions['P9'] = "D"
        self.instructions['P9'].style = self.col_center_header
        self.instructions['Q9'] = "E"
        self.instructions['Q9'].style = self.col_center_header
        self.instructions['S9'] = "J"
        self.instructions['S9'].style = self.col_center_header
        self.instructions['T9'] = "K"
        self.instructions['T9'].style = self.col_center_header
        self.instructions['U9'] = "L"
        self.instructions['U9'].style = self.col_center_header
        self.instructions['V9'] = "M"
        self.instructions['V9'].style = self.col_center_header
        self.instructions['W9'] = "N"
        self.instructions['W9'].style = self.col_center_header
        self.instructions['X9'] = "O"
        self.instructions['X9'].style = self.col_center_header
        i = 10
        # instructions name
        self.instructions.merge_cells('A' + str(i) + ':A' + str(i + 1))  # merge box for name
        self.instructions['A10'] = "kubrick, s"
        self.instructions['A10'].style = self.input_name
        self.instructions['A11'].style = self.input_name
        # instructions list
        self.instructions.merge_cells('B' + str(i) + ':B' + str(i + 1))  # merge box for list type input
        self.instructions['B10'] = "wal"
        self.instructions['B10'].style = self.input_s
        self.instructions['B11'].style = self.input_s
        # instructions weekly
        self.instructions.merge_cells('C' + str(i) + ':C' + str(i + 1))  # merge box for weekly input
        self.instructions['C10'] = 75.00
        self.instructions['C10'].style = self.input_s
        self.instructions['C11'].style = self.input_s
        self.instructions['C10'].number_format = "#,###.00;[RED]-#,###.00"
        # instructions saturday
        self.instructions.merge_cells('D' + str(i + 1) + ':E' + str(i + 1))  # merge box for sat 5200
        self.instructions['D' + str(i)] = ""  # leave time
        self.instructions['D' + str(i)].style = self.input_s
        self.instructions['D' + str(i)].number_format = "#,###.00;[RED]-#,###.00"
        self.instructions['E' + str(i)] = ""  # leave type
        self.instructions['E' + str(i)].style = self.input_s
        self.instructions['D' + str(i + 1)] = 13.00  # 5200 time
        self.instructions['D' + str(i + 1)].style = self.input_s
        self.instructions['D' + str(i + 1)].number_format = "#,###.00;[RED]-#,###.00"
        # instructions sunday
        self.instructions.merge_cells('F' + str(i + 1) + ':G' + str(i + 1))  # merge box for sun 5200
        self.instructions['F' + str(i)] = ""  # leave time
        self.instructions['F' + str(i)].style = self.input_s
        self.instructions['F' + str(i)].number_format = "#,###.00;[RED]-#,###.00"
        self.instructions['G' + str(i)] = ""  # leave type
        self.instructions['G' + str(i)].style = self.input_s
        self.instructions['F' + str(i + 1)] = ""  # 5200 time
        self.instructions['F' + str(i + 1)].style = self.input_s
        self.instructions['F' + str(i + 1)].number_format = "#,###.00;[RED]-#,###.00"
        # instructions monday
        self.instructions.merge_cells('H' + str(i + 1) + ':I' + str(i + 1))  # merge box for mon 5200
        self.instructions['H' + str(i)] = 8  # leave time
        self.instructions['H' + str(i)].style = self.input_s
        self.instructions['H' + str(i)].number_format = "#,###.00;[RED]-#,###.00"
        self.instructions['I' + str(i)] = "h"  # leave type
        self.instructions['I' + str(i)].style = self.input_s
        self.instructions['H' + str(i + 1)] = ""  # 5200 time
        self.instructions['H' + str(i + 1)].style = self.input_s
        self.instructions['H' + str(i + 1)].number_format = "#,###.00;[RED]-#,###.00"
        # instructions tuesday
        self.instructions.merge_cells('J' + str(i + 1) + ':K' + str(i + 1))  # merge box for tue 5200
        self.instructions['J' + str(i)] = ""  # leave time
        self.instructions['J' + str(i)].style = self.input_s
        self.instructions['J' + str(i)].number_format = "#,###.00;[RED]-#,###.00"
        self.instructions['K' + str(i)] = ""  # leave type
        self.instructions['K' + str(i)].style = self.input_s
        self.instructions['J' + str(i + 1)] = 14  # 5200 time
        self.instructions['J' + str(i + 1)].style = self.input_s
        self.instructions['J' + str(i + 1)].number_format = "#,###.00;[RED]-#,###.00"
        # instructions wednesday
        self.instructions.merge_cells('L' + str(i + 1) + ':M' + str(i + 1))  # merge box for wed 5200
        self.instructions['L' + str(i)] = ""  # leave time
        self.instructions['L' + str(i)].style = self.input_s
        self.instructions['L' + str(i)].number_format = "#,###.00;[RED]-#,###.00"
        self.instructions['M' + str(i)] = ""  # leave type
        self.instructions['M' + str(i)].style = self.input_s
        self.instructions['L' + str(i + 1)] = 14  # 5200 time
        self.instructions['L' + str(i + 1)].style = self.input_s
        self.instructions['M' + str(i + 1)].number_format = "#,###.00;[RED]-#,###.00"
        # instructions thursday
        self.instructions.merge_cells('N' + str(i + 1) + ':O' + str(i + 1))  # merge box for thr 5200
        self.instructions['N' + str(i)] = ""  # leave time
        self.instructions['N' + str(i)].style = self.input_s
        self.instructions['N' + str(i)].number_format = "#,###.00;[RED]-#,###.00"
        self.instructions['O' + str(i)] = ""  # leave type
        self.instructions['O' + str(i)].style = self.input_s
        self.instructions['N' + str(i + 1)] = 13  # 5200 time
        self.instructions['N' + str(i + 1)].style = self.input_s
        self.instructions['N' + str(i + 1)].number_format = "#,###.00;[RED]-#,###.00"
        # instructions friday
        self.instructions.merge_cells('P' + str(i + 1) + ':Q' + str(i + 1))  # merge box for fri 5200
        self.instructions['P' + str(i)] = ""  # leave time
        self.instructions['P' + str(i)].style = self.input_s
        self.instructions['P' + str(i)].number_format = "#,###.00;[RED]-#,###.00"
        self.instructions['Q' + str(i)] = ""  # leave type
        self.instructions['Q' + str(i)].style = self.input_s
        self.instructions['P' + str(i + 1)] = 13  # 5200 time
        self.instructions['P' + str(i + 1)].style = self.input_s
        self.instructions['P' + str(i + 1)].number_format = "#,###.00;[RED]-#,###.00"
        # instructions hidden columns
        page = "instructions"
        formula = "=SUM(%s!D%s:P%s)+%s!D%s + %s!H%s + %s!J%s + %s!L%s + " \
                  "%s!N%s + %s!P%s" % (page, str(i + 1), str(i + 1),
                                       page, str(i), page, str(i), page, str(i),
                                       page, str(i), page, str(i), page, str(i))
        self.instructions['R' + str(i)] = formula
        self.instructions['R' + str(i)].style = self.calcs
        self.instructions['R' + str(i)].number_format = "#,###.00;[RED]-#,###.00"
        formula = "=SUM(%s!C%s+%s!D%s+%s!H%s+%s!J%s+%s!L%s+%s!N%s+%s!P%s)" % \
                  (page, str(i), page, str(i), page, str(i),
                   page, str(i), page, str(i), page, str(i),
                   page, str(i))
        self.instructions['R' + str(i + 1)] = formula
        self.instructions['R' + str(i + 1)].style = self.calcs
        self.instructions['R' + str(i + 1)].number_format = "#,###.00;[RED]-#,###.00"
        # instructions weekly self.violations
        self.instructions.merge_cells('S' + str(i) + ':S' + str(i + 1))  # merge box for weekly violation
        formula = "=IF(%s!B%s = \"aux\",0,MAX(IF(%s!R%s>%s!R%s,MAX(%s!R%s-60,0),MAX(%s!R%s-60)),0))" \
                  % (page, str(i), page, str(i), page, str(i + 1), page, str(i),
                     page, str(i + 1))
        self.instructions['S10'] = formula
        self.instructions['S10'].style = self.calcs
        self.instructions['S10'].number_format = "#,###.00;[RED]-#,###.00"
        # instructions daily self.violations
        formula_d = "=IF(OR(%s!B%s=\"%s\",%s!B%s=\"nl\",%s!B%s=\"ptf\",%s!B%s=\"aux\")," \
                    "(SUM(IF(%s!D%s>%s,%s!D%s-%s,0)+IF(%s!H%s>%s,%s!H%s-%s,0)+IF(%s!J%s>%s,%s!J%s-%s,0)" \
                    "+IF(%s!L%s>%s,%s!L%s-%s,0)+IF(%s!N%s>%s,%s!N%s-%s,0)+IF(%s!P%s>%s,%s!P%s-%s,0)))," \
                    "(SUM(IF(%s!D%s>12,%s!D%s-12,0)+IF(%s!H%s>12,%s!H%s-12,0)+IF(%s!J%s>12,%s!J%s-12,0)" \
                    "+IF(%s!L%s>12,%s!L%s-12,0)+IF(%s!N%s>12,%s!N%s-12,0)+IF(%s!P%s>12,%s!P%s-12,0))))" \
                    % (page, str(i), self.wal_12hr_mod,
                       page, str(i), page, str(i), page, str(i),
                       page, str(i + 1), str(self.non_otdl_violation),
                       page, str(i + 1), str(self.non_otdl_violation),
                       page, str(i + 1), str(self.non_otdl_violation),
                       page, str(i + 1), str(self.non_otdl_violation),
                       page, str(i + 1), str(self.non_otdl_violation),
                       page, str(i + 1), str(self.non_otdl_violation),
                       page, str(i + 1), str(self.non_otdl_violation),
                       page, str(i + 1), str(self.non_otdl_violation),
                       page, str(i + 1), str(self.non_otdl_violation),
                       page, str(i + 1), str(self.non_otdl_violation),
                       page, str(i + 1), str(self.non_otdl_violation),
                       page, str(i + 1), str(self.non_otdl_violation),
                       page, str(i + 1),
                       page, str(i + 1), page, str(i + 1),
                       page, str(i + 1), page, str(i + 1), page, str(i + 1),
                       page, str(i + 1), page, str(i + 1), page, str(i + 1),
                       page, str(i + 1), page, str(i + 1), page, str(i + 1))
        self.instructions['T' + str(i)] = formula_d
        self.instructions.merge_cells('T' + str(i) + ':T' + str(i + 1))  # merge box for daily violation
        self.instructions['T' + str(i)].style = self.calcs
        self.instructions['T' + str(i)].number_format = "#,###.00"
        # instructions wed adjustment
        self.instructions.merge_cells('U' + str(i) + ':U' + str(i + 1))  # merge box for wed adj
        formula_e = "=IF(OR(%s!B%s=\"%s\",%s!B%s=\"nl\",%s!B%s=\"ptf\",%s!B%s=\"aux\")," \
                    "IF(AND(%s!S%s-(%s!N%s+%s!N%s+%s!P%s+%s!P%s)>0,%s!L%s>%s)," \
                    "IF(%s!S%s-(%s!N%s+%s!N%s+%s!P%s+%s!P%s)>%s!L%s-%s,%s!L%s-%s,%s!S%s-" \
                    "(%s!N%s+%s!N%s+%s!P%s+%s!P%s)),0)," \
                    "IF(AND(%s!S%s-(%s!N%s+%s!N%s+%s!P%s+%s!P%s)>0,%s!L%s>12)," \
                    "IF(%s!S%s-(%s!N%s+%s!N%s+%s!P%s+%s!P%s)>%s!L%s-12,%s!L%s-12,%s!S%s-" \
                    "(%s!N%s+%s!N%s+%s!P%s+%s!P%s)),0))" \
                    % (page, str(i), self.wal_12hr_mod,
                       page, str(i), page, str(i), page, str(i),
                       page, str(i), page, str(i + 1), page, str(i),
                       page, str(i + 1), page, str(i), page, str(i + 1),
                       str(self.non_otdl_violation),
                       page, str(i), page, str(i + 1), page, str(i),
                       page, str(i + 1), page, str(i), page, str(i + 1),
                       str(self.non_otdl_violation),
                       page, str(i + 1), str(self.non_otdl_violation),
                       page, str(i), page, str(i + 1),
                       page, str(i), page, str(i + 1), page, str(i),
                       page, str(i), page, str(i + 1), page, str(i),
                       page, str(i + 1), page, str(i), page, str(i + 1),
                       page, str(i), page, str(i + 1), page, str(i),
                       page, str(i + 1), page, str(i), page, str(i + 1),
                       page, str(i + 1), page, str(i), page, str(i + 1),
                       page, str(i), page, str(i + 1), page, str(i))
        self.instructions['U' + str(i)] = formula_e
        self.instructions['U' + str(i)].style = self.vert_calcs
        self.instructions['U' + str(i)].number_format = "#,###.00"
        # instructions thr adjustment
        formula_f = "=IF(OR(%s!B%s=\"%s\",%s!B%s=\"nl\",%s!B%s=\"ptf\",%s!B%s=\"aux\")," \
                    "IF(AND(%s!S%s-(%s!P%s+%s!P%s)>0,%s!N%s>%s)," \
                    "IF(%s!S%s-(%s!P%s+%s!P%s)>%s!N%s-%s,%s!N%s-%s,%s!S%s-(%s!P%s+%s!P%s)),0)," \
                    "IF(AND(%s!S%s-(%s!P%s+%s!P%s)>0,%s!N%s>12)," \
                    "IF(%s!S%s-(%s!P%s+%s!P%s)>%s!N%s-12,%s!N%s-12,%s!S%s-(%s!P%s+%s!P%s)),0))" \
                    % (page, str(i), self.wal_12hr_mod,
                       page, str(i), page, str(i), page, str(i),
                       page, str(i), page, str(i + 1), page, str(i),
                       page, str(i + 1), str(self.non_otdl_violation),
                       page, str(i), page, str(i + 1), page, str(i),
                       page, str(i + 1), str(self.non_otdl_violation),
                       page, str(i + 1), str(self.non_otdl_violation),
                       page, str(i), page, str(i + 1), page, str(i),
                       page, str(i), page, str(i + 1), page, str(i),
                       page, str(i + 1), page, str(i), page, str(i + 1), page, str(i),
                       page, str(i + 1), page, str(i + 1), page, str(i),
                       page, str(i + 1), page, str(i)
                       )
        self.instructions.merge_cells('V' + str(i) + ':V' + str(i + 1))  # merge box for thr adj
        self.instructions['V' + str(i)] = formula_f
        self.instructions['V' + str(i)].style = self.vert_calcs
        self.instructions['V' + str(i)].number_format = "#,###.00"
        # instructions fri adjustment
        self.instructions.merge_cells('W' + str(i) + ':W' + str(i + 1))  # merge box for fri adj
        formula_g = "=IF(OR(%s!B%s=\"%s\",%s!B%s=\"nl\",%s!B%s=\"aux\",%s!B%s=\"ptf\")," \
                    "IF(AND(%s!S%s>0,%s!P%s>%s)," \
                    "IF(%s!S%s>%s!P%s-%s,%s!P%s-%s,%s!S%s),0)," \
                    "IF(AND(%s!S%s>0,%s!P%s>12)," \
                    "IF(%s!S%s>%s!P%s-12,%s!P%s-12,%s!S%s),0))" \
                    % (page, str(i), self.wal_12hr_mod,
                       page, str(i), page, str(i), page, str(i),
                       page, str(i), page, str(i + 1), str(self.non_otdl_violation),
                       page, str(i), page, str(i + 1), str(self.non_otdl_violation),
                       page, str(i + 1), str(self.non_otdl_violation),
                       page, str(i), page, str(i), page, str(i + 1), page, str(i),
                       page, str(i + 1), page, str(i + 1), page, str(i))
        self.instructions['W' + str(i)] = formula_g
        self.instructions['W' + str(i)].style = self.vert_calcs
        self.instructions['W' + str(i)].number_format = "#,###.00"
        # instructions total violation
        self.instructions.merge_cells('X' + str(i) + ':X' + str(i + 1))  # merge box for total violation
        formula_h = "=SUM(%s!S%s:T%s)-(%s!U%s+%s!V%s+%s!W%s)" \
                    % (page, str(i), str(i), page, str(i),
                       page, str(i), page, str(i))
        self.instructions['X' + str(i)] = formula_h
        self.instructions['X' + str(i)].style = self.calcs
        self.instructions['X' + str(i)].number_format = "#,###.00"
        self.instructions['D12'] = "F"
        self.instructions['D12'].style = self.col_center_header
        self.instructions.merge_cells('D12:E12')
        self.instructions['F12'] = "F"
        self.instructions['F12'].style = self.col_center_header
        self.instructions.merge_cells('F12:G12')
        self.instructions['H12'] = "F"
        self.instructions['H12'].style = self.col_center_header
        self.instructions.merge_cells('H12:I12')
        self.instructions['J12'] = "F"
        self.instructions['J12'].style = self.col_center_header
        self.instructions.merge_cells('J12:K12')
        self.instructions['L12'] = "F"
        self.instructions['L12'].style = self.col_center_header
        self.instructions.merge_cells('L12:M12')
        self.instructions['N12'] = "F"
        self.instructions['N12'].style = self.col_center_header
        self.instructions.merge_cells('N12:O12')
        self.instructions['P12'] = "F"
        self.instructions['P12'].style = self.col_center_header
        self.instructions.merge_cells('P12:Q12')
        # legend section
        # create text for daily violations:
        text_k_a = "wal, "
        text_k_b = ""
        if self.wal_12hr_mod:
            text_k_a = ""
            text_k_b = " and wal"
        self.instructions.row_dimensions[14].height = 210
        self.instructions['A14'].style = self.instruct_text
        self.instructions.merge_cells('A14:X14')
        self.instructions['A14'] = \
            "Legend: \n" \
            "A.  Name \n" \
            "B.  List: Either otdl, wal, nl, ptf or aux (always use lowercase to preserve " \
            "operation of the formulas).\n" \
            "C.  Weekly 5200 Time: Enter the 5200 time for the week. \n" \
            "D.  Daily Non 5200 Time: Enter daily hours for either holiday, annual sick leave or " \
            "other type of paid leave.\n" \
            "E.  Daily Non 5200 Type: Enter “a” for annual, “s” for sick, “h” for holiday, etc. \n" \
            "F.  Daily 5200 Hours: Enter 5200 hours or hours worked for the day. \n" \
            "G.  No value allowed: No non 5200 times allowed for Sundays.\n" \
            "J.  Weekly Violations: This is the total of self.violations over 60 hours in a week.\n" \
            "K.  Daily Violations: This is the total of daily violations which have exceeded 11.50 " \
            "(for {}nl, ptf or aux)\n" \
            "     or 12 hours in a day (for otdl{}).\n" \
            "L.  Wednesday Adjustment: In cases were the 60 hour limit is reached " \
            "and a daily violation happens (on Wednesday),\n" \
            "     this column deducts one of the violations so to provide a correct remedy.\n" \
            "M.  Thursday Adjustment: In cases were the 60 hour limit is reached and " \
            "a daily violation happens (on Thursday), \n" \
            "     this column deducts one of the violations so to provide a correct remedy.\n" \
            "N.  Friday Adjustment: In cases were the 60 hour limit is reached and " \
            "a daily violation happens (on Friday),\n" \
            "     this column deducts one of the violations so to provide a correct remedy.\n" \
            "O.  Total Violation: This field is the end result of the calculation. " \
            "This is the addition of the total daily  " \
            "violations and the\n" \
            "     weekly violation, it shows the sum of the two. " \
            "This is the value which the steward should seek a remedy for.".format(text_k_a, text_k_b)
        self.instructions['A14'].alignment = Alignment(wrap_text=True, vertical='top')

    def violated_recs(self):
        """
        The violation record set is appended if the carrier has a daily violation or a weekly violation of
        over 60 hours in a week. It consist of 4 arrays: 1. carrier info (name and list), 2. daily hours array,
        3. daily leavetypes and 4 daily leavetimes. The carrier list the status on Saturday.
        """
        twelvehourlimit = ("otdl",)  # the 12 hour limit only applies to otdl carriers
        if self.wal_12hour:  # unless the WAL 12 Hour Violation setting is "on"
            twelvehourlimit = ("otdl", "wal")  # then the 12 hour limit applies to otdl and wal carriers
        i = 0
        while i <= len(self.carrier_list)-1:
            totals_array = ["", "", "", "", "", "", ""]  # daily hours
            leavetype_array = ["", "", "", "", "", "", ""]  # daily leave types
            leavetime_array = ["", "", "", "", "", "", ""]  # daily leave times
            carrier_rings = []
            total = 0.0
            grandtotal = 0.0
            # carrier name, list status for Saturday, and total weekly hours worked
            carrier_array = [self.carrier_list[i][1], self.carrier_list[i][2], 0.0]
            cc = 0
            daily_violation = False
            for day in self.dates:
                for ring in self.rings:
                    if ring[0] == str(day) and ring[1] == self.carrier_list[i][1]:  # find rings for carrier
                        carrier_rings.append(ring)  # add any rings to an array
                        if isfloat(ring[2]):
                            totals_array[cc] = float(ring[2])  # if hours worked is a number, add it as a number
                            if float(ring[2]) > 12 and self.carrier_list[i][2] in twelvehourlimit:
                                daily_violation = True
                            if float(ring[2]) > self.non_otdl_violation and \
                                    self.carrier_list[i][2] not in twelvehourlimit:
                                daily_violation = True
                        else:
                            totals_array[cc] = ring[2]  # if hours worked is empty string, add empty string
                        if ring[6] == "annual":
                            leavetype_array[cc] = "A"
                        if ring[6] == "sick":
                            leavetype_array[cc] = "S"
                        if ring[6] == "holiday":
                            leavetype_array[cc] = "H"
                        if ring[6] == "other":
                            leavetype_array[cc] = "O"
                        if ring[7] == "0.0" or ring[7] == "0":
                            leavetime_array[cc] = ""
                        elif isfloat(ring[7]):
                            leavetime_array[cc] = float(ring[7])
                        else:
                            leavetime_array[cc] = ring[7]
                cc += 1
            for item in carrier_rings:
                if item[2] == "":  # convert empty 5200 strings to zero
                    t = 0.0
                else:
                    t = float(item[2])
                if item[7] == "":  # convert leave time strings to zero
                    lv = 0.0
                else:
                    lv = float(item[7])
                total += t
                grandtotal = grandtotal + t + lv
            carrier_array[2] = total  # append total weekly hours worked to carrier array
            if grandtotal > 60 or daily_violation:  # only append violation recset, if there has been a violation
                violation_recset = [carrier_array, totals_array, leavetype_array, leavetime_array]  # build recset
                self.violation_recsets.append(violation_recset)  # append violations record set
            i += 1
        while len(self.violation_recsets) < self.min_rows:  # if minimum rows haven't been reached
            carrier_array = ["", "", 0.0]  # carrier information plus total weekly hours worked
            totals_array = ["", "", "", "", "", "", ""]  # daily hours
            leavetype_array = ["", "", "", "", "", "", ""]  # daily leave types
            leavetime_array = ["", "", "", "", "", "", ""]  # daily leave times
            violation_recset = [carrier_array, totals_array, leavetype_array, leavetime_array]  # combine
            self.violation_recsets.append(violation_recset)  # append these empty recs into the violations rec sets

    def get_pb_len(self):
        """ get the lenght of the progress bar. """
        self.pb.max_count(len(self.violation_recsets))  # set length of progress bar

    def show_violations(self):
        """ generates the rows of the violations worksheet. """
        summary_i = 7
        i = 10
        for line in self.violation_recsets:
            carrier_name = line[0][0]
            self.pbi += 1
            self.pb.move_count(self.pbi)  # increment progress bar
            self.pb.change_text("Building display for {}".format(carrier_name))
            carrier_list = line[0][1]
            total = line[0][2]
            totals_array = line[1]
            leavetype_array = line[2]
            leavetime_array = line[3]
            self.violations.row_dimensions[i].height = 10  # adjust all row height
            self.violations.row_dimensions[i + 1].height = 10
            self.violations.merge_cells('A' + str(i) + ':A' + str(i + 1))
            self.violations['A' + str(i)] = carrier_name  # name
            self.violations['A' + str(i)].style = self.input_name
            self.violations['A' + str(i+1)].style = self.input_name
            self.violations.merge_cells('B' + str(i) + ':B' + str(i + 1))  # merge box for list
            self.violations['B' + str(i)] = carrier_list  # list
            self.violations['B' + str(i)].style = self.input_s
            self.violations['B' + str(i+1)].style = self.input_s
            self.violations.merge_cells('C' + str(i) + ':C' + str(i + 1))  # merge box for weekly 5200
            self.violations['C' + str(i)] = Convert(total).empty_not_zerofloat()  # total
            self.violations['C' + str(i)].style = self.input_s
            self.violations['C' + str(i+1)].style = self.input_s
            self.violations['C' + str(i)].number_format = "#,###.00;[RED]-#,###.00"
            # saturday
            self.violations.merge_cells('D' + str(i + 1) + ':E' + str(i + 1))  # merge box for sat 5200
            self.violations['D' + str(i)] = leavetime_array[0]  # leave time
            self.violations['D' + str(i)].style = self.input_s
            self.violations['D' + str(i)].number_format = "#,###.00;[RED]-#,###.00"
            self.violations['E' + str(i)] = leavetype_array[0]  # leave type
            self.violations['E' + str(i)].style = self.input_s
            self.violations['D' + str(i + 1)] = totals_array[0]  # 5200 time
            self.violations['D' + str(i + 1)].style = self.input_s
            self.violations['D' + str(i + 1)].number_format = "#,###.00;[RED]-#,###.00"
            # sunday
            self.violations.merge_cells('F' + str(i + 1) + ':G' + str(i + 1))  # merge box for sun 5200
            self.violations['F' + str(i)] = leavetime_array[1]  # leave time
            self.violations['F' + str(i)].style = self.input_s
            self.violations['F' + str(i)].number_format = "#,###.00;[RED]-#,###.00"
            self.violations['G' + str(i)] = leavetype_array[1]  # leave type
            self.violations['G' + str(i)].style = self.input_s
            self.violations['F' + str(i + 1)] = totals_array[1]  # 5200 time
            self.violations['F' + str(i + 1)].style = self.input_s
            self.violations['F' + str(i + 1)].number_format = "#,###.00;[RED]-#,###.00"
            # monday
            self.violations.merge_cells('H' + str(i + 1) + ':I' + str(i + 1))  # merge box for mon 5200
            self.violations['H' + str(i)] = leavetime_array[2]  # leave time
            self.violations['H' + str(i)].style = self.input_s
            self.violations['H' + str(i)].number_format = "#,###.00;[RED]-#,###.00"
            self.violations['I' + str(i)] = leavetype_array[2]  # leave type
            self.violations['I' + str(i)].style = self.input_s
            self.violations['H' + str(i + 1)] = totals_array[2]  # 5200 time
            self.violations['H' + str(i + 1)].style = self.input_s
            self.violations['H' + str(i + 1)].number_format = "#,###.00;[RED]-#,###.00"
            # tuesday
            self.violations.merge_cells('J' + str(i + 1) + ':K' + str(i + 1))  # merge box for tue 5200
            self.violations['J' + str(i)] = leavetime_array[3]  # leave time
            self.violations['J' + str(i)].style = self.input_s
            self.violations['J' + str(i)].number_format = "#,###.00;[RED]-#,###.00"
            self.violations['K' + str(i)] = leavetype_array[3]  # leave type
            self.violations['K' + str(i)].style = self.input_s
            self.violations['J' + str(i + 1)] = totals_array[3]  # 5200 time
            self.violations['J' + str(i + 1)].style = self.input_s
            self.violations['J' + str(i + 1)].number_format = "#,###.00;[RED]-#,###.00"
            # wednesday
            self.violations.merge_cells('L' + str(i + 1) + ':M' + str(i + 1))  # merge box for wed 5200
            self.violations['L' + str(i)] = leavetime_array[4]  # leave time
            self.violations['L' + str(i)].style = self.input_s
            self.violations['L' + str(i)].number_format = "#,###.00;[RED]-#,###.00"
            self.violations['M' + str(i)] = leavetype_array[4]  # leave type
            self.violations['M' + str(i)].style = self.input_s
            self.violations['L' + str(i + 1)] = totals_array[4]  # 5200 time
            self.violations['L' + str(i + 1)].style = self.input_s
            self.violations['L' + str(i + 1)].number_format = "#,###.00;[RED]-#,###.00"
            # thursday
            self.violations.merge_cells('N' + str(i + 1) + ':O' + str(i + 1))  # merge box for thr 5200
            self.violations['N' + str(i)] = leavetime_array[5]  # leave time
            self.violations['N' + str(i)].style = self.input_s
            self.violations['N' + str(i)].number_format = "#,###.00;[RED]-#,###.00"
            self.violations['O' + str(i)] = leavetype_array[5]  # leave type
            self.violations['O' + str(i)].style = self.input_s
            self.violations['N' + str(i + 1)] = totals_array[5]  # 5200 time
            self.violations['N' + str(i + 1)].style = self.input_s
            self.violations['N' + str(i + 1)].number_format = "#,###.00;[RED]-#,###.00"
            # friday
            self.violations.merge_cells('P' + str(i + 1) + ':Q' + str(i + 1))  # merge box for fri 5200
            self.violations['P' + str(i)] = leavetime_array[6]  # leave time
            self.violations['P' + str(i)].style = self.input_s
            self.violations['P' + str(i)].number_format = "#,###.00;[RED]-#,###.00"
            self.violations['Q' + str(i)] = leavetype_array[6]  # leave type
            self.violations['Q' + str(i)].style = self.input_s
            self.violations['P' + str(i + 1)] = totals_array[6]  # 5200 time
            self.violations['P' + str(i + 1)].style = self.input_s
            self.violations['P' + str(i + 1)].number_format = "#,###.00;[RED]-#,###.00"
            # calculated fields
            # hidden columns
            formula_a = "=SUM(%s!D%s:P%s)+%s!D%s + %s!F%s + %s!H%s + %s!J%s + %s!L%s + " \
                        "%s!N%s + %s!P%s" % ("violations", str(i + 1), str(i + 1),
                                             "violations", str(i), "violations", str(i), "violations", str(i),
                                             "violations", str(i), "violations", str(i), "violations", str(i),
                                             "violations", str(i))
            self.violations['R' + str(i)] = formula_a
            self.violations['R' + str(i)].style = self.calcs
            self.violations['R' + str(i)].number_format = "#,###.00;[RED]-#,###.00"
            formula_b = "=SUM(%s!C%s+%s!D%s+%s!F%s+%s!H%s+%s!J%s+%s!L%s+%s!N%s+%s!P%s)" % \
                        ("violations", str(i), "violations", str(i), "violations", str(i),
                         "violations", str(i), "violations", str(i), "violations", str(i),
                         "violations", str(i), "violations", str(i))
            self.violations['R' + str(i + 1)] = formula_b
            self.violations['R' + str(i + 1)].style = self.calcs
            self.violations['R' + str(i + 1)].number_format = "#,###.00;[RED]-#,###.00"
            # weekly violation
            self.violations.merge_cells('S' + str(i) + ':S' + str(i + 1))  # merge box for weekly violation
            formula_c = "=IF(%s!B%s = \"aux\",0,MAX(IF(%s!R%s>%s!R%s,MAX(%s!R%s-60,0),MAX(%s!R%s-60)),0))" \
                        % ("violations", str(i), "violations", str(i), "violations", str(i + 1), "violations", str(i),
                           "violations", str(i + 1))
            self.violations['S' + str(i)] = formula_c
            self.violations['S' + str(i)].style = self.calcs
            self.violations['S' + str(i)].number_format = "#,###.00;[RED]-#,###.00"
            # daily violation
            formula_d = "=IF(OR(%s!B%s=\"%s\",%s!B%s=\"nl\",%s!B%s=\"ptf\",%s!B%s=\"aux\")," \
                        "(SUM(IF(%s!D%s>%s,%s!D%s-%s,0)+IF(%s!F%s>%s,%s!F%s-%s,0)" \
                        "+IF(%s!H%s>%s,%s!H%s-%s,0)+" \
                        "IF(%s!J%s>%s,%s!J%s-%s,0)" \
                        "+IF(%s!L%s>%s,%s!L%s-%s,0)+IF(%s!N%s>%s,%s!N%s-%s,0)+" \
                        "IF(%s!P%s>%s,%s!P%s-%s,0)))," \
                        "(SUM(IF(%s!D%s>12,%s!D%s-12,0)+IF(%s!F%s>12,%s!F%s-12,0)+IF(%s!H%s>12,%s!H%s-12,0)" \
                        "+IF(%s!J%s>12,%s!J%s-12,0)" \
                        "+IF(%s!L%s>12,%s!L%s-12,0)+IF(%s!N%s>12,%s!N%s-12,0)+IF(%s!P%s>12,%s!P%s-12,0))))"\
                        % ("violations", str(i), self.wal_12hr_mod,
                           "violations", str(i), "violations", str(i), "violations", str(i),
                           "violations", str(i + 1), str(self.non_otdl_violation),
                           "violations", str(i + 1), str(self.non_otdl_violation),
                           "violations", str(i + 1), str(self.non_otdl_violation),
                           "violations", str(i + 1), str(self.non_otdl_violation),
                           "violations", str(i + 1), str(self.non_otdl_violation),
                           "violations", str(i + 1), str(self.non_otdl_violation),
                           "violations", str(i + 1), str(self.non_otdl_violation),
                           "violations", str(i + 1), str(self.non_otdl_violation),
                           "violations", str(i + 1), str(self.non_otdl_violation),
                           "violations", str(i + 1), str(self.non_otdl_violation),
                           "violations", str(i + 1), str(self.non_otdl_violation),
                           "violations", str(i + 1), str(self.non_otdl_violation),
                           "violations", str(i + 1), str(self.non_otdl_violation),
                           "violations", str(i + 1), str(self.non_otdl_violation),
                           "violations", str(i + 1), "violations", str(i + 1), "violations", str(i + 1),
                           "violations", str(i + 1), "violations", str(i + 1), "violations", str(i + 1),
                           "violations", str(i + 1), "violations", str(i + 1), "violations", str(i + 1),
                           "violations", str(i + 1), "violations", str(i + 1), "violations", str(i + 1),
                           "violations", str(i + 1), "violations", str(i + 1))
            self.violations['T' + str(i)] = formula_d
            self.violations.merge_cells('T' + str(i) + ':T' + str(i + 1))  # merge box for daily violation
            self.violations['T' + str(i)].style = self.calcs
            self.violations['T' + str(i)].number_format = "#,###.00"
            # wed adjustment
            self.violations.merge_cells('U' + str(i) + ':U' + str(i + 1))  # merge box for wed adj
            formula_e = "=IF(OR(%s!B%s=\"%s\",%s!B%s=\"nl\",%s!B%s=\"ptf\",%s!B%s=\"aux\")," \
                        "IF(AND(%s!S%s-(%s!N%s+%s!N%s+%s!P%s+%s!P%s)>0,%s!L%s>%s)," \
                        "IF(%s!S%s-(%s!N%s+%s!N%s+%s!P%s+%s!P%s)>%s!L%s-%s,%s!L%s-%s,%s!S%s-" \
                        "(%s!N%s+%s!N%s+%s!P%s+%s!P%s)),0)," \
                        "IF(AND(%s!S%s-(%s!N%s+%s!N%s+%s!P%s+%s!P%s)>0,%s!L%s>12)," \
                        "IF(%s!S%s-(%s!N%s+%s!N%s+%s!P%s+%s!P%s)>%s!L%s-12,%s!L%s-12,%s!S%s-" \
                        "(%s!N%s+%s!N%s+%s!P%s+%s!P%s)),0))" \
                        % ("violations", str(i), self.wal_12hr_mod,
                           "violations", str(i), "violations", str(i), "violations", str(i),
                           "violations", str(i), "violations", str(i + 1), "violations", str(i),
                           "violations", str(i + 1), "violations", str(i), "violations", str(i + 1),
                           str(self.non_otdl_violation),
                           "violations", str(i), "violations", str(i + 1), "violations", str(i),
                           "violations", str(i + 1), "violations", str(i), "violations", str(i + 1),
                           str(self.non_otdl_violation),
                           "violations", str(i + 1),
                           str(self.non_otdl_violation),
                           "violations", str(i), "violations", str(i + 1),
                           "violations", str(i), "violations", str(i + 1), "violations", str(i),
                           "violations", str(i), "violations", str(i + 1), "violations", str(i),
                           "violations", str(i + 1), "violations", str(i), "violations", str(i + 1),
                           "violations", str(i), "violations", str(i + 1), "violations", str(i),
                           "violations", str(i + 1), "violations", str(i), "violations", str(i + 1),
                           "violations", str(i + 1), "violations", str(i), "violations", str(i + 1),
                           "violations", str(i), "violations", str(i + 1), "violations", str(i))
            self.violations['U' + str(i)] = formula_e
            self.violations['U' + str(i)].style = self.vert_calcs
            self.violations['U' + str(i)].number_format = "#,###.00"
            # thr adjustment
            formula_f = "=IF(OR(%s!B%s=\"%s\",%s!B%s=\"nl\",%s!B%s=\"ptf\",%s!B%s=\"aux\")," \
                        "IF(AND(%s!S%s-(%s!P%s+%s!P%s)>0,%s!N%s>%s)," \
                        "IF(%s!S%s-(%s!P%s+%s!P%s)>%s!N%s-%s,%s!N%s-%s,%s!S%s-(%s!P%s+%s!P%s)),0)," \
                        "IF(AND(%s!S%s-(%s!P%s+%s!P%s)>0,%s!N%s>12)," \
                        "IF(%s!S%s-(%s!P%s+%s!P%s)>%s!N%s-12,%s!N%s-12,%s!S%s-(%s!P%s+%s!P%s)),0))" \
                        % ("violations", str(i), self.wal_12hr_mod,
                           "violations", str(i), "violations", str(i), "violations", str(i),
                           "violations", str(i), "violations", str(i + 1), "violations", str(i),
                           "violations", str(i + 1), str(self.non_otdl_violation),
                           "violations", str(i), "violations", str(i + 1), "violations", str(i),
                           "violations", str(i + 1), str(self.non_otdl_violation),
                           "violations", str(i + 1), str(self.non_otdl_violation),
                           "violations", str(i),
                           "violations", str(i + 1), "violations", str(i),
                           "violations", str(i), "violations", str(i + 1), "violations", str(i),
                           "violations", str(i + 1),
                           "violations", str(i), "violations", str(i + 1), "violations", str(i),
                           "violations", str(i + 1), "violations", str(i + 1), "violations", str(i),
                           "violations", str(i + 1), "violations", str(i)
                           )
            self.violations.merge_cells('V' + str(i) + ':V' + str(i + 1))  # merge box for thr adj
            self.violations['V' + str(i)] = formula_f
            self.violations['V' + str(i)].style = self.vert_calcs
            self.violations['V' + str(i)].number_format = "#,###.00"
            # fri adjustment
            self.violations.merge_cells('W' + str(i) + ':W' + str(i + 1))  # merge box for fri adj
            formula_g = "=IF(OR(%s!B%s=\"%s\",%s!B%s=\"nl\",%s!B%s=\"ptf\",%s!B%s=\"aux\")," \
                        "IF(AND(%s!S%s>0,%s!P%s>%s)," \
                        "IF(%s!S%s>%s!P%s-%s,%s!P%s-%s,%s!S%s),0)," \
                        "IF(AND(%s!S%s>0,%s!P%s>12)," \
                        "IF(%s!S%s>%s!P%s-12,%s!P%s-12,%s!S%s),0))" \
                        % ("violations", str(i), self.wal_12hr_mod,
                           "violations", str(i), "violations", str(i),
                           "violations", str(i),
                           "violations", str(i), "violations", str(i + 1),
                           str(self.non_otdl_violation),
                           "violations", str(i), "violations", str(i + 1),
                           str(self.non_otdl_violation),
                           "violations", str(i + 1),
                           str(self.non_otdl_violation),
                           "violations", str(i), "violations", str(i), "violations",  str(i + 1), "violations", str(i),
                           "violations", str(i + 1), "violations", str(i + 1), "violations", str(i))
            self.violations['W' + str(i)] = formula_g
            self.violations['W' + str(i)].style = self.vert_calcs
            self.violations['W' + str(i)].number_format = "#,###.00"
            # total violation
            self.violations.merge_cells('X' + str(i) + ':X' + str(i + 1))  # merge box for total violation

            formula_h = "=IF(AND(%s!O4=\"yes\"," \
                        "OR(%s!B%s=\"otdl\", %s!B%s=\"%s\")),\"exempt\",SUM(%s!S%s:T%s)-(%s!U%s+%s!V%s+%s!W%s))" \
                        % ("violations", "violations", str(i),
                           "violations", str(i), self.wal_dec_exempt_mod,
                           "violations", str(i), str(i), "violations", str(i),
                           "violations", str(i), "violations", str(i))

            # formula_h = "=SUM(%s!S%s:T%s)-(%s!U%s+%s!V%s+%s!W%s)" \
            #             % ("violations", str(i), str(i), "violations", str(i),
            #                "violations", str(i), "violations", str(i))
            self.violations['X' + str(i)] = formula_h
            self.violations['X' + str(i)].style = self.calcs
            self.violations['X' + str(i)].number_format = "#,###.00"
            formula_i = "=IF(%s!A%s = 0,\"\",%s!A%s)" % ("violations", str(i), "violations", str(i))
            self.summary['A' + str(summary_i)] = formula_i
            self.summary['A' + str(summary_i)].style = self.input_name
            formula_j = "=%s!X%s" % ("violations", str(i))
            self.summary['B' + str(summary_i)] = formula_j
            self.summary['B' + str(summary_i)].style = self.input_s
            self.summary['B' + str(summary_i)].number_format = "#,###.00"
            self.summary.row_dimensions[summary_i].height = 10  # adjust all row height
            i += 2
            summary_i += 1
        i += 1
        self.violations.merge_cells('L' + str(i) + ':T' + str(i))  # label for cumulative violations
        self.violations['L' + str(i)] = "Cumulative Total Violations:  "
        self.violations['L' + str(i)].style = self.date_dov_title
        self.violations.merge_cells('U' + str(i) + ':X' + str(i))  # total violation summary at bottom of page
        formula_h = "=SUM(%s!X%s:X%s)" \
                    % ("violations", "9", str(i-2))
        self.violations['U' + str(i)] = formula_h
        self.violations['U' + str(i)].style = self.calcs
        self.violations['X' + str(i)].style = self.calcs
        self.violations['U' + str(i)].number_format = "#,###.00"
        self.violations.row_dimensions[i].height = 20  # adjust all row height

    def save_open(self):
        """ save the spreadsheet and open """
        self.pbi += 1
        self.pb.move_count(self.pbi)  # increment progress bar
        self.pb.change_text("Saving...")
        self.pb.stop()
        xl_filename = "kb_om" + str(format(projvar.invran_date_week[0], "_%y_%m_%d")) + ".xlsx"
        try:
            self.wb.save(dir_path('over_max_spreadsheet') + xl_filename)
            messagebox.showinfo("Spreadsheet generator",
                                "Your spreadsheet was successfully generated. \n"
                                "File is named: {}".format(xl_filename),
                                parent=self.frame)
            if sys.platform == "win32":  # open the text document
                os.startfile(dir_path('over_max_spreadsheet') + xl_filename)
            if sys.platform == "linux":
                subprocess.call(["xdg-open", 'kb_sub/over_max_spreadsheet/' + xl_filename])
            if sys.platform == "darwin":
                subprocess.call(["open", dir_path('over_max_spreadsheet') + xl_filename])
        except PermissionError:
            messagebox.showerror("Spreadsheet generator",
                                 "The spreadsheet was not generated. \n"
                                 "Suggestion: "
                                 "Make sure that identically named spreadsheets are closed "
                                 "(the file can't be overwritten while open).",
                                 parent=self.frame)


class ImpManSpreadsheet4:
    """
    This is a spreadsheet built to placate step B for Region 4.
    """
    def __init__(self):
        self.frame = None  # the frame of parent
        self.pb = None  # progress bar object
        self.pbi = 0  # progress bar count index
        self.startdate = None  # start date of the investigation
        self.enddate = None  # ending date of the investigation
        self.dates = []  # all days of the investigation
        self.carrierlist = []  # all carriers in carrier list
        self.carrier_breakdown = []  # all carriers in carrier list broken down into appropiate list
        self.mod_carrierlist = []
        self.tol_ot_ownroute = 0.0  # get tolerances from tolerances table.
        self.tol_ot_offroute = 0.0
        self.tol_availability = 0.0
        self.min_man4_nl = 0  # get minimum rows from tolerances table
        self.min_man4_wal = 0
        self.min_man4_otdl = 0
        self.min_man4_aux = 0
        self.wb = None  # the workbook object
        self.ws_list = []  # "saturday", "sunday", "monday", "tuesday", "wednesday", "thursday", "friday"
        self.day_of_week = []  # seven day array for weekly investigations/ one day array for daily investigations
        # styles for worksheet
        self.ws_header = None  # style
        self.list_header = None  # style
        self.date_dov = None  # style
        self.date_dov_title = None  # style
        self.col_header = None  # style
        self.input_name = None  # style
        self.input_s = None  # style
        self.calcs = None  # style
        self.quad_top = None  # style
        self.quad_bottom = None  # style
        self.quad_left = None  # style
        self.quad_right = None  # style
        self.col_header_left = None  # style
        self.col_header = None  # style
        self.footer_left = None  # style
        self.footer_right = None  # style
        self.footer_mid = None  # style
        self.day = None  # build worksheet - loop once for each day
        self.i = 0  # build worksheet loop iteration
        self.lsi = 0  # list loop iteration
        self.pref = ("nl", "wal", "otdl", "aux")
        self.row = 0
        # cell for summary quadrants page
        self.cellc9 = None  # non otdl own route violations
        self.cellf9 = None  # non otdl off route violations
        self.cellf11 = None  # wal off route violations
        self.cellj9 = None  # aux availability to 10 hours
        self.cellm9 = None  # aux availability to 11.5 hours
        self.cellj11 = None  # otdl availability to 10 hours
        self.cellm11 = None  # otdl availability to 12 hours
        self.cellf16 = None  # carriers out past dispatch of value
        self.celln16 = None  # otdl/aux availability to DOV

        self.ot_list = ("NON OTDL", "Work Assignment", "Auxiliary", "OTDL")  # list loop iteration
        self.page_titles = ("NON-OTDL Employees that worked overtime",
                            "Work Assignment Employees that worked off their assignment",
                            "Auxiliary Employees who were available to work OT",
                            "OTDL Employees who were available to work OT")
        self.pref = ("nl", "wal", "aux", "otdl")
        self.pb4_nl_wal = True  # page break between no list and work assignment
        self.pb4_wal_aux = True  # page break between work assignment and otdl
        self.pb4_aux_otdl = True  # page break between otdl and auxiliary
        self.min4_ss_nl = 0  # minimum rows for "no list"
        self.min4_ss_wal = 0  # minimum rows for work assignment list
        self.min4_ss_otdl = 0  # minimum rows for overtime desired list
        self.min4_ss_aux = 0  # minimum rows for auxiliary
        self.row_number = 1
        self.carrier = None  # current iteration of carrier's name is assigned self.carrier
        self.list_ = None  # current iteration of carrier's list status is assigned self.carrier
        self.route = None  # current iteration of carrier's route is assigned self.carrier
        self.rings = []  # assign as self.rings
        self.totalhours = 0.0  # set default as an empty string
        self.bt = ""
        self.rs = ""
        self.et = ""
        self.codes = ""
        self.moves = ""
        self.overtime = 0.0  # the amount of overtime worked by the carrier
        self.onroute = 0.0  # the amount of overworked on the carrier's own route.
        self.offroute = 0.0  # empty string or calculated time that carrier spent off their assignment
        self.offroute_adj = 0.0  # self.offroute adjusted for pivot time, ns days, and whole days off bid assignment
        self.otherroute_array = []  # a list of routes where carrier worked off assignment
        self.otherroute = ""  # the off assignment route the carrier worked on - formated for the cell
        self.avail_10 = 0.0  # otdl/aux availability to 10 hours
        self.avail_115 = 0.0  # aux availability to 11.50 hours
        self.avail_12 = 0.0  # otdl availability to 12 hours.
        self.lvtype = ""
        self.lvtime = ""
        self.first_row = 0  # record the number of the first row for totals formulas in footers
        self.last_row = 0  # record the number of the last row for totals formulas in footers
        # build a dictionary for displaying list statuses on spreadsheet
        self.list_dict = {'': '', 'nl': 'non list', 'wal': 'wal', 'otdl': 'otdl', 'aux': 'cca', 'ptf': 'ptf'}
        self.display_limiter = "show all"  # show all, only workdays, only mandates
        self.display_counter = 0  # count the number of rows displayed per list loop
        self.listrange = []  # records the first row, last row and summary row of each list
        self.dayrange = []  # records the listranges for the day by appending listranges after each listloop.
        self.dovarray = []  # build a list of 7 dov times. One for each day.

    def create(self, frame):
        """ a master method for running other methods in proper order."""
        self.frame = frame
        if not self.ask_ok():  # abort if user selects cancel from askokcancel
            return
        self.pb = ProgressBarDe(label="Building Improper Mandates Spreadsheet")
        self.pb.max_count(100)  # set length of progress bar
        self.pb.start_up()  # start the progress bar
        self.pbi = 1
        self.pb.move_count(self.pbi)  # increment progress bar
        self.pb.change_text("Gathering Data... ")
        self.get_dates()
        self.get_settings()  # get tolerances, minimum rows and page breaks from tolerances table
        self.get_pb_max_count()  # set the length of the progress bar
        self.get_carrierlist()
        self.get_carrier_breakdown()  # breakdown carrier list into no list, wal, otdl, aux
        self.get_tolerances()  # get tolerances, minimum rows and page break preferences from tolerances table
        self.get_dov()  # get the dispatch of value for each day
        self.get_styles()
        self.build_workbook()
        self.set_dimensions()
        self.build_ws_loop()  # loop once for each day
        self.save_open()

    def ask_ok(self):
        """ ends process if user cancels """
        if messagebox.askokcancel("Spreadsheet generator",
                                  "Do you want to generate an \nImproper Mandates No. 4 Spreadsheet?",
                                  parent=self.frame):
            return True
        return False

    def get_dates(self):
        """ get the dates from the project variables """
        self.startdate = projvar.invran_date  # set daily investigation range as default - get start date
        self.enddate = projvar.invran_date  # get end date
        self.dates = [projvar.invran_date, ]  # create an array of days - only one day if daily investigation range
        if projvar.invran_weekly_span:  # if the investigation range is weekly
            date = projvar.invran_date_week[0]
            self.startdate = projvar.invran_date_week[0]
            self.enddate = projvar.invran_date_week[6]
            self.dates = []
            for _ in range(7):  # create an array with all the days in the weekly investigation range
                self.dates.append(date)
                date += timedelta(days=1)

    def get_settings(self):
        """ get spreadsheet tolerances, row minimums and page break prefs from tolerance table """
        sql = "SELECT tolerance FROM tolerances"
        result = inquire(sql)
        self.tol_ot_ownroute = float(result[0][0])  # overtime on own route tolerance
        self.tol_ot_offroute = float(result[1][0])  # overtime off own route tolerance
        self.tol_availability = float(result[2][0])  # availability tolerance
        self.min4_ss_nl = int(result[32][0])  # minimum rows for no list
        self.min4_ss_wal = int(result[33][0])  # mimimum rows for work assignment
        self.min4_ss_otdl = int(result[34][0])  # minimum rows for otdl
        self.min4_ss_aux = int(result[35][0])  # minimum rows for auxiliary
        self.pb4_nl_wal = Convert(result[36][0]).str_to_bool()  # page break between no list and work assignment
        self.pb4_wal_aux = Convert(result[37][0]).str_to_bool()  # page break between work assignment and otdl
        self.pb4_aux_otdl = Convert(result[38][0]).str_to_bool()  # page break between otdl and auxiliary
        self.display_limiter = result[39][0]

    def get_pb_max_count(self):
        """ set length of progress bar """
        self.pb.max_count((len(self.dates)*4)+1)  # once for each list in each day, plus saving

    def get_carrierlist(self):
        """ get record sets for all carriers """
        self.carrierlist = CarrierList(self.startdate, self.enddate, projvar.invran_station).get()

    def get_carrier_breakdown(self):
        """ breakdown carrier list into no list, wal, otdl, aux """
        timely_rec = []
        for day in self.dates:
            nl_array = []
            wal_array = []
            aux_array = []
            otdl_array = []
            for carrier in self.carrierlist:
                for rec in reversed(carrier):
                    if Convert(rec[0]).dt_converter() <= day:
                        timely_rec = rec
                if timely_rec[2] == "nl":
                    nl_array.append(timely_rec)
                if timely_rec[2] == "wal":
                    wal_array.append(timely_rec)
                if timely_rec[2] == "otdl":
                    otdl_array.append(timely_rec)
                if timely_rec[2] == "aux" or timely_rec[2] == "ptf":
                    aux_array.append(timely_rec)
            daily_breakdown = [nl_array, wal_array, aux_array, otdl_array]
            self.carrier_breakdown.append(daily_breakdown)

    def get_tolerances(self):
        """ get spreadsheet tolerances, row minimums and page break prefs from tolerance table """
        sql = "SELECT tolerance FROM tolerances"
        result = inquire(sql)
        self.tol_ot_ownroute = float(result[0][0])  # overtime on own route tolerance
        self.tol_ot_offroute = float(result[1][0])  # overtime off own route tolerance
        self.tol_availability = float(result[2][0])  # availability tolerance
        self.min_man4_nl = int(result[3][0])  # minimum rows for no list
        self.min_man4_wal = int(result[4][0])  # mimimum rows for work assignment
        self.min_man4_otdl = int(result[5][0])  # minimum rows for otdl
        self.min_man4_aux = int(result[6][0])  # minimum rows for auxiliary

    def get_dov(self):
        """ get the dov records currently in the database """
        days = ("sat", "sun", "mon", "tue", "wed", "thu", "fri")
        for i in range(len(days)):
            sql = "SELECT * FROM dov WHERE eff_date <= '%s' AND station = '%s' AND day = '%s' " \
                  "ORDER BY eff_date DESC" % \
                  (projvar.invran_date_week[0], projvar.invran_station, days[i])
            result = inquire(sql)
            for rec in result:
                if rec[0] == Convert(projvar.invran_date_week[0]).dt_to_str():
                    self.dovarray.append(rec[3])
                    break
                elif rec[4] == "False":
                    self.dovarray.append(rec[3])
                    break
                else:
                    continue

    def get_styles(self):
        """ Named styles for workbook """
        bd = Side(style='thin', color="80808080")  # defines borders
        self.ws_header = NamedStyle(name="ws_header", font=Font(bold=True, name='Arial', size=12))
        self.list_header = NamedStyle(name="list_header", font=Font(bold=True, name='Arial', size=10))
        self.date_dov = NamedStyle(name="date_dov", font=Font(name='Arial', size=8))
        self.date_dov_title = NamedStyle(name="date_dov_title", font=Font(bold=True, name='Arial', size=8),
                                         alignment=Alignment(horizontal='right'))
        self.col_header_left = NamedStyle(name="col_header_left", font=Font(bold=True, name='Arial', size=8),
                                          alignment=Alignment(horizontal='left', vertical='bottom'))
        self.col_header = NamedStyle(name="col_header", font=Font(bold=True, name='Arial', size=8),
                                     alignment=Alignment(horizontal='center', vertical='bottom'))
        self.input_name = NamedStyle(name="input_name", font=Font(name='Arial', size=8),
                                     border=Border(left=bd, top=bd, right=bd, bottom=bd))
        self.input_s = NamedStyle(name="input_s", font=Font(name='Arial', size=8),
                                  border=Border(left=bd, top=bd, right=bd, bottom=bd),
                                  alignment=Alignment(horizontal='right'))
        self.calcs = NamedStyle(name="calcs", font=Font(name='Arial', size=8),
                                border=Border(left=bd, top=bd, right=bd, bottom=bd),
                                fill=PatternFill(fgColor='e5e4e2', fill_type='solid'),
                                alignment=Alignment(horizontal='right'))

        self.quad_top = NamedStyle(name="quad_top", font=Font(name='Arial', size=10),
                                   alignment=Alignment(horizontal='left', vertical='top'),
                                   border=Border(left=bd, top=bd, right=bd))
        self.quad_bottom = NamedStyle(name="quad_bottom", font=Font(name='Arial', size=10),
                                      alignment=Alignment(horizontal='right'),
                                      border=Border(left=bd, bottom=bd, right=bd))
        self.quad_left = NamedStyle(name="quad_left", font=Font(name='Arial', size=10),
                                    alignment=Alignment(horizontal='left', vertical='top'),
                                    border=Border(left=bd, bottom=bd, top=bd))
        self.quad_right = NamedStyle(name="quad_right", font=Font(name='Arial', size=10),
                                     alignment=Alignment(horizontal='right', vertical='top'),
                                     border=Border(top=bd, bottom=bd, right=bd))
        self.footer_left = NamedStyle(name="footer_left", font=Font(bold=True, name='Arial', size=8),
                                      alignment=Alignment(horizontal='left'),
                                      border=Border(left=bd, bottom=bd, top=bd))
        self.footer_right = NamedStyle(name="footer_right", font=Font(bold=True, name='Arial', size=8),
                                       alignment=Alignment(horizontal='right'),
                                       border=Border(top=bd, bottom=bd, right=bd))
        self.footer_mid = NamedStyle(name="footer_mid", font=Font(bold=True, name='Arial', size=8),
                                     alignment=Alignment(horizontal='right'),
                                     border=Border(top=bd, bottom=bd))

    def build_workbook(self):
        """ build the workbook object """
        day_finder = ["Sat", "Sun", "Mon", "Tue", "Wed", "Thu", "Fri"]
        day_of_week = ["saturday", "sunday", "monday", "tuesday", "wednesday", "thursday", "friday"]
        i = 0
        self.wb = Workbook()  # define the workbook
        if not projvar.invran_weekly_span:  # if investigation range is daily
            for ii in range(len(day_finder)):
                if projvar.invran_date.strftime("%a") == day_finder[ii]:  # find the correct day
                    i = ii
            self.ws_list.append(self.wb.active)  # create first worksheet
            self.ws_list[0].title = day_of_week[i]  # title first worksheet
            self.day_of_week.append(day_of_week[i])  # create self.day_of_week array with one day
        if projvar.invran_weekly_span:  # if investigation range is weekly
            for day in day_of_week:
                self.day_of_week.append(day)  # create self.day_of_week array with seven days
            self.ws_list.append(self.wb.active)  # create first worksheet
            self.ws_list[0].title = "saturday"  # title first worksheet
            for i in range(1, 7):  # create worksheet for remaining six days
                self.ws_list.append(self.wb.create_sheet(day_of_week[i]))  # create subsequent worksheets
                self.ws_list[i].title = day_of_week[i]  # title subsequent worksheets

    def set_dimensions(self):
        """ set the orientation and dimensions of the workbook """
        for i in range(len(self.dates)):
            self.ws_list[i].set_printer_settings(paper_size=1, orientation='landscape')  # set orientation
            self.ws_list[i].oddFooter.center.text = "&A"  # include the footer
            self.ws_list[i].column_dimensions["A"].width = 4  # column width
            self.ws_list[i].column_dimensions["B"].width = 4
            self.ws_list[i].column_dimensions["C"].width = 9
            self.ws_list[i].column_dimensions["D"].width = 9
            self.ws_list[i].column_dimensions["E"].width = 5
            self.ws_list[i].column_dimensions["F"].width = 4
            self.ws_list[i].column_dimensions["G"].width = 9
            self.ws_list[i].column_dimensions["H"].width = 9
            self.ws_list[i].column_dimensions["I"].width = 9
            self.ws_list[i].column_dimensions["J"].width = 9
            self.ws_list[i].column_dimensions["K"].width = 9
            self.ws_list[i].column_dimensions["L"].width = 5
            self.ws_list[i].column_dimensions["M"].width = 4
            self.ws_list[i].column_dimensions["N"].width = 9
            self.ws_list[i].column_dimensions["O"].width = 9
            self.ws_list[i].column_dimensions["P"].width = 9
            self.ws_list[i].row_dimensions[8].height = 45  # adjust row height
            self.ws_list[i].row_dimensions[10].height = 45  # adjust row height

    def build_ws_loop(self):
        """ this loops once for each list. """
        self.i = 0
        for day in self.dates:
            self.dayrange = []  # initialize array for holding all start/stop/summary rows for all four list.
            self.day = day
            self.build_ws_headers()
            self.build_ws_quads()
            self.pagebreak(force=True)  # force the page break
            self.list_loop()  # loops four times. once for each list.
            self.fill_quads()  # write formulas for the quadrants on the cover sheet.
            self.i += 1

    def build_ws_headers(self):
        """ worksheet headers """
        cell = self.ws_list[self.i].cell(row=1, column=3)
        cell.value = "Improper Mandate Worksheet"
        cell.style = self.ws_header
        self.ws_list[self.i].merge_cells('C1:O1')
        cell = self.ws_list[self.i].cell(row=3, column=3)  # create date label
        cell.value = "Date:  "
        cell.style = self.date_dov_title
        cell = self.ws_list[self.i].cell(row=3, column=4)  # display date
        cell.value = format(self.day, "%A  %m/%d/%y")
        cell.style = self.date_dov
        self.ws_list[self.i].merge_cells('D3:H3')
        cell = self.ws_list[self.i].cell(row=3, column=9)  # pay period label
        cell.value = "Pay Period:  "
        cell.style = self.date_dov_title
        self.ws_list[self.i].merge_cells('I3:J3')
        cell = self.ws_list[self.i].cell(row=3, column=11)  # display pay period
        cell.value = projvar.pay_period
        cell.style = self.date_dov
        self.ws_list[self.i].merge_cells('K3:O3')
        cell = self.ws_list[self.i].cell(row=4, column=3)  # station label
        cell.value = "Station:  "
        cell.style = self.date_dov_title
        cell = self.ws_list[self.i].cell(row=4, column=4)  # display station
        cell.value = projvar.invran_station
        cell.style = self.date_dov
        self.ws_list[self.i].merge_cells('D4:H4')

    def build_ws_quads(self):
        """ build the two quadrants and other elements on the coversheet """
        # Violations Quadrants
        # Top Left
        self.ws_list[self.i].merge_cells('C8:E8')
        cell = self.ws_list[self.i].cell(row=8, column=3)  # NON-OTDL Violations on own route Page 1
        cell.value = "NON-OTDL \nViolations on own route \nPage 1"
        cell.style = self.quad_top
        self.ws_list[self.i].merge_cells('C9:E9')
        self.cellc9 = self.ws_list[self.i].cell(row=9, column=3)
        self.cellc9.value = ""
        self.cellc9.style = self.quad_bottom
        self.cellc9.number_format = "#,###.00;[RED]-#,###.00"
        # Top Right
        cell = self.ws_list[self.i].cell(row=8, column=6)  # NON-OTDL Violations off own route Page 1
        cell.value = "NON-OTDL \nViolations off own route \nPage 1"
        cell.style = self.quad_top
        self.ws_list[self.i].merge_cells('F8:H8')
        self.cellf9 = self.ws_list[self.i].cell(row=9, column=6)  # value filled later
        self.cellf9.style = self.quad_bottom
        self.ws_list[self.i].merge_cells('F9:H9')
        self.cellf9.number_format = "#,###.00;[RED]-#,###.00"
        # Bottom Left
        self.ws_list[self.i].merge_cells('C10:E10')
        cell = self.ws_list[self.i].cell(row=10, column=3)  # Blank
        cell.value = ""
        cell.style = self.quad_top
        cell = self.ws_list[self.i].cell(row=11, column=3)
        cell.value = ""
        cell.style = self.quad_bottom
        self.ws_list[self.i].merge_cells('C11:E11')
        # Bottom Right
        cell = self.ws_list[self.i].cell(row=10, column=6)  # Work assignment Violations off own route Page 2
        cell.value = "Work Assignment \nViolations off own route \nPage 2"
        cell.style = self.quad_top
        self.ws_list[self.i].merge_cells('F10:H10')
        self.cellf11 = self.ws_list[self.i].cell(row=11, column=6)
        self.cellf11.value = ""
        self.cellf11.style = self.quad_bottom
        self.ws_list[self.i].merge_cells('F11:H11')
        self.cellf11.number_format = "#,###.00;[RED]-#,###.00"
        # Totals Left
        cell = self.ws_list[self.i].cell(row=12, column=3)
        formula = "= %s!C%s" % (self.day_of_week[self.i], "9")
        cell.value = formula
        cell.style = self.quad_bottom
        self.ws_list[self.i].merge_cells('C12:E12')
        cell.number_format = "#,###.00;[RED]-#,###.00"
        # Totals Right
        cell = self.ws_list[self.i].cell(row=12, column=6)
        formula = "= SUM(%s!F%s + %s!F%s)" % (self.day_of_week[self.i], "9", self.day_of_week[self.i], "11")
        cell.value = formula
        cell.style = self.quad_bottom
        self.ws_list[self.i].merge_cells('F12:H12')
        cell.number_format = "#,###.00;[RED]-#,###.00"
        # Availability Quadrants
        # Top Left
        cell = self.ws_list[self.i].cell(row=8, column=10)  # CCAs Available to 10 hours Page 3
        cell.value = "CCAs \nAvailable to 10 hours \nPage 3"
        cell.style = self.quad_top
        self.ws_list[self.i].merge_cells('J8:L8')
        self.cellj9 = self.ws_list[self.i].cell(row=9, column=10)  # value filled later
        self.cellj9.style = self.quad_bottom
        self.ws_list[self.i].merge_cells('J9:L9')
        self.cellj9.number_format = "#,###.00;[RED]-#,###.00"
        # Top Right
        cell = self.ws_list[self.i].cell(row=8, column=13)  # CCAs Available to 11.5 hours Page 3
        cell.value = "CCAs \nAvailable to 11.5 hours \nPage 3"
        cell.style = self.quad_top
        self.ws_list[self.i].merge_cells('M8:O8')
        self.cellm9 = self.ws_list[self.i].cell(row=9, column=13)  # value filled later
        self.cellm9.style = self.quad_bottom
        self.ws_list[self.i].merge_cells('M9:O9')
        self.cellm9.number_format = "#,###.00;[RED]-#,###.00"
        # Bottom Left
        cell = self.ws_list[self.i].cell(row=10, column=10)  # OTDL Available to 10 hours Page 4
        cell.value = "OTDL \nAvailable to 10 hours \nPage 4"
        cell.style = self.quad_top
        self.ws_list[self.i].merge_cells('J10:L10')
        self.cellj11 = self.ws_list[self.i].cell(row=11, column=10)  # value filled later
        self.cellj11.style = self.quad_bottom
        self.ws_list[self.i].merge_cells('J11:L11')
        self.cellj11.number_format = "#,###.00;[RED]-#,###.00"
        # Bottom Right
        cell = self.ws_list[self.i].cell(row=10, column=13)  # OTDL Available to 12 hours Page 4
        cell.value = "OTDL \nAvailable to 12 hours \nPage 4"
        cell.style = self.quad_top
        self.ws_list[self.i].merge_cells('M10:O10')
        self.cellm11 = self.ws_list[self.i].cell(row=11, column=13)  # value filled later
        self.cellm11.style = self.quad_bottom
        self.ws_list[self.i].merge_cells('M11:O11')
        self.cellm11.number_format = "#,###.00;[RED]-#,###.00"
        # Totals Left
        cell = self.ws_list[self.i].cell(row=12, column=10)
        formula = "= SUM(%s!J%s + %s!J%s)" % (self.day_of_week[self.i], "9", self.day_of_week[self.i], "11")
        cell.value = formula
        cell.style = self.quad_bottom
        self.ws_list[self.i].merge_cells('J12:L12')
        cell.number_format = "#,###.00;[RED]-#,###.00"
        # Totals Right
        cell = self.ws_list[self.i].cell(row=12, column=13)
        formula = "= SUM(%s!M%s + %s!M%s)" % (self.day_of_week[self.i], "9", self.day_of_week[self.i], "11")
        cell.value = formula
        cell.style = self.quad_bottom
        self.ws_list[self.i].merge_cells('M12:O12')
        cell.number_format = "#,###.00;[RED]-#,###.00"
        # DOV box
        cell = self.ws_list[self.i].cell(row=14, column=3)  # Dispatch of Value
        cell.value = "Dispatch of Value"
        cell.style = self.quad_top
        self.ws_list[self.i].merge_cells('C14:E15')
        cell = self.ws_list[self.i].cell(row=16, column=3)  # aquired by get_dov()
        cell.value = self.dovarray[self.i]
        cell.style = self.quad_bottom
        self.ws_list[self.i].merge_cells('C16:E16')
        cell.number_format = "#,###.00;[RED]-#,###.00"

        # DOV Violations box
        cell = self.ws_list[self.i].cell(row=14, column=6)  # Dispatch of Value
        cell.value = "Carriers out past \nDispatch of Value"
        cell.style = self.quad_top
        self.ws_list[self.i].merge_cells('F14:H15')
        self.cellf16 = self.ws_list[self.i].cell(row=16, column=6)  # value filled later - # of carriers out past DOV
        self.cellf16.style = self.quad_bottom
        self.ws_list[self.i].merge_cells('F16:H16')
        self.cellf16.number_format = "#,##0"

        # Straight Time Available:
        cell = self.ws_list[self.i].cell(row=14, column=10)  # Straight Time Available:
        cell.value = "Straight Time Available:"
        cell.style = self.quad_left
        self.ws_list[self.i].merge_cells('J14:M14')
        cell = self.ws_list[self.i].cell(row=14, column=14)  # straight time available formula
        cell.value = ""
        cell.style = self.quad_right
        self.ws_list[self.i].merge_cells('N14:O14')

        # Available to DOV:
        cell = self.ws_list[self.i].cell(row=16, column=10)  # Available to DOV:
        cell.value = "Available to DOV:"
        cell.style = self.quad_left
        self.ws_list[self.i].merge_cells('J16:M16')

        self.celln16 = self.ws_list[self.i].cell(row=16, column=14)  # avaiable to dov formula
        self.celln16.style = self.quad_right
        self.ws_list[self.i].merge_cells('N16:O16')
        self.celln16.number_format = "#,###.00;[RED]-#,###.00"
        self.row = 19  # starts on row 19 to give room to the quadrants

    def list_loop(self):
        """ loops four times. once for each list. """
        self.lsi = 0  # iterations of the list loop method
        for _ in self.ot_list:  # loops for nl, wal, otdl and aux
            self.list_and_column_headers()  # builds headers for list and columns
            self.carrierlist_mod()
            self.get_first_row()
            self.row_number = 1  # initialize the row number that appears on the far left column
            self.carrierloop()  # loop once to fill a row with carrier rings data
            self.fill_for_minrows()  # fill in blank rows to fullfill minrows requirement
            self.get_listrange()
            self.dayrange.append(self.listrange)  # get rows information for each list loop.
            self.build_footer()
            self.pagebreak()
            self.increment_progbar()
            self.lsi += 1
        self.lsi = 0  # reset list loop iteration

    def list_and_column_headers(self):
        """ builds headers for list and column """
        n_14heads = ("On Own \nRoute", "Off Route", "Available \nto 10", "Available \nto 10")
        o_15heads = ("Off Route", "Other Route", "Available \nto 11.50", "Available \nto 12.00")
        p_16heads = ("Other Route", "", "Available \nto DOV", "Available \nto DOV")

        cell = self.ws_list[self.i].cell(row=self.row, column=1)
        cell.value = self.page_titles[self.lsi]  # Displays the page title for each list,
        cell.style = self.list_header
        cell = self.ws_list[self.i].cell(row=self.row, column=16)
        cell.value = "Page {}".format(self.lsi+1)  # Displays the page title for each list,
        cell.style = self.list_header
        self.row += 2
        self.ws_list[self.i].row_dimensions[self.row].height = 30  # adjust row height
        cell = self.ws_list[self.i].cell(row=self.row, column=2)  # Name Header
        cell.value = "Name"
        cell.style = self.col_header_left
        self.ws_list[self.i].merge_cells('B' + str(self.row) + ':D' + str(self.row))

        cell = self.ws_list[self.i].cell(row=self.row, column=5)  # OTDL List/Status
        cell.value = "OTDL \nList"
        if self.ot_list[self.lsi] == "Auxiliary":  # Aux has header variation
            cell.value = "status"
        cell.style = self.col_header_left
        self.ws_list[self.i].merge_cells('E' + str(self.row) + ':F' + str(self.row))

        cell = self.ws_list[self.i].cell(row=self.row, column=7)  # Route Assigned
        cell.value = "Route \nAssigned"
        cell.style = self.col_header_left

        cell = self.ws_list[self.i].cell(row=self.row, column=8)  # BT
        cell.value = "BT"
        cell.style = self.col_header

        cell = self.ws_list[self.i].cell(row=self.row, column=9)  # MV
        cell.value = "MV"
        cell.style = self.col_header

        cell = self.ws_list[self.i].cell(row=self.row, column=10)  # RS
        cell.value = "RS"
        cell.style = self.col_header

        cell = self.ws_list[self.i].cell(row=self.row, column=11)  # ET
        cell.value = "ET"
        cell.style = self.col_header

        cell = self.ws_list[self.i].cell(row=self.row, column=12)  # Overtime Worked
        cell.value = "Overtime \nWorked"
        cell.style = self.col_header_left
        self.ws_list[self.i].merge_cells('L' + str(self.row) + ':M' + str(self.row))

        cell = self.ws_list[self.i].cell(row=self.row, column=14)  # On/off route, avail to 10
        cell.value = n_14heads[self.lsi]
        cell.style = self.col_header_left

        cell = self.ws_list[self.i].cell(row=self.row, column=15)  # off/other route, avail to 11.5/12
        cell.value = o_15heads[self.lsi]
        cell.style = self.col_header_left

        if self.lsi != 1:  # do not display this column for work assignment carriers
            cell = self.ws_list[self.i].cell(row=self.row, column=16)  # other route, avail to DOV
            cell.value = p_16heads[self.lsi]
            cell.style = self.col_header_left
        self.row += 1  # increment the row so first name starts on fresh line.

    def carrierlist_mod(self):
        """ add empty carrier records to carrier list until quantity matches minrows preference """
        self.mod_carrierlist = self.carrier_breakdown[self.i][self.lsi]

    def get_first_row(self):
        """ record the number of the first row for totals formulas in footers """
        self.first_row = self.row

    def carrierloop(self):
        """ loop for each carrier """
        self.display_counter = 0  # count the number of rows displayed
        for carrier in self.mod_carrierlist:
            self.carrier = carrier[1]  # current iteration of carrier list is assigned self.carrier
            self.list_ = carrier[2]  # get the list status of the carrier
            self.route = carrier[4]  # get the route of the carrier
            self.display_route()  # will alter self.route if the route is a floater/t6 string.
            self.get_rings()  # get individual carrier rings for the day
            self.number_crunching()  # do calculations to get overtime and availability
            if self.qualify():  # test the rings to see if they need to be displayed
                self.display_recs()  # build the carrier and the rings row into the spreadsheet
                self.display_counter += 1
                self.row += 1

    def fill_for_minrows(self):
        """ fill in blank rows to fullfill minrows requirement. """
        self.carrier = ""  # current iteration of carrier list is assigned self.carrier
        self.list_ = ""  # get the list status of the carrier
        self.route = ""  # get the route of the carrier
        self.totalhours = 0.0  # set default as an empty string
        self.bt = ""  # begin tour
        self.rs = ""  # return to station
        self.et = ""  # end tour
        self.codes = ""  # codes from carrier rings
        self.moves = ""  # moves from carrier rings
        self.overtime = 0.0  # the total overtime worked
        self.onroute = 0.0  # the amount of overtime worked on the carrier's own route.
        self.offroute = 0.0  # total time spend off route
        self.offroute_adj = 0.0  # self.offroute adjusted for pivot time, ns days, and whole days off bid assignment
        self.otherroute_array = []  # a list of routes where carrier worked off assignment
        self.otherroute = ""  # a formatted display for routes worked off assignment
        self.avail_10 = 0.0  # otdl/aux availability to 10 hours
        self.avail_115 = 0.0  # aux availability to 11.50 hours
        self.avail_12 = 0.0  # otdl availability to 12 hours.
        blank_lines = []  # make an array for blank lines
        if self.pref[self.lsi] in ("nl",):  # if "no list"
            minrows = self.min4_ss_nl
        elif self.pref[self.lsi] in ("wal",):  # if "work assignment list"
            minrows = self.min4_ss_wal
        elif self.pref[self.lsi] in ("otdl",):  # if "overtime desired list"
            minrows = self.min4_ss_otdl
        else:  # if "auxiliary"
            minrows = self.min4_ss_aux
        while self.display_counter < minrows:  # until carrier list quantity matches minrows
            add_this = ('', '', '', '', '', '')
            blank_lines.append(add_this)  # append empty recs to carrier list
            self.display_counter += 1
        for _ in blank_lines:
            self.display_recs()  # put the carrier and the first part of rings into the spreadsheet
            self.row += 1

    def display_route(self):
        """ formats route number for floater/t6 strings into a short version """
        if self.route:
            route = self.route.split("/")
            if len(route) == 5:
                self.route = "T6: {} +".format(route[0])

    def get_rings(self):
        """ get individual carrier rings for the day """
        self.rings = Rings(self.carrier, self.dates[self.i]).get_for_day()  # assign as self.rings
        self.totalhours = 0.0  # set default as an empty string
        self.bt = ""
        self.rs = ""
        self.et = ""
        self.codes = ""
        self.moves = ""
        self.lvtype = ""
        self.lvtime = ""
        if self.rings[0]:  # if rings record is not blank
            self.totalhours = float(Convert(self.rings[0][2]).zero_not_empty())
            self.bt = self.rings[0][9]
            self.rs = self.rings[0][3]
            self.et = self.rings[0][10]
            self.codes = self.rings[0][4]
            self.moves = self.rings[0][5]
            self.lvtype = self.rings[0][6]
            self.lvtime = self.rings[0][7]

    def number_crunching(self):
        """ crunch numbers to get overtime, off route, other route and availability"""
        self.overtime = 0.0  # the total overtime worked
        self.onroute = 0.0  # the amount of overtime worked on the carrier's own route.
        self.offroute = 0.0  # total time spend off route
        self.offroute_adj = 0.0
        self.otherroute_array = []  # a list of routes where carrier worked off assignment
        self.otherroute = ""  # a formatted display for routes worked off assignment
        self.avail_10 = 0.0  # otdl/aux availability to 10 hours
        self.avail_115 = 0.0  # aux availability to 11.50 hours
        self.avail_12 = 0.0  # otdl availability to 12 hours.
        self.calc_overtime()  # calculate the amount of overtime worked
        if self.pref[self.lsi] in ("nl", "wal"):
            if self.moves:
                self.calc_offroute()  # calculate the time that the carrier spent off their route and get other route
                self.format_otherroute()  # format the self.other route so that if fits in the spreadsheet cell
            self.calc_offroute_adj()  # adj for pivot time or if code is nsday or whole day spent off route
        if self.pref[self.lsi] == "nl":
            self.calc_onroute()  # calculate the overtime worked on carrier's own route.
        if self.pref[self.lsi] in ("otdl", "aux"):
            self.calc_availability()
            self.moves = ""  # empty self.moves for otdl and aux carriers.

    def calc_overtime(self):
        """ calculates the amount of overtime worked. if it is the carrier's ns day, then the full day is overtime. """
        if self.codes == "ns day":
            self.overtime = self.totalhours
        else:
            self.overtime = max(self.totalhours - 8, 0)

    def calc_offroute(self):
        """ calculate the time that the carrier spent off their route assignment, get other route """
        moves = self.moves.split(",")
        move_sets = int(len(moves)/3)  # get the number of triads in the moves array
        count = 0
        for _ in range(move_sets):
            offroute = float(moves[count+1]) - float(moves[count])  # calculate off route time per triad
            self.offroute += offroute  # add triad time off route
            self.otherroute_array.append(moves[count+2])
            count += 3
        self.offroute = round(self.offroute, 2)
        if self.offroute >= self.totalhours:
            self.offroute = self.totalhours
        self.moves = moves[0]  # replace moves with the first time moved off route

    def calc_offroute_adj(self):
        """ calculate the off route overtime for ns days or if the whole day is spent off own route. """
        self.offroute_adj = min(self.overtime, self.offroute)  # will adjust for pivot time
        if self.codes == "ns day":  # if it is the ns day, then whole day is off route
            self.offroute_adj = self.totalhours
            self.otherroute_array.append("ns day")
            self.otherroute = "ns day"
            self.moves = self.bt
        if self.totalhours:
            # if self.totalhours <= self.offroute:  # if the whole day is off route
            if self.offroute == self.totalhours:  # if the whole day is off route
                self.offroute_adj = self.totalhours
                self.otherroute = "off bid"

    def calc_onroute(self):
        """ calculate the overtime the carrier worked on their own route. """
        if self.codes == "ns day":
            self.onroute = 0
        else:
            self.onroute = max(self.overtime - self.offroute, 0)

    def format_otherroute(self):
        """ format the self.other route so that if fits in the spreadsheet cell. """
        if self.otherroute_array:
            if len(self.otherroute_array) > 1:
                # format like "1024 + 1"
                self.otherroute = self.otherroute_array[0] + "+" + str(len(self.otherroute_array) - 1)
            else:
                # format like "1024"
                self.otherroute = self.otherroute_array[0]

    def calc_availability(self):
        """ calculate otdl and aux availability """
        if not self.totalhours and self.codes == "no call":  # if the carrier was not scheduled for the day
            self.avail_10 = 10  # otdl/aux availability to 10 hours
            self.avail_115 = 11.5  # aux availability to 11.50 hours
            self.avail_12 = 12  # otdl availability to 12 hours.
            return
        if self.codes in ("light", "excused", "sch chg", "annual", "sick"):  # if carrier excused for day
            self.avail_10 = 0  # otdl/aux availability to 0 hours
            self.avail_115 = 0  # aux availability to 0 hours
            self.avail_12 = 0  # otdl availability to 0 hours.
            return
        if not self.totalhours:
            return
        self.avail_10 = max(10 - self.totalhours, 0)
        self.avail_115 = max(11.5 - self.totalhours, 0)
        self.avail_12 = max(12 - self.totalhours, 0)

    def qualify(self):
        """ check to see if the carrier information needs to be displayed. """
        if self.pref[self.lsi] in ("otdl", "aux"):  # display all for otdl and aux
            return True
        if self.display_limiter == "show all":  # display all if the limiter is set to "show all"
            return True
        if self.display_limiter == "only workdays":  # display only days when the carrier worked.
            if self.totalhours:
                return True
            return False
        if self.display_limiter == "only mandates":
            if self.pref[self.lsi] == "nl":
                if self.overtime or self.offroute_adj:
                    return True
                return False
            if self.pref[self.lsi] == "wal":
                if self.offroute_adj:
                    return True
                return False

    def display_recs(self):
        """ put the carrier and the first part of rings into the spreadsheet - it's show time! """
        cell = self.ws_list[self.i].cell(row=self.row, column=1)  # row number
        cell.value = "{}.".format(self.row_number)
        cell.style = self.input_s
        self.row_number += 1  # increment the row number
        cell = self.ws_list[self.i].cell(row=self.row, column=2)  # name
        cell.value = self.carrier
        cell.style = self.input_name
        self.ws_list[self.i].merge_cells('B' + str(self.row) + ':D' + str(self.row))
        cell = self.ws_list[self.i].cell(row=self.row, column=5)  # list status
        cell.value = self.list_dict[self.list_]
        cell.style = self.input_s
        self.ws_list[self.i].merge_cells('E' + str(self.row) + ':F' + str(self.row))
        cell = self.ws_list[self.i].cell(row=self.row, column=7)  # list status
        cell.value = self.route
        cell.style = self.input_s
        cell = self.ws_list[self.i].cell(row=self.row, column=8)  # begin tour
        cell.value = Convert(self.bt).str_to_floatoremptystr()
        cell.style = self.input_s
        cell.number_format = "#,###.00;[RED]-#,###.00"
        cell = self.ws_list[self.i].cell(row=self.row, column=9)  # move
        cell.value = Convert(self.moves).str_to_floatoremptystr()
        cell.style = self.input_s
        cell.number_format = "#,###.00;[RED]-#,###.00"
        cell = self.ws_list[self.i].cell(row=self.row, column=10)  # return to station
        cell.value = Convert(self.rs).str_to_floatoremptystr()
        cell.style = self.input_s
        cell.number_format = "#,###.00;[RED]-#,###.00"
        cell = self.ws_list[self.i].cell(row=self.row, column=11)  # end tour
        cell.value = Convert(self.et).str_to_floatoremptystr()
        cell.style = self.input_s
        cell.number_format = "#,###.00;[RED]-#,###.00"
        cell = self.ws_list[self.i].cell(row=self.row, column=12)  # overtime worked
        cell.value = Convert(self.overtime).str_to_floatoremptystr()
        cell.style = self.input_s
        cell.number_format = "#,###.00;[RED]-#,###.00"
        self.ws_list[self.i].merge_cells('L' + str(self.row) + ':M' + str(self.row))
        column = 14
        if self.pref[self.lsi] == "nl":
            cell = self.ws_list[self.i].cell(row=self.row, column=column)  # on route
            cell.value = Convert(self.onroute).str_to_floatoremptystr()
            cell.style = self.input_s
            cell.number_format = "#,###.00;[RED]-#,###.00"
            column += 1
        if self.pref[self.lsi] in ("nl", "wal"):
            cell = self.ws_list[self.i].cell(row=self.row, column=column)  # off route
            cell.value = Convert(self.offroute_adj).str_to_floatoremptystr()
            cell.style = self.input_s
            cell.number_format = "#,###.00;[RED]-#,###.00"
            column += 1
        if self.pref[self.lsi] in ("nl", "wal"):
            cell = self.ws_list[self.i].cell(row=self.row, column=column)  # other route
            cell.value = self.otherroute
            cell.style = self.input_s
            cell.number_format = "#,###.00;[RED]-#,###.00"
        if self.pref[self.lsi] in ("otdl", "aux"):
            cell = self.ws_list[self.i].cell(row=self.row, column=column)  # availability to 10
            cell.value = self.avail_10
            cell.style = self.input_s
            cell.number_format = "#,###.00;[RED]-#,###.00"
            column += 1
            if self.pref[self.lsi] == "otdl":  # change value dependant on otdl or aux
                value = self.avail_12
            else:
                value = self.avail_115
            cell = self.ws_list[self.i].cell(row=self.row, column=column)  # availability to 12 or 11.5
            cell.value = value
            cell.style = self.input_s
            cell.number_format = "#,###.00;[RED]-#,###.00"
            column += 1
            formula = "=IF(%s!J%s = \"\", \"\", IF(%s!C%s = \"\", \"no dov\", MAX(%s!C%s-%s!J%s, 0)))" % \
                      (self.day_of_week[self.i], str(self.row),
                       self.day_of_week[self.i], "16",
                       self.day_of_week[self.i], "16",
                       self.day_of_week[self.i], str(self.row))
            cell = self.ws_list[self.i].cell(row=self.row, column=column)  # availability to DOV
            cell.value = formula
            cell.style = self.input_s
            cell.number_format = "#,###.00;[RED]-#,###.00"
        self.get_last_row()  # record the number of the last row for total formulas in footers

    def get_last_row(self):
        """ record the number of the last row for totals formulas in footers """
        self.last_row = self.row

    def get_listrange(self):
        """ get the first row and last row of the current list and store them into an array called listrange.
        later, the summary row will be added."""
        self.listrange = []
        self.listrange.append(self.first_row)
        self.listrange.append(self.last_row)
        self.listrange.append(self.row)  # put the 3rd and final row number into the listrange - summary row

    def build_footer(self):
        """ call the footer depending on the list. """
        if self.pref[self.lsi] == "nl":
            self.nl_footer()
        elif self.pref[self.lsi] == "wal":
            self.wal_footer()
        elif self.pref[self.lsi] == "otdl":
            self.otdl_footer()
        else:
            self.aux_footer()
        self.row += 1

    def nl_footer(self):
        """ build the non list footer. """
        self.ws_list[self.i].row_dimensions[self.row].height = 20  # adjust row height
        cell = self.ws_list[self.i].cell(row=self.row, column=2)  # totals label
        cell.value = "     Totals:"
        cell.style = self.footer_left
        self.ws_list[self.i].merge_cells('B' + str(self.row) + ':M' + str(self.row))
        cell = self.ws_list[self.i].cell(row=self.row, column=14)  # totals own route
        formula = "=SUM(%s!N%s:N%s)" % (self.day_of_week[self.i], self.listrange[0], self.listrange[1])
        cell.value = formula
        cell.style = self.footer_mid
        cell.number_format = "#,###.00;[RED]-#,###.00"
        cell = self.ws_list[self.i].cell(row=self.row, column=15)  # totals off route
        formula = "=SUM(%s!O%s:O%s)" % (self.day_of_week[self.i], self.listrange[0], self.listrange[1])
        cell.value = formula
        cell.style = self.footer_mid
        cell.number_format = "#,###.00;[RED]-#,###.00"
        cell = self.ws_list[self.i].cell(row=self.row, column=16)  # blank formated cell
        cell.value = ""
        cell.style = self.footer_right

    def wal_footer(self):
        """ build the work assignment footer """
        self.ws_list[self.i].row_dimensions[self.row].height = 20  # adjust row height
        cell = self.ws_list[self.i].cell(row=self.row, column=2)  # totals label
        cell.value = "     Totals:"
        cell.style = self.footer_left
        self.ws_list[self.i].merge_cells('B' + str(self.row) + ':M' + str(self.row))

        cell = self.ws_list[self.i].cell(row=self.row, column=14)  # totals off route
        formula = "=SUM(%s!N%s:N%s)" % (self.day_of_week[self.i], self.listrange[0], self.listrange[1])
        cell.value = formula
        cell.style = self.footer_mid
        cell.number_format = "#,###.00;[RED]-#,###.00"
        cell = self.ws_list[self.i].cell(row=self.row, column=15)  # blank formated cell
        cell.style = self.footer_right

    def aux_footer(self):
        """ build the footer for auxiliary - cca and ptf - carriers. """
        self.ws_list[self.i].row_dimensions[self.row].height = 20  # adjust row height
        cell = self.ws_list[self.i].cell(row=self.row, column=2)  # totals label
        cell.value = "     Totals:"
        cell.style = self.footer_left
        self.ws_list[self.i].merge_cells('B' + str(self.row) + ':M' + str(self.row))
        cell = self.ws_list[self.i].cell(row=self.row, column=14)  # availability to 10
        formula = "=SUM(%s!N%s:N%s)" % (self.day_of_week[self.i], self.listrange[0], self.listrange[1])
        cell.value = formula
        cell.style = self.footer_mid
        cell.number_format = "#,###.00;[RED]-#,###.00"
        cell = self.ws_list[self.i].cell(row=self.row, column=15)  # availability to 11.50
        formula = "=SUM(%s!O%s:O%s)" % (self.day_of_week[self.i], self.listrange[0], self.listrange[1])
        cell.value = formula
        cell.style = self.footer_mid
        cell.number_format = "#,###.00;[RED]-#,###.00"
        cell = self.ws_list[self.i].cell(row=self.row, column=16)  # availability to DOV
        formula = "=SUM(%s!P%s:P%s)" % (self.day_of_week[self.i], self.listrange[0], self.listrange[1])
        cell.value = formula
        cell.style = self.footer_right
        cell.number_format = "#,###.00;[RED]-#,###.00"

    def otdl_footer(self):
        """ build the overtime desired list footer. """
        self.ws_list[self.i].row_dimensions[self.row].height = 20  # adjust row height
        cell = self.ws_list[self.i].cell(row=self.row, column=2)  # totals label
        cell.value = "     Totals:"
        cell.style = self.footer_left
        self.ws_list[self.i].merge_cells('B' + str(self.row) + ':M' + str(self.row))
        cell = self.ws_list[self.i].cell(row=self.row, column=14)  # availability to 10
        formula = "=SUM(%s!N%s:N%s)" % (self.day_of_week[self.i], self.listrange[0], self.listrange[1])
        cell.value = formula
        cell.style = self.footer_mid
        cell.number_format = "#,###.00;[RED]-#,###.00"
        cell = self.ws_list[self.i].cell(row=self.row, column=15)  # availability to 12
        formula = "=SUM(%s!O%s:O%s)" % (self.day_of_week[self.i], self.listrange[0], self.listrange[1])
        cell.value = formula
        cell.style = self.footer_mid
        cell.number_format = "#,###.00;[RED]-#,###.00"
        cell = self.ws_list[self.i].cell(row=self.row, column=16)  # availability to DOV
        formula = "=SUM(%s!P%s:P%s)" % (self.day_of_week[self.i], self.listrange[0], self.listrange[1])
        cell.value = formula
        cell.style = self.footer_right
        cell.number_format = "#,###.00;[RED]-#,###.00"

    def fill_quads(self):
        """ write formulas for the quadrants on the cover sheet. %s!J%s"""
        formula = "=%s!N%s" % (self.day_of_week[self.i], self.dayrange[0][2])  # cell N+nl summary
        self.cellc9.value = formula  # non otdl own route violations
        formula = "=%s!O%s" % (self.day_of_week[self.i], self.dayrange[0][2])  # cell O+nl summary
        self.cellf9.value = formula  # non otdl off route violations
        formula = "=%s!N%s" % (self.day_of_week[self.i], self.dayrange[1][2])  # cell O+wal summary
        self.cellf11.value = formula  # wal off route violations
        formula = "=%s!N%s" % (self.day_of_week[self.i], self.dayrange[2][2])  # cell N+aux summary
        self.cellj9.value = formula  # aux availability to 10 hours
        formula = "=%s!O%s" % (self.day_of_week[self.i], self.dayrange[2][2])  # cell O+aux summary
        self.cellm9.value = formula  # aux availability to 11.5 hours
        formula = "=%s!N%s" % (self.day_of_week[self.i], self.dayrange[3][2])  # cell N+otdl summary
        self.cellj11.value = formula  # otdl availability to 10 hours
        formula = "=%s!O%s" % (self.day_of_week[self.i], self.dayrange[3][2])  # cell O+otdl summary
        self.cellm11.value = formula  # otdl availability to 12 hours
        formula = "=SUM(%s!P%s + %s!P%s)" \
                  % (self.day_of_week[self.i], self.dayrange[2][2],  # cell P+aux summary
                     self.day_of_week[self.i], self.dayrange[3][2])  # cell P+otdl summary
        self.celln16.value = formula
        formula = "=MAX(" \
                  "COUNTIF(%s!J%s:J%s, \">\"&%s!C%s) + " \
                  "COUNTIF(%s!J%s:J%s, \">\"&%s!C%s) + " \
                  "COUNTIF(%s!J%s:J%s, \">\"&%s!C%s) + " \
                  "COUNTIF(%s!J%s:J%s, \">\"&%s!C%s)," \
                  "0)" \
                  % (self.day_of_week[self.i], self.dayrange[0][0], self.dayrange[0][1], self.day_of_week[self.i], "16",
                     self.day_of_week[self.i], self.dayrange[1][0], self.dayrange[1][1], self.day_of_week[self.i], "16",
                     self.day_of_week[self.i], self.dayrange[2][0], self.dayrange[2][1], self.day_of_week[self.i], "16",
                     self.day_of_week[self.i], self.dayrange[3][0], self.dayrange[3][1], self.day_of_week[self.i], "16")
        self.cellf16.value = formula  # carriers out past dispatch of value

    def pagebreak(self, force=False):
        """ create a page break if consistant with user preferences. If page break is True, then the page
        break can not be skipped. This ensures that there is always a page break after the summary page."""
        if self.pref[self.lsi] == "nl" and not self.pb4_nl_wal:
            if not force:
                self.row += 1
                return
        if self.pref[self.lsi] == "wal" and not self.pb4_wal_aux:
            if not force:
                self.row += 1
                return
        if self.pref[self.lsi] == "aux" and not self.pb4_aux_otdl:
            if not force:
                self.row += 1
                return
        if self.pref[self.lsi] == "otdl":
            if not force:
                self.row += 1
                return
        try:
            self.ws_list[self.i].page_breaks.append(Break(id=self.row))
        except AttributeError:
            self.ws_list[self.i].row_breaks.append(Break(id=self.row))  # effective for windows
        self.row += 1

    def increment_progbar(self):
        """ move the progress bar, update with info on what is being done """
        lst = ("No List", "Work Assignment", "Overtime Desired", "Auxiliary")
        self.pbi += 1
        self.pb.move_count(self.pbi)  # increment progress bar
        self.pb.change_text("Building day {}: list: {}".format(self.day.strftime("%A"), lst[self.lsi]))

    def save_open(self):
        """ name and open the excel file """
        self.pbi += 1
        self.pb.move_count(self.pbi)  # increment progress bar
        self.pb.change_text("Saving...")
        self.pb.stop()
        r = "_w"
        if not projvar.invran_weekly_span:  # if investigation range is daily
            r = "_d"
        xl_filename = "man4" + str(format(self.dates[0], "_%y_%m_%d")) + r + ".xlsx"
        try:
            self.wb.save(dir_path('mandates_4') + xl_filename)
            messagebox.showinfo("Spreadsheet generator",
                                "Your spreadsheet was successfully generated. \n"
                                "File is named: {}".format(xl_filename),
                                parent=self.frame)
            if sys.platform == "win32":
                os.startfile(dir_path('mandates_4') + xl_filename)
            if sys.platform == "linux":
                subprocess.call(["xdg-open", 'kb_sub/mandates_4/' + xl_filename])
            if sys.platform == "darwin":
                subprocess.call(["open", dir_path('mandates_4') + xl_filename])
        except PermissionError:
            messagebox.showerror("Spreadsheet generator",
                                 "The spreadsheet was not opened. \n"
                                 "Suggestion: "
                                 "Make sure that identically named spreadsheets are closed "
                                 "(the file can't be overwritten while open).",
                                 parent=self.frame)


class OffbidSpreadsheet:
    """
    Create a spreadsheet for calculating and detecting situations where carriers do no get 8 hours of work
    on their own bid assignments due to off route assignments.
    """
    def __init__(self):
        self.frame = None  # the frame of parent
        self.pb = None  # progress bar object
        self.pbi = 0  # progress bar count index
        self.startdate = None  # start date of the investigation
        self.enddate = None  # ending date of the investigation
        self.dates = []  # all days of the investigation
        self.carrierlist = []  # all carriers in carrier list
        self.wb = None  # the workbook object
        self.row = 1
        self.violation_number = 0
        self.ws_header = None  # style
        self.date_dov = None  # style
        self.date_dov_title = None  # style
        self.name_header = None  # style
        self.col_header = None  # style
        self.input_name = None  # style
        self.input_s = None  # style
        self.calcs = None  # style
        self.list_header = None  # style
        self.offbid = None  # worksheet for the analysis of off bid violations
        # self.instructions = None  # worksheet for instructions
        self.carrier = ""  # carrier name
        self.route = ""  # carrier route
        self.rings = []  # carrier rings queried from database
        self.totalhours = ""  # carrier rings - 5200 time
        self.codes = ""  # carrier rings - code/note
        self.moves = ""  # carrier rings - moves on and off route with route
        self.move_i = 0  # increments rows for multiple move functionality
        self.i = 0  # the day being investigated as a number 0 - 6.
        self.max_pivot = 0.0  # the maximum allowed pivot.
        self.distinct_pages = None
        self.move_set = 0  # extra rows used by the multiple moves display
        self.no_qualifiers = True  # is True as long as no violations have been found for any carriers.

    def create(self, frame):
        """ a master method for running the other methods in proper order. """
        self.frame = frame
        if not self.ask_ok():  # abort if user selects cancel from askokcancel
            return
        self.pb = ProgressBarDe(label="Building Off Bid Assignment Spreadsheet")
        self.pb.max_count(1000)  # set length of progress bar
        self.pb.start_up()  # start the progress bar
        self.pbi = 1
        self.pb.move_count(self.pbi)  # increment progress bar
        self.pb.change_text("Initializing... ")
        self.get_dates()
        self.get_carrierlist()
        self.get_pb_max_count()
        self.get_maxpivot_distinctpage()
        self.get_styles()
        self.build_workbook()
        self.carrierloop()
        if self.no_qualifiers:
            self.no_violations()
        self.save_open()

    def ask_ok(self):
        """ ends process if user cancels """
        if messagebox.askokcancel("Spreadsheet generator",
                                  "Do you want to generate an \n"
                                  "Off Bid Assignment Spreadsheet?",
                                  parent=self.frame):
            return True
        return False

    def get_dates(self):
        """ get the dates from the project variables """
        self.startdate = projvar.invran_date  # set daily investigation range as default - get start date
        self.enddate = projvar.invran_date  # get end date
        self.dates = [projvar.invran_date, ]  # create an array of days - only one day if daily investigation range
        if projvar.invran_weekly_span:  # if the investigation range is weekly
            date = projvar.invran_date_week[0]
            self.startdate = projvar.invran_date_week[0]
            self.enddate = projvar.invran_date_week[6]
            self.dates = []
            for _ in range(7):  # create an array with all the days in the weekly investigation range
                self.dates.append(date)
                date += timedelta(days=1)

    def get_carrierlist(self):
        """ get record sets for all carriers """
        self.carrierlist = CarrierList(self.startdate, self.enddate, projvar.invran_station).get()

    def get_maxpivot_distinctpage(self):
        """ get the maximum pivot and distinct page value from the database. """
        sql = "SELECT tolerance FROM tolerances WHERE category = 'offbid_maxpivot'"
        result = inquire(sql)
        self.max_pivot = float(result[0][0])
        sql = "SELECT tolerance FROM tolerances WHERE category = 'offbid_distinctpage'"
        result = inquire(sql)
        self.distinct_pages = Convert(result[0][0]).str_to_bool()

    def get_pb_max_count(self):
        """ set length of progress bar """
        self.pb.max_count(len(self.carrierlist)+1)  # once for each list in each day, plus reference, summary and saving

    def get_styles(self):
        """ Named styles for workbook """
        bd = Side(style='thin', color="80808080")  # defines borders
        self.ws_header = NamedStyle(name="ws_header", font=Font(bold=True, name='Arial', size=12))
        self.date_dov = NamedStyle(name="date_dov", font=Font(name='Arial', size=8))
        self.date_dov_title = NamedStyle(name="date_dov_title", font=Font(bold=True, name='Arial', size=8),
                                         alignment=Alignment(horizontal='right'))
        self.name_header = NamedStyle(name="name_header", font=Font(bold=True, name='Arial', size=8, color='666666'),
                                      alignment=Alignment(horizontal='right'))
        self.col_header = NamedStyle(name="col_header", font=Font(bold=True, name='Arial', size=8, color='666666'),
                                     alignment=Alignment(horizontal='center'),
                                     border=Border(left=bd, top=bd, right=bd, bottom=bd))
        self.input_name = NamedStyle(name="input_name", font=Font(bold=True, name='Arial', size=10),
                                     alignment=Alignment(horizontal='left'))
        self.input_s = NamedStyle(name="input_s", font=Font(name='Arial', size=8),
                                  border=Border(left=bd, top=bd, right=bd, bottom=bd),
                                  alignment=Alignment(horizontal='right'))
        self.calcs = NamedStyle(name="calcs", font=Font(name='Arial', size=8),
                                border=Border(left=bd, top=bd, right=bd, bottom=bd),
                                fill=PatternFill(fgColor='e5e4e2', fill_type='solid'),
                                alignment=Alignment(horizontal='right'))
        self.list_header = NamedStyle(name="list_header", font=Font(bold=True, name='Arial', size=10))

    def build_workbook(self):
        """ creates the workbook object """
        self.wb = Workbook()  # define the workbook
        self.offbid = self.wb.active  # create first worksheet
        self.offbid.title = "off bid"  # title first worksheet
        self.offbid.oddFooter.center.text = "&A"
        # self.instructions = self.wb.create_sheet("instructions")

    def carrierloop(self):
        """ loop for each carrier """
        for carrier in self.carrierlist:
            self.carrier = carrier[0][1]  # current iteration of carrier list is assigned self.carrier
            self.pbi += 1
            self.pb.move_count(self.pbi)  # increment progress bar
            self.pb.change_text("Checking: {}".format(self.carrier))
            self.route = carrier[0][4]  # get the route of the carrier
            self.get_rings()  # get individual carrier rings for the day - define self.rings
            if self.qualify():  # test the rings to see if there is a violation during the week
                self.no_qualifiers = False
                self.conditional_header()  # insert the header if proper conditions apply
                self.violation_number += 1  # increment the row number
                self.display_recs()  # build the carrier and the rings row into the spreadsheet
                self.conditional_pagebreak()

    def conditional_header(self):
        """ insert the header on certain conditions"""
        if self.violation_number == 0 or self.distinct_pages:
            self.build_headers()

    def build_headers(self):
        """ worksheet headers """
        cell = self.offbid.cell(row=self.row, column=1)
        cell.value = "Off Bid Assignment Worksheet"
        cell.style = self.ws_header
        self.offbid.merge_cells('A' + str(self.row) + ':E' + str(self.row))
        self.row += 2
        cell = self.offbid.cell(row=self.row, column=1)
        cell.value = "Date:  "  # create date/ pay period/ station header
        cell.style = self.date_dov_title
        cell = self.offbid.cell(row=self.row, column=2)
        date_string = self.dates[0].strftime("%x")
        # The date can be one day or a service week (a range of 7 days)
        if len(self.dates) > 1:
            date_string = self.dates[0].strftime("%x") + " - " + self.dates[6].strftime("%x")
        cell.value = date_string  # fill in the date/s
        cell.style = self.date_dov
        self.offbid.merge_cells('B' + str(self.row) + ':D' + str(self.row))
        cell = self.offbid.cell(row=self.row, column=5)
        cell.value = "Pay Period:  "
        cell.style = self.date_dov_title
        self.offbid.merge_cells('E' + str(self.row) + ':F' + str(self.row))
        cell = self.offbid.cell(row=self.row, column=7)
        cell.value = projvar.pay_period
        cell.style = self.date_dov
        self.offbid.merge_cells('G' + str(self.row) + ':H' + str(self.row))
        self.row += 1
        cell = self.offbid.cell(row=self.row, column=1)
        cell.value = "Station:  "
        cell.style = self.date_dov_title
        cell = self.offbid.cell(row=self.row, column=2)
        cell.value = projvar.invran_station
        cell.style = self.date_dov
        self.offbid.merge_cells('B' + str(self.row) + ':D' + str(self.row))
        self.row += 2

    def get_rings(self):
        """ get individual carrier rings for the day - define self.rings"""
        self.rings = []
        for date in self.dates:
            rings = Rings(self.carrier, date).get_for_day()  # assign as rings
            totalhours = 0.0  # set default as an empty string
            codes = ""
            moves = ""
            if rings[0]:  # if rings record is not blank
                totalhours = float(Convert(rings[0][2]).zero_not_empty())
                codes = rings[0][4]
                moves = rings[0][5]
            to_add = [date, totalhours, codes, moves]
            self.rings.append(to_add)

    def qualify(self):
        """ test to see if the carrier/rings need to be displayed. """
        qualify = False
        for i in range(len(self.rings)):  # loops for each day in the investigation range
            if self.number_crunching(i):  # returns True if there is a violation
                qualify = True
            else:  # if there is no violations for the day - insert False into self.rings
                self.rings[i].append(False)
        if qualify:  # if there is a violation for at least one day
            return True
        return False

    def number_crunching(self, i):
        """ do calculations to determine off route, on route and violation values.
            returns True if there is a violation. Adds violation boolean to self.rings. """
        offroute = 0.0  # this is the total time spent off the carrier's route
        if not self.rings[i][1]:  # if the total hours is zero - the violation is zero
            return False
        if self.rings[i][2] == "ns day":  # if it is the carrier's ns day - violation is zero
            return False
        if not self.rings[i][3]:  # if the moves is empty, then the violation is zero
            return False
        index = 0  # set the index to 1. This will point to an element in the moves array.
        totalhours = self.rings[i][1]  # simplify the variable name
        moves = Convert(self.rings[i][3]).string_to_array()  # simplify the variable name
        while index < len(moves):  # calculate the total time off route
            offroute += float(moves[index+1]) - float(moves[index])
            index += 3
        ownroute = max(totalhours - offroute, 0)   # calculate the total time spent on route
        violation = max(8 - ownroute, 0)  # calculate the total violation
        if violation > self.max_pivot:
            self.rings[i].append(True)
            return True
        return False

    def display_recs(self):
        """ build the carrier and ring recs into the spreadsheet. """
        cell = self.offbid.cell(row=self.row, column=1)  # carrier label
        cell.value = "carrier:  "
        cell.style = self.name_header
        cell = self.offbid.cell(row=self.row, column=2)  # carrier name input
        cell.value = self.carrier
        cell.style = self.input_name
        self.offbid.merge_cells('B' + str(self.row) + ':E' + str(self.row))
        cell = self.offbid.cell(row=self.row, column=6)  # route label
        cell.value = "route:  "
        cell.style = self.name_header
        cell = self.offbid.cell(row=self.row, column=7)  # route input
        cell.value = self.route
        cell.style = self.input_name
        self.offbid.merge_cells('G' + str(self.row) + ':J' + str(self.row))
        self.row += 1
        # use loops and an array to build the column headers
        column_headers = ("day", "date", "5200", "mv off", "mv on", "route", "off rt", "on rt", "violation")
        for i in range(9):
            cell = self.offbid.cell(row=self.row, column=i + 2)  # column headers
            cell.value = column_headers[i]
            cell.style = self.col_header
        self. row += 1
        self.display_daily()  # create a row/s to display the daily information on the violation
        self.row += 1

    def display_daily(self):
        """ display the daily ring recs for the carrier. """
        for i in range(len(self.rings)):
            if self.rings[i][4]:
                cell = self.offbid.cell(row=self.row, column=2)  # day
                cell.value = self.rings[i][0].strftime("%a")
                cell.style = self.input_s
                cell = self.offbid.cell(row=self.row, column=3)  # date
                cell.value = self.rings[i][0].strftime("%m/%d/%Y")
                cell.style = self.input_s
                cell = self.offbid.cell(row=self.row, column=4)  # 5200
                cell.value = self.rings[i][1]
                cell.style = self.input_s
                cell.number_format = "#,###.00;[RED]-#,###.00"
                self.display_moves(i)
                cell = self.offbid.cell(row=self.row, column=9)  # on route
                formula = "=MAX(D" + str(self.row) + "-H" + str(self.row) + ",0)"
                cell.value = formula
                cell.style = self.calcs
                cell.number_format = "#,###.00;[RED]-#,###.00"
                cell = self.offbid.cell(row=self.row, column=10)  # violation
                formula = "=IF(AND(D" + str(self.row) + ">0,H" + str(self.row) + ">0),8-I" + str(self.row) + ",0)"
                cell.value = formula
                cell.style = self.calcs
                cell.number_format = "#,###.00;[RED]-#,###.00"
                self.row += (self.move_set - 1)  # correct for increment after last move set.
                self.row += 1

    def display_moves(self, i):
        """ display the moves. include contingencies for multiple moves.
            also displays the off route formula column as that changes with multiple moves"""
        moves = Convert(self.rings[i][3]).string_to_array()
        set_count = Moves().count_movesets(moves)
        if len(moves) > 3:
            moves = ["*", "*", "*"] + moves
        move_place = 0
        self.move_set = 0  # extra rows used to display multiple moves/ incremented in self.display_moves()
        for move in moves:
            cell = self.offbid.cell(row=self.row + self.move_set, column=5 + move_place)  # move off
            cell.value = move
            cell.style = self.input_s
            if move_place == 2:
                formulacell = self.offbid.cell(row=self.row + self.move_set, column=8)  # formula cell
                formulacell.style = self.calcs
                if not self.move_set and set_count > 1:  # if this is the first row of a multiple row
                    formula = "=SUM(H" + str(self.row + 1) + ":H" + str(self.row + set_count) + ")"
                else:
                    formula = "=SUM(F" + str(self.row + self.move_set) + "-E" + str(self.row + self.move_set) + ")"
                formulacell.value = formula
                formulacell.number_format = "#,###.00;[RED]-#,###.00"
                move_place = 0
                self.move_set += 1
            else:
                move_place += 1
                cell.number_format = "#,###.00;[RED]-#,###.00"

    def conditional_pagebreak(self):
        """ insert a page break if the correct conditions apply """
        if self.distinct_pages:
            try:
                self.offbid.page_breaks.append(Break(id=self.row))
                self.row += 1
            except AttributeError:
                self.offbid.row_breaks.append(Break(id=self.row))  # effective for windows
                self.row += 1

    def no_violations(self):
        """ if self.no_qualifiers is True after all carriers have been checked, this will display a message
            saying that no violations occured. """
        self.build_headers()
        cell = self.offbid.cell(row=self.row, column=2)  # No off bid violation found
        cell.value = "No off bid violations found"
        cell.style = self.list_header
        self.offbid.merge_cells('B' + str(self.row) + ':J' + str(self.row))

    def save_open(self):
        """ name and open the excel file """
        self.pbi += 1
        self.pb.move_count(self.pbi)  # increment progress bar
        self.pb.change_text("Saving...")
        self.pb.stop()
        r = "_w"
        if not projvar.invran_weekly_span:  # if investigation range is daily
            r = "_d"
        xl_filename = "kb_offbid_" + str(format(self.dates[0], "_%y_%m_%d")) + r + ".xlsx"
        try:
            self.wb.save(dir_path('off_bid') + xl_filename)
            messagebox.showinfo("Spreadsheet generator",
                                "Your spreadsheet was successfully generated. \n"
                                "File is named: {}".format(xl_filename),
                                parent=self.frame)
            if sys.platform == "win32":
                os.startfile(dir_path('off_bid') + xl_filename)
            if sys.platform == "linux":
                subprocess.call(["xdg-open", 'kb_sub/off_bid/' + xl_filename])
            if sys.platform == "darwin":
                subprocess.call(["open", dir_path('off_bid') + xl_filename])
        except PermissionError:
            messagebox.showerror("Spreadsheet generator",
                                 "The spreadsheet was not opened. \n"
                                 "Suggestion: "
                                 "Make sure that identically named spreadsheets are closed "
                                 "(the file can't be overwritten while open).",
                                 parent=self.frame)


class OtAvailSpreadsheet:
    """ this will generate a spreadsheet that will display the hours and paid leave an otdl carrier has worked, 
    and use this to tally how much availability that carrier has day to day. """
    def __init__(self):
        self.frame = None
        self.pb = None  # progress bar object
        self.pbi = 0  # progress bar count index
        self.carrier_list = []  # build a carrier list
        self.ot_carrier = None
        self.wb = None  # workbook object
        self.availability = None  # workbook object sheet
        self.startdate = None  # start of the investigation range
        self.enddate = None  # ending day of the investigation range
        self.dates = []  # all the dates of the investigation range
        self.rings = []  # all rings for all carriers in the carrier list
        self.ws_header = None  # style
        self.date_dov = None  # style
        self.date_dov_title = None  # style
        self.name_header = None  # style
        self.col_header = None  # style
        self.input_name = None  # style
        self.input_s = None  # style
        self.calcs = None  # style
        self.list_header = None  # style
        self.violation_recsets = []  # carrier info, daily hours, leavetypes and leavetimes
        self.row = 1

    def create(self, frame):
        """ master method for calling methods"""
        self.frame = frame
        if not self.ask_ok():
            return
        self.get_dates()
        self.get_carrierlist()
        self.pb = ProgressBarDe(label="OTDL Weekly Availability Spreadsheet")
        self.pb.max_count(len(self.carrier_list))  # set length of progress bar
        self.pb.start_up()  # start the progress bar
        self.pbi = 1
        self.pb.move_count(self.pbi)  # increment progress bar
        self.pb.change_text("Gathering Data... ")
        self.build_workbook()
        self.get_styles()
        self.set_dimensions()
        self.build_headers()
        self.carrierloop()
        self.save_open()

    def ask_ok(self):
        """ continue if user selects ok. """
        if messagebox.askokcancel("Spreadsheet generator",
                                  "Do you want to generate a spreadsheet for OTDL availability?",
                                  parent=self.frame):
            return True
        return False
        
    def get_dates(self):
        """ get the dates of the investigation range from the project variables. """
        date = projvar.invran_date_week[0]
        self.startdate = projvar.invran_date_week[0]
        self.enddate = projvar.invran_date_week[6]
        for _ in range(7):
            self.dates.append(date)
            date += timedelta(days=1)

    def get_carrierlist(self):
        """ call the carrierlist class from kbtoolbox module to get the carrier list """
        carrierlist = CarrierList(self.startdate, self.enddate, projvar.invran_station).get()
        for carrier in carrierlist:
            for rec in carrier:
                if rec[2] == "otdl":
                    self.carrier_list.append(carrier[0])  # add record for each otdl carrier in carrier list
                    break

    def get_styles(self):
        """ Named styles for workbook """
        bd = Side(style='thin', color="80808080")  # defines borders
        self.ws_header = NamedStyle(name="ws_header", font=Font(bold=True, name='Arial', size=12))
        self.date_dov = NamedStyle(name="date_dov", font=Font(name='Arial', size=8))
        self.date_dov_title = NamedStyle(name="date_dov_title", font=Font(bold=True, name='Arial', size=8),
                                         alignment=Alignment(horizontal='right'))
        self.name_header = NamedStyle(name="name_header", font=Font(bold=True, name='Arial', size=8, color='666666'),
                                      alignment=Alignment(horizontal='right'))
        self.col_header = NamedStyle(name="col_header", font=Font(bold=True, name='Arial', size=8, color='666666'),
                                     alignment=Alignment(horizontal='center'),
                                     border=Border(left=bd, top=bd, right=bd, bottom=bd))
        self.input_name = NamedStyle(name="input_name", font=Font(bold=True, name='Arial', size=10),
                                     alignment=Alignment(horizontal='left'))
        self.input_s = NamedStyle(name="input_s", font=Font(name='Arial', size=8),
                                  border=Border(left=bd, top=bd, right=bd, bottom=bd),
                                  alignment=Alignment(horizontal='right'))
        self.calcs = NamedStyle(name="calcs", font=Font(name='Arial', size=8),
                                border=Border(top=bd, right=bd, bottom=bd, left=bd),
                                fill=PatternFill(fgColor='e5e4e2', fill_type='solid'),
                                alignment=Alignment(horizontal='right'))
        self.list_header = NamedStyle(name="list_header", font=Font(bold=True, name='Arial', size=10))

    def build_headers(self):
        """ self.availability worksheet header - format cells """
        self.pb.change_text("Building Availability...")
        self.availability.merge_cells('A1:O1')
        self.availability['A1'] = "OTDL Weekly Availability Worksheet"
        self.availability['A1'].style = self.ws_header
        self.availability['A3'] = "Date:"  # date label
        self.availability['A3'].style = self.date_dov_title
        self.availability.merge_cells('B3:H3')  # blank field for date
        self.availability['B3'] = self.dates[0].strftime("%x") + " - " + self.dates[6].strftime("%x")
        self.availability['B3'].style = self.date_dov
        self.availability.merge_cells('I3:L3')  # pay period label
        self.availability['I3'] = "Pay Period:"
        self.availability['I3'].style = self.date_dov_title  # blank field for pay period
        self.availability.merge_cells('M3:O3')
        self.availability['M3'] = projvar.pay_period
        self.availability['M3'].style = self.date_dov
        self.availability['A4'] = "Station:"  # station label
        self.availability['A4'].style = self.date_dov_title
        self.availability.merge_cells('B4:O4')  # blank field for station
        self.availability['B4'] = projvar.invran_station
        self.availability['B4'].style = self.date_dov

    def build_workbook(self):
        """ creates the workbook object """
        self.pb.change_text("Building workbook...")
        self.wb = Workbook()  # define the workbook
        self.availability = self.wb.active  # create first worksheet
        self.availability.title = "availability"  # title first worksheet
        self.availability.oddFooter.center.text = "&A"
        
    def set_dimensions(self):
        """ adjust the height and width on the violations/ instructions page """
        for x in range(2, 4):
            self.availability.row_dimensions[x].height = 10  # adjust all row height
        sheets = (self.availability, )
        for sheet in sheets:
            sheet.column_dimensions["A"].width = 16
            sheet.column_dimensions["B"].width = 6
            sheet.column_dimensions["C"].width = 2
            sheet.column_dimensions["D"].width = 6
            sheet.column_dimensions["E"].width = 2
            sheet.column_dimensions["F"].width = 6
            sheet.column_dimensions["G"].width = 2
            sheet.column_dimensions["H"].width = 6
            sheet.column_dimensions["I"].width = 2
            sheet.column_dimensions["J"].width = 6
            sheet.column_dimensions["K"].width = 2
            sheet.column_dimensions["L"].width = 6
            sheet.column_dimensions["M"].width = 2
            sheet.column_dimensions["N"].width = 6
            sheet.column_dimensions["O"].width = 2

    def carrierloop(self):
        """ loop for each carrier """
        self.row = 6  # allow space for headers
        first_page = True
        carriers_displayed = 0
        for carrier in self.carrier_list:
            self.pbi += 1
            self.pb.move_count(self.pbi)  # increment progress bar
            self.pb.change_text("Checking: {}".format(self.ot_carrier))
            if carrier[2] == "otdl":
                self.ot_carrier = carrier[1]  # current iteration of carrier list is assigned self.carrier
                self.display_recs()
                carriers_displayed += 1
            if first_page and carriers_displayed == 5:  # allow only five carriers per page.
                self.make_pagebreak()  # insert a page break
                carriers_displayed = 0  # reinitialize the counter
                first_page = False
            if not first_page and carriers_displayed == 6:
                self.make_pagebreak()  # insert a page break
                carriers_displayed = 0  # reinitialize the counter

    def display_recs(self):
        """ build the carrier and ring recs into the spreadsheet. """
        merge_first = ("B", "D", "F", "H", "J", "L", "N")
        merge_second = ("C", "E", "G", "I", "K", "M", "O")
        col_increment = 2
        cell = self.availability.cell(row=self.row, column=1)  # carrier label
        cell.value = "carrier:  "
        cell.style = self.name_header
        cell = self.availability.cell(row=self.row, column=2)  # carrier name input
        cell.value = self.ot_carrier
        cell.style = self.input_name
        self.availability.merge_cells('B' + str(self.row) + ':O' + str(self.row))
        self.row += 1
        # row headers
        cell = self.availability.cell(column=1, row=self.row + 1)  # paid leave label
        cell.value = "paid leave: "
        cell.style = self.name_header
        self.availability.row_dimensions[self.row + 1].height = 12  # adjust all row height
        cell = self.availability.cell(column=1, row=self.row + 2)  # hours worked label
        cell.value = "hours worked: "
        cell.style = self.name_header
        self.availability.row_dimensions[self.row + 2].height = 12  # adjust all row height
        cell = self.availability.cell(column=1, row=self.row + 3)  # cumulative hours label
        cell.value = "cumulative hours: "
        cell.style = self.name_header
        self.availability.row_dimensions[self.row + 3].height = 12  # adjust all row height
        cell = self.availability.cell(column=1, row=self.row + 4)  # cumulative hours label
        cell.value = "available weekly: "
        cell.style = self.name_header
        self.availability.row_dimensions[self.row + 4].height = 12  # adjust all row height
        cell = self.availability.cell(column=1, row=self.row + 5)  # cumulative hours label
        cell.value = "available daily: "
        cell.style = self.name_header
        self.availability.row_dimensions[self.row + 5].height = 12  # adjust all row height
        # use loops and an array to build the column headers
        column_headers = ("sat", "sun", "mon", "tue", "wed", "thu", "fri")
        for i in range(7):
            # ------------------------------------------------------------------------------------- column headers row
            cell = self.availability.cell(column=i + col_increment, row=self.row)
            cell.value = column_headers[i]
            cell.style = self.col_header
            self.availability.merge_cells(str(merge_first[i]) + str(self.row) + ":" +
                                          str(merge_second[i]) + str(self.row))
            # ----------------------------------------------------------------------------------------- paid leave row
            rings = self.get_rings(self.dates[i])  # get the total hours, leave type and leave hours for the carrier
            cell = self.availability.cell(column=i + col_increment, row=self.row + 1)  # column headers row
            cell.value = self.format_time(rings[2])  # format and display leave time
            cell.style = self.input_s
            cell.number_format = "#,###.00;[RED]-#,###.00"
            cell = self.availability.cell(column=i + 1 + col_increment, row=self.row + 1)  # column headers row
            cell.value = self.leave_code(rings[1])  # format and display leave code
            cell.style = self.col_header
            # ---------------------------------------------------------------------------------------- total hours row
            cell = self.availability.cell(column=i + col_increment, row=self.row + 2)  # column headers row
            cell.value = self.format_time(rings[0])
            cell.style = self.input_s
            cell.number_format = "#,###.00;[RED]-#,###.00"
            self.availability.merge_cells(str(merge_first[i]) + str(self.row + 2) + ":" +
                                          str(merge_second[i]) + str(self.row + 2))
            # --------------------------------------------------------------------------------------- cumulative hours
            cell = self.availability.cell(column=i + col_increment, row=self.row + 3)
            cell.value = self.cum_formula(i, self.row)  # get the formula for the cell
            cell.style = self.calcs
            cell.number_format = "#,###.00;[RED]-#,###.00"
            self.availability.merge_cells(str(merge_first[i]) + str(self.row + 3) + ":" +
                                          str(merge_second[i]) + str(self.row + 3))
            # -------------------------------------------------------------------------------------------- availability
            cell = self.availability.cell(column=i + col_increment, row=self.row + 4)
            cell.value = self.avail_formula(i, self.row)  # get the formula for the cell
            cell.style = self.calcs
            cell.number_format = "#,###.00;[RED]-#,###.00"
            self.availability.merge_cells(str(merge_first[i]) + str(self.row + 4) + ":" +
                                          str(merge_second[i]) + str(self.row + 4))
            # --------------------------------------------------------------------------------------daily availability
            cell = self.availability.cell(column=i + col_increment, row=self.row + 5)
            cell.value = self.avail_daily(i, self.row)  # get the formula for the cell
            cell.style = self.calcs
            cell.number_format = "#,###.00;[RED]-#,###.00"
            self.availability.merge_cells(str(merge_first[i]) + str(self.row + 5) + ":" +
                                          str(merge_second[i]) + str(self.row + 5))
            col_increment += 1  # move over two columns
        self.row += 7

    def get_rings(self, date):
        """ get individual carrier rings for the day - define self.rings"""
        sql = "SELECT total, leave_type, leave_time FROM rings3 WHERE carrier_name = '%s' AND rings_date = '%s' " \
              "ORDER BY rings_date, carrier_name" % (self.ot_carrier, date)
        rings = inquire(sql)
        totalhours = 0.0  # set default as an empty string
        lv_type = ""
        lv_hours = ""
        if rings:  # if rings record is not blank
            totalhours = float(Convert(rings[0][0]).zero_not_empty())
            lv_type = rings[0][1]
            lv_hours = rings[0][2]
        return [totalhours, lv_type, lv_hours]

    @staticmethod
    def leave_code(leave):
        """ converts the leave type to a one letter code. """
        if leave == "annual":
            return "A"
        elif leave == "sick":
            return "S"
        elif leave == "holiday":
            return "H"
        elif leave == "other":
            return "O"
        elif leave == "combo":
            return "C"
        elif leave == "none":
            return ""
        else:
            return ""

    @staticmethod
    def format_time(time):
        """ format the time for leave time and total time """
        if time == "0.0" or time == "0":
            return ""
        elif isfloat(time):
            return float(time)
        else:
            return time

    @staticmethod
    def cum_formula(day, row):
        """ return a formula for cumulative hours """
        if day == 0:  # if the day is saturday
            return "=SUM(%s!B%s+B%s)" % ('availability', str(row + 1), str(row + 2))
        if day == 1:  # if the day is sunday
            return "=SUM(%s!B%s+D%s+D%s)" % ('availability', str(row + 3), str(row + 1), str(row + 2))
        if day == 2:  # if the day is monday
            return "=SUM(%s!D%s+F%s+F%s)" % ('availability', str(row + 3), str(row + 1), str(row + 2))
        if day == 3:  # if the day is tuesday
            return "=SUM(%s!F%s+H%s+H%s)" % ('availability', str(row + 3), str(row + 1), str(row + 2))
        if day == 4:  # if the day is wednesday
            return "=SUM(%s!H%s+J%s+J%s)" % ('availability', str(row + 3), str(row + 1), str(row + 2))
        if day == 5:  # if the day is thursday
            return "=SUM(%s!J%s+L%s+L%s)" % ('availability', str(row + 3), str(row + 1), str(row + 2))
        if day == 6:  # if the day is friday
            return "=SUM(%s!L%s+N%s+N%s)" % ('availability', str(row + 3), str(row + 1), str(row + 2))

    @staticmethod
    def avail_formula(day, row):
        """ return a formula for cumulative hours """
        if day == 0:  # if the day is saturday
            return "=MAX(%s-%s!B%s, 0)" % (str(60), 'availability', str(row + 3))
        if day == 1:  # if the day is sunday
            return "=MAX(%s-%s!D%s, 0)" % (str(60), 'availability', str(row + 3))
        if day == 2:  # if the day is monday
            return "=MAX(%s-%s!F%s, 0)" % (str(60), 'availability', str(row + 3))
        if day == 3:  # if the day is tuesday
            return "=MAX(%s-%s!H%s, 0)" % (str(60), 'availability', str(row + 3))
        if day == 4:  # if the day is wednesday
            return "=MAX(%s-%s!J%s, 0)" % (str(60), 'availability', str(row + 3))
        if day == 5:  # if the day is thursday
            return "=MAX(%s-%s!L%s, 0)" % (str(60), 'availability', str(row + 3))
        if day == 6:  # if the day is friday
            return "=MAX(%s-%s!N%s, 0)" % (str(60), 'availability', str(row + 3))

    @staticmethod
    def avail_daily(day, row):
        """ return a formula for cumulative hours """
        if day == 0:  # if the day is saturday
            return "=IF(%s!C%s=\"\",MIN(MAX(%s-%s!B%s, 0), %s!B%s),0)" % \
                   ('availability', str(row + 1), str(12), 'availability', str(row + 2), 'availability', str(row + 4))
        if day == 1:  # if the day is sunday
            return "=IF(%s!E%s=\"\",MIN(MAX(%s-%s!D%s, 0), %s!D%s),0)" % \
                   ('availability', str(row + 1), str(12), 'availability', str(row + 2), 'availability', str(row + 4))
        if day == 2:  # if the day is monday
            return "=IF(%s!G%s=\"\",MIN(MAX(%s-%s!F%s, 0), %s!F%s),0)" % \
                   ('availability', str(row + 1), str(12), 'availability', str(row + 2), 'availability', str(row + 4))
        if day == 3:  # if the day is tuesday
            return "=IF(%s!I%s=\"\",MIN(MAX(%s-%s!H%s, 0), %s!H%s),0)" % \
                   ('availability', str(row + 1), str(12), 'availability', str(row + 2), 'availability', str(row + 4))
        if day == 4:  # if the day is wednesday
            return "=IF(%s!K%s=\"\",MIN(MAX(%s-%s!J%s, 0), %s!J%s),0)" % \
                   ('availability', str(row + 1), str(12), 'availability', str(row + 2), 'availability', str(row + 4))
        if day == 5:  # if the day is thursday
            return "=IF(%s!M%s=\"\",MIN(MAX(%s-%s!L%s, 0), %s!L%s),0)" % \
                   ('availability', str(row + 1), str(12), 'availability', str(row + 2), 'availability', str(row + 4))
        if day == 6:  # if the day is friday
            return "=IF(%s!O%s=\"\",MIN(MAX(%s-%s!N%s, 0), %s!N%s),0)" % \
                   ('availability', str(row + 1), str(12), 'availability', str(row + 2), 'availability', str(row + 4))

    def make_pagebreak(self):
        """ create a page break """
        try:
            self.availability.page_breaks.append(Break(id=self.row))
            self.row += 1
        except AttributeError:
            self.availability.row_breaks.append(Break(id=self.row))  # effective for windows
            self.row += 1

    def save_open(self):
        """ save the spreadsheet and open """
        self.pbi += 1
        self.pb.move_count(self.pbi)  # increment progress bar
        self.pb.change_text("Saving...")
        self.pb.stop()
        xl_filename = "kb_wa" + str(format(projvar.invran_date_week[0], "_%y_%m_%d")) + ".xlsx"
        try:
            self.wb.save(dir_path('weekly_availability') + xl_filename)
            messagebox.showinfo("Spreadsheet generator",
                                "Your spreadsheet was successfully generated. \n"
                                "File is named: {}".format(xl_filename),
                                parent=self.frame)
            if sys.platform == "win32":  # open the text document
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
                                 parent=self.frame)
    