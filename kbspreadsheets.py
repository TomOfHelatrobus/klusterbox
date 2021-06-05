import projvar
from kbtoolbox import inquire, CarrierList, dir_path, isfloat, Convert, CarrierRecFilter, Rings, SpeedSettings, \
    ProgressBarDe
from tkinter import messagebox
import os
import sys
import subprocess
from datetime import timedelta
from operator import itemgetter
# non standard libraries
from openpyxl import Workbook
from openpyxl.worksheet.pagebreak import Break
from openpyxl.styles import NamedStyle, Font, Border, Side, Alignment, PatternFill, Protection


class ImpManSpreadsheet:
    def __init__(self):
        self.frame = None  # the frame of parent
        self.startdate = None  # start date of the investigation
        self.enddate = None  # ending date of the investigation
        self.dates = []  # all days of the investigation
        self.carrierlist = []  # all carriers in carrier list
        self.carrier_breakdown = []  # all carriers in carrier list broken down into appropiate list
        self.wb = None  # the workbook object
        self.ws_list = []  # "saturday", "sunday", "monday", "tuesday", "wednesday", "thursday", "friday"
        self.summary = None
        self.reference = None
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

    def create(self, frame):
        self.frame = frame
        self.ask_ok()
        self.get_dates()
        self.get_carrierlist()
        self.get_carrier_breakdown()
        self.get_tolerances()  # get tolerances, minimum rows and page break preferences from tolerances table
        self.get_styles()
        self.build_workbook()
        self.set_dimensions()
        self.build_refs()
        self.build_ws_loop()  # calls list loop and carrier loop
        self.build_summary_header()
        self.save_open()

    def ask_ok(self):
        if messagebox.askokcancel("Spreadsheet generator",
                                  "Do you want to generate a spreadsheet?",
                                  parent=self.frame):
            return True
        return False

    def get_dates(self):
        self.startdate = projvar.invran_date
        self.enddate = projvar.invran_date
        self.dates = [projvar.invran_date, ]
        if projvar.invran_weekly_span:
            date = projvar.invran_date_week[0]
            self.startdate = projvar.invran_date_week[0]
            self.enddate = projvar.invran_date_week[6]
            self.dates = []
            for i in range(7):
                self.dates.append(date)
                date += timedelta(days=1)

    def get_carrierlist(self):  # get record sets for all carriers
        self.carrierlist = CarrierList(self.startdate, self.enddate, projvar.invran_station).get()

    def get_carrier_breakdown(self):
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

    def get_tolerances(self):  # get spreadsheet tolerances, row minimums and page break prefs from tolerance table
        sql = "SELECT tolerance FROM tolerances"
        result = inquire(sql)
        self.tol_ot_ownroute = float(result[0][0])  # overtime on own route tolerance
        self.tol_ot_offroute = float(result[1][0])  # overtime off own route tolerance
        self.tol_availability = float(result[2][0])  # availability tolerance
        self.min_ss_nl = int(result[3][0])  # minimum rows for no list
        self.min_ss_wal = int(result[4][0])  # mimimum rows for work assignment
        self.min_ss_otdl = int(result[5][0])  # minimum rows for otdl
        self.min_ss_aux = int(result[6][0])  # minimum rows for auxiliary
        self.pb_nl_wal = bool(result[21][0])  # page break between no list and work assignment
        self.pb_wal_otdl = bool(result[22][0])  # page break between work assignment and otdl
        self.pb_otdl_aux = bool(result[23][0])  # page break between otdl and auxiliary

    def get_styles(self):  # Named styles for workbook
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
        day_finder = ["Sat", "Sun", "Mon", "Tue", "Wed", "Thu", "Fri"]
        day_of_week = ["saturday", "sunday", "monday", "tuesday", "wednesday", "thursday", "friday"]
        i = 0
        self.wb = Workbook()  # define the workbook
        if not projvar.invran_weekly_span:  # if investigation range is daily
            for ii in range(len(day_finder)):
                if projvar.invran_date.strftime("%a") == day_finder[ii]:  # find the correct day
                    i = ii
            self.ws_list[i] = self.wb.active  # create first worksheet
            self.ws_list[i].title = day_of_week[i]  # title first worksheet
        if projvar.invran_weekly_span:  # if investigation range is weekly
            self.ws_list.append(self.wb.active)  # create first worksheet
            self.ws_list[0].title = "saturday"  # title first worksheet
            for i in range(1, 7):  # create worksheet for remaining six days
                self.ws_list.append(self.wb.create_sheet(day_of_week[i]))  # create subsequent worksheets
                self.ws_list[i].title = day_of_week[i]  # title subsequent worksheets
        self.summary = self.wb.create_sheet("summary")
        self.reference = self.wb.create_sheet("reference")

    def set_dimensions(self):
        for i in range(len(self.dates)):
            self.ws_list[i].oddFooter.center.text = "&A"
            self.ws_list[i].column_dimensions["A"].width = 14
            self.ws_list[i].column_dimensions["B"].width = 5
            self.ws_list[i].column_dimensions["C"].width = 6
            self.ws_list[i].column_dimensions["D"].width = 6
            self.ws_list[i].column_dimensions["E"].width = 6
            self.ws_list[i].column_dimensions["F"].width = 6
            self.ws_list[i].column_dimensions["G"].width = 6
            self.ws_list[i].column_dimensions["H"].width = 6
            self.ws_list[i].column_dimensions["I"].width = 6
            self.ws_list[i].column_dimensions["J"].width = 6
            self.ws_list[i].column_dimensions["K"].width = 6
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
        self.reference.column_dimensions["E"].width = 6

    def build_refs(self):
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
        self.reference['B7'].style = self.list_header
        self.reference['B7'] = "Code Guide"
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
        
    def build_ws_loop(self):
        self.i = 0
        for day in self.dates:
            self.day = day
            self.build_ws_headers()
            self.list_loop()  # loops four times. once for each list.
            self.i += 1

    def build_ws_headers(self):  # worksheet headers
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

    def list_loop(self):  # loops four times. once for each list.
        self.lsi = 0  # iterations of the list loop method
        self.row = 6
        for _ in self.ot_list:  # loops for nl, wal, otdl and aux
            self.list_and_column_headers()  # builds headers for list and column
            self.carrierlist_mod()
            self.carrierloop()
            self.build_footer()
            self.pagebreak()
            self.lsi += 1

    def list_and_column_headers(self):  # builds headers for list and column
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

    def column_header_non(self):  # column headers specific for non otdl
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
        cell.value = "OT off rt"
        cell.style = self.col_header
        self.row += 1

    def column_header_ot(self):  # column headers specific for otdl or aux
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

    def carrierlist_mod(self):  # add empty carrier records to carrier list until quantity matches minrows preference
        self.mod_carrierlist = self.carrier_breakdown[self.i][self.lsi]
        minrows = 0  # initialize minrows
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

    def carrierloop(self):
        for carrier in self.mod_carrierlist:
            self.carrier = carrier[1]  # current iteration of carrier list is assigned self.carrier
            self.get_rings()  # get individual carrier rings for the day
            self.display_recs()
            if self.pref[self.lsi] in ("nl", "wal"):
                self.get_movesarray()
                self.display_moves()
                self.display_formulas_non()
            else:
                self.display_formulas_ot()
            self.increment_rows()

    def increment_rows(self):  # increment the rows counter
        self.row += 1
        self.row += self.move_i  # add 1 plus any the added rows from multiple moves
        self.move_i = 0  # reset the row incrementor for multiple move functionality

    def get_rings(self):  # get individual carrier rings for the day
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

    def display_recs(self):  # put the carrier and the first part of rings into the spreadsheet
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

    def get_movesarray(self):  # builds sets of moves for each triad
        multiple_sets = False  # is there more than one triad?
        self.movesarray = []  # re initialized - a list of tuples of move sets
        moves_array = []  # initialized - the moves string converted into an array
        day_of_week = ["saturday", "sunday", "monday", "tuesday", "wednesday", "thursday", "friday"]
        move_off = ""  # if empty set, use default values
        move_back = ""
        move_route = ""
        formula = "=SUM(%s!F%s - %s!E%s)" % (day_of_week[self.i], self.row, day_of_week[self.i], self.row)
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
                          (day_of_week[self.i], self.row + 1, int(self.row + len(moves_array) / 3))
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
                    formula = "=SUM(%s!F%s - %s!E%s)" % (day_of_week[self.i], self.row + formula_row_i,
                                                         day_of_week[self.i], self.row + formula_row_i)
                    add_this = (move_off, move_back, move_route, formula)
                    self.movesarray.append(add_this)
                    formula_row_i += 1  # increment the row in the formula after each moves_set
                i += 1  # increment i

    def display_moves(self):
        for move_set in self.movesarray:
            for move_cell in range(4):
                move = move_set[move_cell]
                cell = self.ws_list[self.i].cell(row=self.row + self.move_i, column=5 + move_cell)
                if move_cell in (0, 1):  # format move times as floats or empty strings
                    cell.value = Convert(move).empty_not_zerofloat()  # insert an iteration of self.movesarray
                else:  # do not alter route or formula elements of move sets
                    cell.value = move  # insert an iteration of self.movesarray
                cell.style = self.input_s  # assign worksheet style
                if move_cell != 2:
                    cell.number_format = "#,###.00;[RED]-#,###.00"
            self.move_i += 1
        self.move_i -= 1  # correction

    def display_formulas_non(self):
        day_of_week = ["saturday", "sunday", "monday", "tuesday", "wednesday", "thursday", "friday"]
        ot_formula = "=IF(%s!B%s =\"ns day\", %s!C%s, MAX(%s!C%s - 8, 0))" \
                  % (day_of_week[self.i], str(self.row), day_of_week[self.i], str(self.row),
                     day_of_week[self.i], str(self.row))
        if self.pref[self.lsi] == "nl":  # use alternate formula for non list carriers
            ot_formula = "=IF(%s!B%s =\"ns day\", %s!C%s,IF(%s!C%s <= 8 + reference!C3, 0, MAX(%s!C%s - 8, 0)))" \
                         % (day_of_week[self.i], str(self.row), day_of_week[self.i], str(self.row),
                            day_of_week[self.i], str(self.row), day_of_week[self.i], str(self.row))
        off_rt_formula = "=%s!H%s" % (day_of_week[self.i], str(self.row))  # copy data from column H/ MV total
        ot_off_rt_formula = "=IF(OR(%s!B%s=\"ns day\",%s!J%s >= %s!C%s), " \
                  "%s!C%s, IF(%s!C%s <= 8 + reference!C4, 0, " \
                  "MIN(MAX(%s!C%s - 8, 0),IF(%s!J%s <= reference!C4,0, %s!J%s))))" \
                  % (day_of_week[self.i], str(self.row), day_of_week[self.i], str(self.row),
                     day_of_week[self.i], str(self.row), day_of_week[self.i], str(self.row),
                     day_of_week[self.i], str(self.row), day_of_week[self.i], str(self.row),
                     day_of_week[self.i], str(self.row), day_of_week[self.i], str(self.row))
        formulas = (ot_formula, off_rt_formula, ot_off_rt_formula)
        column_i = 0
        for formula in formulas:
            cell = self.ws_list[self.i].cell(row=self.row, column=9 + column_i)
            cell.value = formula  # insert an iteration of formulas
            cell.style = self.input_s  # assign worksheet style
            cell.number_format = "#,###.00;[RED]-#,###.00"
            column_i += 1

    def display_formulas_ot(self):
        day_of_week = ["saturday", "sunday", "monday", "tuesday", "wednesday", "thursday", "friday"]
        max_hrs = 12  # maximum hours for otdl carriers
        if self.pref[self.lsi] == "aux":  # alter formula by list preference
            max_hrs = 11.5  # maximux hours for auxiliary carriers
        formula_ten = "=IF(OR(%s!B%s = \"light\", %s!B%s = \"excused\", %s!B%s = \"sch chg\", " \
                     "%s!B%s = \"annual\", %s!B%s = \"sick\", %s!C%s >= 10 - reference!C5), 0, " \
                     "IF(%s!B%s = \"no call\", 10, " \
                     "IF(%s!C%s = 0, 0, MAX(10 - %s!C%s, 0))))" % \
                     (day_of_week[self.i], str(self.row), day_of_week[self.i], str(self.row),
                      day_of_week[self.i], str(self.row), day_of_week[self.i], str(self.row),
                      day_of_week[self.i], str(self.row), day_of_week[self.i], str(self.row),
                      day_of_week[self.i], str(self.row), day_of_week[self.i], str(self.row),
                      day_of_week[self.i], str(self.row))
        formula_max = "=IF(OR(%s!B%s = \"light\",%s!B%s = \"excused\", %s!B%s = \"sch chg\", %s!B%s = \"annual\", " \
                      "%s!B%s = \"sick\", %s!C%s >= %s - reference!C5), 0, IF(%s!B%s = \"no call\", %s, " \
                      "IF(%s!C%s = 0, 0, MAX(%s - %s!C%s, 0))))" % \
                      (day_of_week[self.i], str(self.row), day_of_week[self.i], str(self.row),
                      day_of_week[self.i], str(self.row), day_of_week[self.i], str(self.row),
                      day_of_week[self.i], str(self.row), day_of_week[self.i], str(self.row),
                      max_hrs, day_of_week[self.i], str(self.row),
                      max_hrs, day_of_week[self.i], str(self.row),
                      max_hrs, day_of_week[self.i], str(self.row))
        formulas = (formula_ten, formula_max)
        column_i = 0
        for formula in formulas:
            cell = self.ws_list[self.i].cell(row=self.row, column=5 + column_i)
            cell.value = formula  # insert an iteration formulas
            cell.style = self.input_s  # assign worksheet style
            cell.number_format = "#,###.00;[RED]-#,###.00"
            column_i += 1

    def build_footer(self):
        if self.pref[self.lsi] == "nl":
            self.nl_footer()
            
    def nl_footer(self):
        cell = self.ws_list[self.i].cell(row=self.row, column=8)
        cell.value = "Total NL Overtime"
        cell.style = self.col_header
        formula = ""
        # formula = "=SUM(%s!I8:I%s)" % (day_of_week[i], cello)
        cell = self.ws_list[self.i].cell(row=self.row, column=9)
        cell.value = formula  # OT
        # nl_ot_row.append(str(self.row))  # get the cello information to reference in summary tab
        # nl_ot_day.append(i)
        cell.style = self.calcs
        cell.number_format = "#,###.00;[RED]-#,###.00"
        self.row += 2
        cell = self.ws_list[self.i].cell(row=self.row, column=10)
        cell.value = "Total NL Mandates"
        cell.style = self.col_header
        formula = ""
        # formula = "=SUM(%s!K8:K%s)" % (day_of_week[i], cello)
        cell = self.ws_list[self.i].cell(row=self.row, column=11)
        cell.value = formula  # OT off route
        cell.style = self.calcs
        cell.number_format = "#,###.00;[RED]-#,###.00"
        nl_totals = self.row
        self.row += 1

    def pagebreak(self):  # create a page break if consistant with user preferences
        if self.pref[self.lsi] == "nl" and not self.pb_nl_wal:
            return
        if self.pref[self.lsi] == "wal" and not self.pb_wal_otdl:
            return
        if self.pref[self.lsi] == "otdl" and not self.pb_otdl_aux:
            return
        if self.pref[self.lsi] == "aux":
            return
        try:
            self.ws_list[self.i].page_breaks.append(Break(id=self.row))
            # print("page break")
        except AttributeError:
            self.ws_list[self.i].row_breaks.append(Break(id=self.row))  # effective for windows
            # print("row break")

    def build_summary_header(self):  # summary headers
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

    def save_open(self):  # name the excel file
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


def spreadsheet(frame, list_carrier, r_rings):
    date = projvar.invran_date_week[0]
    dates = []  # array containing days.
    if projvar.invran_weekly_span:  # if investigation range is weekly
        for i in range(7):
            dates.append(date)
            date += timedelta(days=1)
    if not projvar.invran_weekly_span:  # if investigation range is daily
        dates.append(projvar.invran_date)
    if r_rings == "x":
        if projvar.invran_weekly_span:  # if investigation range is weekly
            sql = "SELECT * FROM rings3 WHERE rings_date BETWEEN '%s' AND '%s' ORDER BY rings_date, carrier_name" \
                  % (projvar.invran_date_week[0], projvar.invran_date_week[6])
        else:
            sql = "SELECT * FROM rings3 WHERE rings_date = '%s' ORDER BY rings_date, " \
                  "carrier_name" \
                  % projvar.invran_date
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
    i = 0
    wb = Workbook()  # define the workbook
    if not projvar.invran_weekly_span:  # if investigation range is daily
        for ii in range(len(day_finder)):
            if projvar.invran_date.strftime("%a") == day_finder[ii]:  # find the correct day
                i = ii
        ws_list[i] = wb.active  # create first worksheet
        ws_list[i].title = day_of_week[i]  # title first worksheet
        summary = wb.create_sheet("summary")
        reference = wb.create_sheet("reference")
    if projvar.invran_weekly_span:  # if investigation range is weekly
        ws_list[0] = wb.active  # create first worksheet
        ws_list[0].title = "saturday"  # title first worksheet
        for i in range(1, len(ws_list)):  # create worksheet for remaining six days
            ws_list[i] = wb.create_sheet(ws_list[i])
            # i = 0
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
            if list_carrier[ii][0][0] <= str(day):
                candidates.append(list_carrier[ii][0])  # put name into candidates array
            jump = "no"  # triggers an analysis of the candidates array
            if ii != len(list_carrier) - 1:  # if the loop has not reached the end of the list
                if list_carrier[ii][0][1] == list_carrier[ii + 1][0][1]:  # if the name current and next name are same
                    jump = "yes"  # bypasses an analysis of the candidates array
            if jump == "no":  # review the list of candidates
                winner = max(candidates, key=itemgetter(0))  # select the most recent
                if winner[5] == projvar.invran_station:
                    daily_list.append(
                    winner)  # add the record if it matches the station
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
        cell = ws_list[i].cell(row=1, column=1)
        cell.value = "Improper Mandate Worksheet"
        cell.style = ws_header
        ws_list[i].merge_cells('A1:E1')
        cell = ws_list[i].cell(row=3, column=1)
        cell.value = "Date:  "  # create date/ pay period/ station header
        cell.style = date_dov_title
        cell = ws_list[i].cell(row=3, column=2)
        cell.value = format(day, "%A  %m/%d/%y")
        cell.style = date_dov
        ws_list[i].merge_cells('B3:D3')
        cell = ws_list[i].cell(row=3, column=5)
        cell.value = "Pay Period:  "
        cell.style = date_dov_title
        ws_list[i].merge_cells('E3:F3')
        cell = ws_list[i].cell(row=3, column=7)
        cell.value = projvar.pay_period
        cell.style = date_dov
        ws_list[i].merge_cells('G3:H3')
        cell = ws_list[i].cell(row=4, column=1)
        cell.value = "Station:  "
        cell.style = date_dov_title
        cell = ws_list[i].cell(row=4, column=2)
        cell.value = projvar.invran_station
        cell.style = date_dov
        ws_list[i].merge_cells('B4:D4')
        # no list carriers *********************************************************************************************
        cell = ws_list[i].cell(row=6, column=1)
        cell.value = "No List Carriers"
        cell.style = list_header
        # column headers
        cell = ws_list[i].cell(row=7, column=1)
        cell.value = "Name"
        cell.style = col_header
        cell = ws_list[i].cell(row=7, column=2)
        cell.value = "note"
        cell.style = col_header
        cell = ws_list[i].cell(row=7, column=3)
        cell.value = "5200"
        cell.style = col_header
        cell = ws_list[i].cell(row=7, column=4)
        cell.value = "RS"
        cell.style = col_header
        cell = ws_list[i].cell(row=7, column=5)
        cell.value = "MV off"
        cell.style = col_header
        cell = ws_list[i].cell(row=7, column=6)
        cell.value = "MV on"
        cell.style = col_header
        cell = ws_list[i].cell(row=7, column=7)
        cell.value = "Route"
        cell.style = col_header
        cell = ws_list[i].cell(row=7, column=8)
        cell.value = "MV total"
        cell.style = col_header
        cell = ws_list[i].cell(row=7, column=9)
        cell.value = "OT"
        cell.style = col_header
        cell = ws_list[i].cell(row=7, column=10)
        cell.value = "off rt"
        cell.style = col_header
        cell = ws_list[i].cell(row=7, column=11)
        cell.value = "OT off rt"
        cell.style = col_header
        oi = 8  # rows: start at 8th row
        move_totals = []  # list of totals of each set of moves
        ot_total = 0  # running total for OT
        ot_off_total = 0  # running total for OT off route
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
                        cc = 0
                        for e in range(int(len(s_moves) / 3)):  # tally totals for each set of moves
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
                            for mt in move_totals:  # calc total off route work.
                                off_route += float(mt)
                        ot_total += ot
                        ot_off_route = min(off_route, ot)  # calculate the ot off route
                        ot_off_total += ot_off_route
                        move_count = (int(len(s_moves) / 3))  # find the number of sets of moves
                        # output to the gui
                        cell = ws_list[i].cell(row=oi, column=1)
                        cell.value = each[1]  # name
                        cell.style = input_name
                        if each[4] == "none":
                            code = ""  # leave code field blank if 'none'
                        else:
                            code = each[4]
                        cell = ws_list[i].cell(row=oi, column=2)
                        cell.value = code  # code
                        cell.style = input_s
                        cell = ws_list[i].cell(row=oi, column=3)
                        if time5200 == 0:
                            cell.value = ""  # 5200
                        else:
                            cell.value = float(time5200)  # 5200
                        cell.style = input_s
                        cell.number_format = "#,###.00;[RED]-#,###.00"
                        cell = ws_list[i].cell(row=oi, column=4)
                        if isfloat(each[3]):
                            cell.value = float(each[3])
                        else:
                            cell.value = each[3]
                        cell.style = input_s
                        cell.number_format = "#,###.00;[RED]-#,###.00"
                        count = 0
                        if move_count == 0:  # if there are no moves then format the empty cells
                            cell = ws_list[i].cell(row=oi, column=5)
                            cell.value = ""  # move off
                            cell.style = input_s
                            cell.number_format = "#,###.00;[RED]-#,###.00"
                            count += 1
                            cell = ws_list[i].cell(row=oi, column=6)
                            cell.value = ""  # move on
                            cell.style = input_s
                            cell.number_format = "#,###.00;[RED]-#,###.00"
                            count += 1
                            cell = ws_list[i].cell(row=oi, column=7)
                            cell.value = ""  # route
                            cell.style = input_s
                            count += 1
                            formula = "=SUM(%s!F%s - %s!E%s)" % (day_of_week[i], str(oi), day_of_week[i], str(oi))
                            cell = ws_list[i].cell(row=oi, column=8)
                            cell.value = formula  # move total
                            cell.style = input_s
                            cell.number_format = "#,###.00;[RED]-#,###.00"
                        elif move_count == 1:  # if there is only one set of moves
                            cell = ws_list[i].cell(row=oi, column=5)
                            cell.value = float(s_moves[0])  # move off
                            cell.style = input_s
                            cell.number_format = "#,###.00;[RED]-#,###.00"
                            count += 1
                            cell = ws_list[i].cell(row=oi, column=6)
                            cell.value = float(s_moves[1])  # move on
                            cell.style = input_s
                            cell.number_format = "#,###.00;[RED]-#,###.00"
                            count += 1
                            cell = ws_list[i].cell(row=oi, column=7)
                            cell.value = str(s_moves[2])  # route
                            cell.style = input_s
                            count += 1
                            formula = "=SUM(%s!F%s - %s!E%s)" % (day_of_week[i], str(oi), day_of_week[i], str(oi))
                            cell = ws_list[i].cell(row=oi, column=8)
                            cell.value = formula  # move total
                            cell.style = input_s
                            cell.number_format = "#,###.00;[RED]-#,###.00"
                        else:  # There are multiple moves
                            cell = ws_list[i].cell(row=oi, column=5)
                            cell.value = "*"  # move off
                            cell.style = input_s
                            cell = ws_list[i].cell(row=oi, column=6)
                            cell.value = "*"  # move on
                            cell.style = input_s
                            cell = ws_list[i].cell(row=oi, column=7)
                            cell.value = "*"  # route
                            cell.style = input_s
                            formula = "=SUM(%s!H%s:H%s)" % (day_of_week[i], str(oi + move_count), str(oi + 1))
                            cell = ws_list[i].cell(row=oi, column=8)
                            cell.value = formula  # move total
                            cell.style = input_s
                            cell.number_format = "#,###.00;[RED]-#,###.00"
                            formula = "=IF(%s!B%s =\"ns day\", %s!C%s,IF(%s!C%s <= 8 + reference!C3, 0, " \
                                      "MAX(%s!C%s - 8, 0)))" \
                                      % (day_of_week[i], str(oi), day_of_week[i], str(oi), day_of_week[i], str(oi),
                                         day_of_week[i], str(oi))
                            cell = ws_list[i].cell(row=oi, column=9)
                            cell.value = formula  # overtime
                            cell.style = calcs
                            cell.number_format = "#,###.00;[RED]-#,###.00"
                            formula = "=%s!H%s" % (day_of_week[i], str(oi))  # copy data from column H/ MV total
                            cell = ws_list[i].cell(row=oi, column=10)
                            cell.value = formula  # off route
                            cell.style = calcs
                            cell.number_format = "#,###.00;[RED]-#,###.00"
                            # formula for OT off route
                            formula = "=IF(OR(%s!B%s=\"ns day\",%s!J%s >= %s!C%s),%s!C%s, " \
                                      "IF(%s!C%s <= 8 + reference!C4, 0, MIN" \
                                      "(MAX(%s!C%s - 8, 0),IF(%s!J%s <= reference!C4,0, %s!J%s))))" \
                                      % (day_of_week[i], str(oi), day_of_week[i], str(oi), day_of_week[i], str(oi),
                                         day_of_week[i], str(oi), day_of_week[i], str(oi),
                                         day_of_week[i], str(oi), day_of_week[i], str(oi), day_of_week[i], str(oi))
                            cell = ws_list[i].cell(row=oi, column=11)
                            cell.value = formula  # OT off route
                            cell.style = calcs
                            cell.number_format = "#,###.00;[RED]-#,###.00"
                            oi += 1
                            for ii in range(move_count):  # if there are multiple moves, create + populate cells
                                cell = ws_list[i].cell(row=oi, column=5)
                                cell.value = float(s_moves[count])  # move off
                                cell.style = input_s
                                cell.number_format = "#,###.00;[RED]-#,###.00"
                                count += 1
                                cell = ws_list[i].cell(row=oi, column=6)
                                cell.value = float(s_moves[count])  # move on
                                cell.style = input_s
                                cell.number_format = "#,###.00;[RED]-#,###.00"
                                count += 1
                                cell = ws_list[i].cell(row=oi, column=7)
                                cell.value = str(s_moves[count])  # route
                                cell.style = input_s
                                count += 1
                                formula = "=SUM(%s!F%s - %s!E%s)" % (day_of_week[i], str(oi), day_of_week[i], str(oi))
                                cell = ws_list[i].cell(row=oi, column=8)
                                cell.value = formula  # move total
                                cell.style = input_s
                                cell.number_format = "#,###.00;[RED]-#,###.00"
                                if ii < move_count - 1:
                                    oi += 1  # create another row
                            oi += 1
                        if move_count < 2:
                            # input formula for overtime
                            formula = "=IF(%s!B%s =\"ns day\", %s!C%s,IF(%s!C%s <= 8+ reference!C3, 0, " \
                                      "MAX(%s!C%s - 8, 0)))" \
                                      % (day_of_week[i], str(oi), day_of_week[i], str(oi), day_of_week[i], str(oi),
                                         day_of_week[i], str(oi))
                            cell = ws_list[i].cell(row=oi, column=9)
                            cell.value = formula  # overtime
                            cell.style = calcs
                            cell.number_format = "#,###.00;[RED]-#,###.00"
                            # formula for off route
                            formula = "=SUM(%s!F%s - %s!E%s)" \
                                      % (day_of_week[i], str(oi), day_of_week[i], str(oi))
                            cell = ws_list[i].cell(row=oi, column=10)
                            cell.value = formula  # off route
                            cell.style = calcs
                            cell.number_format = "#,###.00;[RED]-#,###.00"
                            # formula for OT off route
                            formula = "=IF(OR(%s!B%s=\"ns day\",%s!J%s >= %s!C%s),%s!C%s, IF(%s!C%s <= 8 + " \
                                      "reference!C4, 0, MIN" \
                                      "(MAX(%s!C%s - 8, 0),IF(%s!J%s <= reference!C4,0, %s!J%s))))" \
                                      % (day_of_week[i], str(oi), day_of_week[i], str(oi), day_of_week[i], str(oi),
                                         day_of_week[i], str(oi), day_of_week[i], str(oi),
                                         day_of_week[i], str(oi), day_of_week[i], str(oi), day_of_week[i], str(oi))
                            cell = ws_list[i].cell(row=oi, column=11)
                            cell.value = formula  # OT off route
                            cell.style = calcs
                            cell.number_format = "#,###.00;[RED]-#,###.00"
                            oi += 1
            #  if there is no match, then just printe the name.
            if match == "miss":
                cell = ws_list[i].cell(row=oi, column=1)
                cell.value = line[1]  # name
                cell.style = input_name
                cell = ws_list[i].cell(row=oi, column=2)
                cell.style = input_s
                cell.number_format = "#,###.00;[RED]-#,###.00"
                cell = ws_list[i].cell(row=oi, column=3)
                cell.style = input_s
                cell.number_format = "#,###.00;[RED]-#,###.00"
                cell = ws_list[i].cell(row=oi, column=4)
                cell.style = input_s
                cell.number_format = "#,###.00;[RED]-#,###.00"
                cell = ws_list[i].cell(row=oi, column=5)
                cell.style = input_s
                cell.number_format = "#,###.00;[RED]-#,###.00"
                cell = ws_list[i].cell(row=oi, column=6)
                cell.style = input_s
                cell.number_format = "#,###.00;[RED]-#,###.00"
                cell = ws_list[i].cell(row=oi, column=7)
                cell.style = input_s
                formula = "=SUM(%s!F%s - %s!E%s)" % (day_of_week[i], str(oi), day_of_week[i], str(oi))
                cell = ws_list[i].cell(row=oi, column=8)
                cell.value = formula  # move total
                cell.style = input_s
                cell.number_format = "#,###.00;[RED]-#,###.00"
                formula = "=IF(%s!B%s =\"ns day\", %s!C%s,IF(%s!C%s <= 8 + reference!C3, 0, MAX(%s!C%s - 8, 0)))" \
                          % (day_of_week[i], str(oi), day_of_week[i], str(oi), day_of_week[i], str(oi),
                             day_of_week[i], str(oi))
                cell = ws_list[i].cell(row=oi, column=9)
                cell.value = formula  # overtime
                cell.style = calcs
                cell.number_format = "#,###.00;[RED]-#,###.00"
                formula = "=SUM(%s!F%s - %s!E%s)" \
                          % (day_of_week[i], str(oi), day_of_week[i], str(oi))
                cell = ws_list[i].cell(row=oi, column=10)
                cell.value = formula  # off route
                cell.style = calcs
                cell.number_format = "#,###.00;[RED]-#,###.00"
                formula = "=IF(OR(%s!B%s=\"ns day\",%s!J%s >= %s!C%s),%s!C%s, IF(%s!C%s <= 8 + reference!C4, 0, MIN" \
                          "(MAX(%s!C%s - 8, 0),IF(%s!J%s <= reference!C4,0, %s!J%s))))" \
                          % (day_of_week[i], str(oi), day_of_week[i], str(oi), day_of_week[i], str(oi),
                             day_of_week[i], str(oi), day_of_week[i], str(oi),
                             day_of_week[i], str(oi), day_of_week[i], str(oi), day_of_week[i], str(oi))
                cell = ws_list[i].cell(row=oi, column=11)
                cell.value = formula  # OT off route
                cell.style = calcs
                cell.number_format = "#,###.00;[RED]-#,###.00"
                oi += 1
        nl_oi_end = oi
        nl_oi_diff = nl_oi_end - nl_oi_start  # find how many lines exist in nl
        # if the minimum number of rows are not reached, insert blank rows
        e_range = min_ss_nl - nl_oi_diff
        if e_range <= 0:
            e_range = 0
        for e in range(e_range):
            cell = ws_list[i].cell(row=oi, column=1)
            cell.value = ""  # name
            cell.style = input_name
            cell = ws_list[i].cell(row=oi, column=2)
            cell.style = input_s
            cell.number_format = "#,###.00;[RED]-#,###.00"
            cell = ws_list[i].cell(row=oi, column=3)
            cell.style = input_s
            cell.number_format = "#,###.00;[RED]-#,###.00"
            cell = ws_list[i].cell(row=oi, column=4)
            cell.style = input_s
            cell.number_format = "#,###.00;[RED]-#,###.00"
            cell = ws_list[i].cell(row=oi, column=5)
            cell.style = input_s
            cell.number_format = "#,###.00;[RED]-#,###.00"
            cell = ws_list[i].cell(row=oi, column=6)
            cell.style = input_s
            cell.number_format = "#,###.00;[RED]-#,###.00"
            cell = ws_list[i].cell(row=oi, column=7)
            cell.style = input_s
            formula = "=SUM(%s!F%s - %s!E%s)" % (day_of_week[i], str(oi), day_of_week[i], str(oi))
            cell = ws_list[i].cell(row=oi, column=8)
            cell.value = formula  # move total
            cell.style = input_s
            cell.number_format = "#,###.00;[RED]-#,###.00"
            formula = "=IF(%s!B%s =\"ns day\", %s!C%s,IF(%s!C%s <= 8 + reference!C3, 0, MAX(%s!C%s - 8, 0)))" \
                      % (day_of_week[i], str(oi), day_of_week[i], str(oi), day_of_week[i], str(oi),
                         day_of_week[i], str(oi))
            cell = ws_list[i].cell(row=oi, column=9)
            cell.value = formula  # overtime
            cell.style = calcs
            cell.number_format = "#,###.00;[RED]-#,###.00"
            formula = "=SUM(%s!F%s - %s!E%s)" \
                      % (day_of_week[i], str(oi), day_of_week[i], str(oi))
            cell = ws_list[i].cell(row=oi, column=10)
            cell.value = formula  # off route
            cell.style = calcs
            cell.number_format = "#,###.00;[RED]-#,###.00"
            formula = "=IF(OR(%s!B%s=\"ns day\",%s!J%s >= %s!C%s),%s!C%s, IF(%s!C%s <= 8 + reference!C4, 0, MIN" \
                      "(MAX(%s!C%s - 8, 0),IF(%s!J%s <= reference!C4,0, %s!J%s))))" \
                      % (day_of_week[i], str(oi), day_of_week[i], str(oi), day_of_week[i], str(oi),
                         day_of_week[i], str(oi), day_of_week[i], str(oi),
                         day_of_week[i], str(oi), day_of_week[i], str(oi), day_of_week[i], str(oi))
            cell = ws_list[i].cell(row=oi, column=11)
            cell.value = formula  # OT off route
            cell.style = calcs
            cell.number_format = "#,###.00;[RED]-#,###.00"
            oi += 1
        cello = str(oi - 1)
        oi += 1
        cell = ws_list[i].cell(row=oi, column=8)
        cell.value = "Total NL Overtime"
        cell.style = col_header
        formula = "=SUM(%s!I8:I%s)" % (day_of_week[i], cello)
        cell = ws_list[i].cell(row=oi, column=9)
        cell.value = formula  # OT
        nl_ot_row.append(str(oi))  # get the cello information to reference in summary tab
        nl_ot_day.append(i)
        cell.style = calcs
        cell.number_format = "#,###.00;[RED]-#,###.00"
        oi += 2
        cell = ws_list[i].cell(row=oi, column=10)
        cell.value = "Total NL Mandates"
        cell.style = col_header
        formula = "=SUM(%s!K8:K%s)" % (day_of_week[i], cello)
        cell = ws_list[i].cell(row=oi, column=11)
        cell.value = formula  # OT off route
        cell.style = calcs
        cell.number_format = "#,###.00;[RED]-#,###.00"
        nl_totals = oi
        oi += 1
        try:
            ws_list[i].page_breaks.append(Break(id=oi))
        except:
            ws_list[i].row_breaks.append(Break(id=oi))
        oi += 1
        # # work assignment carriers **********************************************************************
        cell = ws_list[i].cell(row=oi, column=1)
        cell.value = "Work Assignment Carriers"
        cell.style = list_header
        oi += 1
        # column headers
        cell = ws_list[i].cell(row=oi, column=1)
        cell.value = "Name"
        cell.style = col_header
        cell = ws_list[i].cell(row=oi, column=2)
        cell.value = "note"
        cell.style = col_header
        cell = ws_list[i].cell(row=oi, column=3)
        cell.value = "5200"
        cell.style = col_header
        cell = ws_list[i].cell(row=oi, column=4)
        cell.value = "RS"
        cell.style = col_header
        cell = ws_list[i].cell(row=oi, column=5)
        cell.value = "MV off"
        cell.style = col_header
        cell = ws_list[i].cell(row=oi, column=6)
        cell.value = "MV on"
        cell.style = col_header
        cell = ws_list[i].cell(row=oi, column=7)
        cell.value = "Route"
        cell.style = col_header
        cell = ws_list[i].cell(row=oi, column=8)
        cell.value = "MV total"
        cell.style = col_header
        cell = ws_list[i].cell(row=oi, column=9)
        cell.value = "OT"
        cell.style = col_header
        cell = ws_list[i].cell(row=oi, column=10)
        cell.value = "off rt"
        cell.style = col_header
        cell = ws_list[i].cell(row=oi, column=11)
        cell.value = "OT off rt"
        cell.style = col_header
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
                        cc = 0
                        for e in range(int(len(s_moves) / 3)):  # tally totals for each set of moves
                            total = float(s_moves[cc + 1]) - float(s_moves[cc])
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
                            for mt in move_totals:  # calc total off route work.
                                off_route += float(mt)
                        ot_total += ot
                        ot_off_route = min(off_route, ot)  # calculate the ot off route
                        ot_off_total += ot_off_route
                        move_count = (int(len(s_moves) / 3))  # find the number of sets of moves
                        # output to the gui
                        cell = ws_list[i].cell(row=oi, column=1)
                        cell.value = each[1]  # name
                        cell.style = input_name

                        if each[4] == "none":
                            code = ""  # leave code field blank if 'none'
                        else:
                            code = each[4]
                        cell = ws_list[i].cell(row=oi, column=2)
                        cell.value = code  # code
                        cell.style = input_s

                        cell = ws_list[i].cell(row=oi, column=3)
                        if time5200 == 0:
                            cell.value = ""  # 5200
                        else:
                            cell.value = float(time5200)  # 5200
                        cell.style = input_s
                        cell.number_format = "#,###.00;[RED]-#,###.00"

                        cell = ws_list[i].cell(row=oi, column=4)
                        if isfloat(each[3]):
                            cell.value = float(each[3])
                        else:
                            cell.value = each[3]
                        cell.style = input_s
                        cell.number_format = "#,###.00;[RED]-#,###.00"
                        count = 0
                        if move_count == 0:  # if there are no moves then format the empty cells
                            cell = ws_list[i].cell(row=oi, column=5)
                            cell.value = ""  # move off
                            cell.style = input_s
                            cell.number_format = "#,###.00;[RED]-#,###.00"
                            count += 1
                            cell = ws_list[i].cell(row=oi, column=6)
                            cell.value = ""  # move on
                            cell.style = input_s
                            cell.number_format = "#,###.00;[RED]-#,###.00"
                            count += 1
                            cell = ws_list[i].cell(row=oi, column=7)
                            cell.value = ""  # route
                            cell.style = input_s
                            count += 1
                            formula = "=SUM(%s!F%s - %s!E%s)" % (day_of_week[i], str(oi), day_of_week[i], str(oi))
                            cell = ws_list[i].cell(row=oi, column=8)
                            cell.value = formula  # move total
                            cell.style = input_s
                            cell.number_format = "#,###.00;[RED]-#,###.00"
                        elif move_count == 1:  # if there is only one set of moves
                            cell = ws_list[i].cell(row=oi, column=5)
                            cell.value = float(s_moves[0])  # move off
                            cell.style = input_s
                            cell.number_format = "#,###.00;[RED]-#,###.00"
                            count += 1
                            cell = ws_list[i].cell(row=oi, column=6)
                            cell.value = float(s_moves[1])  # move on
                            cell.style = input_s
                            cell.number_format = "#,###.00;[RED]-#,###.00"
                            count += 1
                            cell = ws_list[i].cell(row=oi, column=7)
                            cell.value = str(s_moves[2])  # route
                            cell.style = input_s
                            count += 1
                            formula = "=SUM(%s!F%s - %s!E%s)" % (day_of_week[i], str(oi), day_of_week[i], str(oi))
                            cell = ws_list[i].cell(row=oi, column=8)
                            cell.value = formula  # move total
                            cell.style = input_s
                            cell.number_format = "#,###.00;[RED]-#,###.00"
                        else:  # There are multiple moves
                            cell = ws_list[i].cell(row=oi, column=5)
                            cell.value = "*"  # move off
                            cell.style = input_s
                            cell = ws_list[i].cell(row=oi, column=6)
                            cell.value = "*"  # move on
                            cell.style = input_s
                            cell = ws_list[i].cell(row=oi, column=7)
                            cell.value = "*"  # route
                            cell.style = input_s
                            formula = "=SUM(%s!H%s:H%s)" % (day_of_week[i], str(oi + move_count), str(oi + 1))
                            cell = ws_list[i].cell(row=oi, column=8)
                            cell.value = formula  # move total
                            cell.style = input_s
                            cell.number_format = "#,###.00;[RED]-#,###.00"
                            # input formula for overtime
                            formula = "=IF(%s!B%s =\"ns day\", %s!C%s, MAX(%s!C%s - 8, 0))" \
                                      % (day_of_week[i], str(oi), day_of_week[i], str(oi), day_of_week[i], str(oi))
                            cell = ws_list[i].cell(row=oi, column=9)
                            cell.value = formula  # overtime
                            cell.style = calcs
                            cell.number_format = "#,###.00;[RED]-#,###.00"
                            formula = "=%s!H%s" % (day_of_week[i], str(oi))  # copy data from column H/ MV total
                            cell = ws_list[i].cell(row=oi, column=10)
                            cell.value = formula  # off route
                            cell.style = calcs
                            cell.number_format = "#,###.00;[RED]-#,###.00"
                            # formula for OT off route
                            formula = "=IF(OR(%s!B%s=\"ns day\",%s!J%s >= %s!C%s),%s!C%s, IF(%s!C%s <= 8 + " \
                                      "reference!C4, 0, MIN" \
                                      "(MAX(%s!C%s - 8, 0),IF(%s!J%s <= reference!C4,0, %s!J%s))))" \
                                      % (day_of_week[i], str(oi), day_of_week[i], str(oi), day_of_week[i], str(oi),
                                         day_of_week[i], str(oi), day_of_week[i], str(oi),
                                         day_of_week[i], str(oi), day_of_week[i], str(oi), day_of_week[i], str(oi))
                            cell = ws_list[i].cell(row=oi, column=11)
                            cell.value = formula  # OT off route
                            cell.style = calcs
                            cell.number_format = "#,###.00;[RED]-#,###.00"
                            oi += 1
                            for ii in range(move_count):  # if there are multiple moves, create + populate cells
                                cell = ws_list[i].cell(row=oi, column=5)
                                cell.value = float(s_moves[count])  # move off
                                cell.style = input_s
                                cell.number_format = "#,###.00;[RED]-#,###.00"
                                count += 1
                                cell = ws_list[i].cell(row=oi, column=6)
                                cell.value = float(s_moves[count])  # move on
                                cell.style = input_s
                                cell.number_format = "#,###.00;[RED]-#,###.00"
                                count += 1
                                cell = ws_list[i].cell(row=oi, column=7)
                                cell.value = str(s_moves[count])  # route
                                cell.style = input_s
                                count += 1
                                formula = "=SUM(%s!F%s - %s!E%s)" % (day_of_week[i], str(oi), day_of_week[i], str(oi))
                                cell = ws_list[i].cell(row=oi, column=8)
                                cell.value = formula  # move total
                                cell.style = input_s
                                cell.number_format = "#,###.00;[RED]-#,###.00"
                                if ii < move_count - 1:
                                    oi += 1
                            oi += 1
                        if move_count < 2:
                            # input formula for overtime
                            formula = "=IF(%s!B%s =\"ns day\", %s!C%s, MAX(%s!C%s - 8, 0))" \
                                      % (day_of_week[i], str(oi), day_of_week[i], str(oi), day_of_week[i], str(oi))
                            cell = ws_list[i].cell(row=oi, column=9)
                            cell.value = formula  # overtime
                            cell.style = calcs
                            cell.number_format = "#,###.00;[RED]-#,###.00"
                            # formula for off route
                            formula = "=SUM(%s!F%s - %s!E%s)" \
                                      % (day_of_week[i], str(oi), day_of_week[i], str(oi))
                            cell = ws_list[i].cell(row=oi, column=10)
                            cell.value = formula  # off route
                            cell.style = calcs
                            cell.number_format = "#,###.00;[RED]-#,###.00"
                            # formula for OT off route
                            formula = "=IF(OR(%s!B%s=\"ns day\",%s!J%s >= %s!C%s),%s!C%s, IF(%s!C%s <= 8 + " \
                                      "reference!C4, 0, MIN" \
                                      "(MAX(%s!C%s - 8, 0),IF(%s!J%s <= reference!C4,0, %s!J%s))))" \
                                      % (day_of_week[i], str(oi), day_of_week[i], str(oi), day_of_week[i], str(oi),
                                         day_of_week[i], str(oi), day_of_week[i], str(oi),
                                         day_of_week[i], str(oi), day_of_week[i], str(oi), day_of_week[i], str(oi))
                            cell = ws_list[i].cell(row=oi, column=11)
                            cell.value = formula  # OT off route
                            cell.style = calcs
                            cell.number_format = "#,###.00;[RED]-#,###.00"
                            oi += 1
            #  if there is no match, then just printe the name.
            if match == "miss":
                cell = ws_list[i].cell(row=oi, column=1)
                cell.value = line[1]  # name
                cell.style = input_name
                cell = ws_list[i].cell(row=oi, column=2)
                cell.style = input_s
                cell.number_format = "#,###.00;[RED]-#,###.00"
                cell = ws_list[i].cell(row=oi, column=3)
                cell.style = input_s
                cell.number_format = "#,###.00;[RED]-#,###.00"
                cell = ws_list[i].cell(row=oi, column=4)
                cell.style = input_s
                cell.number_format = "#,###.00;[RED]-#,###.00"
                cell = ws_list[i].cell(row=oi, column=5)
                cell.style = input_s
                cell.number_format = "#,###.00;[RED]-#,###.00"
                cell = ws_list[i].cell(row=oi, column=6)
                cell.style = input_s
                cell.number_format = "#,###.00;[RED]-#,###.00"
                cell = ws_list[i].cell(row=oi, column=7)
                cell.style = input_s
                cell.number_format = "####"
                formula = "=SUM(%s!F%s - %s!E%s)" % (day_of_week[i], str(oi), day_of_week[i], str(oi))
                cell = ws_list[i].cell(row=oi, column=8)
                cell.value = formula  # move total
                cell.style = input_s
                cell.number_format = "#,###.00;[RED]-#,###.00"
                formula = "=IF(%s!B%s =\"ns day\", %s!C%s, MAX(%s!C%s - 8, 0))" \
                          % (day_of_week[i], str(oi), day_of_week[i], str(oi), day_of_week[i], str(oi))
                cell = ws_list[i].cell(row=oi, column=9)
                cell.value = formula  # overtime
                cell.style = calcs
                cell.number_format = "#,###.00;[RED]-#,###.00"
                formula = "=SUM(%s!F%s - %s!E%s)" \
                          % (day_of_week[i], str(oi), day_of_week[i], str(oi))
                cell = ws_list[i].cell(row=oi, column=10)
                cell.value = formula  # off route
                cell.style = calcs
                cell.number_format = "#,###.00;[RED]-#,###.00"
                formula = "=IF(OR(%s!B%s=\"ns day\",%s!J%s >= %s!C%s),%s!C%s, IF(%s!C%s <= 8 + reference!C4, 0, MIN" \
                          "(MAX(%s!C%s - 8, 0),IF(%s!J%s <= reference!C4,0, %s!J%s))))" \
                          % (day_of_week[i], str(oi), day_of_week[i], str(oi), day_of_week[i], str(oi),
                             day_of_week[i], str(oi), day_of_week[i], str(oi),
                             day_of_week[i], str(oi), day_of_week[i], str(oi), day_of_week[i], str(oi))
                cell = ws_list[i].cell(row=oi, column=11)
                cell.value = formula  # OT off route
                cell.style = calcs
                cell.number_format = "#,###.00;[RED]-#,###.00"
                oi += 1
        wal_oi_end = oi
        wal_oi_diff = wal_oi_end - wal_oi_start  # find how many lines exist in nl
        # if the minimum number of rows are not reached, insert blank rows
        e_range = min_ss_wal - wal_oi_diff
        if e_range <= 0:
            e_range = 0
        for e in range(e_range):
            cell = ws_list[i].cell(row=oi, column=1)
            cell.style = input_name
            cell = ws_list[i].cell(row=oi, column=2)
            cell.style = input_s
            cell.number_format = "#,###.00;[RED]-#,###.00"
            cell = ws_list[i].cell(row=oi, column=3)
            cell.style = input_s
            cell.number_format = "#,###.00;[RED]-#,###.00"
            cell = ws_list[i].cell(row=oi, column=4)
            cell.style = input_s
            cell.number_format = "#,###.00;[RED]-#,###.00"
            cell = ws_list[i].cell(row=oi, column=5)
            cell.style = input_s
            cell.number_format = "#,###.00;[RED]-#,###.00"
            cell = ws_list[i].cell(row=oi, column=6)
            cell.style = input_s
            cell.number_format = "#,###.00;[RED]-#,###.00"
            cell = ws_list[i].cell(row=oi, column=7)
            cell.style = input_s
            cell.number_format = "####"
            formula = "=SUM(%s!F%s - %s!E%s)" % (day_of_week[i], str(oi), day_of_week[i], str(oi))
            cell = ws_list[i].cell(row=oi, column=8)
            cell.value = formula  # move total
            cell.style = input_s
            cell.number_format = "#,###.00;[RED]-#,###.00"
            formula = "=IF(%s!B%s =\"ns day\", %s!C%s,IF(%s!C%s <= 8 + reference!C3, 0, MAX(%s!C%s - 8, 0)))" \
                      % (day_of_week[i], str(oi), day_of_week[i], str(oi), day_of_week[i], str(oi),
                         day_of_week[i], str(oi))
            cell = ws_list[i].cell(row=oi, column=9)
            cell.value = formula  # overtime
            cell.style = calcs
            cell.number_format = "#,###.00;[RED]-#,###.00"
            formula = "=SUM(%s!F%s - %s!E%s)" \
                      % (day_of_week[i], str(oi), day_of_week[i], str(oi))
            cell = ws_list[i].cell(row=oi, column=10)
            cell.value = formula  # off route
            cell.style = calcs
            cell.number_format = "#,###.00;[RED]-#,###.00"
            formula = "=IF(OR(%s!B%s=\"ns day\",%s!J%s >= %s!C%s),%s!C%s, IF(%s!C%s <= 8 + reference!C4, 0, MIN" \
                      "(MAX(%s!C%s - 8, 0),IF(%s!J%s <= reference!C4,0, %s!J%s))))" \
                      % (day_of_week[i], str(oi), day_of_week[i], str(oi), day_of_week[i], str(oi),
                         day_of_week[i], str(oi), day_of_week[i], str(oi),
                         day_of_week[i], str(oi), day_of_week[i], str(oi), day_of_week[i], str(oi))
            cell = ws_list[i].cell(row=oi, column=11)
            cell.value = formula  # OT off route
            cell.style = calcs
            cell.number_format = "#,###.00;[RED]-#,###.00"
            oi += 1
        cello = str(oi - 1)
        oi += 1
        cell = ws_list[i].cell(row=oi, column=10)
        cell.value = "Total WAL Mandates"
        cell.style = col_header
        formula = "=SUM(%s!K%s:K%s)" % (day_of_week[i], top_cell, cello)
        cell = ws_list[i].cell(row=oi, column=11)
        cell.value = formula  # OT off route
        cell.style = calcs
        cell.number_format = "#,###.00;[RED]-#,###.00"
        formula = "=SUM(%s!K%s + %s!K%s)" % (day_of_week[i], str(oi), day_of_week[i], str(nl_totals))
        oi += 2
        cell = ws_list[i].cell(row=oi, column=10)
        cell.value = "Total Mandates"
        cell.style = col_header
        cell = ws_list[i].cell(row=oi, column=11)
        cell.value = formula  # total ot off route for nl and wal
        cell.style = calcs
        cell.number_format = "#,###.00;[RED]-#,###.00"
        man_ot_day.append(i)  # get the cello information to reference in the summary tab
        man_ot_row.append(oi)
        oi += 1
        try:
            ws_list[i].page_breaks.append(Break(id=oi))
        except:
            ws_list[i].row_breaks.append(Break(id=oi))
        oi += 1
        #  overtime desired list xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
        cell = ws_list[i].cell(row=oi, column=1)
        cell.value = "Overtime Desired List Carriers"
        cell.style = list_header
        oi += 1
        # column headers
        cell = ws_list[i].cell(row=oi, column=5)
        cell.value = "Availability to:"
        cell.style = col_header
        oi += 1
        cell = ws_list[i].cell(row=oi, column=1)
        cell.value = "Name"
        cell.style = col_header
        cell = ws_list[i].cell(row=oi, column=2)
        cell.value = "note"
        cell.style = col_header
        cell = ws_list[i].cell(row=oi, column=3)
        cell.value = "5200"
        cell.style = col_header
        cell = ws_list[i].cell(row=oi, column=4)
        cell.value = "RS"
        cell.style = col_header
        cell = ws_list[i].cell(row=oi, column=5)
        cell.value = "to 10"
        cell.style = col_header
        cell = ws_list[i].cell(row=oi, column=6)
        cell.value = "to 12"
        cell.style = col_header
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
                        cell = ws_list[i].cell(row=oi, column=1)
                        cell.value = each[1]  # name
                        cell.style = input_name
                        if each[4] == "none":
                            code = ""  # leave code field blank if 'none'
                        else:
                            code = each[4]
                        cell = ws_list[i].cell(row=oi, column=2)
                        cell.value = code  # code
                        cell.style = input_s
                        cell = ws_list[i].cell(row=oi, column=3)
                        if each[2].strip() == "":
                            cell.value = each[2]  # 5200
                        else:
                            cell.value = float(each[2])  # 5200
                        cell.style = input_s
                        cell.number_format = "#,###.00;[RED]-#,###.00"
                        if each[3].strip() == "":
                            rs = ""  # handle empty RS strings
                        else:
                            rs = float(each[3])
                        cell = ws_list[i].cell(row=oi, column=4)
                        cell.value = rs  # rs
                        cell.style = input_s
                        cell.number_format = "#,###.00;[RED]-#,###.00"
                        formula = "=IF(OR(%s!B%s = \"light\",%s!B%s = \"excused\", " \
                                  "%s!B%s = \"sch chg\", %s!B%s = \"annual\", " \
                                  "%s!B%s = \"sick\", %s!C%s >= 10 - reference!C5), 0, IF(%s!B%s = \"no call\", " \
                                  "10, IF(%s!C%s = 0, 0, MAX(10 - %s!C%s, 0))))" % (
                                      day_of_week[i], str(oi), day_of_week[i], str(oi),
                                      day_of_week[i], str(oi), day_of_week[i], str(oi),
                                      day_of_week[i], str(oi), day_of_week[i], str(oi),
                                      day_of_week[i], str(oi), day_of_week[i], str(oi),
                                      day_of_week[i], str(oi))
                        cell = ws_list[i].cell(row=oi, column=5)
                        cell.value = formula  # availability to 10
                        cell.style = calcs
                        cell.number_format = "#,###.00;[RED]-#,###.00"
                        formula = "=IF(OR(%s!B%s = \"light\",%s!B%s = \"excused\", " \
                                  "%s!B%s = \"sch chg\", %s!B%s = \"annual\", " \
                                  "%s!B%s = \"sick\", %s!C%s >= 12 - reference!C5), 0, IF(%s!B%s = \"no call\", " \
                                  "12, IF(%s!C%s = 0, 0, MAX(12 - %s!C%s, 0))))" % (
                                      day_of_week[i], str(oi), day_of_week[i], str(oi),
                                      day_of_week[i], str(oi), day_of_week[i], str(oi),
                                      day_of_week[i], str(oi), day_of_week[i], str(oi),
                                      day_of_week[i], str(oi), day_of_week[i], str(oi),
                                      day_of_week[i], str(oi))
                        cell = ws_list[i].cell(row=oi, column=6)
                        cell.value = formula  # availability to 12
                        cell.style = calcs
                        cell.number_format = "#,###.00;[RED]-#,###.00"
                        oi += 1
            # if there is no match, then just printe the name.
            if match == "miss":
                cell = ws_list[i].cell(row=oi, column=1)
                cell.value = line[1]  # name
                cell.style = input_name
                cell = ws_list[i].cell(row=oi, column=2)
                cell.style = input_s
                cell.number_format = "#,###.00;[RED]-#,###.00"
                cell = ws_list[i].cell(row=oi, column=3)
                cell.style = input_s
                cell.number_format = "#,###.00;[RED]-#,###.00"
                cell = ws_list[i].cell(row=oi, column=4)
                cell.style = input_s
                cell.number_format = "#,###.00;[RED]-#,###.00"
                formula = "=IF(OR(%s!B%s = \"light\",%s!B%s = \"excused\", " \
                          "%s!B%s = \"sch chg\", %s!B%s = \"annual\", " \
                          "%s!B%s = \"sick\", %s!C%s >= 10 - reference!C5), 0, IF(%s!B%s = \"no call\", " \
                          "10, IF(%s!C%s = 0, 0, MAX(10 - %s!C%s, 0))))" % (
                              day_of_week[i], str(oi), day_of_week[i], str(oi),
                              day_of_week[i], str(oi), day_of_week[i], str(oi),
                              day_of_week[i], str(oi), day_of_week[i], str(oi),
                              day_of_week[i], str(oi), day_of_week[i], str(oi),
                              day_of_week[i], str(oi))
                cell = ws_list[i].cell(row=oi, column=5)
                cell.value = formula  # availability to 10
                cell.style = calcs
                cell.number_format = "#,###.00;[RED]-#,###.00"
                formula = "=IF(OR(%s!B%s = \"light\",%s!B%s = \"excused\", " \
                          "%s!B%s = \"sch chg\", %s!B%s = \"annual\", " \
                          "%s!B%s = \"sick\", %s!C%s >= 12 - reference!C5), 0, IF(%s!B%s = \"no call\", " \
                          "12, IF(%s!C%s = 0, 0, MAX(12 - %s!C%s, 0))))" % (
                              day_of_week[i], str(oi), day_of_week[i], str(oi),
                              day_of_week[i], str(oi), day_of_week[i], str(oi),
                              day_of_week[i], str(oi), day_of_week[i], str(oi),
                              day_of_week[i], str(oi), day_of_week[i], str(oi),
                              day_of_week[i], str(oi))
                cell = ws_list[i].cell(row=oi, column=6)
                cell.value = formula  # availability to 12
                cell.style = calcs
                cell.number_format = "#,###.00;[RED]-#,###.00"
                oi += 1
        otdl_oi_end = oi
        otdl_oi_diff = otdl_oi_end - otdl_oi_start  # find how many lines exist in otdl
        # if the minimum number of rows are not reached, insert blank rows
        e_range = min_ss_otdl - otdl_oi_diff
        if e_range <= 0:
            e_range = 0
        for e in range(e_range):
            cell = ws_list[i].cell(row=oi, column=1)
            cell.value = ""  # name
            cell.style = input_name
            cell = ws_list[i].cell(row=oi, column=2)
            cell.style = input_s
            cell.number_format = "#,###.00;[RED]-#,###.00"
            cell = ws_list[i].cell(row=oi, column=3)
            cell.style = input_s
            cell.number_format = "#,###.00;[RED]-#,###.00"
            cell = ws_list[i].cell(row=oi, column=4)
            cell.style = input_s
            cell.number_format = "#,###.00;[RED]-#,###.00"
            formula = "=IF(OR(%s!B%s = \"light\",%s!B%s = \"excused\", %s!B%s = \"sch chg\", %s!B%s = \"annual\", " \
                      "%s!B%s = \"sick\", %s!C%s >= 10 - reference!C5), 0, IF(%s!B%s = \"no call\", " \
                      "10, IF(%s!C%s = 0, 0, MAX(10 - %s!C%s, 0))))" % (
                          day_of_week[i], str(oi), day_of_week[i], str(oi),
                          day_of_week[i], str(oi), day_of_week[i], str(oi),
                          day_of_week[i], str(oi), day_of_week[i], str(oi),
                          day_of_week[i], str(oi), day_of_week[i], str(oi),
                          day_of_week[i], str(oi))
            cell = ws_list[i].cell(row=oi, column=5)
            cell.value = formula  # availability to 10
            cell.style = calcs
            cell.number_format = "#,###.00;[RED]-#,###.00"
            formula = "=IF(OR(%s!B%s = \"light\",%s!B%s = \"excused\", %s!B%s = \"sch chg\", %s!B%s = \"annual\", " \
                      "%s!B%s = \"sick\", %s!C%s >= 12 - reference!C5), 0, IF(%s!B%s = \"no call\", " \
                      "12, IF(%s!C%s = 0, 0, MAX(12 - %s!C%s, 0))))" % (
                          day_of_week[i], str(oi), day_of_week[i], str(oi),
                          day_of_week[i], str(oi), day_of_week[i], str(oi),
                          day_of_week[i], str(oi), day_of_week[i], str(oi),
                          day_of_week[i], str(oi), day_of_week[i], str(oi),
                          day_of_week[i], str(oi))
            cell = ws_list[i].cell(row=oi, column=6)
            cell.value = formula  # availability to 12
            cell.style = calcs
            cell.number_format = "#,###.00;[RED]-#,###.00"
            oi += 1
        oi += 1
        cello = str(oi - 2)
        cell = ws_list[i].cell(row=oi, column=4)
        cell.value = "Total OTDL Availability"
        cell.style = col_header
        formula = "=SUM(%s!E%s:E%s)" % (day_of_week[i], top_cell, cello)
        otdl_total = oi
        cell = ws_list[i].cell(row=oi, column=5)
        cell.value = formula  # availability to 10
        cell.style = calcs
        cell.number_format = "#,###.00;[RED]-#,###.00"
        formula = "=SUM(%s!F%s:F%s)" % (day_of_week[i], top_cell, cello)
        cell = ws_list[i].cell(row=oi, column=6)
        cell.value = formula  # availability to 12
        cell.style = calcs
        cell.number_format = "#,###.00;[RED]-#,###.00"
        oi += 1
        try:
            ws_list[i].page_breaks.append(Break(id=oi))
        except:
            ws_list[i].row_breaks.append(Break(id=oi))
        oi += 1
        # Auxiliary assistance xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
        cell = ws_list[i].cell(row=oi, column=1)
        cell.value = "Auxiliary Assistance"
        cell.style = list_header
        oi += 1
        # column headers
        cell = ws_list[i].cell(row=oi, column=5)
        cell.value = "Availability to:"
        cell.style = col_header
        oi += 1
        cell = ws_list[i].cell(row=oi, column=1)
        cell.value = "Name"
        cell.style = col_header
        cell = ws_list[i].cell(row=oi, column=2)
        cell.value = "note"
        cell.style = col_header
        cell = ws_list[i].cell(row=oi, column=3)
        cell.value = "5200"
        cell.style = col_header
        cell = ws_list[i].cell(row=oi, column=4)
        cell.value = "RS"
        cell.style = col_header
        cell = ws_list[i].cell(row=oi, column=5)
        cell.value = "to 10"
        cell.style = col_header
        cell = ws_list[i].cell(row=oi, column=6)
        cell.value = "to 11.5"
        cell.style = col_header
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
                        cell = ws_list[i].cell(row=oi, column=1)
                        cell.value = each[1]  # name
                        cell.style = input_name
                        if each[4] == "none":
                            code = ""  # leave code field blank if 'none'
                        else:
                            code = each[4]
                        cell = ws_list[i].cell(row=oi, column=2)
                        cell.value = code  # code
                        cell.style = input_s
                        cell = ws_list[i].cell(row=oi, column=3)
                        if each[2].strip() == "":
                            cell.value = each[2]  # 5200
                        else:
                            cell.value = float(each[2])  # 5200
                        cell.style = input_s
                        cell.number_format = "#,###.00;[RED]-#,###.00"
                        if each[3].strip() == "":
                            rs = ""  # handle empty RS strings
                        else:
                            rs = float(each[3])
                        cell = ws_list[i].cell(row=oi, column=4)
                        cell.value = rs  # rs
                        cell.style = input_s
                        cell.number_format = "#,###.00;[RED]-#,###.00"
                        formula = "=IF(OR(%s!B%s = \"light\",%s!B%s = \"excused\", " \
                                  "%s!B%s = \"sch chg\", %s!B%s = \"annual\", " \
                                  "%s!B%s = \"sick\", %s!C%s >= 10 - reference!C5), 0, IF(%s!B%s = \"no call\", " \
                                  "10, IF(%s!C%s = 0, 0, MAX(10 - %s!C%s, 0))))" % (
                                      day_of_week[i], str(oi), day_of_week[i], str(oi),
                                      day_of_week[i], str(oi), day_of_week[i], str(oi),
                                      day_of_week[i], str(oi), day_of_week[i], str(oi),
                                      day_of_week[i], str(oi), day_of_week[i], str(oi),
                                      day_of_week[i], str(oi))
                        cell = ws_list[i].cell(row=oi, column=5)
                        cell.value = formula  # availability to 10
                        cell.style = calcs
                        cell.number_format = "#,###.00;[RED]-#,###.00"
                        formula = "=IF(OR(%s!B%s = \"light\",%s!B%s = \"excused\", " \
                                  "%s!B%s = \"sch chg\", %s!B%s = \"annual\", " \
                                  "%s!B%s = \"sick\", %s!C%s >= 11.5 - reference!C5), 0, IF(%s!B%s = \"no call\", " \
                                  "11.5, IF(%s!C%s = 0, 0, MAX(11.5 - %s!C%s, 0))))" % (
                                      day_of_week[i], str(oi), day_of_week[i], str(oi),
                                      day_of_week[i], str(oi), day_of_week[i], str(oi),
                                      day_of_week[i], str(oi), day_of_week[i], str(oi),
                                      day_of_week[i], str(oi), day_of_week[i], str(oi),
                                      day_of_week[i], str(oi))
                        cell = ws_list[i].cell(row=oi, column=6)
                        cell.value = formula  # availability to 12
                        cell.style = calcs
                        cell.number_format = "#,###.00;[RED]-#,###.00"
                        oi += 1
            # if there is no match, then just printe the name.
            if match == "miss":
                if match == "miss":
                    cell = ws_list[i].cell(row=oi, column=1)
                    cell.value = line[1]  # name
                    cell.style = input_name
                    cell = ws_list[i].cell(row=oi, column=2)
                    cell.style = input_s
                    cell.number_format = "#,###.00;[RED]-#,###.00"
                    cell = ws_list[i].cell(row=oi, column=3)
                    cell.style = input_s
                    cell.number_format = "#,###.00;[RED]-#,###.00"
                    cell = ws_list[i].cell(row=oi, column=4)
                    cell.style = input_s
                    cell.number_format = "#,###.00;[RED]-#,###.00"
                    formula = "=IF(OR(%s!B%s = \"light\",%s!B%s = \"excused\", " \
                              "%s!B%s = \"sch chg\", %s!B%s = \"annual\", " \
                              "%s!B%s = \"sick\", %s!C%s >= 10 - reference!C5), 0, IF(%s!B%s = \"no call\", " \
                              "10, IF(%s!C%s = 0, 0, MAX(10 - %s!C%s, 0))))" % (
                                  day_of_week[i], str(oi), day_of_week[i], str(oi),
                                  day_of_week[i], str(oi), day_of_week[i], str(oi),
                                  day_of_week[i], str(oi), day_of_week[i], str(oi),
                                  day_of_week[i], str(oi), day_of_week[i], str(oi),
                                  day_of_week[i], str(oi))
                    cell = ws_list[i].cell(row=oi, column=5)
                    cell.value = formula  # availability to 10
                    cell.style = calcs
                    cell.number_format = "#,###.00;[RED]-#,###.00"
                    formula = "=IF(OR(%s!B%s = \"light\",%s!B%s = \"excused\", " \
                              "%s!B%s = \"sch chg\", %s!B%s = \"annual\", " \
                              "%s!B%s = \"sick\", %s!C%s >= 11.5 - reference!C5), 0, IF(%s!B%s = \"no call\", " \
                              "11.5, IF(%s!C%s = 0, 0, MAX(11.5 - %s!C%s, 0))))" % (
                                  day_of_week[i], str(oi), day_of_week[i], str(oi),
                                  day_of_week[i], str(oi), day_of_week[i], str(oi),
                                  day_of_week[i], str(oi), day_of_week[i], str(oi),
                                  day_of_week[i], str(oi), day_of_week[i], str(oi),
                                  day_of_week[i], str(oi))
                    cell = ws_list[i].cell(row=oi, column=6)
                    cell.value = formula  # availability to 12
                    cell.style = calcs
                    cell.number_format = "#,###.00;[RED]-#,###.00"
                    oi += 1
        aux_oi_end = oi
        aux_oi_diff = aux_oi_end - aux_oi_start  # find how many lines exist in aux
        # if the minimum number of rows are not reached, insert blank rows
        e_range = min_ss_aux - aux_oi_diff
        if e_range <= 0:
            e_range = 0
        for e in range(e_range):
            cell = ws_list[i].cell(row=oi, column=1)
            cell.value = ""  # name
            cell.style = input_name
            cell = ws_list[i].cell(row=oi, column=2)
            cell.style = input_s
            cell.number_format = "#,###.00;[RED]-#,###.00"
            cell = ws_list[i].cell(row=oi, column=3)
            cell.style = input_s
            cell.number_format = "#,###.00;[RED]-#,###.00"
            cell = ws_list[i].cell(row=oi, column=4)
            cell.style = input_s
            cell.number_format = "#,###.00;[RED]-#,###.00"
            formula = "=IF(OR(%s!B%s = \"light\",%s!B%s = \"excused\", %s!B%s = \"sch chg\", %s!B%s = \"annual\", " \
                      "%s!B%s = \"sick\", %s!C%s >= 10 - reference!C5), 0, IF(%s!B%s = \"no call\", " \
                      "10, IF(%s!C%s = 0, 0, MAX(10 - %s!C%s, 0))))" % (
                          day_of_week[i], str(oi), day_of_week[i], str(oi),
                          day_of_week[i], str(oi), day_of_week[i], str(oi),
                          day_of_week[i], str(oi), day_of_week[i], str(oi),
                          day_of_week[i], str(oi), day_of_week[i], str(oi),
                          day_of_week[i], str(oi))
            cell = ws_list[i].cell(row=oi, column=5)
            cell.value = formula  # availability to 10
            cell.style = calcs
            cell.number_format = "#,###.00;[RED]-#,###.00"
            formula = "=IF(OR(%s!B%s = \"light\",%s!B%s = \"excused\", %s!B%s = \"sch chg\", %s!B%s = \"annual\", " \
                      "%s!B%s = \"sick\", %s!C%s >= 12 - reference!C5), 0, IF(%s!B%s = \"no call\", " \
                      "12, IF(%s!C%s = 0, 0, MAX(12 - %s!C%s, 0))))" % (
                          day_of_week[i], str(oi), day_of_week[i], str(oi),
                          day_of_week[i], str(oi), day_of_week[i], str(oi),
                          day_of_week[i], str(oi), day_of_week[i], str(oi),
                          day_of_week[i], str(oi), day_of_week[i], str(oi),
                          day_of_week[i], str(oi))
            cell = ws_list[i].cell(row=oi, column=6)
            cell.value = formula  # availability to 12
            cell.style = calcs
            cell.number_format = "#,###.00;[RED]-#,###.00"
            oi += 1
        oi += 1
        cello = str(oi - 2)
        cell = ws_list[i].cell(row=oi, column=4)
        cell.value = "Total AUX Availability"
        cell.style = col_header
        formula = "=SUM(%s!E%s:E%s)" % (day_of_week[i], top_cell, cello)
        aux_total = oi
        cell = ws_list[i].cell(row=oi, column=5)
        cell.value = formula  # availability to 10
        cell.style = calcs
        cell.number_format = "#,###.00;[RED]-#,###.00"
        formula = "=SUM(%s!F%s:F%s)" % (day_of_week[i], top_cell, cello)
        cell = ws_list[i].cell(row=oi, column=6)
        cell.value = formula  # availability to 11.5
        cell.style = calcs
        cell.number_format = "#,###.00;[RED]-#,###.00"
        oi += 2
        cell = ws_list[i].cell(row=oi, column=4)
        cell.value = "Total Availability"
        cell.style = col_header
        formula = "=SUM(%s!E%s + %s!E%s)" % (day_of_week[i], otdl_total, day_of_week[i], aux_total)
        cell = ws_list[i].cell(row=oi, column=5)
        cell.value = formula  # availability to 10
        cell.style = calcs
        cell.number_format = "#,###.00;[RED]-#,###.00"
        av_to_10_day.append(i)
        av_to_10_row.append(oi)
        formula = "=SUM(%s!F%s + %s!F%s)" % (day_of_week[i], otdl_total, day_of_week[i], aux_total)
        cell = ws_list[i].cell(row=oi, column=6)
        cell.value = formula  # availability to 11.5
        cell.style = calcs
        cell.number_format = "#,###.00;[RED]-#,###.00"
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
    summary['B5'] = projvar.pay_period
    summary['B5'].style = date_dov
    summary.merge_cells('B5:D5')
    summary['A6'] = "Station:  "
    summary['A6'].style = date_dov_title
    summary['B6'] = projvar.invran_station
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
    range_num = 0
    if projvar.invran_weekly_span:  # if investigation range is weekly
        range_num = 7
    if not projvar.invran_weekly_span:  # if investigation range is daily
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
                                  % ('summary', str(row), 'summary', str(row), 'summary',
                                     str(row), 'summary', str(row))
        summary['D' + str(row)].style = calcs
        summary['D' + str(row)].number_format = "#,###.00;[RED]-#,###.00"
        summary['F' + str(row)] = "=%s!F%s" % (day_of_week[av_to_12_day[i]], av_to_12_row[i])  # availability to 12
        summary['F' + str(row)].style = input_s
        summary['F' + str(row)].number_format = "#,###.00;[RED]-#,###.00"
        summary['G' + str(row)] = "=%s!K%s" % (day_of_week[man_ot_day[i]], man_ot_row[i])  # total mandates
        summary['G' + str(row)].style = input_s
        summary['G' + str(row)].number_format = "#,###.00;[RED]-#,###.00"
        summary['H' + str(row)] = "=IF(%s!F%s<%s!G%s,%s!F%s,%s!G%s)" \
                                  % ('summary', str(row), 'summary', str(row), 'summary',
                                     str(row), 'summary', str(row))
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
    if not projvar.invran_weekly_span:  # if investigation range is daily
        r = "_d"
    xl_filename = "kb" + str(format(dates[0], "_%y_%m_%d")) + r + ".xlsx"
    if messagebox.askokcancel("Spreadsheet generator",
                              "Do you want to generate a spreadsheet?",
                              parent=frame):
        try:
            wb.save(dir_path('spreadsheets') + xl_filename)
            messagebox.showinfo("Spreadsheet generator",
                                "Your spreadsheet was successfully generated. \n"
                                "File is named: {}".format(xl_filename),
                                parent=frame)
        except PermissionError:
            messagebox.showerror("Spreadsheet generator",
                                 "The spreadsheet was not generated. \n"
                                 "Suggestion: "
                                 "Make sure that identically named spreadsheets are closed "
                                 "(the file can't be overwritten while open).",
                                 parent=frame)
        try:
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
                                 parent=frame)


class OvermaxSpreadsheet:
    def __init__(self):
        self.frame = None
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

    def create(self, frame):  # master method for calling methods
        self.frame = frame
        if not self.ask_ok():
            return
        self.get_dates()
        self.get_carrierlist()
        self.get_rings()
        self.get_minrows()
        self.get_styles()
        self.build_workbook()
        self.set_dimensions()
        self.build_summary()
        self.build_violations()
        self.build_instructions()
        self.violated_recs()
        self.show_violations()
        # self.build_joint()
        self.save_open()

    def ask_ok(self):
        if messagebox.askokcancel("Spreadsheet generator",
                                  "Do you want to generate a spreadsheet?",
                                  parent=self.frame):
            return True
        return False

    def get_dates(self):
        date = projvar.invran_date_week[0]
        self.startdate = projvar.invran_date_week[0]
        self.enddate = projvar.invran_date_week[6]
        for i in range(7):
            self.dates.append(date)
            date += timedelta(days=1)

    def get_carrierlist(self):
        carrierlist = CarrierList(self.startdate, self.enddate, projvar.invran_station).get()
        for carrier in carrierlist:
            self.carrier_list.append(carrier[0])  # add the first record for each carrier in rec set

    def get_rings(self):
        sql = "SELECT * FROM rings3 WHERE rings_date BETWEEN '%s' AND '%s' ORDER BY rings_date, carrier_name" \
              % (projvar.invran_date_week[0], projvar.invran_date_week[6])
        self.rings = inquire(sql)

    def get_minrows(self):  # get minimum rows
        sql = "SELECT tolerance FROM tolerances"
        result = inquire(sql)
        self.min_rows = int(result[14][0])

    def get_styles(self):  # Named styles for workbook
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
                                border=Border(left=bd, right=bd, top=bd, bottom=bd))
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
        self.wb = Workbook()  # define the workbook
        self.violations = self.wb.active  # create first worksheet
        self.violations.title = "violations"  # title first worksheet
        self.violations.oddFooter.center.text = "&A"
        self.summary = self.wb.create_sheet("summary")
        self.summary.oddFooter.center.text = "&A"
        self.instructions = self.wb.create_sheet("instructions")

    def set_dimensions(self):
        for x in range(2, 8):
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
            sheet.column_dimensions["X"].width = 5

    def build_summary(self):
        # summary worksheet - format cells
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
        # self.violations worksheet - format cells
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
        self.violations.merge_cells('D6:Q6')
        self.violations['D6'] = "Daily Paid Leave times with type"
        self.violations['D6'].style = self.col_center_header
        self.violations.merge_cells('D7:Q7')
        self.violations['D7'] = "Daily 5200 times"
        self.violations['D7'].style = self.col_center_header
        self.violations['A8'] = "name"
        self.violations['A8'].style = self.col_header
        self.violations['B8'] = "list"
        self.violations['B8'].style = self.col_header
        self.violations.merge_cells('C5:C8')
        self.violations['C5'] = "Weekly\n5200"
        self.violations['C5'].style = self.vert_header
        self.violations.merge_cells('D8:E8')
        self.violations['D8'] = "sat"
        self.violations['D8'].style = self.col_center_header
        self.violations.merge_cells('F8:G8')
        self.violations['F8'] = "sun"
        self.violations['F8'].style = self.col_center_header
        self.violations.merge_cells('H8:I8')
        self.violations['H8'] = "mon"
        self.violations['H8'].style = self.col_center_header
        self.violations.merge_cells('J8:K8')
        self.violations['J8'] = "tue"
        self.violations['J8'].style = self.col_center_header
        self.violations.merge_cells('L8:M8')
        self.violations['L8'] = "wed"
        self.violations['L8'].style = self.col_center_header
        self.violations.merge_cells('N8:O8')
        self.violations['N8'] = "thr"
        self.violations['N8'].style = self.col_center_header
        self.violations.merge_cells('P8:Q8')
        self.violations['P8'] = "fri"
        self.violations['P8'].style = self.col_center_header
        self.violations.merge_cells('S4:S8')
        self.violations['S4'] = " Weekly\nViolation"
        self.violations['S4'].style = self.vert_header
        self.violations.merge_cells('T4:T8')
        self.violations['T4'] = "Daily\nViolation"
        self.violations['T4'].style = self.vert_header
        self.violations.merge_cells('U4:U8')
        self.violations['U4'] = "Wed Adj"
        self.violations['U4'].style = self.vert_header
        self.violations.merge_cells('V4:V8')
        self.violations['V4'] = "Thr Adj"
        self.violations['V4'].style = self.vert_header
        self.violations.merge_cells('W4:W8')
        self.violations['W4'] = "Fri Adj"
        self.violations['W4'].style = self.vert_header
        self.violations.merge_cells('X4:X8')
        self.violations['X4'] = "Total\nViolation"
        self.violations['X4'].style = self.vert_header

    def build_instructions(self):
        # format the instructions cells
        self.instructions.merge_cells('A1:R1')
        self.instructions['A1'] = "12 and 60 Hour Violations Instructions"
        self.instructions['A1'].style = self.ws_header
        self.instructions.row_dimensions[3].height = 165
        self.instructions['A3'].style = self.instruct_text
        self.instructions.merge_cells('A3:X3')
        self.instructions['A3'] = "Instructions: \n" \
                              "1. Fill in the name \n" \
                              "2. Fill in the list. Enter either “otdl”,”wal”,”nl”,“aux” or “ptf” in list columns. " \
                              "Use only lowercase. \n" \
                              "   If you do not enter anything, the default is “otdl\n" \
                              "\totdl = overtime desired list\n" \
                              "\twal = work assignment list\n" \
                              "\tnl = no list \n" \
                              "\taux = auxiliary (this would be a cca or city carrier assistant).\n" \
                              "\tptf = part time flexible" \
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
                              "7. Field O will show the violation in hours which you should seek a remedy for. \n"
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
        # instructions list
        self.instructions.merge_cells('B' + str(i) + ':B' + str(i + 1))  # merge box for list type input
        self.instructions['B10'] = "wal"
        self.instructions['B10'].style = self.input_s
        # instructions weekly
        self.instructions.merge_cells('C' + str(i) + ':C' + str(i + 1))  # merge box for weekly input
        self.instructions['C10'] = 75.00
        self.instructions['C10'].style = self.input_s
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
        formula = "=MAX(IF(%s!R%s>%s!R%s,MAX(%s!R%s-60,0),MAX(%s!R%s-60)),0)" % \
                  (page, str(i), page, str(i + 1), page, str(i), page, str(i + 1),)
        self.instructions['S10'] = formula
        self.instructions['S10'].style = self.calcs
        self.instructions['S10'].number_format = "#,###.00;[RED]-#,###.00"
        # instructions daily self.violations
        formula_d = "=IF(OR(%s!B%s=\"wal\",%s!B%s=\"nl\",%s!B%s=\"ptf\",%s!B%s=\"aux\")," \
                    "(SUM(IF(%s!D%s>11.5,%s!D%s-11.5,0)+IF(%s!H%s>11.5,%s!H%s-11.5,0)+IF(%s!J%s>11.5,%s!J%s-11.5,0)" \
                    "+IF(%s!L%s>11.5,%s!L%s-11.5,0)+IF(%s!N%s>11.5,%s!N%s-11.5,0)+IF(%s!P%s>11.5,%s!P%s-11.5,0)))," \
                    "(SUM(IF(%s!D%s>12,%s!D%s-12,0)+IF(%s!H%s>12,%s!H%s-12,0)+IF(%s!J%s>12,%s!J%s-12,0)" \
                    "+IF(%s!L%s>12,%s!L%s-12,0)+IF(%s!N%s>12,%s!N%s-12,0)+IF(%s!P%s>12,%s!P%s-12,0))))" \
                    % (page, str(i), page, str(i), page, str(i), page, str(i),
                       page, str(i + 1), page, str(i + 1), page, str(i + 1),
                       page, str(i + 1), page, str(i + 1), page, str(i + 1),
                       page, str(i + 1), page, str(i + 1), page, str(i + 1),
                       page, str(i + 1), page, str(i + 1), page, str(i + 1),
                       page, str(i + 1), page, str(i + 1), page, str(i + 1),
                       page, str(i + 1), page, str(i + 1), page, str(i + 1),
                       page, str(i + 1), page, str(i + 1), page, str(i + 1),
                       page, str(i + 1), page, str(i + 1), page, str(i + 1))
        self.instructions['T' + str(i)] = formula_d
        self.instructions.merge_cells('T' + str(i) + ':T' + str(i + 1))  # merge box for daily violation
        self.instructions['T' + str(i)].style = self.calcs
        self.instructions['T' + str(i)].number_format = "#,###.00"
        # instructions wed adjustment
        self.instructions.merge_cells('U' + str(i) + ':U' + str(i + 1))  # merge box for wed adj
        formula_e = "=IF(OR(%s!B%s=\"wal\",%s!B%s=\"nl\",%s!B%s=\"ptf\",%s!B%s=\"aux\")," \
                    "IF(AND(%s!S%s-(%s!N%s+%s!N%s+%s!P%s+%s!P%s)>0,%s!L%s>11.5)," \
                    "IF(%s!S%s-(%s!N%s+%s!N%s+%s!P%s+%s!P%s)>%s!L%s-11.5,%s!L%s-11.5,%s!S%s-" \
                    "(%s!N%s+%s!N%s+%s!P%s+%s!P%s)),0)," \
                    "IF(AND(%s!S%s-(%s!N%s+%s!N%s+%s!P%s+%s!P%s)>0,%s!L%s>12)," \
                    "IF(%s!S%s-(%s!N%s+%s!N%s+%s!P%s+%s!P%s)>%s!L%s-12,%s!L%s-12,%s!S%s-" \
                    "(%s!N%s+%s!N%s+%s!P%s+%s!P%s)),0))" \
                    % (page, str(i), page, str(i), page, str(i), page, str(i),
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
        self.instructions['U' + str(i)] = formula_e
        self.instructions['U' + str(i)].style = self.vert_calcs
        self.instructions['U' + str(i)].number_format = "#,###.00"
        # instructions thr adjustment
        formula_f = "=IF(OR(%s!B%s=\"wal\",%s!B%s=\"nl\",%s!B%s=\"ptf\",%s!B%s=\"aux\")," \
                    "IF(AND(%s!S%s-(%s!P%s+%s!P%s)>0,%s!N%s>11.5)," \
                    "IF(%s!S%s-(%s!P%s+%s!P%s)>%s!N%s-11.5,%s!N%s-11.5,%s!S%s-(%s!P%s+%s!P%s)),0)," \
                    "IF(AND(%s!S%s-(%s!P%s+%s!P%s)>0,%s!N%s>12)," \
                    "IF(%s!S%s-(%s!P%s+%s!P%s)>%s!N%s-12,%s!N%s-12,%s!S%s-(%s!P%s+%s!P%s)),0))" \
                    % (page, str(i), page, str(i), page, str(i), page, str(i),
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
        self.instructions.merge_cells('V' + str(i) + ':V' + str(i + 1))  # merge box for thr adj
        self.instructions['V' + str(i)] = formula_f
        self.instructions['V' + str(i)].style = self.vert_calcs
        self.instructions['V' + str(i)].number_format = "#,###.00"
        # instructions fri adjustment
        self.instructions.merge_cells('W' + str(i) + ':W' + str(i + 1))  # merge box for fri adj
        formula_g = "=IF(OR(%s!B%s=\"wal\",%s!B%s=\"nl\",%s!B%s=\"aux\",%s!B%s=\"ptf\")," \
                    "IF(AND(%s!S%s>0,%s!P%s>11.5)," \
                    "IF(%s!S%s>%s!P%s-11.5,%s!P%s-11.5,%s!S%s),0)," \
                    "IF(AND(%s!S%s>0,%s!P%s>12)," \
                    "IF(%s!S%s>%s!P%s-12,%s!P%s-12,%s!S%s),0))" \
                    % (page, str(i), page, str(i), page, str(i), page, str(i),
                       page, str(i), page, str(i + 1), page, str(i),
                       page, str(i + 1), page, str(i + 1), page, str(i),
                       page, str(i), page, str(i + 1), page, str(i),
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
        self.instructions.row_dimensions[14].height = 180
        self.instructions['A14'].style = self.instruct_text
        self.instructions.merge_cells('A14:X14')
        self.instructions['A14'] = "Legend: \n" \
                           "A.  Name \n" \
                           "B.  List: Either otdl, wal, nl, ptf or aux (always use lowercase to preserve " \
                           "operation of the formulas).\n" \
                           "C.  Weekly 5200 Time: Enter the 5200 time for the week. \n" \
                           "D.  Daily Non 5200 Time: Enter daily hours for either holiday, annual sick leave or " \
                           "other type of paid leave.\n" \
                           "E.  Daily Non 5200 Type: Enter “a” for annual, “s” for sick, “h” for holiday, etc. \n" \
                           "F.  Daily 5200 Hours: Enter 5200 hours or hours worked for the day. \n" \
                           "G.  No value allowed: No non 5200 times allowed for Sundays.\n" \
                           "J.   Weekly Violations: This is the total of self.violations over 60 hours in a week.\n" \
                           "K.  Daily Violations: This is the total of daily violations which have exceeded 11.50 " \
                           "(for wal, nl, ptf or aux)\n" \
                           "     or 12 hours in a day (for otdl).\n" \
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
                           "This is the value which the steward should seek a remedy for."

    def violated_recs(self):
        """
        The violation record set is appended if the carrier has a daily violation or a weekly violation of
        over 60 hours in a week. It consist of 4 arrays: 1. carrier info (name and list), 2. daily hours array,
        3. daily leavetypes and 4 daily leavetimes. The carrier list the status on Saturday.
        """
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
                            if float(ring[2]) > 12 and self.carrier_list[i][2] == "otdl":
                                daily_violation = True
                            if float(ring[2]) > 11.5 and self.carrier_list[i][2] != "otdl":
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
                total = total + t
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

    def show_violations(self):
        summary_i = 7
        i = 9
        for line in self.violation_recsets:
            carrier_name = line[0][0]
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
            self.violations.merge_cells('B' + str(i) + ':B' + str(i + 1))  # merge box for list
            self.violations['B' + str(i)] = carrier_list  # list
            self.violations['B' + str(i)].style = self.input_s
            self.violations.merge_cells('C' + str(i) + ':C' + str(i + 1))  # merge box for weekly 5200
            self.violations['C' + str(i)] = Convert(total).empty_not_zerofloat()  # total
            self.violations['C' + str(i)].style = self.input_s
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
            formula_c = "=MAX(IF(%s!R%s>%s!R%s,MAX(%s!R%s-60,0),MAX(%s!R%s-60)),0)" \
                        % ("violations", str(i), "violations", str(i + 1), "violations", str(i),
                           "violations", str(i + 1))
            self.violations['S' + str(i)] = formula_c
            self.violations['S' + str(i)].style = self.calcs
            self.violations['S' + str(i)].number_format = "#,###.00;[RED]-#,###.00"
            # daily violation
            formula_d = "=IF(OR(%s!B%s=\"wal\",%s!B%s=\"nl\",%s!B%s=\"ptf\",%s!B%s=\"aux\")," \
                        "(SUM(IF(%s!D%s>11.5,%s!D%s-11.5,0)+IF(%s!F%s>11.5,%s!F%s-11.5,0)" \
                        "+IF(%s!H%s>11.5,%s!H%s-11.5,0)+" \
                        "IF(%s!J%s>11.5,%s!J%s-11.5,0)" \
                        "+IF(%s!L%s>11.5,%s!L%s-11.5,0)+IF(%s!N%s>11.5,%s!N%s-11.5,0)+" \
                        "IF(%s!P%s>11.5,%s!P%s-11.5,0)))," \
                        "(SUM(IF(%s!D%s>12,%s!D%s-12,0)+IF(%s!F%s>12,%s!F%s-12,0)+IF(%s!H%s>12,%s!H%s-12,0)" \
                        "+IF(%s!J%s>12,%s!J%s-12,0)" \
                        "+IF(%s!L%s>12,%s!L%s-12,0)+IF(%s!N%s>12,%s!N%s-12,0)+IF(%s!P%s>12,%s!P%s-12,0))))" \
                        % ("violations", str(i), "violations", str(i), "violations", str(i), "violations", str(i),
                           "violations", str(i + 1), "violations", str(i + 1), "violations", str(i + 1),
                           "violations", str(i + 1), "violations", str(i + 1), "violations", str(i + 1),
                           "violations", str(i + 1), "violations", str(i + 1), "violations", str(i + 1),
                           "violations", str(i + 1), "violations", str(i + 1), "violations", str(i + 1),
                           "violations", str(i + 1), "violations", str(i + 1), "violations", str(i + 1),
                           "violations", str(i + 1), "violations", str(i + 1), "violations", str(i + 1),
                           "violations", str(i + 1), "violations", str(i + 1), "violations", str(i + 1),
                           "violations", str(i + 1), "violations", str(i + 1), "violations", str(i + 1),
                           "violations", str(i + 1), "violations", str(i + 1),
                           "violations", str(i + 1), "violations", str(i + 1))
            self.violations['T' + str(i)] = formula_d
            self.violations.merge_cells('T' + str(i) + ':T' + str(i + 1))  # merge box for daily violation
            self.violations['T' + str(i)].style = self.calcs
            self.violations['T' + str(i)].number_format = "#,###.00"
            # wed adjustment
            self.violations.merge_cells('U' + str(i) + ':U' + str(i + 1))  # merge box for wed adj
            formula_e = "=IF(OR(%s!B%s=\"wal\",%s!B%s=\"nl\",%s!B%s=\"ptf\",%s!B%s=\"aux\")," \
                        "IF(AND(%s!S%s-(%s!N%s+%s!N%s+%s!P%s+%s!P%s)>0,%s!L%s>11.5)," \
                        "IF(%s!S%s-(%s!N%s+%s!N%s+%s!P%s+%s!P%s)>%s!L%s-11.5,%s!L%s-11.5,%s!S%s-" \
                        "(%s!N%s+%s!N%s+%s!P%s+%s!P%s)),0)," \
                        "IF(AND(%s!S%s-(%s!N%s+%s!N%s+%s!P%s+%s!P%s)>0,%s!L%s>12)," \
                        "IF(%s!S%s-(%s!N%s+%s!N%s+%s!P%s+%s!P%s)>%s!L%s-12,%s!L%s-12,%s!S%s-" \
                        "(%s!N%s+%s!N%s+%s!P%s+%s!P%s)),0))" \
                        % ("violations", str(i), "violations", str(i), "violations", str(i), "violations", str(i),
                           "violations", str(i), "violations", str(i + 1), "violations", str(i),
                           "violations", str(i + 1), "violations", str(i), "violations", str(i + 1),
                           "violations", str(i), "violations", str(i + 1), "violations", str(i),
                           "violations", str(i + 1), "violations", str(i), "violations", str(i + 1),
                           "violations", str(i + 1), "violations", str(i), "violations", str(i + 1),
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
            formula_f = "=IF(OR(%s!B%s=\"wal\",%s!B%s=\"nl\",%s!B%s=\"ptf\",%s!B%s=\"aux\")," \
                        "IF(AND(%s!S%s-(%s!P%s+%s!P%s)>0,%s!N%s>11.5)," \
                        "IF(%s!S%s-(%s!P%s+%s!P%s)>%s!N%s-11.5,%s!N%s-11.5,%s!S%s-(%s!P%s+%s!P%s)),0)," \
                        "IF(AND(%s!S%s-(%s!P%s+%s!P%s)>0,%s!N%s>12)," \
                        "IF(%s!S%s-(%s!P%s+%s!P%s)>%s!N%s-12,%s!N%s-12,%s!S%s-(%s!P%s+%s!P%s)),0))" \
                        % ("violations", str(i), "violations", str(i), "violations", str(i), "violations", str(i),
                           "violations", str(i), "violations", str(i + 1), "violations", str(i),
                           "violations", str(i + 1),
                           "violations", str(i), "violations", str(i + 1), "violations", str(i),
                           "violations", str(i + 1), "violations", str(i + 1), "violations", str(i),
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
            formula_g = "=IF(OR(%s!B%s=\"wal\",%s!B%s=\"nl\",%s!B%s=\"ptf\",%s!B%s=\"aux\")," \
                        "IF(AND(%s!S%s>0,%s!P%s>11.5)," \
                        "IF(%s!S%s>%s!P%s-11.5,%s!P%s-11.5,%s!S%s),0)," \
                        "IF(AND(%s!S%s>0,%s!P%s>12)," \
                        "IF(%s!S%s>%s!P%s-12,%s!P%s-12,%s!S%s),0))" \
                        % ("violations", str(i), "violations", str(i), "violations", str(i), "violations", str(i),
                           "violations", str(i), "violations", str(i + 1), "violations", str(i),
                           "violations", str(i + 1), "violations", str(i + 1), "violations", str(i),
                           "violations", str(i), "violations", str(i + 1), "violations", str(i),
                           "violations", str(i + 1), "violations", str(i + 1), "violations", str(i))
            self.violations['W' + str(i)] = formula_g
            self.violations['W' + str(i)].style = self.vert_calcs
            self.violations['W' + str(i)].number_format = "#,###.00"
            # total violation
            self.violations.merge_cells('X' + str(i) + ':X' + str(i + 1))  # merge box for total violation
            formula_h = "=SUM(%s!S%s:T%s)-(%s!U%s+%s!V%s+%s!W%s)" \
                        % ("violations", str(i), str(i), "violations", str(i),
                           "violations", str(i), "violations", str(i))
            self.violations['X' + str(i)] = formula_h
            self.violations['X' + str(i)].style = self.calcs
            self.violations['X' + str(i)].number_format = "#,###.00"
            # =IF($violations.A13 = 0,"",$violations.A13)
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

    def save_open(self):
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


class SpeedSheetGen:
    def __init__(self, frame, full_report):
        self.frame = frame
        self.full_report = full_report  # true - all inclusive, false - carrier recs only
        self.pb = ProgressBarDe()  # create the progress bar object
        self.db = SpeedSettings()  # calls values from tolerance table
        self.date = projvar.invran_date
        self.day_array = (str(projvar.invran_date.strftime("%a")).lower(),)  # if invran_weekly_span == False
        self.range = "day"
        if projvar.invran_weekly_span:  # if investigation range is weekly
            self.day_array = ("sat", "sun", "mon", "tue", "wed", "thu", "fri")
            self.range = "week"
        self.rotate_mode = self.db.speedcell_ns_rotate_mode  # NS day mode preference: True-rotating or False-fixed
        self.ns_pref = "r"  # "r" for rotating
        if not self.rotate_mode:
            self.ns_pref = "f"  # "f" for fixed
        self.dlsn_dict = {"sat": "sat", "mon": "mon", "tue": "tue", "wed": "wed", "thu": "thu", "fri": "fri",
                          "rsat": "sat", "rmon": "mon", "rtue": "tue", "rwed": "wed", "rthu": "thu", "rfri": "fri",
                          "fsat": "sat", "fmon": "mon", "ftue": "tue", "fwed": "wed", "fthu": "thu", "ffri": "fri",
                          "  ": "none", "": "none"}
        self.id_recset = []
        self.car_recs = []
        self.speedcell_count = 0
        self.ws_list = []
        self.ws_titles = []
        self.wb = Workbook()  # define the workbook
        self.ws_header = ""
        self.list_header = ""  # spreadsheet styles
        self.date_dov = ""
        self.date_dov_title = ""
        self.car_col_header = ""
        self.bold_name = ""
        self.input_name = ""
        self.col_header = ""
        self.input_s = ""
        self.input_ns = ""
        self.filename = ""

    def gen(self):
        self.get_id_recset()  # get carrier list and format for speedsheets
        self.get_car_recs()  # sort carrier list by worksheet
        self.speedcell_count = self.count()  # get a count of rows for progress bar
        self.make_workbook_object()  # generate and open the workbook
        self.name_styles()  # define the spreadsheet styles
        self.make_workbook()  # generate and open the workbook
        self.stopsaveopen()  # stop, save and open

    def get_id_recset(self):  # get filtered/ condensed record set and employee id
        carriers = CarrierList(projvar.invran_date_week[0], projvar.invran_date_week[6],
                               projvar.invran_station).get()  # first get a carrier list
        for c in carriers:
            # filter out any recs where list status is unchanged
            filtered_recs = CarrierRecFilter(c, projvar.invran_date_week[0]).filter_nonlist_recs()
            # condense multiple recs into format used by speedsheets
            condensed_recs = CarrierRecFilter(filtered_recs, projvar.invran_date_week[0]).condense_recs(
                self.db.speedcell_ns_rotate_mode)
            self.id_recset.append(self.add_id(condensed_recs))  # merge carriers with emp id

    @staticmethod
    def add_id(recset):  # put the employee id and carrier records together in a list
        carrier = recset[1]
        sql = "SELECT emp_id FROM name_index WHERE kb_name = '%s'" % carrier
        result = inquire(sql)
        if len(result) == 1:
            addthis = (result[0][0], recset)
        else:
            addthis = ("", recset)  # if there is no employee id, insert an empty string
        return addthis

    def get_car_recs(self):  # sort carrier records by the worksheets they will be put on
        self.car_recs = [self.order_by_id()]  # combine the id_rec arrays for emp id and alphabetical
        if not self.db.abc_breakdown:
            order_abc = self.order_alphabetically()  # sort the id_recset alphabetically
        else:
            order_abc = self.order_by_abc_breakdown()  # sort the id_recset w/o emp id by abc breakdown
        for abc in order_abc:
            self.car_recs.append(abc)

    def order_by_id(self):  # order id_recset by employee id
        ordered_recs = []
        for rec in self.id_recset:  # loop through the carrier list
            if rec[0] != "":  # if the item for employee id is not empty
                ordered_recs.append(rec)  # add the record set to the array
        ordered_recs.sort(key=itemgetter(0))  # sort the array by the employee id
        return ordered_recs

    def order_alphabetically(self):  # order id recset alphabetically into one tab
        alpha_recset = ["alpha_array", ]
        alpha_recset[0] = []
        for rec in self.id_recset:
            if rec[0] == "":
                alpha_recset[0].append(rec)
        return alpha_recset

    def order_by_abc_breakdown(self):  # sort id recset alphabetically into multiple tabs
        abc_recset = ["a_array", "b_array", "cd_array", "efg_array", "h_array", "ijk_array", "m_array",
                      "nop_array", "qr_array", "s_array", "tuv_array", "w_array", "xyz_array"]
        for i in range(len(abc_recset)):
            abc_recset[i] = []
        for rec in self.id_recset:
            if rec[0] == "":
                if rec[1][1][0] == "a":  # sort names without emp ids into lettered arrays
                    abc_recset[0].append(rec)
                elif rec[1][1][0] == "b":
                    abc_recset[1].append(rec)
                elif rec[1][1][0] == "rec" or rec[1][1][0] == "d":
                    abc_recset[2].append(rec)
                elif rec[1][1][0] == "e" or rec[1][1][0] == "f" or rec[1][1][0] == "g":
                    abc_recset[3].append(rec)
                elif rec[1][1][0] == "h":
                    abc_recset[4].append(rec)
                elif rec[1][1][0] == "i" or rec[1][1][0] == "j" or rec[1][1][0] == "k":
                    abc_recset[5].append(rec)
                elif rec[1][1][0] == "m":
                    abc_recset[6].append(rec)
                elif rec[1][1][0] == "n" or rec[1][1][0] == "o" or rec[1][1][0] == "p":
                    abc_recset[7].append(rec)
                elif rec[1][1][0] == "q" or rec[1][1][0] == "r":
                    abc_recset[8].append(rec)
                elif rec[1][1][0] == "s":
                    abc_recset[9].append(rec)
                elif rec[1][1][0] == "t" or rec[1][1][0] == "u" or rec[1][1][0] == "v":
                    abc_recset[10].append(rec)
                elif rec[1][1][0] == "w":
                    abc_recset[11].append(rec)
                else:
                    abc_recset[12].append(rec)
        return abc_recset

    def count_minrow_array(self):  # gets a count of minimum row info for each SpeedSheet tab
        minrow_array = [self.db.min_empid, ]  # get minimum rows for employee id sheet
        if not self.db.abc_breakdown:
            minrow_array.append(self.db.min_alpha)  # get minimum rows for alphabetical sheet
        else:
            for i in range(len(self.car_recs) - 1):  # get minimum rows for abc breakdown sheets
                minrow_array.append(self.db.min_abc)
        return minrow_array

    def count_car_recs(self):  # gets a count of carrier records for each SpeedSheet tab
        car_recs = [len(self.car_recs[0]), ]  # get count of carrier recs for employee id sheet
        if not self.db.abc_breakdown:
            car_recs.append(len(self.car_recs[1]))  # get count of carriers for alphabetical sheet
        else:
            for i in range(1, len(self.car_recs)):  # get count of carriers for abc breakdown
                car_recs.append(len(self.car_recs[i]))
        return car_recs

    def count(self):  # compare the minimum row and carrier records arrays to get the number of SpeedCells
        speedcell_count = 0  # initialized the count
        minrows = self.count_minrow_array()  # get minimum row count
        carrecs = self.count_car_recs()  # get count of carriers
        for i in range(len(minrows)):  # loop through results
            speedcell_count += max(minrows[i], carrecs[i])  # take the larger of the two
        return speedcell_count  # return total number of speedcells to be generated

    def mv_to_speed(self, triset):  # format mv triads for output to speedsheets
        if triset == "":
            return triset  # do nothing if blank
        else:
            return self.mv_format(triset)  # send to mv_format for formating if not blank

    @staticmethod
    def mv_format(triset):  # format mv triads for output to speedsheets
        mv_array = triset.split(",")  # split by commas
        mv_str = ""  # the move string
        i = 1  # initiate counter
        for mv in mv_array:
            mv_str += mv
            if i % 3 != 0:
                mv_str += "+"  # put + between items in move triads
            elif i % 3 == 0 and i != len(mv_array):
                mv_str += "/"  # put / between move triads
            else:
                mv_str += ""  # if at the end
            i += 1  # increment counter
        return mv_str

    def make_workbook_object(self):
        if not self.db.abc_breakdown:
            self.ws_list = ["emp_id", "alphabet"]
            self.ws_titles = ["by employee id", "alphabetically"]
        else:
            self.ws_list = ["emp_id", "a", "b", "cd", "efg", "h", "ijk", "m", "nop", "qr", "s", "tuv", "w", "xyz"]
            self.ws_titles = ["employee id", "a", "b", "c,d", "e,f,g", "h", "i,j,k",
                              "m", "n,o,p", "q,r,", "s", "t,u,v", "w", "x,y,z"]
        self.ws_list[0] = self.wb.active  # create first worksheet
        self.ws_list[0].title = self.ws_titles[0]  # title first worksheet
        self.ws_list[0].protection.sheet = True
        for i in range(1, len(self.ws_list)):  # loop to create all other worksheets
            self.ws_list[i] = self.wb.create_sheet(self.ws_titles[i])
            # self.ws_list[i].protection.sheet = True

    def title(self):  # generate title and filename
        if self.full_report and self.range == "week":
            title = "Speedsheet - All Inclusive Weekly"
            self.filename = "speed_" + str(format(projvar.invran_date_week[0], "%y_%m_%d")) + "_all_w" + ".xlsx"
        elif self.full_report and self.range == "day":
            title = "Speedsheet - All Inclusive Daily"
            self.filename = "speed_" + str(format(projvar.invran_date, "%y_%m_%d")) + "_all_d" + ".xlsx"
        else:
            title = "Speedsheet - Carriers"
            self.filename = "speed_" + str(format(projvar.invran_date_week[0], "%y_%m_%d")) + "_carrier" + ".xlsx"
        return title

    def name_styles(self):  # Named styles for workbook
        bd = Side(style='thin', color="80808080")  # defines borders
        self.ws_header = NamedStyle(name="ws_header", font=Font(bold=True, name='Arial', size=12))
        self.list_header = NamedStyle(name="list_header", font=Font(bold=True, name='Arial', size=10))
        self.date_dov = NamedStyle(name="date_dov", font=Font(name='Arial', size=8))
        self.date_dov_title = NamedStyle(name="date_dov_title", font=Font(bold=True, name='Arial', size=8),
                                         alignment=Alignment(horizontal='right'))
        # other color options: yellow: faf818, blue: 18fafa, green: 18fa20, grey: ababab
        if self.full_report:  # color carrier cells
            self.car_col_header = NamedStyle(name="car_col_header", font=Font(bold=True, name='Arial', size=8),
                                             fill=PatternFill(fgColor='18fafa', fill_type='solid'),
                                             border=Border(left=bd, top=bd, right=bd, bottom=bd),
                                             alignment=Alignment(horizontal='left'))
            self.bold_name = NamedStyle(name="bold_name", font=Font(name='Arial', size=8, bold=True),
                                        fill=PatternFill(fgColor='18fafa', fill_type='solid'),
                                        border=Border(left=bd, top=bd, right=bd, bottom=bd))
            self.input_name = NamedStyle(name="input_name", font=Font(name='Arial', size=8),
                                         fill=PatternFill(fgColor='18fafa', fill_type='solid'),
                                         border=Border(left=bd, top=bd, right=bd, bottom=bd))
        else:  # do not color carrier cells
            self.car_col_header = NamedStyle(name="car_col_header", font=Font(bold=True, name='Arial', size=8),
                                             border=Border(left=bd, top=bd, right=bd, bottom=bd),
                                             alignment=Alignment(horizontal='left'))
            self.bold_name = NamedStyle(name="bold_name", font=Font(name='Arial', size=8, bold=True),
                                        border=Border(left=bd, top=bd, right=bd, bottom=bd))
            self.input_name = NamedStyle(name="input_name", font=Font(name='Arial', size=8),
                                         border=Border(left=bd, top=bd, right=bd, bottom=bd))
        self.col_header = NamedStyle(name="col_header", font=Font(bold=True, name='Arial', size=8),
                                     border=Border(left=bd, top=bd, right=bd, bottom=bd),
                                     alignment=Alignment(horizontal='left'))

        self.input_s = NamedStyle(name="input_s", font=Font(name='Arial', size=8),
                                  border=Border(left=bd, top=bd, right=bd, bottom=bd),
                                  alignment=Alignment(horizontal='right'))
        self.input_ns = NamedStyle(name="input_ns", font=Font(bold=True, name='Arial', size=8, color='ff0000'),
                                   border=Border(left=bd, top=bd, right=bd, bottom=bd),
                                   alignment=Alignment(horizontal='right'))

    def make_workbook(self):
        pi = 0
        empty_sc = 0
        self.pb.max_count(self.speedcell_count)  # set length of progress bar
        self.pb.start_up()  # start the progress bar
        for i in range(len(self.ws_list)):
            # format cell widths
            self.ws_list[i].oddFooter.center.text = "&A"
            self.ws_list[i].column_dimensions["A"].width = 8
            self.ws_list[i].column_dimensions["B"].width = 8
            self.ws_list[i].column_dimensions["C"].width = 8
            self.ws_list[i].column_dimensions["D"].width = 8
            self.ws_list[i].column_dimensions["E"].width = 8
            self.ws_list[i].column_dimensions["F"].width = 8
            self.ws_list[i].column_dimensions["G"].width = 8
            self.ws_list[i].column_dimensions["H"].width = 8
            self.ws_list[i].column_dimensions["I"].width = 8
            self.ws_list[i].column_dimensions["J"].width = 8
            self.ws_list[i].column_dimensions["K"].width = 8
            cell = self.ws_list[i].cell(column=1, row=1)
            cell.value = self.title()
            cell.style = self.ws_header
            self.ws_list[i].merge_cells('A1:E1')
            # create date/ pay period/ station header
            cell = self.ws_list[i].cell(row=2, column=1)  # date label
            cell.value = "Date:  "
            cell.style = self.date_dov_title
            cell = self.ws_list[i].cell(row=2, column=2)  # date
            # if investigation range is daily
            cell.value = "{}".format(projvar.invran_date.strftime("%m/%d/%Y"))
            if projvar.invran_weekly_span:
                cell.value = "{} through {}".format(projvar.invran_date_week[0].strftime("%m/%d/%Y"),
                                                    projvar.invran_date_week[6].strftime("%m/%d/%Y"))
            cell.style = self.date_dov
            self.ws_list[i].merge_cells('B2:E2')
            cell = self.ws_list[i].cell(row=2, column=6)  # pay period label
            cell.value = "PP:  "
            cell.style = self.date_dov_title
            cell = self.ws_list[i].cell(row=2, column=7)  # pay period
            cell.value = projvar.pay_period
            cell.style = self.date_dov
            cell = self.ws_list[i].cell(row=2, column=8)  # station label
            cell.value = "Station:  "
            cell.style = self.date_dov_title
            cell = self.ws_list[i].cell(row=2, column=9)  # station
            cell.value = projvar.invran_station
            cell.style = self.date_dov
            self.ws_list[i].merge_cells('I2:J2')
            # apply title - show how carriers are sorted
            cell = self.ws_list[i].cell(row=3, column=1)
            if i == 0:
                cell.value = "Carriers listed by Employee ID"
            else:
                cell.value = "Carriers listed Alphabetically: {}".format(self.ws_titles[i])
            cell.style = self.list_header
            self.ws_list[i].merge_cells('A3:E3')
            if i == 0:  # only execute on the first sheet of the workbook
                cell = self.ws_list[i].cell(row=3, column=6)  #
                cell.value = "ns day preference (r=rotating/f=fixed): "  # ns day preference
                cell.style = self.date_dov_title
                self.ws_list[i].merge_cells('F3:I3')
                cell = self.ws_list[i].cell(row=3, column=10)  #
                cell.value = self.ns_pref
                cell.style = self.date_dov
            # Headers for Carrier List
            cell = self.ws_list[i].cell(row=4, column=1)  # header day
            cell.value = "Days"
            cell.style = self.car_col_header
            cell = self.ws_list[i].cell(row=4, column=2)  # header carrier name
            cell.value = "Carrier Name"
            cell.style = self.car_col_header
            self.ws_list[i].merge_cells('B4:D4')
            cell = self.ws_list[i].cell(row=4, column=5)  # header list type
            cell.value = "List"
            cell.style = self.car_col_header
            cell = self.ws_list[i].cell(row=4, column=6)  # header ns day
            cell.value = "NS Day"
            cell.style = self.car_col_header
            cell = self.ws_list[i].cell(row=4, column=7)  # header route
            cell.value = "Route/s"
            cell.style = self.car_col_header
            self.ws_list[i].merge_cells('G4:I4')
            cell = self.ws_list[i].cell(row=4, column=10)  # header emp id
            cell.value = "Emp id"
            cell.style = self.car_col_header
            row = 5  # start at row 5 after the page header display
            if self.full_report:  # only include rings headers on all inclusive not carrier only
                # Headers for Rings
                cell = self.ws_list[i].cell(row=5, column=1)  # header day
                cell.value = "Day"
                cell.style = self.col_header
                cell = self.ws_list[i].cell(row=5, column=2)  # header 5200
                cell.value = "5200"
                cell.style = self.col_header
                cell = self.ws_list[i].cell(row=5, column=3)  # header MOVES
                cell.value = "MOVES"
                cell.style = self.col_header
                self.ws_list[i].merge_cells('C5:F5')
                cell = self.ws_list[i].cell(row=5, column=7)  # header RS
                cell.value = "RS"
                cell.style = self.col_header
                cell = self.ws_list[i].cell(row=5, column=8)  # header codes
                cell.value = "CODE"
                cell.style = self.col_header
                cell = self.ws_list[i].cell(row=5, column=9)  # header leave type
                cell.value = "LV type"
                cell.style = self.col_header
                cell = self.ws_list[i].cell(row=5, column=10)  # header leave time
                cell.value = "LV time"
                cell.style = self.col_header
                row = 6  # update start at row 6 after the page header display if all inclusive
            # freeze panes
            self.ws_list[i].freeze_panes = self.ws_list[i].cell(row=row, column=1)  # ['A5] or ['A6']
            if i == 0:
                rowcount = self.db.min_empid  # get minimum speedcell count for employee id tab
            elif i != 0 and not self.db.abc_breakdown:
                rowcount = self.db.min_alpha  # get minimum speedcell count for alphabetical tab
            else:
                rowcount = self.db.min_abc  # get minimum speedcell count for each abc breakdown tab
            rowcounter = max(rowcount, len(self.car_recs[i]))
            for r in range(rowcounter):
                if r < len(self.car_recs[i]):  # if the carrier records are not exhausted
                    eff_date = self.car_recs[i][r][1][0]  # carrier effective date
                    car_name = self.car_recs[i][r][1][1]  # carrier name
                    car_list = self.car_recs[i][r][1][2]  # carrier list status
                    car_ns = self.car_recs[i][r][1][3]  # carrier ns day
                    car_route = self.car_recs[i][r][1][4]  # carrier route
                    car_empid = self.car_recs[i][r][0]  # carrier employee id number
                    self.pb.move_count(pi)  # increment progress bar
                    self.pb.change_text("Formatting Speedcell for {}".format(car_name))
                else:
                    eff_date = ""  # enter blanks once records are exhausted
                    car_name = ""
                    car_list = ""
                    car_ns = ""
                    car_route = ""
                    car_empid = ""
                    self.pb.move_count(pi)  # increment progress bar
                    empty_sc += 1  # increment counter for empty speedcells
                    self.pb.change_text("Formatting empty Speedcell #{}".format(empty_sc))
                pi += 1  # progress bar counter
                cell = self.ws_list[i].cell(row=row, column=1)  # carrier effective date
                cell.value = eff_date
                cell.style = self.input_name
                cell.protection = Protection(locked=False)
                cell = self.ws_list[i].cell(row=row, column=2)  # carrier name
                cell.value = car_name
                cell.style = self.bold_name
                cell.protection = Protection(locked=False)
                self.ws_list[i].merge_cells('B' + str(row) + ':' + 'D' + str(row))
                cell = self.ws_list[i].cell(row=row, column=5)  # carrier list status
                cell.value = car_list
                cell.style = self.input_name
                cell.protection = Protection(locked=False)
                cell = self.ws_list[i].cell(row=row, column=6)  # carrier ns day
                cell.value = car_ns
                cell.style = self.input_name
                cell.protection = Protection(locked=False)
                cell = self.ws_list[i].cell(column=7, row=row)  # carrier route
                cell.value = car_route
                cell.style = self.input_name
                cell.protection = Protection(locked=False)
                self.ws_list[i].merge_cells('G' + str(row) + ':' + 'I' + str(row))
                cell = self.ws_list[i].cell(column=10, row=row)  # carrier emp id
                cell.value = car_empid
                cell.style = self.input_name
                cell.protection = Protection(locked=False)
                row += 1
                ring_recs = []
                if self.full_report:
                    if r < len(self.car_recs[i]):  # if the carrier records are not exhausted
                        if self.range == "day":  # if the investigation range is for a day
                            ring_recs = Rings(car_name, self.date).get_for_day()  # get the rings for the carrier
                        else:  # if the investigation range is for a week
                            ring_recs = Rings(car_name, self.date).get_for_week()  # get the rings for the carrier
                    for d in range(len(self.day_array)):
                        if r < len(self.car_recs[i]) and ring_recs[d] != []:  # if the carrier records are not exhausted
                            ring_5200 = ring_recs[d][2]  # rings 5200
                            ring_move = self.mv_to_speed(ring_recs[d][5])  # format rings MOVES
                            ring_rs = ring_recs[d][3]  # rings RS
                            ring_code = ring_recs[d][4]  # rings CODES
                            ring_lvty = ring_recs[d][6]  # rings LEAVE TYPE
                            ring_lvtm = ring_recs[d][7]  # rings LEAVE TIME
                        else:
                            ring_5200 = ""  # rings 5200
                            ring_move = ""  # rings MOVES
                            ring_rs = ""  # rings RS
                            ring_code = ""  # rings CODES
                            ring_lvty = ""  # rings LEAVE TYPE
                            ring_lvtm = ""  # rings LEAVE TIME
                        cell = self.ws_list[i].cell(column=1, row=row)  # rings day
                        cell.value = self.day_array[d]
                        if self.day_array[d] == self.dlsn_dict[car_ns.lower()]:  # if it is the nsday
                            cell.style = self.input_ns  # display it red and bold
                        else:
                            cell.style = self.input_s
                        cell = self.ws_list[i].cell(column=2, row=row)  # rings 5200
                        cell.value = ring_5200
                        cell.style = self.input_s
                        cell.protection = Protection(locked=False)
                        cell = self.ws_list[i].cell(column=3, row=row)  # rings moves
                        cell.value = ring_move
                        cell.style = self.input_s
                        cell.protection = Protection(locked=False)
                        self.ws_list[i].merge_cells('C' + str(row) + ':' + 'F' + str(row))
                        cell = self.ws_list[i].cell(column=7, row=row)  # rings RS
                        cell.value = ring_rs
                        cell.style = self.input_s
                        cell.protection = Protection(locked=False)
                        cell = self.ws_list[i].cell(column=8, row=row)  # rings code
                        cell.value = ring_code
                        cell.style = self.input_s
                        cell.protection = Protection(locked=False)
                        cell = self.ws_list[i].cell(column=9, row=row)  # rings lv type
                        cell.value = ring_lvty
                        cell.style = self.input_s
                        cell.protection = Protection(locked=False)
                        cell = self.ws_list[i].cell(column=10, row=row)  # rings lv time
                        cell.value = ring_lvtm
                        cell.style = self.input_s
                        cell.protection = Protection(locked=False)
                        row += 1
        self.pb.stop()

    def stopsaveopen(self):
        try:
            self.wb.save(dir_path('speedsheets') + self.filename)
            messagebox.showinfo("Speedsheet Generator",
                                "Your speedsheet was successfully generated. \n"
                                "File is named: {}".format(self.filename),
                                parent=self.frame)
            if sys.platform == "win32":
                os.startfile(dir_path('speedsheets') + self.filename)
            if sys.platform == "linux":
                subprocess.call(["xdg-open", 'kb_sub/speedsheets/' + self.filename])
            if sys.platform == "darwin":
                subprocess.call(["open", dir_path('speedsheets') + self.filename])
        except PermissionError:
            messagebox.showerror("Speedsheet generator",
                                 "The speedsheet was not generated. \n"
                                 "Suggestion: \n"
                                 "Make sure that identically named speedsheets are closed \n"
                                 "(the file can't be overwritten while open).\n",
                                 parent=self.frame)
