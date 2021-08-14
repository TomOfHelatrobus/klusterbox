# custom modules
import projvar  # holds variables, including root, for use in all modules
from kbtoolbox import *
from tkinter import messagebox
import os
# Spreadsheet Libraries
from openpyxl import load_workbook
from openpyxl import Workbook
from openpyxl.styles import NamedStyle, Font, Border, Side, Alignment, PatternFill, Protection
from openpyxl.worksheet.pagebreak import Break


class QuarterRecs:
    def __init__(self, carrier, startdate, enddate, station):
        self.carrier = carrier
        self.start = startdate
        self.end = enddate
        self.station = station

    def get_filtered_recs(self, ls):
        lst = ("otdl", )
        if ls == "non":
            lst = ("nl", "wal")
        if ls == "aux":
            lst = ("aux", )
        if ls == "ptf":
            lst = ("ptf", )
        recset = self.get_recs()
        if recset is None:
            return
        for rec in recset:
            if rec[2] in lst:
                return recset

    def get_recs(self):
        # get all records in the investigation range - the entire quarter
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


class OTEquitSpreadsheet:
    def __init__(self):
        self.frame = None
        self.date = None
        self.station = None
        self.year = None
        self.month = None
        self.startdate = None
        self.enddate = None
        self.quarter = None
        self.full_quarter = None
        self.startdate_index = []
        self.enddate_index = []
        self.carrierlist = []
        self.recset = []
        self.minrows = 1
        self.otcalcpref = "off_route"  # preference for overtime calculation - "off_route" or "all"
        self.carrier_overview = []  # a list of carrier's name, status and makeups
        self.date_array = []  # a list of all days in the quarter as a datetimes
        self.front_padding = 0  # number of empty triad to fill worksheet prior to startdate
        self.end_padding = 0  # number of empty trids to fill worksheet prior to enddate
        self.ringrefset = []  # multidimensional array - daily rings/refusals for each otdl carrier
        self.dates_breakdown = []  # a list of dates for display on spreadsheets
        self.week = ("w01", "w02", "w03", "w04", "w05", "w06", "w07", "w08", "w09", "w10", "w11", "w12",
                     "w13", "w14", "w15")
        self.week_label = ("01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12",
                     "13", "14", "15")
        self.triad_week_index = 0
        self.triad_column = 0
        self.triad_row = 0
        self.footer_row = 0
        self.wb = None  # workbook
        self.overview = None  # first worksheet which summarizes all the following worksheets
        self.ws = None  # worksheet
        self.instructions = None  # last worksheet which provides instructions
        self.ws_header = None  # NamedStyles for spreadsheet
        self.date_dov = None
        self.date_dov_title = None
        self.col_header = None
        self.col_center_header = None
        self.col_header_instructions = None
        self.input_name = None
        self.input_s = None
        self.input_center = None
        self.calcs = None
        self.ws_name = None
        self.ref_ot = None
        self.instruct_text = None

    def create(self, frame, date, station):
        self.frame = frame
        if not self.ask_ok():  # abort if user selects cancel from askokcancel
            return
        self.station = station
        self.date = date  # a datetime object from the quarter is passed and used as date
        self.breakdown_date()  # the passed datetime object is broken down into year and month
        self.define_quarter()  # the year and month are used to generate quarter and full quarter
        self.get_dates()  # use quarter information to get start and end date
        self.get_carrierlist()  # generate a raw list of carriers at station before or on end date.
        self.get_recsets()  # filter the carrierlist to get only otdl carriers
        self.get_settings()  # get minimum rows and ot calculation preference
        self.get_carier_overview()  # build a list of carrier's name, status and makeups
        self.carrier_overview_add()  # adds empty sets so the lenght of carrier overview = minimum rows.
        self.get_date_array()  # get a list of all days in the quarter as datetime objects
        self.get_front_padding()  # get number of empty triads needed to pad prior to startdate
        self.get_end_padding()  # get number of empty triads needed to pad after enddate
        self.get_ringrefset()  # build multidimensional array - daily rings/refusals for each carrier
        self.get_date_breakdown()  # build an array of dates for display on the spreadsheets
        self.get_footer_row()  # get the row where the footer will be
        self.build_workbook()  # build the spreadsheet and define the worksheets.
        self.set_dimensions_overview()  # column widths for overview sheet
        self.set_dimensions_weekly()  # column widths for weekly worksheets
        self.get_styles()  # define workbook styles
        self.build_header_overview()  # build the header for overview worksheet
        self.build_columnheader_overview()  # build column headers for overview worksheet
        self.build_main_overview()  # build main body for overview worksheet
        self.build_overview_footer()  # create the footer at the bottom of the worksheet for averages and totals
        self.build_header_weeklysheets()  # build the header for the weekly worksheet
        self.build_columnheader_worksheets()  # build column headers for weekly worksheet
        self.build_main_worksheets()  # build the main parts of the worksheet save the triad groups
        self.triads_delegator()  # orders which triads to build from build triads
        self.build_worksheet_footer()
        self.build_instructions()
        self.save_open()  # save and open the spreadsheet

    def ask_ok(self):
        if messagebox.askokcancel("Spreadsheet generator",
                                  "Do you want to generate a spreadsheet?",
                                  parent=self.frame):
            return True
        return False

    def breakdown_date(self):  # breakdown the date into year and month
        self.year = int(self.date.strftime("%Y"))
        self.month = int(self.date.strftime("%m"))

    def define_quarter(self):
        self.quarter = Quarter(self.month).find()  # convert the month into a quarter - 1 through 4.
        self.full_quarter = str(self.year) + "-" + str(self.quarter)  # create a string expressing the year - quarter

    def get_dates(self):
        self.startdate_index = (datetime(self.year, 1, 1), datetime(self.year, 4, 1), datetime(self.year, 7, 1),
                     datetime(self.year, 10, 1))
        self.enddate_index = (datetime(self.year, 3, 31), datetime(self.year, 6, 30), datetime(self.year, 9, 30),
                   datetime(self.year, 12, 31))
        self.startdate = self.startdate_index[int(self.quarter) - 1]
        self.enddate = self.enddate_index[int(self.quarter) - 1]

    def starting_day(self):  # returns the column position of the startdate as an odd number (5 to 17)
        days = ("Saturday", "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday")
        i = 6  #
        for day in days:  # loop through tuple of days"

            if self.startdate.strftime("%A") == day:  # if the startdate matches the day
                return i  # returns the column of the first date
            i -= 1  # count down from Saturday

    def get_carrierlist(self):
        self.carrierlist = CarrierList(self.startdate, self.enddate, self.station).get_distinct()

    def get_recsets(self):
        for carrier in self.carrierlist:
            rec = QuarterRecs(carrier[0], self.startdate, self.enddate, self.station).get_filtered_recs("otdl")
            if rec:
                self.recset.append(rec)

    def get_settings(self):  # get minimum rows and ot calculation preference
        sql = "SELECT tolerance FROM tolerances"
        result = inquire(sql)
        self.minrows = int(result[25][0])
        self.otcalcpref = result[26][0]

    def get_status(self, recs):  # returns true if the carrier's last record is otdl and the station is correct.
        if recs[0][2] == "otdl" and recs[0][5] == self.station:
            return True
        return False

    def get_pref(self, recs):
        carrier = recs[0][1]
        sql = "SELECT preference FROM otdl_preference WHERE carrier_name = '%s' and quarter = '%s' and station = '%s'" \
              % (carrier, self.full_quarter, self.station)
        pref = inquire(sql)
        if not pref:  # if there is no record in the database, create one
            sql = "INSERT INTO otdl_preference (quarter, carrier_name, preference, station, makeups) " \
                  "VALUES('%s', '%s', '%s', '%s', '%s')" \
                  % (self.full_quarter, carrier, "12", self.station, "")
            commit(sql)  # enter the new record in the dbae
            return "12"  # return 12 hour preference
        else:
            return pref[0][0]  # return the pulled from the database.

    def get_makeups(self, carrier):
        sql = "SELECT makeups FROM otdl_preference WHERE carrier_name = '%s' and quarter = '%s' and station = '%s'" \
              % (carrier, self.full_quarter, self.station)
        makeup = inquire(sql)
        if makeup:
            return makeup[0][0]
        else:
            return ""

    def get_carier_overview(self):  # build a list of carrier's name, status and makeups
        for recs in self.recset:  # loop through the recsets
            carrier = recs[0][1]  # get the carrier name
            status = "off"  # default status is off
            if self.get_status(recs):  # if the carrier is currently on the otdl
                status = self.get_pref(recs)  # pull the otdl preference from the database
            makeup = self.get_makeups(carrier)
            add_this = (carrier, status, makeup)
            self.carrier_overview.append(add_this)

    def carrier_overview_add(self):  # adds empty sets so the lenght of carrier overview = minimum rows.
        while len(self.carrier_overview) < self.minrows:
            add_this = ("", "", "")
            self.carrier_overview.append(add_this)

    def get_date_array(self):  # get a list of all days in the quarter as datetime objects
        running_date = self.startdate
        while running_date <= self.enddate:
            self.date_array.append(running_date)
            running_date += timedelta(days=1)

    def get_front_padding(self):  # get the number of empty triads to put before startdate to fill worksheet
        self.front_padding = 6 - self.starting_day()

    def get_end_padding(self):  # get number of empty triads needed after enddate to fill worksheet
        self.end_padding = 105 - (self.front_padding + len(self.date_array))

    def get_ringrefset(self):    # build multidimensional array - daily rings/refusals for each otdl carrier
        for i in range(len(self.carrier_overview)):
            self.ringrefset.append([])  # each carrier has an array
            self.get_daily_ringrefs(i)

    def get_overtime(self, total, moves, code):  # find the overtime pending ot calculation preference and ns day code
        if self.otcalcpref == "off_route":
            return Overtime().proper_overtime(total, moves, code)
        else:
            return Overtime().straight_overtime(total, code)

    def get_daily_ringrefs(self, index):
        daily_ringref = []
        carrier = self.carrier_overview[index][0]  # get the carrier name using carrier overview md array and index
        for _ in range(self.front_padding):  # insert front padding so empty cells fill worksheet
            add_this = ["", "", ""]
            daily_ringref.append(add_this)
        for date in self.date_array:  # get the ringrefs from the database or empty if none
            overtime = ""
            sql = "SELECT total, code, moves FROM rings3 WHERE rings_date = '%s' AND carrier_name = '%s'" \
                  % (date, carrier)
            results = inquire(sql)
            if results:
                total = results[0][0]
                code = results[0][1]
                moves = Moves().timeoffroute(results[0][2])  # calculate the time off route
                overtime = self.get_overtime(total, moves, code)  # find the overtime
            ref_type = ""
            ref_time = ""
            sql = "SELECT refusal_type, refusal_time FROM refusals WHERE refusal_date = '%s' AND carrier_name = '%s'" \
                  % (date, carrier)
            ref_results = inquire(sql)
            if ref_results:
                ref_type = ref_results[0][0]
                ref_time = ref_results[0][1]
            add_this = [overtime, ref_type, ref_time]
            daily_ringref.append(add_this)
        for _ in range(self.end_padding):  # insert front padding so empty cells fill worksheet
            add_this = ["", "", ""]
            daily_ringref.append(add_this)
        self.ringrefset[index] = daily_ringref

    def get_date_breakdown(self):
        date = self.startdate
        days_til_end = self.starting_day()
        if not days_til_end:
            display_date = date.strftime("%m/%d/%Y")
        else:
            enddate = date + timedelta(days=days_til_end)
            display_date = date.strftime("%m/%d/%Y") + " through " + enddate.strftime("%m/%d/%Y")
        self.dates_breakdown.append(display_date)
        date += timedelta(days=days_til_end + 1)
        for _ in range(14):  # loop once for each week
            enddate = date + timedelta(days=6)
            display_date = date.strftime("%m/%d/%Y") + " through " + enddate.strftime("%m/%d/%Y")
            self.dates_breakdown.append(display_date)
            date += timedelta(weeks=1)

    def get_footer_row(self):  # get the number of the row where the footer will go.
        self.footer_row = (len(self.carrier_overview) * 2) + 7

    def build_workbook(self):
        self.ws = []
        self.wb = Workbook()  # define the workbook
        self.overview = self.wb.active  # create first worksheet
        self.overview.title = "overview"  # give the first worksheet a name
        for i in range(15):  # create worksheet for remaining 15 weeks
            self.ws.append(self.wb.create_sheet(self.week[i]))  # create subsequent worksheets
            self.ws[i].title = self.week[i]  # title subsequent worksheets
        self.instructions = self.wb.create_sheet("instructions")

    def set_dimensions_overview(self):
        self.overview.column_dimensions["A"].width = 6
        self.overview.column_dimensions["B"].width = 12
        self.overview.column_dimensions["C"].width = 6
        self.overview.column_dimensions["D"].width = 7
        self.overview.column_dimensions["E"].width = 3
        self.overview.column_dimensions["F"].width = 11
        self.overview.column_dimensions["G"].width = 12
        self.overview.column_dimensions["H"].width = 12

    def set_dimensions_weekly(self):
        for i in range(15):
            self.ws[i].column_dimensions["A"].width = 6
            self.ws[i].column_dimensions["B"].width = 14
            self.ws[i].column_dimensions["C"].width = 5
            self.ws[i].column_dimensions["D"].width = 3
            self.ws[i].column_dimensions["E"].width = 2
            self.ws[i].column_dimensions["F"].width = 4
            self.ws[i].column_dimensions["G"].width = 2
            self.ws[i].column_dimensions["H"].width = 4
            self.ws[i].column_dimensions["I"].width = 2
            self.ws[i].column_dimensions["J"].width = 4
            self.ws[i].column_dimensions["K"].width = 2
            self.ws[i].column_dimensions["L"].width = 4
            self.ws[i].column_dimensions["M"].width = 2
            self.ws[i].column_dimensions["N"].width = 4
            self.ws[i].column_dimensions["O"].width = 2
            self.ws[i].column_dimensions["P"].width = 4
            self.ws[i].column_dimensions["Q"].width = 2
            self.ws[i].column_dimensions["R"].width = 4
            self.ws[i].column_dimensions["S"].width = 7

    def get_styles(self):  # Named styles for workbook
        bd = Side(style='thin', color="80808080")  # defines borders
        self.ws_header = NamedStyle(name="ws_header", font=Font(bold=True, name='Arial', size=12))
        self.date_dov = NamedStyle(name="date_dov", font=Font(name='Arial', size=8))
        self.date_dov_title = NamedStyle(name="date_dov_title", font=Font(bold=True, name='Arial', size=8),
                                    alignment=Alignment(horizontal='right'))
        self.col_header = NamedStyle(name="col_header", font=Font(bold=True, name='Arial', size=8))
        self.col_center_header = NamedStyle(name="col_center_header", font=Font(bold=True, name='Arial', size=8),
                                       alignment=Alignment(horizontal='center'))
        self.col_header_instructions = \
            NamedStyle(name="col_header_instructions", font=Font(bold=True, name='Arial', size=10))
        self.input_name = NamedStyle(name="input_name", font=Font(name='Arial', size=8),
                                border=Border(left=bd, right=bd, top=bd, bottom=bd),
                                alignment=Alignment(horizontal='left'))
        self.input_s = NamedStyle(name="input_s", font=Font(name='Arial', size=8),
                             border=Border(left=bd, right=bd, top=bd, bottom=bd),
                             alignment=Alignment(horizontal='right'))
        self.input_center = NamedStyle(name="input_center", font=Font(name='Arial', size=8),
                                  border=Border(left=bd, right=bd, top=bd, bottom=bd),
                                  alignment=Alignment(horizontal='center'))
        self.calcs = NamedStyle(name="calcs", font=Font(name='Arial', size=8),
                           border=Border(left=bd, right=bd, top=bd, bottom=bd),
                           fill=PatternFill(fgColor='e5e4e2', fill_type='solid'),
                           alignment=Alignment(horizontal='right'))
        self.ws_name = NamedStyle(name="ws_name", font=Font(name='Arial', size=8),
                                border=Border(left=bd, right=bd, top=bd, bottom=bd),
                                fill=PatternFill(fgColor='e5e4e2', fill_type='solid'),
                                alignment=Alignment(horizontal='left'))
        self.ref_ot = NamedStyle(name="ref_ot", font=Font(name='Arial', size=8),
                                border=Border(left=bd, right=bd, top=bd, bottom=bd),
                                fill=PatternFill(fgColor='e5e4e2', fill_type='solid'),
                                alignment=Alignment(horizontal='center'))
        self.instruct_text = NamedStyle(name="instruct_text", font=Font(name='Arial', size=10),
                                   alignment=Alignment(horizontal='left', vertical='top'))

    def build_header_overview(self):  # build the header for overview worksheet
        cell = self.overview.cell(row=1, column=1)  # page title
        cell.value = "OTDL Equitability Worksheet"
        cell.style = self.ws_header
        self.overview.merge_cells('A1:E1')
        cell = self.overview.cell(row=2, column=1)  # date
        cell.value = "dates: "
        cell.style = self.date_dov_title
        cell = self.overview.cell(row=2, column=2)  # fill in dates
        date = self.startdate_index[self.quarter - 1].strftime("%m/%d/%Y") + " through " + \
               self.enddate_index[self.quarter - 1].strftime("%m/%d/%Y")
        cell.value = date
        cell.style = self.date_dov
        self.overview.merge_cells('B2:E2')
        cell = self.overview.cell(row=2, column=6)  # station
        cell.value = "station: "
        cell.style = self.date_dov_title
        cell = self.overview.cell(row=2, column=7)  # fill in station
        cell.value = self.station
        cell.style = self.date_dov
        self.overview.merge_cells('G2:H2')
        cell = self.overview.cell(row=3, column=6)  # number of carriers
        cell.value = "# of carriers active on otdl: "
        cell.style = self.date_dov_title
        self.overview.merge_cells('F3:G3')
        cell = self.overview.cell(row=3, column=8)  # fill in number of carriers
        lastnum = ((len(self.carrier_overview)*2)+5)
        formula = "=COUNTIF(%s!%s%s:%s!%s%s, %s)+COUNTIF(%s!%s%s:%s!%s%s, %s)" \
                % ("overview", "C", str(6), "overview", "C", str(lastnum), str(12),
                   "overview", "C", str(6), "overview", "C", str(lastnum), str(10))
        cell.value = formula
        cell.style = self.calcs

    def build_columnheader_overview(self):
        cell = self.overview.cell(row=5, column=1)  # name
        cell.value = "name"
        cell.style = self.col_header
        self.overview.merge_cells('A5:B5')
        cell = self.overview.cell(row=5, column=3)  # status
        cell.value = "status"
        cell.style = self.col_center_header
        cell = self.overview.cell(row=5, column=4)  # make up
        cell.value = "make up"
        cell.style = self.col_center_header
        cell = self.overview.cell(row=5, column=5)  # refusals/overtime
        cell.value = "refusals/overtime"
        cell.style = self.col_center_header
        self.overview.merge_cells('E5:F5')
        cell = self.overview.cell(row=5, column=7)  # opportunities
        cell.value = "opportunities"
        cell.style = self.col_center_header
        cell = self.overview.cell(row=5, column=8)  # diff from avg
        cell.value = "diff from avg"
        cell.style = self.col_center_header

    def build_main_overview(self):
        row = 6
        for i in range(len(self.carrier_overview)):
            self.overview.row_dimensions[row].height = 13
            self.overview.row_dimensions[row+1].height = 13
            cell = self.overview.cell(row=row, column=1)  # name
            cell.value = self.carrier_overview[i][0]
            cell.style = self.input_name
            self.overview.merge_cells('A' + str(row) + ':' + 'B' + str(row+1))
            cell = self.overview.cell(row=row, column=3)  # status
            cell.value = Handler(self.carrier_overview[i][1]).str_to_int_or_str()
            cell.style = self.input_center
            self.overview.merge_cells('C' + str(row) + ':' + 'C' + str(row + 1))
            cell = self.overview.cell(row=row, column=4)  # make up
            cell.value = Handler(self.carrier_overview[i][2]).str_to_float_or_str()
            cell.style = self.input_center
            cell.number_format = "#,###.00;[RED]-#,###.00"
            self.overview.merge_cells('D' + str(row) + ':' + 'D' + str(row + 1))
            cell = self.overview.cell(row=row, column=5)  # refusals label
            cell.value = "ref"
            cell.style = self.ref_ot
            cell = self.overview.cell(row=row+1, column=5)  # OT label
            cell.value = "ot"
            cell.style = self.ref_ot
            cell = self.overview.cell(row=row, column=6)  # refusals
            formula = "=IF(OR(C%s=12, C%s=10)," \
                      "SUM(%s!S%s, %s!S%s, %s!S%s, %s!S%s, %s!S%s, " \
                      "%s!S%s, %s!S%s, %s!S%s, %s!S%s, %s!S%s, " \
                      "%s!S%s, %s!S%s, %s!S%s, %s!S%s, %s!S%s)," \
                      "0" \
                      % (str(row), str(row),
                         self.week[0], str(row), self.week[1], str(row), self.week[2], str(row),
                         self.week[3], str(row), self.week[4], str(row), self.week[5], str(row),
                         self.week[6], str(row), self.week[7], str(row), self.week[8], str(row),
                         self.week[9], str(row), self.week[10], str(row), self.week[11], str(row),
                         self.week[12], str(row), self.week[13], str(row), self.week[14], str(row))
            cell.value = formula
            cell.style = self.calcs
            cell.number_format = "#,###.00;[RED]-#,###.00"
            cell = self.overview.cell(row=row + 1, column=6)  # overtime
            formula = "=IF(OR(C%s=12, C%s=10)," \
                      "SUM(%s!S%s, %s!S%s, %s!S%s, %s!S%s, %s!S%s, " \
                      "%s!S%s, %s!S%s, %s!S%s, %s!S%s, %s!S%s, " \
                      "%s!S%s, %s!S%s, %s!S%s, %s!S%s, %s!S%s)," \
                      "0" \
                      % (str(row), str(row),
                         self.week[0], str(row+1), self.week[1], str(row+1), self.week[2], str(row+1),
                         self.week[3], str(row+1), self.week[4], str(row+1), self.week[5], str(row+1),
                         self.week[6], str(row+1), self.week[7], str(row+1), self.week[8], str(row+1),
                         self.week[9], str(row+1), self.week[10], str(row+1), self.week[11], str(row+1),
                         self.week[12], str(row+1), self.week[13], str(row+1), self.week[14], str(row+1))
            cell.value = formula
            cell.style = self.calcs
            cell.number_format = "#,###.00;[RED]-#,###.00"
            cell = self.overview.cell(row=row, column=7)  # opportunities
            formula = "=IF(OR(C%s=12, C%s=10),(F%s+F%s)-D%s,0)" % \
                      (str(row), str(row), str(row), str(row+1), str(row))
            cell.value = formula
            cell.style = self.calcs
            cell.number_format = "#,###.00;[RED]-#,###.00"
            self.overview.merge_cells('G' + str(row) + ":" + 'G' + str(row+1))
            cell = self.overview.cell(row=row, column=8)  # difference from average
            formula = "=IF(A%s=\"\",0, IF(OR(C%s=12,C%s=10),G%s-$G$%s,\"off list\"))" \
                      % (str(row), str(row), str(row), str(row), str(self.footer_row+2))  # last row is avg footer
            cell.value = formula
            cell.style = self.calcs
            cell.number_format = "#,###.00;[RED]-#,###.00"
            self.overview.merge_cells('H' + str(row) + ":" + 'H' + str(row + 1))
            row += 2

    def get_totalovertime_formula(self, row, column):  # gives formulas for totals counting skipping rows.
        """
        the row argument is the starting row of the count,
        the column is given as a number and matched to a letter with the dictionary
        """
        column_dict = {5: "E", 6: "F", 7: "G", 9: "I", 11: "K", 13: "M", 15: "O", 17: "Q"}
        string = "=SUM("
        while row < self.footer_row-2:
            string += "{}{},".format(column_dict[column], row)
            row += 2
        string += "{}{})".format(column_dict[column], row)
        return string

    def build_overview_footer(self):  # create the footer at the bottom of the worksheet for averages and totals
        cell = self.overview.cell(row=self.footer_row, column=5)  # label total overtime
        cell.value = "total overtime:"
        cell.style = self.date_dov_title
        cell = self.overview.cell(row=self.footer_row, column=6)  # calculate total overtime
        formula = self.get_totalovertime_formula(7, 6)
        cell.value = formula
        cell.style = self.calcs
        cell.number_format = "#,###.00;[RED]-#,###.00"
        cell = self.overview.cell(row=self.footer_row, column=7)  # calculate total opportunities
        formula = "=SUM(G%s:G%s)" % (str(6), str(self.footer_row - 2))
        cell.value = formula
        cell.style = self.calcs
        cell.number_format = "#,###.00;[RED]-#,###.00"
        cell = self.overview.cell(row=self.footer_row, column=8)  # label total opportunities
        cell.value = "  :total opportunities"
        cell.style = self.col_header
        cell = self.overview.cell(row=self.footer_row+2, column=5)  # label average overtime
        cell.value = "average overtime:"
        cell.style = self.date_dov_title
        cell = self.overview.cell(row=self.footer_row+2, column=6)  # calculate average overtime
        formula = "=F%s/$H$3" % self.footer_row
        cell.value = formula
        cell.style = self.calcs
        cell.number_format = "#,###.00;[RED]-#,###.00"
        cell = self.overview.cell(row=self.footer_row+2, column=7)  # calculate average opportunities
        formula = "=G%s/$H$3" % (str(self.footer_row))
        cell.value = formula
        cell.style = self.calcs
        cell.number_format = "#,###.00;[RED]-#,###.00"
        cell = self.overview.cell(row=self.footer_row+2, column=8)  # label average opportunities
        cell.value = "  :average opportunities"
        cell.style = self.col_header

    def build_header_weeklysheets(self):  # build the header for the weekly worksheet
        for i in range(15):
            cell = self.ws[i].cell(row=1, column=1)  # page title
            cell.value = "OTDL Equitability Worksheet"
            cell.style = self.ws_header
            self.ws[i].merge_cells('A1:H1')
            cell = self.ws[i].cell(row=1, column=16)  # week
            cell.value = "week: "
            cell.style = self.date_dov_title
            self.ws[i].merge_cells('P1:R1')
            cell = self.ws[i].cell(row=1, column=19)  # fill in week
            cell.value = self.week_label[i]
            cell.style = self.date_dov
            cell = self.ws[i].cell(row=2, column=1)  # date
            cell.value = "dates: "
            cell.style = self.date_dov_title
            cell = self.ws[i].cell(row=2, column=2)  # fill in date
            cell.value = self.dates_breakdown[i]
            cell.style = self.date_dov
            self.ws[i].merge_cells('B2:J2')
            cell = self.ws[i].cell(row=2, column=11)  # station
            cell.value = "station: "
            cell.style = self.date_dov_title
            self.ws[i].merge_cells('K2:M2')
            cell = self.ws[i].cell(row=2, column=14)  # fill in station
            cell.value = self.station
            cell.style = self.date_dov
            self.ws[i].merge_cells('N2:S2')
            cell = self.ws[i].cell(row=3, column=10)  # number of carriers
            cell.value = "# of carriers active on otdl: "
            cell.style = self.date_dov_title
            self.ws[i].merge_cells('J3:R3')
            cell = self.ws[i].cell(row=3, column=19)  # calculate number of carriers
            formula = "=%s!%s%s" % ("overview", "H", "3")
            cell.value = formula
            cell.style = self.calcs

    def build_columnheader_worksheets(self):
        for i in range(15):
            cell = self.ws[i].cell(row=5, column=1)  # name
            cell.value = "name"
            cell.style = self.col_header
            self.ws[i].merge_cells('A5:B5')
            cell = self.ws[i].cell(row=5, column=3)  # status
            cell.value = "status"
            cell.style = self.col_header
            self.ws[i].merge_cells('C5:D5')
            cell = self.ws[i].cell(row=5, column=5)  # sat
            cell.value = "sat"
            cell.style = self.col_center_header
            self.ws[i].merge_cells('E5:F5')
            cell = self.ws[i].cell(row=5, column=7)  # sun
            cell.value = "sun"
            cell.style = self.col_center_header
            self.ws[i].merge_cells('G5:H5')
            cell = self.ws[i].cell(row=5, column=9)  # mon
            cell.value = "mon"
            cell.style = self.col_center_header
            self.ws[i].merge_cells('I5:J5')
            cell = self.ws[i].cell(row=5, column=11)  # tue
            cell.value = "tue"
            cell.style = self.col_center_header
            self.ws[i].merge_cells('K5:L5')
            cell = self.ws[i].cell(row=5, column=13)  # wed
            cell.value = "wed"
            cell.style = self.col_center_header
            self.ws[i].merge_cells('M5:N5')
            cell = self.ws[i].cell(row=5, column=15)  # thu
            cell.value = "thu"
            cell.style = self.col_center_header
            self.ws[i].merge_cells('O5:P5')
            cell = self.ws[i].cell(row=5, column=17)  # fri
            cell.value = "fri"
            cell.style = self.col_center_header
            self.ws[i].merge_cells('Q5:R5')
            cell = self.ws[i].cell(row=5, column=19)  # weekly
            cell.value = "weekly"
            cell.style = self.col_center_header

    def build_main_worksheets(self):
        for i in range(len(self.ws)):  # for each worksheet
            row = 6
            for _ in range(len(self.carrier_overview)):  # for each carrier or empty set (get reach minimum rows)
                cell = self.ws[i].cell(row=row, column=1)  # name
                formula = "=IF(%s!A%s = \"\",\"\",%s!A%s)" % ("overview", str(row), "overview", str(row))
                cell.value = formula
                cell.style = self.ws_name
                self.ws[i].merge_cells('A' + str(row) + ":" + "B" + str(row+1))
                cell = self.ws[i].cell(row=row, column=3)  # status
                formula = "=IF(%s!C%s = \"\",\"\",%s!C%s)" % ("overview", str(row), "overview", str(row))
                cell.value = formula
                cell.style = self.ref_ot
                self.ws[i].merge_cells('C' + str(row) + ":" + "C" + str(row+1))
                cell = self.ws[i].cell(row=row, column=4)  # ref label
                cell.value = "ref"
                cell.style = self.ref_ot
                cell = self.ws[i].cell(row=row+1, column=4)  # ref label
                cell.value = "ot"
                cell.style = self.ref_ot
                cell = self.ws[i].cell(row=row, column=19)  # ref weekly total
                formula = "=SUM(%s%s,%s%s,%s%s,%s%s,%s%s,%s%s,%s%s)" \
                          % ("F", str(row), "H", str(row), "J", str(row), "L", str(row), "N", str(row),
                             "P", str(row), "R", str(row))
                cell.value = formula
                cell.style = self.calcs
                cell.number_format = "#,###.00;[RED]-#,###.00"
                cell = self.ws[i].cell(row=row + 1, column=19)  # overtime weekly total
                formula = "=SUM(%s%s:%s%s)" % ("E", str(row+1), "Q", str(row+1))
                cell.value = formula
                cell.style = self.calcs
                cell.number_format = "#,###.00;[RED]-#,###.00"
                row += 2

    @staticmethod
    def get_triad_merge(row, column):  # returns a string for for merge in triad group
        column -= 5
        letter = ("E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R")
        text = letter[column] + str(row+1) + ":" + letter[column+1] + str(row+1)
        return text

    def get_triad_refset(self, carrier, date):
        ringrefset = ("", "", "")  # default is empty set
        if self.carrier_overview[carrier][0]:  # if the carrier overview is not an empty set
            ringrefset = self.ringrefset[carrier][date]  # get the ring ref set for that carrier and date
        return ringrefset

    def triads_delegator(self):
        self.triad_row = 6
        for c in range(len(self.carrier_overview)):  # for each carrier including empty sets for minimum rows
            self.triad_week_index = 0  # week starts at zero
            self.triad_column = 5  # column starts at first day due to front padding
            for i in range(105):  # loop for self.front_padding + len(self.date_array) + self.end_padding
                self.build_triads(c, i)
            self.triad_row += 2

    def build_triads(self, c, i):
        ringrefset = self.get_triad_refset(c, i)
        # refusal indicator field
        cell = self.ws[self.triad_week_index].cell(row=self.triad_row, column=self.triad_column)
        cell.value = ringrefset[1]
        cell.style = self.input_s
        # refusal time field
        cell = self.ws[self.triad_week_index].cell(row=self.triad_row, column=self.triad_column + 1)
        cell.value = Handler(ringrefset[2]).str_to_float_or_str()
        cell.style = self.input_s
        cell.number_format = "#,###.00;[RED]-#,###.00"
        # overtime field
        cell = self.ws[self.triad_week_index].cell(row=self.triad_row + 1, column=self.triad_column)
        cell.value = Handler(ringrefset[0]).str_to_float_or_str()
        cell.style = self.input_s
        cell.number_format = "#,###.00;[RED]-#,###.00"
        # merge overtime field
        self.ws[self.triad_week_index].merge_cells(self.get_triad_merge(self.triad_row, self.triad_column))
        self.triad_column += 2
        if self.triad_column >= 19:
            self.triad_column = 5
            self.triad_week_index += 1

    def build_worksheet_footer(self):
        column_dict = {5: "E", 6: "F", 7: "G", 9: "I", 11: "K", 13: "M", 15: "O", 17: "Q"}
        for i in range(len(self.ws)):  # for each worksheet
            cell = self.ws[i].cell(row=self.footer_row, column=4)  # total overtime label
            cell.value = "total overtime:  "
            cell.style = self.date_dov_title
            cell = self.ws[i].cell(row=self.footer_row+2, column=4)  # average overtime label
            cell.value = "average overtime:  "
            cell.style = self.date_dov_title
            column_array = (5, 7, 9, 11, 13, 15, 17)
            for col in column_array:  # loop though for each column
                cell = self.ws[i].cell(row=self.footer_row, column=col)
                formula = self.get_totalovertime_formula(7, col)
                cell.value = formula
                cell.style = self.calcs
                cell.number_format = "#,###.00;[RED]-#,###.00"
                merge_string = self.get_triad_merge(self.footer_row-1, col)  # subtact 1 from row to match function
                self.ws[i].merge_cells(merge_string)
                cell = self.ws[i].cell(row=self.footer_row+2, column=col)
                formula = "=%s%s/$S$3" % (column_dict[col], self.footer_row)
                cell.value = formula
                cell.style = self.calcs
                cell.number_format = "#,###.00;[RED]-#,###.00"
                merge_string = self.get_triad_merge(self.footer_row + 1, col)  # subtact 1 from row to match function
                self.ws[i].merge_cells(merge_string)

    def build_instructions(self):
        cell = self.instructions.cell(row=1, column=1)  # page title
        cell.value = "OTDL Equitability Worksheet"
        cell.style = self.ws_header
        self.instructions.merge_cells('A1:E1')
        cell = self.instructions.cell(row=3, column=1)  # page title
        cell.value = "Instructions"
        cell.style = self.col_header_instructions
        self.instructions.merge_cells('A1:E1')
        cell = self.instructions.cell(row=5, column=1)
        text = "CAUTION: Do not write in grayed out cells. These cells have formulas. Writing in \n" \
               "them will delete those formulas. If this happens, do a CTRL Z to undo.\n\n" \
                "1. NAME:  Enter the carrier names on the first page only. Formulas on other pages \n" \
               "will import the name so that you don’t have to write it 15 times.\n\n" \
                "2. STATUS:  Enter the status on the first page only. Again formulas will do the work \n" \
                "and copy it to other pages. Enter “12”, “10” or “off” for 12 hour preference, 10 hour \n" \
                "preference or the carrier has gotten off the list. Leave the field blank if there is no \n" \
                "carrier. Failure to follow these instructions will result in the formulas calculating the \n" \
                "averages not working properly.\n\n" \
                "3.  MAKE UP:  This applies to make up opportunities from grievance settlements. \n" \
               "See JCAM Article 8.5.C.2 Remedies. This value is given in hours and clicks e.g. \n" \
               "5.32 or 14.00.\n\n" \
                "4. REFUSALS/OVERTIME: This displays refusals and overtime worked.\n\n" \
                "5. OPPORTUNITIES:   This displays total opportunities for overtime. This is all \n" \
               "refusals + overtime.\n\n" \
                "6. DIFF FROM AVERAGE:  This cell uses formulas to calculate the average \n" \
               "overtime of all carriers and the individual carrier’s difference from that. If they have \n" \
               "more than average, the number will be positive otherwise it will be negative. This will \n" \
               "be the core of your case in your grievance for OTDL equitability violations.\n" \
                "There are 15 worksheets. Each on represents a service week. Start with the first \n" \
                "week and proceed day by day.\n\n" \
                "7. For each day and each carrier there are are groups of three cells.\n\n" \
                "TOP LEFT (smaller) CELL:  This is a one letter explanation for any refusal: You can\n" \
                "use your own system, but my suggestions are: “p” for preference (the carrier \n" \
                "refused on the grounds that they are on the 10 hour list), “a” for annual \n" \
                "(the carrier missed overtime opportunities due to being on annual leave), “s” for sick \n" \
                "leave, for non scheduled day (the carrier worked on their nonscheduled day – see JCAM \n" \
                "Article 8.5.C.2.d Not Counted Toward “Equitability), “x” for exceptional \n" \
                "circumstances (see JCAM Article 8.5.E - Exceptional Situations May Excuse \n" \
                "Mandatory Overtime) and “r” for refusal (the carrier tells the supervisor to go jump \n" \
                "in a lake).\n\n" \
                "TOP RIGHT CELL: This is the amount of overtime that the carrier refused or was \n" \
                "unable to work.\n\n" \
                "BOTTOM CELL: This is overtime worked. Normally these is overtime worked off \n" \
                "the carrier’s own route, but it could mean any overtime depending on any local \n" \
                "agreements with management (LMOUs).\n\n" \
                "At the very bottom , there are totals and averages for the day. These are for your \n" \
                "information.\n"
        cell.value = text
        cell.style = self.instruct_text
        self.instructions.merge_cells('A5:I47')

    def save_open(self):  # name the excel file
        quarter = self.full_quarter.replace(" ", "")
        xl_filename = "ot_equit_" + quarter + ".xlsx"
        try:
            self.wb.save(dir_path('ot_equitability') + xl_filename)
            messagebox.showinfo("Spreadsheet generator",
                                "Your spreadsheet was successfully generated. \n"
                                "File is named: {}".format(xl_filename),
                                parent=self.frame)
            if sys.platform == "win32":
                os.startfile(dir_path('ot_equitability') + xl_filename)
            if sys.platform == "linux":
                subprocess.call(["xdg-open", 'kb_sub/ot_equitability/' + xl_filename])
            if sys.platform == "darwin":
                subprocess.call(["open", dir_path('ot_equitability') + xl_filename])
        except PermissionError:
            messagebox.showerror("Spreadsheet generator",
                                 "The spreadsheet was not opened. \n"
                                 "Suggestion: "
                                 "Make sure that identically named spreadsheets are closed "
                                 "(the file can't be overwritten while open).",
                                 parent=self.frame)

