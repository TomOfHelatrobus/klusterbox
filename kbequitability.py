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
        self.carrierlist = []
        self.recset = []
        self.minrows = 1
        self.otcalcpref = "off_route"  # preference for overtime calculation - "off_route" or "all"
        self.carrier_overview = []  # a list of carrier's name, status and makeups
        self.date_array = []  # a list of all days in the quarter as a datetimes
        self.ringrefset = []  # multidimensional array - daily rings/refusals for each otdl carrier
        self.wb = None

    def create(self, frame, date, station):
        self.frame = frame
        self.station = station
        self.date = date  # a datetime object from the quarter is passed and used as date
        self.date_breakdown()  # the passed datetime object is broken down into year and month
        self.define_quarter()  # the year and month are used to generate quarter and full quarter
        self.get_dates()  # use quarter information to get start and end date
        self.get_carrierlist()  # generate a raw list of carriers at station before or on end date.
        self.get_recsets()  # filter the carrierlist to get only otdl carriers
        self.get_settings()  # get minimum rows and ot calculation preference
        self.get_carier_overview()  # build a list of carrier's name, status and makeups
        self.get_date_array()  # get a list of all days in the quarter as datetime objects
        self.get_ringrefset()  # build multidimensional array - daily rings/refusals for each carrier
        print(self.ringrefset)

    def date_breakdown(self):  # breakdown the date into year and month
        self.year = int(self.date.strftime("%Y"))
        self.month = int(self.date.strftime("%m"))

    def define_quarter(self):
        self.quarter = Quarter(self.month).find()  # convert the month into a quarter - 1 through 4.
        self.full_quarter = str(self.year) + " - " + str(self.quarter)  # create a string expressing the year - quarter

    def get_dates(self):
        startdate = (datetime(self.year, 1, 1), datetime(self.year, 4, 1), datetime(self.year, 7, 1),
                     datetime(self.year, 10, 1))
        enddate = (datetime(self.year, 3, 31), datetime(self.year, 6, 30), datetime(self.year, 9, 30),
                   datetime(self.year, 12, 31))
        self.startdate = startdate[int(self.quarter) - 1]
        self.enddate = enddate[int(self.quarter) - 1]

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
            return makeup[0]
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

    def get_date_array(self):
        running_date = self.startdate
        while running_date <= self.enddate:
            self.date_array.append(running_date)
            running_date += timedelta(days=1)

    def get_ringrefset(self):
        for i in range(len(self.carrier_overview)):
            self.ringrefset.append([])
            self.get_daily_ringrefs(i)

    def get_daily_ringrefs(self, index):
        daily_ringref = []
        for date in self.date_array:
            total = ""
            code = ""
            moves = ""
            ref_type = ""
            ref_time = ""
            add_this = [total, code, moves, ref_type, ref_time]
            daily_ringref.append(add_this)
        self.ringrefset[index] = daily_ringref
