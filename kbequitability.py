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
