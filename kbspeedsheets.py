import projvar  # custom libraries
from kbtoolbox import inquire, CarrierList, dir_path, isfloat, Convert, CarrierRecFilter, Rings, SpeedSettings, \
    ProgressBarDe
# standard libraries
from tkinter import messagebox
import os
import sys
import subprocess
from operator import itemgetter
# non standard libraries
from openpyxl import Workbook
from openpyxl.styles import NamedStyle, Font, Border, Side, Alignment, PatternFill, Protection


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
        if projvar.invran_weekly_span:  # if the investigation range is weekly
            start = projvar.invran_date_week[0]  # use sat - fri to curate the carrier list
            end = projvar.invran_date_week[6]
        else:  # if the investigation range is one day
            start = projvar.invran_date  # use the one day to curate the carrier list
            end = projvar.invran_date
        carriers = CarrierList(start, end, projvar.invran_station).get()  # first get a carrier list
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
            self.ws_titles = ["by employee id", "a", "b", "c,d", "e,f,g", "h", "i,j,k",
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
        elif not self.full_report and self.range == "week":
            title = "Speedsheet - Carriers"
            self.filename = "speed_" + str(format(projvar.invran_date_week[0], "%y_%m_%d")) + "_carrier_w" + ".xlsx"
        else:  # if not self.full_report and self.range == "day":
            title = "Speedsheet - Carriers"
            self.filename = "speed_" + str(format(projvar.invran_date, "%y_%m_%d")) + "_carrier_d" + ".xlsx"
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


class OpenText:
    def __init__(self):
        self.frame = None

    def open_docs(self, frame, doc):  # opens docs in the about_klusterbox() function
        self.frame = frame
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
                                 parent=self.frame)
