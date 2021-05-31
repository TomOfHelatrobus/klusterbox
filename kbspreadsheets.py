import projvar
from kbtoolbox import inquire, CarrierList, dir_path, isfloat
from tkinter import messagebox
import os
import sys
import subprocess
from datetime import timedelta
# non standard libraries
from openpyxl import Workbook
from openpyxl.styles import NamedStyle, Font, Border, Side, Alignment, PatternFill


class OvermaxSpreadsheet:
    def __init__(self):
        self.frame = None
        self.carrier_list = []
        self.wb = None  # workbook object
        self.violations = None  # workbook object sheet
        self.instructions = None  # workbook object sheet
        self.summary = None  # workbook object sheet
        self.startdate = None
        self.enddate = None
        self.dates = []
        self.rings = []
        self.min_rows = 0
        self.ws_header = None  # styles
        self.date_dov = None
        self.date_dov_title = None
        self.col_header = None
        self.col_center_header = None
        self.vert_header = None
        self.input_name = None
        self.input_s = None
        self.calcs = None
        self.vert_calcs = None
        self.instruct_text = None

    def create(self, frame):  # master method for calling methods
        self.frame = frame
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
        self.build_joint()
        self.save_open()

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

    def build_joint(self):
        summary_i = 7
        i = 9
        row_count = 0
        for line in self.carrier_list:
            # if there is a ring to match the carrier/ date then printe
            carrier_rings = []
            total = 0.0
            grandtotal = 0.0
            totals_array = ["", "", "", "", "", "", ""]
            leavetype_array = ["", "", "", "", "", "", ""]
            leavetime_array = ["", "", "", "", "", "", ""]
            cc = 0
            daily_violation = False
            for day in self.dates:
                for ring in self.rings:
                    if ring[0] == str(day) and ring[1] == line[1]:  # find if there are rings for the carrier
                        carrier_rings.append(ring)  # add any rings to an array
                        if isfloat(ring[2]):
                            totals_array[cc] = float(ring[2])
                            if float(ring[2]) > 12 and line[2] == "otdl":
                                daily_violation = True
                            if float(ring[2]) > 11.5 and line[2] != "otdl":
                                daily_violation = True
                        else:
                            totals_array[cc] = ring[2]
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

            if grandtotal > 60 or daily_violation is True:
                row_count += 1
                # output to the gui
                self.violations.row_dimensions[i].height = 10  # adjust all row height
                self.violations.row_dimensions[i + 1].height = 10
                self.violations.merge_cells('A' + str(i) + ':A' + str(i + 1))
                self.violations['A' + str(i)] = line[1]  # name
                self.violations['A' + str(i)].style = self.input_name
                self.violations.merge_cells('B' + str(i) + ':B' + str(i + 1))  # merge box for list
                self.violations['B' + str(i)] = line[2]  # list
                self.violations['B' + str(i)].style = self.input_s
                self.violations.merge_cells('C' + str(i) + ':C' + str(i + 1))  # merge box for weekly 5200
                self.violations['C' + str(i)] = float(total)  # total
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
                formula_i = "=%s!A%s" % ("violations", str(i))
                self.summary['A' + str(summary_i)] = formula_i
                self.summary['A' + str(summary_i)].style = self.input_name
                formula_j = "=%s!X%s" % ("violations", str(i))
                self.summary['B' + str(summary_i)] = formula_j
                self.summary['B' + str(summary_i)].style = self.input_s
                self.summary['B' + str(summary_i)].number_format = "#,###.00"
                self.summary.row_dimensions[summary_i].height = 10  # adjust all row height
                i += 2
                summary_i += 1
        # insert rows if minimum rows is not reached
        if row_count < self.min_rows:
            add_rows = self.min_rows - row_count
        else:
            add_rows = 0
        for add in range(add_rows):
            # output to the gui
            self.violations.row_dimensions[i].height = 10  # adjust all row height
            self.violations.row_dimensions[i + 1].height = 10
            self.violations.merge_cells('A' + str(i) + ':A' + str(i + 1))
            self.violations['A' + str(i)] = ""  # name
            self.violations['A' + str(i)].style = self.input_name
            self.violations.merge_cells('B' + str(i) + ':B' + str(i + 1))  # merge box for list
            self.violations['B' + str(i)] = ""  # list
            self.violations['B' + str(i)].style = self.input_s
            self.violations.merge_cells('C' + str(i) + ':C' + str(i + 1))  # merge box for weekly 5200
            self.violations['C' + str(i)] = ""  # total
            self.violations['C' + str(i)].style = self.input_s
            self.violations['C' + str(i)].number_format = "#,###.00;[RED]-#,###.00"
            # saturday
            self.violations.merge_cells('D' + str(i + 1) + ':E' + str(i + 1))  # merge box for sat 5200
            self.violations['D' + str(i)] = ""  # leave time
            self.violations['D' + str(i)].style = self.input_s
            self.violations['D' + str(i)].number_format = "#,###.00;[RED]-#,###.00"
            self.violations['E' + str(i)] = ""  # leave type
            self.violations['E' + str(i)].style = self.input_s
            self.violations['D' + str(i + 1)] = ""  # 5200 time
            self.violations['D' + str(i + 1)].style = self.input_s
            self.violations['D' + str(i + 1)].number_format = "#,###.00;[RED]-#,###.00"
            # sunday
            self.violations.merge_cells('F' + str(i + 1) + ':G' + str(i + 1))  # merge box for sun 5200
            self.violations['F' + str(i)] = ""  # leave time
            self.violations['F' + str(i)].style = self.input_s
            self.violations['F' + str(i)].number_format = "#,###.00;[RED]-#,###.00"
            self.violations['G' + str(i)] = ""  # leave type
            self.violations['G' + str(i)].style = self.input_s
            self.violations['F' + str(i + 1)] = ""  # 5200 time
            self.violations['F' + str(i + 1)].style = self.input_s
            self.violations['F' + str(i + 1)].number_format = "#,###.00;[RED]-#,###.00"
            # monday
            self.violations.merge_cells('H' + str(i + 1) + ':I' + str(i + 1))  # merge box for mon 5200
            self.violations['H' + str(i)] = ""  # leave time
            self.violations['H' + str(i)].style = self.input_s
            self.violations['H' + str(i)].number_format = "#,###.00;[RED]-#,###.00"
            self.violations['I' + str(i)] = ""  # leave type
            self.violations['I' + str(i)].style = self.input_s
            self.violations['H' + str(i + 1)] = ""  # 5200 time
            self.violations['H' + str(i + 1)].style = self.input_s
            self.violations['H' + str(i + 1)].number_format = "#,###.00;[RED]-#,###.00"
            # tuesday
            self.violations.merge_cells('J' + str(i + 1) + ':K' + str(i + 1))  # merge box for tue 5200
            self.violations['J' + str(i)] = ""  # leave time
            self.violations['J' + str(i)].style = self.input_s
            self.violations['J' + str(i)].number_format = "#,###.00;[RED]-#,###.00"
            self.violations['K' + str(i)] = ""  # leave type
            self.violations['K' + str(i)].style = self.input_s
            self.violations['J' + str(i + 1)] = ""  # 5200 time
            self.violations['J' + str(i + 1)].style = self.input_s
            self.violations['J' + str(i + 1)].number_format = "#,###.00;[RED]-#,###.00"
            # wednesday
            self.violations.merge_cells('L' + str(i + 1) + ':M' + str(i + 1))  # merge box for wed 5200
            self.violations['L' + str(i)] = ""  # leave time
            self.violations['L' + str(i)].style = self.input_s
            self.violations['L' + str(i)].number_format = "#,###.00;[RED]-#,###.00"
            self.violations['M' + str(i)] = ""  # leave type
            self.violations['M' + str(i)].style = self.input_s
            self.violations['L' + str(i + 1)] = ""  # 5200 time
            self.violations['L' + str(i + 1)].style = self.input_s
            self.violations['M' + str(i + 1)].number_format = "#,###.00;[RED]-#,###.00"
            # thursday
            self.violations.merge_cells('N' + str(i + 1) + ':O' + str(i + 1))  # merge box for thr 5200
            self.violations['N' + str(i)] = ""  # leave time
            self.violations['N' + str(i)].style = self.input_s
            self.violations['N' + str(i)].number_format = "#,###.00;[RED]-#,###.00"
            self.violations['O' + str(i)] = ""  # leave type
            self.violations['O' + str(i)].style = self.input_s
            self.violations['N' + str(i + 1)] = ""  # 5200 time
            self.violations['N' + str(i + 1)].style = self.input_s
            self.violations['N' + str(i + 1)].number_format = "#,###.00;[RED]-#,###.00"
            # friday
            self.violations.merge_cells('P' + str(i + 1) + ':Q' + str(i + 1))  # merge box for fri 5200
            self.violations['P' + str(i)] = ""  # leave time
            self.violations['P' + str(i)].style = self.input_s
            self.violations['P' + str(i)].number_format = "#,###.00;[RED]-#,###.00"
            self.violations['Q' + str(i)] = ""  # leave type
            self.violations['Q' + str(i)].style = self.input_s
            self.violations['P' + str(i + 1)] = ""  # 5200 time
            self.violations['P' + str(i + 1)].style = self.input_s
            self.violations['P' + str(i + 1)].number_format = "#,###.00;[RED]-#,###.00"
            # calculated fields
            # hidden columns
            formula_a = "=SUM(%s!D%s:P%s)+%s!D%s + %s!H%s + %s!J%s + %s!L%s + " \
                        "%s!N%s + %s!P%s" % ("violations", str(i + 1), str(i + 1),
                                             "violations", str(i), "violations", str(i), "violations", str(i),
                                             "violations", str(i), "violations", str(i), "violations", str(i))
            self.violations['R' + str(i)] = formula_a
            self.violations['R' + str(i)].style = self.calcs
            self.violations['R' + str(i)].number_format = "#,###.00;[RED]-#,###.00"
            formula_b = "=SUM(%s!C%s+%s!D%s+%s!H%s+%s!J%s+%s!L%s+%s!N%s+%s!P%s)" % \
                        ("violations", str(i), "violations", str(i), "violations", str(i),
                         "violations", str(i), "violations", str(i), "violations", str(i),
                         "violations", str(i))
            self.violations['R' + str(i + 1)] = formula_b
            self.violations['R' + str(i + 1)].style = self.calcs
            self.violations['R' + str(i + 1)].number_format = "#,###.00;[RED]-#,###.00"
            # weekly violation
            self.violations.merge_cells('S' + str(i) + ':S' + str(i + 1))  # merge box for weekly violation
            formula_c = "=MAX(IF(%s!R%s>%s!R%s,MAX(%s!R%s-60,0),MAX(%s!R%s-60)),0)" \
                        % ("violations", str(i), "violations", str(i + 1), "violations", str(i),
                           "violations", str(i + 1),)
            self.violations['S' + str(i)] = formula_c
            self.violations['S' + str(i)].style = self.calcs
            self.violations['S' + str(i)].number_format = "#,###.00;[RED]-#,###.00"
            # daily violation
            formula_d = "=IF(OR(%s!B%s=\"wal\",%s!B%s=\"nl\",%s!B%s=\"ptf\",%s!B%s=\"aux\")," \
                        "(SUM(IF(%s!D%s>11.5,%s!D%s-11.5,0)+IF(%s!H%s>11.5,%s!H%s-11.5,0)" \
                        "+IF(%s!J%s>11.5,%s!J%s-11.5,0)+IF(%s!L%s>11.5,%s!L%s-11.5,0)" \
                        "+IF(%s!N%s>11.5,%s!N%s-11.5,0)+IF(%s!P%s>11.5,%s!P%s-11.5,0)))," \
                        "(SUM(IF(%s!D%s>12,%s!D%s-12,0)+IF(%s!H%s>12,%s!H%s-12,0)+IF(%s!J%s>12,%s!J%s-12,0)" \
                        "+IF(%s!L%s>12,%s!L%s-12,0)+IF(%s!N%s>12,%s!N%s-12,0)+IF(%s!P%s>12,%s!P%s-12,0))))" \
                        % ("violations", str(i), "violations", str(i), "violations", str(i), "violations", str(i),
                           "violations", str(i + 1), "violations", str(i + 1), "violations", str(i + 1),
                           "violations", str(i + 1), "violations", str(i + 1), "violations", str(i + 1),
                           "violations", str(i + 1), "violations", str(i + 1), "violations", str(i + 1),
                           "violations", str(i + 1), "violations", str(i + 1), "violations", str(i + 1),
                           "violations", str(i + 1), "violations", str(i + 1), "violations", str(i + 1),
                           "violations", str(i + 1), "violations", str(i + 1), "violations", str(i + 1),
                           "violations", str(i + 1), "violations", str(i + 1), "violations", str(i + 1),
                           "violations", str(i + 1), "violations", str(i + 1), "violations", str(i + 1))
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
            formula_i = "=IF(%s!A%s=0,\"\",%s!A%s)" % ("violations", str(i), "violations", str(i))
            self.summary['A' + str(summary_i)] = formula_i
            self.summary['A' + str(summary_i)].style = self.input_name
            formula_j = "=%s!X%s" % ("violations", str(i))
            self.summary['B' + str(summary_i)] = formula_j
            self.summary['B' + str(summary_i)].style = self.input_s
            self.summary['B' + str(summary_i)].number_format = "#,###.00"
            self.summary.row_dimensions[summary_i].height = 10  # adjust all row height
            i += 2
            summary_i += 1
        # display totals for all violations
        self.violations.merge_cells('P' + str(i + 1) + ':T' + str(i + 1))
        self.violations['P' + str(i + 1)] = "Total Violations"
        self.violations['P' + str(i + 1)].style = self.col_header
        self.violations.merge_cells('V' + str(i + 1) + ':X' + str(i + 1))
        formula_k = "=SUM(%s!X%s:X%s)" % ("violations", "9", str(i))
        self.violations['V' + str(i + 1)] = formula_k
        self.violations['V' + str(i + 1)].style = self.calcs
        self.violations['V' + str(i + 1)].number_format = "#,###.00"
        self.violations.row_dimensions[i].height = 10  # adjust all row height
        self.violations.row_dimensions[i + 1].height = 10  # adjust all row height

    def save_open(self):
        xl_filename = "kb_om" + str(format(projvar.invran_date_week[0], "_%y_%m_%d")) + ".xlsx"
        if messagebox.askokcancel("Spreadsheet generator",
                                  "Do you want to generate a spreadsheet?",
                                  parent=self.frame):
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
