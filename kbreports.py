import projvar
from kbtoolbox import inquire, CarrierList, dt_converter, NsDayDict, dir_path
from tkinter import messagebox, simpledialog
import os
import sys
import subprocess
from datetime import datetime, timedelta


class Reports:
    def __init__(self, frame):
        self.frame = frame
        self.start_date = projvar.invran_date
        self.end_date = projvar.invran_date
        if projvar.invran_weekly_span:
            self.start_date = projvar.invran_date_week[0]
            self.end_date = projvar.invran_date_week[6]
        self.carrier_list = []

    def get_carrierlist(self):
        # get carrier list
        self.carrier_list = CarrierList(self.start_date, self.end_date, projvar.invran_station).get()

    @staticmethod
    def rpt_dt_limiter(date, first_date):  # return the first day if it is earlier than the date
        if date < first_date:
            return first_date
        else:
            return date

    @staticmethod
    def rpt_ns_fixer(nsday_code):  # remove the day from the ns_code if fixed
        if "fixed" in nsday_code:
            fix = nsday_code.split(":")
            return fix[0]
        else:
            return nsday_code

    def rpt_carrier(self):  # Generate and display a report of carrier routes and nsday
        self.get_carrierlist()
        ns_dict = NsDayDict.get_custom_nsday()  # get the ns day names from the dbase
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")  # create a file name
        filename = "report_carrier_route_nsday" + "_" + stamp + ".txt"
        report = open(dir_path('report') + filename, "w")
        report.write("\nCarrier Route and NS Day Report\n\n\n")
        report.write('   Showing results for:\n')
        report.write('      Station: {}\n'.format(projvar.invran_station))
        if not projvar.invran_weekly_span:  # if investigation range is daily
            f_date = projvar.invran_date
            report.write('      Date: {}\n'.format(f_date.strftime("%m/%d/%Y")))
        else:  # if investigation range is weekly
            f_date = projvar.invran_date_week[0]  # use the first day of the service week
            report.write('      Dates: {} through {}\n'
                         .format(projvar.invran_date_week[0].strftime("%m/%d/%Y"),
                                 projvar.invran_date_week[6].strftime("%m/%d/%Y")))
        report.write('      Pay Period: {}\n\n'.format(projvar.pay_period))
        report.write('{:>4} {:<23} {:<13} {:<29} {:<10}\n'.format("", "Carrier Name", "N/S Day", "Route/s",
                                                                  "Start Date"))
        report.write('     ------------------------------------------------------------------- ----------\n')
        i = 1
        for line in self.carrier_list:
            ii = 0
            for rec in reversed(line):
                if not ii:
                    report.write('{:>4} {:<23} {:<4} {:<8} {:<29}\n'
                                 .format(i, rec[1], projvar.ns_code[rec[3]], self.rpt_ns_fixer(ns_dict[rec[3]]),
                                         rec[4]))
                else:
                    report.write('{:>4} {:<23} {:<4} {:<8} {:<29} {:<10}\n'
                                 .format("", rec[1], projvar.ns_code[rec[3]], self.rpt_ns_fixer(ns_dict[rec[3]]),
                                         rec[4], self.rpt_dt_limiter(dt_converter(rec[0]), f_date).strftime("%A")))
                ii += 1
            if i % 3 == 0:
                report.write('     ------------------------------------------------------------------- ----------\n')
            i += 1
        report.close()
        if sys.platform == "win32":  # open the text document
            os.startfile(dir_path('report') + filename)
        if sys.platform == "linux":
            subprocess.call(["xdg-open", 'kb_sub/report/' + filename])
        if sys.platform == "darwin":
            subprocess.call(["open", dir_path('report') + filename])

    def rpt_carrier_route(self):  # Generate and display a report of carrier routes
        self.get_carrierlist()
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = "report_carrier_route" + "_" + stamp + ".txt"
        report = open(dir_path('report') + filename, "w")
        report.write("\nCarrier Route Report\n\n\n")
        report.write('   Showing results for:\n')
        report.write('      Station: {}\n'.format(projvar.invran_station))
        if not projvar.invran_weekly_span:  # if investigation range is daily
            f_date = projvar.invran_date
            report.write('      Date: {}\n'.format(f_date.strftime("%m/%d/%Y")))
        else:
            f_date = projvar.invran_date_week[0]
            report.write('      Date: {} through {}\n'
                         .format(projvar.invran_date_week[0].strftime("%m/%d/%Y"),
                                 projvar.invran_date_week[6].strftime("%m/%d/%Y")))
        report.write('      Pay Period: {}\n\n'.format(projvar.pay_period))
        report.write('{:>4}  {:<22} {:<29}\n'.format("", "Carrier Name", "Route/s"))
        report.write('      ---------------------------------------------------- -------------------\n')
        i = 1
        for line in self.carrier_list:
            ii = 0
            for rec in reversed(line):  # reverse order so earliest one appears first
                if not ii:  # if the first record
                    report.write('{:>4}  {:<22} {:<29}\n'.format(i, rec[1], rec[4]))
                else:  # if not the first record, use alternate format
                    report.write('{:>4}  {:<22} {:<29} effective {:<10}\n'
                                 .format("", rec[1], rec[4],
                                         self.rpt_dt_limiter(dt_converter(rec[0]), f_date).strftime("%A")))
                ii += 1
            if i % 3 == 0:
                report.write('      ---------------------------------------------------- -------------------\n')
            i += 1
        report.close()
        if sys.platform == "win32":  # open the text document
            os.startfile(dir_path('report') + filename)
        if sys.platform == "linux":
            subprocess.call(["xdg-open", 'kb_sub/report/' + filename])
        if sys.platform == "darwin":
            subprocess.call(["open", dir_path('report') + filename])

    def rpt_carrier_nsday(self):  # Generate and display a report of carrier ns day
        self.get_carrierlist()
        ns_dict = NsDayDict.get_custom_nsday()  # get the ns day names from the dbase
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = "report_carrier_nsday" + "_" + stamp + ".txt"
        report = open(dir_path('report') + filename, "w")
        report.write("\nCarrier NS Day\n\n\n")
        report.write('   Showing results for:\n')
        report.write('      Station: {}\n'.format(projvar.invran_station))
        if not projvar.invran_weekly_span:  # if investigation range is daily
            f_date = projvar.invran_date
            report.write('      Date: {}\n'.format(f_date.strftime("%m/%d/%Y")))
        else:
            f_date = projvar.invran_date_week[0]
            report.write('      Date: {} through {}\n'
                         .format(projvar.invran_date_week[0].strftime("%m/%d/%Y"),
                                 projvar.invran_date_week[6].strftime("%m/%d/%Y")))
        report.write('      Pay Period: {}\n\n'.format(projvar.pay_period))
        report.write('{:>4}  {:<22} {:<17}\n'.format("", "Carrier Name", "N/S Day"))
        report.write('      ----------------------------------------  -------------------\n')
        i = 1
        for line in self.carrier_list:
            ii = 0
            for rec in reversed(line):
                if not ii:
                    report.write('{:>4}  {:<22} {:<5}{:<12}\n'
                                 .format(i, rec[1], projvar.ns_code[rec[3]], self.rpt_ns_fixer(ns_dict[rec[3]])))
                else:
                    report.write('{:>4}  {:<22} {:<5}{:<12}  effective {:<10}\n'
                                 .format("", rec[1], projvar.ns_code[rec[3]], self.rpt_ns_fixer(ns_dict[rec[3]]),
                                         self.rpt_dt_limiter(dt_converter(rec[0]), f_date).strftime("%A")))
                ii += 1
            if i % 3 == 0:
                report.write('      ----------------------------------------  -------------------\n')
            i += 1
        report.close()
        if sys.platform == "win32":  # open the text document
            os.startfile(dir_path('report') + filename)
        if sys.platform == "linux":
            subprocess.call(["xdg-open", 'kb_sub/report/' + filename])
        if sys.platform == "darwin":
            subprocess.call(["open", dir_path('report') + filename])

    def rpt_carrier_by_list(self):
        self.get_carrierlist()
        list_dict = {"nl": "No List", "wal": "Work Assignment List",
                     "otdl": "Overtime Desired List", "ptf": "Part Time Flexible", "aux": "Auxiliary Carrier"}
        # initialize arrays for data sorting
        otdl_array = []
        wal_array = []
        nl_array = []
        ptf_array = []
        aux_array = []
        for line in self.carrier_list:
            for carrier in line:
                if carrier[2] == "otdl":
                    otdl_array.append(carrier)
                if carrier[2] == "wal":
                    wal_array.append(carrier)
                if carrier[2] == "nl":
                    nl_array.append(carrier)
                if carrier[2] == "ptf":
                    ptf_array.append(carrier)
                if carrier[2] == "aux":
                    aux_array.append(carrier)
        array_var = nl_array + wal_array + otdl_array + ptf_array + aux_array  #
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")  # create a file name
        filename = "report_carrier_by_list" + "_" + stamp + ".txt"
        report = open(dir_path('report') + filename, "w")
        report.write("\nCarrier by List\n\n")
        report.write('   Showing results for:\n')
        report.write('      Station: {}\n'.format(projvar.invran_station))
        if not projvar.invran_weekly_span:  # if investigation range is daily
            f_date = projvar.invran_date
            report.write('      Date: {}\n'.format(f_date.strftime("%m/%d/%Y")))
        else:
            f_date = projvar.invran_date_week[0]
            report.write('      Dates: {} through {}\n'
                         .format(projvar.invran_date_week[0].strftime("%m/%d/%Y"),
                                 projvar.invran_date_week[6].strftime("%m/%d/%Y")))
        report.write('      Pay Period: {}\n'.format(projvar.pay_period))
        i = 1
        last_list = ""  # this is a indicator for when a new list is starting
        for line in array_var:
            if last_list != line[2]:  # if the new record is in a different list that the last
                report.write('\n\n      {:<20}\n\n'
                             .format(list_dict[line[2]]))  # write new headers
                report.write('{:>4}  {:<22} {:>4}\n'.format("", "Carrier Name", "List"))
                report.write('      ---------------------------  -------------------\n')
                i = 1
            if dt_converter(line[0]) not in projvar.invran_date_week:
                report.write('{:>4}  {:<22} {:>4}\n'.format(i, line[1], line[2]))
            else:
                report.write('{:>4}  {:<22} {:>4}  effective {:<10}\n'
                             .format(i, line[1], line[2],
                                     self.rpt_dt_limiter(dt_converter(line[0]), f_date).strftime("%A")))
            if i % 3 == 0:
                report.write('      ---------------------------  -------------------\n')
            last_list = line[2]
            i += 1
        report.close()
        if sys.platform == "win32":  # open the text document
            os.startfile(dir_path('report') + filename)
        if sys.platform == "linux":
            subprocess.call(["xdg-open", 'kb_sub/report/' + filename])
        if sys.platform == "darwin":
            subprocess.call(["open", dir_path('report') + filename])

    @staticmethod
    def rpt_carrier_history(carrier):
        sql = "SELECT effective_date, list_status, ns_day, route_s, station" \
              " FROM carriers WHERE carrier_name = '%s' ORDER BY effective_date DESC" % carrier
        results = inquire(sql)
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = "report_carrier_history" + "_" + stamp + ".txt"
        report = open(dir_path('report') + filename, "w")
        report.write("\nCarrier Status Change History\n\n")
        report.write('   Showing all status changes in the klusterbox database for {}\n\n'.format(carrier))
        report.write('{:<16}{:<8}{:<10}{:<31}{:<25}\n'
                     .format("Date Effective", "List", "N/S Day", "Route/s", "Station"))
        report.write('----------------------------------------------------------------------------------\n')
        i = 1
        for line in results:
            report.write('{:<16}{:<8}{:<10}{:<31}{:<25}\n'
                         .format(dt_converter(line[0]).strftime("%m/%d/%Y"), line[1], line[2], line[3], line[4]))
            if i % 3 == 0:
                report.write('----------------------------------------------------------------------------------\n')
            i += 1
        report.close()
        if sys.platform == "win32":  # open the text document
            os.startfile(dir_path('report') + filename)
        if sys.platform == "linux":
            subprocess.call(["xdg-open", 'kb_sub/report/' + filename])
        if sys.platform == "darwin":
            subprocess.call(["open", dir_path('report') + filename])

    def pay_period_guide(self):
        i = 0
        year = simpledialog.askinteger("Pay Period Guide", "Enter the year you want generated.", parent=self.frame,
                                       minvalue=2, maxvalue=9999)
        if year is not None:
            firstday = datetime(1, 12, 22, 0, 0, 0)
            while int(firstday.strftime("%Y")) != year - 1:
                firstday += timedelta(weeks=52)
                if int(firstday.strftime("%m")) <= 12 and int(firstday.strftime("%d")) <= 12:
                    firstday += timedelta(weeks=2)
            filename = "pp_guide" + "_" + str(year) + ".txt"  # create the filename for the text doc
            report = open(dir_path('pp_guide') + filename, "w")  # create the document
            report.write("\nPay Period Guide\n")
            report.write("Year: " + str(year) + "\n")
            report.write("---------------------------------------------\n\n")
            report.write("                 START (Sat):   END (Fri):         \n")
            for i in range(1, 27):
                # calculate dates
                wk1_start = firstday
                wk1_end = firstday + timedelta(days=6)
                wk2_start = firstday + timedelta(days=7)
                wk2_end = firstday + timedelta(days=13)
                report.write("PP: " + str(i).zfill(2) + "\n")
                report.write(
                    "\t week 1: " + wk1_start.strftime("%b %d, %Y") + " - " + wk1_end.strftime("%b %d, %Y") + "\n")
                report.write(
                    "\t week 2: " + wk2_start.strftime("%b %d, %Y") + " - " + wk2_end.strftime("%b %d, %Y") + "\n")
                # increment the first day by two weeks
                firstday += timedelta(days=14)
            # handle cases where there are 27 pay periods
            if int(firstday.strftime("%m")) <= 12 and int(firstday.strftime("%d")) <= 12:
                i += 1
                wk1_start = firstday
                wk1_end = firstday + timedelta(days=6)
                wk2_start = firstday + timedelta(days=7)
                wk2_end = firstday + timedelta(days=13)
                report.write("PP: " + str(i).zfill(2) + "\n")
                report.write(
                    "\t week 1: " + wk1_start.strftime("%b %d, %Y") + " - " + wk1_end.strftime("%b %d, %Y") + "\n")
                report.write(
                    "\t week 2: " + wk2_start.strftime("%b %d, %Y") + " - " + wk2_end.strftime("%b %d, %Y") + "\n")
            report.close()
            if sys.platform == "win32":
                os.startfile(dir_path('pp_guide') + filename)
            if sys.platform == "linux":
                subprocess.call(["xdg-open", 'kb_sub/pp_guide/' + filename])
            if sys.platform == "darwin":
                subprocess.call(["open", dir_path('pp_guide') + filename])


class Messenger:
    def __init__(self, frame):
        self.frame = frame

    def location_klusterbox(self):  # provides the location of the program
        archive = ""
        dbase = None
        path = None
        if sys.platform == "darwin":
            if projvar.platform == "macapp":
                path = "Applications"
                dbase = os.path.expanduser("~") + '/Documents/.klusterbox/' + 'mandates.sqlite'
                archive = os.path.expanduser("~") + '/Documents/klusterbox'
            if projvar.platform == "py":
                path = os.getcwd()
                dbase = os.getcwd() + '/kb_sub/mandates.sqlite'
                archive = os.getcwd() + '/kb_sub'
        else:
            if projvar.platform == "winapp":
                path = os.getcwd()
                dbase = os.path.expanduser("~") + '\Documents\.klusterbox\\' + 'mandates.sqlite'
                archive = os.path.expanduser("~") + '\Documents\klusterbox'
            else:
                path = os.getcwd()
                dbase = os.getcwd() + '\kb_sub\mandates.sqlite'
                archive = os.getcwd() + '\kb_sub'

        messagebox.showinfo("KLUSTERBOX ",
                            "On this computer Klusterbox is located at:\n"
                            "{}\n\nThe Klusterbox database is located at \n"
                            "{}\n\nThe Klusterbox archive is located at \n"
                            "{}".format(path, dbase, archive),
                            parent=self.frame)

    def tolerance_info(self, switch):
        text = ""
        if switch == "OT_own_route":
            text = "Sets the tolerance for no list carrier overtime\n" \
                   "\n" \
                   "Enter a value in clicks between 0 and .99"
        if switch == "OT_off_route":
            text = "Sets the tolerance for no list and work assignment \n" \
                   "list carriers for overtime off their own routes.\n\n" \
                   "Enter a value in clicks between 0 and .99"
        if switch == "availability":
            text = "Sets the tolerance for availability of otdl and " \
                   "aux carriers. Applies to availability to 10, 11.5 \n" \
                   "and 12 hour columns.\n\n" \
                   "Enter a value in clicks between 0 and .99"
        if switch == "min_nl":
            text = "Sets the minimum number of rows for the No List " \
                   "section of the spreadsheet. \n\n" \
                   "Enter a value between 0 and 100"
        if switch == "min_wal":
            text = "Sets the minimum number of rows for the Work Assignment " \
                   "section of the spreadsheet. \n\n" \
                   "Enter a value between 0 and 100"
        if switch == "min_otdl":
            text = "Sets the minimum number of rows for the OT Desired " \
                   "section of the spreadsheet. \n\n" \
                   "Enter a value between 0 and 100"
        if switch == "min_aux":
            text = "Sets the minimum number of rows for the Auxiliary " \
                   "section of the spreadsheet. \n\n" \
                   "Enter a value between 0 and 100"
        if switch == "min_overmax":
            text = "Sets the minimum number of rows for the " \
                   "12 and 60 Hour Violations spreadsheet. \n\n" \
                   "Enter a value between 0 and 100"
        messagebox.showinfo("About Tolerances", text, parent=self.frame)