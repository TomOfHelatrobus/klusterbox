"""
a klusterbox module: Klusterbox Employee Everything Report Reader for CSV files
provides code for reading employee everything reports in the csv format.
such as the employee everything report reader
"""
from kbtoolbox import dir_path, dir_filedialog, ee_ns_detect, find_pp, Handler
import sys
import csv
import os
import subprocess
from tkinter import filedialog, messagebox
from datetime import timedelta
from kbcsv_repair import CsvRepair
from operator import itemgetter


class MaxHr:
    """
    generates a report for 12/60 hour violations
    """
    def __init__(self):
        self.frame = None  # get tkinter frame for messageboxes
        self.path = dir_filedialog()
        self.file_path = None
        self.csv_fix = None  # create a CsvRepair object
        # returns a file path for a checked and, if needed, fixed csv file.
        self.file_path = None
        self.filename = None
        self.report = None
        self.day_xlr = {"Saturday": "sat", "Sunday": "sun", "Monday": "mon", "Tuesday": "tue",
                        "Wednesday": "wed", "Thursday": "thr", "Friday": "fri"}
        self.leave_xlr = {"49": "owcp   ", "55": "annual ", "56": "sick   ", "58": "holiday",
                          "59": "lwop   ", "60": "lwop   "}
        self.maxhour = []
        self.max_aux_day = []
        self.max_ft_day = []
        self.extra_hours = []
        self.all_extra = []
        self.adjustment = []
        self.target_file = None
        self.a_file = None
        self.line = None
        self.pp = ""
        self.ft = ""
        self.days = ("Saturday", "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday")
        self.day_hours = []
        self.good_id = ""

    def run(self, frame):
        """ master method for controlling other methods """
        self.frame = frame
        if not self.get_filepath():  # get the file
            return
        self.check_filetype()  # check extension, make sure the path is not an empty string.
        self.generate_proxyfile()  # create a proxy file to repair any double spaces or empty arrays
        self.get_file()  # get the file object
        if self.skimmer():  # break down file into line and analyze.
            self.make_report()
            self.write()
            self.open()

    def get_filepath(self):
        """ get the file and return True, else return False """
        self.file_path = filedialog.askopenfilename(initialdir=self.path, filetypes=[("Excel files", "*.csv *.xls")])
        if not self.file_path:
            return False
        return True

    def check_filetype(self):
        """ makes sure the file is an csv or xls file. """
        if self.file_path == "":
            return False
        if self.file_path[-4:].lower() == ".csv" or self.file_path[-4:].lower() == ".xls":
            return True
        messagebox.showerror("Report Generator",
                             "The file you have selected is not a .csv or .xls file.\n"
                             "You must select a file with a .csv or .xls extension.",
                             parent=self.frame)
        return False

    def generate_proxyfile(self):
        """ create a proxy file to repair any double spaces or empty arrays in the original file. """
        self.csv_fix = CsvRepair()  # create a CsvRepair object
        self.file_path = self.csv_fix.run(self.file_path)

    def get_file(self):
        """ get the file object """
        self.target_file = open(self.file_path, newline="")
        self.a_file = csv.reader(self.target_file)

    def skimmer(self):
        """ break down into lines and analyze """
        self.good_id = "no"
        cc = 0
        for self.line in self.a_file:
            if cc == 0:
                if not self.check_for_eereport():
                    return False
            if cc == 2:  # on the second line
                self.pp = self.line[0]  # find the pay period
                self.pp = self.pp.strip()  # strip whitespace out of pay period information
            if cc != 0:  # on all but the first line
                self.anaylsis()
            cc += 1
        return True

    def check_for_eereport(self):
        """ make sure the file is an employee everything report. """
        if self.line[0][:8] != "TAC500R3":
            messagebox.showwarning("File Selection Error",
                                   "The selected file does not appear to be an "
                                   "Employee Everything report.",
                                   parent=self.frame)
            self.target_file.close()
            self.csv_fix.destroy()
            return False
        return True

    def anaylsis(self):
        """ anaylsize the report line by line """

        if self.line[18] == "Base" and self.good_id and len(self.day_hours) > 0:
            # find fri hours for friday adjustment
            fri_hrs = 0
            for t in self.day_hours:  # get the friday hours
                if t[3] == "Friday":
                    fri_hrs += float(t[2])
            # find thu hours for thursday adjustment
            thu_hrs = 0
            for t in self.day_hours:  # find the thursday hours
                if t[3] == "Thursday":
                    thu_hrs += float(t[2])
            # find wed hours for wednesday adjustment
            wed_hrs = 0
            for t in self.day_hours:  # find the wednesday hours
                if t[3] == "Wednesday":
                    wed_hrs += float(t[2])
            # find the weekly total by adding daily totals
            wkly_total = 0
            for t in self.day_hours:
                wkly_total += float(t[2])
            if wkly_total > 60:
                add_maxhr = (self.day_hours[0][0].lower(), self.day_hours[0][1].lower(), wkly_total)
                self.maxhour.append(add_maxhr)
                for item in self.extra_hours:  # get any extra hours codes for non-5200 hours list
                    self.all_extra.append(item)
                # find the all adjustments
                if self.ft:
                    # find friday adjustment
                    fri_post_60 = float(wkly_total - 60)
                    if fri_hrs > 12:
                        fri_over = fri_hrs - 12
                        if fri_over < fri_post_60:
                            fri_adj = fri_over
                        else:
                            fri_adj = fri_post_60
                        add_adjustment = ("fri", self.day_hours[0][0].lower(), self.day_hours[0][1].lower(), fri_adj)
                        self.adjustment.append(add_adjustment)
                    # find the thursday adjustment
                    thu_post_60 = float(wkly_total - 60) - fri_hrs
                    if thu_hrs > 12 and thu_post_60 > 0:
                        thu_over = thu_hrs - 12
                        if thu_over < thu_post_60:
                            thu_adj = thu_over
                        else:
                            thu_adj = thu_post_60
                        add_adjustment = ("thu", self.day_hours[0][0].lower(), self.day_hours[0][1].lower(), thu_adj)
                        self.adjustment.append(add_adjustment)
                    # find the wednesday adjustment
                    wed_post_60 = float(wkly_total - 60) - fri_hrs - thu_hrs
                    if wed_hrs > 12 and wed_post_60 > 0:
                        wed_over = wed_hrs - 12
                        if wed_over < wed_post_60:
                            wed_adj = wed_over
                        else:
                            wed_adj = wed_post_60
                        add_adjustment = (
                            "wed", self.day_hours[0][0].lower(), self.day_hours[0][1].lower(), wed_adj)
                        self.adjustment.append(add_adjustment)
            del self.day_hours[:]
            del self.extra_hours[:]
        # find first self.line of specific carrier
        if self.line[18] == "Base" and self.line[19] in ("844", "134", "434"):
            self.good_id = self.line[4]  # remember id of carriers who are FT or aux carriers
            if self.line[19] in ("844", "434"):
                self.ft = False
            else:
                self.ft = True
        if self.good_id == self.line[4] and self.line[18] != "Base":
            if self.line[18] in self.days:  # get the hours for each day
                spt_20 = self.line[20].split(':')  # split to get code and hours
                hr_type = spt_20[0][1] + spt_20[0][2]  # parse hour code to 2 digits
                # if hr_type in hr_codes:  # compare to array of hour codes
                if hr_type == "52":  # compare to array of hour codes
                    if float(spt_20[1]) > 11.5 and not self.ft:
                        add_max_aux = (self.line[5].lower(), self.line[6].lower(), self.line[18], spt_20[1])
                        self.max_aux_day.append(add_max_aux)
                    if float(spt_20[1]) > 12 and self.ft:
                        add_max_ft = (self.line[5].lower(), self.line[6].lower(), self.line[18], spt_20[1])
                        self.max_ft_day.append(add_max_ft)
                    if self.ft:  # increment daily totals to find weekly total
                        add_day_hours = (self.line[5].lower(), self.line[6].lower(), spt_20[1], self.line[18])
                        self.day_hours.append(add_day_hours)
                extra_hour_codes = ("49", "55", "56", "58")  # paid leave types only , (lwop "59", "60")
                if hr_type in extra_hour_codes and self.ft:  # if there is holiday pay
                    add_day_hours = (self.line[5].lower(), self.line[6].lower(), spt_20[1], self.line[18])
                    self.day_hours.append(add_day_hours)
                    add_extra_hours = (self.line[5].lower(), self.line[6].lower(), self.line[18], hr_type, spt_20[1])
                    self.extra_hours.append(add_extra_hours)  # track non 5200 hours

    def make_report(self):
        """ make report object """
        pp_str = self.pp[:-3] + "_" + self.pp[4] + self.pp[5] + "_" + self.pp[6]
        self.filename = "max" + "_" + pp_str + ".txt"
        self.report = open(dir_path('over_max') + self.filename, "w")

    def write(self):
        """ write the report """
        weekly_max = []  # array hold each carrier's hours for the week
        daily_max = []  # array hold each carrier's sum of maximum daily hours for the week
        if len(self.maxhour) > 0 or len(self.max_ft_day) > 0 or len(self.max_aux_day) > 0:
            self.report.write("12 and 60 Hour Violations\n\n")
            # disply pay period
            self.report.write("pay period: " + self.pp[:-3] + " " + self.pp[4] + self.pp[5] + "-" + self.pp[6] + "\n")
            pp_date = find_pp(int(self.pp[:-3]), self.pp[-3:])  # send year and pp to get the date
            pp_date_end = pp_date + timedelta(days=6)  # add six days to get the last part of the range
            self.report.write(
                "week of: " + pp_date.strftime("%x") + " - " + pp_date_end.strftime("%x") + "\n")  # printe date
            self.report.write("\n60 hour violations \n\n")
            self.report.write("name                              total   over\n")
            self.report.write("-----------------------------------------------\n")
            if len(self.maxhour) == 0:
                self.report.write("no violations" + "\n")
            else:
                diff_total = 0
                self.maxhour.sort(key=itemgetter(0))
                for item in self.maxhour:
                    tabs = 30 - (len(item[0]))
                    period = "."
                    period += tabs * "."
                    diff = float(item[2]) - 60
                    diff_total += diff
                    self.report.write(item[0] + ", " + item[1] + period + "{0:.2f}".format(float(item[2])) + "   "
                                      + "{0:.2f}".format(float(diff)).rjust(5, " ") + "\n")
                    wmax_add = [item[0], item[1], diff]
                    weekly_max.append(wmax_add)  # catch totals of violations for the week
                self.report.write("\n" + "                                   total:  " +
                                  "{0:.2f}".format(float(diff_total)) + "\n")
            self.all_extra.sort(key=itemgetter(0))
            self.report.write("\nNon 5200 codes contributing to 60 hour violations  \n\n")
            self.report.write("day   name                            hr type   hours\n")
            self.report.write("-----------------------------------------------------\n")
            if len(self.all_extra) == 0:
                self.report.write("no contributions" + "\n")
            for i in range(len(self.all_extra)):
                tabs = 28 - (len(self.all_extra[i][0]))
                period = "."
                period += tabs * "."
                self.report.write(self.day_xlr[self.all_extra[i][2]] + "   " + self.all_extra[i][0] +
                                  ", " + self.all_extra[i][1] + period + self.leave_xlr[self.all_extra[i][3]] +
                                  "  " + "{0:.2f}".format(float(self.all_extra[i][4])).rjust(5, " ") + "\n")
            self.report.write("\n\n12 hour full time carrier violations \n\n")
            self.report.write("day   name                        total   over   sum\n")
            self.report.write("-----------------------------------------------------\n")
            if len(self.max_ft_day) == 0:
                self.report.write("no violations" + "\n")
            diff_sum = 0
            sum_total = 0
            self.max_ft_day.sort(key=itemgetter(0))
            for i in range(len(self.max_ft_day)):
                jump = "no"  # triggers an analysis of the candidates array
                diff = float(self.max_ft_day[i][3]) - 12
                diff_sum += diff
                if i != len(self.max_ft_day) - 1:  # if the loop has not reached the end of the list
                    # if the name current and next name are the same
                    if self.max_ft_day[i][0] == self.max_ft_day[i + 1][0] and \
                            self.max_ft_day[i][1] == self.max_ft_day[i + 1][1]:
                        jump = "yes"  # bypasses an analysis of the candidates array
                        tabs = 24 - (len(self.max_ft_day[i][0]))
                        period = "."
                        period += tabs * "."
                        self.report.write(self.day_xlr[self.max_ft_day[i][2]] + "   " +
                                          self.max_ft_day[i][0] + ", " +
                                          self.max_ft_day[i][1] +
                                          period + "{0:.2f}".format(
                            float(self.max_ft_day[i][3])) + "   " + "{0:.2f}".format(float(diff)) + "\n")
                if jump == "no":
                    tabs = 24 - (len(self.max_ft_day[i][0]))
                    period = "."
                    period += tabs * "."
                    self.report.write(self.day_xlr[self.max_ft_day[i][2]] + "   " +
                                      self.max_ft_day[i][0] + ", " +
                                      self.max_ft_day[i][1] + period +
                                      "{0:.2f}".format(float(self.max_ft_day[i][3])) + "   " +
                                      "{0:.2f}".format(float(diff)) + "   " +
                                      "{0:.2f}".format(float(diff_sum)) + "\n")
                    dmax_add = [self.max_ft_day[i][0], self.max_ft_day[i][1], diff_sum]
                    daily_max.append(dmax_add)  # catch sum of daily violations for the week
                    sum_total += diff_sum
                    diff_sum = 0
            self.report.write("\n" + "                                         total:  " +
                              "{0:.2f}".format(float(sum_total))
                              + "\n")
            self.report.write("\n11.50 hour auxiliary carrier violations \n\n")
            self.report.write("day   name                        total   over   sum\n")
            self.report.write("-----------------------------------------------------\n")
            if len(self.max_aux_day) == 0:
                self.report.write("no violations" + "\n")
            diff_sum = 0
            sum_total = 0
            self.max_aux_day.sort(key=itemgetter(0))
            for i in range(len(self.max_aux_day)):
                jump = "no"  # triggers an analysis of the candidates array
                diff = float(self.max_aux_day[i][3]) - 11.5
                diff_sum += diff
                if i != len(self.max_aux_day) - 1:  # if the loop has not reached the end of the list
                    # if the current and next name are the same
                    if self.max_aux_day[i][0] == self.max_aux_day[i + 1][0] and \
                            self.max_aux_day[i][1] == self.max_aux_day[i + 1][1]:
                        jump = "yes"  # bypasses an analysis of the candidates array
                        tabs = 24 - (len(self.max_aux_day[i][0]))
                        period = "."
                        period += tabs * "."
                        self.report.write(self.day_xlr[self.max_aux_day[i][2]] + "   " +
                                          self.max_aux_day[i][0] + ", " +
                                          self.max_aux_day[i][1] + period +
                                          "{0:.2f}".format(float(self.max_aux_day[i][3])) + "   " +
                                          "{0:.2f}".format(float(diff)) + "\n")
                if jump == "no":
                    tabs = 24 - (len(self.max_aux_day[i][0]))
                    period = "."
                    period += tabs * "."
                    self.report.write(self.day_xlr[self.max_aux_day[i][2]] + "   " +
                                      self.max_aux_day[i][0] + ", " +
                                      self.max_aux_day[i][1] + period +
                                      "{0:.2f}".format(float(self.max_aux_day[i][3])) + "   " +
                                      "{0:.2f}".format(float(diff)) + "   " +
                                      "{0:.2f}".format(float(diff_sum)) + "\n")
                    dmax_add = [self.max_aux_day[i][0], self.max_aux_day[i][1], diff_sum]
                    daily_max.append(dmax_add)  # catch sum of daily violations for the week
                    sum_total += diff_sum
                    diff_sum = 0
            self.report.write(
                "\n" + "                                         total:  " +
                "{0:.2f}".format(float(sum_total)) + "\n")
            weekly_and_daily = []
            d_max_remove = []
            w_max_remove = []
            # find the write the adjustments
            # get the adjustment
            self.adjustment.sort(key=itemgetter(1))
            adj_sum = 0
            adj_total = []
            self.report.write("\nPost 60 Hour Adjustments \n\n")
            self.report.write("day   name                   daily adj    total\n")
            self.report.write("-----------------------------------------------\n")
            if len(self.adjustment) == 0:
                self.report.write("no adjustments" + "\n")
            for i in range(len(self.adjustment)):
                jump = "no"  # triggers an analysis of the adjustment array
                adj_sum += self.adjustment[i][3]
                if i != len(self.adjustment) - 1:  # if the loop has not reached the end of the list
                    # if the current and next name are the same
                    if self.adjustment[i][1] == self.adjustment[i + 1][1] and \
                            self.adjustment[i][2] == self.adjustment[i + 1][2]:
                        jump = "yes"  # bypasses an analysis of the candidates array
                        tabs = 24 - (len(self.adjustment[i][1]))
                        period = "."
                        period += tabs * "."
                        self.report.write(self.adjustment[i][0] + "   " + self.adjustment[i][1] + ", " +
                                          self.adjustment[i][2] + period + "{0:.2f}"
                                          .format(float(self.adjustment[i][3])) + "\n")
                if jump == "no":
                    tabs = 24 - (len(self.adjustment[i][1]))
                    period = "."
                    period += tabs * "."
                    self.report.write(self.adjustment[i][0] + "   " + self.adjustment[i][1] + ", "
                                      + self.adjustment[i][2] + period + "{0:.2f}".format(float(self.adjustment[i][3]))
                                      + "     " + "{0:.2f}".format(float(adj_sum)) + "\n")
                    adj_add = [self.adjustment[i][1], self.adjustment[i][2], adj_sum]
                    adj_sum = 0
                    adj_total.append(adj_add)  # catch sum of adjustments for the week
            for w_max in weekly_max:  # find the total violation
                for d_max in daily_max:
                    # look for names with both weekly and daily violations
                    if w_max[0] + w_max[1] == d_max[0] + d_max[1]:
                        wk_dy_sum = w_max[2] + d_max[2]  # add the weekly and daily
                        to_add = [w_max[0], w_max[1], wk_dy_sum]
                        weekly_and_daily.append(to_add)
                        d_max_remove.append(d_max)
                        w_max_remove.append(w_max)
            weekly_max = [x for x in weekly_max if x not in w_max_remove]
            daily_max = [x for x in daily_max if x not in d_max_remove]
            d_max_remove = []
            w_max_remove = []
            for d_max in daily_max:
                for w_max in weekly_max:
                    if w_max[0] + w_max[1] == d_max[0] + d_max[1]:  # if the names match
                        wk_dy_sum = w_max[2] + d_max[2]  # add the weekly and daily
                        to_add = [w_max[0], w_max[1], wk_dy_sum]
                        weekly_and_daily.append(to_add)
                        d_max_remove.append(d_max)
                        w_max_remove.append(w_max)
            weekly_max = [x for x in weekly_max if x not in w_max_remove]  # remove
            daily_max = [x for x in daily_max if x not in d_max_remove]
            joint_max = (weekly_max + daily_max + weekly_and_daily)  # add all arrays to get the final array

            joint_max.sort(key=itemgetter(0, 1))
            for j in joint_max:  # cycle through the totals and adjustments
                for a in adj_total:
                    if j[0] + j[1] == a[0] + a[1]:  # if the names match
                        j[2] -= a[2]  # subtract the adjustment from the total
            self.report.write("\n\nTotal of the two violations (with adjustments)\n\n")
            self.report.write("name                              total\n")
            self.report.write("---------------------------------------\n")
            if len(joint_max) == 0:
                self.report.write("no violations" + "\n")
            great_total = 0
            for item in joint_max:
                tabs = 30 - (len(item[0]))
                period = "."
                period += tabs * "."
                great_total += item[2]
                # self.report.write(item[0] + ", " + item[1] + period +
                #                   "{0:.2f}".format(float(item[2])).rjust(5, ".") + "\n")
                str_format = "{0:.2f}".format(float(item[2])).rjust(5, ".")
                self.report.write("{}, {}{}{}\n".format(item[0], item[1], period, str_format))
            self.report.write("\n" + "                           total:  " +
                              "{0:.2f}".format(float(great_total)) + "\n")

    def open(self):
        """ open the file.  """
        self.target_file.close()
        self.csv_fix.destroy()
        self.report.close()
        try:
            if sys.platform == "win32":
                os.startfile(dir_path('over_max') + self.filename)
            if sys.platform == "linux":
                subprocess.call(["xdg-open", 'kb_sub/over_max/' + self.filename])
            if sys.platform == "darwin":
                subprocess.call(["open", dir_path('over_max') + self.filename])
        except PermissionError:
            messagebox.showerror("Report Generator",
                                 "The report was not generated.",
                                 parent=self.frame)


def ee_skimmer(frame):
    """ This the function called for the csv everything report reader """
    days = ("Saturday", "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday")
    mv_codes = ("BT", "MV", "ET")
    carrier = []
    path = dir_filedialog()
    file_path = filedialog.askopenfilename(initialdir=path, filetypes=[("Excel files", "*.csv *.xls")])
    if file_path[-4:].lower() == ".csv" or file_path[-4:].lower() == ".xls":
        pass
    else:
        messagebox.showerror("Report Generator",
                             "The file you have selected is not a .csv or .xls file.\n"
                             "You must select a file with a .csv or .xls extension.",
                             parent=frame)
        return
    csv_fix = CsvRepair()  # create a CsvRepair object
    # returns a file path for a checked and, if needed, fixed csv file.
    file_path = csv_fix.run(file_path)
    target_file = open(file_path, newline="")
    with target_file as file:
        a_file = csv.reader(file)
        cc = 0
        good_id = "no"
        for line in a_file:
            # print(cc, line)
            if cc == 0:
                if line[0][:8] != "TAC500R3":
                    messagebox.showwarning("File Selection Error",
                                           "The selected file does not appear to be an "
                                           "Employee Everything report.",
                                           parent=frame)
                    target_file.close()  # close the opened file which is no longer being read
                    csv_fix.destroy()  # destroy the CsvRepair object and the proxy csv file
                    return
            if cc == 2:
                pp = line[0]  # find the pay period
                filename = "ee_reader" + "_" + pp + ".txt"
                try:
                    report = open(dir_path('ee_reader') + filename, "w")
                except (PermissionError, FileNotFoundError):
                    messagebox.showwarning("Report Generator",
                                           "The Employee Everything Report Reader "
                                           "was not generated.",
                                           parent=frame)
                    target_file.close()  # close the opened file which is no longer being read
                    csv_fix.destroy()  # destroy the CsvRepair object and the proxy csv file
                    return
                report.write("\nEmployee Everything Report Reader\n")
                report.write(
                    "pay period: " + pp[:-3] + " " + pp[4] + pp[5] + "-" + pp[6] + "\n\n")  # printe pay period
            if cc != 0:
                if good_id != line[4] and good_id != "no":  # if new carrier or employee
                    ee_analysis(carrier, report)  # trigger analysis
                    del carrier[:]  # empty array
                    good_id = "no"  # reset trigger
                # find first line of specific carrier
                if line[18] == "Base" and line[19] in ("844", "134", "434"):
                    good_id = line[4]  # set trigger to id of carriers who are FT or aux carriers
                    carrier.append(line)  # gather times and moves for anaylsis
                if good_id == line[4] and line[18] != "Base":
                    if line[18] in days:  # get the hours for each day
                        carrier.append(line)  # gather times and moves for anaylsis
                    if line[19] in mv_codes and line[32] != "(W)Ring Deleted From PC":
                        carrier.append(line)  # gather times and moves for anaylsis
            cc += 1
        ee_analysis(carrier, report)  # when loop ends, run final analysis
        del carrier[:]  # empty array
        report.close()
        if sys.platform == "win32":
            os.startfile(dir_path('ee_reader') + filename)
        if sys.platform == "linux":
            subprocess.call(["xdg-open", 'kb_sub/ee_reader/' + filename])
        if sys.platform == "darwin":
            subprocess.call(["open", dir_path('ee_reader') + filename])
    target_file.close()  # close the opened file which is no longer being read
    csv_fix.destroy()  # destroy the CsvRepair object and the proxy csv file


def ee_analysis(array, report):
    """ This is called by the ee_skimmer """
    listt = None
    ns_day = None
    route = None
    days = ("Saturday", "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday")
    hr_codes = ("52", "55", "56", "59", "60")
    code_dict = {"52": "total ", "55": "annual", "56": "sick  ", "59": "lwop  ", "60": "lwop  "}
    mv_codes = ("BT", "MV", "ET")
    moves_array = []
    for line in array:
        if line[19] and line[19] not in mv_codes and len(moves_array) > 0:
            find_move_sets(moves_array)  # call function to analyse moves
            del moves_array[:]
        # find first line of specific carrier
        if line[18] == "Base" and line[19] == "844" \
                or line[18] == "Base" and line[19] == "134" \
                or line[18] == "Base" and line[19] == "434":
            if line[19] == "844":
                listt = "aux"
                route = ""
                ns_day = ""
            elif line[19] == "434":
                listt = "ptf"
                route = ""
                ns_day = ""
            else:
                listt = "FT"
                ns_day = ee_ns_detect(array)  # call function to find the ns day
                if line[23].zfill(2) == "01":
                    route = line[25].zfill(6)
                    route = route[1] + route[2] + route[4] + route[5]
                    route = Handler(route).routes_adj()
                if line[23].zfill(2) == "02":
                    route = "floater"
            report.write("================================================\n")
            report.write(line[5].lower() + ", " + line[6].lower() + "\n")  # write name
            report.write(listt + "\n")
            if listt == "FT":
                report.write("route:" + route + "\n")
                if ns_day is None:
                    report.write("Klusterbox failed to detect ns day!")
                else:
                    report.write("ns day:" + ns_day + "\n")
            # report.write("================================================\n")
        if line[18] in days:
            spt_20 = line[20].split(':')  # split to get code and hours
            hr_type = spt_20[0][1] + spt_20[0][2]  # parse hour code to 2 digits
            if hr_type in hr_codes:  # compare to array of hour codes
                report.write("------------------------------------------------\n")
                if line[18] == ns_day:  # if the day is the ns day...
                    report.write("{}{}{}{}\n".format(line[18].ljust(12, " "), code_dict[hr_type].ljust(10, " "),
                                                     "{0:.2f}".format(float(spt_20[1])).ljust(6, " "),
                                                     "ns day".rjust(17, " ")))
                else:  # if the day is NOT the ns day...
                    report.write("{}{}{}\n".format(line[18].ljust(12, " "), code_dict[hr_type].ljust(10, " "),
                                                   "{0:.2f}".format(float(spt_20[1])).ljust(6, " ")))
                # report.write("------------------------------------------------\n")
        if line[19] in mv_codes and line[32] != "(W)Ring Deleted From PC":  # printe rings
            r_route = line[24].zfill(6)
            r_route = r_route[1] + r_route[2] + r_route[4] + r_route[5]  # reformat route to 4 digit format
            if route != r_route and listt == "FT" and route != "floater" and r_route != "0000":
                off_route = "off"  # marker for off route work
            else:
                off_route = ""  # no marker for off route work
            # make array and call function to makes moves sets
            mv_data = (line[19], float(line[21]), move_translator(line[23][:-4]), off_route)
            moves_array.append(mv_data)
            report.write(
                "\t{}{}{}{}{}\n".format(line[19].ljust(2, " "), "{00:.2f}".format(float(line[21])).rjust(8, " "),
                                        move_translator(line[23][:-4]).rjust(12, " "), r_route.rjust(6, " "),
                                        off_route.rjust(6, " ")))
    if len(moves_array) > 0:
        # call function to analyse moves
        find_move_sets(moves_array)
        del moves_array[:]


def move_translator(num):
    """ makes 721, 722 codes readable. """
    move_xlr = {"721": "to office", "722": "to street", "354": "standby", "622": "to travel", "613": "steward"}
    if num in move_xlr:  # if the code is in the dictionary...
        return move_xlr[num]  # translate it
    else:  # if the code is not in the dictionary...
        return num  # just return the code

    
def find_move_sets(moves):
    """
    this finds the triads of moves off and back on the carrier's route.
    """
    mv_sets = []
    pair = "closed"
    for line in moves:
        if line[3] == "off" and pair == "closed":
            mv_sets.append(line[1])
            pair = "open"
        if pair == "open":
            if line[3] == "":
                mv_sets.append(line[1])
                pair = "closed"
