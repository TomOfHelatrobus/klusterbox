from kbtoolbox import *
import csv
import os
from tkinter import filedialog
from kbcsv_repair import CsvRepair
from operator import itemgetter
# Spreadsheet Libraries
from openpyxl import load_workbook
from openpyxl import Workbook
from openpyxl.styles import NamedStyle, Font, Border, Side, Alignment, PatternFill


def max_hr(frame):  # generates a report for 12/60 hour violations
    path = dir_filedialog()
    file_path = filedialog.askopenfilename(initialdir=path, filetypes=[("Excel files", "*.csv *.xls")])
    csv_fix = CsvRepair()  # create a CsvRepair object
    # returns a file path for a checked and, if needed, fixed csv file.
    file_path = csv_fix.run(file_path)
    day_xlr = {"Saturday": "sat", "Sunday": "sun", "Monday": "mon", "Tuesday": "tue", "Wednesday": "wed",
               "Thursday": "thr", "Friday": "fri"}
    leave_xlr = {"49": "owcp   ", "55": "annual ", "56": "sick   ", "58": "holiday", "59": "lwop   ", "60": "lwop   "}
    maxhour = []
    max_aux_day = []
    max_ft_day = []
    extra_hours = []
    all_extra = []
    adjustment = []
    target_file = None
    pp = ""
    ft = ""
    days = ("Saturday", "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday")
    day_hours = []
    if file_path[-4:].lower() == ".csv" or file_path[-4:].lower() == ".xls":
        target_file = open(file_path, newline="")
        a_file = csv.reader(target_file)
        # with open(file_path, newline="") as file:
        #     a_file = csv.reader(file)
        cc = 0
        good_id = "no"
        for line in a_file:
            if cc == 0:
                if line[0][:8] != "TAC500R3":
                    messagebox.showwarning("File Selection Error",
                                           "The selected file does not appear to be an "
                                           "Employee Everything report.",
                                           parent=frame)
                    target_file.close()
                    csv_fix.destroy()
                    return
            if cc == 2:  # on the second line
                pp = line[0]  # find the pay period
                pp = pp.strip()  # strip whitespace out of pay period information
            if cc != 0:  # on all but the first line
                if line[18] == "Base" and good_id and len(day_hours) > 0:
                    # find fri hours for friday adjustment
                    fri_hrs = 0
                    for t in day_hours:  # get the friday hours
                        if t[3] == "Friday":
                            fri_hrs += float(t[2])
                    # find thu hours for thursday adjustment
                    thu_hrs = 0
                    for t in day_hours:  # find the thursday hours
                        if t[3] == "Thursday":
                            thu_hrs += float(t[2])
                    # find wed hours for wednesday adjustment
                    wed_hrs = 0
                    for t in day_hours:  # find the wednesday hours
                        if t[3] == "Wednesday":
                            wed_hrs += float(t[2])
                    # find the weekly total by adding daily totals
                    wkly_total = 0
                    for t in day_hours:
                        wkly_total += float(t[2])
                    if wkly_total > 60:
                        add_maxhr = (day_hours[0][0].lower(), day_hours[0][1].lower(), wkly_total)
                        maxhour.append(add_maxhr)
                        for item in extra_hours:  # get any extra hours codes for non-5200 hours list
                            all_extra.append(item)
                        # find the all adjustments
                        if ft:
                            # find friday adjustment
                            fri_post_60 = float(wkly_total - 60)
                            if fri_hrs > 12:
                                fri_over = fri_hrs - 12
                                if fri_over < fri_post_60:
                                    fri_adj = fri_over
                                else:
                                    fri_adj = fri_post_60
                                add_adjustment = ("fri", day_hours[0][0].lower(), day_hours[0][1].lower(), fri_adj)
                                adjustment.append(add_adjustment)
                            # find the thursday adjustment
                            thu_post_60 = float(wkly_total - 60) - fri_hrs
                            if thu_hrs > 12 and thu_post_60 > 0:
                                thu_over = thu_hrs - 12
                                if thu_over < thu_post_60:
                                    thu_adj = thu_over
                                else:
                                    thu_adj = thu_post_60
                                add_adjustment = ("thu", day_hours[0][0].lower(), day_hours[0][1].lower(), thu_adj)
                                adjustment.append(add_adjustment)
                            # find the wednesday adjustment
                            wed_post_60 = float(wkly_total - 60) - fri_hrs - thu_hrs
                            if wed_hrs > 12 and wed_post_60 > 0:
                                wed_over = wed_hrs - 12
                                if wed_over < wed_post_60:
                                    wed_adj = wed_over
                                else:
                                    wed_adj = wed_post_60
                                add_adjustment = (
                                    "wed", day_hours[0][0].lower(), day_hours[0][1].lower(), wed_adj)
                                adjustment.append(add_adjustment)
                    del day_hours[:]
                    del extra_hours[:]
                # find first line of specific carrier
                if line[18] == "Base" and line[19] in ("844", "134", "434"):
                    good_id = line[4]  # remember id of carriers who are FT or aux carriers
                    if line[19] in ("844", "434"):
                        ft = False
                    else:
                        ft = True
                if good_id == line[4] and line[18] != "Base":
                    if line[18] in days:  # get the hours for each day
                        spt_20 = line[20].split(':')  # split to get code and hours
                        hr_type = spt_20[0][1] + spt_20[0][2]  # parse hour code to 2 digits
                        # if hr_type in hr_codes:  # compare to array of hour codes
                        if hr_type == "52":  # compare to array of hour codes
                            if float(spt_20[1]) > 11.5 and not ft:
                                add_max_aux = (line[5].lower(), line[6].lower(), line[18], spt_20[1])
                                max_aux_day.append(add_max_aux)
                            if float(spt_20[1]) > 12 and ft:
                                add_max_ft = (line[5].lower(), line[6].lower(), line[18], spt_20[1])
                                max_ft_day.append(add_max_ft)
                            if ft:  # increment daily totals to find weekly total
                                add_day_hours = (line[5].lower(), line[6].lower(), spt_20[1], line[18])
                                day_hours.append(add_day_hours)
                        extra_hour_codes = ("49", "55", "56", "58")  # paid leave types only , (lwop "59", "60")
                        if hr_type in extra_hour_codes and ft:  # if there is holiday pay
                            add_day_hours = (line[5].lower(), line[6].lower(), spt_20[1], line[18])
                            day_hours.append(add_day_hours)
                            add_extra_hours = (line[5].lower(), line[6].lower(), line[18], hr_type, spt_20[1])
                            extra_hours.append(add_extra_hours)  # track non 5200 hours
            cc += 1
    elif file_path == "":
        target_file.close()
        csv_fix.destroy()
        return
    else:
        messagebox.showerror("Report Generator",
                             "The file you have selected is not a .csv or .xls file.\n"
                             "You must select a file with a .csv or .xls extension.",
                             parent=frame)
        target_file.close()
        csv_fix.destroy()
        return
    # find the weekly total by adding daily totals for last carrier
    if len(day_hours) > 0:
        wkly_total = 0
        for t in day_hours:
            wkly_total += float(t[2])
        if wkly_total > 60:
            add_maxhr = (day_hours[0][0].lower(), day_hours[0][1].lower(), wkly_total)
            maxhour.append(add_maxhr)
            for item in extra_hours:  # get any extra hours codes for non-5200 hours list
                all_extra.append(item)
        del day_hours[:]
        del extra_hours[:]

    if len(maxhour) == 0 and len(max_ft_day) == 0 and len(max_aux_day) == 0:
        messagebox.showwarning("Report Generator",
                               "No violations were found. "
                               "The report was not generated.",
                               parent=frame)
        target_file.close()
        csv_fix.destroy()
        return
    weekly_max = []  # array hold each carrier's hours for the week
    daily_max = []  # array hold each carrier's sum of maximum daily hours for the week
    if len(maxhour) > 0 or len(max_ft_day) > 0 or len(max_aux_day) > 0:
        pp_str = pp[:-3] + "_" + pp[4] + pp[5] + "_" + pp[6]
        filename = "max" + "_" + pp_str + ".txt"
        report = open(dir_path('over_max') + filename, "w")
        report.write("12 and 60 Hour Violations Report\n\n")
        report.write("pay period: " + pp[:-3] + " " + pp[4] + pp[5] + "-" + pp[6] + "\n")  # printe pay period
        pp_date = find_pp(int(pp[:-3]), pp[-3:])  # send year and pp to get the date
        pp_date_end = pp_date + timedelta(days=6)  # add six days to get the last part of the range
        report.write(
            "week of: " + pp_date.strftime("%x") + " - " + pp_date_end.strftime("%x") + "\n")  # printe date
        report.write("\n60 hour violations \n\n")
        report.write("name                              total   over\n")
        report.write("-----------------------------------------------\n")
        if len(maxhour) == 0:
            report.write("no violations" + "\n")
        else:
            diff_total = 0
            maxhour.sort(key=itemgetter(0))
            for item in maxhour:
                tabs = 30 - (len(item[0]))
                period = "."
                period = period + (tabs * ".")
                diff = float(item[2]) - 60
                diff_total = diff_total + diff
                report.write(item[0] + ", " + item[1] + period + "{0:.2f}".format(float(item[2]))
                             + "   " + "{0:.2f}".format(float(diff)).rjust(5, " ") + "\n")
                wmax_add = [item[0], item[1], diff]
                weekly_max.append(wmax_add)  # catch totals of violations for the week
            report.write("\n" + "                                   total:  " + "{0:.2f}".format(float(diff_total))
                         + "\n")
        all_extra.sort(key=itemgetter(0))
        report.write("\nNon 5200 codes contributing to 60 hour violations  \n\n")
        report.write("day   name                            hr type   hours\n")
        report.write("-----------------------------------------------------\n")
        if len(all_extra) == 0:
            report.write("no contributions" + "\n")
        for i in range(len(all_extra)):
            tabs = 28 - (len(all_extra[i][0]))
            period = "."
            period = period + (tabs * ".")
            report.write(day_xlr[all_extra[i][2]] + "   " + all_extra[i][0] + ", " + all_extra[i][1] + period +
                         leave_xlr[all_extra[i][3]] + "  " + "{0:.2f}".format(float(all_extra[i][4])).rjust(5, " ")
                         + "\n")
        report.write("\n\n12 hour full time carrier violations \n\n")
        report.write("day   name                        total   over   sum\n")
        report.write("-----------------------------------------------------\n")
        if len(max_ft_day) == 0:
            report.write("no violations" + "\n")
        diff_sum = 0
        sum_total = 0
        max_ft_day.sort(key=itemgetter(0))
        for i in range(len(max_ft_day)):
            jump = "no"  # triggers an analysis of the candidates array
            diff = float(max_ft_day[i][3]) - 12
            diff_sum = diff_sum + diff
            if i != len(max_ft_day) - 1:  # if the loop has not reached the end of the list
                # if the name current and next name are the same
                if max_ft_day[i][0] == max_ft_day[i + 1][0] and max_ft_day[i][1] == max_ft_day[i + 1][1]:
                    jump = "yes"  # bypasses an analysis of the candidates array
                    tabs = 24 - (len(max_ft_day[i][0]))
                    period = "."
                    period = period + (tabs * ".")
                    report.write(day_xlr[max_ft_day[i][2]] + "   " + max_ft_day[i][0] + ", " + max_ft_day[i][1] +
                                 period + "{0:.2f}".format(
                        float(max_ft_day[i][3])) + "   " + "{0:.2f}".format(float(diff)) + "\n")
            if jump == "no":
                tabs = 24 - (len(max_ft_day[i][0]))
                period = "."
                period = period + (tabs * ".")
                report.write(day_xlr[max_ft_day[i][2]] + "   " + max_ft_day[i][0] + ", " + max_ft_day[i][1] + period
                             + "{0:.2f}".format(float(max_ft_day[i][3])) + "   " + "{0:.2f}".format(float(diff)) +
                             "   " + "{0:.2f}".format(float(diff_sum)) + "\n")
                dmax_add = [max_ft_day[i][0], max_ft_day[i][1], diff_sum]
                daily_max.append(dmax_add)  # catch sum of daily violations for the week
                sum_total = sum_total + diff_sum
                diff_sum = 0
        report.write("\n" + "                                         total:  " + "{0:.2f}".format(float(sum_total))
                     + "\n")
        report.write("\n11.50 hour auxiliary carrier violations \n\n")
        report.write("day   name                        total   over   sum\n")
        report.write("-----------------------------------------------------\n")
        if len(max_aux_day) == 0:
            report.write("no violations" + "\n")
        diff_sum = 0
        sum_total = 0
        max_aux_day.sort(key=itemgetter(0))
        for i in range(len(max_aux_day)):
            jump = "no"  # triggers an analysis of the candidates array
            diff = float(max_aux_day[i][3]) - 11.5
            diff_sum = diff_sum + diff
            if i != len(max_aux_day) - 1:  # if the loop has not reached the end of the list
                # if the current and next name are the same
                if max_aux_day[i][0] == max_aux_day[i + 1][0] and max_aux_day[i][1] == max_aux_day[i + 1][1]:
                    jump = "yes"  # bypasses an analysis of the candidates array
                    tabs = 24 - (len(max_aux_day[i][0]))
                    period = "."
                    period = period + (tabs * ".")
                    report.write(day_xlr[max_aux_day[i][2]] + "   " + max_aux_day[i][0] + ", "
                                 + max_aux_day[i][1] + period + "{0:.2f}".format(float(max_aux_day[i][3]))
                                 + "   " + "{0:.2f}".format(float(diff)) + "\n")
            if jump == "no":
                tabs = 24 - (len(max_aux_day[i][0]))
                period = "."
                period = period + (tabs * ".")
                report.write(day_xlr[max_aux_day[i][2]] + "   " + max_aux_day[i][0] + ", "
                             + max_aux_day[i][1] + period + "{0:.2f}".format(float(max_aux_day[i][3]))
                             + "   " + "{0:.2f}".format(float(diff)) + "   " + "{0:.2f}".format(float(diff_sum))
                             + "\n")
                dmax_add = [max_aux_day[i][0], max_aux_day[i][1], diff_sum]
                daily_max.append(dmax_add)  # catch sum of daily violations for the week
                sum_total = sum_total + diff_sum
                diff_sum = 0
        report.write(
            "\n" + "                                         total:  " + "{0:.2f}".format(float(sum_total)) + "\n")
        weekly_and_daily = []
        d_max_remove = []
        w_max_remove = []
        # find the write the adjustments
        # get the adjustment
        adjustment.sort(key=itemgetter(1))
        adj_sum = 0
        adj_total = []
        report.write("\nPost 60 Hour Adjustments \n\n")
        report.write("day   name                   daily adj    total\n")
        report.write("-----------------------------------------------\n")
        if len(adjustment) == 0:
            report.write("no adjustments" + "\n")
        for i in range(len(adjustment)):
            jump = "no"  # triggers an analysis of the adjustment array
            adj_sum = adj_sum + adjustment[i][3]
            if i != len(adjustment) - 1:  # if the loop has not reached the end of the list
                # if the current and next name are the same
                if adjustment[i][1] == adjustment[i + 1][1] and adjustment[i][2] == adjustment[i + 1][2]:
                    jump = "yes"  # bypasses an analysis of the candidates array
                    tabs = 24 - (len(adjustment[i][1]))
                    period = "."
                    period = period + (tabs * ".")
                    report.write(adjustment[i][0] + "   " + adjustment[i][1] + ", "
                                 + adjustment[i][2] + period + "{0:.2f}".format(float(adjustment[i][3])) + "\n")
            if jump == "no":
                tabs = 24 - (len(adjustment[i][1]))
                period = "."
                period = period + (tabs * ".")
                report.write(adjustment[i][0] + "   " + adjustment[i][1] + ", "
                             + adjustment[i][2] + period + "{0:.2f}".format(float(adjustment[i][3]))
                             + "     " + "{0:.2f}".format(float(adj_sum))
                             + "\n")
                adj_add = [adjustment[i][1], adjustment[i][2], adj_sum]
                adj_sum = 0
                adj_total.append(adj_add)  # catch sum of adjustments for the week
        for w_max in weekly_max:  # find the total violation
            for d_max in daily_max:
                if w_max[0] + w_max[1] == d_max[0] + d_max[1]:  # look for names with both weekly and daily violations
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
                    j[2] = j[2] - a[2]  # subtract the adjustment from the total
        report.write("\n\nTotal of the two violations (with adjustments)\n\n")
        report.write("name                              total\n")
        report.write("---------------------------------------\n")
        if len(joint_max) == 0:
            report.write("no violations" + "\n")
        great_total = 0
        for item in joint_max:
            tabs = 30 - (len(item[0]))
            period = "."
            period = period + (tabs * ".")
            great_total = great_total + item[2]
            report.write(item[0] + ", " + item[1] + period + "{0:.2f}".format(float(item[2])).rjust(5, ".") + "\n")
        report.write(
            "\n" + "                           total:  " + "{0:.2f}".format(float(great_total)) + "\n")
        report.close()
        try:
            if sys.platform == "win32":
                os.startfile(dir_path('over_max') + filename)
            if sys.platform == "linux":
                subprocess.call(["xdg-open", 'kb_sub/over_max/' + filename])
            if sys.platform == "darwin":
                subprocess.call(["open", dir_path('over_max') + filename])
        except PermissionError:
            messagebox.showerror("Report Generator",
                                 "The report was not generated.",
                                 parent=frame)
    target_file.close()
    csv_fix.destroy()


def ee_skimmer(frame):
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


def ee_ns_detect(array):  # finds the ns day from ee reports
    days = ("Saturday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday")
    ns_candidates = ["Saturday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
    for d in days:
        hr_52 = 0  # straight hours
        hr_53 = 0  # overtime hours
        hr_43 = 0  # penalty hours
        for line in array:
            if line[18] in ns_candidates:
                ns_candidates.remove(line[18])
            if line[18] == d:
                spt_20 = line[20].split(':')  # split to get code and hours
                if spt_20[0] == "05200":
                    hr_52 = spt_20[1]
                if spt_20[0] == "05300":
                    hr_53 = spt_20[1]
                if spt_20[0] == "04300":
                    hr_43 = spt_20[1]
        if float(hr_52) != 0:
            summ = float(hr_53) + float(hr_43)
            if float(hr_52) == round(summ, 2):
                return d
    if len(ns_candidates) == 1:
        return ns_candidates[0]


def move_translator(num):  # makes 721, 722 codes readable.
    move_xlr = {"721": "to office", "722": "to street", "354": "standby", "622": "to travel", "613": "steward"}
    if num in move_xlr:  # if the code is in the dictionary...
        return move_xlr[num]  # translate it
    else:  # if the code is not in the dictionary...
        return num  # just return the code

    
def find_move_sets(moves):
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


def wkly_avail(frame):  # creates a spreadsheet which shows weekly otdl availability
    tacs_pp = None
    tacs_station = None
    path = dir_filedialog()
    file_path = filedialog.askopenfilename(initialdir=path, filetypes=[("Excel files", "*.csv *.xls")])
    if file_path[-4:].lower() == ".csv" or file_path[-4:].lower() == ".xls":
        pass
    else:
        messagebox.showerror("Report Generator",
                             "The file you have selected is not a .csv or .xls file.\n"
                             "You must select a file with a .csv or .xls extension.",
                             parent=frame)
        return False
    csv_fix = CsvRepair()  # create a CsvRepair object
    # returns a file path for a checked and, if needed, fixed csv file.
    file_path = csv_fix.run(file_path)
    target_file = open(file_path, newline="")
    a_file = csv.reader(target_file)
    cc = 0
    for line in a_file:
        if cc == 0 and line[0][:8] != "TAC500R3":
            messagebox.showwarning("File Selection Error",
                                   "The selected file does not appear to be an "
                                   "Employee Everything report.",
                                   parent=frame)
            target_file.close()
            csv_fix.destroy()
            return False
        if cc == 3:
            tacs_pp = line[0]  # find the pay period
            tacs_station = line[2]  # find the station
            break
        cc += 1
    cc = 0
    range_days = ["Saturday", "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
    for line in a_file:  # find the range
        if line[18] in range_days:
            range_days.remove(line[18])
        if cc == 150:
            break  # survey 150 lines before breaking to anaylize results.
        cc += 1
    if len(range_days) > 5:
        messagebox.showwarning("File Selection Error",
                               "Employee Everything Reports that cover only one day /n"
                               "are not supported in version {} of Klusterbox.".format(version),
                               parent=frame)
        target_file.close()
        csv_fix.destroy()
        return False
    else:
        t_range = True
    year = int(tacs_pp[:-3])  # set the globals
    pp = tacs_pp[-3:]
    t_date = find_pp(year, pp)  # returns the starting date of the pp when given year and pay period
    s_year = t_date.strftime("%Y")
    s_mo = t_date.strftime("%m")
    s_day = t_date.strftime("%d")
    sql = "SELECT kb_station FROM station_index WHERE tacs_station = '%s'" % tacs_station
    station = inquire(sql)  # check to see if station has match in station index
    if not station:
        messagebox.showwarning("Error",
                               "This station has not been matched with Auto Data Entry.",
                               parent=frame)
        target_file.close()
        csv_fix.destroy()
        return False
    set_globals(s_year, s_mo, s_day, t_range, station[0][0], "None")  # set the investigation range
    # get the otdl list from the carriers table
    sql = "SELECT carrier_name FROM carriers WHERE effective_date <= '%s' and station = '%s' and list_status = '%s'" \
          "ORDER BY carrier_name, effective_date desc" % (projvar.invran_date_week[6], projvar.invran_station, 'otdl')
    results = inquire(sql)  # call function to access database
    unique_carriers = []  # create non repeating list of otdl carriers
    for name in results:
        if name[0] not in unique_carriers:
            unique_carriers.append(name[0])
    wkly_list = []  # initialize arrays for data sorting
    otdl_list = []  # pull info from ee for these carriers
    on_list = "no"
    station_anchor = "no"
    for name in unique_carriers:
        ot_wkly = []
        sql = "SELECT emp_id FROM name_index WHERE kb_name='%s'" % name
        results = inquire(sql)
        if results:  # record emp id to otdl carrier info
            ot_wkly.append(results[0][0])
        else:  # mark otdl carriers who don't have emp id available
            ot_wkly.append("no index")
        sql = "SELECT effective_date,list_status,station FROM carriers " \
              "WHERE carrier_name='%s' and effective_date<='%s'" \
              "ORDER BY effective_date desc" % (name, projvar.invran_date_week[6])
        results = inquire(sql)
        ot_wkly.append(name)
        for date in projvar.invran_date_week:  # loop for each day of the week
            for rec in results:  # loop for each record starting from the latest
                if rec[2] == projvar.invran_station:  # if there is a station match
                    station_anchor = "yes"  # mark the carrier as attached to station
                if datetime.strptime(rec[0],
                                     '%Y-%m-%d %H:%M:%S') <= date:  # if the rec is at or earlier than investigation.
                    if rec[1] == "otdl":  # note whether otdl or not.
                        ot_wkly.append("otdl")
                        on_list = "yes"
                    else:
                        ot_wkly.append("")
                    break  # stop. we only want the first
        if on_list == "yes" and station_anchor == "yes":
            wkly_list.append(ot_wkly)  # fill in array with carrier and otdl data
            otdl_list.append(ot_wkly[0])  # add to list of carriers who will be researched
        on_list = "no"  # reset
        station_anchor = "no"  # reset
    not_indexed = []
    for name in wkly_list:  # check to see if there are any otdl carriers who do not have a rec in name index
        if name[0] == "no index":
            not_indexed.append(name[1])  # add any names who do not into an array
    if len(not_indexed) != 0:  # message box info that some otdl do not have a record in the name index
        messagebox.showwarning("Missing Data",
                               "There are {} name/s which have not been matched with their employee id."
                               " Please exit and run the Auto Data Entry Feature to ensure that all carriers have "
                               " employee ids entered into Klusterbox.".format(len(not_indexed)),
                               parent=frame)
    if len(otdl_list) == 0:
        messagebox.showwarning("Empty OTDL",
                               "Klusterbox has no records of any otdl carriers for {} station "
                               "for the week of {}. This could mean that: \n1. The carrier list is empty. Run the "
                               "Automatic Data Entry Feature, selecting the Employee Everything Report you used here "
                               " to remedy this. You do not have to enter the rings data at the final step "
                               " \n2. The Name Index which matches the carrier name to the employee id "
                               "empty. As in #1, run the Automatic Data Entry Feature to fix this.\n3. "
                               "The carrier list has no otdl carriers "
                               "designated. Use the Multi Input Feature to designate otdl carriers. \n"
                               "This Weekly Availability Report can not be generated without a list of otdl carriers. "
                               "Build the carrier list/otdl before re-running Weekly Availability."
                               .format(projvar.invran_station, projvar.invran_date_week[0].strftime("%b %d, %Y")),
                               parent=frame)
        target_file.close()
        csv_fix.destroy()
        return True
    else:  # if there is an otdl then build array holding hours for each day
        days = ("Saturday", "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday")
        extra_hour_codes = ("49", "52", "55", "56", "57", "58", "59", "60")
        running_total = 0
        target_file = open(file_path, newline="")
        with target_file as file:
            a_file = csv.reader(file)
            cc = 0
            all_otdl = []
            good_id = "no"
            day_over = "empty"
            long_day = ["Saturday", "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
            sat = 0
            sun = 0
            mon = 0
            tue = 0
            wed = 0
            thr = 0
            fri = 0
            day_run = [sat, sun, mon, tue, wed, thr, fri]
            for line in a_file:
                """ since the CsvRepair creates an empty line this try statment skips it."""
                try:
                    if line[4]:
                        pass
                except IndexError:
                    continue
                if cc != 0 and line[4].zfill(8) in otdl_list:  # if the emp_id matches ones we are looking for
                    if line[18] == "Base" and good_id != "no":
                        sql = "SELECT kb_name FROM name_index WHERE emp_id='%s'" % good_id
                        result = inquire(sql)  # get the kb name with the emp id
                        all_day_run = []
                        for i in range(7):
                            all_day_run.append(day_run[i])
                        all_day_run.append(day_over)
                        # to_add = ([result[0][0]] + all_day_run + [day_over])
                        to_add = ([result[0][0]] + all_day_run)
                        all_otdl.append(to_add)
                        for i in range(len(long_day)):
                            day_run[i] = 0  # empty each day in day run
                        day_over = "empty"  # reset
                        running_total = 0  # reset
                    # find first line of specific carrier
                    if line[18] == "Base" and line[19] in ("844", "134", "434"):
                        good_id = line[4].zfill(8)  # remember id of carriers who are FT or aux carriers
                    if good_id == line[4].zfill(8) and line[18] != "Base":
                        if line[18] in days:  # get the hours for each day
                            spt_20 = line[20].split(':')  # split to get code and hours
                            hr_type = spt_20[0][1] + spt_20[0][2]  # parse hour code to 2 digits
                            if hr_type in extra_hour_codes:  # if hr_type in hr_codes:
                                running_total += float(spt_20[1])
                                i = 0
                                for ld in long_day:
                                    if ld == line[18]:
                                        day_run[i] += float(spt_20[1])
                                    i += 1
                            if day_over == "empty" and running_total > 60:
                                day_over = line[18]
                cc += 1
            # add to the all_otdl for the final carrier after the last line of the file is read
            if good_id != "no":
                sql = "SELECT kb_name FROM name_index WHERE emp_id='%s'" % good_id
                result = inquire(sql)  # get the kb name with the emp id
                all_day_run = []  # gets the total hours for each day
                for i in range(7):
                    all_day_run.append(day_run[i])
                all_day_run.append(day_over)
                # to_add = ([result[0][0]] + all_day_run + [day_over])  # add name, daily totals, day over
                to_add = ([result[0][0]] + all_day_run)
                all_otdl.append(to_add)
        all_otdl.sort(key=itemgetter(0))  # sort the all otdl array by carrier name
        # define spreadsheet cell formats
        bd = Side(style='thin', color="80808080")  # defines borders
        ws_header = NamedStyle(name="ws_header", font=Font(bold=True, name='Arial', size=14))
        date_dov = NamedStyle(name="date_dov", font=Font(name='Arial', size=10))
        date_dov_title = NamedStyle(name="date_dov_title", font=Font(bold=True, name='Arial', size=10),
                                    alignment=Alignment(horizontal='right'))
        col_header = NamedStyle(name="col_header", font=Font(bold=True, name='Arial', size=10),
                                alignment=Alignment(horizontal='center'))
        col_name = NamedStyle(name="col_name", font=Font(bold=True, name='Arial', size=10),
                              alignment=Alignment(horizontal='left'))
        col_mod = NamedStyle(name="col_mod", font=Font(bold=True, name='Arial', size=10),
                             alignment=Alignment(horizontal='center'),
                             fill=PatternFill(fgColor='FFFFE0', fill_type='solid'),
                             border=Border(left=bd, top=bd, right=bd, bottom=bd))
        input_name = NamedStyle(name="input_name", font=Font(name='Arial', size=10),
                                border=Border(left=bd, top=bd, right=bd, bottom=bd))
        input_s = NamedStyle(name="input_s", font=Font(name='Arial', size=10),
                             border=Border(left=bd, top=bd, right=bd, bottom=bd),
                             alignment=Alignment(horizontal='right'))
        calcs = NamedStyle(name="calcs", font=Font(name='Arial', size=10),
                           border=Border(left=bd, top=bd, right=bd, bottom=bd),
                           fill=PatternFill(fgColor='FFFFE0', fill_type='solid'),
                           alignment=Alignment(horizontal='right'))
        wb = Workbook()  # define the workbook
        wkly_total = wb.active  # create first worksheet
        wkly_total.title = "over_60"  # title first worksheet
        cell = wkly_total.cell(row=1, column=1)
        cell.value = "Weekly Availability Summary"
        cell.style = ws_header
        wkly_total.merge_cells('A1:E1')
        wkly_total['A3'] = "Date:  "  # create date/ pay period/ station header
        wkly_total['A3'].style = date_dov_title
        range_of_dates = format(projvar.invran_date_week[0], "%A  %m/%d/%y") + " - " + \
                         format(projvar.invran_date_week[6], "%A  %m/%d/%y")
        wkly_total['B3'] = range_of_dates
        wkly_total['B3'].style = date_dov
        wkly_total.merge_cells('B3:H3')
        date = datetime(int(projvar.invran_year), int(projvar.invran_month), int(projvar.invran_day))
        projvar.pay_period = pp_by_date(date)
        wkly_total['E4'] = "Pay Period:  "
        wkly_total['E4'].style = date_dov_title
        wkly_total.merge_cells('E4:F4')
        wkly_total['G4'] = projvar.pay_period
        wkly_total['G4'].style = date_dov
        wkly_total.merge_cells('G4:H4')
        wkly_total['A4'] = "Station:  "
        wkly_total['A4'].style = date_dov_title
        wkly_total['B4'] = projvar.invran_station
        wkly_total['B4'].style = date_dov
        wkly_total.merge_cells('B4:D4')
        oi = 6
        # column headers - first row
        wkly_total["A" + str(oi)] = "carrier name"  # carrier name
        wkly_total["B" + str(oi)] = "sat"
        wkly_total["C" + str(oi)] = "sun"
        wkly_total["D" + str(oi)] = "mon"
        wkly_total["E" + str(oi)] = "tue"
        wkly_total["F" + str(oi)] = "wed"
        wkly_total["G" + str(oi)] = "thr"
        wkly_total["H" + str(oi)] = "fri"
        wkly_total["I" + str(oi)] = "day over"  # the day of the violation
        # column headers - second row
        wkly_total["B" + str(oi + 1)] = "cumulative totals"
        wkly_total.merge_cells('B7:H7')
        wkly_total["I" + str(oi + 1)] = "to 60"  # the day of the violation
        # format headers
        wkly_total["A" + str(oi)].style = col_name
        wkly_total["B" + str(oi)].style = col_header
        wkly_total["C" + str(oi)].style = col_header
        wkly_total["D" + str(oi)].style = col_header
        wkly_total["E" + str(oi)].style = col_header
        wkly_total["F" + str(oi)].style = col_header
        wkly_total["G" + str(oi)].style = col_header
        wkly_total["H" + str(oi)].style = col_header
        wkly_total["I" + str(oi)].style = col_header
        wkly_total["B" + str(oi + 1)].style = col_mod
        wkly_total["I" + str(oi + 1)].style = col_mod
        # column widths
        wkly_total.column_dimensions["A"].width = 18
        wkly_total.column_dimensions["B"].width = 7
        wkly_total.column_dimensions["C"].width = 7
        wkly_total.column_dimensions["D"].width = 7
        wkly_total.column_dimensions["E"].width = 7
        wkly_total.column_dimensions["F"].width = 7
        wkly_total.column_dimensions["G"].width = 7
        wkly_total.column_dimensions["H"].width = 7
        wkly_total.column_dimensions["I"].width = 10
        oi += 2
        for otdl in all_otdl:
            # first of two rows
            wkly_total["A" + str(oi)] = otdl[0]  # carrier name
            wkly_total["B" + str(oi)] = otdl[1]
            wkly_total["C" + str(oi)] = otdl[2]
            wkly_total["D" + str(oi)] = otdl[3]
            wkly_total["E" + str(oi)] = otdl[4]
            wkly_total["F" + str(oi)] = otdl[5]
            wkly_total["G" + str(oi)] = otdl[6]
            wkly_total["H" + str(oi)] = otdl[7]
            if otdl[8] == "empty":  # handle "empty" violation days
                violation_day = ""
            else:
                violation_day = otdl[8]
            wkly_total["I" + str(oi)] = violation_day  # the day of the violation
            # format each cell with style
            wkly_total["A" + str(oi)].style = input_name
            wkly_total["B" + str(oi)].style = input_s
            wkly_total["C" + str(oi)].style = input_s
            wkly_total["D" + str(oi)].style = input_s
            wkly_total["E" + str(oi)].style = input_s
            wkly_total["F" + str(oi)].style = input_s
            wkly_total["G" + str(oi)].style = input_s
            wkly_total["H" + str(oi)].style = input_s
            wkly_total["I" + str(oi)].style = input_s
            # set number format for each cell
            wkly_total["B" + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
            wkly_total["C" + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
            wkly_total["D" + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
            wkly_total["E" + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
            wkly_total["F" + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
            wkly_total["G" + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
            wkly_total["H" + str(oi)].number_format = "#,###.00;[RED]-#,###.00"
            # second of two rows - incluces running totals
            formula = "=%s!B%s" % ('over_60', str(oi))
            wkly_total["B" + str(oi + 1)] = formula
            formula = "=SUM(%s!C%s+%s!B%s)" % ('over_60', str(oi), 'over_60', str(oi + 1))
            wkly_total["C" + str(oi + 1)] = formula
            formula = "=SUM(%s!D%s+%s!C%s)" % ('over_60', str(oi), 'over_60', str(oi + 1))
            wkly_total["D" + str(oi + 1)] = formula
            formula = "=SUM(%s!E%s+%s!D%s)" % ('over_60', str(oi), 'over_60', str(oi + 1))
            wkly_total["E" + str(oi + 1)] = formula
            formula = "=SUM(%s!F%s+%s!E%s)" % ('over_60', str(oi), 'over_60', str(oi + 1))
            wkly_total["F" + str(oi + 1)] = formula
            formula = "=SUM(%s!G%s+%s!F%s)" % ('over_60', str(oi), 'over_60', str(oi + 1))
            wkly_total["G" + str(oi + 1)] = formula
            formula = "=SUM(%s!H%s+%s!G%s)" % ('over_60', str(oi), 'over_60', str(oi + 1))
            wkly_total["H" + str(oi + 1)] = formula
            formula = "=MAX(60-%s!H%s,0)" % ('over_60', str(oi + 1))
            wkly_total["I" + str(oi + 1)] = formula
            # format each cell of the second row
            wkly_total["B" + str(oi + 1)].style = calcs
            wkly_total["C" + str(oi + 1)].style = calcs
            wkly_total["D" + str(oi + 1)].style = calcs
            wkly_total["E" + str(oi + 1)].style = calcs
            wkly_total["F" + str(oi + 1)].style = calcs
            wkly_total["G" + str(oi + 1)].style = calcs
            wkly_total["H" + str(oi + 1)].style = calcs
            wkly_total["I" + str(oi + 1)].style = calcs
            # set number format for each cell
            wkly_total["B" + str(oi + 1)].number_format = "#,###.00;[RED]-#,###.00"
            wkly_total["C" + str(oi + 1)].number_format = "#,###.00;[RED]-#,###.00"
            wkly_total["D" + str(oi + 1)].number_format = "#,###.00;[RED]-#,###.00"
            wkly_total["E" + str(oi + 1)].number_format = "#,###.00;[RED]-#,###.00"
            wkly_total["F" + str(oi + 1)].number_format = "#,###.00;[RED]-#,###.00"
            wkly_total["G" + str(oi + 1)].number_format = "#,###.00;[RED]-#,###.00"
            wkly_total["H" + str(oi + 1)].number_format = "#,###.00;[RED]-#,###.00"
            wkly_total["I" + str(oi + 1)].number_format = "#,###.00;[RED]-#,###.00"
            oi += 2
        if len(not_indexed) > 0:
            wkly_total["A" + str(oi)] = "Carriers not included (not in name index):"
            wkly_total.merge_cells('A' + str(oi) + ':D' + str(oi))
            oi += 1
            for name in not_indexed:
                wkly_total['A' + str(oi)] = name
                wkly_total.merge_cells('A' + str(oi) + ':D' + str(oi))
                oi += 1
        # name the excel file
        xl_filename = "kb_wa" + str(format(projvar.invran_date_week[0], "_%y_%m_%d")) + ".xlsx"
        ok = messagebox.askokcancel("Spreadsheet generator",
                                    "Do you want to generate a spreadsheet?",
                                    parent=frame)
        if ok:
            try:
                wb.save(dir_path('weekly_availability') + xl_filename)
                messagebox.showinfo("Spreadsheet generator",
                                    "Your spreadsheet was successfully generated. \n"
                                    "File is named: {}".format(xl_filename),
                                    parent=frame)
                if sys.platform == "win32":
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
                                     parent=frame)
        target_file.close()
        csv_fix.destroy()
        return True
