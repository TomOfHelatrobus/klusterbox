"""
a klusterbox module: Klusterbox Converter for Employee Everything Reports from PDF to CSV format
this module contains the pdf converter which reads employee everything reports in the pdf format and converts them
into csv formatted employee everything reports which can be read by the automatic data entry, auto overmax finder and
the employee everything reader.
"""
from kbtoolbox import inquire, dir_filedialog, find_pp, PdfConverterFix, titlebar_icon
# Standard Libraries
from tkinter import messagebox, filedialog, ttk, Label, Tk
from datetime import timedelta
import os
import csv
from io import StringIO  # change from cStringIO to io for py 3x
import time
import re
# PDF Converter Libraries
from pdfminer.pdfparser import PDFParser
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter, resolve1
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage


def pdf_converter_pagecount(filepath):
    """ gives a page count for pdf_to_text """
    file = open(filepath, 'rb')
    parser = PDFParser(file)
    document = PDFDocument(parser)
    page_count = resolve1(document.catalog['Pages'])['Count']  # This will give you the count of pages
    return page_count


def pdf_to_text(frame, filepath):
    """ Called by pdf_converter() to read pdfs with pdfminer """
    text = None
    codec = 'utf-8'
    password = ""
    maxpages = 0
    caching = (True, True)
    pagenos = set()
    laparams = (
        LAParams(
            line_overlap=.1,  # best results
            char_margin=2,
            line_margin=.5,
            word_margin=.5,
            boxes_flow=0,
            detect_vertical=True,
            all_texts=True),
        LAParams(
            line_overlap=.5,  # default settings
            char_margin=2,
            line_margin=.5,
            word_margin=.5,
            boxes_flow=.5  # detect_vertical=False (default), all_texts=False (default)
            )
    )
    for i in range(2):
        retstr = StringIO()
        rsrcmgr = PDFResourceManager()
        device = TextConverter(rsrcmgr, retstr, codec=codec, laparams=laparams[i])
        interpreter = PDFPageInterpreter(rsrcmgr, device)
        page_count = pdf_converter_pagecount(filepath)  # get page count
        with open(filepath, 'rb') as filein:
            # create progressbar
            pb_root = Tk()  # create a window for the progress bar
            pb_root.geometry("%dx%d+%d+%d" % (450, 75, 200, 300))
            pb_root.title("Klusterbox PDF Converter - reading pdf")
            titlebar_icon(pb_root)  # place icon in titlebar
            Label(pb_root, text="This process takes several minutes. Please wait for results.") \
                .grid(row=0, column=0, columnspan=2, sticky="w")
            pb_label = Label(pb_root, text="Reading PDF: ")  # make label for progress bar
            pb_label.grid(row=1, column=0, sticky="w")
            pb = ttk.Progressbar(pb_root, length=350, mode="determinate")  # create progress bar
            pb.grid(row=1, column=1, sticky="w")
            pb_text = Label(pb_root, text="", anchor="w")
            pb_text.grid(row=2, column=0, columnspan=2, sticky="w")
            pb["maximum"] = page_count  # set length of progress bar
            pb.start()
            count = 0
            # check_extractable=True (default setting)
            for page in PDFPage.get_pages(filein, pagenos, maxpages=maxpages, password=password, caching=caching[i]):
                interpreter.process_page(page)
                pb["value"] = count  # increment progress bar
                pb_text.config(text="Reading page: {}/{}".format(count, page_count))
                pb_root.update()
                count += 1
            text = retstr.getvalue()
            device.close()
            retstr.close()
        pb.stop()  # stop and destroy the progress bar
        pb_label.destroy()  # destroy the label for the progress bar
        pb.destroy()
        pb_root.destroy()
        # test the results
        text = text.replace("", "")
        page = text.split("")  # split the document into page
        result = re.search("Restricted USPS T&A Information(.*)Employee Everything Report", page[0], re.DOTALL)
        try:
            station = result.group(1).strip()
            break
        except:
            if i < 1:
                result = messagebox.askokcancel("Klusterbox PDF Converter",
                                                "PDF Conversion has failed and will not generate a file.  \n\n"
                                                "We will try again.",
                                                parent=frame)
                if not result:
                    return text
            else:
                messagebox.showerror("Klusterbox PDF Converter",
                                     "PDF Conversion has failed and will not generate a file.  \n\n"
                                     "You will either have to obtain the Employee Everything Report "
                                     "in the csv format from management or manually enter in the "
                                     "information",
                                     parent=frame)

    return text


def pdf_converter_reorder_founddays(found_days):
    """ makes sure the days are in the proper order. """
    new_order = []
    correct_series = ("Saturday", "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday")
    for cs in correct_series:
        if cs in found_days:
            new_order.append(cs)
    return new_order


def pdf_converter_path_generator(file_path, add_on, extension):
    """ generate csv file name and path """
    file_parts = file_path.split("/")  # split path into folders and file
    file_name_xten = file_parts[len(file_parts) - 1]  # get the file name from the end of the path
    file_name = file_name_xten[:-4]  # remove the file extension from the file name
    file_name = file_name.replace("_raw_kbpc", "")
    path = file_path[:-len(file_name_xten)]  # get the path back to the source folder
    new_fname = file_name + add_on  # add suffix to to show converted pdf to csv
    new_file_path = path + new_fname + extension  # new path with modified file name
    return new_file_path


def pdf_converter_short_name(file_path):
    """ get the last part of the file name"""
    file_parts = file_path.split("/")  # split path into folders and file
    file_name_xten = file_parts[len(file_parts) - 1]  # get the file name from the end of the path
    return file_name_xten


def pdf_converter(frame):
    """ I have to break this up at some point. """
    kbpc_rpt = None
    kbpc_rpt_file_path = None
    kbpc_raw_rpt_file_path = None
    movecode_holder = None
    date_holder = []
    underscore_slash_result = None
    yyppwk = None
    # inquire as to if the pdf converter reports have been opted for by the user
    sql = "SELECT tolerance FROM tolerances WHERE category ='%s'" % "pdf_error_rpt"
    result = inquire(sql)
    gen_error_report = result[0][0]
    sql = "SELECT tolerance FROM tolerances WHERE category ='%s'" % "pdf_raw_rpt"
    result = inquire(sql)
    gen_raw_report = result[0][0]
    starttime = time.time()  # start the timer
    # make it possible for user to select text file
    sql = "SELECT tolerance FROM tolerances WHERE category ='%s'" % "pdf_text_reader"
    result = inquire(sql)
    allow_txt_reader = result[0][0]
    if allow_txt_reader == "on":
        preference = messagebox.askyesno("PDF Converter",
                                         "Did you want to read from a text file of data output by pdfminer?",
                                         parent=frame)
    else:
        preference = False
    if not preference:  # user opts to read from pdf file
        path = dir_filedialog()
        file_path = filedialog.askopenfilename(initialdir=path,
                                               filetypes=[("PDF files", "*.pdf")])  # get the pdf file
        new_file_path = pdf_converter_path_generator(file_path, "_kbpc", ".csv")  # generate csv file name and path
        short_file_name = pdf_converter_short_name(new_file_path)
        # if the file path already exist - ask for confirmation
        if os.path.exists(new_file_path):
            if not messagebox.askokcancel("Possible File Name Discrepancy",
                                          "There is already a file named {}. "
                                          "If you proceed, the file will be overwritten. "
                                          "Did you want to proceed?".format(short_file_name),
                                          parent=frame):
                return
        # warn user that the process can take several minutes
        if not messagebox.askokcancel("PDF Converter", "This process will take several minutes. "
                                                       "Did you want to proceed?",
                                      parent=frame):
            return
        else:
            text = pdf_to_text(frame, file_path)  # read the pdf with pdfminer
    else:  # user opts to read from text file
        path = dir_filedialog()
        file_path = filedialog.askopenfilename(initialdir=path,
                                               filetypes=[("text files", "*.txt")])  # get the pdf file
        new_file_path = pdf_converter_path_generator(file_path, "_kbpc", ".csv")  # generate csv file name and path
        short_file_name = pdf_converter_short_name(new_file_path)
        # if the file path already exist - ask for confirmation
        if os.path.exists(new_file_path):
            if not messagebox.askokcancel(
                    "Possible File Name Discrepancy",
                    "There is already a file named {}. If you proceed, the file will be overwritten. "
                    "Did you want to proceed?".format(short_file_name),
                    parent=frame):
                return
        gen_raw_report = "off"  # since you are reading a raw report, turn off the generator
        with open(file_path, 'r') as file:  # read the txt file and put it in the text variable
            text = file.read()
    # put the raw output from the pdf conversion into a text file
    if gen_raw_report == "on":
        kbpc_raw_rpt_file_path = pdf_converter_path_generator \
            (file_path, "_raw_kbpc", ".txt")  # generate csv file name and path
        kbpc_raw_rpt = open(kbpc_raw_rpt_file_path, "w")
        kbpc_raw_rpt.write("KLUSTERBOX PDF CONVERSION REPORT \n\n")
        kbpc_raw_rpt.write("Raw output from pdf miner\n\n")
        datainput = "subject file: {}\n\n".format(file_path)
        kbpc_raw_rpt.write(datainput)
        kbpc_raw_rpt.write(text)
        kbpc_raw_rpt.close()
    # create text document for data extracted from the raw pdfminer output
    if gen_error_report == "on":
        kbpc_rpt_file_path = pdf_converter_path_generator(file_path, "_kbpc", ".txt")  # generate csv file name and path
        kbpc_rpt = open(kbpc_rpt_file_path, "w")
        kbpc_rpt.write("KLUSTERBOX PDF CONVERSION REPORT \n\n")
        kbpc_rpt.write("Data extracted from pdfminer output and error reports\n\n")
        datainput = "subject file: {}\n\n".format(file_path)
        kbpc_rpt.write(datainput)
    # define csv writer parameters
    csv.register_dialect('myDialect',
                         delimiter=',',
                         quoting=csv.QUOTE_NONE,
                         skipinitialspace=True,
                         lineterminator="\r"
                         )
    # create the csv file and write the first line
    line = ["TAC500R3 - Employee Everything Report"]
    with open(new_file_path, 'w') as writeFile:
        writer = csv.writer(writeFile, dialect='myDialect')
        writer.writerow(line)
    # define csv writer parameters
    csv.register_dialect('myDialect',
                         delimiter=',',
                         quoting=csv.QUOTE_ALL,
                         skipinitialspace=True,
                         lineterminator=",\r"
                         )
    line = ["YrPPWk", "Finance No", "Organization Name", "Sub-Unit", "Employee Id", "Last Name", "FI", "MI",
            "Pay Loc/Fin Unit", "Var. EAS", "Borrowed", "Auto H/L", "Annual Lv Bal", "Sick Lv Bal", "LWOP Lv Bal",
            "FMLA Hrs", "FMLA Used", "SLDC Used", "Job", "D/A", "LDC", "Oper/Lu", "RSC", "Lvl", "FLSA", "Route #",
            "Loaned Fin #", "Effective Start", "Effective End", "Begin Tour", "End Tour", "Lunch Amt", "1261 Ind",
            "Lunch Ind", "Daily Sched Ind", "Time Zone", "FTF", "OOS", "Day", ]
    with open(new_file_path, 'a') as writeFile:
        writer = csv.writer(writeFile, dialect='myDialect')
        writer.writerow(line)
    text = text.replace("", "")
    page = text.split("")  # split the document into pages
    # whole_line = []
    page_num = 1  # initialize var to count pages
    eid_count = 0  # initialize var to count underscore dash items
    # underscore_slash = []  # arrays for building daily array
    daily_underscoreslash = []
    mv_holder = []
    time_holder = []
    timezone_holder = []
    finance_holder = []
    foundday_holder = []
    daily_array = []
    franklin_array = []
    mv_desigs = ("BT", "MV", "ET", "OT", "OL", "IL", "DG")
    days = ("Saturday", "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday")
    saved_pp = ""  # hold the pp to identify if it changes
    pp_days = []  # array of date/time objs for each day in the week
    found_days = []  # array for holding days worked
    base_time = []  # array for holding hours worked during the day
    eid = ""  # hold the employee id
    lastname = ""  # holds the last name of the employee
    fi = ""
    jobs = []  # holds the d/a code
    routes = []  # holds the route
    level = []  # hold the level (one or two normally)
    base_temp = ("Base", "Temp")
    eid_label = False
    lookforname = False
    lookforfi = False
    lookforroute = False
    lookfor2route = False
    lookforlevel = False
    lookfor2level = False
    base_counter = 0
    base_chg = 0
    lookfortimes = False
    unprocessedrings = ""
    new_page = False
    unprocessed_counter = 0
    mcgrath_indicator = False
    mcgrath_carryover = ""
    rod_rpt = []  # error reports
    frank_rpt = []
    rose_rpt = []
    robert_rpt = []
    stevens_rpt = []
    carroll_rpt = []
    nguyen_rpt = []
    salih_rpt = []
    unruh_rpt = []
    mcgrath_rpt = []
    unresolved = []
    basecounter_error = []
    failed = []
    daily_array_days = []  # build an array of formatted days with just month/ day
    result = re.search('Restricted USPS T&A Information(.*?)Employee Everything Report', page[0], re.DOTALL)
    try:
        station = result.group(1).strip()
    except:
        messagebox.showerror("Klusterbox PDF Converter",
                             "This file does not appear to be an Employee Everything Report. \n\n"
                             "The PDF Converter will not generate a file",
                             parent=frame)
        os.remove(new_file_path)
        if gen_error_report == "on":
            kbpc_rpt.close()
            os.remove(kbpc_rpt_file_path)
        if gen_raw_report == "on":
            os.remove(kbpc_raw_rpt_file_path)
        return
    # start the progress bar
    pb_root = Tk()  # create a window for the progress bar
    pb_root.geometry("%dx%d+%d+%d" % (450, 75, 200, 300))
    pb_root.title("Klusterbox PDF Converter - translating pdf")
    titlebar_icon(pb_root)  # place icon in titlebar
    Label(pb_root, text="This process takes several minutes. Please wait for results.").pack(anchor="w", padx=20)
    pb_label = Label(pb_root, text="Translating PDF: ")  # make label for progress bar
    pb_label.pack(anchor="w", padx=20)
    pb = ttk.Progressbar(pb_root, length=400, mode="determinate")  # create progress bar
    pb.pack(anchor="w", padx=20)
    pb["maximum"] = len(page) - 1  # set length of progress bar
    pb.start()
    pb_count = 0
    for a in page:
        if gen_error_report == "on":
            kbpc_rpt.write(
                "\nxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
                "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx\n")
        if a[0:6] == "Report" or a[0:6] == "":
            pass
        else:
            if gen_error_report == "on":
                kbpc_rpt.write("Out of Sequence Problem!\n")
            eid_count = 0
        if gen_error_report == "on":
            datainput = "Page: {}\n".format(page_num)
            kbpc_rpt.write(datainput)
        try:  # if the page has no station information, then break the loop.
            result = re.search("Restricted USPS T&A Information(.*)Employee Everything Report", a, re.DOTALL)
            station = result.group(1).strip()
            station = station.split('\n')[0]
            if len(station) == 0:
                result = re.search("Employee Everything Report(.*)Weekly", a, re.DOTALL)
                station = result.group(1).strip()
                station = station.split('\n')[0]
        except:
            break
        # get the pay period
        try:
            result = re.search("YrPPWk:\nSub-Unit:\n\n(.*)\n", a)
            yyppwk = result.group(1)
        except:
            try:
                result = re.search("YrPPWk:\n\n(.*)\n\nFin. #:", a)
                yyppwk = result.group(1)
            except:
                try:
                    result = re.findall(r'[0-9][0-9][0-9][0-9]-[0-9][0-9]-[0-9]', text)
                    yyppwk = result[-1]
                except:
                    pass
        if saved_pp != yyppwk:
            exploded = yyppwk.split("-")  # break up the year/pp string from the ee rpt pdf
            year = exploded[0]  # get the year
            if gen_error_report == "on":
                datainput = "Year: {}\n".format(year)
                kbpc_rpt.write(datainput)
            pp = exploded[1]  # get the pay period
            if gen_error_report == "on":
                datainput = "Pay Period: {}\n".format(pp)
                kbpc_rpt.write(datainput)
            pp_wk = exploded[2]  # get the week of the pay period
            if gen_error_report == "on":
                datainput = "Pay Period Week: {}\n".format(pp_wk)
                kbpc_rpt.write(datainput)
            pp += pp_wk  # join the pay period and the week
            first_date = find_pp(int(year), pp)  # get the first day of the pay period
            if gen_error_report == "on":
                datainput = "{}\n".format(str(first_date))
                kbpc_rpt.write(datainput)
            pp_days = []  # build an array of date/time objects for each day in the pay period
            daily_array_days = []  # build an array of formatted days with just month/ day
            for _ in range(7):
                pp_days.append(first_date)
                daily_array_days.append(first_date.strftime("%m/%d"))
                first_date += timedelta(days=1)
            if gen_error_report == "on":
                datainput = "Days in Pay Period: {}\n".format(pp_days)
                kbpc_rpt.write(datainput)
            saved_pp = yyppwk  # hold the year/pp to check if it changes
        page_num += 1
        b = a.split("\n\n")
        for c in b:
            # find, categorize and record daily times
            if lookfortimes:
                if re.match(r"0[0-9]{4}\:\s0[0-9]{2}\.[0-9]{2}$", c):
                    to_add = [base_counter, c]
                    base_time.append(to_add)
                    base_chg = base_counter  # value to check for errors+
                # solve for robertson basetime problem / Base followed by H/L
                elif re.match(r"0[0-9]{4}\:\s0[0-9]{2}\.[0-9]{2}\n0[0-9]{4}\:\s0[0-9]{2}\.[0-9]{2}", c):
                    if "\n" not in c:  # check that there are no multiple times in the line
                        to_add = [base_counter, c]
                        base_time.append(to_add)
                        base_chg = base_counter  # value to check for errors
                        robert_rpt.append(lastname)  # data for robertson baseline problem
                    elif "\n" in c:  # if there are multiple times in the line
                        split_base = c.split("\n")  # split the times by the line break
                        for sb in split_base:  # add each time individually
                            to_add = [base_counter, sb]  # combine the base counter with the time
                            base_time.append(to_add)  # add that time to the array of base times
                            base_chg = base_counter  # value to check for errors
                else:
                    base_counter += 1
                    lookfortimes = False
            if re.match(r"Base", c):
                lookfortimes = True
            # solve for stevens problem / H/L base times not being read
            if len(finance_holder) == 0 and re.match(r"H/L\s", c):  # set trap to catch daily times
                lookfortimes = True
                stevens_rpt.append(lastname)
            checker = False
            one_mistake = False
            underscore_slash = c.split("\n")
            for us in underscore_slash:  # loop through items to detect matches
                if re.match(r"[0-1][0-9]\/[0-9][0-9]", us) or us == "__/__":
                    checker = True
                else:
                    one_mistake = True
            if len(underscore_slash) > 1 and checker == True and one_mistake == False:
                daily_underscoreslash.append(underscore_slash)
            # underscore_slash = []
            d = c.split("\n")
            for e in d:
                try:
                    # build the daily array
                    if re.match(r"[0-9]{6}$", e) and len(movecode_holder) != 0:  # get the route following the chain
                        movecode_holder.append(e)
                        route_holder = movecode_holder
                        if unprocessedrings == "":
                            daily_array.append(route_holder)
                        else:
                            unprocessed_counter += 1  # handle carroll problem
                            carroll_rpt.append(lastname)  # append carroll report
                    movecode_holder = []
                    if len(finance_holder) != 0:  # get the move code following the chain
                        if re.match(r"[0-9]{4}\-[0-9]{2}$", e):
                            finance_holder.append(e)
                            movecode_holder = finance_holder
                        # solve for robertson problem / "H/L" is in move code
                        if re.match(r"H/L", e):  # if the move code is a higher level assignment
                            finance_holder.append(e)
                            finance_holder.append("000000")  # insert zeros for route number
                            if unprocessedrings == "":
                                daily_array.append(
                                    finance_holder)  # skip getting the route and create append daily array
                            else:
                                unprocessed_counter += 1  # handle carroll problem
                                carroll_rpt.append(lastname)  # append carroll report
                    finance_holder = []
                    if len(timezone_holder) != 0:  # get the finance number following the chain
                        timezone_holder.append(e)
                        finance_holder = timezone_holder
                    timezone_holder = []
                    if re.match(r"[A-Z]{2}T", e) and len(time_holder) != 0:  # look for the time zone following chain
                        time_holder.append(e)
                        timezone_holder = time_holder
                    # solve for salih problem / missing time zone in ...
                    elif len(time_holder) != 0 and unprocessedrings != "":
                        unprocessed_counter += 1  # unprocessed rings
                        salih_rpt.append(lastname)
                    time_holder = []
                    # look for time following date/mv desig
                    if re.match(r" [0-2][0-9]\.[0-9][0-9]$", e) and len(date_holder) != 0:
                        date_holder.append(e)
                        time_holder = date_holder
                    # look for items in franklin array to solve for franklin problem
                    if len(franklin_array) > 0 and re.match(r"[0-1][0-9]\/[0-3][0-9]$",
                                                            e):  # if franklin array and date
                        frank = franklin_array.pop(0)  # pop out the earliest mv desig
                        mv_holder = [eid, frank]
                    # solve for rodriguez problem / multiple consecutive mv desigs
                    if len(franklin_array) > 0:
                        if re.match(r"0[0-9]{4}$", e) or re.match(r"0[0-9]{2}$",
                                                                  e) or e in mv_desigs:  # look for move desig
                            franklin_array.append(e)
                            rod_rpt.append(lastname)
                    date_holder = []
                    if re.match(r"[0-1][0-9]\/[0-3][0-9]$", e) and len(
                            mv_holder) != 0:  # look for date following move desig
                        mv_holder.append(e)
                        date_holder = mv_holder
                    # solve for franklin problem: two mv desigs appear consecutively
                    if len(mv_holder) > 0:
                        if re.match(r"0[0-9]{4}$", e) or re.match(r"0[0-9]{2}$",
                                                                  e) or e in mv_desigs:  # look for move desig
                            franklin_array.append(mv_holder[1])
                            franklin_array.append(e)
                            frank_rpt.append(lastname)
                    mv_holder = []
                    if len(franklin_array) == 0:
                        if re.match(r"0[0-9]{4}$", e) or re.match(r"0[0-9]{2}$",
                                                                  e) or e in mv_desigs:  # look for move desig
                            mv_holder.append(eid)
                            mv_holder.append(e)  # place in a holder and check the next line for a date
                    # solve for rose problem: mv desig and date appearing on same line
                    if re.match(r"0[0-9]{4}\s[0-2][0-9]\/[0-9][0-9]$", e):
                        rose = e.split(" ")
                        mv_holder.append(eid)  # add the emp id to the daily array
                        mv_holder.append(rose[0])  # add the mv desig to the daily array
                        mv_holder.append(rose[1])  # add the date to the mv desig array
                        date_holder = mv_holder  # transfer array items to date holder
                        rose_rpt.append(lastname)
                    if e in days:  # find and record all days on the report
                        if eid_label:
                            found_days.append(e)
                        if not eid_label:
                            foundday_holder.append(e)
                    if e == "Processed Clock Rings":
                        eid_count = 0
                    if e == "Employee ID":
                        eid_label = True
                        if gen_error_report == "on":
                            if len(jobs) > 0:
                                datainput = "Jobs: {}\n".format(jobs)
                                kbpc_rpt.write(datainput)
                            if len(routes) > 0:
                                datainput = "Routes: {}\n".format(routes)
                                kbpc_rpt.write(datainput)
                            if len(level) > 0:
                                datainput = "Levels: {}\n".format(level)
                                kbpc_rpt.write(datainput)
                            if len(base_time) > 0:
                                kbpc_rpt.write("Base / Times:")
                                for bt in base_time:
                                    datainput = "{}\n".format(bt)
                                    kbpc_rpt.write(datainput)
                        if len(daily_underscoreslash) > 0:  # bind all underscore slash items in one array
                            underscore_slash_result = sum(daily_underscoreslash, [])
                        # write to csv file
                        prime_info = [yyppwk.replace("-", ""), '"{}"'.format("000000"), '"{}"'.format(station),
                                      '"{}"'.format("0000"), '"{}"'.format(eid), '"{}"'.format(lastname),
                                      '"{}"'.format(fi[:1]),
                                      '"_"', '"010/0000"', '"N"', '"N"', '"N"', '"0"', '"0"', '"0"', '"0"', '"0"',
                                      '"0"']
                        count = 0
                        for array in daily_array:
                            array.append(underscore_slash_result[count])
                            array.append(underscore_slash_result[count + 1])
                            count += 2
                        if base_chg + 1 != len(found_days):  # add to basecounter error array
                            to_add = (lastname, base_chg, len(found_days))
                            if len(found_days) > 0:
                                basecounter_error.append(to_add)
                        # set up array for each day in the week
                        csv_sat = []
                        csv_sun = []
                        csv_mon = []
                        csv_tue = []
                        csv_wed = []
                        csv_thr = []
                        csv_fri = []
                        csv_output = [csv_sat, csv_sun, csv_mon, csv_tue, csv_wed, csv_thr, csv_fri]
                        # reorder the found days to ensure the correct order
                        found_days = pdf_converter_reorder_founddays(found_days)
                        # fix problem with miscounted base times
                        high_array = []
                        for bt in base_time:
                            high_array.append(bt[0])
                        if len(high_array) > 0:
                            high_num = max(high_array)
                            comp_array = []
                            for i in range(high_num + 1):
                                comp_array.append(i)
                            del_array = []
                            for num in comp_array:
                                if num in high_array:
                                    del_array.append(num)
                            error_array = comp_array
                            error_array = [x for x in error_array if x not in del_array]
                            error_array.reverse()
                            if len(error_array) > 0:
                                for error_num in error_array:
                                    for bt in base_time:
                                        if bt[0] > error_num:
                                            bt[0] -= 1
                        # load the multi array with array for each day
                        if len(foundday_holder) > 0:
                            # solve for nguyen problem / day of week occurs prior to "employee id" label
                            found_days += foundday_holder
                            ordered_days = ["Saturday", "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday",
                                            "Friday"]
                            for day in days:  # re order days into correct order
                                if day not in found_days:
                                    ordered_days.remove(day)
                            found_days = ordered_days
                            # foundday_holder = []
                            nguyen_rpt.append(lastname)
                        if len(found_days) > 0:  # printe out found days
                            # reorder the found days to ensure the correct order
                            found_days = pdf_converter_reorder_founddays(found_days)
                            if gen_error_report == "on":
                                datainput = "Found days: {}\n".format(found_days)
                                kbpc_rpt.write(datainput)
                        if gen_error_report == "on":
                            datainput = "proto emp id counter: {}\n".format(eid_count)
                            kbpc_rpt.write(datainput)
                        for i in range(7):
                            for bt in base_time:
                                if found_days[bt[0]] == days[i]:
                                    csv_output[i].append(bt)
                            for da in daily_array:
                                if da[2] == pp_days[i].strftime("%m/%d"):
                                    csv_output[i].append(da)
                        for co in csv_output:  # for each time in the array, printe a line
                            for array in co:
                                if gen_error_report == "on":
                                    datainput = "{}\n".format(array)
                                    kbpc_rpt.write(datainput)
                                # put the data into the csv file
                                if len(array) == 2:  # if the line comes from base/time data
                                    add_this = [found_days[int(array[0])], '"_0-00"', '"{}"'.format(array[1])]
                                    whole_line = prime_info + add_this
                                    with open(new_file_path, 'a') as writeFile:  # add the line to the csv file
                                        writer = csv.writer(writeFile, dialect='myDialect')
                                        writer.writerow(whole_line)
                                if len(array) == 10:  # if the line comes from daily array
                                    if array[9] != "__/__":
                                        end_notes = "(W)Ring Deleted From PC"
                                    else:
                                        end_notes = ""
                                    add_this = ["000-00", '"{}"'.format(array[1]),
                                                '"{}"'.format(
                                                    pp_days[daily_array_days.index(array[2])].strftime(
                                                        "%d-%b-%y").upper()),
                                                '"{}"'.format(array[3].strip()), '"{}"'.format(array[5]),
                                                '"{}"'.format(array[6]),
                                                '"{}"'.format(array[7]), '""', '""', '""', '"0"', '""', '""', '"0"',
                                                '"{}"'.format(end_notes)]
                                    whole_line = prime_info + add_this
                                    with open(new_file_path, 'a') as writeFile:  # add the line to the csv file
                                        writer = csv.writer(writeFile, dialect='myDialect')
                                        writer.writerow(whole_line)
                        # define csv writer parameters
                        csv.register_dialect('myDialect',
                                             delimiter=',',
                                             quotechar="'",
                                             skipinitialspace=True,
                                             lineterminator=",\r"
                                             )
                        if len(jobs) > 0:
                            for i in range(len(jobs)):
                                base_line = [base_temp[i], '"{}"'.format(jobs[i].replace("-", "").strip()),
                                             '"0000"', '"7220-10"',
                                             '"Q0"', '"{}"'.format(level[i]), '"N"', '"{}"'.format(routes[i]), '""',
                                             '"0000000"',
                                             '"0000000"', '"0"', '"0"', '"0"', '"N"', '"N"', '"N"', '"MDT"', '"N"']
                                whole_line = prime_info + base_line
                                with open(new_file_path, 'a') as writeFile:
                                    writer = csv.writer(writeFile, dialect='myDialect')
                                    writer.writerow(whole_line)
                        found_days = []  # initialized arrays
                        lookfortimes = False
                        base_time = []
                        eid = ""
                        base_chg = 0
                        base_counter = 0
                        daily_array = []
                        daily_underscoreslash = []
                        unprocessed_counter = 0
                        jobs = []
                        level = []
                        if gen_error_report == "on":
                            datainput = "{}\n".format(e)
                            kbpc_rpt.write(datainput)
                        eid_count = 0
                    if lookforfi:  # look for first initial
                        if re.fullmatch("[A-Z]\s[A-Z]", e) or re.fullmatch("([A-Z])", e):
                            if gen_error_report == "on":
                                datainput = "FI: {}\n".format(e)
                                kbpc_rpt.write(datainput)
                            fi = e
                            lookforfi = False
                    if lookforname:  # look for the name
                        if re.fullmatch(r"([A-Z]+)", e) \
                                or re.fullmatch(r"([A-Z]+.[A-Z]+)", e) \
                                or re.fullmatch(r"([A-Z]+.[A-Z]+.[A-Z]+)", e) \
                                or re.fullmatch(r"([A-Z]+.[A-Z]+.[A-Z]+.[A-Z]+)", e) \
                                or re.fullmatch(r"([A-Z]+.[A-Z]+.[A-Z]+.[A-Z]+.[A-Z]+)", e):
                            lastname = e.replace("'", " ")
                            if gen_error_report == "on":
                                datainput = "Name: {}\n".format(e)
                                kbpc_rpt.write(datainput)
                            lookforname = False
                            lookforfi = True
                    if re.match(r"\s[0-9]{2}\-[0-9]$", e):  # find the job or d/a code - there might be two
                        jobs.append(e)
                    if lookfor2route:  # look for temp route
                        if re.match(r"[0-9]{6}$", e):
                            routes.append(e)  # add route to routes array
                        lookfor2route = False
                    if lookforroute:  # look for main route
                        if re.match(r"[0-9]{6}$", e):  #
                            routes.append(e)  # add route to routes array
                            lookfor2route = True
                        lookforroute = False
                    if e == "Route #":  # set trap to catch route # on the next line
                        lookforroute = True
                    if lookfor2level:  # intercept the second level
                        if re.match(r"[0-9]{2}$", e):
                            level.append(e)
                        lookfor2level = False
                    if lookforlevel:  # intercept the level
                        if re.match(r"[0-9]{2}$", e):
                            level.append(e)
                            lookfor2level = True  # set trap to catch the second level next line
                        lookforlevel = False
                    if e == "Lvl":  # set trap to catch Lvl on the next line
                        lookforlevel = True
                    if eid != "" and new_page == False:
                        if re.match(r"[0-9]{8}", e):  # find the underscore dash string
                            eid_count += 1
                        if re.match(r"xxx\-xx\-[0-9]{4}", e):
                            eid_count += 1
                        if re.match(r"XXX\-XX\-[0-9]{4}", e):
                            eid_count += 1
                        if e == "___-___-____":
                            eid_count += 1
                        # solve for rose problem: time object is fused to emp id object - just increment the eid counter
                        if re.match(r"\s[0-9]{2}\.[0-9]{10}", e) \
                                or re.match(r"__.__[0-9]{8}", e) \
                                or re.match(r"__._____-___-____", e):
                            eid_count += 1
                            rose_rpt.append(lastname)
                    # solve for carroll problem/ unprocessed rings do not have underscore slash counterparts
                    if e == "Un-Processed Rings":  # after unprocessed rings label, add no new rings to daily array
                        unprocessedrings = eid
                    if re.match(r"[0-9]{8}", e):  # find the emp id / it is the first 8 digit number on the page
                        if eid_count == 0:
                            eid = e
                            if gen_error_report == "on":
                                datainput = "Employee ID: {}\n".format(e)
                                kbpc_rpt.write(datainput)
                            lookforname = True
                            if eid != unprocessedrings:  # set unprocessedrings and new_page variables
                                unprocessedrings = ""
                                new_page = False
                            else:
                                new_page = True
                                eid_count += 1  # increment the eid counter to stop new eid from being set
                                if gen_error_report == "on": kbpc_rpt.write("NEW PAGE!!!\n")
                except:
                    failed.append(lastname)
                    datainput = "READING FAILURE: {}\n".format(e)
                    kbpc_rpt.write(datainput)
        if gen_error_report == "on":  # write to error report
            datainput = "Station: {}\n".format(station)
            kbpc_rpt.write(datainput)
            datainput = "Pay Period: {}\n".format(yyppwk)
            kbpc_rpt.write(datainput)  # show the pay period
            if len(jobs) > 0:
                datainput = "Jobs: {}\n".format(jobs)
                kbpc_rpt.write(datainput)
            if len(routes) > 0:
                datainput = "Routes: {}\n".format(routes)
                kbpc_rpt.write(datainput)
            if len(level) > 0:
                datainput = "Levels: {}\n".format(level)
                kbpc_rpt.write(datainput)
        # define csv writer parameters
        csv.register_dialect('myDialect',
                             delimiter=',',
                             quotechar="'",
                             skipinitialspace=True,
                             lineterminator=",\r"
                             )
        # write to csv file
        prime_info = [yyppwk.replace("-", ""), '"{}"'.format("000000"), '"{}"'.format(station),
                      '"{}"'.format("0000"), '"{}"'.format(eid), '"{}"'.format(lastname), '"{}"'.format(fi[:1]),
                      '"_"', '"010/0000"', '"N"', '"N"', '"N"', '"0"', '"0"', '"0"', '"0"', '"0"', '"0"']
        if len(jobs) > 0:
            # if the route count is less than the jobs count, fill the route count
            routes = PdfConverterFix(routes).route_filler(len(jobs))
            for i in range(len(jobs)):
                base_line = [base_temp[i], '"{}"'.format(jobs[i].replace("-", "").strip()), '"0000"', '"7220-10"',
                             '"Q0"', '"{}"'.format(level[i]), '"N"', '"{}"'.format(routes[i]), '""', '"0000000"',
                             '"0000000"', '"0"', '"0"', '"0"', '"N"', '"N"', '"N"', '"MDT"', '"N"']
                whole_line = prime_info + base_line
                with open(new_file_path, 'a') as writeFile:
                    writer = csv.writer(writeFile, dialect='myDialect')
                    writer.writerow(whole_line)
        if len(foundday_holder) > 0:
            # solve for nguyen problem / day of week occurs prior to "employee id" label
            found_days += foundday_holder
            ordered_days = ["Saturday", "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
            for day in days:  # re order days into correct order
                if day not in found_days:
                    ordered_days.remove(day)
            found_days = ordered_days
            # foundday_holder = []
            nguyen_rpt.append(lastname)
        if len(found_days) > 0:  # printe out found days
            # reorder the found days to ensure the correct order
            found_days = pdf_converter_reorder_founddays(found_days)
            if gen_error_report == "on":
                datainput = "Found days: {}\n".format(found_days)
                kbpc_rpt.write(datainput)
        if gen_error_report == "on":
            datainput = "proto emp id counter: {}\n".format(eid_count)
            kbpc_rpt.write(datainput)
        if len(daily_underscoreslash) > 0:  # bind all underscore slash items in one array
            underscore_slash_result = sum(daily_underscoreslash, [])
        if mcgrath_indicator and len(underscore_slash_result) > 0:  # solve for mcgrath indicator
            mcgrath_carryover.append(underscore_slash_result[0])  # add underscore slash to carryover
            mcgrath_indicator = False  # reset the indicator
            if gen_error_report == "on":
                datainput = "MCGRATH CARRYOVER: {}\n".format(mcgrath_carryover)
                kbpc_rpt.write(datainput)  # printe out a notice.
            del underscore_slash_result[0]  # delete the ophan underscore slash
        count = 0
        for array in daily_array:
            array.append(underscore_slash_result[count])
            try:
                array.append(underscore_slash_result[count + 1])
            except:  # solve for the mcgrath problem
                mcgrath_carryover = array
                mcgrath_indicator = True
                mcgrath_rpt.append(lastname)
                if gen_error_report == "on":
                    kbpc_rpt.write("MCGRATH ERROR DETECTED!!!\n")
            # if mcgrath_indicator == False:
            count += 2
        if mcgrath_carryover in daily_array:  # if there is a carryover, remove the daily array item from the list
            daily_array.remove(mcgrath_carryover)
        if not mcgrath_indicator and mcgrath_carryover != "":  # if there is a carryover to be added
            daily_array.insert(0, mcgrath_carryover)  # put the carryover at the front of the daily array
            mcgrath_carryover = ""  # reset the carryover
            eid_count += 1  # increment the emp id counter
        # set up array for each day in the week
        csv_sat = []
        csv_sun = []
        csv_mon = []
        csv_tue = []
        csv_wed = []
        csv_thr = []
        csv_fri = []
        csv_output = [csv_sat, csv_sun, csv_mon, csv_tue, csv_wed, csv_thr, csv_fri]
        # reorder the found days to ensure the correct order
        found_days = pdf_converter_reorder_founddays(found_days)
        # fix problem with miscounted base times
        high_array = []
        for bt in base_time:
            high_array.append(bt[0])
        if len(high_array) > 0:
            high_num = max(high_array)
            comp_array = []
            for i in range(high_num + 1):
                comp_array.append(i)
            del_array = []
            for num in comp_array:
                if num in high_array:
                    del_array.append(num)
            error_array = comp_array
            error_array = [x for x in error_array if x not in del_array]
            error_array.reverse()
            if len(error_array) > 0:
                for error_num in error_array:
                    for bt in base_time:
                        if bt[0] > error_num:
                            bt[0] -= 1
        # load the multi array with array for each day
        for i in range(7):
            for bt in base_time:
                if found_days[bt[0]] == days[i]:
                    csv_output[i].append(bt)
            for da in daily_array:
                if da[2] == pp_days[i].strftime("%m/%d"):
                    csv_output[i].append(da)
        for co in csv_output:  # for each time in the array, printe a line
            for array in co:
                if gen_error_report == "on":
                    datainput = "{}\n".format(str(array))
                    kbpc_rpt.write(datainput)
                # put the data into the csv file
                if len(array) == 2:  # if the line comes from base/time data
                    add_this = [found_days[int(array[0])], '"_0-00"', '"{}"'.format(array[1])]
                    whole_line = prime_info + add_this
                    with open(new_file_path, 'a') as writeFile:  # add the line to the csv file
                        writer = csv.writer(writeFile, dialect='myDialect')
                        writer.writerow(whole_line)
                if len(array) == 10:  # if the line comes from daily array
                    if array[9] != "__/__":
                        end_notes = "(W)Ring Deleted From PC"
                    else:
                        end_notes = ""
                    add_this = ["000-00", '"{}"'.format(array[1]),
                                '"{}"'.format(pp_days[daily_array_days.index(array[2])].strftime("%d-%b-%y").upper()),
                                '"{}"'.format(array[3].strip()), '"{}"'.format(array[5]), '"{}"'.format(array[6]),
                                '"{}"'.format(array[7]), '""', '""', '""', '"0"', '""', '""', '"0"',
                                '"{}"'.format(end_notes)]
                    whole_line = prime_info + add_this
                    with open(new_file_path, 'a') as writeFile:  # add the line to the csv file
                        writer = csv.writer(writeFile, dialect='myDialect')
                        writer.writerow(whole_line)
        # Handle Carroll problems
        if not mcgrath_indicator:
            if eid_count == 1:  # handle widows
                eid_count = 0
                if gen_error_report == "on":
                    datainput = "WIDOW HANDLING: Carroll Mod emp id counter: {}\n".format(eid_count)
                    kbpc_rpt.write(datainput)
            elif eid_count % 2 != 0:  # handle eid counts where there has been a cut off
                eid_count += 1
                if gen_error_report == "on":
                    datainput = "CUT OFF CONTROL: Carroll Mod emp id counter: {}\n".format(eid_count)
                    kbpc_rpt.write(datainput)
        else:
            eid_count -= 1
        eid_count -= unprocessed_counter * 2

        if unprocessed_counter > 0:
            if gen_error_report == "on":
                datainput = "Unprocessed Rings: {}\n".format(unprocessed_counter)
                kbpc_rpt.write(datainput)
            if len(daily_array) == eid_count / 2:
                pass
            # Solve for Unruh error / when a underscore dash is missing after unprocessed rings
            elif len(daily_array) == max((eid_count + 2) / 2, 0):
                if gen_error_report == "on":
                    datainput = "Unruh Mod emp id counter: {}\n".format(eid_count + 2)
                    kbpc_rpt.write(datainput)
                    kbpc_rpt.write("UNRUH PROBLEM DETECTED!!!")
                unruh_rpt.append(lastname)
            else:
                if gen_error_report == "on":
                    kbpc_rpt.write(
                        "FRANKLIN ERROR DETECTED!!! ALERT! (Unprocessed counter)!\n")
                unresolved.append(lastname)
        else:
            if len(daily_array) != max(eid_count / 2, 0):
                if gen_error_report == "on":
                    kbpc_rpt.write("FRANKLIN ERROR DETECTED!!! ALERT! ALERT!\n")
                unresolved.append(lastname)
        if base_chg + 1 != len(found_days):  # add to basecounter error array
            to_add = (lastname, base_chg, len(found_days))
            if len(found_days) > 0:
                basecounter_error.append(to_add)
        if gen_error_report == "on":
            datainput = "daily array lenght: {}\n".format(len(daily_array))
            kbpc_rpt.write(datainput)
        # initialize arrays
        found_days = []
        foundday_holder = []
        base_time = []
        eid = ""
        eid_label = False
        # perez_switch = False
        base_counter = 0
        base_chg = 0
        daily_array = []
        daily_underscoreslash = []
        unprocessed_counter = 0
        jobs = []
        routes = []
        level = []
        franklin_array = []
        if gen_error_report == "on":
            datainput = "emp id counter: {}\n".format(max(eid_count, 0))
            kbpc_rpt.write(datainput)
        pb["value"] = pb_count  # increment progress bar
        pb_root.update()
        pb_count += 1
    # end loop
    endtime = time.time()
    pb.stop()  # stop and destroy the progress bar
    pb_label.destroy()  # destroy the label for the progress bar
    pb.destroy()
    pb_root.destroy()
    if gen_error_report == "on":
        kbpc_rpt.write("Potential Problem Reports _________________________________________________\n")
        datainput = "runtime: {} seconds\n".format(round(endtime - starttime, 4))
        kbpc_rpt.write(datainput)
        kbpc_rpt.write("Franklin Problems: Consecutive MV Desigs \n")
        datainput = "\t>>> {}\n".format(frank_rpt)
        kbpc_rpt.write(datainput)
        kbpc_rpt.write("Rodriguez Problem: This is the Franklin Problem X 4. \n")
        datainput = "\t>>> {}\n".format(rod_rpt)
        kbpc_rpt.write(datainput)
        kbpc_rpt.write("Rose Problem: The MV Desig and date are on the same line.\n")
        datainput = "\t>>> {}\n".format(rose_rpt)
        kbpc_rpt.write(datainput)
        kbpc_rpt.write("Robertson Baseline Problem: The base count is jumping when H/L basetimes "
                       "are put into the basetime array.\n")
        datainput = "\t>>> {}\n".format(robert_rpt)
        kbpc_rpt.write(datainput)
        kbpc_rpt.write("Stevens Problem: Basetimes begining with H/L do not show up and are "
                       "not entered into the basetime array.\n")
        datainput = "\t>>> {}\n".format(stevens_rpt)
        kbpc_rpt.write(datainput)
        kbpc_rpt.write("Carroll Problem: Unprocessed rings at the end of the page do not contain __/__ or times.'n")
        datainput = ">>> {}\n".format(carroll_rpt)
        kbpc_rpt.write(datainput)
        kbpc_rpt.write("Nguyen Problem: Found day appears above the Emp ID.\n")
        datainput = "\t>>> {}\n".format(nguyen_rpt)
        kbpc_rpt.write(datainput)
        kbpc_rpt.write("Unruh Problem: Underscore dash cut off in unprecessed rings.\n")
        datainput = "\t>>> {}\n".format(unruh_rpt)
        kbpc_rpt.write(datainput)
        kbpc_rpt.write(
            "Salih Problem: Unprocessed rings are missing a timezone, so that unprocessed rings counter is not"
            " incremented.\n")
        datainput = "\t>>> {}\n".format(salih_rpt)
        kbpc_rpt.write(datainput)
        kbpc_rpt.write("McGrath Problem: \n")
        datainput = " \t>>> {}\n".format(mcgrath_rpt)
        kbpc_rpt.write(datainput)
        datainput = "Unresolved: {}\n".format(unresolved)
        kbpc_rpt.write(datainput)
        datainput = "Base Counter Error: {}\n".format(basecounter_error)
        kbpc_rpt.write(datainput)
    if len(failed) > 0:  # create messagebox to show any errors
        failed_daily = ""
        for f in failed:
            failed_daily = failed_daily + " \n " + f
        messagebox.showerror("Klusterbox PDF Converter",
                             "Errors have occured for the following carriers {}."
                             .format(failed_daily),
                             parent=frame)
    # create messagebox for completion
    messagebox.showinfo("Klusterbox PDF Converter",
                        "The PDF Convertion is complete. "
                        "The file name is {}. ".format(short_file_name),
                        parent=frame)
