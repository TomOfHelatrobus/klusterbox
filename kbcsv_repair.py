"""
a klusterbox module: Klusterbox Checking and Repair for Employee Everything Reports in CSV format
this module contains code for the csv repair file which reads a csv employee everything report, then writes a new file
skipping blank lines or arrays no values.
"""
from kbtoolbox import BuildPath, ProgressBarDe
import csv
import os
from os.path import exists


class CsvRepair:
    """
    This class will remove all blank lines or lines of commas from employee everything report csv files.
    It will also remove any duplicate Base or Temp lines for the same carrier thus limiting all carriers to
    only one Base line and only one Temp line.
    """
    def __init__(self):
        self.frame = None
        self.file_path = None
        self.new_filepath = None
        self.a_file = None
        self.baseid = []  # an array of employee ids where Base line has been read
        self.tempid = []  # an array of employee ids where Temp line has been read
        self.eid = ""  # the employee id, fourth column in the row.
        self.build_i = 0  # build index - tracks the number of lines written. on 3 get first employee id.
        self.write_it = True  # if True, will write rows instead of caching them for further analysis
        self.cache = []  # a cache of rows saved for analysis
        self.lastgoodid = None  # the last good id seen
        self.lastgoodname = None  # the last good name seen - not ET, MV a timezone or anything lowercase

    def run(self, file_path):
        """ this runs the classes when called. """
        self.file_path = file_path  # after testing uncomment this and put file path into arguments
        self.get_newfilepath()  # create a name for the path
        self.delete_filepath()
        self.destroy()  # if the file already exist. destroy it.
        self.get_file()
        self.test_write()
        return self.new_filepath

    def get_newfilepath(self):
        """ this creates a new file path """
        path_splice = self.file_path.split("/")  # split the file path into an array
        filename = path_splice.pop()  # remove the old filename from the end of the file path
        newfilename = ""
        if self.file_path[-4:].lower() == ".xls":
            newfilename = filename.replace(".xls", "_fixed.xls")  # change the old file name
        if self.file_path[-4:].lower() == ".csv":
            newfilename = filename.replace(".csv", "_fixed.csv")  # change the old file name
        path_splice.append(newfilename)  # add the new file name to the end of the path
        self.new_filepath = BuildPath().build(path_splice)  # get the new file path

    def delete_filepath(self):
        """ delete the filepath if it exist to avoid permission error """
        if exists(self.new_filepath):
            try:
                os.remove(self.new_filepath)
            except PermissionError:
                self.new_filepath.close()
                os.remove(self.new_filepath)

    def get_file(self):
        """ read the csv file and assign to self.a_file attribute """
        self.a_file = csv.reader(open(self.file_path, newline=""))

    def build_csv(self, line):
        """ writes lines to the csv file """
        csv.register_dialect('myDialect', lineterminator="\r")
        with open(self.new_filepath, 'a') as f:
            writer = csv.writer(f, dialect='myDialect')
            writer.writerow(line)
        self.build_i += 1  # keep track of how many lines have been written to the new csv file

    def csvrowcount(self):
        """ gets the number of rows from the csv file """
        with open(self.file_path) as f:
            return sum(1 for _ in f)

    def test_write(self):
        """ this checks the old csv and writing appropiate lines and skipping blank lines."""
        pb = ProgressBarDe(title="Checking csv file", label="Reading and writing proxy file: ")  # create progressbar
        rowcount = self.csvrowcount()  # gets the number of rows from the csv file
        pb.max_count(rowcount)  # set length of progress bar
        pb.start_up()
        firstline = True
        i = 0  # index - used for updating the process bar
        for line in self.a_file:
            pb.move_count(i)  # increment progress bar
            update = "row: " + str(i) + "/" + str(rowcount)  # build string for status message update
            pb.change_text(update)  # update status message on progress bar
            i += 1
            if firstline:  # do not subject the first line to any test
                self.build_csv(line)
                firstline = False
            elif self.testfordupfirstline(line):  # returns True if this is a duplicate first line
                pass
            elif self.testforblank(line):  # returns True if the line is blank
                pass
            elif self.testforempties(line):  # returns True if an array in the multidimensional array is not empty
                pass
            elif self.testfordupbase(line):  # returns True if redundant Base lines are found
                pass
            elif self.testforduptemp(line):  # returns True if redundant Temp lines are found
                pass
            else:
                self.get_firsteid(line)  # get the first employee id number from the third line - column 4
                self.checkforneweid(line)  # check if the row has a new employee id.
                if self.write_it:  # if the first line is a Base line then write, otherwise cache
                    line = self.testforbadname(line)  # rewrite line if the name is BT or MV
                    self.build_csv(line)  # writes lines to the csv file
                else:
                    self.cache_rows(line)
        pb.stop()  # stop and destroy the progress bar

    def get_firsteid(self, line):
        """
        get_firsteid, checkforneweid, checkforbase, cache_rows and cache_analysis all work together to
        repair the Denton Problem which can be created in the pdf converter. This happens when the csv file
        is generated in a such a way that the Base and Temp lines occur in the middle of the times and
        rings. these methods work to indentify when base is not the first row, collects all lines for that
        carrier, then re arranges them in proper order.

        get the first employee id number from the third line - column 4 """
        if self.build_i == 2:  # this is the first row of carrier information
            self.eid = line[4]

    def checkforneweid(self, line):
        """ check for a new employee id """
        if self.eid != line[4]:
            if self.cache:  # if there is something in the cache
                self.cache_analysis()
            # get the base and temp lines of the carrier with the new employee id.
            self.checkforbase(line)
        self.eid = line[4]

    def checkforbase(self, line):
        """ checks that the first line of a carrier's information is a base line """
        if line[18] not in ("Base", ):
            self.write_it = False

    def cache_rows(self, line):
        """ cache rows if the base line is not the first line. """
        self.cache.append(line)

    def cache_analysis(self):
        """ go through the cache to put it in the correct order """
        basetemp = ("Base", "Temp")
        for type_ in basetemp:  # look for "base" first, then "temp" next
            for row in self.cache:
                if row[18] == type_:  # if base/temp are found
                    row = self.testforbadname(row)  # rewrite row if the name is BT or MV
                    self.build_csv(row)  # write the line
        for row in self.cache:  # write the remaining rows, omitting base/temp rows.
            if row[18] not in basetemp:
                row = self.testforbadname(row)  # rewrite row if the name is BT or MV
                self.build_csv(row)
        self.cache = []  # empty out the cache
        self.write_it = True  # reset the write_it variable so the cache is not appended.

    @staticmethod
    def testfordupfirstline(line):
        """ check if there is a duplicate first line"""
        if "TAC500R3" in line[0] or "Employee Everything Report" in line[0]:
            return True
        return False

    @staticmethod
    def testforblank(line):
        """ returns True if the line is blank """
        if line:
            return False
        return True

    @staticmethod
    def testforempties(line):
        """ returns True if an array in the multidimensional array is not empty"""
        try:
            for i in range(len(line)):  # check each array in the multidimensional array
                if line[i]:
                    return False
            return True
        except IndexError:
            return False

    def testfordupbase(self, line):
        """
        looks for "Base" or "Temp" in line[18]. if found, the employee compared to a list of other employee ids and
        is added if it is not found. Then the method returns False. Otherwise, True is returned to show that this
        is a redundant "Base" or "Temp" line for a carrier.
        """
        if line[18] == "Base":
            if line[4] not in self.baseid:
                self.baseid.append(line[4])  # add to list of carriers with a base line
                return False
            return True

    def testforduptemp(self, line):
        if line[18] == "Temp":
            if line[4] not in self.tempid:
                self.tempid.append(line[4])  # add to list of carriers with a temp line
                return False
            return True

    def testforbadname(self, line):
        """ check to see if the name is BT or MV. If so, then swap it for the lastgoodname, lastgoodid. """
        if line[5] not in ("BT", "MV", "CDT", "CST", "MDT", "MST", "PDT", "PST", "HDT", "HST") \
                and line[5].isupper():
            self.lastgoodname = line[5]
            self.lastgoodid = line[4]
            return line
        else:
            line[4] = self.lastgoodid
            line[5] = self.lastgoodname
            return line

    def destroy(self):
        """ remove the new file when it is no longer needed """
        if os.path.isfile(self.new_filepath):
            os.remove(self.new_filepath)
