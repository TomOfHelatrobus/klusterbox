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
        i = 0
        for line in self.a_file:
            pb.move_count(i)  # increment progress bar
            update = "row: " + str(i) + "/" + str(rowcount)  # build string for status message update
            pb.change_text(update)  # update status message on progress bar
            i += 1
            if firstline:  # do not subject the first line to any test
                self.build_csv(line)
                firstline = False
            elif self.testforblank(line):  # returns True if the line is blank
                pass
            elif self.testforempties(line):  # returns True if an array in the multidimensional array is not empty
                pass
            elif self.testfordupbase(line):  # returns True if redundant Base lines are found
                pass
            elif self.testforduptemp(line):  # returns True if redundant Temp lines are found
                pass
            else:
                self.build_csv(line)  # writes lines to the csv file
        pb.stop()  # stop and destroy the progress bar

    @staticmethod
    def testforblank(line):  # returns True if the line is blank
        if line:
            return False
        return True

    @staticmethod
    def testforempties(line):  # returns True if an array in the multidimensional array is not empty
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

    def destroy(self):
        """ remove the new file when it is no longer needed """
        if os.path.isfile(self.new_filepath):
            os.remove(self.new_filepath)
