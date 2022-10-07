""" counts lines of all python files in the project folder.
the script must be in the target project folder to use.
enter 'python countlines.py' at prompt """
import os


class CountLines:
    """ counts the number of lines of all python files in the project folder."""

    def __init__(self):
        self.all_files = []  # a list of python files in the project folder
        self.total_count = 0  # a running count of all the lines of all the python files.
        self.file_index = 0  # this is an index for all_files array

    def run(self):
        """ master method for running other methods."""
        self.get_files()
        for file in self.all_files:
            self.count_lines(file)
            self.file_index += 1
        self.summary_report()

    def get_files(self):
        """ get a list of all files in the directory. put all names into a list. """
        for file in os.listdir():
            if file.endswith(".py"):
                if file != "countlines.py":  # do not include the countlines.py file
                    self.all_files.append(file)

    def count_lines(self, file):
        """ counts the lines of a file """
        with open(file, encoding="utf8") as f:  # open the file with proper utf8 encoding
            num_lines = sum(1 for _ in f)  # count the number of lines
            name = self.get_name()  # adds periods to the end of the name to improve reabability
            print("   {:<30} {}".format(name, num_lines))
            self.total_count += num_lines

    def summary_report(self):
        """ displays a summary report at the end """
        print("")
        print("   Number of Files ..............", len(self.all_files))
        print("   Total Lines ..................", self.total_count)

    def get_name(self):
        """ returns the name with periods following it to help readability"""
        name = self.all_files[self.file_index]
        fix_name = ""
        for i in range(30):
            if i < len(name):  # enter name, letter by letter
                fix_name += name[i]
            elif i == len(name):  # enter a blank space immediately after name
                fix_name += " "
            else:
                fix_name += "."  # enter a period after the name
        return fix_name


CountLines().run()
