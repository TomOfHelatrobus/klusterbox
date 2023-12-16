""" this module is not for general release. it takes existing data from the db and makes aliases for all the names that
appear in the database. """

from kbtoolbox import inquire, commit
from tqdm import tqdm
from time import sleep


class AliasMaker:
    """ this fix checks all the names in the database in 7 tables for uppper case
    values and replaces any it finds with lower case spellings of the name.
    """

    def __init__(self):
        self.carrier_list = []  # list in self.name_array for carriers table.
        self.infc_awards = []  # list in self.name_array for informalc_awards2 table.
        self.infc_grvs = []  # list in self.name_array for informalc_grievances table.
        self.infc_payouts = []  # list in self.name_array for informalc_payouts table.
        self.name_index = []  # list in self.name_array for name_index table.
        self.otdl_pref = []  # list in self.name_array for "otdl_preference table.
        self.refuse_list = []  # list in self.name_array for refusals table.
        self.rings_list = []  # list in self.name_array for rings3 table.
        self.all_stations = []  # a list of all stations in the database
        self.random_name_array = []  # a list of random names generated from the aliases.txt doc.
        self.distinct_array = []  # a list of distinct names from all tables in self.tablelist
        self.station_dict = {}  # a dictionary where true station names are keys and aliases are values
        self.stationtables = ["carriers", "dov", "informalc_grievances", "otdl_preference", "station_index", "stations"]
        self.station_convention = ["station", "station", "station", "station", "kb_station", "station"]
        self.alias_dict = {}  # a dictionary where true names are keys and alias are values
        # ensure that the tableslist, name_convention and name_array all have the same number of elements.
        self.tablelist = ["carriers", "informalc_awards2", "informalc_grievances", "informalc_payouts", "name_index",
                          "otdl_preference", "refusals", "rings3"]
        self.name_convention = ["carrier_name", "carrier_name", "grievant", "carrier_name", "kb_name", "carrier_name",
                                "carrier_name", "carrier_name"]
        self.name_array = [self.carrier_list, self.infc_awards, self.infc_grvs, self.infc_payouts, self.name_index,
                           self.otdl_pref, self.refuse_list, self.rings_list]
        self.station_iterations = 0
        self.iterations = 0
        self.station_aliases = [  # a list of station aliases
            "kluster park",
            "kluster oaks",
            "klusterville",
            "kluster hills",
            "klusterton",
            "klusterberg"
        ]

    def run(self):
        """ a master method for running other methods in proper order"""
        if not self.areyousure():  # if the user chooses to not proceed.
            exit()
        if not self.check_station_arrays():  # avoid index errors if the arrays don't have similar lengths.
            exit()
        self.get_stations()
        if not self.create_station_dict():  # if there are not enough fake station names
            exit()
        if not self.check_arrays():  # avoid index errors if the arrays don't have similar lengths.
            exit()
        if not self.fetch_names():  # if there is an error in reading 'aliases.txt' doc.
            exit()
        self.delete_null_names()
        self.get_carriers()
        if not self.get_distinct_list():  # if there are not enough fake names for all real names
            exit()
        self.create_name_dict()
        self.alias_stations()
        self.alias_carriers()

    @staticmethod
    def areyousure():
        """ prompt user for input to ensure they want to proceed. """
        print("Alias Maker")
        print("  This program will create aliases for all station and carrier names in the klusterbox database.")
        print("  This can not be undone. Be sure to make a copy of mandates.sqlite before proceeding...")
        answer = input("  Are you sure you want to proceed? If so, enter 'yes' >>> ")
        if answer.lower().strip() != "yes":
            print("  Probably a good idea. Alias Maker will terminate. ")
            return False
        return True

    def check_station_arrays(self):
        """ ensure that the stationtables and station_convention all have the same number of elements. """
        if len(self.stationtables) == len(self.station_convention):
            self.station_iterations = len(self.stationtables)
            return True
        print("ERROR: There are unequal array lengths for staitionlist and station_convention. \n"
              "Alias Maker will terminate.")
        return False

    def get_stations(self):
        """ get a list of stations """
        sql = "SELECT * FROM stations"
        results = inquire(sql)
        for rec in results:
            self.all_stations.append(rec[0])  # get all stations in database.
        self.all_stations.remove("out of station")  # remove out of station.

    def create_station_dict(self):
        """ create a dictionary of station aliases. """
        if len(self.all_stations) > len(self.station_aliases):
            print("ERROR: There are {} stations and only {} station aliases. \n"
                  "Please add station aliases to __init__ in alias_maker.py module. "
                  .format(len(self.all_stations), len(self.station_aliases)))
            return False
        i = 0
        for station in self.all_stations:  # loop though each distinct name
            self.station_dict[station] = self.station_aliases[i]
            i += 1
        print("SUCCESS: There are {} stations aliases ready for use".format(len(self.station_aliases)))
        return True

    def check_arrays(self):
        """ ensure that the tableslist, name_convention and name_array all have the same number of elements. """
        if len(self.tablelist) == len(self.name_convention):
            if len(self.name_convention) == len(self.name_array):
                self.iterations = len(self.tablelist)
                return True
        print("ERROR: There are unequal array lengths for tablelist, name_convention and name_array. \n"
              "Alias Maker will terminate.")
        return False

    def fetch_names(self):
        """ this will read the names on the text file 'aliases.txt' and make each line into a element in the
        name_array list. """
        try:
            file = open("aliases.txt", "r")
        except FileNotFoundError:
            print("ERROR: The \'aliases.txt\' file is not in the project folder. Alias Maker will terminate. ")
            return False
        content = file.readlines()
        for name in content:
            name = name.strip("\n")
            name = name.strip()
            name = name.lower()
            self.random_name_array.append(name)
        if not self.random_name_array:
            print("WARNING: There were no names in aliases.txt document. Alias Maker will terminate. ")
            return False
        else:
            print("SUCCESS: There are {} aliases ready for use... ".format(len(self.random_name_array)))
        file.close()
        return True

    def delete_null_names(self):
        """ delete all the records where the name is null """
        for i in range(self.iterations):  # loop for each table
            sql = "SELECT DISTINCT {} from {}".format(self.name_convention[i], self.tablelist[i])
            results = inquire(sql, returnerror=True)  # use kwarg to check for operational error
            # check that the table exist, if it does not then use 'continue' to skip iteration.
            if results == "OperationalError":  # if operational error is returned - return False
                continue
            if results:
                for carrier in results:
                    if carrier[0] is None:
                        sql = "DELETE FROM {} WHERE {} IS NULL".format(self.tablelist[i], self.name_convention[i])
                        commit(sql)

    def get_carriers(self):
        """ get a list of distinct names from the carriers table. """
        for i in range(self.iterations):  # loop for each table
            sql = "SELECT DISTINCT {} from {}".format(self.name_convention[i], self.tablelist[i])
            results = inquire(sql, returnerror=True)  # use kwarg to check for operational error
            # check that the table exist, if it does not then use 'continue' to skip iteration.
            if results == "OperationalError":  # if operational error is returned - then skip
                continue
            if results:
                for carrier in results:
                    if carrier[0]:
                        self.name_array[i].append(carrier[0])

    def get_distinct_list(self):
        """ loop through all name arrays and get a list of distinct names """
        for table in self.name_array:
            for name in table:
                if name not in self.distinct_array:
                    self.distinct_array.append(name)
        if len(self.random_name_array) < len(self.distinct_array):
            print("ERROR: There are not a sufficent amount of aliases to continue. "
                  "You need {} names. Alias Maker will terminate. ".format(len(self.distinct_array)))
            return False
        return True

    def create_name_dict(self):
        """ create a dictionary of name aliases. """
        i = 0
        for name in self.distinct_array:  # loop though each distinct name
            surname = name.split(",")[0]  # separate the last name from the first initial/name
            self.alias_dict[surname] = self.random_name_array[i]
            i += 1

    def alias_stations(self):
        """ replace the station name with an alias in the db. """
        for i in range(self.station_iterations):  # loop for each table
            sleep(.5)
            for ii in tqdm(range(len(self.all_stations)), colour="#00ffff",
                           desc="station aliases {}: ".format(self.stationtables[i])):
                # print(self.stationtables[i] + "/////////////////////////////////////////////////////")
                sql = "UPDATE {} SET {} = '%s' WHERE {} = '%s'"\
                      .format(self.stationtables[i], self.station_convention[i], self.station_convention[i]) \
                      % (self.station_dict[self.all_stations[ii]], self.all_stations[ii])
                # print(sql)
                commit(sql)

    def alias_carriers(self):
        """ check if the name is all lower, if not, update the record. """
        for i in range(self.iterations):  # loop for each table
            # print(self.tablelist[i] + "/////////////////////////////////////////////////////")
            sleep(.5)
            for ii in tqdm(range(len(self.name_array[i])), colour="#00ffff",
                           desc="carrier aliases {}: ".format(self.tablelist[i])):
                if self.name_array[i][ii] == "class action":
                    alias = "class action"
                else:
                    splitname = self.name_array[i][ii].split(",")
                    alias = "{},{}".format(self.alias_dict[splitname[0]], splitname[1])
                sql = "UPDATE {} SET {} = '%s' WHERE {} = '%s'"\
                    .format(self.tablelist[i], self.name_convention[i], self.name_convention[i]) \
                      % (alias, self.name_array[i][ii])
                # print(sql)
                commit(sql)


if __name__ == "__main__":
    """ this is where the program starts """
    AliasMaker().run()
