"""
This is a module in the klusterbox library. This module will provide version specific updates and fixes.
A method will look up the value for the fixes column in the tolerances table, if the value in the table is less than
the version number, fixes will be looked for and applied.  Once fixes are applied, then the value in the tolerance
table will be updated to match the current version number, so that the update only occurs once.
"""

from kbtoolbox import isfloat, inquire, commit


class Fixes:
    """ check the version number and compare it to the number for the fix. """
    def __init__(self):
        self.version = None  # the version number is passed from main
        self.lastfix = None  # the last fix is fetched from the database, expressed as version number.

    def check(self, version):
        """ get the version number and compare it to the version number of checks """
        if not isfloat(version):  # version numbers must be compatable with a float. e.g. 5.08
            return
        self.version = float(version)
        self.get()  # get the number of the most resently done fix from the tolerances table
        # compare the version number to the version number of the fix and what's been marked done in the dbase
        if not self.compare():
            return
        self.update_lastfix()

    def get(self):
        """ get the number of the most resently done fix from the tolerances table."""
        sql = "SELECT tolerance from tolerances WHERE category = '%s'" % "lastfix"
        result = inquire(sql)
        self.lastfix = float(result[0][0])  # update the lastfix value in the tolerances table.

    def compare(self):
        """ compare the version number to the version number of the fix and what's been marked done in the dbase. """
        if self.lastfix >= self.version:
            return False
        if self.version >= 5.0:
            V5000Fix().run()
            V5000FixA().run()
        # if self.version >= 5.08:
        #     V5008Fix().run()
        return True

    def update_lastfix(self):
        """ update the lastfix value in the tolerances table. """
        sql = "UPDATE tolerances SET tolerance = '%s' WHERE category = 'lastfix'" % self.version
        commit(sql)


class V5000Fix:
    """ this fix checks all the names in the database in 7 tables for uppper case
    values and replaces any it finds with lower case spellings of the name.
    """

    def __init__(self):
        self.carrier_list = []
        self.infc_awards = []
        self.infc_payouts = []
        self.name_index = []
        self.otdl_pref = []
        self.refuse_list = []
        self.rings_list = []
        # ensure that the tableslist, name_convention and name_array all have the same number of elements.
        self.tablelist = ["carriers", "informalc_awards", "informalc_payouts", "name_index", "otdl_preference",
                          "refusals", "rings3"]
        self.name_convention = ["carrier_name", "carrier_name", "carrier_name", "kb_name", "carrier_name",
                                "carrier_name", "carrier_name"]
        self.name_array = [self.carrier_list, self.infc_awards, self.infc_payouts, self.name_index, self.otdl_pref,
                           self.refuse_list, self.rings_list]
        self.iterations = 0

    def run(self):
        """ a master method for running other methods in proper order"""
        if not self.check_arrays():  # avoid index errors if the arrays don't have similar lengths.
            return
        self.delete_null_names()
        self.get_carriers()
        self.fix_carriers()

    def check_arrays(self):
        """ ensure that the tableslist, name_convention and name_array all have the same number of elements. """
        if len(self.tablelist) == len(self.name_convention):
            if len(self.name_convention) == len(self.name_array):
                self.iterations = len(self.tablelist)
                return True
        return False

    def delete_null_names(self):
        """ delete all the records where the name is null """
        for i in range(self.iterations):
            sql = "SELECT DISTINCT {} from {}".format(self.name_convention[i], self.tablelist[i])
            results = inquire(sql)
            if results:
                for carrier in results:
                    if carrier[0] is None:
                        sql = "DELETE FROM {} WHERE {} IS NULL".format(self.tablelist[i], self.name_convention[i])
                        commit(sql)

    def get_carriers(self):
        """ get a list of distinct names from the carriers table. """
        for i in range(self.iterations):
            sql = "SELECT DISTINCT {} from {}".format(self.name_convention[i], self.tablelist[i])
            results = inquire(sql)
            if results:
                for carrier in results:
                    if not carrier[0].islower():
                        self.name_array[i].append(carrier[0])

    def fix_carriers(self):
        """ check if the name is all lower, if not, update the record. """
        for i in range(self.iterations):
            for carrier in self.name_array[i]:
                if not carrier.islower():
                    sql = "UPDATE {} SET {} = '%s' WHERE {} = '%s'"\
                              .format(self.tablelist[i], self.name_convention[i], self.name_convention[i]) \
                          % (carrier.lower(), carrier)
                    commit(sql)


class V5000FixA:
    """
    replace the null values in rings 3 in bt and et columns with empty strings.
    also replace null values in rings 3 in leave type and time with empty string and empty float.
    """

    def __init__(self):
        pass

    def run(self):
        """ master method for running other methods in proper order. """
        if self.check_for_null():
            self.update_null_to_emptystring()
        if self.check_for_null_leave():
            self.update_null_leavetime_type()

    @staticmethod
    def check_for_null():
        """ returns true if there are null values in bt """
        sql = "SELECT * FROM rings3 WHERE bt IS NULL"
        results = inquire(sql)
        if results:
            return True
        return False

    @staticmethod
    def update_null_to_emptystring():
        """ change any null values in bt and et to empty strings in the rings3 table. """
        sql = "UPDATE rings3 SET bt = '' WHERE bt IS NULL"
        commit(sql)
        sql = "UPDATE rings3 SET et = '' WHERE et IS NULL"
        commit(sql)

    @staticmethod
    def check_for_null_leave():
        """ returns true if there are null values in leave time or type. """
        sql = "SELECT * FROM rings3 WHERE leave_type IS NULL"
        results = inquire(sql)
        if results:
            return True
        return False

    @staticmethod
    def update_null_leavetime_type():
        """ this converts leave type to an empty string and leave time to an empty float. """
        types = ""
        times = float(0.0)
        sql = "UPDATE rings3 SET leave_type='%s',leave_time='%s'" \
              "WHERE leave_type IS NULL" \
              % (types, times)
        commit(sql)

class V5008Fix:
    """ this is a migration of the a table used for informal c into two new tables.
    informalc_grv will be copied, and the contents written into informalc_grievances and informalc_settlements. """
