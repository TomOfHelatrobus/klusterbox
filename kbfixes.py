"""
This is a module in the klusterbox library. This module will provide version specific updates and fixes.
A method will look up the value for the fixes column in the tolerances table, if the value in the table is less than
the version number, fixes will be looked for and applied.  Once fixes are applied, then the value in the tolerance
table will be updated to match the current version number, so that the update only occurs once.
"""

from kbtoolbox import isfloat, inquire, commit, ProgressBarDe, Convert


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
        if self.version >= 5.08:
            V5008Fix().run()
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
            if results == "OperationalError":  # if operational error is returned - return False
                continue
            if results:
                for carrier in results:
                    if not carrier[0].islower():
                        self.name_array[i].append(carrier[0])

    def fix_carriers(self):
        """ check if the name is all lower, if not, update the record. """
        for i in range(self.iterations):  # loop for each table
            for carrier in self.name_array[i]:
                if not carrier.islower():
                    sql = "SELECT * FROM {} WHERE {} = %s".format(self.tablelist[i], self.name_convention[i]) % carrier
                    results = inquire(sql, returnerror=True)  # use kwarg to check for operational error
                    # check that the table exist, if it does not then use 'continue' to skip iteration.
                    if results == "OperationalError":  # if operational error is returned - return False
                        continue
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

    def __init__(self):
        self.transfer_array = []
        self.distinct_grievances = []
        self.distinct_settlements = []
        self.pb = None
        self.awards_transfer_array = []  # hold all recs from the informalc awards table

    def run(self):
        """ a master method for controlling other methods """
        self.informalc_grv_xfer()  # transfer for informalc grv records to grievances and settlements
        self.informalc_awards_xfer()  # transfer informalc awards recs to informalc awards2 table

    def informalc_grv_xfer(self):
        """ a run a progress bar and the transfer for informalc grv records to grievances and settlements """
        if not self.get_transfer_array():
            return
        self.pb = ProgressBarDe(title="Informal C Data Transfer", label="transfer all grievance/settlement data")
        self.pb.max_count(len(self.transfer_array))
        self.pb.start_up()
        self.get_distinct()
        self.transfer_records()
        self.pb.stop()

    def get_transfer_array(self):
        """ get an array of information to be moved to new tables. """
        sql = "SELECT * FROM informalc_grv"  # if the table does not exist - abort transfer
        self.transfer_array = inquire(sql, returnerror=True)  # use kwarg to check for operational error
        if self.transfer_array == "OperationalError":  # if operational error is returned - return False
            return False
        if not self.transfer_array:  # if there is nothing in the results - return False
            return False
        return True

    def get_distinct(self):
        """ because we don't want to write duplicate records get a list of distinct records from informalc
        grievances and informalc settlements so we can check if a record exist before we create it. """
        sql = "SELECT DISTINCT grv_no FROM informalc_grievances"
        results = inquire(sql)
        for r in results:
            self.distinct_grievances.append(*r)
        sql = "SELECT DISTINCT grv_no FROM informalc_settlements"
        results = inquire(sql)
        for r in results:
            self.distinct_settlements.append(*r)

    def transfer_records(self):
        """ insert data into informalc_grievances, informalc_settlements and informalc_gats tables.
        when the transfer is complete, delete the informalc_grv table"""
        i = 1
        for rec in self.transfer_array:
            self.pb.change_text("processing: {}".format(rec[0]))
            self.pb.move_count(i)
            if rec[0] not in self.distinct_grievances:  # if the grievance number is not in the list of distinct
                sql = "INSERT INTO informalc_grievances " \
                      "(grievant, station, grv_no, startdate, enddate, meetingdate, issue, article) " \
                      "VALUES ('', '%s', '%s', '%s', '%s', '', '%s', '')" % \
                      (rec[4], rec[0], rec[1], rec[2], rec[7])
                commit(sql)
            if rec[0] not in self.distinct_settlements:
                sql = "INSERT INTO informalc_settlements " \
                      "(grv_no, level, date_signed, decision, proofdue, docs) " \
                      "VALUES('%s', '%s', '%s', 'monetary remedy', '', '%s')" % \
                      (rec[0], rec[8], rec[3], rec[6])
                commit(sql)
                if rec[5]:  # if there is a gats number
                    # insert the grievance number and gats number into the informalc gats table
                    sql = "INSERT INTO informalc_gats (grv_no, gats_no) VALUES('%s', '%s')" % (rec[0], rec[5])
                    commit(sql)
            i += 1
        sql = "DROP TABLE informalc_grv"
        commit(sql)

    def informalc_awards_xfer(self):
        """ a run a progress bar and transfer recs from informalc awards to informalc awards2 """
        if not self.get_awards_transfer_array():
            return
        self.pb = ProgressBarDe(title="Informal C Data Transfer", label="Transfer all monetary awards data")
        self.pb.max_count(len(self.awards_transfer_array))
        self.pb.start_up()
        self.transfer_award_recs()
        self.pb.stop()

    def get_awards_transfer_array(self):
        """ get an array of information to be moved to new tables. """
        sql = "SELECT * FROM informalc_awards"  # if informalc_awards does not exist - abort transfer
        self.awards_transfer_array = inquire(sql, returnerror=True)  # use kwarg to check for operational error
        if self.awards_transfer_array == "OperationalError":  # if operational error is returned - return False
            return False
        if not self.awards_transfer_array:  # if there is nothing in the results - return False
            return False
        return True

    def transfer_award_recs(self):
        """ insert data into informalc_awards table. when the transfer is complete, delete the informalc_grv table"""
        grv_no, carrier_name, hours, rate, amount = "", "", "", "", ""

        def awards_converter():
            """ convert recs for informalc awards to the format for informalc awards2 """
            award = ""
            if hours and rate:
                award = Convert(hours).hundredths() + "/" + Convert(rate).hundredths()
            if amount:
                award = Convert(amount).hundredths()
            return [grv_no, carrier_name, award, ""]

        i = 1
        for rec in self.awards_transfer_array:
            self.pb.change_text("processing: {}".format(rec[0]))
            self.pb.move_count(i)
            sql = "SELECT * FROM informalc_awards2 WHERE grv_no = '%s' and carrier_name = '%s'" % (rec[0], rec[1])
            result = inquire(sql)  # check if there is a pre existing record.
            if not result:  # if not, then proceed
                grv_no, carrier_name, hours, rate, amount = rec[0], rec[1], rec[2], rec[3], rec[4]
                if grv_no != "None":  # only input if the grv no is not "None"
                    mod_rec = awards_converter()
                    sql = "INSERT INTO informalc_awards2 " \
                          "(grv_no, carrier_name, award, gats_discrepancy) " \
                          "VALUES('%s', '%s', '%s', '%s')" % \
                          (mod_rec[0], mod_rec[1], mod_rec[2], mod_rec[3])
                    commit(sql)
            i += 1
        sql = "DROP TABLE informalc_awards"
        commit(sql)
