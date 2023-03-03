""" This is an auxiliary program for the klusterbox project which merges tables of the mandates.sqlite database"""
import sqlite3
import os
from tqdm import tqdm


class MergeTables:
    """ This is a class for merging tables"""

    def __init__(self):
        self.dbase_receiving = "undefined"  # this is the name of the recieving database
        self.table_receiving = "undefined"  # this is the name of the recieving table
        self.dbase_sending = "undefined"  # this is the name of the sending database
        self.table_sending = "undefined"  # this is the name of the sending table
        self.option = ""  # this is the number of the option from the option menu
        self.dbase_array = []  # this contains a list of tables in the database
        self.path = ""
        self.column_names = []  # this is a list of the column names
        self.sending_data = []  # this is the data fetched from the sending dbase/table
        self.receiving_rec_count = 0  # a count of the recs in the receiving table
        self.checking_sql = ""  # this is the sql for the select/check before merge
        self.insert_sql = ""  # this is the sql for the insert during the merge
        self.distinctids = []  # this is a list of distinct employee ids used for name index cleaner
        self.distinctname_sending = []  # this is a list of distinct names used for for name index cleaner
        self.distinctname_receiving = []  # this is a list of distinct names used for for name index cleaner
        self.doublerecs = []  # this is a list of records where the employee id occurs more than once.
        self.element_counter = []  # this counts the number of non empty elements in the double recs array
        self.max_value = ""  # this is the maxumum value in the element counter
        self.nameindexdeletereport = 0  # this is a count of records deleted by nameindex cleaner.
        self.sending_nameindex = []  # all rows from the name index of the sending table
        self.receiving_nameindex = []  # all rows from the name index of the receiving table

    def run(self):
        self.start_message()
        self.main_menu()

    @staticmethod
    def inquire(path, sql):
        """ query the database """
        db = sqlite3.connect(path)
        cursor = db.cursor()
        try:
            cursor.execute(sql)
            results = cursor.fetchall()
            return results
        except sqlite3.OperationalError:
            print("Database Error\n"
                  "Unable to access database.\n"
                  "Attempted Query: {}".format(sql))
        db.close()

    @staticmethod
    def commit(path, sql):
        """write to the database"""
        db = sqlite3.connect(path)
        cursor = db.cursor()
        try:
            cursor.execute(sql)
            db.commit()
            db.close()
        except sqlite3.OperationalError:
            print("Database Error\n"
                  "Unable to access database.\n"
                  "Attempted Query: {}".format(sql))

    @staticmethod
    def start_message():
        """ this displays a message upon the start of the program in the terminal. """
        text = "Merge Tables \n" \
               "This is a program for merging similar tables in differant databases. \n"
        print(text)

    def main_menu(self):
        """ display a menu of options and get the users choice """
        statustext = "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx\n" \
                     "Recieving database/ table: {}/ {} \n" \
                     "Sending database/ table: {}/ {}\n" \
                     "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx" \
            .format(self.dbase_receiving, self.table_receiving, self.dbase_sending, self.table_sending)
        print(statustext)
        menutext = "\t1. Choose recieving database \n" \
                   "\t2. Choose recieving table \n" \
                   "\t3. Choose sending database \n" \
                   "\t4. Choose sending table \n" \
                   "\t5. Merge \n" \
                   "\t6. Exit \n" \
                   "\t7. Merge all1"
        print(menutext)
        try:
            self.option = input(">>>Enter the number of the option: ")
        except KeyboardInterrupt:
            print("\n")
            self.end()
        self.router()

    def router(self):
        """ this method directs the action based on the option from the main menu """
        if self.option == "1":
            self.get_dbase(True)  # if True - get receiving database
        if self.option == "2":
            self.get_table(True)  # if True - use receiving dbase to get table
        if self.option == "3":
            self.get_dbase(False)  # if False - get sending database
        if self.option == "4":
            self.get_table(False)  # if False - use sending dbase to get table
        if self.option == "5":
            self.merge()
        if self.option == "6":
            self.end()
        if self.option == "7":
            self.merge_all()
        if self.option in ("1", "2", "3", "4", "5", "6", "7"):
            self.main_menu()
        else:
            print("{} is not a valid option. \n\n".format(self.option))
            self.main_menu()

    def get_dbase(self, receiving):
        """ this gets the location of the recieving database using a prompt """
        self.path = self.dbase_path()
        if not self.path:  # if the result of dbase path does not show a database in the chosen directory
            print("\t !!! The selected directory does not contain any sqlite databases...")
            self.path = ""
            self.main_menu()
        self.dbase_select(receiving)

    def dbase_select(self, receiving):
        """ provides the user a list of databases and prompts user to enter a number reflecting that database. """
        i = 1
        for each in self.dbase_array:
            print(i, ": ", each)
            i += 1
        answer = input(">>>Enter number of database you wish to merge: ")
        if answer.lower() in ("q", "x", "quit", "exit", "go back"):
            self.main_menu()
        if not self.isint(answer):
            print("\t !!! {} is not a valid response. ".format(answer))
            self.dbase_select(receiving)
        if not self.isinrange(self.dbase_array, int(answer) - 1):
            print("\t !!! {} is not in the given range of responses. ".format(answer))
            self.dbase_select(receiving)
        if receiving:
            self.dbase_receiving = str(self.path + "/" + self.dbase_array[int(answer) - 1])
        else:
            self.dbase_sending = str(self.path + "/" + self.dbase_array[int(answer) - 1])
        self.dbase_check()  # ensure that sending and receiving dbases are not the same.
        self.main_menu()

    def dbase_check(self):
        """ ensure that the receiving and sending databases are not the same"""
        if self.dbase_receiving == self.dbase_sending:
            print("\t !!! The receiving and sending databases can not be the same.")
            self.dbase_sending = "undefined"

    def dbase_path(self):
        """ get and verify the path to the database. """
        path = input(">>>Enter the path to the database: ")
        if path.lower() in ("q", "x", "quit", "exit", "go back"):
            self.main_menu()
        self.dbase_array = []
        try:
            for x in os.listdir(path):  # get the files in the designated directory
                if x.endswith(".sqlite"):  # look for any sqlite files
                    self.dbase_array.append(x)  # add those files to the dbase_array
        except FileNotFoundError:
            print("\t !!! The directory you designated does not exist! ")
            self.main_menu()
        if len(self.dbase_array):
            return path
        else:
            return False

    def get_table(self, receiving):
        """ this gets the location of the recieving table using a prompt """
        sql = "SELECT name FROM sqlite_master WHERE type='table';"
        if receiving:
            result = self.inquire(self.dbase_receiving, sql)
        else:
            result = self.inquire(self.dbase_sending, sql)
        tables_array = []
        for each in result:
            tables_array.append(each[0])
        i = 1
        for each in tables_array:
            print(i, ": ", each)
            i += 1
        answer = input(">>>Enter the number of the table you want to merge: ")
        if not self.isint(answer):
            print("\t !!! {} is not a valid response. ".format(answer))
            self.get_table(receiving)
        if receiving:
            self.table_receiving = tables_array[int(answer) - 1]
        else:
            self.table_sending = tables_array[int(answer) - 1]
        self.main_menu()

    def merge(self):
        """ this method executes the merge """
        if not self.verifycolumnnames():  # this counts the number of columns in each table
            return
        if not self.verifytablenames():  # this checks that both tables have the same name.
            return
        self.fetchsendingdata()  # get all the data from the sending table as a multidimensional array
        self.checkandinput()

    def verifycolumnnames(self):
        """ this counts the number of columns in each table """
        result_array = []
        dbase_array = (self.dbase_receiving, self.dbase_sending)
        table_array = (self.table_receiving, self.table_sending)
        result = None  # holds the results of the sql search
        for i in range(2):
            sql = "PRAGMA table_info('%s')" % table_array[i]  # get table info. returns an array of columns.
            result = self.inquire(dbase_array[i], sql)
            result_array.append(result)
        if result_array[0] != result_array[1]:
            print("\t !!! Table column names verification: failed")
            print("\t !!! The sending and receiving tables do not have the same column names.")
            return False
        self.column_names = result
        # print("\t >>> Table column names verification: passed")
        return True

    def verifytablenames(self):
        """ this checks that both tables have the same name. """
        if self.table_receiving != self.table_sending:
            print("\t !!! Table names verification: failed")
            print("\t !!! The sending and receiving tables do not have the same name. ")
            return False
        # print("\t >>> Table names verification: passed")
        return True

    def fetchsendingdata(self):
        """ this method gets the data from the sending database """
        sql = "SELECT * FROM '%s'" % self.table_sending
        result = self.inquire(self.dbase_sending, sql)
        self.sending_data = result

    def sql_insert_writer(self, row):
        """ this method writes the sql for the insert """
        i = 0
        sql_insert_columns = "("
        sql_insert_values = "("
        for column in self.column_names:
            sql_insert_columns += column[1]
            sql_insert_values += "'%s'" % self.sending_data[row][i]
            if i != len(self.column_names) - 1:
                sql_insert_columns += ", "
                sql_insert_values += ", "
            i += 1
        sql_insert_columns += ")"
        sql_insert_values += ")"
        self.insert_sql = "INSERT INTO {} {} VALUES{}" \
            .format(self.table_receiving, sql_insert_columns, sql_insert_values)

    def sql_select_writer(self, row):
        """ this method writes the sql for the select """
        i = 0
        sql_select_where = ""
        for column in self.column_names:
            sql_select_where += column[1] + " = '%s'" % self.sending_data[row][i]
            if i != len(self.column_names) - 1:
                sql_select_where += " and "
            i += 1
        self.checking_sql = "SELECT * FROM {} WHERE {}" \
            .format(self.table_receiving, sql_select_where)

    def checkandinput(self):
        """ this performs a check to remove duplicates and enters non duplicates. """
        dups = 0  # a count of duplicate records
        recs = 0  # a count of records inputted into
        print("Initiating checks and inputting data into receiving {} table ... ".format(self.table_receiving))
        for i in tqdm(range(len(self.sending_data))):
            self.sql_select_writer(i)  # this method writes the sql for the select
            result = self.inquire(self.dbase_receiving, self.checking_sql)
            if not result:  # if there is not a record in the receiving db
                self.sql_insert_writer(i)  # this method writes the sql for the insert
                self.commit(self.dbase_receiving, self.insert_sql)  # insert the rec into receiving db
                recs += 1
            else:
                dups += 1  # this counts the number of duplicates
        print("\t records merged: {}\t duplicates rejected: {}".format(recs, dups))

    def merge_all(self):
        """ this cleans duplicates out of the name index table """
        self.nameindexdeletereport = 0  # zero out the count of records deleted
        if not self.nameindex_start():  # check that dbs are defined
            return  # if dbs are not defined - return
        self.get_nameindex()  # get all rows in the name index table from both dbs and place in arrays
        self.nameindex_sort()  # go line by line through name index - take appropriate action
        self.merge_mandates()  # merge all the tables of the mandates database 
        return

    def nameindex_start(self):
        """ this is the introduction and intial check for name index table cleaner tool method """
        self.distinctids = []  # this is a list of distinct employee ids used for name index cleaner
        self.doublerecs = []  # this is a list of records where the employee id occurs more than once.
        print("Name index table cleaner: \n"
              "This will clean the duplicates out of the name index table \n"
              "Set the target database to the receiving database.")
        if self.dbase_receiving == "undefined":
            print("/t !!! Receiving database is not defined. Select option 1 to set receiving database.")
            return False
        if self.dbase_sending == "undefined":
            print("/t !!! Sending database is not defined. Select option 2 to set sending database.")
            return False
        return True

    def get_nameindex(self):
        """ get all rows in the name index table from both dbs and place in arrays """
        sql = "SELECT * FROM name_index"
        self.sending_nameindex = self.inquire(self.dbase_sending, sql)
        self.receiving_nameindex = self.inquire(self.dbase_receiving, sql)

    def nameindex_sort(self):
        """ this will go through the name index of the sending db row by row and determine which action to take
        based on how the row compares to rows in the name index of the receiving database. Index 0 is tacs_name, 
        index 1 is kb_name and index 2 is employee id"""
        for sendrow in self.sending_nameindex:  # send row is a row from the name index from sending db
            # print(sendrow)
            actiontaken = False
            for rcvrow in self.receiving_nameindex:  # rcvrow is a row from the name index from receiving db
                # all criteria match
                if sendrow[0] == rcvrow[0] and sendrow[1] == rcvrow[1] and sendrow[2] == rcvrow[2]:
                    actiontaken = True  # if all match, then do nothing
                # the kb_name and emp_id match, but tacs_name does not match
                elif not actiontaken and sendrow[1] == rcvrow[1] and sendrow[2] == rcvrow[2]:
                    if self.sendrow_greater(sendrow, rcvrow):  # if the sendrow is better data
                        self.delete_from_nameindex(rcvrow)  # delete the row from receiving db
                        self.insert_into_nameindex(sendrow)  # insert sending row into receiving db
                        actiontaken = True
                # the kb_name matches, but the emp_id does not match
                elif not actiontaken and sendrow[1] == rcvrow[1] and sendrow[2] != rcvrow[2]:
                    self.delete_from_nameindex(sendrow)  # delete the row from receiving db by emp id
                    sendrow = self.altername(sendrow)  # change the kb_name in the array
                    self.insert_into_nameindex(sendrow)  # insert the array into the receiving name index
                    self.name_change(sendrow[1], rcvrow[1])  # change name in sending db
                    actiontaken = True
                # if there is an employee id match, but kb name doesn't match
                elif not actiontaken and sendrow[2] == rcvrow[2] and sendrow[1] != rcvrow[1]:
                    actiontaken = True
                else:
                    pass
            if not actiontaken:  # if there are no matches anywhere in the receiving table
                self.insert_into_nameindex(sendrow)

    @staticmethod
    def sendrow_greater(sendrow, rcvrow):
        """ this will count the elements in sendrow/rcvrow to see which has more elements. Returns True if the
        sendrow has more elements that the rcvrow."""
        counter = []
        sendreceive = (sendrow, rcvrow)
        for row in sendreceive:  # go row by row
            count = 0
            for element in row:  # iterate once for each element of the list
                if element:  # if there is something there
                    count += 1  # add one to the count
            counter.append(count)  # add the count to the counter - there should only be two elements when finished
        if counter[0] > counter[1]:  # if the sendrow is greater than the receiverow, then return
            return True
        return False

    def delete_from_nameindex(self, rcvrow):
        """ this will delete the row that is sent to it by the employee id number.  """
        sql = "DELETE FROM name_index WHERE emp_id = '%s'" % rcvrow[2]
        # print(sql)
        self.commit(self.dbase_receiving, sql)

    def insert_into_nameindex(self, sendrow):
        """ this will insert the row that is sent to it.  """
        sql = "INSERT INTO name_index (tacs_name, kb_name, emp_id) " \
              "VALUES('%s','%s','%s')" \
              % (sendrow[0], sendrow[1], sendrow[2])
        # print(sql)
        self.commit(self.dbase_receiving, sql)

    @staticmethod
    def altername(sendrow):
        """ this will add the emp_id to the carrier's kb_name so to make it distinct """
        name = sendrow[1] + " " + sendrow[2][-4:]  # add the last four digits of the employee id number to the name
        return [sendrow[0], name, sendrow[2]]

    def table_exist(self, table):
        """ check to see if the table exist - if so, return True, else False """
        sql = "SELECT name FROM sqlite_master WHERE type='table' AND name='%s'" % table
        result = self.inquire(self.dbase_sending, sql)
        if result:
            return True
        return False

    def name_change(self, newname, oldname):
        """ this will change the name in all relevent tables of the sending database. The name index table has
        already been altered - so do not include it. """
        tables = ("carriers", "informalc_awards", "informalc_payouts", "otdl_preference", "refusals", "rings3",
                  "seniority")
        columns = ("carrier_name", "carrier_name", "carrier_name", "carrier_name", "carrier_name", "carrier_name",
                   "name")
        for i in range(len(tables)):
            if not self.table_exist(tables[i]):  # check to see if the table exist
                print("{} table does not exist in sending db".format(tables[i]))
            else:  # if the table exist
                sql = "SELECT {} FROM {} WHERE {} = '%s'".format(columns[i], tables[i], columns[i]) % oldname
                result = self.inquire(self.dbase_sending, sql)  # look for record
                if result:
                    sql = "UPDATE {} SET {} = '%s' WHERE {} = '%s'".format(tables[i], columns[i], columns[i]) \
                          % (newname, oldname)
                    self.commit(self.dbase_sending, sql)

    def merge_mandates(self):
        """ merge all the tables of the mandates database """
        tables = ("carriers", "dov", "informalc_awards", "informalc_grv", "informalc_payouts", "name_index",
                  "otdl_preference", "refusals", "rings3", "seniority", "station_index", "stations")
        for table in tables:
            self.table_sending = table  # set the sending table to an iteration of tables tuple
            self.table_receiving = table  # set the receiving table to an iteration of the tables tuple
            self.merge()  # run the merge method to check and merge all tables in the tables tuple.

    @staticmethod
    def end():
        """ this displays a message upon the termination of the program in the terminal. """
        text = "Thank you for using Merge Tables \n" \
               "Goodbye. \n"
        print(text)
        quit()

    @staticmethod
    def isint(value):
        """ checks if the argument is an integer"""
        try:
            int(value)
            return True
        except (ValueError, TypeError):
            return False

    @staticmethod
    def isinrange(array, index):
        """ checks if the index is in range"""
        try:
            if not array[index] is "????????":
                return True
        except IndexError:
            return False


if __name__ == "__main__":
    MergeTables().run()
