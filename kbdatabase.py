"""
a klusterbox module: Klusterbox Database and Platform Variable Setup
This module contains classes and functions to set up the klusterbox database named mandates.sqlite which is located
either in the hidden .klusterbox folder in documents or in th kb_sub folder.
"""
import projvar
from kbtoolbox import inquire, commit, ProgressBarDe
import os


class DataBase:
    """ checks the database for missing tables, columns and values. Enters the same if they are missing. """
    def __init__(self):
        self.pbar_counter = 0
        self.pbar = None

    def setup(self):
        """ checks for database tables and columns then creates them if they do not exist. """
        self.pbar = ProgressBarDe(label="Building Database", text="Starting Up")
        self.pbar.max_count(135)
        self.pbar.start_up()
        self.globals()  # pb increment: 1
        self.tables()  # pb increment: 23
        self.stations()  # pb increment: 1
        self.tolerances()  # pb increment: 45
        self.rings()  # pb increment: 1
        self.skippers()  # pb increment: 1
        self.ns_config()  # pb increment: 6
        self.mousewheel()  # pb increment: 1
        self.list_of_stations()  # pb increment: 1
        self.informalc()  # pb increment: 1
        self.informalc_issues()  # pb increment 35
        self.informalc_decisions()  # pb increment 18
        self.dov()  # pb increment: 1
        self.pbar.stop()

    def globals(self):
        """ defines the project variable for investigation range date week """
        self.pbar_counter += 1
        self.pbar.move_count(self.pbar_counter)
        self.pbar.change_text("Setting up: Global Variables ")
        projvar.invran_date_week = []

    def tables(self):
        """
        Make sure the count of the tables_sql and tables_text match or you will get a list out of sequence error
        """
        tables_sql = (
            'CREATE table IF NOT EXISTS stations (station varchar primary key)',
            'CREATE table IF NOT EXISTS carriers (effective_date date, carrier_name varchar, list_status varchar, '
            'ns_day varchar, route_s varchar, station varchar)',
            'CREATE table IF NOT EXISTS rings3 (rings_date date, carrier_name varchar, total varchar, rs varchar, '
            'code varchar, moves varchar, leave_type varchar, leave_time varchar, refusals varchar, bt varchar, '
            'et varchar)',
            'CREATE table IF NOT EXISTS name_index (tacs_name varchar, kb_name varchar, emp_id varchar)',
            'CREATE table IF NOT EXISTS seniority (name varchar, senior_date varchar)',
            'CREATE table IF NOT EXISTS station_index (tacs_station varchar, kb_station varchar, finance_num varchar)',
            'CREATE table IF NOT EXISTS skippers (code varchar primary key, description varchar)',
            'CREATE table IF NOT EXISTS ns_configuration (ns_name varchar primary key, fill_color varchar, '
            'custom_name varchar)',
            'CREATE table IF NOT EXISTS tolerances (row_id integer primary key, category varchar, tolerance varchar)',
            'CREATE table IF NOT EXISTS otdl_preference (quarter varchar, carrier_name varchar, preference varchar, '
            'station varchar, makeups varchar)',
            'CREATE table IF NOT EXISTS refusals (refusal_date varchar, carrier_name varchar, refusal_type varchar, '
            'refusal_time varchar)',

            'CREATE table IF NOT EXISTS informalc_grv (grv_no varchar, indate_start varchar, indate_end varchar,'
            'date_signed varchar, station varchar, gats_number varchar, docs varchar, description varchar, '
            'level varchar)',
            'CREATE table IF NOT EXISTS informalc_awards (grv_no varchar,carrier_name varchar, hours varchar, '
            'rate varchar, amount varchar)',
            'CREATE table IF NOT EXISTS informalc_payouts(year varchar, pp varchar, payday varchar, '
            'carrier_name varchar, hours varchar, rate varchar, amount varchar)',

            'CREATE table IF NOT EXISTS informalc_grievances(grievant varchar, station varchar, grv_no varchar, '
            'startdate varchar, enddate varchar, meetingdate varchar, issue varchar, article varchar)',

            'CREATE table IF NOT EXISTS informalc_settlements(grv_no varchar, level varchar, date_signed varchar, '
            'decision varchar, proofdue varchar, docs varchar, gats_number varchar)',
            'CREATE table IF NOT EXISTS informalc_remedies (grv_no varchar, carrier_name varchar, hours varchar, '
            'rate varchar, dollars varchar, gatsverified varchar, payoutverified varchar)',
            'CREATE table IF NOT EXISTS informalc_batchindex (main varchar,sub varchar)',
            'CREATE table IF NOT EXISTS informalc_noncindex (settlement varchar, followup varchar)',
            'CREATE table IF NOT EXISTS informalc_remandindex (remanded varchar, followup varchar)',
            'CREATE table IF NOT EXISTS informalc_issuescategories (ssindex, article varchar, '
            'issue varchar primary key, standard boolean)',
            'CREATE table IF NOT EXISTS informalc_decisioncategories (ssindex, type varchar, '
            'decision varchar primary key, standard boolean)',
            'CREATE table IF NOT EXISTS informalc_payout(year varchar, pp varchar, payday varchar, '
            'carrier_name varchar, hours varchar, rate varchar, amount varchar)',
            
            'CREATE table IF NOT EXISTS dov(eff_date date, station varchar, day varchar, dov_time varchar, '
            'temp varchar)'
        )

        tables_text = (
            "Setting up: Tables - Station",
            "Setting up: Tables - Carriers",
            "Setting up: Tables - Rings",
            "Setting up: Tables - Name Indexes",
            "Setting up: Tables - Seniority",
            "Setting up: Tables - Station Indexes",
            "Setting up: Tables - Skippers",
            "Setting up: Tables - NS Configurations",
            "Setting up: Tables - Tolerances...",
            "Setting up: Tables - OTDL Preference",
            "Setting up: Tables - Refusals",
            "Setting up: Tables - Informal C",
            "Setting up: Tables - Informal C Awards",
            "Setting up: Tables - Informal C Payouts",
            "Setting up: Tables - Informal C Grievances",
            "Setting up: Tables - Informal C Settlements",
            "Setting up: Tables - Informal C Remedies",
            "Setting up: Tables - Informal C Batch Index",
            "Setting up: Tables - Informal C Noncompliance Index",
            "Setting up: Tables - Informal C Remand Index",
            "Setting up: Tables - Informal C Issue Categories",
            "Setting up: Tables - Informal C Decision Categories",
            "Setting up: Tables - Informal C Payout",
            "Setting up: Tables - DOV"
        )
        for i in range(len(tables_sql)):
            self.pbar_counter += 1
            self.pbar.move_count(self.pbar_counter)
            self.pbar.change_text(tables_text[i])
            commit(tables_sql[i])

    def stations(self):
        """ checks for columns in station table and creates out of station value if it doesn't exist. """
        self.pbar_counter += 1
        self.pbar.move_count(self.pbar_counter)
        self.pbar.change_text("Setting up: Tables - Station > add out of station")
        sql = 'INSERT OR IGNORE INTO stations (station) VALUES ("out of station")'
        commit(sql)

    def tolerances(self):
        """ checks the tolerances table and inputs values if they do not exist. """
        tolerance_array = (
            (0, "ot_own_rt", 0),
            (1, "ot_tol", 0),
            (2, "av_tol", 0),
            (3, "min_ss_nl", 25),
            (4, "min_ss_wal", 25),
            (5, "min_ss_otdl", 25),
            (6, "min_ss_aux", 25),
            (7, "allow_zero_top", "False"),  # obsolete
            (8, "allow_zero_bottom", "True"),  # obsolete
            (9, "pdf_error_rpt", "off"),
            (10, "pdf_raw_rpt", "off"),
            (11, "pdf_text_reader", "off"),
            (12, "ns_auto_pref", "rotation"),
            (13, "mousewheel", -1),
            (14, "min_ss_overmax", 30),
            (15, "abc_breakdown", "False"),
            (16, "min_spd_empid", 50),
            (17, "min_spd_alpha", 50),
            (18, "min_spd_abc", 10),
            (19, "speedcell_ns_rotate_mode", "True"),
            (20, "ot_rings_limiter", 0),
            (21, "pb_nl_wal", "True"),
            (22, "pb_wal_otdl", "True"),
            (23, "pb_otdl_aux", "True"),
            (24, "invran_mode", "simple"),
            (25, "min_ot_equit", 19),
            (26, "ot_calc_pref", "off_route"),
            (27, "min_ot_dist", 25),
            (28, "ot_calc_pref_dist", "off_route"),
            (29, "tourrings", 0),
            (30, "spreadsheet_pref", "Mandates"),
            (31, "lastfix", "1.000"),
            (32, "min4_ss_nl", 19),
            (33, "min4_ss_wal", 19),
            (34, "min4_ss_otdl", 19),
            (35, "min4_ss_aux", 19),
            (36, "pb4_nl_wal", "True"),
            (37, "pb4_wal_aux", "True"),
            (38, "pb4_aux_otdl", "True"),
            (39, "man4_dis_limit", "show all"),
            (40, "speedsheets_fullreport", "False"),
            (41, "offbid_distinctpage", "True"),
            (42, "offbid_maxpivot", 2.0),
            (43, "triad_routefirst", "False"),  # when False, the route is displayed at the end of the route triad
            (44, "wal_12_hour", "True"),  # when True, wal 12/60 violations happen after 12 hr, else after 11.50 hrs
            (45, "wal_dec_exempt", "False")
            # increment self.pbar.max_count() in self.setup() if you add more records.
        )
        for tol in tolerance_array:
            self.pbar_counter += 1
            self.pbar.move_count(self.pbar_counter)
            self.pbar.change_text("Setting up: Tables - Tolerances {}".format(tol[1]))
            sql = 'INSERT OR IGNORE INTO tolerances (row_id, category, tolerance) ' \
                  'VALUES ("%s", "%s", "%s")' % (tol[0], tol[1], tol[2])
            commit(sql)

    def rings(self):
        """ sets up the rings table """
        self.pbar_counter += 1
        self.pbar.move_count(self.pbar_counter)
        self.pbar.change_text("Setting up: Tables - Rings > leave time/type")
        # modify table for legacy version which did not have leave type and leave time columns of
        # table.
        sql = 'PRAGMA table_info(rings3)'  # get table info. returns an array of columns.
        result = inquire(sql)
        if len(result) < 7:  # if there are not enough columns, add the leave type column
            sql = 'ALTER table rings3 ADD COLUMN leave_type varchar'
            commit(sql)
        if len(result) < 8:  # if there are not enough columns, add the leave time column
            sql = 'ALTER table rings3 ADD COLUMN leave_time varchar'
            commit(sql)
        if len(result) < 9:  # if there are not enough columns, add the refusals column
            sql = 'ALTER table rings3 ADD COLUMN refusals varchar'
            commit(sql)
        if len(result) < 10:  # if there are not enough columns, add the bt column
            sql = 'ALTER table rings3 ADD COLUMN bt varchar'
            commit(sql)
        if len(result) < 11:  # if there are not enough columns, add the et column
            sql = 'ALTER table rings3 ADD COLUMN et varchar'
            commit(sql)

    def skippers(self):
        """ put records in the skippers table """
        skip_these = (("354", "stand by"), ("613", "stewards time"), ("743", "route maintenance"))
        for rec in skip_these:
            self.pbar_counter += 1
            self.pbar.move_count(self.pbar_counter)
            self.pbar.change_text("Setting up: Tables - Skippers > {}".format(rec[0]))
            sql = 'INSERT OR IGNORE INTO skippers(code, description) VALUES ("%s", "%s")' % (rec[0], rec[1])
            commit(sql)

    def ns_config(self):
        """ sets rotating non scheduled days if not set """
        ns_sql = (
            ("yellow", "gold", "yellow"),
            ("blue", "navy", "blue"),
            ("green", "forest green", "green"),
            ("brown", "saddle brown", "brown"),
            ("red", "red3", "red"),
            ("black", "gray10", "black")
        )
        for ns in ns_sql:
            self.pbar_counter += 1
            self.pbar.move_count(self.pbar_counter)
            self.pbar.change_text("Setting up: Tables - NS Configurations {}".format(ns[0]))
            sql = 'INSERT OR IGNORE INTO ns_configuration(ns_name,fill_color,custom_name)VALUES("%s", "%s", "%s")'\
                  % (ns[0], ns[1], ns[2])
            commit(sql)

    def mousewheel(self):
        """ initialize mousewheel - mouse wheel scroll direction """
        self.pbar_counter += 1
        self.pbar.move_count(self.pbar_counter)
        self.pbar.change_text("Setting up: Mousewheel")
        sql = "SELECT tolerance FROM tolerances WHERE category = '%s'" % "mousewheel"
        results = inquire(sql)
        projvar.mousewheel = int(results[0][0])

    def list_of_stations(self):
        """ sets up the list of stations project variable. """
        self.pbar_counter += 1
        self.pbar.move_count(self.pbar_counter)
        self.pbar.change_text("Setting up: List of Stations")
        sql = "SELECT * FROM stations ORDER BY station"
        results = inquire(sql)
        # define and populate list of stations variable
        projvar.list_of_stations = []
        for stat in results:
            projvar.list_of_stations.append(stat[0])

    def informalc(self):
        """ check the informalc table and make changes if needed"""
        self.pbar_counter += 1
        self.pbar.move_count(self.pbar_counter)
        self.pbar.change_text("Setting up: Tables - Informal C > update ")
        # modify table for legacy version which did not have level column of informalc_grv table.
        sql = 'PRAGMA table_info(informalc_grv)'  # get table info. returns an array of columns.
        result = inquire(sql)
        if len(result) <= 8:  # if there are not enough columns add the leave type and leave time columns
            sql = 'ALTER table informalc_grv ADD COLUMN level varchar'
            commit(sql)

    def informalc_issues(self):
        """  set up the standard issues for informal c issue categories
        fields are 'ssindex'(speedsheet index), 'article', 'issue' and 'standard' """
        issues = (
            ("1", "2", "discrimination", True),
            ("2", "5", "unilateral action", True),
            ("3", "8", "improper mandating", True),
            ("4", "8", "12/60 hour violations", True),
            ("5", "8", "otdl equitability", True),
            ("6", "8", "out of schedule pay", True),
            ("7", "8", "schedule change", True),
            ("8", "10", "denied leave", True),
            ("9", "10", "improper awol", True),
            ("10", "10", "improper annual leave", True),
            ("11", "10", "denied sick leave", True),
            ("12", "10", "denied annual leave", True),
            ("13", "10", "medical documentation", True),
            ("14", "11", "denied holiday pay", True),
            ("15", "11", "holiday scheduling", True),
            ("16", "12", "denied transfer", True),
            ("17", "13", "denied reassignment", True),
            ("18", "13", "denied accommodation", True),
            ("19", "14", "health and safety", True),
            ("20", "15", "failure to meet", True),
            ("21", "15", "non compliance", True),
            ("22", "16", "discipline", True),
            ("23", "16", "letter of warning", True),
            ("24", "16", "suspension", True),
            ("25", "16", "removal", True),
            ("26", "16", "emergency placement", True),
            ("27", "17", "stewards rights", True),
            ("28", "17", "denied information", True),
            ("29", "17", "denied time", True),
            ("30", "17", "weingarten violation", True),
            ("31", "26", "uniform allowance", True),
            ("32", "41", "off bid violation", True),
            ("33", "41", "denied opt", True),
            ("34", "41", "posting violation", True),
            ("35", "41", "improper hold down", True)  # 35 count
        )
        # increment self.pbar.max_count() in self.setup() if you add more records.
        for iss in issues:
            self.pbar_counter += 1
            self.pbar.move_count(self.pbar_counter)
            self.pbar.change_text("Setting up: Tables - Informal C Issues {}".format(iss[1]))
            sql = 'INSERT OR IGNORE INTO informalc_issuescategories (ssindex, article, issue, standard) ' \
                  'VALUES ("%s", "%s", "%s", "%s")' % (iss[0], iss[1], iss[2], iss[3])
            commit(sql)

    def informalc_decisions(self):
        """  writes the standard decisions to the informal c table for decisions
        the columns are 'ssindex' (speeed sheet index), type (of grienance), 'decision' and 'standard'. """
        decisions = (
            ("1", "general", "favorable", True),
            ("2", "general", "unfavorable", True),
            ("3", "general", "monetary remedy", True),
            ("4", "general", "adjustment", True),
            ("5", "general", "language", True),
            ("6", "general", "cease and desist", True),
            ("7", "general", "withdrawn", True),
            ("8", "general", "no violation", True),
            ("9", "general", "moot", True),
            ("10", "general", "remanded", True),
            ("11", "general", "???", True),
            ("12", "general", "bullshit", True),
            ("13", "16", "expunged", True),
            ("14", "16", "discussion", True),
            ("15", "16", "limited retention", True),
            ("16", "16", "time served", True),
            ("17", "16", "back pay", True),
            ("18", "16", "sustained", True)  # 18 count
        )
        # increment self.pbar.max_count() in self.setup() if you add more records.
        for des in decisions:
            self.pbar_counter += 1
            self.pbar.move_count(self.pbar_counter)
            self.pbar.change_text("Setting up: Tables - Informal C Issues {}".format(des[1]))
            sql = 'INSERT OR IGNORE INTO informalc_decisioncategories (ssindex, type, decision, standard) ' \
                  'VALUES ("%s", "%s", "%s", "%s")' % (des[0], des[1], des[2], des[3])
            commit(sql)

    def dov(self):
        """ check if minimum records are in the dov table. These are seven recs, one for each day, for the
        year 1. This ensures that there is always a record in the dov table. the default value is hard coded. """
        self.pbar_counter += 1
        self.pbar.move_count(self.pbar_counter)
        for station in projvar.list_of_stations:  # for each station with a record in station table
            self.pbar.change_text("Setting up: Tables - DOV > {} default values".format(station))
            if station != "out of station":
                DovBase().minimum_recs(station)  # make sure the minimum recs are in the DOV table.


def setup_plaformvar():
    """ set up platform variable """
    projvar.platform = "py"  # initialize projvar.platform variable
    split_home = os.getcwd().split("\\")
    if os.path.isdir('Applications/klusterbox.app') and os.getcwd() == "/":  # if it is a mac app
        projvar.platform = "macapp"
    elif len(split_home) > 2:
        if split_home[1] == "Program Files (x86)" and split_home[2] == "klusterbox":
            projvar.platform = "winapp"
        elif split_home[1] == "Program Files" and split_home[2] == "klusterbox":
            projvar.platform = "winapp"
        else:
            projvar.platform = "py"  # if it is running as a .py or .exe outside program files/applications
    else:
        projvar.platform = "py"  # if it is running as a .py or .exe outside program files/applications


def setup_dirs_by_platformvar():
    """ create directories if they don't exist """
    if projvar.platform == "macapp":
        if not os.path.isdir(os.path.join(os.path.sep, os.path.expanduser("~"), 'Documents')):
            os.makedirs(os.path.join(os.path.sep, os.path.expanduser("~"), 'Documents'))
        if not os.path.isdir(os.path.join(os.path.sep, os.path.expanduser("~"), 'Documents', 'klusterbox')):
            os.makedirs(os.path.join(os.path.sep, os.path.expanduser("~"), 'Documents', 'klusterbox'))
        if not os.path.isdir(os.path.join(os.path.sep, os.path.expanduser("~"), 'Documents', '.klusterbox')):
            os.makedirs(os.path.join(os.path.sep, os.path.expanduser("~"), 'Documents', '.klusterbox'))
    if projvar.platform == "winapp":
        if not os.path.isdir(os.path.expanduser("~") + '\\Documents'):
            os.makedirs(os.path.expanduser("~") + '\\Documents')
        if not os.path.isdir(os.path.expanduser("~") + '\\Documents\\klusterbox'):
            os.makedirs(os.path.expanduser("~") + '\\Documents\\klusterbox')
        if not os.path.isdir(os.path.expanduser("~") + '\\Documents\\.klusterbox'):
            os.makedirs(os.path.expanduser("~") + '\\Documents\\.klusterbox')
    if projvar.platform == "py":
        if not os.path.isdir('kb_sub'):
            os.makedirs('kb_sub')


class DovBase:
    """
    deals with DOV table in the mandates.sql database
    """

    def __init__(self):
        self.station = None
        self.day = ("sat", "sun", "mon", "tue", "wed", "thu", "fri")
        self.defaulttime = "20.50"

    def minimum_recs(self, station):
        """ places 7 records dated year 1 in the DOV table so that the database always has a default record
         for any day. """
        self.station = station
        for day in self.day:
            # check if the minimum record is in the database/ there is one for each day
            sql = "SELECT * FROM dov WHERE station = '%s' AND eff_date = '%s' AND day = '%s'" % \
                  (self.station, "0001-01-01 00:00:00", day)
            result = inquire(sql)
            if not result:
                # if the minimum record is not in the database, then add it.
                sql = "INSERT INTO dov (eff_date, station, day, dov_time, temp) " \
                      "VALUES('%s', '%s', '%s', '%s', '%s')" % \
                      ("0001-01-01 00:00:00", self.station, day, self.defaulttime, False)
                commit(sql)
