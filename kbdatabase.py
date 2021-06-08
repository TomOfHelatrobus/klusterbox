import projvar
from kbtoolbox import titlebar_icon, inquire, commit, ProgressBarDe
import os


class DataBase:
    def __init__(self):
        self.pbar_counter = 0
        self.pbar = None

    def setup(self):
        self.pbar = ProgressBarDe(title="Building Klusterbox Database", label="Running Setup: ", text="Starting Up")
        self.pbar.max_count(45)
        self.pbar.start_up()
        self.globals()
        self.tables()
        self.stations()
        self.tolerances()
        self.rings()
        self.ns_config()
        self.mousewheel()
        self.list_of_stations()
        self.pbar.stop()

    def globals(self):
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
            'code varchar, moves varchar, leave_type varchar, leave_time varchar)',
            'CREATE table IF NOT EXISTS name_index (tacs_name varchar, kb_name varchar, emp_id varchar)',
            'CREATE table IF NOT EXISTS station_index (tacs_station varchar, kb_station varchar, finance_num varchar)',
            'CREATE table IF NOT EXISTS skippers (code varchar primary key, description varchar)',
            'CREATE table IF NOT EXISTS ns_configuration (ns_name varchar primary key, fill_color varchar, '
            'custom_name varchar)',
            'CREATE table IF NOT EXISTS tolerances (row_id integer primary key, category varchar, tolerance varchar)'
        )

        tables_text = (
            "Setting up: Tables - Station",
            "Setting up: Tables - Carriers",
            "Setting up: Tables - Rings",
            "Setting up: Tables - Name Indexes",
            "Setting up: Tables - Station Indexes",
            "Setting up: Tables - Skippers",
            "Setting up: Tables - NS Configurations",
            "Setting up: Tables - Tolerances..."
        )
        for i in range(len(tables_sql)):
            self.pbar_counter += 1
            self.pbar.move_count(self.pbar_counter)
            self.pbar.change_text(tables_text[i])
            commit(tables_sql[i])

    def stations(self):
        self.pbar_counter += 1
        self.pbar.move_count(self.pbar_counter)
        self.pbar.change_text("Setting up: Tables - Station > add out of station")
        sql = 'INSERT OR IGNORE INTO stations (station) VALUES ("out of station")'
        commit(sql)

    def tolerances(self):
        tolerance_array = (
            (0, "ot_own_rt", .25),
            (1, "ot_tol", .25),
            (2, "av_tol", .25),
            (3, "min_ss_nl", 25),
            (4, "min_ss_wal", 25),
            (5, "min_ss_otdl", 25),
            (6, "min_ss_aux", 25),
            (7, "allow_zero_top", "False"),
            (8, "allow_zero_bottom", "True"),
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
            (24, "invran_mode", "simple")
        )
        for tol in tolerance_array:
            self.pbar_counter += 1
            self.pbar.move_count(self.pbar_counter)
            self.pbar.change_text("Setting up: Tables - Tolerances {}".format(tol[1]))
            sql = 'INSERT OR IGNORE INTO tolerances (row_id, category, tolerance) ' \
                  'VALUES ("%s", "%s", "%s")' % (tol[0], tol[1], tol[2])
            commit(sql)

    def rings(self):
        self.pbar_counter += 1
        self.pbar.move_count(self.pbar_counter)
        self.pbar.change_text("Setting up: Tables - Rings > leave time/type")
        # modify table for legacy version which did not have leave type and leave time columns of rings3 table.
        sql = 'PRAGMA table_info(rings3)'  # get table info. returns an array of columns.
        result = inquire(sql)
        if len(result) <= 6:  # if there are not enough columns add the leave type and leave time columns
            sql = 'ALTER table rings3 ADD COLUMN leave_type varchar'
            commit(sql)
            sql = 'ALTER table rings3 ADD COLUMN leave_time varchar'
            commit(sql)

    def skippers(self):
        # put records in the skippers table
        skip_these = (("354", "stand by"), ("613", "stewards time"), ("743", "route maintenance"))
        for rec in skip_these:
            self.pbar_counter += 1
            self.pbar.move_count(self.pbar_counter)
            self.pbar.change_text("Setting up: Tables - Skippers > {}".format(rec[0]))
            sql = "INSERT OR IGNORE INTO skippers(code, description) VALUES ('%s','%s')" % (rec[0], rec[1])
            commit(sql)

    def ns_config(self):
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

    def mousewheel(self):  # initialize mousewheel - mouse wheel scroll direction
        self.pbar_counter += 1
        self.pbar.move_count(self.pbar_counter)
        self.pbar.change_text("Setting up: Mousewheel")
        sql = "SELECT tolerance FROM tolerances WHERE category = '%s'" % "mousewheel"
        results = inquire(sql)
        projvar.mousewheel = int(results[0][0])

    def list_of_stations(self):
        self.pbar_counter += 1
        self.pbar.move_count(self.pbar_counter)
        self.pbar.change_text("Setting up: List of Stations")
        sql = "SELECT * FROM stations ORDER BY station"
        results = inquire(sql)
        # define and populate list of stations variable
        projvar.list_of_stations = []
        for stat in results:
            projvar.list_of_stations.append(stat[0])


def setup_plaformvar():
    # set up platform variable
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
    # create directories if they don't exist
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
