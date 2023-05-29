"""
a klusterbox module: Project Variables
holds all project variables for the klusterbox project
call "import projvar" in any file using these variables
to call:   x = projvar.invran_year
to set:    projvar.invran_year = x
"""

root = None  # tkinter root
invran_year = None  # the year of the investigation range
invran_month = None  # the month of the investigation range
invran_day = None  # the day of the investigation range
invran_weekly_span = None  # the span of the investigation True - weekly, False - daily
invran_station = None  # the name of the station
invran_date_week = []  # a list of seven days in the investigation range
invran_date = None  # the day of a daily investigation
list_of_stations = []  # list of all stations
ns_code = {}  # dictionary of ns days
pay_period = None  # the pay period
platform = "py"  # the platform that klusterbox is running on
mousewheel = -1  # configure the mousewheel for natural - "1" or reverse - "-1" scrolling
try_absorber = False  # a variable that absorbs errors in try/except statements.
