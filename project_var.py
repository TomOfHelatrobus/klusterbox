"""
holds all project variables for the klusterbox project
call "import project_var" in any file using these variables
to call:   x = project_var.case_year
to set:    project_var.case_year = x
"""
case_year = None  # the year of the investigation range
case_month = None  # the month of the investigation range
case_day = None  # the day of the investigation range
case_weekly_span = True  # the span of the investigation True - weekly, False - daily
case_station = None  # the name of the station
case_date_week = []  # a list of seven days in the investigation range
case_date = None  # the day of a daily investigation
list_of_stations = []  # list of all statons
ns_code = {}  # dictionary of ns days
pay_period = None  # the pay period
platform = "py"  # the platform that klusterbox is running on
mousewheel = -1  # configure the mousewheel for natural - "1" or reverse - "-1" scrolling

