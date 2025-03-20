from datetime import datetime


class Operation:
    def __init__(self):
        self.__name = None  # (str) ExcelDataLoader yükler
        self.__compatible_jigs = None  # ExcelDataLoader yükler
        self.__required_skills = None  # (str) ExcelDataLoader yükler
        self.__required_man_hours = None  # (float) ExcelDataLoader yükler
        self.__min_workers = None  # (int) ExcelDataLoader yükler
        self.__max_workers = None  # (int)ExcelDataLoader yükler
        self.__predecessors = None  # (list of operation  objects) ExcelDataLoader yükler
        self.__uncompleted_predecessors = None
        self.__previous_operations = []
        self.__successors = None  # (list of operation names)
        self.__completed = False  # (bool if completed: True) AddProgress yükler
        self.__assigned_jig = None  # (Holds jig object)
        self.__start_datetime = None  # (date, time)
        self.__end_datetime = None  # (date, time)
        self.__assigned_workers = None  # (List holding worker objects)
        self.__required_worker = None  # (int) Calculated and set by Maincontroller called in read_excel
        self.__operating_duration = None  # (float) Calculated and set by Maincontroller called in add product in mainscreen
        self.__early_start = 0
        self.__early_finish = 0
        self.__late_start = float('inf')
        self.__late_finish = float('inf')
        self.__slack = None

    def set_name(self, _name):
        self.__name = _name

    def get_name(self):
        return self.__name

    def set_compatible_jigs(self, _compatible_jigs):
        self.__compatible_jigs = _compatible_jigs

    def get_compatible_jigs(self):
        return self.__compatible_jigs

    def set_required_skills(self, _required_skills):
        self.__required_skills = _required_skills

    def get_required_skills(self):
        return self.__required_skills

    def set_required_man_hours(self, _required_man_hours):
        self.__required_man_hours = _required_man_hours

    def get_required_man_hours(self):
        return self.__required_man_hours

    def set_min_workers(self, _min_workers):
        self.__min_workers = _min_workers

    def get_min_workers(self):
        return self.__min_workers

    def set_max_workers(self, _max_workers):
        self.__max_workers = _max_workers

    def get_max_workers(self):
        return self.__max_workers

    def set_predecessors(self, _predecessors):
        self.__predecessors = _predecessors

    def get_predecessors(self):
        return self.__predecessors

    def set_uncompleted_predecessors(self, _uncompleted_predecessors):
        self.__uncompleted_predecessors = _uncompleted_predecessors

    def get_uncompleted_predecessors(self):
        return self.__uncompleted_predecessors

    def set_previous_operations(self, _prevOps):
        self.__previous_operations = _prevOps

    def get_previous_operations(self):
        return self.__previous_operations

    def set_successors(self, _successors):
        self.__successors = _successors

    def get_successors(self):
        return self.__successors

    def remove_predecessor(self, _pre_name):
        for predecessor in self.__predecessors:
            if predecessor == _pre_name:
                self.__predecessors.remove(predecessor)

    def set_completed(self, _completed):
        self.__completed = _completed

    def get_completed(self):
        return self.__completed

    def set_assigned_jig(self, _jig):
        self.__assigned_jig = _jig

    def get_assigned_jig(self):
        return self.__assigned_jig

    def set_required_worker(self, _shiftWorker):
        self.__required_worker = _shiftWorker

    def get_required_worker(self):
        return self.__required_worker

    def set_operating_duration(self, _operatingDuration):
        self.__operating_duration = _operatingDuration

    def get_operating_duration(self):
        return self.__operating_duration

    def get_early_start(self):
        return self.__early_start

    def set_early_start(self, _es):
        self.__early_start = _es

    def set_late_finish(self, _lf):
        self.__late_finish = _lf

    def get_late_finish(self):
        return self.__late_finish

    def set_late_start(self, _es):
        self.__early_start = _es

    def get_late_start(self):
        return self.__early_start

    def set_early_finish(self, _lf):
        self.__late_finish = _lf

    def get_early_finish(self):
        return self.__late_finish

    def set_slack(self, _sl):
        self.__slack = _sl

    def get_slack(self):
        return self.__slack

    def set_start_datetime(self, date, time):
        self.__start_datetime = date, time

    def get_start_datetime(self):
        return self.__start_datetime

    def set_end_datetime(self, date, time):
        self.__end_datetime = date, time

    def get_end_datetime(self):
        return self.__end_datetime