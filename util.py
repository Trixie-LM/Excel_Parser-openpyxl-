import time
import re


def _get_cell_value(char_column, number_line, sheet):
    return sheet[char_column + str(number_line)].value


def get_boundary_values(diapason, value):
    group_id = 0
    if value == 'min_column':
        group_id = 1
    elif value == 'min_row':
        group_id = 2
    elif value == 'max_column':
        group_id = 3
    elif value == 'max_row':
        group_id = 4
    value = re.search(r'(\D)(\d+):(\D)(\d+)', diapason).group(group_id)
    return value


class Timer:
    def __init__(self):
        self._start_time: float = time.time()
        self._last_time: float = self._start_time
        self._times: dict[str, float] = {}

    def tick(self, key: str):
        tm = time.time()
        self._times[key] = tm - self._last_time
        self._last_time = tm

    @property
    def total(self) -> float:
        return self._last_time - self._start_time

    @property
    def details(self) -> str:
        return ', '.join('%s: %.1f секунд' % (key, tm) for key, tm in self._times.items())

    @property
    def info(self) -> str:
        return '%.1f секунд\n(%s)' % (time.time() - self._start_time, self.details)
