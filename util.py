import re


def get_cell_value(char_column, number_line, sheet):
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
