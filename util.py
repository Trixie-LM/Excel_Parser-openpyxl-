from openpyxl import load_workbook


def get_cell_value(char_column, number_line, sheet):
    return sheet[char_column + str(number_line)].value
