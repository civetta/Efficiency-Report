from openpyxl.utils import get_column_letter
from openpyxl.styles import Font
from openpyxl.styles import PatternFill
from copy import copy


def copier(old_cell, new_cell):
    '''Copies both the value, font style, and fill from old cell
    into the new cell'''
    new_cell.value = old_cell.value
    new_cell.font = copy(old_cell.font)
    new_cell.fill = copy(old_cell.fill)