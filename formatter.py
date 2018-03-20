import openpyxl
from openpyxl.utils import get_column_letter
import re
from datetime import datetime
from openpyxl.styles import Font
from openpyxl.styles import colors
from openpyxl.styles import alignment
import numpy as np
from openpyxl.chart import BarChart, Series, Reference
from openpyxl.styles import Border, Side, PatternFill, Font, GradientFill, Alignment


def formatter(wb):
    ws = wb.get_sheet_by_name("Summary")
