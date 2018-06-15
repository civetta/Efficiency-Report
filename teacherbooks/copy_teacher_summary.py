from openpyxl.styles import Font
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter

def copy_summary(teacherbook, wb, teachername):
    teacher_summary = teacherbook.create_sheet('Weekly Summary')