#!usr/bin/python

import argparse
import logging
import openpyxl
import os
import re

from openpyxl.styles import (Alignment, Color, fills, PatternFill)


class Course:

    def __init__(self):
        self.abbrev = None
        self.name = None
        self.tutor = None
        self.location = None
        self.count = 0
        self.count_to_method = {0:self.set_abbrev, 1:self.set_name,
                              2:self.set_tutor, 3:self.set_location}

    def record(self, entry):
        self.count_to_method[self.count](entry)

    def set_abbrev(self, abbrev):
        self.abbrev = abbrev
        self.count += 1

    def set_name(self, name):
        self.name = name
        self.count += 1

    def set_tutor(self, tutor):
        self.tutor = tutor
        self.count += 1

    def set_location(self, loc):
        self.location = loc
        self.count += 1

def fuck_timetable(args):

    """Assumes info box in top RH corner is bounded on L & R side
    by cells containing only the word CUT and END_CUT, respectively."""

    course_list = dict()
    in_cut = False

    sheet_number = args.year
    if sheet_number == 4:
        sheet_number = 3

    timetable = openpyxl.load_workbook(os.path.abspath(args.path))
    year_sheet = timetable[timetable.get_sheet_names()[sheet_number-1]]

    for row in year_sheet:
        for cell in row:

            if not isinstance(cell.value, basestring):
                continue

            cell_contains = {val.strip('([ ])') for val in cell.value.split()}
            if 'CUT' in cell_contains:
                in_cut = True
                cell.value = None
                course = Course()
                continue
            if 'END_CUT' in cell_contains:
                in_cut = False
                cell.value = None
                course_list[course.abbrev] = course
                continue
            if any(['Week' in cell_contains, 'Time' in cell_contains,
                    'Mon' in cell_contains, 'Tues' in cell_contains,
                    'Wed' in cell_contains, 'Thur' in cell_contains,
                    'Fri' in cell_contains,
                    re.search('[0-9]+\-[0-9]+', cell.value)]):
                continue
            if in_cut:
               derryck.info('Recording {0}'.format(cell.value))
               course.record(cell.value)
               continue

            if not isinstance(args.courses, list):
                courses = set([unicode(args.courses)])
            courses = set(unicode(course) for course in args.courses)
            courses_in_cell = courses.intersection(cell_contains)
            if len(courses_in_cell) > 0:
                cell.value = ' '.join(courses_in_cell)
                cell.alignment = Alignment(horizontal='center')
                continue
            cell.value = None
            cell.fill = PatternFill(patternType=fills.FILL_SOLID,
                                    fgColor=Color('FFFFFF'))

    timetable.save(os.path.abspath(args.path))
    derryck.info(course_list)
    return course_list

if __name__ == '__main__':

    derryck = logging.getLogger('Derryck')
    derryck.addHandler(logging.StreamHandler())
    derryck.setLevel(logging.INFO)

    desc_str = """Removes all but specified courses from Derryck-issued
                  timetables, doing so on a row-by-row basis. Cells in a given
                  row bounded by CUT and END_CUT will be ignored. Requires
                  openpyxl."""
    ep_str = 'Example: "python <path> 4 -c GR QO" removes all but GR, QO.'
    parser = argparse.ArgumentParser(description=desc_str, epilog=ep_str)
    parser.add_argument('path', help='Path to timetable.')
    parser.add_argument('year', type=int, help='Year of study.')
    parser.add_argument('--courses', '-c', metavar='C', nargs='+',
                        help='Relevant courses (abbrevs).')
    args = parser.parse_args()
    course_list = fuck_timetable(args)
