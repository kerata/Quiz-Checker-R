#!/usr/local/bin/python3

from os.path import join, dirname, abspath
import argparse
import xlrd

from sheet_helper import connect_GSheets, spreadsheet_id
from sheet_helper import section1_sheet_id, section2_sheet_id
from sheet_helper import color_default, color_wrong

from other.student_info import section1_grade, section2_grade, one2two, two2one


def create_student_grades(file_name, id_col='StudentID'):
    if not file_name.endswith('.xlsx'):
        file_name = join(dirname(abspath(__file__)), file_name + '.xlsx')
    # Open the workbook
    xl_workbook = xlrd.open_workbook(file_name)

    # List sheet names, and pull a sheet by name
    # sheet_names = xl_workbook.sheet_names()
    # xl_sheet = xl_workbook.sheet_by_name(sheet_names[0])

    # Or grab the first sheet by index
    #  (sheets are zero-indexed)
    #
    xl_sheet = xl_workbook.sheet_by_index(0)

    # Pull the first row by index
    #  (rows/columns are also zero-indexed)
    #
    row = xl_sheet.row(0)  # 1st row

    # Print 1st row values and types
    #
    from xlrd.sheet import ctype_text

    studentIDColInx = 0
    studentScoreColInx = 0
    for idx, cell_obj in enumerate(row):
        cell_type_str = ctype_text.get(cell_obj.ctype, 'unknown type')
        if cell_obj.value == id_col:
            studentIDColInx = idx
        elif cell_obj.value == "Score":
            studentScoreColInx = idx

    student_grades = {}
    # num_cols = xl_sheet.ncols  # Number of columns
    for row_idx in range(1, xl_sheet.nrows):  # Iterate through rows
        try:
            cell_obj = xl_sheet.cell(row_idx, studentIDColInx)
            studentID = str(int(cell_obj.value))
            cell_obj = xl_sheet.cell(row_idx, studentScoreColInx)
            studentScore = int(cell_obj.value)
            student_grades[studentID] = studentScore
        except ValueError:
            print('{0} -> {1}'.format(file_name, row_idx))
    return student_grades


def dump_attendance(aliens, no_grade):
    print('Section 1')
    for studentID, [grade, isTrue] in sorted(section1_grade.items(), reverse=True):
        print('{1}\t{0}{2}'.format(grade if isTrue else 0, studentID,
            '' if isTrue else '\t{}\tWrong Section'.format(grade)))

    print('\nSection 2')
    for studentID, [grade, isTrue] in sorted(section2_grade.items(), reverse=True):
        print('{1}\t{0}{2}'.format(grade if isTrue else 0, studentID,
            '' if isTrue else '\t{}\tWrong Section'.format(grade)))

    print('\nAliens')
    print('StudentID\t{}\tSection'.format('grade' if no_grade else 'is_here'))
    for studentID, [grade, section] in sorted(aliens.items(), reverse=True):
        print('{1}\t{0}{2}'.format(grade, studentID, '\t{}\tNot from class!'.format(section)))


def update_sheet(aliens, sheet_title):
    service = connect_GSheets()

    header_range = 'Section1!1:1'
    header_res = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id,
                                                     range=header_range).execute()

    col_idx = header_res['values'][0].index(sheet_title)

    requests = []
    requests.append({
        'updateCells': {
            'range': {
                'sheetId': section1_sheet_id,
                'startRowIndex': 1,
                'endRowIndex': len(section1_grade) + 1,
                'startColumnIndex': col_idx,
                'endColumnIndex': col_idx + 1
            },
            'fields': 'userEnteredValue.numberValue,userEnteredFormat.backgroundColor',
            'rows': [
                {
                    'values': {
                        'userEnteredValue': { 'numberValue': grade if is_true else 0 },
                        'userEnteredFormat': { 'backgroundColor': color_default if is_true else color_wrong }
                    }
                } for StudentID, [grade, is_true] in sorted(section1_grade.items(), reverse=True)
            ]
        }
    })
    requests.append({
        'updateCells': {
            'range': {
                'sheetId': section2_sheet_id,
                'startRowIndex': 1,
                'endRowIndex': len(section2_grade) + 1,
                'startColumnIndex': col_idx,
                'endColumnIndex': col_idx + 1
            },
            'fields': 'userEnteredValue.numberValue,userEnteredFormat.backgroundColor',
            'rows': [
                {
                    'values': {
                        'userEnteredValue': { 'numberValue': grade if is_true else 0 },
                        'userEnteredFormat': { 'backgroundColor': color_default if is_true else color_wrong }
                    }
                } for StudentID, [grade, is_true] in sorted(section2_grade.items(), reverse=True)
            ]
        }
    })

    attendance_result = service.spreadsheets().batchUpdate(
                                                spreadsheetId=spreadsheet_id,
                                                body={ 'requests': requests }).execute()


def calculate_student_grades_all(student_grades_s1, student_grades_s2, no_grade, sheet_title, verbose=0):
    aliens = {}
    for studentID, grade in student_grades_s1.items():
        # Correct section...
        if studentID in section1_grade:
            section1_grade[studentID] = [1 if no_grade else grade, studentID not in one2two]
        elif studentID in two2one:
            section2_grade[studentID] = [1 if no_grade else grade, True]
        # Wrong section...
        elif studentID in section2_grade:
            section2_grade[studentID] = [grade, False]
        # Not exists...
        else:
            aliens[studentID] = [1 if no_grade else grade, 1]

    for studentID, grade in student_grades_s2.items():
        # Correct section...
        if studentID in section2_grade:
            section2_grade[studentID] = [1 if no_grade else grade, studentID not in two2one]
        elif studentID in one2two:
            section1_grade[studentID] = [1 if no_grade else grade, True]
        # Wrong section...
        elif studentID in section1_grade:
            section1_grade[studentID] = [grade, False]
        # Not exists...
        else:
            aliens[studentID] = [1 if no_grade else grade, 2]

    if verbose > 0:
        dump_attendance(aliens, no_grade)
    else:
        update_sheet(aliens, sheet_title)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="xlsx to grades for Cmpe 140")
    parser.add_argument("student_answer_locations", nargs="+", help="File location xlsx files", type=str)
    parser.add_argument("-ng", "--no_grade", help="Should check only attandance",
                        action="store_true", default=True)
    parser.add_argument('-i', '--id_col', default='Student ID', help='text on column with id', type=str)
    parser.add_argument('-s', '--sheet_title', help="sheet title to update", type=str)
    parser.add_argument('-v', '--verbose', default=0, action='count')

    args = parser.parse_args()

    student_answers = []
    for student_answer_location in args.student_answer_locations:
        student_answers.append(create_student_grades(student_answer_location, args.id_col))

    calculate_student_grades_all(student_answers[0], student_answers[1],
                                 args.no_grade, args.sheet_title, args.verbose)
