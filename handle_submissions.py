#!/usr/local/bin/python3

import os
from os.path import join, dirname, abspath
from subprocess import Popen, PIPE
import argparse
import re

from sheet_helper import connect_GSheets, spreadsheet_id
from sheet_helper import color_header, color_id, color_default, color_white
from sheet_helper import get_border_request


def download_questions(all_student_answers, r_cridentials):
    """
    Connects to the server which requires connection to CMPE private network.
    Downloads all files uploaded for the questions and saves them into the
    parameter called all_student_answers.

    This supports only png and txt files.

    Args:
        all_student_answers: Contains file locations for every student
        r_cridentials: Is a tuple contains username and password for server
    """
    import requests

    credentials = tuple(r_cridentials.split('@'))

    session = requests.session()
    for student_answers in all_student_answers:
        for student_ID, answers in student_answers.items():
            for i in range(len(answers)):
                if answers[i] is not None:
                    response = session.get(answers[i], auth=credentials)
                    if answers[i].endswith('.r') or answers[i].endswith('.R'):
                        answers[i] = ('.R', response.text if response.ok else None)
                    # Denote txt files with .R extension
                    elif answers[i].endswith('.txt'):
                        answers[i] = ('.R', response.text if response.ok else None)
                    # Download and save file, requires to have png format...
                    elif answers[i].endswith('.png'):
                        content = b''
                        for chunk in response:
                            content += chunk
                        answers[i] = ('.png', None if content == b'' else content)
                    elif answers[i].endswith('.jpg'):
                        content = b''
                        for chunk in response:
                            content += chunk
                        answers[i] = ('.jpg', None if content == b'' else content)
                    elif answers[i].endswith('.jpeg'):
                        content = b''
                        for chunk in response:
                            content += chunk
                        answers[i] = ('.jpeg', None if content == b'' else content)


def create_student_answers(file_name, question_counts, target, id_col='Student ID'):
    """
    Parses the given excel to get student numbers and links of the files of
    submissions. Contains more code than required as remainder for xlrd package
    that has been used in this function.

    Args:
        :param file_name: Excel file name to parse
        :param question_counts: Question and file submission counts
        :param target: Name of the columns contains uploaded files
        :param id_col: Name of the column contains student ID
    """
    import xlrd

    # Open the workbook
    file_w_ext = file_name + '' if file_name.endswith('.xlsx') else '.xlsx'
    corrected_file_name = join(dirname(abspath(__file__)), file_w_ext)

    xl_workbook = xlrd.open_workbook(corrected_file_name)

    xl_sheet = xl_workbook.sheet_by_index(0)

    # Pull the first row by index
    #  (rows/columns are also zero-indexed)
    # Header row
    row = xl_sheet.row(0)  # 1st row

    # from xlrd.sheet import ctype_text

    student_ID_col_idx = 0
    student_answer_col_inx = []
    for idx, cell_obj in enumerate(row):
        # cell_type_str = ctype_text.get(cell_obj.ctype, 'unknown type')
        # location of the student ID coln
        if cell_obj.value == id_col:
            student_ID_col_idx = idx
        # locations of the question submission colns
        elif target in cell_obj.value:
            student_answer_col_inx.append(idx)

    student_answers = {}
    # num_cols = xl_sheet.ncols  # Number of columns
    for row_idx in range(1, xl_sheet.nrows):  # Iterate through rows
        try:
            # Get the student ID
            cell_obj = xl_sheet.cell(row_idx, student_ID_col_idx)
            student_ID = str(int(cell_obj.value))

            # Get all file urls
            answers = []
            for answer_count, ans_col in zip(question_counts, student_answer_col_inx):
                # Get list of names and url location of files
                answer = xl_sheet.cell(row_idx, ans_col).value

                # If there any file exists, there should be new line within
                # the answers
                if '\n' in answer:
                    # Split and extract urls without considering file names
                    # Answer structure:
                    # <File Name>
                    # <File Location in URL>    <- We Want These!
                    # <Empty Line>
                    # <File Name>
                    # <File Location in URL>    <- We Want These!
                    # <Empty Line>
                    _answers = answer.split('\n')[1::3]

                    # Put None to rest of the files if there are no submissions
                    answers += _answers + [None] * (answer_count - len(_answers))
                else:
                    answers += [None] * answer_count

            # Add submitted files under the student ID
            student_answers[student_ID] = answers
        except ValueError:
            print(f'{file_name} -> {row_idx}')
    return student_answers


def create_answers_class(folder_name, question_counts, student_answers):
    """

    Args:
        :param folder_name:
        :param question_counts:
        :param student_answers:
    """
    # Create question folders
    question_folders = []
    for i in range(len(question_counts)):
        question_folder = join(folder_name, f'q0{i + 1}')
        question_folders.append(question_folder)
        os.makedirs(question_folder, exist_ok=True)

        gitignore_file_path = join(question_folder, '.gitignore')
        if not os.path.exists(gitignore_file_path):
            with open(gitignore_file_path, 'w') as f:
                f.write('# Ignore everything in this directory\n')
                f.write('*\n')
                f.write('# Except this file\n')
                f.write('!.gitignore')

    # Write answers to files in designated folders
    for student_ID, answers in student_answers.items():
        loc = 0
        for idx, question_folder in enumerate(question_folders):
            for i in range(question_counts[idx]):
                # For getting attendance for null submissions
                if answers[loc] is None:
                    file_name = join(question_folder, f'{student_ID}.R')
                    if not os.path.isfile(file_name):
                        with open(file_name, 'w+') as f:
                            f.write('# AUTO GENERATED ANSWER FOR NULL SUBMISSION')
                else:
                    (extension, answer) = answers[loc]
                    file_name = join(question_folder, student_ID + extension)

                    counter = 1
                    while os.path.isfile(file_name):
                        file_name = join(question_folder, f'{student_ID}_{str(counter)}{extension}')
                        counter += 1

                    with open(file_name, 'wb+' if extension.lower() in ['.png', '.jpg', '.jpeg'] else 'w+') as f:
                        f.write(answer)
                loc += 1
                if loc == len(answers):
                    break


def create_answers_dataset(root_folder, question_counts, all_student_answers):
    """

    Args:
        :param root_folder:
        :param question_counts:
        :param all_student_answers:
    """
    # Check for the answers and sub-folder for saving student submission
    os.makedirs(root_folder, exist_ok=True)

    # When there are no sections, mostly for exams
    if len(all_student_answers) == 1:
        assert all_student_answers[0] is not None, "Answers do not exist!"
        create_answers_class(root_folder, question_counts[0],
                             all_student_answers[0])
    # Create for sectioned ones
    else:
        for section, (q_counts, s_answers) in enumerate(zip(question_counts,
                                                            all_student_answers)):
            assert s_answers is not None, f'Answers do not exist! Section {section}'

            # Create section folders
            section_folder = join(root_folder, f'sect0{section + 1}')
            os.makedirs(section_folder, exist_ok=True)

            create_answers_class(section_folder, q_counts, s_answers)


def question_counts(counts, counts_validity, section_count=2):
    """
    counts:          112,212 => [[1, 1, 2], [2, 1, 2]]
    counts_validity: TTF,TTF => [[True, True, False], [True, True, False]]

    counts:          112 => [[1, 1, 2], [1, 1, 2]]
    counts_validity: None => [[True, True, True], [True, True, True]]

    Args:
        :param counts: A string which shows question counts and required submissions
        :param counts_validity:
        :param section_count:
    """

    if ',' in counts:
        return [[int(c) for c in sec] for sec in counts.split(',')],\
               [[v == 'T' for v in sec] for sec in counts_validity.split(',')]
    else:
        validity = [True] * len(counts) if counts_validity is None\
            else [v == 'T' for v in counts_validity]
        return [[int(c) for c in counts]] * section_count, [validity] * section_count


def push_result2sheet(service, spreadsheet_id, _range, results):
    def cell_object(ir, i, row, cell):
        if ir == 0:  # header
            return {
                'userEnteredValue': {'stringValue': cell},
                'userEnteredFormat': {
                    'backgroundColor': color_header,
                    'horizontalAlignment': 'LEFT',
                    'textFormat': {
                        'bold': True
                    }
                }
            }
        elif i == 0:  # not header first column, id(number)
            return {
                'userEnteredValue': {'numberValue': cell},
                'userEnteredFormat': {
                    'backgroundColor': color_id,
                    'horizontalAlignment': 'LEFT',
                    'textFormat': {
                        'bold': True
                    }
                }
            }
        elif (i + 1) == len(row):  # last column, score(number)
            return {
                'userEnteredValue': {'numberValue': cell},
                'userEnteredFormat': {'backgroundColor': color_default}
            }
        else:  # rest of the cells
            return {
                'userEnteredValue': {'stringValue': cell},
                'userEnteredFormat': {'backgroundColor': color_white}
            }

    requests = [
        {
            'updateCells': {
                'range': _range,
                'fields': 'userEnteredValue.stringValue,' +
                          'userEnteredValue.numberValue,' +
                          'userEnteredFormat.backgroundColor,' +
                          'userEnteredFormat.horizontalAlignment,' +
                          'userEnteredFormat.textFormat.bold',
                'rows': [
                    {
                        'values': [
                            cell_object(ir, i, row, cell) for i, cell in enumerate(row)
                        ]
                    } for ir, row in enumerate(results)
                ]
            }
        }, get_border_request(_range)
    ]

    attendance_result = service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={'requests': requests}).execute()

    return attendance_result


def publish_results(name, results):
    from googleapiclient.errors import HttpError
    service = connect_GSheets()

    sheet_id = None
    try:
        # add a new sheet for new results
        body = {
            'requests': [{'addSheet': {'properties': {'title': name}}}]
        }

        add_sheet_response = service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body=body).execute()

        sheet_id = add_sheet_response['replies'][0]['addSheet']['properties']['sheetId']

        # add new columns and adjust the width of all columns
        body = {
            'requests': [
                {
                    'appendDimension': {
                        'sheetId': sheet_id,
                        'dimension': 'COLUMNS',
                        'length': 10
                    }
                },
                {
                    "updateDimensionProperties": {
                        "range": {
                            "sheetId": sheet_id,
                            "dimension": "COLUMNS",
                            "startIndex": 0,
                            "endIndex": 36
                        },
                        "properties": {"pixelSize": 50},
                        "fields": "pixelSize"
                    }
                }
            ]
        }
        service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
    except HttpError as http_err:
        spreadsheet_info = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()

        for sheet in spreadsheet_info['sheets']:
            if sheet['properties']['title'] == name:
                sheet_id = sheet['properties']['sheetId']
                break
        print(f'sheet exists with id: {sheet_id}')

    start_cell_loc = [0, 0]
    for sect_res in results:
        max_col = max([len(q_res[0]) for q_res in sect_res])
        for q_res in sect_res:
            # set the range of cells to fill in for each question
            _range = {
                'sheetId': sheet_id,
                'startRowIndex': start_cell_loc[0],
                'endRowIndex': start_cell_loc[0] + len(q_res),
                'startColumnIndex': start_cell_loc[1],
                'endColumnIndex': start_cell_loc[1] + max_col
            }
            start_cell_loc[0] += len(q_res)

            push_result2sheet(service, spreadsheet_id, _range, q_res)
        # return the first row for next section
        start_cell_loc[0] = 0
        # start from widest col number...
        start_cell_loc[1] += max_col


def run_checker_scripts(name, section_count, counts, counts_validity, checker_params_str, max_point):
    # from rpy2.robjects import r, pandas2ri
    # from rpy2 import robjects
    #
    # pandas2ri.activate()
    #
    # r.source('quiz_checker_env.R')
    # check_quizzes = robjects.globalenv['check_quizzes']
    # result = check_quizzes("Quiz3", "01", "01", answer_list="io/in_Quiz3_sect01_q01.RData", should_output_file=r.F)
    # result = pandas2ri.ri2py(result)

    answer_params = ['Rscript', f'reference_implementations/{name}.R', name]
    # Run to create answers
    with Popen(answer_params, stdout=PIPE) as proc:
        answer_locs = proc.stdout.read().decode().split('\n')

    checker_params = ['Rscript', 'quiz_checker_env.R']
    checker_params.extend(['name', name])
    checker_params.extend(['max_point', max_point])
    if checker_params_str != '':
        checker_params.extend(checker_params_str.split(' '))

    results = []
    counter = 0
    # Run the checker
    # for every section
    for s in range(section_count):
        sect_res = []
        cs_params = checker_params.copy()
        cs_params.extend(['section', 'NULL' if len(counts) == 1 else f'{(s+1):02d}'])
        # for every question
        for q in range(len(counts[s])):
            if not counts_validity[s][q]:
                continue
            csq_params = cs_params.copy()
            csq_params.extend(['question', f'{(q+1):02d}'])
            # answer files for each question should be sorted
            # by first the question and then the question numbers..
            # e.g. [sect1_q1, sect1_q2, sect2_q1, sect2_q2]
            csq_params.extend(['answer_list', answer_locs[counter]])
            with Popen(csq_params, stdout=PIPE) as proc:
                counter += 1
                console_out = proc.stdout.read().decode().split('\n')

                while 'student_ids final_result' not in console_out[0]:
                    console_out.pop(0)
                sect_res.append(console_out)

        results.append(sect_res)

    return results


def parse_Rscript_results(results):
    # header regex
    # p1 = re.compile(r'(\d*)\s*\"(.*?)\"')
    p0 = re.compile(r'(\d*).*')
    p1 = re.compile(r'(?:\"(.*?)\")|(NA)')
    # value regex
    p2 = re.compile(r'\s*(.*?)\s')

    p_results = []
    for sec_res in results:
        s_results = []
        for q_res in sec_res:
            q_result = [[]]
            for line in q_res:
                if line == '':
                    continue
                # when line of a result does not fit to console, output is
                # divided into sub-parts as new lines leading with the
                # actual *line number* which should be taken as a key to
                # merge multiple lines of actual console output
                line_num = 0

                num = p0.match(line)
                if len(num.group(1)) == 0:
                    q_result[line_num].extend([a for a in p2.findall(line + ' ')])
                    continue

                matches = p1.findall(line)
                line_num = int(num.group(1)) if num.group(1) != '' else line_num
                values = [col.strip() if na == '' else na for col, na in matches]

                if len(q_result) == line_num:
                    q_result.append(values)
                else:
                    q_result[line_num].extend(values)

            s_results.append(q_result)
        p_results.append(s_results)

    return p_results


def download_create_dataset(student_answers, credentials, name, counts, verbose):
    # Skip download part if folder exists...
    root_path = join(join(os.getcwd(), 'answers'), name)
    if os.path.exists(root_path):
        return

    # Verbose
    if verbose > 0:
        for student_answer in student_answers:
            print(student_answer)
        print(counts)

    # Download questions
    download_questions(student_answers, credentials)

    # Verbose
    if verbose > 1:
        for student_answer in student_answers:
            for _id, ans in student_answer.items():
                print('\n' + _id)
                for i in range(len(ans)):
                    if ans[i] is None:
                        continue
                    (_type, _val) = ans[i]
                    if _type not in ['.png', '.jpg', '.jpeg']:
                        if verbose > 2:
                            print(_val)
                        else:
                            print(_type)
                    else:
                        print(_type)

    # Create dataset for quiz/exam, (write to files)
    create_answers_dataset(root_path, counts, student_answers)


def main(args):
    if args.file_locs is not None:
        # Parse xlsx file
        student_answers = [create_student_answers(SA_location, q_counts,
                                                  args.target, args.id_col)
                           for q_counts, SA_location in zip(args.counts,
                                                            args.file_locs)]

        # Download files from server and save them with respect to project structure
        download_create_dataset(student_answers, args.credentials,
                                args.name, args.counts, args.verbose)

    for f in os.listdir('io') + os.listdir('results'):
        if args.name in f:
            os.remove(os.path.join('results' if f.endswith('csv') else 'io', f))

    # Run R scripts to generate answer key and grade submissions
    results = run_checker_scripts(args.name, len(args.counts), args.counts,
                                  args.counts_validity, args.checker_params, args.max_point)

    # Create a sheet in admin spreadsheet and post results into it
    publish_results(args.name, parse_Rscript_results(results))


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="xlsx to answers for Cmpe 140")

    parser.add_argument('-fl', "--file_locs", nargs="+", help="File location xlsx files", type=str)
    parser.add_argument('-cr', '--credentials', help='credentials to servers, @ separated name and password', type=str)
    parser.add_argument('-n', '--name', default='Quiz1', help='name of the quiz', type=str)
    parser.add_argument('-i', '--id_col', default='Student ID', help='text on column with id', type=str)
    parser.add_argument('-t', '--target', default='Code', help='target field name to download', type=str)
    parser.add_argument('-cv', '--counts_validity', default=None, help='which questions to control, for questions 111, '
                                                                       'TTF checks first and second but not third')
    parser.add_argument('-c', '--counts', default='1',
                        help='question counts Ex: 212 => 3 questions with 2, 1, 2 answers for both sections' +
                             'if 112,212 => 1, 1, 2 answers for section 1 and 2, 1, 2 answers for second')
    parser.add_argument('-cp', '--checker_params',
                        help='parameters of the quiz_checker separated with spaces <name> <value> pairs', default="",
                        type=str)
    parser.add_argument('-sc', '--sect_count', default=2)
    parser.add_argument('-mp', '--max_point', help="Max point of a question", default='20')
    parser.add_argument('-v', '--verbose', default=0, action='count')

    args = parser.parse_args()

    # parse question counts
    args.counts, args.counts_validity = question_counts(args.counts, args.counts_validity, args.sect_count)

    main(args)
