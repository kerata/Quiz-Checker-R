#!/usr/local/bin/python3

import smtplib
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import mimetypes
import os
import time

from argparse import ArgumentParser

from dashtable import data2rst

from other.student_info import gmail_user, gmail_app_password, gmail_password
from other.student_info import sect01_emails, sect02_emails, one2two, two2one

from sheet_helper import connect_GSheets, spreadsheet_id
from sheet_helper import section1_sheet_id, section2_sheet_id
from sheet_helper import color_default, color_wrong

EXCEPTIONS = {'ABSENCE': 0, 'SWITCH': 1}
PS_REGULAR = 'PS: Please forward this email to either your assistant or Professor if you have any objections.'
PS_EXAM = 'PS: Please respond to this email with your objections BEFORE the DEADLINE. You ' +\
          'have to be VERY CLEAR in your objections. Your objections should be ORDERED.'


def get_exceptions(path):
    if path is None:
        return {}

    with open(path, 'r') as f:
        ex = {}
        for line in f.readlines():
            if line == '':
                continue
            for stu in line.split(','):
                _key, _val = [v.strip() for v in stu.split(':')]
                if _key in ex:
                    ex[_key].append(EXCEPTIONS[_val])
                else:
                    ex[_key] = [EXCEPTIONS[_val]]
        return ex


def get_file_paths(folder_name, section_count):
    """
    Return either a dictionary or array of dictionaries depending
    on the section counts
    """
    def get_folder_contents(path, _files):
        _files = {} if _files is None else _files
        q_name = path.split('/')[-1]

        for file_path in os.listdir(path):
            if file_path.startswith('.'):
                continue

            _id, ext = file_path.split('.')

            full_path = os.path.join(path, file_path)
            candidate_name = f'{_id}_{q_name}.{ext}'

            if '_' in _id:
                _id = _id.split('_')[0]

            if _id in _files:
                _files[_id].append((full_path, candidate_name))
            else:
                _files[_id] = [(full_path, candidate_name)]

        return _files

    # section 1 means no section
    root_folder = f'answers/{folder_name}'
    files = None

    if section_count == 1:
        for q_folder in os.listdir(root_folder):
            if not os.path.isdir(f'{root_folder}/{q_folder}'):
                continue
            files = get_folder_contents(f'{root_folder}/{q_folder}', files)
        files = [files]
    else:
        files = []
        for s_folder in os.listdir(root_folder):
            if s_folder.startswith('.') or not os.path.isdir(f'{root_folder}/{s_folder}'):
                continue
            rel_folder = f'{root_folder}/{s_folder}'
            sub_files = None
            for q_folder in os.listdir(rel_folder):
                if q_folder.startswith('.') or not os.path.isdir(f'{rel_folder}/{q_folder}'):
                    continue
                sub_files = get_folder_contents(f'{rel_folder}/{q_folder}', sub_files)
            files.append(sub_files)

    return files


class QuestionHeader(object):
    PREFIX = ['INJ__', 'FNC__', 'wrapper_']

    def __init__(self, row):
        self.id = row[0]
        self.final_result = row[1]
        self.compilation_error = row[2]
        self.tests = []
        for i in range(3, len(row)-1, 2):
            self.tests.append(row[i])
        self.total_point = row[-1]

    def normalized_tests(self):
        n_tests = self.tests
        for p in self.PREFIX:
            n_tests = [t[len(p):] if t.startswith(p) else t for t in n_tests]
        return n_tests

    def __repr__(self):
        header = f'{self.id}'
        res = ' '.join([n for n in self.tests])
        return f'{header} {res} {self.total_point}'

    def __eq__(self, other):
        return all([self.id == other.id, self.final_result == other.final_result,
                    self.compilation_error == other.compilation_error, self.total_point == other.total_point,
                    *[t1 == t2 for t1, t2 in zip(self.tests, other.tests)]])


class StudentInfo(object):
    def __init__(self, student_id, question_grades=None, headers=None):
        self.id = student_id
        self.mail = None
        self.attended = True
        self.correct_section = False

        self.question_results = []
        self.total_point = 0

        self.question_grades = question_grades

        if headers is not None and len(headers) > 1:
            for i in range(len(headers)-1):
                self.add_null_question(headers[:(i+1)])

    def add_question(self, row):
        if self.question_grades:
            self.question_results.append(QuestionResult(row, self.question_grades[len(self.question_results)]))
        else:
            self.question_results.append(QuestionResult(row))
        self.total_point += self.question_results[-1].total_point

    def add_null_question(self, headers):
        if len(self.question_results) != len(headers):
            null_row = ['FALSE', 'NULL SUBMISSION'] + ['NA', ''] * len(headers[-1].tests) + ['0']
            self.question_results.append(QuestionResult(null_row))

    def table_texts(self, name, question_headers, total_score=None, question_grades=None):
        if question_grades and total_score is None:
            total_score = sum([e for r in question_grades for e in (r[0] if r else [20])])
        score_perc_question = total_score / len(question_headers)
        tables = []
        overall = [
            [f'{name}', '', 'Info/Result', '']
        ]
        overall_spans = [
            [[0, 0], [0, 1]], [[0, 2], [0, 3]]
        ]
        for i, (q_result, q_header) in enumerate(zip(self.question_results, question_headers)):
            recalculated_points = question_grades[i][0] if question_grades and question_grades[i] else None

            has_sourced = q_result.compilation_error == ''
            table = [
                [f'Q{i+1}', '', 'Info/Result', '']
            ]
            spans = [
                [[0, 0], [0, 1]], [[0, 2], [0, 3]]
            ]
            if has_sourced:
                for q_i, (test, (p, e)) in enumerate(zip(q_header.normalized_tests(), q_result.points)):
                    score_perc = recalculated_points[q_i] if recalculated_points else \
                        score_perc_question / len(q_result.points)
                    spans.append([[q_i+1, 0], [q_i+1, 1]])
                    if e == '':
                        # True or False, without source error
                        table.append([
                            f'{test}({score_perc:.2f})', '', p, ''
                        ])
                        spans.append([[q_i+1, 2], [q_i+1, 3]])
                    else:
                        table.append([
                            f'{test}({score_perc:.2f})', '', 'Error', e
                        ])
            else:
                table.append([
                    'Source Error', '', q_result.compilation_error, ''
                ])
                spans.append([[1, 0], [1, 1]])
                spans.append([[1, 2], [1, 3]])
            table.append([
                f'Total({sum(recalculated_points) if recalculated_points else score_perc_question:.2f})', '',
                q_result.total_point, ''
            ])
            spans.append([[len(table) - 1, 0], [len(table) - 1, 1]])
            spans.append([[len(table) - 1, 2], [len(table) - 1, 3]])

            overall.append([
                f'Q{i+1}({sum(recalculated_points) if recalculated_points else score_perc_question:.2f})',
                '', q_result.total_point, ''
            ])
            overall_spans.append([[len(overall) - 1, 0], [len(overall) - 1, 1]])
            overall_spans.append([[len(overall) - 1, 2], [len(overall) - 1, 3]])

            tables.append(data2rst(table, spans=spans, center_headers=True,
                                   center_cells=True, use_headers=True))

        if len(question_headers) > 1:
            overall.append([
                f'Total({total_score:.2f})', '', sum([q_r.total_point for q_r in self.question_results]), ''
            ])
            overall_spans.append([[len(overall) - 1, 0], [len(overall) - 1, 1]])
            overall_spans.append([[len(overall) - 1, 2], [len(overall) - 1, 3]])

            tables.append(data2rst(overall, spans=overall_spans, center_headers=True,
                                   center_cells=True, use_headers=True))
        return tables

    def __repr__(self):
        res = '\n'.join([qr.__repr__() for qr in self.question_results])
        return f'{self.id}\n{res}'

    def __eq__(self, other):
        return all([self.id == other.id, self.mail == other.mail, self.attended == other.attended,
                    self.correct_section == other.correct_section, self.total_point == other.total_point,
                    self.question_grades == other.question_grades, self.question_results == other.question_results])


class QuestionResult(object):
    def __init__(self, row, recalculate=None):
        # Question related information
        self.is_compiled = row[0] == 'TRUE'
        self.compilation_error = row[1]
        self.points = []
        for i, loc in enumerate(range(2, len(row)-1, 2)):
            if not recalculate:
                self.points.append((row[loc], row[loc+1]))
            else:
                if row[loc] in ['TRUE', 'FALSE', 'NA', '']:
                    self.points.append(('FALSE' if row[loc] == '' else row[loc], row[loc+1]))
                else:
                    given_grade = float(row[loc])
                    if given_grade == 0:
                        self.points.append(('FALSE', row[loc + 1]))
                    elif given_grade == recalculate[1][i]:
                        self.points.append(('TRUE', row[loc + 1]))
                    else:
                        self.points.append(((given_grade/recalculate[1][i])*recalculate[0][i], row[loc + 1]))

        self.total_point = float(row[-1]) if not recalculate else \
            sum([0 if g in ['FALSE', 'NA', ''] else recalculate[0][i] if g == 'TRUE' else g
                 for i, (g, _) in enumerate(self.points)])

    def table_repr(self, lengths):
        res = ' '.join([r.rjust(left) for left, (r, _) in zip(lengths, self.points)])
        return f'{res} {self.total_point}'

    def __repr__(self):
        res = ' '.join([f'{r: <6}' for (r, _) in self.points])
        return f'{res} {self.total_point}'

    def __eq__(self, other):
        return all([self.is_compiled == other.is_compiled, self.compilation_error == other.compilation_error,
                    self.total_point == other.total_point,
                    *[p1 == p2 for p1, p2 in zip(self.points, other.points)]])


def get_scores(sheet_title, score_title, exceptions, question_grades=None):
    service = connect_GSheets()

    # Quarries that could be helpful for getting or updating values in the spreadsheet
    # result = service.spreadsheets()\
    #     .get(spreadsheetId=spreadsheet_id, ranges=f'{sheet_title}!1:1', includeGridData=True)\
    #     .execute()
    # first_row_cells = result['sheets'][0]['data'][0]['rowData']
    # for cell in first_row_cells:
    #     cell_note = cell['values'][0]['note']
    #     cell_value = cell['values'][0]['effectiveValue']
    #     cell_actual = cell['values'][0]['userEnteredValue']

    # get resulting sheet
    result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=sheet_title).execute()

    scores = []
    question_headers = []
    # For each section do
    for i, s_loc in enumerate([i for i, x in enumerate(result['values'][0]) if x == 'student_ids']):
        # Add for each sections
        question_headers.append([])
        section_results = {}
        q_end = s_loc
        for row in result['values']:
            if not len(row) > s_loc:
                # End of the section
                break

            if row[s_loc] == 'student_ids':
                for s in section_results.values():
                    s.add_null_question(question_headers[i])

                q_end = s_loc
                # find the range for current question
                while row[q_end] != score_title:
                    q_end += 1
                # Add question header to related section
                question_headers[-1].append(QuestionHeader(row[s_loc:q_end + 1]))
                continue

            # Pass if empty rows comes in when section student counts are different
            if all([e == '' for e in row[s_loc:q_end + 1]]):
                continue

            if row[s_loc] not in section_results:
                # Create student
                student = StudentInfo(row[s_loc], question_grades, question_headers[i])

                # Add student to the section
                section_results[student.id] = student

            # Add question to the student information in section
            section_results[row[s_loc]].add_question(row[s_loc + 1:q_end + 1])

        for s in section_results.values():
            s.add_null_question(question_headers[i])
        # Add section to the scores array
        scores.append(section_results)

    if len(scores) == 1:
        rooster = sect01_emails.copy()
        rooster.update(sect02_emails)
        for s_id in list(scores[0].keys()):
            if s_id in rooster:
                scores[0][s_id].correct_section = True
                scores[0][s_id].mail = rooster[s_id]
            else:
                scores[0].pop(s_id)
                print(f'{s_id} not exists!')
        # Look for the absent students
        remaining_students = set(rooster.keys()).difference(set(scores[0].keys()))
        # add them to the section as unattended
        for s_id in remaining_students:
            student = StudentInfo(s_id, question_grades)
            student.attended = False
            student.mail = rooster[s_id]
            scores[0][s_id] = student
        return question_headers, scores

    for i in range(len(scores)):
        s_scores, other_scores = scores if i == 0 else reversed(scores)
        rooster, other_rooster = [sect01_emails, sect02_emails] if i == 0 else [sect02_emails, sect01_emails]
        sent, received = [one2two, two2one] if i == 0 else [two2one, one2two]
        for s_id in list(s_scores.keys()):
            if s_id in rooster:
                # get and set the mail from the section rooster if correct section
                if not s_scores[s_id].mail:
                    s_scores[s_id].mail = rooster[s_id]
                    s_scores[s_id].correct_section = s_id not in sent or \
                                                     (s_id in exceptions and
                                                      any([e == EXCEPTIONS['SWITCH'] for e in exceptions[s_id]]))
            elif s_id in other_rooster:
                # get and set the mail from the other section rooster if incorrect section
                s_scores[s_id].mail = other_rooster[s_id]
                # control if it has an exception, set correct section if they are excused
                s_scores[s_id].correct_section = s_id in received or \
                                                 (s_id in exceptions and
                                                  any([e == EXCEPTIONS['SWITCH'] for e in exceptions[s_id]]))

                other_scores[s_id] = s_scores.pop(s_id)
            else:
                s_scores.pop(s_id)
                # Log if student does not exists
                print(f'{s_id} not exists!')

        # Look for the absent students
        remaining_students = set(rooster.keys()).difference(set(s_scores.keys()))
        # add them to the section as unattended
        for s_id in remaining_students:
            student = StudentInfo(s_id, question_grades)
            student.attended = False
            student.mail = rooster[s_id]
            s_scores[s_id] = student

    return question_headers, scores


def attach_file(msg, path, c_name):
    # Guess the content type based on the file's extension.  Encoding
    # will be ignored, although we should check for simple things like
    # gzip'd or compressed files.
    ctype, encoding = mimetypes.guess_type(path)
    if ctype is None or encoding is not None:
        # No guess could be made, or the file is encoded (compressed), so
        # use a generic bag-of-bits type.
        ctype = 'application/octet-stream'

    maintype, subtype = ctype.split('/', 1)
    with open(path, 'rb') as fp:
        bin_content = fp.read()
        if ctype in ['image/png', 'image/jpeg', 'image/jpg']:
            part = MIMEImage(bin_content, subtype)
        else:
            try:
                if bin_content.decode() == '# AUTO GENERATED ANSWER FOR NULL SUBMISSION':
                    return
                else:
                    part = MIMEBase(maintype, subtype)
                    part.set_payload(bin_content)
            except Exception as e:
                part = MIMEBase(maintype, subtype)
                part.set_payload(bin_content)

        part.add_header('Content-Disposition',
                        'attachment; filename="{}"'.format(c_name if c_name is not None else path.split('/')[-1]))
        msg.attach(part)


def build_mail_body(_id, student, exam_name, files, question_headers, total_score,
                    question_grades=None, ignore_correct_section=True):
    msg = MIMEMultipart('mixed')

    if student.attended:
        if ignore_correct_section or student.correct_section:
            score_tables = student.table_texts(exam_name, question_headers, total_score, question_grades)

            content = ['\n'.join(['Hello,', '',
                                  f'Please find attached your {exam_name} submission(s) and your score.', '',
                                  f'Student ID: {_id}', '']),
                       '\n'.join(['', 'Best regards,', 'CmpE140 team.', '',
                                  PS_EXAM if exam_name in ['Final', 'Midterm'] or 'assignment' in exam_name.lower() else PS_REGULAR])]
            msg.attach(MIMEText(content[0], 'plain'))
            msg.attach(MIMEText('\n'.join(['<html>', '<head></head>', '<body>', '<pre style="font-family: monospace">',
                                           *[f'{table}' for table in score_tables],
                                           '</pre>', '</body>', '</html>']), 'html'))
            msg.attach(MIMEText(content[1], 'plain'))

            content = ''.join([content[0], '\n'.join([f'{table}' for table in score_tables]), content[1]])
        else:
            content = '\n'.join(['Hello,', '',
                                 'You received 0 points because you have attended to wrong section.', '',
                                 f'Student ID: {_id}', f'Total score: 0/{total_score}', '',
                                 'Best regards,', 'CmpE140 team.', '',
                                 PS_EXAM if exam_name in ['Final', 'Midterm'] or 'assignment' in exam_name.lower() else PS_REGULAR])
            msg.attach(MIMEText(content, 'plain'))
    else:
        if question_grades:
            total_score = sum([e for r in question_grades for e in (r[0] if r else [20])])
        content = '\n'.join(['Hello,', '',
                             f'You received 0 points because you did not attended to {exam_name}.', '',
                             f'Student ID: {_id}',
                             f'Total score: 0/{total_score}', '',
                             'Best regards,', 'CpeE140 team.', '',
                             PS_EXAM if exam_name in ['Final', 'Midterm'] or 'assignment' in exam_name.lower() else PS_REGULAR])
        msg.attach(MIMEText(content, 'plain'))

    msg['Subject'] = f'[NOREPLY] CmpE 140 {exam_name} Result'
    msg['From'] = gmail_user
    msg['To'] = student.mail

    # attach files if exists
    if files:
        for (file_path, file_name) in files:
            attach_file(msg, file_path, file_name)

    # for testing purposes
    dump = '\n'.join([student.mail, content, ', '.join([f'{fp}: {fn}' for fp, fn in files]) if files else '', '\n'])
    dump += '=' * 30 + '\n' * 3

    return msg, dump


def send_mails(exam_name, scores, question_headers,
               files, server, delay, total_score,
               question_grades=None):
    for i, (score_set, header_set, file_set) in enumerate(zip(scores, question_headers, files)):
        should_override = i == 0
        print(f'Mailing started: {i}. Section')

        dump = ''
        for _id, student in score_set.items():
            # send mail to student
            if not student.mail:
                continue

            switch_set = len(scores) > 1 and ((_id in one2two and i == 0) or (_id in two2one and i == 1))

            file = (files[1-i].get(_id, None) if files[1-i] else None) if switch_set else \
                (file_set.get(_id, None) if file_set else None)

            mail_body, msg_dump = build_mail_body(_id, student, exam_name, file,
                                                  question_headers[1-i] if switch_set else header_set,
                                                  total_score, question_grades)

            dump += msg_dump

            try:
                server.send_message(mail_body)
                print(_id)
            except Exception as e:
                print(f'Mail NOT sent to: {_id} :: {e}')

            # sleep between mails
            time.sleep(delay)
        with open(f'{exam_name}.mail', 'w' if should_override else 'a') as f:
            f.write(dump)


def update_score_list(name, attdn_name, m_scores, s_update_scores=True, ignore_wrong_sec=True):
    service = connect_GSheets()

    header_range = 'Section1!1:1'
    header_res = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id,
                                                     range=header_range).execute()

    if len(m_scores) == 1:
        scores = [{_id: m_scores[0][_id] for _id in s} for s in [sect01_emails, sect02_emails]]
    else:
        scores = m_scores

    requests = []
    if s_update_scores:
        col_idx = header_res['values'][0].index(name)
        requests.extend([
          {
            'updateCells': {
              'range': {
                'sheetId': section_id,
                'startRowIndex': 1,
                'endRowIndex': len(s_scores) + 1,
                'startColumnIndex': col_idx,
                'endColumnIndex': col_idx + 1
              },
              'fields': 'userEnteredValue.numberValue,userEnteredFormat.backgroundColor',
              'rows': [
                {
                  'values': {
                    'userEnteredValue': {
                        'numberValue': s_scores[k].total_point if ignore_wrong_sec or s_scores[k].correct_section else 0
                    },
                    'userEnteredFormat': {
                      'backgroundColor': color_wrong if s_scores[k].attended and not s_scores[k].correct_section else color_default
                    }
                  }
                } for k in sorted(list(s_scores.keys()), reverse=True)
              ]
            }
          } for section_id, s_scores in zip([section1_sheet_id, section2_sheet_id], scores)
        ])

    if attdn_name:
        attdn_idx = header_res['values'][0].index(attdn_name)
        requests.extend([
          {
            'updateCells': {
              'range': {
                'sheetId': section_id,
                'startRowIndex': 1,
                'endRowIndex': len(s_scores) + 1,
                'startColumnIndex': attdn_idx,
                'endColumnIndex': attdn_idx + 1
              },
              'fields': 'userEnteredValue.numberValue,userEnteredFormat.backgroundColor',
              'rows': [
                {
                  'values': {
                    'userEnteredValue': {'numberValue': 1 if ignore_wrong_sec or s_scores[k].correct_section else 0},
                    'userEnteredFormat': {
                      'backgroundColor': color_wrong if s_scores[k].attended and not s_scores[k].correct_section else color_default
                    }
                  }
                } for k in sorted(list(s_scores.keys()), reverse=True)
              ]
            }
          } for section_id, s_scores in zip([section1_sheet_id, section2_sheet_id], scores)
        ])

    # Request sect1 grades, sect2 grades, sect1 attendance, and
    # sect2 attendance with respect the given order
    service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body={'requests': requests}).execute()


def get_question_grades(recalculate, recalculate_numbers):
    """
    Example for single question:
    let there be a question with 4 sub-questions that have graded by 5, 5, 5, and 10
    but actually total grade for the question should be 20 and last question should be
    scaled to make max point 5.
    For this example recalculate should be 5,5,5,5:5,5,5,10
    where the first group 5,5,5,5 shows points for the sub-questions
    whereas 5,5,5,10 shows current grading schemes for sub-questions
    We should scale sub-question-4 so that
    student who took 10 from it has to take 5 and the one who took 6 should took 3 instead.
    We should update the total point accordingly as well in the end.
    """

    # Multiple questions should be separated by ";"
    question_grades = [[[float(g) for g in gr.split(',')] for gr in qg.split(':')]
                       for qg in recalculate.split(';')] if recalculate else None
    if recalculate_numbers is not None:
        n, w = recalculate_numbers.split(':')
        w = [int(v) for v in w.split(';')]
        it = iter(question_grades)
        question_grades = [next(it) if i in w else None for i in range(int(n))]

    return question_grades


def main(args):
    # parse recalculate arguments
    question_grades = get_question_grades(args.recalculate, args.recalculate_numbers)

    # get scores from the sheets
    question_headers, scores = get_scores(args.sheet_title, args.score_title,
                                          get_exceptions(args.exceptions), question_grades)

    # Update admin spreadsheet to fill final scores and attendance for given Quiz/Exam
    update_score_list(args.name, args.attdn_name, scores, args.update_scores)

    if args.only_update:
        return

    try:
        # Log in to gmail client
        server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
        server.ehlo()

        # server.login(gmail_user, gmail_password)
        server.login(gmail_user, gmail_app_password)

        # locate files
        files = [None] * len(scores) if args.hide_files else get_file_paths(args.folder_name, len(scores))

        send_mails(args.name, scores, question_headers, files, server, args.delay, args.total_score, question_grades)

        # log out from client
        server.quit()
    except Exception as e:
        print(e)


if __name__ == '__main__':
    parser = ArgumentParser("Script for sending results by emails")

    parser.add_argument('-n', '--name', type=str, default="Quiz1",
                        help="Name of the exam/quiz")
    parser.add_argument('-ou', '--only_update', action='store_true')
    parser.add_argument('-us', '--update_scores', action='store_false',
                        help="Use if you want to forbid to update scores")
    parser.add_argument('-an', '--attdn_name', type=str,
                        help="Name of Lab for attendances to be filled")
    parser.add_argument('-fn', '--folder_name', type=str,
                        help="Name of containing folder")
    parser.add_argument('-s', '--sheet_title', type=str,
                        help="Title of the sheet")
    parser.add_argument('-st', '--score_title', type=str, default="points",
                        help="Title of the scores")
    parser.add_argument('-tc', '--total_score', type=int, default=20,
                        help="Total value of the score")
    parser.add_argument('-r', '--recalculate',
                        help="Required for recalculation of grades Ex: 1,1,1,1:5,5,5,5;2,2,2,2:5,5,5,5;2,2,2,2:5,5,5,5")
    parser.add_argument('-rn', '--recalculate_numbers',
                        help="Required for recalculation of grades, represents the question numbers Ex: 3:0;1;2")
    parser.add_argument('-hf', '--hide_files', action='store_true', default=False,
                        help="Should withheld files when sending the emails")
    parser.add_argument('-d', '--delay', type=float, default=1,
                        help="Delay between mails in sec")
    parser.add_argument('-e', '--exceptions',
                        help='File location of the students who have permission')

    p_args = parser.parse_args()

    p_args.folder_name = p_args.folder_name if p_args.folder_name else p_args.name
    p_args.sheet_title = p_args.sheet_title if p_args.sheet_title else p_args.name
    main(p_args)
