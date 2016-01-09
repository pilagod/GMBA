__author__ = 'pilagod'

from .Exam import *
from .ShareFunc import *

class Student:
    def __init__(self, serial_no, student_id, name, english_name, birth, gender, email, phone, department, level, study_year, rank, gpa,
                 exchange_term, toefl=TOEFL(), ielts=IELTS(), other="", remark="", wills=""):
        self.serial_no = serial_no
        self.student_id = student_id
        self.name = name
        self.english_name = english_name
        self.birth = birth
        self.gender = gender
        self.email = email
        self.phone = phone
        self.department = department
        self.level = level
        self.study_year = study_year
        self.rank = rank
        self.gpa = gpa
        self.exchange_term = exchange_term
        self.toefl = toefl
        self.ielts = ielts
        # self.toeic = toeic
        # self.jlpt = jlpt
        self.other = other
        self.remark = remark
        self.wills = wills

    def __repr__(self):
        return self.__class__.__name__ + "({0.serial_no}, {0.student_id}, " \
               "{0.name}, {0.english_name}, {0.birth}, " \
               "{0.gender}, {0.email}, {0.phone}, " \
               "{0.department}, {0.level}, {0.study_year}, " \
               "{0.rank}, {0.gpa}, {0.exchange_term}, {0.toefl}, {0.ielts}, " \
               "{0.remark}, {0.wills})".format(self)


def getStudentData(sheet):

    col_name = {}
    for col in range(sheet.ncols):
        col_name[caseAndSpaceIndif(sheet.cell(0, col).value)] = col

    students = {}
    for row in range(1, sheet.nrows):
        # For TOEFL => Total, Listening, Speaking, Reading, Writing
        toefl = TOEFL()
        toefl.setScores(str(sheet.cell(row, col_name[caseAndSpaceIndif('TOEFL(T,L,S,R,W)')]).value))

        # For IELTS => Total, Listening, Speaking, Reading, Writing
        ielts = IELTS()
        ielts.setScores(str(sheet.cell(row, col_name[caseAndSpaceIndif('IELTS(T,L,S,R,W)')]).value))

        # For TOEIC => Total, Listening, Reading
        # toeic = TOEIC()
        # toeic.setScores(str(sheet.cell(row, col_name[caseAndSpaceIndif('TOEIC(T,L,R)')]).value))

        # For JLPT
        # jlpt = 0
        # jlpt = sheet.cell(row, col_name['JLPT']).value if sheet.cell(row, col_name['JLPT']).value != "" else 0

        # For Wills
        # wills = sheet.cell(row, col_name['Wills']).value.split('/')
        wills = []
        i = 1
        while True:
            try:
                will = sheet.cell(row, col_name[caseAndSpaceIndif('Will') + str(i)]).value
                if not will:
                    break
                else:
                    wills.append(will)
                    i += 1
            except KeyError:
                break

        # Set Student Data
        students[sheet.cell(row, col_name[caseAndSpaceIndif('Serial No')]).value] = Student(
            sheet.cell(row, col_name[caseAndSpaceIndif('Serial No')]).value,
            sheet.cell(row, col_name[caseAndSpaceIndif('Student ID')]).value,
            sheet.cell(row, col_name[caseAndSpaceIndif('Name')]).value,
            sheet.cell(row, col_name[caseAndSpaceIndif('Name(English)')]).value,
            sheet.cell(row, col_name[caseAndSpaceIndif('Date of Birth')]).value,
            sheet.cell(row, col_name[caseAndSpaceIndif('Gender')]).value,
            sheet.cell(row, col_name[caseAndSpaceIndif('E-mail')]).value,
            sheet.cell(row, col_name[caseAndSpaceIndif('Cell Phone')]).value,
            sheet.cell(row, col_name[caseAndSpaceIndif('Department')]).value,
            sheet.cell(row, col_name[caseAndSpaceIndif('Level of Study')]).value.replace(" ", "").lower(),
            sheet.cell(row, col_name[caseAndSpaceIndif('Year of Study')]).value,
            sheet.cell(row, col_name[caseAndSpaceIndif('Rank')]).value,
            sheet.cell(row, col_name[caseAndSpaceIndif('GPA')]).value,
            sheet.cell(row, col_name[caseAndSpaceIndif('Exchange Term')]).value.replace(" ", "").lower().split(','),
            toefl, # sheet.cell(row, col_name['TOEFL(T/L/S/R/W)']).value,
            ielts, # sheet.cell(row, col_name['IELTS(T/L/S/R/W)']).value,
            # toeic, # sheet.cell(row, col_name['TOEIC(T/L/R)']).value,
            # jlpt, # sheet.cell(row, col_name['JLPT']).value,
            str(sheet.cell(row, col_name[caseAndSpaceIndif('Other')]).value),
            sheet.cell(row, col_name[caseAndSpaceIndif('Remark')]).value,
            wills # sheet.cell(row, col_name['Wills']).value
        )

    # print(students)

    return students