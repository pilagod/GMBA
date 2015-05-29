__author__ = 'pilagod'

from .Exam import *

class Student:
    def __init__(self, serial_no, student_id, name, english_name, birth, gender, email, phone, department, level, study_year, rank, gpa,
                 toefl=TOEFL(), ielts=IELTS(), toeic=TOEIC(), jlpt=0, remark="", wills=""):
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
        self.toefl = toefl
        self.ielts = ielts
        self.toeic = toeic
        self.jlpt = jlpt
        self.remark = remark
        self.wills = wills

    def __repr__(self):
        return self.__class__.__name__ + "({0.serial_no}, {0.student_id}, " \
               "{0.name}, {0.english_name}, {0.birth}, " \
               "{0.gender}, {0.email}, {0.phone}, " \
               "{0.department}, {0.level}, {0.study_year}, " \
               "{0.rank}, {0.gpa}, {0.toefl}, {0.ielts}, " \
               "{0.toeic}, {0.jlpt}, {0.remark}, {0.wills})".format(self)


def getStudentData(sheet):

    col_name = {}
    for col in range(sheet.ncols):
        col_name[sheet.cell(0, col).value] = col

    students = {}
    for row in range(1, sheet.nrows):
        # For TOEFL => Total, Listening, Speaking, Reading, Writing
        toefl = TOEFL()
        toefl.setScores(sheet.cell(row, col_name['TOEFL(T/L/S/R/W)']).value)

        # For IELTS => Total, Listening, Speaking, Reading, Writing
        ielts = IELTS()
        ielts.setScores(sheet.cell(row, col_name['IELTS(T/L/S/R/W)']).value)

        # For TOEIC => Total, Listening, Reading
        toeic = TOEIC()
        toeic.setScores(sheet.cell(row, col_name['TOEIC(T/L/R)']).value)

        # For JLPT
        jlpt = 0
        jlpt = sheet.cell(row, col_name['JLPT']).value if sheet.cell(row, col_name['JLPT']).value != "" else 0

        # For Wills
        wills = sheet.cell(row, col_name['Wills']).value.split('/')

        # Set Student Data
        students[sheet.cell(row, col_name['Serial No']).value] = Student(
            sheet.cell(row, col_name['Serial No']).value,
            sheet.cell(row, col_name['Student ID']).value,
            sheet.cell(row, col_name['Name']).value,
            sheet.cell(row, col_name['Name(English)']).value,
            sheet.cell(row, col_name['Date of Birth']).value,
            sheet.cell(row, col_name['Gender']).value,
            sheet.cell(row, col_name['E-mail']).value,
            sheet.cell(row, col_name['Cell Phone']).value,
            sheet.cell(row, col_name['Department']).value,
            sheet.cell(row, col_name['Level of Study']).value.replace(" ", ""),
            sheet.cell(row, col_name['Year of Study']).value,
            sheet.cell(row, col_name['Rank']).value,
            sheet.cell(row, col_name['GPA']).value,
            toefl, # sheet.cell(row, col_name['TOEFL(T/L/S/R/W)']).value,
            ielts, # sheet.cell(row, col_name['IELTS(T/L/S/R/W)']).value,
            toeic, # sheet.cell(row, col_name['TOEIC(T/L/R)']).value,
            jlpt, # sheet.cell(row, col_name['JLPT']).value,
            sheet.cell(row, col_name['Remark']).value,
            wills # sheet.cell(row, col_name['Wills']).value
        )

    return students