__author__ = 'pilagod'

from .Exam import *
from .ShareFunc import *

class SchoolRequirement:
    def __init__(self, school_code, country, school_name, school_name_chinese, level, exchange_term, slots,
                 gpa=[0, 0], toefl=[TOEFL(), TOEFL()], ielts=[IELTS(), IELTS()], toeic=[TOEIC(), TOEIC()], jlpt=[-1, -1], working_experience=0, english_taught="No", others=""):
        self.school_code = school_code
        self.country = country
        self.school_name = school_name
        self.school_name_chinese = school_name_chinese
        self.level = level
        self.exchange_term = exchange_term
        self.slots = slots
        self.gpa = gpa
        self.toefl = toefl
        self.ielts = ielts
        self.toeic = toeic
        self.jlpt = jlpt
        self.working_experience = working_experience
        self.english_taught = english_taught
        self.others = others

    def __repr__(self):
        return self.__class__.__name__ + "({0.school_code}, {0.country}, " \
                "{0.school_name}, {0.school_name_chinese}, " \
                "{0.level}, {0.gpa}, {0.toefl}, {0.ielts}, " \
                "{0.toeic}, {0.jlpt}, {0.working_experience}, " \
                "{0.english_taught}, {0.others}, {0.exchange_term}, " \
                "{0.slots})".format(self)



def getDataRelatedToLevel(result, data):
    data_string = data.replace(" ", "").split('/')
    if isinstance(result[0], EnglishExam):
        if len(data_string) > 1 :
            result[0].setScores(data_string[0][data_string[0].find(':')+1:])
            result[1].setScores(data_string[1][data_string[1].find(':')+1:])
        else:
            if data_string[0] == "":
                result[0] = None
            else:
                result[0].setScores(data_string[0])

            result[1] = None
    else:
        if len(data_string) > 1 :
            result[0] = float(data_string[0][data_string[0].find(':')+1:])
            result[1] = float(data_string[1][data_string[1].find(':')+1:])
        else:
            if(data_string[0] == ""):
                result[0] = -1
            else:
                result[0] = float(data_string[0])

            result[1] = -1

    return result

def getSchoolData(sheet):
    col_name = {}
    for col in range(sheet.ncols):
        col_name[caseAndSpaceIndif(sheet.cell(0, col).value)] = col
        # print(caseAndSpaceIndif(sheet.cell(0, col).value))

    school_requirements = {}
    for row in range(1, sheet.nrows):
        # print(sheet.cell(row, col_name["School Code"]).value)
        level = sheet.cell(row, col_name[caseAndSpaceIndif("Requirement: Under / Grad")]).value.replace(" ", "").split('/')

        # For TOEFL
        toefl = [TOEFL(), TOEFL()]
        toefl = getDataRelatedToLevel(toefl, sheet.cell(row, col_name[caseAndSpaceIndif("Requirement: TOEFL(T,L,S,R,W)")]).value)

        # For IELTS
        ielts = [IELTS(), IELTS()]
        ielts = getDataRelatedToLevel(ielts, sheet.cell(row, col_name[caseAndSpaceIndif("Requirement: IELTS(T,L,S,R,W)")]).value)

        # For TOEIC
        toeic = [TOEIC(), TOEIC()]
        toeic = getDataRelatedToLevel(toeic, sheet.cell(row, col_name[caseAndSpaceIndif("Requirement: TOEIC(T,L,R)")]).value)

        # For JLPT
        jlpt = [-1, -1]
        jlpt = getDataRelatedToLevel(jlpt, str(sheet.cell(row, col_name[caseAndSpaceIndif("Requirement: JLPT")]).value))

        # For Exchange Term
        exchange = sheet.cell(row, col_name[caseAndSpaceIndif("Exchange term")]).value.replace(" ", "").lower().split('/')
        exchange[0] = exchange[0][exchange[0].find(':')+1:].split(',')
        if len(exchange) > 1:
            exchange[1] = exchange[1][exchange[1].find(':')+1:].split(',')

        # print(exchange)

        # For Slots
        slots = [-1, -1]
        getDataRelatedToLevel(slots, str(sheet.cell(row, col_name[caseAndSpaceIndif("Slots")]).value))

        school_requirements[caseAndSpaceIndif(sheet.cell(row, col_name[caseAndSpaceIndif("School Code")]).value)] = SchoolRequirement(
            sheet.cell(row, col_name[caseAndSpaceIndif("School Code")]).value,
            sheet.cell(row, col_name[caseAndSpaceIndif("Country")]).value,
            sheet.cell(row, col_name[caseAndSpaceIndif("School Name")]).value,
            sheet.cell(row, col_name[caseAndSpaceIndif("School Name(Chinese)")]).value,
            level, # sheet.cell(row, col_name["Requirement: Under / Grad"]).value,
            exchange, # sheet.cell(row, col_name[caseAndSpaceIndif("Exchange term")]).value,
            slots, # sheet.cell(row, col_name["Slots"]).value,
            sheet.cell(row, col_name[caseAndSpaceIndif("Requirement: GPA")]).value if sheet.cell(row, col_name[caseAndSpaceIndif("Requirement: GPA")]).value != "" else None,
            toefl, # sheet.cell(row, col_name["Requirement: TOEFL"]).value,
            ielts, # sheet.cell(row, col_name["Requirement: IELTS"]).value,
            toeic, # sheet.cell(row, col_name["Requirement: TOEIC"]).value,
            jlpt, # sheet.cell(row, col_name["Requirement: JLPT"]).value,
            sheet.cell(row, col_name[caseAndSpaceIndif("Requirement: Working Experience(years)")]).value,
            sheet.cell(row, col_name[caseAndSpaceIndif("Requirement: English-taught program offered")]).value,
            sheet.cell(row, col_name[caseAndSpaceIndif("Requirement: Others")]).value
        )

    return school_requirements


def requirementLevelTest(student_level, requirement_levels):
    if requirement_levels[0] == '':
        return True

    # if student_level in requirement_levels:
    #     return True
    for requirement_level in requirement_levels:
        if requirement_level == 'U' and student_level == "Under":
            return True
        elif requirement_level == 'G' and student_level in ["Master", "Graduate"]:
            print("requirementLevelTest", student_level)
            return True
        elif requirement_level == 'M' and student_level == "Master":
            return True

    print("requirementLevelTest: (False, {0}, {1})".format(student_level, requirement_levels))
    return False

def requirementGpaTest(student_gpa, requirement_gpa):
    if requirement_gpa == None:
        return True
    elif student_gpa >= requirement_gpa:
        return True

    print("requireGpaTest: (False, {0}, {1})".format(student_gpa, requirement_gpa))
    return False

def requirementScoreTest(student_level, student_score, requirement_scores):
    if requirement_scores[0] == None:
        return False

    # if student_level in Level.Under.value:
    #     if student_score >= requirement_scores[0]:
    #         print("Pass {0} {1}".format(student_score, requirement_scores))
    #         return True
    # elif student_level in Level.Graduate.value:
    #     if requirement_scores[1] == None:
    #         if student_score >= requirement_scores[0]:
    #             print("Pass {0} {1}".format(student_score, requirement_scores))
    #             return True
    #     else:
    #         if student_score >= requirement_scores[1]:
    #             print("Pass {0} {1}".format(student_score, requirement_scores))
    #             return True

    if student_level == "Under":
        # print("({0}, {1}, {2})".format(student_score >= requirement_scores[0], student_score, requirement_scores[0]))
        if student_score >= requirement_scores[0]:
            print("Pass {0} {1}".format(student_score, requirement_scores))
            return True
    elif student_level in ["Master", "Graduate"]:
        if requirement_scores[1] == None:
            if student_score >= requirement_scores[0]:
                print("Pass {0} {1}".format(student_score, requirement_scores))
                return True
        else:
            if student_score >= requirement_scores[1]:
                print("Pass {0} {1}".format(student_score, requirement_scores))
                return True

    print("requirementScoreTest({0}): (False, {1}, {2}, {3})".format(student_score.__class__.__name__, student_level, student_score, requirement_scores))
    return False

# def requirementJLPTScoreTest(student_level, student_score, requirement_scores):
#     if requirement_scores[0] < 0:
#         return False
#
#     if student_level in "Undergraduates":
#         # print("({0}, {1}, {2})".format(student_score >= requirement_scores[0], student_score, requirement_scores[0]))
#         if student_score >= requirement_scores[0]:
#             print("Pass {0} {1}".format(student_score, requirement_scores))
#             return True
#     elif student_level in "Masters" or "Graduates":
#         if requirement_scores[1] < 0:
#             if student_score >= requirement_scores[0]:
#                 print("Pass {0} {1}".format(student_score, requirement_scores))
#                 return True
#         else:
#             if student_score >= requirement_scores[1]:
#                 print("Pass {0} {1}".format(student_score, requirement_scores))
#                 return True
#
#     print("requirementScoreTest(JLPT): (False, {0}, {1}, {2})".format(student_level, student_score, requirement_scores))
#     return False


def requirementSlotTest(student_level, requirement_slots):
    # if student_level in Level.Under.value or requirement_slots[1] < 0:
    #     if requirement_slots[0] > 0:
    #         return True
    # elif student_level in Level.Graduate.value:
    #     if requirement_slots[1] > 0:
    #         return True

    if student_level == "Under" or requirement_slots[1] < 0:
        if requirement_slots[0] > 0:
            return True
    elif student_level in ["Master", "Graduate"]:
        if requirement_slots[1] > 0:
            return True

    print("requirementSlotTest: (False, {0}, {1})".format(student_level, requirement_slots))
    return False

def requirementExchangeTermTest(student_level, student_term, requirement_term):
    if len(student_term[0]) < 1:
        return True

    # if student_level in Level.Under.value:
    #     for term in student_term:
    #         if term in requirement_term[0]:
    #             return True
    #
    # if student_level in Level.Graduate.value:
    #     if len(requirement_term) > 1:
    #         for term in student_term:
    #             if term in requirement_term[1]:
    #                 return True
    #     else:
    #         for term in student_term:
    #             if term in requirement_term[0]:
    #                 return True

    if student_level == "Under":
        for term in student_term:
            if term in requirement_term[0]:
                return True

    if student_level in ["Master", "Graduate"]:
        # if requirement_term[1] and len(requirement_term[1]) > 0:
        if len(requirement_term) > 1:
            for term in student_term:
                if term in requirement_term[1]:
                    return True
        else:
            for term in student_term:
                if term in requirement_term[0]:
                    return True

    print("requirementExchangeTermTest: (False, {0}, {1})".format(student_term, requirement_term))
    return False