# __author__ = 'pilagod'

from userDefinedClass.Interface import *

root = Tk()
app = interface(root)
app.mainloop()

# for s in wb.sheets():
#     print('Sheet:' + s.name)
#     for row in range(s.nrows):
#         values = []
#         for col in range(s.ncols):
#             values.append(str(s.cell(row, col).value))
#         print(','.join(values))
#     print()

# input_grade_book = open_workbook("./data/grade.xlsx")
# input_school_book = open_workbook("./data/school.xlsx")
# grade_sheet = input_grade_book.sheet_by_name('grade')
# school_sheet = input_school_book.sheet_by_name('school')
#
# students = getStudentData(grade_sheet)
# schools = getSchoolData(school_sheet)
#
# placement_results = {}
# for student in (sorted(students.values(), key=operator.attrgetter('rank'))):
#
#     print(TextColors.WARNING + "({0.rank:0.0f}) {0.name} {0.level} GPA({0.gpa}) {0.toefl} {0.ielts} {0.toeic} JLPT({0.jlpt})\n".format(student) + TextColors.BOLD)
#
#     for will in student.wills:
#         # input()
#         print(TextColors.OKBLUE + "School:{0}, Level:{1.level}, Slots:{1.slots}, GPA:{1.gpa}, TOEFL:{1.toefl}, IELTS:{1.ielts}, TOEIC:{1.toeic}, JLPT:{1.jlpt}".format(will, schools[will]) + TextColors.ENDC)
#         # print("({0}, {1}, {2})".format(student.level, school_requirements[will].level, requirementLevelTest(student.level, school_requirements[will].level)))
#         if requirementLevelTest(student.level, schools[will].level) and \
#                 requirementSlotTest(student.level, schools[will].slots) and \
#                 requirementGpaTest(student.gpa, schools[will].gpa):
#
#             requirementToeflTest = requirementScoreTest(student.level, student.toefl, schools[will].toefl)
#             requirementIeltsTest = requirementScoreTest(student.level, student.ielts, schools[will].ielts)
#             requirementToeicTest = requirementScoreTest(student.level, student.toeic, schools[will].toeic)
#             requirementJLPTTest = requirementJLPTScoreTest(student.level, student.jlpt, schools[will].jlpt)
#
#             scoreTests = (requirementToeflTest or requirementIeltsTest or requirementToeicTest or requirementJLPTTest) or \
#                         (((not requirementToeflTest) and (schools[will].toefl[0] is None)) and
#                          ((not requirementIeltsTest) and (schools[will].ielts[0] is None)) and
#                          ((not requirementToeicTest) and (schools[will].toeic[0] is None)) and
#                          ((not requirementJLPTTest) and (schools[will].jlpt[0] < -1)))
#
#             # if (requirementToeflTest or requirementIeltsTest or requirementToeicTest or requirementJLPTTest) or \
#             #         (((not requirementToeflTest) and (schools[will].toefl[0] is None)) and
#             #              ((not requirementIeltsTest) and (schools[will].ielts[0] is None)) and
#             #              ((not requirementToeicTest) and (schools[will].toeic[0] is None)) and
#             #              ((not requirementJLPTTest) and (schools[will].jlpt[0] < -1))):
#
#             if scoreTests or student.remark != "":
#
#                 if schools[will].others != "":
#                     pass_or_not = input(
#                         TextColors.WARNING + "【Serial No】{0.serial_no:0.0f} 【Name】{0.name}".format(student) + TextColors.BOLD + "\n" + \
#                         TextColors.FAIL + schools[will].others + "?(y/n):" + TextColors.ENDC
#                     )
#                     while pass_or_not not in (YES_OPTIONS + NO_OPTIONS):
#                         pass_or_not = input(TextColors.FAIL + "Enter 'y' or 'n'(y/n):" + TextColors.ENDC)
#                     if pass_or_not in NO_OPTIONS:
#                         continue
#
#                 placement_results[student.serial_no] = {
#                     # "rank": student.rank,
#                     # "level": student.level,
#                     "Student ID": student.student_id,
#                     "Name": student.name,
#                     "Name(English)": student.english_name,
#                     "Date of Birth(MM/DD/YY)": student.birth,
#                     "Gender": student.gender,
#                     "E-mail": student.email,
#                     "Cell Phone": student.phone,
#                     "Level of Study": student.level,
#                     "Year of Study": student.study_year,
#                     "Department": student.department,
#                     "Assigned School": will,
#                     "Assigned School(Chinese)": schools[will].school_name_chinese,
#                     "Remark": ""
#                 }
#                 # print(placement＿results[student.serial_no])
#                 # print()
#                 if student.level in "Undergraduates" or schools[will].slots[1] < 0:
#                     schools[will].slots[0] -= 1
#                 elif student.level in "Masters" or "Graduates":
#                     schools[will].slots[1] -= 1
#                 break
#             else:
#                 continue
#         else:
#             continue
#
# for result in placement_results.values():
#     print(result)