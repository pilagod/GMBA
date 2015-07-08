__author__ = 'pilagod'

import operator
from xlwt import *
from xlrd import *
from userDefinedClass.Student import *
from userDefinedClass.School import *
from userDefinedClass.TextColors import *

input_student_book = open_workbook('./data/student.xlsx')
grade_sheet = input_student_book.sheet_by_name("grade")

input_school_book = open_workbook('./data/school.xlsx')
school_sheet = input_school_book.sheet_by_name("school")

students = getStudentData(grade_sheet)
print(students)
schools = getSchoolData(school_sheet)

YES_OPTIONS = ["Yes", "YES", "yes", "Y", "y"]
NO_OPTIONS = ["No", "NO", "no", "N", "n"]

placement_results = {}

for student in (sorted(students.values(), key=operator.attrgetter('rank'))):

    print(TextColors.WARNING + "({0.rank:0.0f}) {0.name} {0.level} GPA({0.gpa}) {0.toefl} {0.ielts}\n".format(student) + TextColors.BOLD)

    for will in student.wills:
        will = will.lower()
        print(TextColors.OKBLUE + "School:{0}, Level:{1.level}, Slots:{1.slots}, GPA:{1.gpa}, TOEFL:{1.toefl}, IELTS:{1.ielts}, TOEIC:{1.toeic}, JLPT:{1.jlpt}".format(will, schools[will]) + TextColors.ENDC)

        if requirementLevelTest(student.level, schools[will].level) and \
                requirementSlotTest(student.level, schools[will].slots) and \
                requirementGpaTest(student.gpa, schools[will].gpa):

            requirementToeflTest = requirementScoreTest(student.level, student.toefl, schools[will].toefl)
            requirementIeltsTest = requirementScoreTest(student.level, student.ielts, schools[will].ielts)
            # requirementToeicTest = requirementScoreTest(student.level, student.toeic, schools[will].toeic)
            # requirementJLPTTest = requirementJLPTScoreTest(student.level, student.jlpt, schools[will].jlpt)

            # print(requirementToeflTest)
            # print(requirementIeltsTest)
            # print(requirementToeicTest)
            # print(requirementJLPTTest)

            # scoreTests = (requirementToeflTest or requirementIeltsTest or requirementToeicTest or requirementJLPTTest) or \
            #              (((not requirementToeflTest) and (schools[will].toefl[0] is None)) and
            #               ((not requirementIeltsTest) and (schools[will].ielts[0] is None)) and
            #               ((not requirementToeicTest) and (schools[will].toeic[0] is None)) and
            #               ((not requirementJLPTTest) and (schools[will].jlpt[0] <= -1)))

            scoreTests = (requirementToeflTest or requirementIeltsTest) or \
                         (((not requirementToeflTest) and (schools[will].toefl[0] is None)) and
                          ((not requirementIeltsTest) and (schools[will].ielts[0] is None)))


            # if (requirementToeflTest or requirementIeltsTest or requirementToeicTest or requirementJLPTTest) or \
            #         (((not requirementToeflTest) and (schools[will].toefl[0] is None)) and
            #              ((not requirementIeltsTest) and (schools[will].ielts[0] is None)) and
            #              ((not requirementToeicTest) and (schools[will].toeic[0] is None)) and
            #              ((not requirementJLPTTest) and (schools[will].jlpt[0] < -1))):

            if scoreTests or student.remark != "":

                print("Scores Test Pass.")

                if schools[will].others.rstrip() != "":
                    pass_or_not = input(
                        TextColors.WARNING + "【Serial No】{0.serial_no:0.0f} 【Student ID】{0.student_id} 【Name】{0.name}".format(student) + TextColors.BOLD + "\n" + \
                        TextColors.FAIL + schools[will].others + "?(y/n):" + TextColors.ENDC
                    )
                    while pass_or_not not in (YES_OPTIONS + NO_OPTIONS):
                        pass_or_not = input(TextColors.FAIL + "Enter 'y' or 'n'(y/n):" + TextColors.ENDC)
                    if pass_or_not in NO_OPTIONS:
                        continue

                placement_results[student.serial_no] = {
                    # "rank": student.rank,
                    # "level": student.level,
                    "Student ID": student.student_id,
                    "Name": student.name,
                    "Name(English)": student.english_name,
                    "Date of Birth(MM/DD/YY)": student.birth,
                    "Gender": student.gender,
                    "E-mail": student.email,
                    "Cell Phone": student.phone,
                    "Level of Study": student.level,
                    "Year of Study": student.study_year,
                    "Department": student.department,
                    "Assigned School": will,
                    "Assigned School(Chinese)": schools[will].school_name_chinese,
                    "Remark": ""
                }
                # print(placement＿results[student.serial_no])
                print("Pass.\n")
                if student.level in "Undergraduates" or schools[will].slots[1] < 0:
                    schools[will].slots[0] -= 1
                elif student.level in "Masters" or "Graduates":
                    schools[will].slots[1] -= 1
                break
            else:
                print("Scores Test Fails.\n")
                continue
        else:
            print("Level, Slots, Gpa Test Falis.\n")
            continue

print(placement_results)

output_book = Workbook(encoding='utf-8')
output_book_result_sheet = output_book.add_sheet('result')
col_names = [
    "Student ID", "Name", "Name(English)", "Date of Birth(MM/DD/YY)",
    "Gender", "E-mail", "Cell Phone", "Level of Study", "Year of Study",
    "Department", "Assigned School", "Assigned School(Chinese)", "Remark"
]

index = 0
for col_name in col_names:
    output_book_result_sheet.row(0).write(index, col_name)
    index += 1

# for key in next(iter (placement＿results.values())).keys():
#     output_book_result_sheet.row(row_num).write(index, key)
#     index += 1

index = 0
row_num = 1

# print(students.values())
# print(placement＿results.values())
#
# print(collections.OrderedDict(sorted(placement＿results.items())))
for key, values in sorted(placement_results.items()):
    print(values)
    for col_name in col_names:
        output_book_result_sheet.row(row_num).write(index, values[col_name])
        index += 1
    index = 0
    row_num +=1

output_book.save('./result.xls')