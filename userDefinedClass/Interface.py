__author__ = 'pilagod'
import operator
from xlwt import *
from xlrd import *
from userDefinedClass.Student import *
from userDefinedClass.School import *
from userDefinedClass.TextColors import *
from userDefinedClass.ShareFunc import *
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox

class interface(Frame):
    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.input_file = ""
        self.output_dir = ""
        self.parent = master
        self.grid()
        self.createWidgets()

    def quit(self):
        self.parent.destroy()

    def createWidgets(self):
        self.inputStudentButton = Button(self)
        self.inputStudentButton["text"] = "Choose Student File"
        self.inputStudentButton["height"] = 1
        self.inputStudentButton["command"] = self.chooseInputStudentFile
        self.inputStudentButton.grid(row=0, column=0)

        self.inputStudentLabel = Label(self)
        # self.inputLabel["text"] = "Choose Student Data File"
        self.inputStudentLabel["width"] = 100
        self.inputStudentLabel["height"] = 1
        self.inputStudentLabel["borderwidth"] = 2
        self.inputStudentLabel["relief"] = SUNKEN
        self.inputStudentLabel.grid(row=0, column=1, columnspan=5)

        self.inputSchoolButton = Button(self)
        self.inputSchoolButton["text"] = "Choose School File"
        self.inputSchoolButton["height"] = 1
        self.inputSchoolButton["command"] = self.chooseInputSchoolFile
        self.inputSchoolButton.grid(row=1, column=0)

        self.inputSchoolLabel = Label(self)
        # self.inputLabel["text"] = "Choose Student Data File"
        self.inputSchoolLabel["width"] = 100
        self.inputSchoolLabel["height"] = 1
        self.inputSchoolLabel["borderwidth"] = 2
        self.inputSchoolLabel["relief"] = SUNKEN
        self.inputSchoolLabel.grid(row=1, column=1, columnspan=5)

        self.outputButton = Button(self)
        self.outputButton["text"] = "Choose Directory"
        self.outputButton["height"] = 1
        self.outputButton["command"] = self.chooseOutputDirectory
        self.outputButton.grid(row=2, column=0)

        self.outputLabel = Label(self)
        # self.outputLabel["text"] = "Choose Result Destination"
        self.outputLabel["width"] = 100
        self.outputLabel["height"] = 1
        self.outputLabel["borderwidth"] = 2
        self.outputLabel["relief"] = SUNKEN
        self.outputLabel.grid(row=2, column=1, columnspan=5)

        self.startButton = Button(self)
        self.startButton["text"] = "Start Placement"
        self.startButton["height"] = 2
        self.startButton["command"] = self.startPlacement
        self.startButton.grid(row=3, column=2, columnspan=2)


    def chooseInputStudentFile(self):
        self.inputStudentLabel["text"] = filedialog.askopenfilename()

    def chooseInputSchoolFile(self):
        self.inputSchoolLabel["text"] = filedialog.askopenfilename()

    def chooseOutputDirectory(self):
        self.outputLabel["text"] = filedialog.askdirectory()

    def startPlacement(self):

        (grade_sheet, school_sheet) = self.__getSheets()
        if grade_sheet == None or school_sheet == None:
            return

        students = getStudentData(grade_sheet)
        schools = getSchoolData(school_sheet)

        # Start Placement
        placement_results = self.__doPlacement(students, schools)

        # Output Placement Result
        self.__outputPlacementResult(placement_results)

    # Output: (grade_sheet, school_sheet)
    def __getSheets(self):

        inputStudentFile = self.inputStudentLabel["text"]
        inputSchoolFile = self.inputSchoolLabel["text"]

        try:
            input_student_book = open_workbook(inputStudentFile)
            grade_sheet = input_student_book.sheet_by_name("grade")

            input_school_book = open_workbook(inputSchoolFile)
            school_sheet = input_school_book.sheet_by_name("school")

        except XLRDError:
            messagebox.showerror("Error", "'File' or 'Sheet Name' Error")
            return (None, None)

        return (grade_sheet, school_sheet)


    # Output: placement_results
    def __doPlacement(self, students, schools):

        YES_OPTIONS = ["Yes", "YES", "yes", "Y", "y"]
        NO_OPTIONS = ["No", "NO", "no", "N", "n"]

        placement_results = {}

        for student in (sorted(students.values(), key=operator.attrgetter('rank'))):

            print(TextColors.WARNING + "({0.rank:0.0f}) {0.name}, {0.level}, {0.exchange_term}, GPA({0.gpa}), {0.toefl}, {0.ielts}\n".format(student) + TextColors.BOLD)

            for will in student.wills:
                will = caseAndSpaceIndif(will)
                print(TextColors.OKBLUE + "School:{0}, Level:{1.level}, Exchange:{1.exchange_term}, Slots:{1.slots}, GPA:{1.gpa}, TOEFL:{1.toefl}, IELTS:{1.ielts}, TOEIC:{1.toeic}".format(will, schools[will]) + TextColors.ENDC)

                if requirementLevelTest(student.level, schools[will].level) and \
                        requirementSlotTest(student.level, schools[will].slots) and \
                        requirementGpaTest(student.gpa, schools[will].gpa) and \
                        requirementExchangeTermTest(student.level, student.exchange_term, schools[will].exchange_term):

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
                            "Level of Study": Level.Under.name if student.level in Level.Under.value else Level.Graduate.name,
                            "Year of Study": student.study_year,
                            "Department": student.department,
                            "Assigned School": schools[will].school_code,
                            "Assigned School(Chinese)": schools[will].school_name_chinese,
                            "Remark": ""
                        }
                        # print(placement＿results[student.serial_no])
                        print("Pass.\n")
                        if student.level in Level.Under.value or schools[will].slots[1] < 0:
                            schools[will].slots[0] -= 1
                        elif student.level in Level.Graduate.value:
                            schools[will].slots[1] -= 1
                        break
                    else:
                        print("Scores Test Fails.\n")
                        continue
                else:
                    print("Level, Slots, Gpa, Exchange Term Test Falis.\n")
                    continue

        return placement_results

    def __outputPlacementResult(self, placement_results):

        outputDir = self.outputLabel["text"]
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

        if outputDir == "":
            outputDir = "./"

        try:
            output_book.save(outputDir + '/result.xls')
        except:
            messagebox.showerror("Error", "Saving Output File Error")

        messagebox.showinfo("Finish!", "Placement Finish!")
        self.quit()




