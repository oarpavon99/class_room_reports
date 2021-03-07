from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from pprint import pprint
from secrets import fcaq_pass, fcaq_email
from tkinter import messagebox
from tkcalendar import Calendar
import os, pickle, gspread, yagmail, datetime, tkinter, sys


#######################################
def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


class ClassroomControl:
    def __init__(self):
        def init_email():
            try:
                yag_temp = yagmail.SMTP(user=fcaq_email, password=fcaq_pass)
            except:
                print("E mail connection was not possible!")
            else:
                print("E-mail connection was successful.")
                return yag_temp

        # This  will instantiate the class that will handle the auth flow for the Google Classroom API
        self.flow = InstalledAppFlow.from_client_secrets_file(
            "creds_00.json",
            scopes=[
                "https://www.googleapis.com/auth/classroom.courses.readonly",
                "https://www.googleapis.com/auth/classroom.rosters.readonly",
                "https://www.googleapis.com/auth/classroom.coursework.students",
                "https://www.googleapis.com/auth/classroom.profile.emails",
                "https://www.googleapis.com/auth/spreadsheets",
                "https://www.googleapis.com/auth/drive",
            ],
        )
        self.yag_service = None
        self.courses = {}
        self.students = {}
        self.guardians = {}
        self.coursework = {}
        self.missed_coursework = {}
        self.coursework_rep = {}
        self.report_coursework = {}
        self.creds = None
        self.yag_service = init_email()
        # Creates the service to connect with the API
        self.create_service(self.get_access_token())
        # Gets courses
        self.get_courses()
        # Gets students
        self.get_students(self.courses)
        # Gets guardians
        self.get_guardians(self.students, self.courses)

    # IMPROVEMENT: USE THE REFRESH TOKEN
    def get_access_token(self):
        if os.path.exists("creds.pickle"):
            with open("creds.pickle", "rb") as file:
                self.creds = pickle.load(file)
        else:
            # This will run a local server to grant access:
            self.flow.run_local_server(
                port=8080, prompt="consent", authorization_prompt_message=""
            )
            with open("creds.pickle", "wb") as file:
                pickle.dump(self.flow.credentials, file)
            self.creds = self.flow.credentials
        return self.creds

    def create_service(self, creds):
        # This creates the service for the Google Classroom
        self.service = build("classroom", "v1", credentials=creds)
        """Once the end-user has authorized access to the chosen scopes, the authorization can be used to authorize 
        the gspread access ALSO! """
        self.client = gspread.authorize(creds)

    # This will get all the current courses based on the courseState and the section name
    # GENERALIZATION HERE, HOW CAN I FILTER THE COURSES BETTER?
    def get_courses(self):
        request = self.service.courses().list(courseStates="ACTIVE", teacherId="me")
        response = request.execute()
        for i, course in enumerate(response["courses"]):
            if "2021" in course["creationTime"]:
                self.courses.setdefault(course["id"], course["name"])
        print(f"self.courses = {self.courses}")

    # This will get the students for each course
    # GENERALIZATION HERE, WHAT HAPPENS IF THERE ARE MORE THAN 30 STUDENTS?
    def get_students(self, courses_dict):
        if os.path.exists("students.pickle"):
            with open("students.pickle", "rb") as file:
                self.students = pickle.load(file)
                print("self.students has been loaded")
        else:
            for course_id in courses_dict.keys():
                request = self.service.courses().students().list(courseId=course_id)
                response = request.execute()
                print(f"GETTING STUDENTS IN CLASS: {courses_dict[course_id]}")
                for j, student in enumerate(response["students"]):
                    email = student["profile"]["emailAddress"]
                    student_id = student["profile"]["id"]
                    name = [
                        student["profile"]["name"]["familyName"].split()[0],
                        student["profile"]["name"]["givenName"].split()[0],
                    ]
                    print(f"Student: {name}")
                    if course_id not in self.students.keys():
                        self.students[course_id] = {}
                    self.students[course_id].setdefault(student_id, [email, name])
            with open("students.pickle", "wb") as file:
                pickle.dump(self.students, file)
                print("self.students has been dumped")
        # pprint(self.students)

    """Guardians are only available for administrators and I am not an administrator, so no possible fetch of the 
    guardians coming from the API can be done. 
    That's why I have to fetch the guardians from the Google Spreadsheet."""

    def get_guardians(self, students_dict, courses_dict):
        # ASSUMPTION: The course name has a format with a letter designating each course, example: 9EGB E - Design
        def get_parallel(course):
            temp = courses_dict[course].split()
            for element in temp:
                if len(element) == 1 and element.isalpha():
                    print(f"{element} will be returned")
                    return element

        ###################################
        if os.path.exists("guardians.pickle"):
            with open("guardians.pickle", "rb") as file:
                self.guardians = pickle.load(file)
                print(f"self.guardians has been loaded")
        else:
            sheet = self.client.open("Students contact information")
            for course in students_dict.keys():
                # Get the spreadsheet for the given course:
                sheet_temp = sheet.worksheet(get_parallel(course))
                # Gets the whole range of  values student - guardians
                values_temp = sheet_temp.get_all_values()
                # print(values_temp)
                for row in values_temp:
                    if "@" not in row[0]:
                        continue
                    else:
                        student_email = row[0]
                        guardian_temp = row[1].split()
                        guardian_emails = []
                        for i in guardian_temp:
                            guardian_emails.append(i.strip().strip(","))
                        self.guardians.setdefault(student_email, guardian_emails)
            with open("guardians.pickle", "wb") as file:
                pickle.dump(self.guardians, file)
                print("self.guardians has been dumped")
        # pprint(self.guardians)

    def get_missed_coursework(self, course_choice):
        request = self.service.courses().courseWork().list(courseId=course_choice)
        response = request.execute()
        for task in response["courseWork"]:
            if "dueDate" in task.keys():
                due_date = datetime.datetime(
                    day=task["dueDate"]["day"],
                    month=task["dueDate"]["month"],
                    year=task["dueDate"]["year"],
                )
            else:
                due_date = datetime.datetime.max
            self.coursework.setdefault(task["id"], [task["title"], due_date])
            print(self.coursework[task["id"]])
        for i in self.coursework.keys():
            request = (
                self.service.courses()
                .courseWork()
                .studentSubmissions()
                .list(courseId=course_choice, courseWorkId=i)
            )
            response = request.execute()
            print(
                f"------------------------------------------------------------------\nGETTING INFO FOR COURSEWORKID: {self.coursework[i][0]}"
            )
            # pprint(response["studentSubmissions"])
            for submission in response["studentSubmissions"]:
                if submission["state"] not in ["TURNED_IN", "RETURNED"]:
                    if datetime.datetime.today() > self.coursework[i][1]:
                        # print("********************************")
                        # print("CASE 0: NOT TURNED IN AND OVERDUE.")
                        # print(f"DUE DATE: {self.coursework[i][1]}")
                        # print(datetime.datetime.today() - self.coursework[i][1])
                        # print(self.students[course_choice][submission["userId"]][0])
                        # print(f"submission[\"state\"]={submission['state']}")
                        if submission["userId"] not in self.missed_coursework.keys():
                            self.missed_coursework.setdefault(
                                submission["userId"],
                                [self.coursework[i]],
                            )
                        else:
                            self.missed_coursework[submission["userId"]].append(
                                self.coursework[i]
                            )
                if submission["state"] == "RETURNED":
                    if "assignedGrade" in submission.keys():
                        if submission["assignedGrade"] < 60:
                            # print("********************************")
                            # print("CASE 1: RETURNED WITH 50 (STUDENT DID NOT DO ANYTHING).")
                            # print(self.students[course_choice][submission["userId"]][0])
                            # print(f"submission[\"assignmentSubmission\"]={submission['assignmentSubmission']}")
                            # print(len(submission["assignmentSubmission"]))
                            if (
                                submission["userId"]
                                not in self.missed_coursework.keys()
                            ):
                                self.missed_coursework.setdefault(
                                    submission["userId"],
                                    [self.coursework[i]],
                                )
                            else:
                                self.missed_coursework[submission["userId"]].append(
                                    self.coursework[i]
                                )
        pprint(self.missed_coursework)

    def get_reports_coursework(self, initial_date, final_date, course_choice):
        print("//////Printing reported tasks:")
        request = self.service.courses().courseWork().list(courseId=course_choice)
        response = request.execute()
        ######################################
        for task in response["courseWork"]:
            if "dueDate" in task.keys():
                due_date = datetime.datetime(
                    day=task["dueDate"]["day"],
                    month=task["dueDate"]["month"],
                    year=task["dueDate"]["year"],
                )
            else:
                due_date = datetime.datetime.max
            if initial_date < due_date < final_date:
                self.coursework_rep.setdefault(task["id"], task["title"])
                print(self.coursework_rep[task["id"]])
        ######################################
        for task in self.coursework_rep.keys():
            request = (
                self.service.courses()
                .courseWork()
                .studentSubmissions()
                .list(courseId=course_choice, courseWorkId=task)
            )
            response = request.execute()
            print(
                f"------------------------------------------------------------------\nGETTING INFO FOR COURSEWORKID: {self.coursework_rep[task][0]}"
            )
            # pprint(response["studentSubmissions"])
            for submission in response["studentSubmissions"]:
                if (
                    submission["state"] == "RETURNED"
                    and "assignedGrade" in submission.keys()
                ):
                    if submission["userId"] not in self.report_coursework.keys():
                        self.report_coursework.setdefault(
                            submission["userId"],
                            [],
                        )
                        self.report_coursework[submission["userId"]].append(
                            [self.coursework_rep[task], submission["assignedGrade"]]
                        )
                    else:
                        self.report_coursework[submission["userId"]].append(
                            [self.coursework_rep[task], submission["assignedGrade"]]
                        )
        # pprint(self.report_coursework)

    """There are several parameters for this function:
    type: string. Allowed values: MISSING: sends only the missing tasks. REPORT: sends all the grades for every student.
    who: string. Allowed values: STUDENT, GUARDIANS, VICE.
    course: string. The course Id for the given parallel.
    student: string. Student Id.
    """

    def send_mails(self, type, who, course, preview_send):
        # This configures the email that will be sent individually to each student
        def config_message(student):
            addressees = [self.students[course][student][0]]
            subject = ""
            ##########################################
            ##########################################
            ##########################################
            if type == "MISSING":
                subject = f"{self.students[course][student][1][1]} {self.students[course][student][1][0]} - Missing Activities. DESIGN & TECH."
                if who == "STUDENT":
                    message = f"Dear {self.students[course][student][1][1]}:\n\n"
                    message += "I hope this email finds you well. You are currently missing the following activities:\n\n"
                    for act in self.missed_coursework[student]:
                        message += f"{act[0]}\nDue date: {(act[1] - datetime.timedelta(days=1)).strftime('%B %d')}.\n"
                    message += (
                        "\nPlease make sure to submit everything as soon as possible, so your grade does not get affected. "
                        "Remember that if you need extra help, you can always stay at office hours after school on "
                        "Monday, Tuesday or Thursday.\n\nOmar Ramirez "
                    )
                elif who == "GUARDIANS" or who == "VICE":
                    message = f"Dear Parents:\n\n"
                    message += "I hope this email finds you well. Your son/daughter is currently missing the following activities:\n\n"
                    for act in self.missed_coursework[student]:
                        message += f"{act[0]}\nDue date: {(act[1] - datetime.timedelta(days=1)).strftime('%B %d')}.\n"
                    message += (
                        "\nThe reason why I'm writing you is because I would like to request your support reminding your son/daughter to finish all the pending tasks."
                        " I am worried that this may affect his/her final grade for the partial and my intention is to prevent this from happenning."
                        " Please remember that if your student needs extra help, he/she can always stay at office hours after school on "
                        "Mondays, Tuesdays or Thursdays. For further information please do not hesitate to contact me.\n\nThanks in advance! \n\nOmar Ramirez "
                    )
                    ##########################################
                    if who == "GUARDIANS":
                        for parent_email in self.guardians[
                            self.students[course][student][0]
                        ]:
                            addressees.append(parent_email)
                    ##########################################
                    elif who == "VICE":
                        for parent_email in self.guardians[
                            self.students[course][student][0]
                        ]:
                            addressees.append(parent_email)
                        addressees.append("alizarralde@fcaq.k12.ec")
            ##########################################
            ##########################################
            ##########################################
            elif type == "REPORT":
                subject = f"{self.students[course][student][1][1]} {self.students[course][student][1][0]} - Google Classroom Grades Report. DESIGN & TECH."
                message = f"Dear {self.students[course][student][1][1]}:\n\n"
                message += "I hope this email finds you well. Find below the current status of your graded Google Classroom tasks:\n\n"
                for act in self.report_coursework[student]:
                    message += f"{act[0]}\nGrade: {act[1]}.\n"
                message += (
                    "\nIf you are missing something, please make sure to submit it as soon as possible, so your partial grade does not get affected. "
                    "Remember that if you need extra help, you can always stay at office hours after school on "
                    "Mondays, Tuesdays or Thursdays.\n\nBest Regards,\nOmar Ramirez "
                )

                if who in ["GUARDIANS", "VICE"]:
                    for parent_email in self.guardians[
                        self.students[course][student][0]
                    ]:
                        addressees.append(parent_email)

            ##########################################
            print(addressees)
            print(subject)
            print(message, "\n\n")
            return addressees, subject, message

        if self.yag_service is not None:
            if type == "MISSING":
                for lazy_student in self.missed_coursework.keys():
                    # Config email
                    send_to, subj, what = config_message(lazy_student)
                    if preview_send == "SEND":
                        # Send email
                        try:
                            self.yag_service.send(
                                to=send_to, subject=subj, contents=what
                            )
                        except:
                            print("NO EMAIL WAS SENT.")
            elif type == "REPORT":
                for student in self.report_coursework.keys():
                    # Config email
                    send_to, subj, what = config_message(student)
                    if preview_send == "SEND":
                        # Send email
                        try:
                            self.yag_service.send(
                                to=send_to, subject=subj, contents=what
                            )
                        except:
                            print("NO EMAIL WAS SENT.")
        else:
            print("NO EMAIL WAS SENT.")


class MainApplication(tkinter.Frame):
    def __init__(self, master):
        super().__init__()
        self.bg = "gray15"
        self.bg2 = "gray80"
        self.bg3 = "RoyalBlue4"
        self.fg = "white"
        self.main_window = tkinter.Frame(master)
        self.main_window.configure(bg=self.bg)
        ############################
        self.type_of_report = tkinter.StringVar(self.main_window)
        self.type_of_report.set("MISSING")
        self.addressees = tkinter.StringVar(self.main_window)
        self.addressees.set("STUDENT")
        self.preview_or_send = tkinter.StringVar(self.main_window)
        self.preview_or_send.set("PREVIEW")
        self.date_init = datetime.datetime.now()
        self.date_final = datetime.datetime.now()
        ############################
        self.frame1 = tkinter.Frame(self.main_window)
        self.frame2 = tkinter.LabelFrame(self.main_window)
        self.frame2b = tkinter.Frame(self.frame2)
        self.frame2c = tkinter.Frame(self.frame2b)
        self.frame3 = tkinter.LabelFrame(self.main_window)
        self.frame31 = tkinter.LabelFrame(self.frame3)
        self.frame32 = tkinter.LabelFrame(self.frame3)
        self.frame33 = tkinter.LabelFrame(self.frame3)
        self.frame34 = tkinter.LabelFrame(self.frame3)
        ############################
        self.frame1.configure(bg=self.bg)
        self.frame3.configure(bd=5, bg=self.bg2)
        self.frame31.configure(bg=self.bg)
        self.frame32.configure(bg=self.bg)
        self.frame33.configure(bg=self.bg)
        ############################
        self.title = tkinter.Label(
            self.frame1,
            text="Google Classroom Grade reports.",
            bg=self.bg,
            font="Consolas 20",
            fg=self.fg,
        )
        ############################
        self.lbl_classes = tkinter.Label(
            self.frame2,
            text="Available Classes",
            font="Gotham 12",
            bg=self.bg,
            fg=self.fg,
        )
        self.classes_scroll_y = tkinter.Scrollbar(self.frame2c, orient="vertical")
        self.classes_scroll_x = tkinter.Scrollbar(self.frame2b, orient="horizontal")
        self.classes_listbox = tkinter.Listbox(
            self.frame2c,
            yscrollcommand=self.classes_scroll_y.set,
            xscrollcommand=self.classes_scroll_x.set,
            height=7,
        )
        self.classes_scroll_y.config(command=self.classes_listbox.yview)
        self.classes_scroll_x.config(command=self.classes_listbox.xview)
        ############################
        self.rb_missing = tkinter.Radiobutton(
            self.frame31,
            text="MISSING",
            font="Gotham 10",
            bg=self.bg2,
            variable=self.type_of_report,
            value="MISSING",
            command=lambda: self.type_of_report.set("MISSING"),
        )
        self.rb_report = tkinter.Radiobutton(
            self.frame31,
            text="REPORT",
            font="Gotham 10",
            bg=self.bg2,
            variable=self.type_of_report,
            value="REPORT",
            command=lambda: self.type_of_report.set("REPORT"),
        )
        ############################
        self.rb_student = tkinter.Radiobutton(
            self.frame32,
            text="Student",
            font="Gotham 10",
            bg=self.bg2,
            variable=self.addressees,
            value="STUDENT",
            command=lambda: self.addressees.set("STUDENT"),
        )
        self.rb_guardians = tkinter.Radiobutton(
            self.frame32,
            text="Guardians",
            font="Gotham 10",
            bg=self.bg2,
            variable=self.addressees,
            value="GUARDIANS",
            command=lambda: self.addressees.set("GUARDIANS"),
        )
        self.rb_vice = tkinter.Radiobutton(
            self.frame32,
            text="Vice-principal",
            font="Gotham 10",
            bg=self.bg2,
            variable=self.addressees,
            value="VICE",
            command=lambda: self.addressees.set("VICE"),
        )
        ############################
        self.rb_preview = tkinter.Radiobutton(
            self.frame33,
            text="PREVIEW",
            font="Gotham 10 bold",
            bg=self.bg2,
            variable=self.preview_or_send,
            value="PREVIEW",
            command=lambda: self.preview_or_send.set("PREVIEW"),
        )
        self.rb_send = tkinter.Radiobutton(
            self.frame33,
            text="SEND",
            font="Gotham 10 bold",
            bg=self.bg2,
            variable=self.preview_or_send,
            value="SEND",
            command=lambda: self.preview_or_send.set("SEND"),
        )
        ############################
        self.button_run = tkinter.Button(
            self.frame34,
            text="Send mails.",
            font="Gotham 20",
            command=self.run,
            bg=self.bg3,
            fg=self.fg,
        )
        ############################
        self.cc = ClassroomControl()
        self.display_elements()

    def display_elements(self):
        self.main_window.grid(row=0, column=0)
        ############################
        self.frame1.grid(row=0, column=0, columnspan=2, padx=10, pady=10)
        self.frame2.grid(row=1, column=0, padx=10, pady=10)
        self.frame3.grid(row=1, column=1)
        self.frame31.grid(row=0, column=0, sticky="W", pady=5, padx=5)
        self.frame32.grid(row=1, column=0, pady=5, padx=5)
        self.frame33.grid(row=2, column=0, sticky="W", pady=5, padx=5)
        self.frame34.grid(row=3, column=0, pady=5, padx=5)
        ############################
        self.title.grid(row=0, column=0)
        ############################
        self.lbl_classes.grid(row=0, column=0)
        self.frame2b.grid(row=1, column=0)
        self.frame2c.pack()
        self.classes_listbox.pack(side="left")
        self.classes_scroll_y.pack(side="right", fill="y")
        self.classes_scroll_x.pack(side="bottom", fill="x")
        for count, course in enumerate(sorted(self.cc.courses.keys())):
            self.classes_listbox.insert(
                tkinter.END,
                self.cc.courses[course],
            )
        ############################
        self.rb_missing.grid(row=0, column=0)
        self.rb_report.grid(row=0, column=1)
        ############################
        self.rb_student.grid(row=0, column=0)
        self.rb_guardians.grid(row=0, column=1)
        self.rb_vice.grid(row=0, column=2)
        ############################
        self.rb_preview.grid(row=0, column=0)
        self.rb_send.grid(row=0, column=1)
        ############################
        self.button_run.grid(row=0, column=0)

    def get_initial_final_date_to_report(self, course_key):
        def get_dates():
            temp = self.calendar_init.get_date().split("/")
            self.date_init = datetime.datetime(
                year=int("20" + temp[2]), month=int(temp[0]), day=int(temp[1])
            )
            temp = self.calendar_final.get_date().split("/")
            self.date_final = datetime.datetime(
                year=int("20" + temp[2]), month=int(temp[0]), day=int(temp[1])
            )
            if self.date_final - self.date_init > datetime.timedelta(days=1):
                message_send = ""
                if self.preview_or_send.get() == "SEND":
                    message_send = (
                        f" to the following addressees: '{self.addressees.get()}'"
                    )
                message = f"You will {self.preview_or_send.get()} a grade report for class '{self.classes_listbox.get(tkinter.ANCHOR)}'{message_send}.\n\nDue date time frame:\n{self.date_init.strftime('%B %d %Y')} to {self.date_final.strftime('%B %d %Y')}.\n\nIs it OK?"
                if messagebox.askokcancel(
                    "Google Classroom Grade reports. - REPORT choose dates.", message
                ):
                    self.calendar_window.destroy()
                    self.cc.get_reports_coursework(
                        initial_date=self.date_init,
                        final_date=self.date_final,
                        course_choice=course_key,
                    )
                    self.cc.send_mails(
                        type="REPORT",
                        who=self.addressees.get(),
                        course=course_key,
                        preview_send=self.preview_or_send.get(),
                    )
                else:
                    self.calendar_window.destroy()
            else:
                messagebox.showerror(
                    "Google Classroom Grade reports. - REPORT choose dates.",
                    "The date range is too small. No possible report generation.",
                )

        ################################
        self.calendar_window = tkinter.Toplevel(self.main_window)
        self.calendar_window.configure(bg=self.bg)
        self.calendar_window.title(
            "Google Classroom Grade reports. - REPORT choose dates."
        )
        self.calendar_window.geometry("+700+350")
        self.calendar_window.iconbitmap(resource_path("icon.ico"))
        self.frame_c0 = tkinter.LabelFrame(self.calendar_window)
        self.frame_c0.configure(bg=self.bg)
        self.frame_c1 = tkinter.LabelFrame(self.calendar_window)
        self.frame_c1.configure(bg=self.bg)
        self.frame_c2 = tkinter.LabelFrame(self.calendar_window)
        self.frame_c2.configure(bg=self.bg)
        self.frame_c3 = tkinter.LabelFrame(self.calendar_window)
        self.frame_c3.configure(bg=self.bg)
        self.calendar_title = tkinter.Label(
            self.frame_c0,
            text="Choose initial and final date for reports.",
            bg=self.bg,
            font="Consolas 15",
            fg=self.fg,
        )
        self.calendar_init_title = tkinter.Label(
            self.frame_c1,
            text="Initial date:",
            bg=self.bg,
            font="Gotham 12",
            fg=self.fg,
        )
        self.calendar_final_title = tkinter.Label(
            self.frame_c2, text="Final date:", bg=self.bg, font="Gotham 12", fg=self.fg
        )
        self.calendar_init = Calendar(self.frame_c1, selectmode="day")
        self.calendar_final = Calendar(self.frame_c2, selectmode="day")
        ############################
        self.get_dates_button = tkinter.Button(
            self.frame_c3,
            text="Get report dates.",
            font="Gotham 15",
            command=get_dates,
            bg=self.bg3,
            fg=self.fg,
        )
        ################################
        self.frame_c0.grid(row=0, column=0, columnspan=2, padx=10, pady=10)
        self.frame_c1.grid(row=1, column=0)
        self.frame_c2.grid(row=1, column=1)
        self.frame_c3.grid(row=2, column=0, columnspan=2)
        ################################
        self.calendar_title.grid(row=0, column=0)
        ################################
        self.calendar_init_title.grid(row=0, column=0)
        self.calendar_init.grid(row=1, column=0)
        ################################
        self.calendar_final_title.grid(row=0, column=0)
        self.calendar_final.grid(row=1, column=0)
        ################################
        self.get_dates_button.grid(row=0, column=0)

    def run(self):
        # Checks if the user chose a class:
        try:
            assert len(self.classes_listbox.get(tkinter.ANCHOR)) > 1
        except AssertionError:
            messagebox.showwarning(
                "Google Classroom Grade reports - No class was selected.",
                "Please select a class.",
            )
        else:
            print("let's run!")
            print(f"self.type_of_report.get() = {self.type_of_report.get()}")
            print(f"self.addressees.get() = {self.addressees.get()}")
            print(f"self.preview_or_send.get() = {self.preview_or_send.get()}")
            print(
                f"self.classes_listbox.get() = {self.classes_listbox.get(tkinter.ANCHOR)}"
            )
            if self.type_of_report.get() == "MISSING":
                course_key = [
                    key
                    for key, course in self.cc.courses.items()
                    if course == self.classes_listbox.get(tkinter.ANCHOR)
                ][0]
                message_send = ""
                if self.preview_or_send.get() == "SEND":
                    message_send = (
                        f" to the following addressees: '{self.addressees.get()}'"
                    )
                message = f"You will {self.preview_or_send.get()} a {self.type_of_report.get()} report for class '{self.classes_listbox.get(tkinter.ANCHOR)}'{message_send}.\n\nIs it OK?"
                if messagebox.askokcancel(
                        f"Google Classroom Grade reports. - {self.type_of_report.get()} tasks.", message
                ):
                    self.cc.get_missed_coursework(course_key)
                    self.cc.send_mails(
                        type="MISSING",
                        who=self.addressees.get(),
                        course=course_key,
                        preview_send=self.preview_or_send.get(),
                    )
            elif self.type_of_report.get() == "REPORT":
                course_key = [
                    key
                    for key, course in self.cc.courses.items()
                    if course == self.classes_listbox.get(tkinter.ANCHOR)
                ][0]
                self.get_initial_final_date_to_report(course_key)



if __name__ == "__main__":
    root = tkinter.Tk()
    root.title("Google Classroom Grade reports.")
    root.geometry("+300+300")
    root.iconbitmap(resource_path("icon.ico"))
    app = MainApplication(root)
    root.mainloop()

