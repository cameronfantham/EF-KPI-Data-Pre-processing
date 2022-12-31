import numpy as np
import pandas as pd
import string

def csv():
    #Verify .csv files in folder
    possible_csvs = ["ByTeacher_DetailData.csv"]
    for i in range(1,10):
        possible_csvs.append("ByTeacher_DetailData ({}).csv".format(i))
    csvs = possible_csvs.copy()
    for i in possible_csvs:
        try:
            pd.read_csv(i, encoding='utf-16')
        except FileNotFoundError:
            csvs.remove(i)
    #Select a .csv file
    if len(csvs) == 1: answer = csvs[0]
    elif len(csvs) == 0: answer = 0
    else:
        print("Found {} .csv files:".format(len(csvs)))
        print(*csvs, sep='\n')
        while True:
            answer = input("Please select one (by name or number) or type exit\n")
            if answer in csvs: break
            elif int(answer) in range(1,len(csvs)+1):
                answer = csvs[int(answer)-1]
                break
            elif answer == "e" or answer == "exit" or answer == "Exit": break
            else:
                print("Invalid answer")
                continue
    return answer
def username(df):
    while True:
        names_list = pd.unique(df["Reservation Teacher"])
        name = input("Please type 'summary', 'all', or enter a full name\n(type exit/e to exit)\n")
        if name.lower() in [x.lower() for x in names_list]: break
        elif name.lower() == "all" or name.lower() == "summary": break
        elif name.lower() == "e" or name.lower() == "exit": return 0
        else:
            print("Name not recognised. Please enter valid name.")
            continue
    return name.lower()

def generate_kpi_report(df, name):
    if name == 0:
        print("Exiting program")
        return 0
    print("Generating KPI Report")
    #Generate report for teacher
    if name != "all" and name != "summary":
        df = df[df["Reservation Teacher"] == name]
        df.drop("Reservation Teacher", axis=1, inplace=True)
    df.fillna('0.0', inplace=True)
    df = df[df["Group ID"] != '0.0']
    df.set_index("Group ID", inplace=True)
    for i in list(df):
        if isinstance(df[i][0], str):
            if "%" in df[i][0]:
                df[i] = df[i].str.rstrip("%").astype(float)

    if name == "all":
        for teacher in np.unique(df["Reservation Teacher"]):
            df_teacher = df.copy()
            df_teacher = df_teacher[df_teacher["Reservation Teacher"] == teacher]
            df_teacher.drop("Reservation Teacher", axis=1, inplace=True)
            writer = pd.ExcelWriter("{} KPI Report.xlsx".format(string.capwords(teacher)), engine='xlsxwriter')
            df_teacher.style.set_properties(**{'text-align': 'center'}).to_excel(writer, sheet_name="Teacher Detail")
            #df.style.applymap(lambda _:'text-align: center').to_excel(writer, sheet_name="Teacher Detail")
            workbook, worksheet = writer.book, writer.sheets["Teacher Detail"]
            workbook.add_format().set_text_wrap()
            for column in df_teacher:
                #column_length = max(df[column].astype(str).max(), len(column))
                column_index = df_teacher.columns.get_loc(column)
                worksheet.set_column(column_index, column_index, len(column))
            worksheet.conditional_format('B2:C19', {'type': '3_color_scale',
                                                    'min_value': '0',
                                                    'max_value': '124.3',
                                                    'min_color': '#FF0F0F',
                                                    'mid_color': '#FFFF00',
                                                    'max_color': '#00F000'})
            worksheet.conditional_format('D2:I19', {'type': '3_color_scale',
                                                    'min_value': '0',
                                                    'mid_value': '50',
                                                    'max_value': '100',
                                                    'min_color': '#FF0F0F',
                                                    'mid_color': '#FFFF00',
                                                    'max_color': '#00F000'})
            writer.save()
    else:
        writer = pd.ExcelWriter("{} KPI Report.xlsx".format(string.capwords(name)), engine='xlsxwriter')
        df.reset_index(inplace=True)
        df.style.set_properties(**{'text-align': 'center'}).to_excel(writer, sheet_name="Teacher Detail")
        #df.style.applymap(lambda _:'text-align: center').to_excel(writer, sheet_name="Teacher Detail")
        df.set_index("Group ID", inplace=True)
        workbook, worksheet = writer.book, writer.sheets["Teacher Detail"]
        workbook.add_format().set_text_wrap()
        for column in df:
            #column_length = max(df[column].astype(str).max(), len(column))
            column_index = df.columns.get_loc(column)
            worksheet.set_column(column_index, column_index, len(column))

        worksheet.conditional_format('B2:C19', {'type': '3_color_scale',
                                                'min_value': '0',
                                                'max_value': '124.3',
                                                'min_color': '#FF0F0F',
                                                'mid_color': '#FFFF00',
                                                'max_color': '#00F000'})
        worksheet.conditional_format('D2:I19', {'type': '3_color_scale',
                                                'min_value': '0',
                                                'mid_value': '50',
                                                'max_value': '100',
                                                'min_color': '#FF0F0F',
                                                'mid_color': '#FFFF00',
                                                'max_color': '#00F000'})
        writer.save()

try:
    #Choose which columns to use
    usecols = ["Reservation Teacher", "Group ID", "%Attendance MarkedOnTime",
               "TeachingTime (ACH)", "%TeacherTaskCompletion", "%PT App Use",
               "%SkillTestCompleted", "%Teacher-Led Skills Completion",
               "%BS CanDo Completion"]
    colnames = {"Reservation Teacher": "Reservation\nTeacher",
             "Group ID": "Group ID",
             "%Attendance MarkedOnTime": "Attendance\nMarked On Time (%)",
             "TeachingTime (ACH)": "Teaching\nTime (ACH)",
             "%TeacherTaskCompletion": "Teacher Task\nCompletion (%)",
             "%PT App Use": "PT App\nUse (%)",
             "%SkillTestCompleted": "Skill Test\nCompleted (%)",
             "%Teacher-Led Skills Completion": "Teacher-Led Skills\nCompletion (%)",
             "%BS CanDo Completion": "BS Can Do\nCompletion (%)"}
    #Read data into pandas data frames
    df = pd.read_csv(csv(), sep='\t', usecols=usecols,
                         index_col=False, encoding='utf-16')

    generate_kpi_report(df, username(df))

#except (FileNotFoundError, ValueError):
#    print("No .csv files found. Exiting program")
except NameError:
    print("Exiting program")
