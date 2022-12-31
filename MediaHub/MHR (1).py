import numpy as np
import pandas as pd
import string as str

def csv():
    #Check .csv files in folder
    possible_csvs = ["MarkedStudent_DetailData.csv"]
    for i in range(1,10):
        possible_csvs.append("MarkedStudent_DetailData ({}).csv".format(i))
    csvs = possible_csvs.copy()
    for i in possible_csvs:
        try: pd.read_csv(i, encoding='utf-16')
        except FileNotFoundError: csvs.remove(i)
    #Select a .csv file
    if len(csvs) == 1: answer = csvs[0]
    elif len(csvs) == 0: answer = 0
    else:
        print("Found {} .csv files:".format(len(csvs)))
        print(*csvs, sep='\n')
        while True:
            answer = input("Please select file by name or number (type exit/e to exit)\n")
            if answer in csvs: break
            elif answer == "e" or answer == "exit" or answer == "Exit": break
            elif int(answer) in range(1,len(csvs)+1):
                answer = csvs[int(answer)-1]
                break
            else:
                print("Invalid answer")
                continue
    return answer
def username(df):
    while True:
        names_list = np.unique(df[["main_teacher", "co_teacher"]].values)
        name = input("Please type 'summary', 'all', or enter a full name\n(type exit/e to exit)\n")
        if name.lower() in [x.lower() for x in names_list]: break
        elif name.lower() == "all" or name.lower() == "summary": break
        elif name.lower() == "e" or name.lower() == "exit": return 0
        else:
            print("Name not recognised. Please enter valid name.")
            continue
    return name.lower()
def generate_mh_report(df, name):
    if name == 0:
        print("Exiting program")
        return 0
    df = df.sort_values(["main_teacher", "co_teacher", "group_code"])
    #Filter data frame for teacher's name
    if name != "all" and name != "summary":
        df = df[(df["main_teacher"] == name) | (df["co_teacher"] == name)]
    df_lt, df_af = df.copy(), df.copy()
    #Filter out 100% values
    hundreds = ["1", "100", "100%", "100.0%", "100.00%", "100.000%"]
    for i in hundreds:
        df_lt = df_lt[(df_lt["%Student ReceivedMedia File(Let'sTalk +Video)"] != i)]
        df_af = df_af[(df_af["%StudentReceived AcademicFeedback"] != i)]
    df_lt = df_lt.drop("%StudentReceived AcademicFeedback", axis=1)
    df_af = df_af.drop("%Student ReceivedMedia File(Let'sTalk +Video)", axis=1)
    print("Generating MediaHub Report")
    if name == "all":
        for i in np.unique(df[["main_teacher", "co_teacher"]].values):
            df_lt_personal = df_lt[(df_lt["main_teacher"] == i) | (df_lt["co_teacher"] == i)]
            df_af_personal = df_af[(df_af["main_teacher"] == i) | (df_af["co_teacher"] == i)]
            writer = pd.ExcelWriter("{} MH Report.xlsx".format(str.capwords(i)), engine='xlsxwriter')
            df_lt_personal.to_excel(writer, sheet_name="Let's Talk")
            df_af_personal.to_excel(writer, sheet_name="Academic Feedback")
            workbook = writer.book
            worksheet1, worksheet2 = writer.sheets["Let's Talk"], writer.sheets["Academic Feedback"]
            formatting = {'type': 'text',
                          'criteria': 'not containing',
                          'value': '@',
                          'format': workbook.add_format({'bg_color': '#ff0000'})}
            worksheet1.conditional_format('A1:Z999', formatting)
            worksheet2.conditional_format('A1:Z999', formatting)
            writer.save()

    else:
        writer = pd.ExcelWriter("{} MH Report.xlsx".format(str.capwords(name)), engine='xlsxwriter')
        df_lt.to_excel(writer, sheet_name="Let's Talk")
        df_af.to_excel(writer, sheet_name="Academic Feedback")
        workbook = writer.book
        worksheet1, worksheet2 = writer.sheets["Let's Talk"], writer.sheets["Academic Feedback"]
        formatting = {'type': 'text',
                      'criteria': 'not containing',
                      'value': '@',
                      'format': workbook.add_format({'bg_color': '#ff0000'})}
        worksheet1.conditional_format('A1:Z999', formatting)
        worksheet2.conditional_format('A1:Z999', formatting)
        writer.save()

try:
    #Choose which columns to use
    usecols = ["group_code", "Student English Name", "main_teacher",
                "co_teacher", "pa_name", "Marked Student",
                "%Student ReceivedMedia File(Let'sTalk +Video)",
                "%StudentReceived AcademicFeedback"]
    #Read data into data frame and generate report/s
    df = pd.read_csv(csv(), sep = '\t', usecols = usecols,
                     index_col = usecols[0], encoding = 'utf-16')
    #Lower all names before filtering
    df["main_teacher"] = df["main_teacher"].str.lower()
    df["co_teacher"] = df["co_teacher"].str.lower()
    generate_mh_report(df, username(df))
except (FileNotFoundError, ValueError):
    print("No .csv files found. Exiting program")
except AttributeError:
    print("Exiting program")
