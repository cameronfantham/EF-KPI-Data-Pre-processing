import numpy as np
import pandas as pd
import string as str
#from pandas.plotting import table
#import matplotlib.pyplot as plt

def csv():
    #Verify .csv files in folder
    possible_csvs = ["MarkedStudent_DetailData.csv"]
    for i in range(1,10):
        possible_csvs.append("MarkedStudent_DetailData ({}).csv".format(i))
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
        name = input("Please enter your full name (type exit/e to exit)\n")
        if name.lower() in [x.lower() for x in names_list]: break
        elif name.lower() == "e" or name.lower() == "exit": return 0
        else:
            print("Name not recognised. Please enter valid name.")
            continue
    return name.lower()
def generate_report(df, name):
    if name == 0:
        print("Exiting program")
        return 0
    #Fliter data frame for teacher's name
    df = df[(df["main_teacher"] == name) | (df["co_teacher"] == name)]
    df_lt, df_af = df.copy(), df.copy()
    #Filter out 100% values
    hundreds = ["1", "100", "100%", "100.0%", "100.00%", "100.000%"]
    for i in hundreds:
        df_lt = df_lt[(df_lt["%Student ReceivedMedia File(Let'sTalk +Video)"] != i)]
        df_af = df_af[(df_af["%StudentReceived AcademicFeedback"] != i)]
    df_lt = df_lt.drop("%StudentReceived AcademicFeedback", axis=1)
    df_af = df_af.drop("%Student ReceivedMedia File(Let'sTalk +Video)", axis=1)
    #Choose report type
    print("Which report would you like? (type exit/e to exit)")
    report = input("Academic Feedback (af), Let's Talk (lt) or both?\n")
    if report.lower() == "lt" or report.lower().replace(" ","").replace("''","") == "letstalk":
        print("Generating Let's Talk Report for {}".format(name))
        writer = pd.ExcelWriter("{} LT Report.xlsx".format(str.capwords(name)), engine='xlsxwriter')
        df_lt.to_excel(writer, sheet_name="Let's Talk")
        workbook, worksheet = writer.book, writer.sheets["Let's Talk"]
        worksheet.conditional_format('A1:H999', {'type': 'text',
                      'criteria': 'not containing',
                      'value': '@',
                      'format': workbook.add_format({'bg_color': '#ff0000'})})
        writer.save()
    if report.lower() == "af" or report.lower().replace(" ","") == "academic feedback":
        print("Generating Academic Feedback Report for {}".format(name))
        writer = pd.ExcelWriter("{} AF Report.xlsx".format(str.capwords(name)), engine='xlsxwriter')
        df_af.to_excel(writer, sheet_name="Academic Feedback")
        workbook, worksheet = writer.book, writer.sheets["Academic Feedback"]
        worksheet.conditional_format('A1:H999', {'type': 'text',
                      'criteria': 'not containing',
                      'value': '@',
                      'format': workbook.add_format({'bg_color': '#ff0000'})})
        writer.save()
    if report.lower() == "both" or report.lower() == "b":
        print("Generating full report for {}".format(name))
        writer = pd.ExcelWriter("{} Report.xlsx".format(str.capwords(name)), engine='xlsxwriter')
        df_lt.to_excel(writer, sheet_name="Let's Talk")
        df_af.to_excel(writer, sheet_name="Academic Feedback")
        workbook = writer.book
        worksheet1, worksheet2 = writer.sheets["Let's Talk"], writer.sheets["Academic Feedback"]
        worksheet1.conditional_format('A1:H999', {'type': 'text',
                      'criteria': 'not containing',
                      'value': '@',
                      'format': workbook.add_format({'bg_color': '#ff0000'})})
        worksheet2.conditional_format('A1:H999', {'type': 'text',
                      'criteria': 'not containing',
                      'value': '@',
                      'format': workbook.add_format({'bg_color': '#ff0000'})})
        writer.save()
    if report.lower() == "exit" or report.lower() == "e":
        print("Exiting program")
        return 0

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
    generate_report(df, username(df))
except (FileNotFoundError, ValueError):
    print("No .csv files found. Exiting program")
except AttributeError:
    print("Exiting program")

#ax = plt.subplot(111, frame_on=False) # no visible frame
#ax.xaxis.set_visible(False)  # hide the x axis
#ax.yaxis.set_visible(False)  # hide the y axis
#table(ax, df.tail())
#plt.savefig('mytable.png', bbox_inches='tight')
