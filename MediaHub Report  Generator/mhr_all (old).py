
import numpy as np
import pandas as pd
import string as str
#import matplotlib.pyplot as plt
#from pandas.plotting import table

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

try:
    #Choose which columns to use
    usecols = ["group_code", "Student English Name", "main_teacher",
                "co_teacher", "pa_name", "Marked Student",
                "%Student ReceivedMedia File(Let'sTalk +Video)",
                "%StudentReceived AcademicFeedback"]
    #Read data into pandas data frames
    df = pd.read_csv(csv(), sep = '\t', usecols = usecols,
                     index_col = usecols[0], encoding = 'utf-16')
    df_lt, df_af = df.copy(), df.copy()
    #Filter out 100% values
    hundreds = ["1", "100", "100%", "100.0%", "100.00%", "100.000%"]
    for i in hundreds:
        df_lt = df_lt[(df_lt["%Student ReceivedMedia File(Let'sTalk +Video)"] != i)]
        df_af = df_af[(df_af["%StudentReceived AcademicFeedback"] != i)]
    df_lt = df_lt.drop("%StudentReceived AcademicFeedback", axis=1)
    df_af = df_af.drop("%Student ReceivedMedia File(Let'sTalk +Video)", axis=1)
    #Generate MediaHub Report for each teacher
    print("Generating MediaHub Report")
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
except (FileNotFoundError, ValueError):
    print("No .csv files found. Exiting program")
except AttributeError:
    print("Exiting program")

#ax = plt.subplot(111, frame_on=False) # no visible frame
#ax.xaxis.set_visible(False)  # hide the x axis
#ax.yaxis.set_visible(False)  # hide the y axis
#table(ax, df.tail())
#plt.savefig('mytable.png', bbox_inches='tight')
