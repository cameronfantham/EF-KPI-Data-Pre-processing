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

try:
    #Choose which columns to use
    usecols = ["Reservation Teacher", "Group ID", "%Attendance MarkedOnTime",
               "TeachingTime (ACH)", "%TeacherTaskCompletion", "%PT App Use",
               "%SkillTestCompleted", "%Teacher-Led Skills Completion",
               "%BS CanDo Completion"]
    names = {"Reservation Teacher": "Reservation\nTeacher",
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
                     index_col=0, encoding='utf-16')
    df.rename(columns=names, inplace=True)
    df.fillna('0.0', inplace=True)
    df = df[df["Group ID"] == "Total"]
    df = df.drop("Group ID", axis=1)
    for i in list(df):
        if isinstance(df[i][0], str):
            df[i] = df[i].str.rstrip("%").astype(float)
    writer = pd.ExcelWriter("TD Report.xlsx", engine='xlsxwriter')
    df.style.set_properties(**{'text-align': 'center'}).to_excel(writer, sheet_name="Teacher Detail")
    #df.style.applymap(lambda _:'text-align: center').to_excel(writer, sheet_name="Teacher Detail")
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

#except (FileNotFoundError, ValueError):
#    print("No .csv files found. Exiting program")
except NameError:
    print("Exiting program")
